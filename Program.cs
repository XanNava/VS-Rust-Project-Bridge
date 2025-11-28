using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

class Program {
	// ─────────────────────────────────────────────────────
	// DEFAULT CONFIG (adjust if needed)
	// ─────────────────────────────────────────────────────

	private static readonly string DefaultCsprojPath =
		@"C:\(Work)\RustVsBridgeHost\RustVsBridgeHost.csproj";

	private static readonly string DefaultRustProjectDir =
		@"C:\(Work)\RustVsBridge";

	private const string DefaultRustExeName = "RustVsBridge";

	private const string DefaultDevenvPath =
		@"C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\devenv.com";

	[DllImport("ole32.dll")]
	private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable pprot);

	[DllImport("ole32.dll")]
	private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

	[DllImport("user32.dll")]
	private static extern bool SetForegroundWindow(IntPtr hWnd);

	[DllImport("user32.dll")]
	private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

	private const int SW_RESTORE = 9;

	// ─────────────────────────────────────────────────────
	// ENTRY POINT
	// ─────────────────────────────────────────────────────

	[STAThread]
	static int Main(string[] args) {
		Console.WriteLine("────────────────────────────────────────────────────────");
		Console.WriteLine("[MAIN] VsRustAutoLink starting.");
		Console.WriteLine("  Raw args: " + (args.Length == 0 ? "<none>" : string.Join(" | ", args)));

		string csprojPath = args.Length >= 1 ? args[0] : DefaultCsprojPath;
		string rustProjectDir = args.Length >= 2 ? args[1] : DefaultRustProjectDir;
		string rustExeName = args.Length >= 3 ? args[2] : DefaultRustExeName;

		csprojPath = Path.GetFullPath(csprojPath);
		rustProjectDir = Path.GetFullPath(rustProjectDir);

		Console.WriteLine("[MAIN] Effective config:");
		Console.WriteLine($"  csprojPath     = {csprojPath}");
		Console.WriteLine($"  rustProjectDir = {rustProjectDir}");
		Console.WriteLine($"  rustExeName    = {rustExeName}");
		Console.WriteLine();

		Console.WriteLine("[CHECK] Rust project directory exists? " + Directory.Exists(rustProjectDir));
		if (!Directory.Exists(rustProjectDir)) {
			Console.WriteLine($"[ERROR] Rust project directory not found: {rustProjectDir}");
			return 1;
		}

		LogRustFilesSnapshot(rustProjectDir);
		EnsureRustMainInteractive(rustProjectDir);

		if (!File.Exists(csprojPath)) {
			Console.WriteLine("[CHECK] C# project not found at: " + csprojPath);
			Console.WriteLine("[ACTION] Creating new C# console host project…");
			try {
				CreateNewConsoleProject(csprojPath);
				Console.WriteLine("[OK] Created new C# console project.");
			}
			catch (Exception ex) {
				Console.WriteLine("[FATAL] Exception while creating new C# project:");
				Console.WriteLine(ex);
				return 1;
			}
		} else {
			Console.WriteLine("[CHECK] C# project already exists: " + csprojPath);
		}

		try {
			Console.WriteLine();
			Console.WriteLine("[STEP] Loading .csproj XML: " + csprojPath);
			var doc = XDocument.Load(csprojPath);
			Console.WriteLine("[OK] Loaded .csproj. Root element: " + (doc.Root?.Name.ToString() ?? "<null>"));

			XNamespace ns = doc.Root?.Name.Namespace ?? "";
			Console.WriteLine("[INFO] Using XML namespace: '" + ns + "'");

			AddOrUpdateRustProjectDir(doc, ns, rustProjectDir);
			AddRustItemGroup(doc, ns);
			AddRustBuildTarget(doc, ns, rustExeName);

			Console.WriteLine("[STEP] Saving updated .csproj.");
			doc.Save(csprojPath);
			Console.WriteLine("[OK] Project file saved.");

			EnsureHostProgramDelegatesToRust(csprojPath, rustProjectDir, rustExeName);

			Console.WriteLine();
			Console.WriteLine("[SUMMARY] Project updated successfully.");

			OpenInVisualStudio(csprojPath);

			Console.WriteLine("[MAIN] VsRustAutoLink finished successfully.");
			return 0;
		}
		catch (Exception ex) {
			Console.WriteLine();
			Console.WriteLine("[FATAL] Exception while updating project:");
			Console.WriteLine(ex);
			return 1;
		}
	}

	// ─────────────────────────────────────────────────────
	// RUST MAIN.RS ENSURER
	// ─────────────────────────────────────────────────────

	private static void EnsureRustMainInteractive(string rustProjectDir) {
		Console.WriteLine();
		Console.WriteLine("[EnsureRustMainInteractive] BEGIN");

		string srcDir = Path.Combine(rustProjectDir, "src");
		string mainPath = Path.Combine(srcDir, "main.rs");

		Console.WriteLine("  rustProjectDir = " + rustProjectDir);
		Console.WriteLine("  srcDir         = " + srcDir);
		Console.WriteLine("  main.rs path   = " + mainPath);

		try {
			if (!Directory.Exists(srcDir)) {
				Console.WriteLine("  src directory does not exist; creating it.");
				Directory.CreateDirectory(srcDir);
			} else {
				Console.WriteLine("  src directory already exists.");
			}

			if (File.Exists(mainPath)) {
				Console.WriteLine("  main.rs already exists; backing it up to main.rs.bak.");
				string backupPath = Path.Combine(srcDir, "main.rs.bak");
				try {
					File.Copy(mainPath, backupPath, overwrite: true);
					Console.WriteLine("  Backup written: " + backupPath);
				}
				catch (Exception exBackup) {
					Console.WriteLine("  WARNING: Failed to create backup main.rs.bak: " + exBackup);
				}
			} else {
				Console.WriteLine("  main.rs does not exist; will create a new one.");
			}

			string rustMain = @"
use std::io::{self, Write};
use std::time::SystemTime;

fn log(msg: &str) {
    let ts = SystemTime::now()
        .duration_since(SystemTime::UNIX_EPOCH)
        .unwrap()
        .as_millis();

    println!(""[LOG {}] {}"", ts, msg);
}

fn print_help() {
    println!(""Available commands:"");
    println!(""  help           - show this help"");
    println!(""  echo <text>    - print <text> back"");
    println!(""  exit           - quit the application"");
}

fn main() {
    log(""RustVsBridge started."");

    println!(""RustVsBridge interactive console."");
    println!(""Type 'help' for a list of commands."");

    let stdin = io::stdin();

    loop {
        print!(""> "");
        let _ = io::stdout().flush();

        let mut input = String::new();
        if let Err(e) = stdin.read_line(&mut input) {
            println!(""Error reading input: {e}"");
            break;
        }

        let line = input.trim();
        if line.is_empty() {
            continue;
        }

        if line.eq_ignore_ascii_case(""help"") {
            print_help();
            continue;
        }

        if line.eq_ignore_ascii_case(""exit"") {
            log(""Received 'exit' command. Shutting down."");
            break;
        }

        if let Some(rest) = line.strip_prefix(""echo "") {
            println!(""{}"", rest);
            continue;
        }

        if line.eq_ignore_ascii_case(""echo"") {
            println!();
            continue;
        }

        println!(""Unknown command: '{line}'. Type 'help' for commands."");
    }

    log(""RustVsBridge finished."");
}
".TrimStart('\r', '\n');

			Console.WriteLine("  Writing interactive main.rs.");
			File.WriteAllText(mainPath, rustMain);
			Console.WriteLine("  main.rs written/updated with interactive console.");
		}
		catch (Exception ex) {
			Console.WriteLine("[EnsureRustMainInteractive] ERROR: " + ex);
		}

		Console.WriteLine("[EnsureRustMainInteractive] END");
	}

	// ─────────────────────────────────────────────────────
	// RUST FILE SNAPSHOT / VALIDATION
	// ─────────────────────────────────────────────────────

	private static void LogRustFilesSnapshot(string rustProjectDir) {
		Console.WriteLine();
		Console.WriteLine("[RustFiles] BEGIN snapshot");
		Console.WriteLine("  Root = " + rustProjectDir);

		if (!Directory.Exists(rustProjectDir)) {
			Console.WriteLine("  Directory does not exist; skipping.");
			Console.WriteLine("[RustFiles] END snapshot");
			return;
		}

		string[] rsFiles;
		try {
			rsFiles = Directory.GetFiles(rustProjectDir, "*.rs", SearchOption.AllDirectories);
		}
		catch (Exception ex) {
			Console.WriteLine("  ERROR enumerating .rs files: " + ex);
			Console.WriteLine("[RustFiles] END snapshot");
			return;
		}

		Console.WriteLine($"  Found {rsFiles.Length} .rs files under Rust project.");

		int maxToShow = 30;
		for (int i = 0; i < rsFiles.Length && i < maxToShow; i++) {
			string full = rsFiles[i];
			string rel;
			try {
				rel = Path.GetRelativePath(rustProjectDir, full);
			}
			catch {
				rel = full;
			}
			string virtualLink = @"Rust\" + rel.Replace(Path.DirectorySeparatorChar, '\\');
			Console.WriteLine($"    [{i + 1:D2}] {rel}  ->  {virtualLink}");
		}

		if (rsFiles.Length > maxToShow) {
			Console.WriteLine($"  … plus {rsFiles.Length - maxToShow} more .rs files not listed.");
		}

		Console.WriteLine("  NOTE: All of these are included via:");
		Console.WriteLine("    <None Include=\"$(RustProjectDir)**\\*.*\">");
		Console.WriteLine("      <Link>Rust\\%(RecursiveDir)%(FileName)%(Extension)</Link>");
		Console.WriteLine("    </None>");
		Console.WriteLine("  After reload, they should appear under a virtual 'Rust' folder in Solution Explorer.");
		Console.WriteLine("[RustFiles] END snapshot");
	}

	// ─────────────────────────────────────────────────────
	// C# HOST PROJECT CREATION
	// ─────────────────────────────────────────────────────

	private static void CreateNewConsoleProject(string csprojPath) {
		string projDir = Path.GetDirectoryName(csprojPath)
						 ?? throw new InvalidOperationException("Invalid csproj path.");

		string projName = Path.GetFileNameWithoutExtension(csprojPath);

		Console.WriteLine("[CreateNewConsoleProject] projDir=" + projDir);
		Console.WriteLine("[CreateNewConsoleProject] projName=" + projName);

		if (!Directory.Exists(projDir)) {
			Console.WriteLine("[CreateNewConsoleProject] Creating directory: " + projDir);
			Directory.CreateDirectory(projDir);
		}

		string csprojText = @"
<Project Sdk=""Microsoft.NET.Sdk"">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>
</Project>";

		Console.WriteLine("[CreateNewConsoleProject] Writing .csproj file…");
		File.WriteAllText(csprojPath, csprojText.Trim());

		string programPath = Path.Combine(projDir, "Program.cs");
		if (!File.Exists(programPath)) {
			string projNamespace = SanitizeNamespace(projName);
			Console.WriteLine("[CreateNewConsoleProject] Program.cs not found, creating stub. Namespace=" + projNamespace);

			string programCode = $@"
using System;

namespace {projNamespace}
{{
    internal class Program
    {{
        static void Main(string[] args)
        {{
            Console.WriteLine(""Hello from {projName} (auto-generated host for Rust project)."");
        }}
    }}
}}";
			File.WriteAllText(programPath, programCode.Trim());
		} else {
			Console.WriteLine("[CreateNewConsoleProject] Program.cs already exists; leaving as-is for now.");
		}
	}

	private static string SanitizeNamespace(string name) {
		char[] chars = name.ToCharArray();
		for (int i = 0; i < chars.Length; i++) {
			if (!char.IsLetterOrDigit(chars[i]) && chars[i] != '_')
				chars[i] = '_';
		}
		if (!char.IsLetter(chars[0]) && chars[0] != '_')
			chars[0] = '_';
		return new string(chars);
	}

	// ─────────────────────────────────────────────────────
	// PROJECT MODIFICATION: RUST WIRING
	// ─────────────────────────────────────────────────────

	private static void AddOrUpdateRustProjectDir(XDocument doc, XNamespace ns, string rustProjectDir) {
		Console.WriteLine();
		Console.WriteLine("[AddOrUpdateRustProjectDir] BEGIN");
		Console.WriteLine("  Incoming RustProjectDir = " + rustProjectDir);

		if (!rustProjectDir.EndsWith(Path.DirectorySeparatorChar.ToString()) &&
			!rustProjectDir.EndsWith(Path.AltDirectorySeparatorChar.ToString())) {
			rustProjectDir += Path.DirectorySeparatorChar;
			Console.WriteLine("  Normalized RustProjectDir with trailing slash: " + rustProjectDir);
		}

		var proj = doc.Root ?? throw new InvalidOperationException("Invalid .csproj: missing root element.");

		var propertyGroup = proj.Element(ns + "PropertyGroup");
		if (propertyGroup == null) {
			Console.WriteLine("  No <PropertyGroup> found; creating new one.");
			propertyGroup = new XElement(ns + "PropertyGroup");
			proj.AddFirst(propertyGroup);
		} else {
			Console.WriteLine("  Found existing <PropertyGroup>.");
		}

		var rustDirProp = propertyGroup.Element(ns + "RustProjectDir");
		if (rustDirProp == null) {
			Console.WriteLine("  <RustProjectDir> not found; creating.");
			rustDirProp = new XElement(ns + "RustProjectDir");
			propertyGroup.Add(rustDirProp);
		} else {
			Console.WriteLine("  Found existing <RustProjectDir> element; updating value.");
		}

		rustDirProp.Value = rustProjectDir;
		Console.WriteLine("  Set <RustProjectDir> to: " + rustProjectDir);
		Console.WriteLine("[AddOrUpdateRustProjectDir] END");
	}

	private static void AddRustItemGroup(XDocument doc, XNamespace ns) {
		Console.WriteLine();
		Console.WriteLine("[AddRustItemGroup] BEGIN");

		var proj = doc.Root ?? throw new InvalidOperationException("Invalid .csproj: missing root element.");

		var rustNoneElements = proj.Elements(ns + "ItemGroup")
			.SelectMany(ig => ig.Elements(ns + "None"))
			.Where(n => {
				var include = (string?)n.Attribute("Include");
				bool isRust = include != null && include.Contains("$(RustProjectDir)", StringComparison.OrdinalIgnoreCase);
				if (isRust) {
					Console.WriteLine("  Found existing Rust-linked <None Include=\"" + include + "\">");
				}
				return isRust;
			})
			.ToList();

		Console.WriteLine("  Existing Rust-linked <None> elements to remove: " + rustNoneElements.Count);

		foreach (var none in rustNoneElements) {
			var parent = none.Parent;
			Console.WriteLine("  Removing <None Include=\"{0}\">", (string?)none.Attribute("Include") ?? "<unknown>");
			none.Remove();

			if (parent != null && !parent.Elements().Any()) {
				Console.WriteLine("  Parent <ItemGroup> is now empty; removing it as well.");
				parent.Remove();
			}
		}

		var newGroup = new XElement(ns + "ItemGroup",
			new XElement(ns + "None",
				new XAttribute("Include", @"$(RustProjectDir)**\*.*"),
				new XElement(ns + "Link", @"Rust\%(RecursiveDir)%(FileName)%(Extension)")
			)
		);

		proj.Add(newGroup);
		Console.WriteLine("  Added fresh <ItemGroup> with <None Include=\"$(RustProjectDir)**\\*.*\">.");
		Console.WriteLine("[AddRustItemGroup] END");
	}

	/// <summary>
	/// Adds/updates the BuildAndRunRust target.
	/// NOW: Only builds Rust (cargo build) on Debug,
	/// and first attempts to kill any running RustVsBridge.exe
	/// so Windows doesn’t block overwriting the EXE.
	/// </summary>
	private static void AddRustBuildTarget(XDocument doc, XNamespace ns, string rustExeName) {
		Console.WriteLine();
		Console.WriteLine("[AddRustBuildTarget] BEGIN");

		var proj = doc.Root ?? throw new InvalidOperationException("Invalid .csproj: missing root element.");

		// Remove any existing BuildAndRunRust target so we can rewrite it cleanly
		var existingTargets = proj.Elements(ns + "Target")
			.Where(t => string.Equals((string?)t.Attribute("Name"), "BuildAndRunRust", StringComparison.OrdinalIgnoreCase))
			.ToList();

		Console.WriteLine("  Existing <Target Name=\"BuildAndRunRust\"> count: " + existingTargets.Count);

		foreach (var t in existingTargets) {
			Console.WriteLine("  Removing old BuildAndRunRust target.");
			t.Remove();
		}

		// Kill any stale RustVsBridge.exe (ignore errors), then build.
		// - taskkill may fail if the process isn't running; that's fine.
		// - '&' means the cd/cargo part runs regardless.
		string buildCommand =
			"taskkill /IM RustVsBridge.exe /F >nul 2>nul || echo No running RustVsBridge.exe to kill." +
			" & cd /d \"$(RustProjectDir)\" & cargo build";

		Console.WriteLine("  New Build command (no run): " + buildCommand);
		Console.WriteLine("  Target will only run when Configuration == Debug.");

		var targetElement = new XElement(ns + "Target",
			new XAttribute("Name", "BuildAndRunRust"),
			new XAttribute("AfterTargets", "Build"),
			new XAttribute("Condition", " '$(Configuration)' == 'Debug' "),

			new XElement(ns + "Exec",
				new XAttribute("Command", buildCommand))
		);

		proj.Add(targetElement);
		Console.WriteLine("  Added/updated <Target Name=\"BuildAndRunRust\"> (build-only, Debug, with pre-kill).");
		Console.WriteLine("[AddRustBuildTarget] END");
	}

	// ─────────────────────────────────────────────────────
	// HOST PROGRAM PATCH
	// ─────────────────────────────────────────────────────

	private static void EnsureHostProgramDelegatesToRust(string csprojPath, string rustProjectDir, string rustExeName) {
		Console.WriteLine();
		Console.WriteLine("[EnsureHostProgramDelegatesToRust] BEGIN (forced overwrite)");

		string projDir = Path.GetDirectoryName(csprojPath)
						 ?? throw new InvalidOperationException("Invalid csproj path.");
		string projName = Path.GetFileNameWithoutExtension(csprojPath);
		string programPath = Path.Combine(projDir, "Program.cs");

		Console.WriteLine("  projDir     = " + projDir);
		Console.WriteLine("  projName    = " + projName);
		Console.WriteLine("  Program.cs  = " + programPath);

		string projNamespace = SanitizeNamespace(projName);
		string rustExePath = Path.Combine(rustProjectDir, "target", "debug", rustExeName + ".exe");
		Console.WriteLine("  Expected Rust exe path = " + rustExePath);

		string newProgramCode = $@"
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace {projNamespace}
{{
    internal class Program
    {{
        static int Main(string[] args)
        {{
            var rustExe = @""{rustExePath.Replace(@"\", @"\\")}"";

            if (!File.Exists(rustExe))
            {{
                Console.Error.WriteLine($""[HOST] Rust executable not found at '{{rustExe}}'. Did you build the Rust project?"");
                return 1;
            }}

            var startInfo = new ProcessStartInfo
            {{
                FileName = rustExe,
                Arguments = string.Join("" "", args.Select(a => ""\"""" + a + ""\"""" )),
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = false,
                WorkingDirectory = Path.GetDirectoryName(rustExe) ?? Environment.CurrentDirectory
            }};

            using var process = new Process {{ StartInfo = startInfo }};

            process.OutputDataReceived += (sender, e) =>
            {{
                if (e.Data != null)
                    Console.WriteLine(e.Data);
            }};

            process.ErrorDataReceived += (sender, e) =>
            {{
                if (e.Data != null)
                    Console.Error.WriteLine(e.Data);
            }};

            Console.WriteLine($""[HOST] Launching Rust process: '{{rustExe}}'"");

            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            process.WaitForExit();

            Console.WriteLine($""[HOST] Rust process exited with code {{process.ExitCode}}."");
            return process.ExitCode;
        }}
    }}
}}";

		Console.WriteLine("  Overwriting Program.cs with Rust host.");
		File.WriteAllText(programPath, newProgramCode.Trim());
		Console.WriteLine("[EnsureHostProgramDelegatesToRust] Program.cs written.");
		Console.WriteLine("[EnsureHostProgramDelegatesToRust] END");
	}

	// ─────────────────────────────────────────────────────
	// VISUAL STUDIO LAUNCH / REUSE
	// ─────────────────────────────────────────────────────

	private static void OpenInVisualStudio(string csprojPath) {
		Console.WriteLine();
		Console.WriteLine("[OpenInVisualStudio] BEGIN");

		try {
			string projDir = Path.GetDirectoryName(csprojPath)
							 ?? throw new InvalidOperationException("Invalid csproj path.");
			Console.WriteLine("  projDir = " + projDir);

			string[] slnFiles = Directory.GetFiles(projDir, "*.sln", SearchOption.TopDirectoryOnly);
			Console.WriteLine("  .sln files found in directory: " + slnFiles.Length);

			foreach (var slnFile in slnFiles)
				Console.WriteLine("    sln: " + slnFile);

			string? solutionPath = slnFiles.FirstOrDefault();
			string toOpen = solutionPath ?? csprojPath;

			Console.WriteLine("  Target to open in VS = " + toOpen);
			Console.WriteLine("  Using devenv.com at: " + DefaultDevenvPath);

			if (solutionPath != null) {
				Console.WriteLine("  Checking for existing VS instance with solution: " + solutionPath);
				object? dte = TryGetDteForSolution(solutionPath);
				if (dte != null) {
					Console.WriteLine("  Found existing DTE for solution; bringing to front.");
					TryActivateMainWindow(dte);
					try { Marshal.ReleaseComObject(dte); }
					catch { }
					Console.WriteLine("  Reused existing VS instance; not launching new devenv.");
					Console.WriteLine("[OpenInVisualStudio] END");
					return;
				}
				Console.WriteLine("  No existing DTE found for solution; will launch new VS instance.");
			} else {
				Console.WriteLine("  No .sln found; will open .csproj directly in VS.");
			}

			string args = $"\"{toOpen}\"";
			RunDevenv("OpenHostProjectOrSolution", args, waitForExit: false);
		}
		catch (Exception ex) {
			Console.WriteLine("[OpenInVisualStudio] Failed to open in VS 2026:");
			Console.WriteLine(ex);
			try {
				Console.WriteLine("[OpenInVisualStudio] Falling back to default file association…");
				var fallback = new ProcessStartInfo {
					FileName = csprojPath,
					UseShellExecute = true
				};
				Process.Start(fallback);
			}
			catch (Exception ex2) {
				Console.WriteLine("[OpenInVisualStudio] Fallback open also failed:");
				Console.WriteLine(ex2);
			}
		}

		Console.WriteLine("[OpenInVisualStudio] END");
	}

	private static string ResolveDevenvPath() {
		try {
			bool exists = File.Exists(DefaultDevenvPath);
			Console.WriteLine("[ResolveDevenvPath] DefaultDevenvPath exists? " + exists);
			if (!exists) {
				Console.WriteLine("[ResolveDevenvPath] WARNING: DefaultDevenvPath does not exist: " + DefaultDevenvPath);
			}
		}
		catch (Exception ex) {
			Console.WriteLine("[ResolveDevenvPath] Exception while checking DefaultDevenvPath:");
			Console.WriteLine(ex);
		}
		return DefaultDevenvPath;
	}

	private static void RunDevenv(string operationLabel, string arguments, bool waitForExit) {
		string devenvPath = ResolveDevenvPath();

		Console.WriteLine($"[RunDevenv:{operationLabel}] Preparing to start devenv.");
		Console.WriteLine($"[RunDevenv:{operationLabel}] Path='{devenvPath}', Args='{arguments}'");
		Console.WriteLine($"[RunDevenv:{operationLabel}] File.Exists(devenvPath)={File.Exists(devenvPath)}");

		var psi = new ProcessStartInfo {
			FileName = devenvPath,
			Arguments = arguments,
			UseShellExecute = true,
			CreateNoWindow = false,
			WorkingDirectory = Path.GetDirectoryName(devenvPath) ?? Environment.CurrentDirectory
		};

		Console.WriteLine($"[RunDevenv:{operationLabel}] WorkingDirectory='{psi.WorkingDirectory}'");

		try {
			using var process = new Process { StartInfo = psi };
			Console.WriteLine($"[RunDevenv:{operationLabel}] Starting process.");
			process.Start();
			Console.WriteLine($"[RunDevenv:{operationLabel}] Started, PID={process.Id}");

			if (waitForExit) {
				Console.WriteLine($"[RunDevenv:{operationLabel}] Waiting for exit…");
				process.WaitForExit();
				Console.WriteLine($"[RunDevenv:{operationLabel}] Process exited with code: {process.ExitCode}");
			} else {
				Console.WriteLine($"[RunDevenv:{operationLabel}] Not waiting for exit (fire-and-forget).");
			}
		}
		catch (Exception ex) {
			Console.WriteLine($"[RunDevenv:{operationLabel}] ERROR starting Visual Studio:");
			Console.WriteLine(ex);
		}
	}

	// ─────────────────────────────────────────────────────
	// DTE DISCOVERY / WINDOW ACTIVATION
	// ─────────────────────────────────────────────────────

	private static object? TryGetDteForSolution(string solutionPath) {
		Console.WriteLine("[TryGetDteForSolution] ENTER. solutionPath=" + solutionPath);

		string normalizedTarget;
		try {
			normalizedTarget = NormalizePath(solutionPath);
		}
		catch (Exception ex) {
			Console.WriteLine("[TryGetDteForSolution] ERROR normalizing target solution path: " + ex);
			return null;
		}

		if (!TryGetRunningObjectTable(out var rot)) {
			Console.WriteLine("[TryGetDteForSolution] GetRunningObjectTable failed.");
			return null;
		}

		if (!TryCreateBindContext(out var ctx)) {
			Console.WriteLine("[TryGetDteForSolution] CreateBindCtx failed.");
			return null;
		}

		try {
			object? dte = FindDteForSolution(rot, ctx, normalizedTarget);
			if (dte != null) {
				Console.WriteLine("[TryGetDteForSolution] MATCH found for solution.");
				return dte;
			}
		}
		catch (Exception ex) {
			Console.WriteLine("[TryGetDteForSolution] ERROR enumerating ROT: " + ex);
		}

		Console.WriteLine("[TryGetDteForSolution] EXIT with null (no matching DTE).");
		return null;
	}

	private static bool TryGetRunningObjectTable(out IRunningObjectTable rot) {
		int hr = GetRunningObjectTable(0, out rot);
		if (hr != 0 || rot == null) {
			Console.WriteLine("[TryGetDteForSolution] GetRunningObjectTable failed. hr=" + hr);
			return false;
		}
		return true;
	}

	private static bool TryCreateBindContext(out IBindCtx ctx) {
		int hr = CreateBindCtx(0, out ctx);
		if (hr != 0 || ctx == null) {
			Console.WriteLine("[TryGetDteForSolution] CreateBindCtx failed. hr=" + hr);
			return false;
		}
		return true;
	}

	private static object? FindDteForSolution(IRunningObjectTable rot, IBindCtx ctx, string normalizedTarget) {
		rot.EnumRunning(out IEnumMoniker enumMoniker);
		if (enumMoniker == null) {
			Console.WriteLine("[FindDteForSolution] EnumRunning returned null.");
			return null;
		}

		IMoniker[] monikers = new IMoniker[1];

		while (enumMoniker.Next(1, monikers, IntPtr.Zero) == 0) {
			var moniker = monikers[0];
			if (moniker == null)
				continue;

			string displayName;
			try {
				moniker.GetDisplayName(ctx, null, out displayName);
			}
			catch (Exception ex) {
				Console.WriteLine("[FindDteForSolution] ERROR getting moniker display name: " + ex);
				continue;
			}

			if (!displayName.Contains("VisualStudio.DTE.", StringComparison.OrdinalIgnoreCase))
				continue;

			Console.WriteLine("[FindDteForSolution] ROT entry: " + displayName);

			try {
				rot.GetObject(moniker, out object dteObject);
				if (dteObject == null) {
					Console.WriteLine("[FindDteForSolution] ROT GetObject returned null for VS entry.");
					continue;
				}

				var dteType = dteObject.GetType();
				object? solutionObj = dteType.InvokeMember(
					"Solution",
					System.Reflection.BindingFlags.GetProperty,
					null,
					dteObject,
					null);

				if (solutionObj == null) {
					Console.WriteLine("[FindDteForSolution] DTE.Solution is null.");
					continue;
				}

				var solutionType = solutionObj.GetType();
				string? fullName = solutionType
					.GetProperty("FullName")?
					.GetValue(solutionObj) as string;

				if (string.IsNullOrEmpty(fullName)) {
					Console.WriteLine("[FindDteForSolution] Solution.FullName is null/empty.");
					continue;
				}

				string normalizedExisting = NormalizePath(fullName);
				Console.WriteLine($"[FindDteForSolution] DTE solution: '{fullName}' -> normalized='{normalizedExisting}'");

				if (normalizedExisting == normalizedTarget) {
					Console.WriteLine("[FindDteForSolution] MATCH found.");
					return dteObject;
				}
			}
			catch (Exception ex) {
				Console.WriteLine("[FindDteForSolution] ERROR inspecting DTE object: " + ex);
			}
		}

		return null;
	}

	private static string NormalizePath(string path) {
		return Path.GetFullPath(path)
			.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
			.ToLowerInvariant();
	}

	private static void TryActivateMainWindow(object dteObject) {
		try {
			var dteType = dteObject.GetType();
			object? mainWindowObject = dteType.InvokeMember(
				"MainWindow",
				System.Reflection.BindingFlags.GetProperty,
				null,
				dteObject,
				null);

			if (mainWindowObject == null) {
				Console.WriteLine("[TryActivateMainWindow] MainWindow is null.");
				return;
			}

			var mainWindowType = mainWindowObject.GetType();
			Console.WriteLine("[TryActivateMainWindow] Activating EnvDTE.MainWindow…");
			mainWindowType.InvokeMember(
				"Activate",
				System.Reflection.BindingFlags.InvokeMethod,
				null,
				mainWindowObject,
				null);

			try {
				object? hwndValue = mainWindowType
					.GetProperty("HWnd")?
					.GetValue(mainWindowObject);

				if (hwndValue is int hwndInt && hwndInt != 0) {
					IntPtr hwnd = new IntPtr(hwndInt);
					Console.WriteLine("[TryActivateMainWindow] MainWindow.HWnd=" + hwndInt + " -> bringing to front.");
					ShowWindowAsync(hwnd, SW_RESTORE);
					SetForegroundWindow(hwnd);
				} else {
					Console.WriteLine("[TryActivateMainWindow] MainWindow.HWnd is null/0; skipping user32 calls.");
				}
			}
			catch (Exception ex) {
				Console.WriteLine("[TryActivateMainWindow] ERROR while using HWnd: " + ex);
			}
		}
		catch (Exception ex) {
			Console.WriteLine("[TryActivateMainWindow] ERROR activating MainWindow: " + ex);
		}
	}
}
