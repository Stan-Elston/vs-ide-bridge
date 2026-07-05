using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json.Nodes;
using static VsIdeBridgeService.ArgBuilder;
using static VsIdeBridgeService.SchemaHelpers;
using VsIdeBridgeService.SystemTools;
using IOPath = System.IO.Path;

namespace VsIdeBridgeService;

internal static partial class ToolCatalog
{
    private const string PythonExeName = "python.exe";
    private const string PythonListEnvsTool = "python_list_envs";
    private const string PythonListPackagesTool = "python_list_packages";

    private static IEnumerable<ToolEntry> PythonNativeTools() =>
        PythonDiscoveryTools()
            .Concat(PythonMutationTools())
            .Concat(CondaTools());

    private static IEnumerable<ToolEntry> PythonDiscoveryTools()
    {
        yield return new("python_eval",
            "NEVER guess at arithmetic — call this instead. Evaluates a Python expression and returns the exact result. " +
            "Pre-imported as globals (no import statement needed): math, statistics, decimal, fractions. " +
            "Examples: \"math.sqrt(2)\", \"round(math.pi, 10)\", \"statistics.mean([1.5, 2.3, 4.8])\", " +
            "\"fractions.Fraction(1,3) + fractions.Fraction(1,6)\", \"math.factorial(20)\".",
            ObjectSchema(
                Req("expression", "Python expression to evaluate. Full builtins are available and import works; math, statistics, decimal, fractions, re, itertools, and collections are pre-imported as globals."),
                Opt(Path, "Optional interpreter path. Defaults to the active interpreter, managed runtime, or first discovered Python.")),
            Python,
            async (id, args, bridge) =>
            {
                string expression = RequireArg(id, args, "expression");
                string python = ResolvePythonInterpreterPath(id, args);
                return await RunPythonEvalAsync(id, python, expression, timeoutMs: 5_000)
                    .ConfigureAwait(false);
            },
            searchHints: BuildSearchHints(
                related: [("python_exec", "Execute a multi-line snippet instead"), (PythonListEnvsTool, "Discover available interpreters")]));

        yield return new("python_exec",
            "Execute a multi-line Python snippet for calculations and data transforms. " +
            "Full builtins are available and import works; math, statistics, decimal, fractions, re, itertools, and collections are pre-imported as globals. " +
            "Assign to a variable named result to get it back as structured JSON alongside any print output.",
            ObjectSchema(
                Req("code", "Multi-line Python snippet. Full builtins are available and import works; math, statistics, decimal, fractions, re, itertools, and collections are already in scope. " +
                    "Assign to \"result\" to return a value as structured JSON."),
                Opt(Path, "Optional interpreter path. Defaults to the active interpreter, managed runtime, or first discovered Python.")),
            Python,
            async (id, args, bridge) =>
            {
                string code = RequireArg(id, args, "code");
                string python = ResolvePythonInterpreterPath(id, args);
                return await RunPythonExecAsync(id, python, code, timeoutMs: 10_000)
                    .ConfigureAwait(false);
            },
            searchHints: BuildSearchHints(
                related: [("python_eval", "Evaluate a single expression instead"), ("read_file", "Read a Python source file before executing it"), (PythonListPackagesTool, "Check available packages before running code")]));

        yield return new(PythonListEnvsTool,
            "Enumerate available Python interpreters: PATH entries, common venv locations under " +
            "the solution and repo root directories, and the bridge-managed Python runtime if present.",
            EmptySchema(), Python,
            async (id, _, bridge) =>
            {
                string solutionDir = ServiceToolPaths.ResolveSolutionDirectory(bridge);
                string repoRoot = ServiceToolPaths.ResolveRepoRootDirectory(bridge);
                List<string> found = FindPythonInterpreters(solutionDir, repoRoot);

                JsonArray arr = [];
                foreach (string p in found)
                    arr.Add(JsonValue.Create(p));

                JsonObject payload = new() { ["interpreters"] = arr, ["count"] = found.Count };
                return MakePythonResult(payload, success: true);
            },
            searchHints: BuildSearchHints(
                workflow: [("python_env_info", "Get version and site-packages for a discovered interpreter"), ("python_set_project_env", "Set the active VS Python interpreter")],
                related: [("python_eval", "Run a quick expression"), ("python_exec", "Execute a snippet")]));

        yield return BuildPythonEnvInfoTool();

        yield return new(PythonListPackagesTool,
            "List installed packages in a Python environment using pip list --format=json.",
            ObjectSchema(Req(Path, "Absolute path to the python.exe / python3 interpreter.")),
            Python,
            async (id, args, bridge) =>
            {
                string python = RequireArg(id, args, Path);
                return await RunPythonModuleAsync(
                    id, python, "pip list --format=json", timeoutMs: 30_000)
                    .ConfigureAwait(false);
            },
            searchHints: BuildSearchHints(
                workflow: [("python_install_package", "Install a missing package"), ("python_remove_package", "Remove an unused package")],
                related: [("python_env_info", "Get interpreter version info"), ("scan_project_dependencies", "Check NuGet and project-level dependency health")]));
    }

    private static ToolEntry BuildPythonEnvInfoTool()
    {
        return new("python_env_info",
            "Return version, executable path, and site-packages directory for a Python interpreter.",
            ObjectSchema(
                Req(Path, "Absolute path to the python.exe / python3 interpreter."),
                OptInt("timeout_ms", "Timeout in milliseconds for slow interpreter startup (default 30000, max 120000).")),
            Python,
            async (id, args, bridge) =>
            {
                string python = RequireArg(id, args, Path);
                int timeoutMs = args?["timeout_ms"]?.GetValue<int?>() ?? 30_000;
                timeoutMs = Math.Max(1_000, Math.Min(timeoutMs, 120_000));
                string script =
                    "import sys, json; " +
                    "import sysconfig; " +
                    "paths = sysconfig.get_paths(); " +
                    "print(json.dumps({" +
                    "'executable': sys.executable," +
                    "'version': sys.version," +
                    "'version_info': list(sys.version_info[:3])," +
                    "'site_packages': [p for p in (paths.get('purelib'), paths.get('platlib')) if p]" +
                    "}))";
                return await RunPythonScriptAsync(id, python, script, timeoutMs)
                    .ConfigureAwait(false);
            },
            searchHints: BuildSearchHints(
                workflow: [(PythonListPackagesTool, "List installed packages for this interpreter")],
                related: [(PythonListEnvsTool, "Discover available interpreters"), ("python_create_env", "Create a new virtual environment")]));
    }

    private static IEnumerable<ToolEntry> PythonMutationTools()
    {
        const string PackagesArgName = "packages";

        yield return new("python_install_package",
            "Install one or more packages into a Python environment using pip.",
            ObjectSchema(
                Req(Path, "Absolute path to the python.exe / python3 interpreter."),
                Req(PackagesArgName, "Space-separated package specifiers, e.g. \"requests>=2.28 flask\".")),
            Python,
            async (id, args, bridge) =>
            {
                string python = RequireArg(id, args, Path);
                string packages = RequireArg(id, args, PackagesArgName);
                return await RunPythonModuleAsync(
                    id, python, $"pip install {packages}", timeoutMs: 120_000)
                    .ConfigureAwait(false);
            },
            searchHints: BuildSearchHints(
                workflow: [(PythonListPackagesTool, "Verify the package was installed")],
                related: [("python_remove_package", "Uninstall a package"), ("python_create_env", "Create a clean environment first"), (PythonListEnvsTool, "Find the right interpreter to install into")]));

        yield return new("python_remove_package",
            "Uninstall one or more packages from a Python environment using pip.",
            ObjectSchema(
                Req(Path, "Absolute path to the python.exe / python3 interpreter."),
                Req(PackagesArgName, "Space-separated package names to remove.")),
            Python,
            async (id, args, bridge) =>
            {
                string python = RequireArg(id, args, Path);
                string packages = RequireArg(id, args, PackagesArgName);
                return await RunPythonModuleAsync(
                    id, python, $"pip uninstall -y {packages}", timeoutMs: 60_000)
                    .ConfigureAwait(false);
            },
            searchHints: BuildSearchHints(
                workflow: [(PythonListPackagesTool, "Verify the package was removed")],
                related: [("python_install_package", "Install a package instead")]));

        yield return new("python_create_env",
            "Create a new virtual environment at the specified path using python -m venv.",
            ObjectSchema(
                Req(Path, "Absolute path for the new virtual environment directory."),
                Opt("base_python", "Optional: path to the base interpreter to use (defaults to system python).")),
            Python,
            async (id, args, bridge) =>
            {
                string targetPath = RequireArg(id, args, Path);
                string? basePython = args?["base_python"]?.GetValue<string>();
                string python = string.IsNullOrWhiteSpace(basePython)
                    ? FindSystemPython()
                    : basePython;
                return await RunPythonModuleAsync(
                    id, python, $"venv \"{targetPath}\"", timeoutMs: 60_000)
                    .ConfigureAwait(false);
            },
            searchHints: BuildSearchHints(
                workflow: [("python_install_package", "Install packages into the new env"), (PythonListEnvsTool, "Verify the new env is discoverable")],
                related: [("python_env_info", "Inspect the new environment"), ("python_set_project_env", "Set the new env as the VS project interpreter")]));
    }

    private static IEnumerable<ToolEntry> CondaTools()
    {
        yield return new("conda_install",
            "Install packages into a conda environment.",
            ObjectSchema(
                Req("packages", "Space-separated package specifiers, e.g. \"numpy scipy\"."),
                Opt("env", "Optional: conda environment name or path. Defaults to base/active env.")),
            Python,
            async (id, args, bridge) =>
            {
                string conda = ResolveCondaExecutable(id);
                string packages = RequireArg(id, args, "packages");
                string? env = args?["env"]?.GetValue<string>();
                string envArg = string.IsNullOrWhiteSpace(env) ? string.Empty
                    : env.Contains(IOPath.DirectorySeparatorChar) || env.Contains('/')
                        ? $"-p \"{env}\""
                        : $"-n \"{env}\"";
                return await RunSubprocessAsync(
                    id, conda, $"install -y {envArg} {packages}".TrimEnd(),
                    workingDirectory: null, timeoutMs: 120_000)
                    .ConfigureAwait(false);
            },
            searchHints: BuildSearchHints(
                workflow: [(PythonListPackagesTool, "Verify the packages were installed")],
                related: [("conda_remove", "Remove conda packages"), (PythonListEnvsTool, "List available environments")]));

        yield return new("conda_remove",
            "Remove packages from a conda environment.",
            ObjectSchema(
                Req("packages", "Space-separated package names to remove."),
                Opt("env", "Optional: conda environment name or path.")),
            Python,
            async (id, args, bridge) =>
            {
                string conda = ResolveCondaExecutable(id);
                string packages = RequireArg(id, args, "packages");
                string? env = args?["env"]?.GetValue<string>();
                string envArg = string.IsNullOrWhiteSpace(env) ? string.Empty
                    : env.Contains(IOPath.DirectorySeparatorChar) || env.Contains('/')
                        ? $"-p \"{env}\""
                        : $"-n \"{env}\"";
                return await RunSubprocessAsync(
                    id, conda, $"remove -y {envArg} {packages}".TrimEnd(),
                    workingDirectory: null, timeoutMs: 60_000)
                    .ConfigureAwait(false);
            },
            searchHints: BuildSearchHints(
                workflow: [(PythonListPackagesTool, "Verify the packages were removed")],
                related: [("conda_install", "Install conda packages instead")]));
    }

    // ── Python helpers ─────────────────────────────────────────────────────────────────

    /// <summary>Run a Python -c script and return a structured result.</summary>
    private static async Task<JsonNode> RunPythonScriptAsync(
        JsonNode? id, string python, string script, int timeoutMs)
    {
        return await RunSubprocessAsync(id, python, $"-c \"{script.Replace("\"", "\\\"")}\"",
            workingDirectory: null, timeoutMs).ConfigureAwait(false);
    }

    private static async Task<JsonNode> RunPythonEvalAsync(
        JsonNode? id, string python, string expression, int timeoutMs)
    {
        string script =
            "import json, math, statistics, decimal, fractions, base64, sys, re, itertools, collections; " +
            "expr = base64.b64decode(sys.stdin.read()).decode('utf-8'); " +
            // Full builtins (and import) are available: omit '__builtins__' so Python injects the real
            // builtins module into globals. The scratchpad already runs a real Python subprocess behind
            // python_exec approval, so a restricted builtins allowlist added friction without real safety.
            "ns = {'math': math, 'statistics': statistics, 'decimal': decimal, 'fractions': fractions, 're': re, 'itertools': itertools, 'collections': collections}; " +
            "value = eval(expr, ns); " +
            "payload = {'expression': expr, 'type': type(value).__name__, 'repr': repr(value)}; " +
            "payload['json'] = json.loads(json.dumps(value, default=str)); " +
            "print(json.dumps(payload))";

        string encodedExpression = Convert.ToBase64String(Encoding.UTF8.GetBytes(expression));
        return await RunSubprocessAsync(id, python, $"-c \"{script.Replace("\"", "\\\"")}\"",
            workingDirectory: null, timeoutMs, encodedExpression).ConfigureAwait(false);
    }

    private static async Task<JsonNode> RunPythonExecAsync(
        JsonNode? id, string python, string code, int timeoutMs)
    {
        string script =
            "import json, math, statistics, decimal, fractions, base64, io, contextlib, sys, re, itertools, collections\n" +
            "source = base64.b64decode(sys.stdin.read()).decode('utf-8')\n" +
            // Full builtins (and import) are available: omit '__builtins__' so Python injects the real
            // builtins module into globals_dict. The scratchpad runs a real Python subprocess behind
            // python_exec approval, so a restricted builtins allowlist added friction without real safety.
            // Single namespace for globals AND locals: exec(source, ns) runs the snippet as if at module
            // scope, so top-level def/class/var names land in the same dict that functions resolve globals
            // against. Passing separate globals/locals dicts makes a helper defined in the snippet fail with
            // NameError on any other top-level name (functions look up __globals__, not the exec locals).
            "ns = {'math': math, 'statistics': statistics, 'decimal': decimal, 'fractions': fractions, 're': re, 'itertools': itertools, 'collections': collections}\n" +
            "stdout_buffer = io.StringIO()\n" +
            "with contextlib.redirect_stdout(stdout_buffer):\n" +
            "    exec(source, ns)\n" +
            "payload = {'stdout': stdout_buffer.getvalue(), 'hasResult': 'result' in ns}\n" +
            "if 'result' in ns:\n" +
            "    value = ns['result']\n" +
            "    payload['resultType'] = type(value).__name__\n" +
            "    payload['resultRepr'] = repr(value)\n" +
            "    payload['resultJson'] = json.loads(json.dumps(value, default=str))\n" +
            "print(json.dumps(payload))";

        string encodedCode = Convert.ToBase64String(Encoding.UTF8.GetBytes(code));
        return await RunSubprocessAsync(id, python, $"-c \"{script.Replace("\"", "\\\"")}\"",
            workingDirectory: null, timeoutMs, encodedCode).ConfigureAwait(false);
    }

    /// <summary>Run a python -m module command (e.g. pip install ...).</summary>
    private static async Task<JsonNode> RunPythonModuleAsync(
        JsonNode? id, string python, string moduleArgs, int timeoutMs)
    {
        return await RunSubprocessAsync(id, python, $"-m {moduleArgs}",
            workingDirectory: null, timeoutMs).ConfigureAwait(false);
    }

    /// <summary>General-purpose subprocess runner for python/conda tools.</summary>
    private static async Task<JsonNode> RunSubprocessAsync(
        JsonNode? id, string exe, string args, string? workingDirectory, int timeoutMs, string? standardInput = null)
    {
        ProcessStartInfo psi = new()
        {
            FileName = exe,
            Arguments = args,
            WorkingDirectory = workingDirectory ?? string.Empty,
            RedirectStandardInput = standardInput is not null,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
        };

        using Process process = Process.Start(psi)
            ?? throw new McpRequestException(id, McpErrorCodes.BridgeError,
                $"Failed to start process '{exe}'.");

        if (standardInput is not null)
        {
            await process.StandardInput.WriteAsync(standardInput).ConfigureAwait(false);
            process.StandardInput.Close();
        }

        Task<string> stdoutTask = process.StandardOutput.ReadToEndAsync();
        Task<string> stderrTask = process.StandardError.ReadToEndAsync();
        Task waitTask = process.WaitForExitAsync();

        if (!ReferenceEquals(
                await Task.WhenAny(waitTask, Task.Delay(timeoutMs)).ConfigureAwait(false),
                waitTask))
        {
            TryKillProcess(process);
            throw new McpRequestException(id, McpErrorCodes.TimeoutError,
                $"'{exe} {args}' timed out after {timeoutMs} ms.");
        }

        string stdout = await stdoutTask.ConfigureAwait(false);
        string stderr = await stderrTask.ConfigureAwait(false);
        bool success = process.ExitCode == 0;

        JsonObject payload = new()
        {
            ["success"] = success,
            ["exitCode"] = process.ExitCode,
            ["stdout"] = stdout,
            ["stderr"] = stderr,
        };

        return MakePythonResult(payload, success);
    }

    private static string ResolvePythonInterpreterPath(JsonNode? id, JsonObject? args)
    {
        string? explicitPath = args?[Path]?.GetValue<string>();
        if (!string.IsNullOrWhiteSpace(explicitPath))
        {
            return explicitPath;
        }

        string? activeInterpreterPath = PythonInterpreterState.LoadActiveInterpreterPath();
        if (!string.IsNullOrWhiteSpace(activeInterpreterPath) && File.Exists(activeInterpreterPath))
        {
            return activeInterpreterPath;
        }

        string managedRuntimePath = IOPath.GetFullPath(
            IOPath.Combine(AppContext.BaseDirectory, "..", "python", "managed-runtime", PythonExeName));
        if (File.Exists(managedRuntimePath))
        {
            return managedRuntimePath;
        }

        string fallbackPython = FindSystemPython();
        if (File.Exists(fallbackPython))
        {
            return fallbackPython;
        }

        throw new McpRequestException(id, McpErrorCodes.BridgeError,
            "Python interpreter not found. Pass 'path' explicitly or install the managed Python runtime.");
    }

    private static JsonNode MakePythonResult(JsonObject payload, bool success)
    {
        int exitCode = payload["exitCode"]?.GetValue<int>() ?? 0;
        string? stdout = payload["stdout"]?.GetValue<string>();
        string? stderr = payload["stderr"]?.GetValue<string>();
        string exitText = $"Python command completed with exit code {exitCode}.";
        string outputText = !string.IsNullOrWhiteSpace(stdout) ? stdout.TrimEnd()
            : !string.IsNullOrWhiteSpace(stderr) ? $"stderr: {stderr.TrimEnd()}"
            : string.Empty;
        string successText = string.IsNullOrWhiteSpace(outputText)
            ? exitText
            : exitText + "\n" + outputText;
        return ToolResultFormatter.StructuredToolResult(payload, isError: !success, successText: successText);
    }

    private static string RequireArg(JsonNode? id, JsonObject? args, string name)
    {
        string? value = args?[name]?.GetValue<string>();
        if (string.IsNullOrWhiteSpace(value))
            throw new McpRequestException(id, McpErrorCodes.InvalidParams,
                $"Missing required argument '{name}'.");
        return value;
    }

    // ── Python interpreter discovery ────────────────────────────────────────────────

    private static List<string> FindPythonInterpreters(string solutionDir, string repoRoot)
    {
        List<string> results = [];

        // 1. PATH scan
        string? pathEnv = Environment.GetEnvironmentVariable("PATH");
        if (!string.IsNullOrWhiteSpace(pathEnv))
        {
            foreach (string dir in pathEnv.Split(IOPath.PathSeparator,
                StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                foreach (string name in new[] { PythonExeName, "python3.exe", "python" })
                {
                    string candidate = IOPath.Combine(dir, name);
                    if (File.Exists(candidate) && !results.Contains(candidate))
                        results.Add(candidate);
                }
            }
        }

        // 2. Common venv locations — search both repo root and solution directory so that
        //    projects whose .sln lives in a build/ subdirectory still find venvs at the repo root.
        List<string> venvRoots = [repoRoot];
        if (!string.Equals(solutionDir, repoRoot, StringComparison.OrdinalIgnoreCase))
            venvRoots.Add(solutionDir);

        foreach (string venvRoot in venvRoots)
        {
            foreach (string rel in new[] { ".venv", "venv", "env", ".env" })
            {
                foreach (string name in new[] { PythonExeName, "python3.exe" })
                {
                    string candidate = IOPath.Combine(venvRoot, rel, "Scripts", name);
                    if (File.Exists(candidate) && !results.Contains(candidate))
                        results.Add(candidate);

                    // Linux/macOS style (bin/)
                    candidate = IOPath.Combine(venvRoot, rel, "bin", name);
                    if (File.Exists(candidate) && !results.Contains(candidate))
                        results.Add(candidate);
                }
            }
        }

        // 3. Bridge-managed Python (AppData\Local\Programs\Python)
        string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        string programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        foreach (string root in new[] { IOPath.Combine(localAppData, "Programs", "Python"),
                                        programFiles })
        {
            if (!Directory.Exists(root)) continue;
            foreach (string dir in Directory.EnumerateDirectories(root, "Python*",
                SearchOption.TopDirectoryOnly))
            {
                string candidate = IOPath.Combine(dir, PythonExeName);
                if (File.Exists(candidate) && !results.Contains(candidate))
                    results.Add(candidate);
            }
        }

        return results;
    }

    private static string FindSystemPython()
    {
        foreach (string name in new[] { PythonExeName, "python3.exe", "python", "python3" })
        {
            string? pathEnv = Environment.GetEnvironmentVariable("PATH");
            if (string.IsNullOrWhiteSpace(pathEnv)) continue;
            foreach (string dir in pathEnv.Split(IOPath.PathSeparator,
                StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                string candidate = IOPath.Combine(dir, name);
                if (File.Exists(candidate)) return candidate;
            }
        }

        throw new InvalidOperationException(
            "Python interpreter not found on PATH. Pass base_python explicitly.");
    }

    // ── Conda resolution ────────────────────────────────────────────────────────────

    private static volatile string? _resolvedConda;

    private static string ResolveCondaExecutable(JsonNode? id)
    {
        if (_resolvedConda is not null) return _resolvedConda;

        // 1. CONDA_EXE env var (set by conda activate)
        string? fromEnv = Environment.GetEnvironmentVariable("CONDA_EXE");
        if (!string.IsNullOrWhiteSpace(fromEnv))
        {
            if (File.Exists(fromEnv)) return _resolvedConda = fromEnv;
            throw new McpRequestException(id, McpErrorCodes.InvalidParams,
                $"CONDA_EXE points to '{fromEnv}' but the file does not exist.");
        }

        // 2. PATH scan
        string? pathEnv = Environment.GetEnvironmentVariable("PATH");
        if (!string.IsNullOrWhiteSpace(pathEnv))
        {
            foreach (string dir in pathEnv.Split(IOPath.PathSeparator,
                StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                foreach (string name in new[] { "conda.exe", "conda.bat", "conda" })
                {
                    string candidate = IOPath.Combine(dir, name);
                    if (File.Exists(candidate)) return _resolvedConda = candidate;
                }
            }
        }

        // 3. Known install locations (Miniconda/Anaconda under user profile + LocalAppData)
        string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        foreach (string root in new[] { userProfile, localAppData })
        {
            if (string.IsNullOrWhiteSpace(root)) continue;
            foreach (string rel in new[]
            {
                @"miniconda3\Scripts\conda.exe",
                @"miniconda3\condabin\conda.bat",
                @"anaconda3\Scripts\conda.exe",
                @"anaconda3\condabin\conda.bat",
                @"Miniconda3\Scripts\conda.exe",
                @"Anaconda3\Scripts\conda.exe",
            })
            {
                string candidate = IOPath.Combine(root, rel);
                if (File.Exists(candidate)) return _resolvedConda = candidate;
            }
        }

        throw new McpRequestException(id, McpErrorCodes.BridgeError,
            "Conda executable not found. Install Miniconda/Anaconda or set CONDA_EXE.");
    }
}
