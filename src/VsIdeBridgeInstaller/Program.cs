using System.ComponentModel;
using System.Diagnostics;
using System.Security.Principal;

namespace VsIdeBridgeInstaller;

internal static class Program
{
    private const string DefaultInstallDir = @"C:\Program Files\VsIdeBridge";
    private const string DefaultServiceName = "VsIdeBridgeHost";
    private const string DefaultVsixId = "StanElston.VsIdeBridge";

    private static int Main(string[] args)
    {
        if (!OperatingSystem.IsWindows())
        {
            Console.Error.WriteLine("ERROR: installer is Windows-only.");
            return 1;
        }

        if (args.Length == 0 || IsHelp(args[0]))
        {
            PrintHelp();
            return 0;
        }

        var verb = args[0].Trim().ToLowerInvariant();
        var options = ParseOptions(args.Skip(1).ToArray());

        try
        {
            EnsureAdmin();

            return verb switch
            {
                "install" => RunInstall(options),
                "uninstall" => RunUninstall(options),
                _ => Fail($"Unknown command '{verb}'.")
            };
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"ERROR: {ex.Message}");
            return 1;
        }
    }

    private static int RunInstall(Dictionary<string, string?> options)
    {
        var repoRoot = GetPathOption(options, "repo-root") ?? FindRepoRoot();
        var configuration = GetOption(options, "configuration") ?? "Debug";
        var installDir = GetPathOption(options, "install-dir") ?? DefaultInstallDir;
        var serviceName = GetOption(options, "service-name") ?? DefaultServiceName;
        var vsixId = GetOption(options, "vsix-id") ?? DefaultVsixId;
        var cliSource = GetPathOption(options, "cli-source")
            ?? Path.Combine(repoRoot, "src", "VsIdeBridgeCli", "bin", configuration, "net8.0");
        var vsixPath = GetPathOption(options, "vsix-path")
            ?? Path.Combine(repoRoot, "src", "VsIdeBridge", "bin", configuration, "net472", "VsIdeBridge.vsix");

        var skipVsix = HasFlag(options, "skip-vsix");
        var installService = HasFlag(options, "install-service");
        var serviceExePath = GetPathOption(options, "service-exe");

        if (!Directory.Exists(cliSource))
        {
            return Fail($"CLI source directory not found: {cliSource}");
        }

        Directory.CreateDirectory(installDir);
        var cliDest = Path.Combine(installDir, "cli");
        CopyDirectory(cliSource, cliDest);
        Console.WriteLine($"Installed CLI files -> {cliDest}");

        if (!skipVsix)
        {
            if (!File.Exists(vsixPath))
            {
                return Fail($"VSIX not found: {vsixPath}");
            }

            InstallVsix(vsixPath);
            Console.WriteLine($"VSIX installed/updated ({vsixId}).");
        }

        if (installService)
        {
            if (string.IsNullOrWhiteSpace(serviceExePath))
            {
                return Fail("--install-service requires --service-exe <absolute path>.");
            }

            if (!Path.IsPathRooted(serviceExePath) || !File.Exists(serviceExePath))
            {
                return Fail($"Service executable not found: {serviceExePath}");
            }

            InstallOrUpdateService(serviceName, serviceExePath);
            Console.WriteLine($"Service '{serviceName}' installed (StartType=Manual).");
        }

        Console.WriteLine("Install complete.");
        return 0;
    }

    private static int RunUninstall(Dictionary<string, string?> options)
    {
        var installDir = GetPathOption(options, "install-dir") ?? DefaultInstallDir;
        var serviceName = GetOption(options, "service-name") ?? DefaultServiceName;
        var vsixId = GetOption(options, "vsix-id") ?? DefaultVsixId;
        var skipVsix = HasFlag(options, "skip-vsix");
        var removeService = HasFlag(options, "remove-service");

        if (!skipVsix)
        {
            UninstallVsix(vsixId);
            Console.WriteLine($"VSIX uninstall attempted ({vsixId}).");
        }

        if (removeService)
        {
            RemoveService(serviceName);
            Console.WriteLine($"Service '{serviceName}' removed if it existed.");
        }

        if (Directory.Exists(installDir))
        {
            Directory.Delete(installDir, recursive: true);
            Console.WriteLine($"Removed install directory: {installDir}");
        }

        Console.WriteLine("Uninstall complete.");
        return 0;
    }

    private static void InstallVsix(string vsixPath)
    {
        var installer = FindVsixInstallerPath();
        var exitCode = RunProcess(installer, $"/quiet \"{vsixPath}\"");
        if (exitCode != 0)
        {
            throw new InvalidOperationException($"VSIX install failed with exit code {exitCode}.");
        }
    }

    private static void UninstallVsix(string vsixId)
    {
        var installer = FindVsixInstallerPath();
        var exitCode = RunProcess(installer, $"/quiet /uninstall:{vsixId}");
        if (exitCode != 0)
        {
            Console.WriteLine($"VSIX uninstall returned exit code {exitCode}. Continuing.");
        }
    }

    private static void InstallOrUpdateService(string serviceName, string serviceExePath)
    {
        RemoveService(serviceName);
        var createArgs = $"create \"{serviceName}\" binPath= \"\"{serviceExePath}\"\" start= demand DisplayName= \"VS IDE Bridge Host\"";
        var createExit = RunProcess("sc.exe", createArgs);
        if (createExit != 0)
        {
            throw new InvalidOperationException($"Failed to create service '{serviceName}'. Exit code: {createExit}");
        }

        var descExit = RunProcess("sc.exe", $"description \"{serviceName}\" \"VS IDE Bridge background host (manual start).\"");
        if (descExit != 0)
        {
            Console.WriteLine($"Failed to set service description. Exit code: {descExit}");
        }
    }

    private static void RemoveService(string serviceName)
    {
        RunProcess("sc.exe", $"stop \"{serviceName}\"");
        RunProcess("sc.exe", $"delete \"{serviceName}\"");
    }

    private static string FindVsixInstallerPath()
    {
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var root = Path.Combine(programFiles, "Microsoft Visual Studio", "18");
        if (!Directory.Exists(root))
        {
            throw new InvalidOperationException("Visual Studio 18 installation path not found.");
        }

        var editions = new[] { "Enterprise", "Professional", "Community", "Preview" };
        foreach (var edition in editions)
        {
            var candidate = Path.Combine(root, edition, "Common7", "IDE", "VSIXInstaller.exe");
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        var fallback = Directory.GetFiles(root, "VSIXInstaller.exe", SearchOption.AllDirectories).FirstOrDefault();
        if (!string.IsNullOrWhiteSpace(fallback))
        {
            return fallback;
        }

        throw new FileNotFoundException("VSIXInstaller.exe not found under Visual Studio 18 install path.");
    }

    private static int RunProcess(string fileName, string arguments)
    {
        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
            }
        };

        process.Start();
        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        process.WaitForExit();

        if (!string.IsNullOrWhiteSpace(stdout))
        {
            Console.WriteLine(stdout.Trim());
        }

        if (!string.IsNullOrWhiteSpace(stderr))
        {
            Console.Error.WriteLine(stderr.Trim());
        }

        return process.ExitCode;
    }

    private static void CopyDirectory(string sourceDir, string destinationDir)
    {
        Directory.CreateDirectory(destinationDir);

        foreach (var file in Directory.GetFiles(sourceDir))
        {
            var fileName = Path.GetFileName(file);
            File.Copy(file, Path.Combine(destinationDir, fileName), overwrite: true);
        }
    }

    private static string FindRepoRoot()
    {
        var current = AppContext.BaseDirectory;
        var directory = new DirectoryInfo(current);
        while (directory is not null)
        {
            var sln = Path.Combine(directory.FullName, "VsIdeBridge.sln");
            if (File.Exists(sln))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new InvalidOperationException("Could not infer repo root. Pass --repo-root <path>.");
    }

    private static void EnsureAdmin()
    {
        using var identity = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);
        if (!principal.IsInRole(WindowsBuiltInRole.Administrator))
        {
            throw new UnauthorizedAccessException("Run this installer from an elevated terminal (Administrator).");
        }
    }

    private static bool IsHelp(string token)
    {
        return token is "-h" or "--help" or "/?";
    }

    private static Dictionary<string, string?> ParseOptions(string[] args)
    {
        var map = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (!arg.StartsWith("--", StringComparison.Ordinal))
            {
                throw new InvalidOperationException($"Unexpected argument: {arg}");
            }

            var key = arg[2..];
            if (i + 1 < args.Length && !args[i + 1].StartsWith("--", StringComparison.Ordinal))
            {
                map[key] = args[++i];
            }
            else
            {
                map[key] = null;
            }
        }

        return map;
    }

    private static string? GetOption(Dictionary<string, string?> options, string key)
    {
        return options.TryGetValue(key, out var value) ? value : null;
    }

    private static string? GetPathOption(Dictionary<string, string?> options, string key)
    {
        var value = GetOption(options, key);
        return string.IsNullOrWhiteSpace(value) ? null : Path.GetFullPath(value);
    }

    private static bool HasFlag(Dictionary<string, string?> options, string key)
    {
        return options.ContainsKey(key);
    }

    private static int Fail(string message)
    {
        Console.Error.WriteLine($"ERROR: {message}");
        return 1;
    }

    private static void PrintHelp()
    {
        Console.WriteLine("vs-ide-bridge-installer");
        Console.WriteLine();
        Console.WriteLine("Commands:");
        Console.WriteLine("  install   Install CLI binaries, VSIX, and optionally register a Windows service");
        Console.WriteLine("  uninstall Uninstall VSIX, remove service (optional), and delete install directory");
        Console.WriteLine();
        Console.WriteLine("Install options:");
        Console.WriteLine("  --repo-root <path>       Repo root (default: inferred)");
        Console.WriteLine("  --configuration <cfg>    Build config (Debug|Release, default: Debug)");
        Console.WriteLine("  --install-dir <path>     Install root (default: C:\\Program Files\\VsIdeBridge)");
        Console.WriteLine("  --cli-source <path>      CLI source folder override");
        Console.WriteLine("  --vsix-path <path>       VSIX file override");
        Console.WriteLine("  --vsix-id <id>           VSIX id (default: StanElston.VsIdeBridge)");
        Console.WriteLine("  --skip-vsix              Skip VSIX install/update");
        Console.WriteLine("  --install-service        Register Windows service");
        Console.WriteLine("  --service-name <name>    Service name (default: VsIdeBridgeHost)");
        Console.WriteLine("  --service-exe <path>     Absolute path to service executable (required with --install-service)");
        Console.WriteLine();
        Console.WriteLine("Uninstall options:");
        Console.WriteLine("  --install-dir <path>     Install root to remove");
        Console.WriteLine("  --vsix-id <id>           VSIX id to uninstall");
        Console.WriteLine("  --skip-vsix              Skip VSIX uninstall");
        Console.WriteLine("  --remove-service         Stop and delete service");
        Console.WriteLine("  --service-name <name>    Service name to remove");
    }
}
