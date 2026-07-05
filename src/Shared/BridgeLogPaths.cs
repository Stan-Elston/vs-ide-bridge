using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace VsIdeBridge.Shared;

public static class BridgeLogPaths
{
    private const string ProductDirectoryName = "VsIdeBridge";
    private const string LogsDirectoryName = "logs";
    private const string McpServerLogFileName = "mcp-server.log";
    private const string RegistryKeyPath = @"SOFTWARE\VsIdeBridge";
    private const string RegistryInstallPathValue = "InstallPath";
    private const string ExtensionLogPrefix = "vs-ide-bridge-";

    public static string GetSharedLogDirectory()
    {
        string? installPath = TryGetRegistryInstallPath();
        if (!string.IsNullOrWhiteSpace(installPath))
            return Path.Combine(installPath, LogsDirectoryName);

        string commonAppData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
        if (!string.IsNullOrWhiteSpace(commonAppData))
            return Path.Combine(commonAppData, ProductDirectoryName, LogsDirectoryName);

        return GetTempLogDirectory();
    }

    public static string? GetInstallPath()
    {
        return TryGetRegistryInstallPath();
    }

    private static string? TryGetRegistryInstallPath()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            return null;
        try
        {
            using RegistryKey? key = Registry.LocalMachine.OpenSubKey(RegistryKeyPath);
            return key?.GetValue(RegistryInstallPathValue) as string;
        }
        catch
        {
            return null;
        }
    }

    public static string GetTempLogDirectory()
    {
        return Path.Combine(Path.GetTempPath(), ProductDirectoryName, LogsDirectoryName);
    }

    public static string GetMcpServerLogPath()
    {
        return Path.Combine(GetSharedLogDirectory(), McpServerLogFileName);
    }

    public static string GetMcpServerTempLogPath()
    {
        return Path.Combine(GetTempLogDirectory(), McpServerLogFileName);
    }

    public static string GetVisualStudioExtensionLogPath()
    {
        return GetVisualStudioExtensionLogPath(DateTime.Now);
    }

    public static string GetVisualStudioExtensionLogPath(DateTime timestamp)
    {
        return Path.Combine(GetSharedLogDirectory(), $"vs-ide-bridge-{timestamp:yyyy-MM-dd}.log");
    }

    /// <summary>
    /// Removes log files from the CommonAppData fallback directory when the active log
    /// directory is the installer-managed location. Prevents stale pre-install logs from
    /// accumulating alongside current logs in a different folder.
    /// Called once per process at startup. Failures are silently swallowed.
    /// </summary>
    public static void CleanupLegacyLogs(string currentLogDir)
    {
        try
        {
            string commonAppData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            if (string.IsNullOrWhiteSpace(commonAppData))
                return;
            string legacyDir = Path.Combine(commonAppData, ProductDirectoryName, LogsDirectoryName);
            if (string.Equals(legacyDir, currentLogDir, StringComparison.OrdinalIgnoreCase))
                return;  // already writing to this location - nothing to clean
            if (!Directory.Exists(legacyDir))
                return;
            foreach (string file in Directory.GetFiles(legacyDir, "vs-ide-bridge-*.log"))
                File.Delete(file);
            foreach (string file in Directory.GetFiles(legacyDir, "mcp-server.log*"))
                File.Delete(file);
        }
        catch (IOException ex) { Debug.WriteLine($"BridgeLogPaths.CleanupLegacyLogs failed: {ex.Message}"); }
        catch (UnauthorizedAccessException ex) { Debug.WriteLine($"BridgeLogPaths.CleanupLegacyLogs failed: {ex.Message}"); }
    }

    /// <summary>
    /// Deletes daily extension log files older than <paramref name="keepDays"/> days.
    /// Called once per process on first log write. Failures are silently swallowed.
    /// </summary>
    public static void PruneExtensionLogs(string logDir, int keepDays = 7)
    {
        try
        {
            if (!Directory.Exists(logDir))
                return;
            DateTime cutoff = DateTime.Today.AddDays(-keepDays);
            foreach (string file in Directory.GetFiles(logDir, "vs-ide-bridge-*.log"))
            {
                string stem = Path.GetFileNameWithoutExtension(file);
                if (stem.Length <= ExtensionLogPrefix.Length)
                    continue;
                if (DateTime.TryParseExact(
#if NETFRAMEWORK
                        stem.Substring(ExtensionLogPrefix.Length),
#else
                        stem[ExtensionLogPrefix.Length..],
#endif
                        "yyyy-MM-dd",
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        out DateTime fileDate)
                    && fileDate < cutoff)
                {
                    File.Delete(file);
                }
            }
        }
        catch (IOException ex) { Debug.WriteLine($"BridgeLogPaths.PruneExtensionLogs failed: {ex.Message}"); }
        catch (UnauthorizedAccessException ex) { Debug.WriteLine($"BridgeLogPaths.PruneExtensionLogs failed: {ex.Message}"); }
    }

    /// <summary>
    /// If <paramref name="logPath"/> exceeds <paramref name="maxBytes"/>, renames it to
    /// <c>.old</c> (replacing any previous backup) so the next write starts a fresh file.
    /// Called once per process at startup. Failures are silently swallowed.
    /// </summary>
    public static void RotateMcpServerLog(string logPath, long maxBytes = 5 * 1024 * 1024)
    {
        try
        {
            if (!File.Exists(logPath))
                return;
            if (new FileInfo(logPath).Length < maxBytes)
                return;
            string oldPath = logPath + ".old";
            if (File.Exists(oldPath))
                File.Delete(oldPath);
            File.Move(logPath, oldPath);
        }
        catch (IOException ex) { Debug.WriteLine($"BridgeLogPaths.RotateMcpServerLog failed: {ex.Message}"); }
        catch (UnauthorizedAccessException ex) { Debug.WriteLine($"BridgeLogPaths.RotateMcpServerLog failed: {ex.Message}"); }
    }
}
