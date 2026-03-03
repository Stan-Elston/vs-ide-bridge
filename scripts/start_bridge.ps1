param(
    [Parameter(Mandatory = $true)]
    [string]$SolutionPath,

    [string]$Configuration = "Debug",

    [int]$StartupTimeoutSeconds = 180,

    [int]$PollIntervalMilliseconds = 1000,

    [switch]$SkipWaitForReady,

    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

function Resolve-BridgeCliPath {
    param(
        [string]$RepoRoot,
        [string]$ConfigurationName
    )

    $preferred = Join-Path $RepoRoot "src\VsIdeBridgeCli\bin\$ConfigurationName\net8.0\vs-ide-bridge.exe"
    if (Test-Path -LiteralPath $preferred) {
        return $preferred
    }

    $fallbacks = @(
        (Join-Path $RepoRoot "src\VsIdeBridgeCli\bin\Debug\net8.0\vs-ide-bridge.exe"),
        (Join-Path $RepoRoot "src\VsIdeBridgeCli\bin\Release\net8.0\vs-ide-bridge.exe")
    )

    foreach ($candidate in $fallbacks) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    throw "Bridge CLI not found. Build the solution first with scripts\build.bat."
}

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$solutionFullPath = [System.IO.Path]::GetFullPath($SolutionPath)
if (-not (Test-Path -LiteralPath $solutionFullPath)) {
    throw "Solution not found: $solutionFullPath"
}

$bridgeCli = Resolve-BridgeCliPath -RepoRoot $repoRoot -ConfigurationName $Configuration

$arguments = @(
    "ensure",
    "--solution", $solutionFullPath,
    "--timeout-ms", ([Math]::Max(1, $StartupTimeoutSeconds) * 1000),
    "--poll-ms", ([Math]::Max(100, $PollIntervalMilliseconds)),
    "--format", "summary"
)

if ($SkipWaitForReady) {
    $arguments += "--skip-ready"
}

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $arguments += @("--out", ([System.IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables($OutputPath))))
}

& $bridgeCli @arguments
exit $LASTEXITCODE
