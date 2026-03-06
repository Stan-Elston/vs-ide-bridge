param(
    [string]$Configuration = "Release",
    [string]$InstallRoot = "C:\Program Files\VsIdeBridge",
    [string]$ServiceName = "VsIdeBridgeService",
    [int]$IdleSoftSeconds = 900,
    [int]$IdleHardSeconds = 1200,
    [string]$VsixId = "StanElston.VsIdeBridge"
)

$ErrorActionPreference = "Stop"

function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($id)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Resolve-RepoRoot {
    return [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
}

function Resolve-VsixInstaller {
    $base = "C:\Program Files\Microsoft Visual Studio\18"
    $editions = @("Enterprise", "Professional", "Community", "Preview")
    foreach ($edition in $editions) {
        $candidate = Join-Path $base "$edition\Common7\IDE\VSIXInstaller.exe"
        if (Test-Path $candidate) { return $candidate }
    }

    $fallback = Get-ChildItem -Path $base -Filter VSIXInstaller.exe -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($fallback) { return $fallback.FullName }
    throw "VSIXInstaller.exe not found under $base"
}

if (-not (Test-Admin)) {
    throw "Run this script from an elevated PowerShell session."
}

$repoRoot = Resolve-RepoRoot
$cliSource = Join-Path $repoRoot "src\VsIdeBridgeCli\bin\$Configuration\net8.0"
$serviceSource = Join-Path $repoRoot "src\VsIdeBridgeService\bin\$Configuration\net8.0-windows"
$vsixPath = Join-Path $repoRoot "src\VsIdeBridge\bin\$Configuration\net472\VsIdeBridge.vsix"

if (-not (Test-Path $cliSource)) { throw "CLI build output not found: $cliSource" }
if (-not (Test-Path $serviceSource)) { throw "Service build output not found: $serviceSource" }
if (-not (Test-Path $vsixPath)) { throw "VSIX not found: $vsixPath" }

$cliDest = Join-Path $InstallRoot "cli"
$serviceDest = Join-Path $InstallRoot "service"
New-Item -ItemType Directory -Force -Path $cliDest, $serviceDest | Out-Null

Copy-Item -Path (Join-Path $cliSource "*") -Destination $cliDest -Recurse -Force
Copy-Item -Path (Join-Path $serviceSource "*") -Destination $serviceDest -Recurse -Force

$serviceExe = Join-Path $serviceDest "VsIdeBridgeService.exe"
if (-not (Test-Path $serviceExe)) {
    throw "Service executable not found after copy: $serviceExe"
}

sc.exe stop $ServiceName | Out-Null
Start-Sleep -Milliseconds 300
sc.exe delete $ServiceName | Out-Null
Start-Sleep -Milliseconds 300

$binPath = "\"$serviceExe\" --idle-soft-seconds $IdleSoftSeconds --idle-hard-seconds $IdleHardSeconds"
sc.exe create $ServiceName binPath= $binPath start= demand DisplayName= "VS IDE Bridge Service" | Out-Null
sc.exe description $ServiceName "VS IDE Bridge service host with idle auto-shutdown." | Out-Null

$vsixInstaller = Resolve-VsixInstaller
$proc = Start-Process -FilePath $vsixInstaller -ArgumentList @('/quiet', $vsixPath) -Wait -PassThru
if ($proc.ExitCode -ne 0) {
    throw "VSIX install failed with exit code $($proc.ExitCode)."
}

Write-Host "Install complete."
Write-Host "Install root: $InstallRoot"
Write-Host "Service name: $ServiceName (manual start)"
Write-Host "VSIX id: $VsixId"
Write-Host "Start service: sc.exe start $ServiceName"
Write-Host "Service log: C:\ProgramData\VsIdeBridge\service.log"
