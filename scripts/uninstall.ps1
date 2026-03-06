param(
    [string]$InstallRoot = "C:\Program Files\VsIdeBridge",
    [string]$ServiceName = "VsIdeBridgeService",
    [string]$VsixId = "StanElston.VsIdeBridge"
)

$ErrorActionPreference = "Stop"

function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($id)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
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

sc.exe stop $ServiceName | Out-Null
Start-Sleep -Milliseconds 300
sc.exe delete $ServiceName | Out-Null

$vsixInstaller = Resolve-VsixInstaller
$proc = Start-Process -FilePath $vsixInstaller -ArgumentList @('/quiet', "/uninstall:$VsixId") -Wait -PassThru
if ($proc.ExitCode -ne 0) {
    Write-Warning "VSIX uninstall returned exit code $($proc.ExitCode)."
}

if (Test-Path $InstallRoot) {
    Remove-Item -Path $InstallRoot -Recurse -Force
}

Write-Host "Uninstall complete."
Write-Host "Removed service: $ServiceName"
Write-Host "Removed install root: $InstallRoot"
Write-Host "Attempted VSIX uninstall: $VsixId"
