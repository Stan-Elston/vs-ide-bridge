@echo off
setlocal

set "CONFIG=%~1"
if "%CONFIG%"=="" set "CONFIG=Debug"

set "ROOT=%~dp0.."
set "SOLUTION=%ROOT%\VsIdeBridge.sln"
set "VSIX_PROJECT=%ROOT%\src\VsIdeBridge\VsIdeBridge.csproj"
set "VSIX_BIN=%ROOT%\src\VsIdeBridge\bin\%CONFIG%\net472"
set "VSIX_OBJ=%ROOT%\src\VsIdeBridge\obj\%CONFIG%\net472"

:: Locate MSBuild via vswhere (works for Community, Professional, Enterprise, Preview)
set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
if not exist "%VSWHERE%" set "VSWHERE=%ProgramFiles%\Microsoft Visual Studio\Installer\vswhere.exe"
:: Literal fallback: ProgramFiles(x86) is not defined in some host environments
:: (e.g. processes spawned by the bridge service with a stripped environment).
if not exist "%VSWHERE%" set "VSWHERE=C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe"

if not exist "%VSWHERE%" (
    echo ERROR: vswhere.exe not found. Install Visual Studio first.
    exit /b 1
)

for /f "usebackq delims=" %%I in (`"%VSWHERE%" -latest -prerelease -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe 2^>nul`) do (
    set "MSBUILD=%%I"
)

if not defined MSBUILD (
    echo ERROR: MSBuild not found via vswhere.
    exit /b 1
)

:: Derive DevEnvDir so HintPath references (e.g. CallHierarchy) resolve correctly
:: when building outside a VS Developer Command Prompt.
for /f "usebackq delims=" %%I in (`"%VSWHERE%" -latest -prerelease -property installationPath 2^>nul`) do (
    set "VS_INSTALL_DIR=%%I"
)
if defined VS_INSTALL_DIR (
    set "DevEnvDir=%VS_INSTALL_DIR%\Common7\IDE\"
)

echo MSBuild : %MSBUILD%
echo Solution: %SOLUTION%
echo Config  : %CONFIG%
if defined DevEnvDir echo DevEnvDir: %DevEnvDir%
echo.

if exist "%VSIX_BIN%" (
    echo Cleaning VSIX output cache: %VSIX_BIN%
    rmdir /s /q "%VSIX_BIN%"
)

if exist "%VSIX_OBJ%" (
    echo Cleaning VSIX intermediate cache: %VSIX_OBJ%
    rmdir /s /q "%VSIX_OBJ%"
)

"%MSBUILD%" "%SOLUTION%" /t:Restore;Rebuild /m /p:Configuration=%CONFIG%
if %ERRORLEVEL% neq 0 exit /b %ERRORLEVEL%

echo Forcing standalone VSIX rebuild: %VSIX_PROJECT%
"%MSBUILD%" "%VSIX_PROJECT%" /t:Rebuild /m /p:Configuration=%CONFIG%
exit /b %ERRORLEVEL%
