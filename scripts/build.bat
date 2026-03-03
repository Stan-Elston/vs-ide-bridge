@echo off
setlocal

set "CONFIG=%~1"
if "%CONFIG%"=="" set "CONFIG=Debug"

set "ROOT=%~dp0.."
set "SOLUTION=%ROOT%\VsIdeBridge.sln"
set "MSBUILD=%ProgramFiles%\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"
if not exist "%MSBUILD%" set "MSBUILD=%ProgramFiles%\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"

if not exist "%MSBUILD%" (
    echo MSBuild not found: "%MSBUILD%"
    exit /b 1
)

"%MSBUILD%" "%SOLUTION%" /restore /m /p:Configuration=%CONFIG%
exit /b %ERRORLEVEL%
