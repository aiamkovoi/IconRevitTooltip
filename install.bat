@echo off
setlocal

:: ============================================================
:: IconRevitTooltip — Install Script
:: Run as Administrator!
:: ============================================================

:: --- Determine if pre-built binaries exist in dist/ ---
set "USE_PREBUILT=0"
if exist "%~dp0dist\IconRevitTooltip.ShellExtension.dll" (
    if exist "%~dp0dist\IconRevitTooltip.Core.dll" (
        if exist "%~dp0dist\OpenMcdf.dll" (
            if exist "%~dp0dist\SharpShell.dll" (
                set "USE_PREBUILT=1"
            )
        )
    )
)

if "%USE_PREBUILT%"=="1" (
    echo Pre-built binaries found in dist\ folder. Skipping compilation.
    goto install_prebuilt
)

:: ============================================================
:: BUILD FROM SOURCE
:: ============================================================
echo No pre-built binaries found. Building from source...
echo.

:: --- Download nuget.exe if missing ---
if not exist "%~dp0nuget.exe" (
    echo Downloading nuget.exe...
    powershell -Command "Invoke-WebRequest 'https://dist.nuget.org/win-x86-commandline/latest/nuget.exe' -OutFile '%~dp0nuget.exe'"
)

echo Restoring NuGet packages...
"%~dp0nuget.exe" restore "%~dp0IconRevitTooltip.slnx" -PackagesDirectory "%~dp0packages"
"%~dp0nuget.exe" install Microsoft.NETFramework.ReferenceAssemblies.net47 -Version 1.0.3 -OutputDirectory "%~dp0packages"
"%~dp0nuget.exe" install ServerRegistrationManager -Version 2.7.2 -OutputDirectory "%~dp0packages"

:: --- Find MSBuild via vswhere ---
echo Searching for MSBuild...
set "MSBUILD_PATH="

:: Try vswhere (ships with VS 2017+ and Build Tools)
set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
if exist "%VSWHERE%" (
    for /f "usebackq tokens=*" %%i in (`"%VSWHERE%" -latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe 2^>nul`) do (
        set "MSBUILD_PATH=%%i"
    )
)

:: Fallback: .NET Framework built-in MSBuild
if "%MSBUILD_PATH%"=="" (
    if exist "%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe" (
        set "MSBUILD_PATH=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe"
        echo Using .NET Framework MSBuild fallback.
    )
)

if "%MSBUILD_PATH%"=="" (
    echo ERROR: MSBuild not found!
    echo Install Visual Studio or Visual Studio Build Tools, or use pre-built binaries in dist\.
    pause
    exit /b 1
)

echo Found MSBuild: %MSBUILD_PATH%

echo Terminating Windows Explorer to release locks on Shell Extension DLLs...
taskkill /f /im explorer.exe 2>nul
timeout /t 1 /nobreak >nul

echo Building IconRevitTooltip...
set "FrameworkPathOverride=%~dp0packages\Microsoft.NETFramework.ReferenceAssemblies.net47.1.0.3\build\.NETFramework\v4.7"

"%MSBUILD_PATH%" "%~dp0IconRevitTooltip.slnx" /p:Configuration=Release /p:TargetFrameworkVersion=v4.7 /p:FrameworkPathOverride="%FrameworkPathOverride%" /v:minimal

if errorlevel 1 (
    echo ERROR: Build failed!
    start explorer.exe
    pause
    exit /b 1
)

echo Copying transient dependencies to ShellExtension folder...
xcopy /y "%~dp0IconRevitTooltip\bin\Release\*.dll" "%~dp0IconRevitTooltip.ShellExtension\bin\Release\" >nul

:: Set paths for registration
set "SHELL_DLL=%~dp0IconRevitTooltip.ShellExtension\bin\Release\IconRevitTooltip.ShellExtension.dll"
set "CORE_DLL=%~dp0IconRevitTooltip.Core\bin\Release\IconRevitTooltip.Core.dll"
set "SHARPSHELL_DLL=%~dp0packages\SharpShell.2.7.2\lib\net40-client\SharpShell.dll"
set "OPENMCDF_DLL=%~dp0packages\OpenMcdf.2.4.1\lib\net40\OpenMcdf.dll"

goto do_install

:: ============================================================
:: INSTALL FROM PRE-BUILT BINARIES
:: ============================================================
:install_prebuilt
echo.

:: Download nuget only for SRM
if not exist "%~dp0nuget.exe" (
    echo Downloading nuget.exe for SRM...
    powershell -Command "Invoke-WebRequest 'https://dist.nuget.org/win-x86-commandline/latest/nuget.exe' -OutFile '%~dp0nuget.exe'"
)
"%~dp0nuget.exe" install ServerRegistrationManager -Version 2.7.2 -OutputDirectory "%~dp0packages"

echo Terminating Windows Explorer to release locks...
taskkill /f /im explorer.exe 2>nul
timeout /t 1 /nobreak >nul

:: Set paths for registration (from dist/)
set "SHELL_DLL=%~dp0dist\IconRevitTooltip.ShellExtension.dll"
set "CORE_DLL=%~dp0dist\IconRevitTooltip.Core.dll"
set "SHARPSHELL_DLL=%~dp0dist\SharpShell.dll"
set "OPENMCDF_DLL=%~dp0dist\OpenMcdf.dll"

:: ============================================================
:: COMMON REGISTRATION
:: ============================================================
:do_install
echo.
echo Installing the Shell Extension...
echo Make sure to run this script as Administrator!

set "SRM_PATH="
for /f "delims=" %%A in ('dir /s /b "%~dp0packages\ServerRegistrationManager.exe" 2^>nul') do (
    set "SRM_PATH=%%A"
    goto found
)
:found

if "%SRM_PATH%"=="" (
    echo ERROR: srm.exe not found!
    start explorer.exe
    pause
    exit /b 1
)

echo Installing assemblies into Global Assembly Cache (GAC)...
powershell -Command "[System.Reflection.Assembly]::Load('System.EnterpriseServices, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a') | Out-Null; $p = New-Object System.EnterpriseServices.Internal.Publish; $p.GacInstall('%SHARPSHELL_DLL%'); $p.GacInstall('%CORE_DLL%'); $p.GacInstall('%OPENMCDF_DLL%')"

"%SRM_PATH%" uninstall "%SHELL_DLL%" -os64 2>nul
"%SRM_PATH%" install "%SHELL_DLL%" -codebase -os64

:: --- Override Revit class-level IQueryInfo handlers ---
:: Read the CLSID that SharpShell registered for our handler
set "OUR_CLSID="
for /f "tokens=3" %%G in ('reg query "HKLM\Software\Classes\.rvt\shellex\{00021500-0000-0000-C000-000000000046}" /ve 2^>nul ^| findstr REG_SZ') do set "OUR_CLSID=%%G"

if "%OUR_CLSID%"=="" (
    echo WARNING: Could not determine our CLSID. Tooltip may not override Revit defaults.
    goto skip_class_override
)

echo.
echo Overriding Revit class-level tooltip handlers with ours...

:: Backup and override Revit.Project
for /f "tokens=3" %%G in ('reg query "HKLM\Software\Classes\Revit.Project\ShellEx\{00021500-0000-0000-C000-000000000046}" /ve 2^>nul ^| findstr REG_SZ') do (
    if not "%%G"=="%OUR_CLSID%" (
        reg add "HKLM\Software\Classes\Revit.Project\ShellEx\{00021500-0000-0000-C000-000000000046}" /v "OriginalHandler" /d "%%G" /f >nul
    )
)
reg add "HKLM\Software\Classes\Revit.Project\ShellEx\{00021500-0000-0000-C000-000000000046}" /ve /d "%OUR_CLSID%" /f >nul

:: Backup and override Revit.Family
for /f "tokens=3" %%G in ('reg query "HKLM\Software\Classes\Revit.Family\ShellEx\{00021500-0000-0000-C000-000000000046}" /ve 2^>nul ^| findstr REG_SZ') do (
    if not "%%G"=="%OUR_CLSID%" (
        reg add "HKLM\Software\Classes\Revit.Family\ShellEx\{00021500-0000-0000-C000-000000000046}" /v "OriginalHandler" /d "%%G" /f >nul
    )
)
reg add "HKLM\Software\Classes\Revit.Family\ShellEx\{00021500-0000-0000-C000-000000000046}" /ve /d "%OUR_CLSID%" /f >nul

echo Class-level handlers overridden successfully.

:skip_class_override
echo Requesting Explorer Restart...
taskkill /f /im explorer.exe 2>nul
start explorer.exe
echo.
echo Done! Hover over a .rvt or .rfa file to test.
pause
