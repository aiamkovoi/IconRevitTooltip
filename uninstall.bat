@echo off
setlocal

:: ============================================================
:: IconRevitTooltip — Uninstall Script
:: Run as Administrator!
:: ============================================================

echo Uninstalling IconRevitTooltip Shell Extension...
echo Make sure to run this script as Administrator!
echo.

:: --- Find SRM ---
set "SRM_PATH="
for /f "delims=" %%A in ('dir /s /b "%~dp0packages\ServerRegistrationManager.exe" 2^>nul') do (
    set "SRM_PATH=%%A"
    goto found
)
:found

if "%SRM_PATH%"=="" (
    echo srm.exe not found. It was probably never compiled or installed.
    pause
    exit /b 1
)

:: --- Restore original Revit handlers ---
echo Restoring original Revit tooltip handlers...

for /f "tokens=3" %%G in ('reg query "HKLM\Software\Classes\Revit.Project\ShellEx\{00021500-0000-0000-C000-000000000046}" /v "OriginalHandler" 2^>nul ^| findstr REG_SZ') do (
    echo Restoring Revit.Project handler to: %%G
    reg add "HKLM\Software\Classes\Revit.Project\ShellEx\{00021500-0000-0000-C000-000000000046}" /ve /d "%%G" /f >nul
    reg delete "HKLM\Software\Classes\Revit.Project\ShellEx\{00021500-0000-0000-C000-000000000046}" /v "OriginalHandler" /f >nul
)

for /f "tokens=3" %%G in ('reg query "HKLM\Software\Classes\Revit.Family\ShellEx\{00021500-0000-0000-C000-000000000046}" /v "OriginalHandler" 2^>nul ^| findstr REG_SZ') do (
    echo Restoring Revit.Family handler to: %%G
    reg add "HKLM\Software\Classes\Revit.Family\ShellEx\{00021500-0000-0000-C000-000000000046}" /ve /d "%%G" /f >nul
    reg delete "HKLM\Software\Classes\Revit.Family\ShellEx\{00021500-0000-0000-C000-000000000046}" /v "OriginalHandler" /f >nul
)

:: --- Unregister and clean GAC ---
echo Removing assemblies from GAC...
powershell -Command "[System.Reflection.Assembly]::Load('System.EnterpriseServices, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a') | Out-Null; $p = New-Object System.EnterpriseServices.Internal.Publish; $p.GacRemove('SharpShell.dll'); $p.GacRemove('IconRevitTooltip.Core.dll'); $p.GacRemove('OpenMcdf.dll')"

:: Unregister from both possible locations
"%SRM_PATH%" uninstall "%~dp0IconRevitTooltip.ShellExtension\bin\Release\IconRevitTooltip.ShellExtension.dll" -os64 2>nul
"%SRM_PATH%" uninstall "%~dp0dist\IconRevitTooltip.ShellExtension.dll" -os64 2>nul

echo Requesting Explorer Restart...
taskkill /f /im explorer.exe
start explorer.exe
echo.
echo Done! Shell extension removed.
pause
