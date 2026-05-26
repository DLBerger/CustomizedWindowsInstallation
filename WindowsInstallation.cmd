@echo off
setlocal

set "SRC=%~dp0"
set "DRV=$WinPEDriver$"

:: ----------------------------------------------------------------------------
::  Validate Driver Folder Presence
:: ----------------------------------------------------------------------------
if not exist "%SRC%%DRV%" goto fail

:: ----------------------------------------------------------------------------
::  Consolidate Common Setup.exe Arguments
:: ----------------------------------------------------------------------------
:: These arguments are identical for both Upgrade and Clean Install scenarios.
set "ARGS=/Eula Accept /DynamicUpdate Disable /Telemetry Disable /InstallDrivers "%SRC%%DRV%""

echo =======================================================================
echo Windows Setup Automation Wrapper
echo Running from: %SRC%
echo Found driver folder at: "%SRC%%DRV%"
echo =======================================================================
echo.

:: ----------------------------------------------------------------------------
::  Select Installation Type (With Timeout)
:: ----------------------------------------------------------------------------
echo Select the type of installation you want to perform:
echo [U] Upgrade (Keeps files, settings, and apps) - DEFAULT
echo [C] Clean Install (WIPES THE DRIVE - Fresh OS installation)
echo.
set "TIMEOUT=30"

choice /C UC /T %TIMEOUT% /D U /M "Enter your choice (Will default to Upgrade in %TIMEOUT% seconds):"

if errorlevel 2 goto confirm_clean
if errorlevel 1 goto do_upgrade
goto fail

:: ----------------------------------------------------------------------------
::  Action: Upgrade Execution
:: ----------------------------------------------------------------------------
:do_upgrade
echo.
echo Launching Windows Setup for In-Place Upgrade...
"%SRC%setup.exe" /Auto Upgrade %ARGS%
goto end

:: ----------------------------------------------------------------------------
::  Action: Clean Install Safety Gate (No Timeout)
:: ----------------------------------------------------------------------------
:confirm_clean
echo.
echo ***********************************************************************
echo                             !!! WARNING !!!
echo ***********************************************************************
echo You have selected "Clean Install". 
echo THIS WILL COMPLETELY WIPE DISK 0! All personal files, applications,
echo and settings will be permanently destroyed.
echo ***********************************************************************
echo.

choice /C YN /M "Are you absolutely sure you want to completely clear this PC and install fresh?"

if errorlevel 2 goto menu_cancel
if errorlevel 1 goto do_clean
goto fail

:do_clean
echo.
echo Launching Windows Setup for Automated Clean Install...
"%SRC%setup.exe" /Auto Clean %ARGS%
goto end

:menu_cancel
echo.
echo Operations cancelled by user. Returning to command prompt...
pause
goto end

:: ----------------------------------------------------------------------------
::  Fail Handler
:: ----------------------------------------------------------------------------
:fail
echo.
echo =======================================================================
echo ERROR: The '%DRV%' folder was not found in:
echo "%SRC%"
echo The process has been halted to prevent missing hardware drivers.
echo =======================================================================
echo.
pause
exit /b 1

:end
endlocal