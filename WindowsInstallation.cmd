@echo off
setlocal

set "SRC=%~dp0"
set "DRV=Drivers"
set "DRIVER_ARGS="
set "TIMEOUT=30"

echo =======================================================================
echo Windows Setup Automation Wrapper
echo Running from: %SRC%
echo =======================================================================
echo.

:: ----------------------------------------------------------------------------
::  Driver Installation (Optional)
:: ----------------------------------------------------------------------------
if not exist "%SRC%%DRV%" goto drivers_done

echo Found driver folder at: "%SRC%%DRV%"
echo.
echo Do you want to install drivers from the Drivers folder?
echo [N] No  - do not install drivers - DEFAULT
echo [Y] Yes - include drivers during setup
echo.

choice /C NY /T %TIMEOUT% /D N /M "Install drivers (Will default to No in %TIMEOUT% seconds):"

if errorlevel 2 set "DRIVER_ARGS=/InstallDrivers "%SRC%%DRV%""

:drivers_done

:: ----------------------------------------------------------------------------
::  Consolidate Common Setup.exe Arguments
:: ----------------------------------------------------------------------------
:: These arguments are identical for both Upgrade and Clean Install scenarios.
set "ARGS=/Eula Accept /DynamicUpdate Disable /Telemetry Disable"

if defined DRIVER_ARGS (
    echo.
    echo Drivers will be installed from: "%SRC%%DRV%"
) else (
    echo.
    echo Drivers will not be installed during setup.
)
echo.

:: ----------------------------------------------------------------------------
::  Select Installation Type (With Timeout)
:: ----------------------------------------------------------------------------
echo Select the type of installation you want to perform:
echo [U] Upgrade (Keeps files, settings, and apps) - DEFAULT
echo [C] Clean Install (WIPES THE DRIVE - Fresh OS installation)
echo.

choice /C UC /T %TIMEOUT% /D U /M "Enter your choice (Will default to Upgrade in %TIMEOUT% seconds):"

if errorlevel 2 goto confirm_clean
if errorlevel 1 goto do_upgrade
goto end

:: ----------------------------------------------------------------------------
::  Action: Upgrade Execution
:: ----------------------------------------------------------------------------
:do_upgrade
echo.
echo Launching Windows Setup for In-Place Upgrade...
"%SRC%setup.exe" /Auto Upgrade %ARGS% %DRIVER_ARGS%
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
goto end

:do_clean
echo.
echo Launching Windows Setup for Automated Clean Install...
"%SRC%setup.exe" /Auto Clean %ARGS% %DRIVER_ARGS%
goto end

:menu_cancel
echo.
echo Operations cancelled by user. Returning to command prompt...
pause
goto end

:end
endlocal
