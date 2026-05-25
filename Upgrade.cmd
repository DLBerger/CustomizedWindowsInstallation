@echo off
setlocal

set "SRC=%~dp0"
set "DRV=$WinPEDriver$"
echo Running in-place upgrade from: %SRC%

:: ----------------------------------------------------------------------------
::  Validate Driver Folder Presence
:: ----------------------------------------------------------------------------
if not exist "%SRC%%DRV%" goto fail

echo Found driver folder at: "%SRC%%DRV%"
echo Launching Windows Setup...

"%SRC%setup.exe" /auto upgrade /eula accept /configfile "%SRC%SetupConfig-Upgrade.ini" /InstallDrivers "%SRC%%DRV%"
goto end

:: ----------------------------------------------------------------------------
::  Fail Handler
:: ----------------------------------------------------------------------------
:fail
echo.
echo =======================================================================
echo ERROR: The '%DRV%' folder was not found in:
echo "%SRC%"
echo The upgrade process has been halted to prevent missing hardware drivers.
echo =======================================================================
echo.
pause
exit /b 1

:end
endlocal