@echo off
setlocal enabledelayedexpansion

set "SRC=%~dp0"
set "FLD=$WinpeDriver$"
set "DRIVERDIR=%SRC%%FLD%"

if not exist "%DRIVERDIR%" (
    echo ERROR: Driver folder not found: "%DRIVERDIR%"
    exit /b 1
)

pnputil /add-driver "%DRIVERDIR%\*.inf" /subdirs /install > "%DRIVERDIR%\pnputil.log" 2>&1
endlocal
exit /b %ERRORLEVEL%