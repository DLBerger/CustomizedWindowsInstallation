@echo off
setlocal enabledelayedexpansion

set "SRC=%~dp0"
set "FLD=Drivers"
set "DRIVERDIR=%SRC%%FLD%"

if not exist "%DRIVERDIR%" (
    echo ERROR: Driver folder not found: "%DRIVERDIR%"
    exit /b 1
)

pnputil /add-driver "%DRIVERDIR%\*.inf" /subdirs /install
endlocal
exit /b %ERRORLEVEL%
