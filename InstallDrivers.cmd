@echo off
setlocal ENABLEDELAYEDEXPANSION

:: -----------------------------------------
:: Locate OS drive safely (no WMIC, no assumptions)
:: -----------------------------------------
for %%D in (C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
    if exist "%%D:\Windows\System32\config\SYSTEM" (
        set "OSDRIVE=%%D:"
        goto :foundOS
    )
)

echo ERROR: Unable to locate OS drive.
exit /b 1

:foundOS

:: -----------------------------------------
:: Driver folder
:: -----------------------------------------
set "SRC=%~dp0"
set "FLD=$WinpeDriver$"
set "DRIVERDIR=%SRC%%FLD%"

if not exist "%DRIVERDIR%" (
    echo ERROR: Driver folder not found: "%DRIVERDIR%"
    exit /b 1
)

:: -----------------------------------------
:: Detect whether system is fully online
:: -----------------------------------------
:: SetupComplete runs before explorer.exe exists
tasklist /FI "IMAGENAME eq explorer.exe" >nul 2>&1
if %ERRORLEVEL%==0 (
    set "ONLINE=1"
) else (
    set "ONLINE=0"
)

:: -----------------------------------------
:: Install drivers
:: -----------------------------------------
if "%ONLINE%"=="1" (
    echo System is fully booted. Using pnputil...
    pnputil /add-driver "%DRIVERDIR%\*.inf" /subdirs /install > "%DRIVERDIR%\pnputil-install.log" 2>&1
) else (
    echo System is in SetupComplete/OOBE. Using DISM offline servicing...
    dism /Image:"%OSDRIVE%\" /Add-Driver /Driver:"%DRIVERDIR%" /Recurse /ForceUnsigned /logpath:"%DRIVERDIR%\dism-import.log"
)

endlocal
exit /b 0
