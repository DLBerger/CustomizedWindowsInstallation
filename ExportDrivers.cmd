@echo off
:: Must be run elevated to work
setlocal
set "SRC=%~dp0"
set "FLD=Drivers"
set "DEST=%SRC%%FLD%"

:: Ensure the output folder exists
if not exist "%DEST%" (
    echo [INFO] Creating destination folder...
    mkdir "%DEST%"
)

echo Windows Driver Export Tool
echo [START] Exporting drivers to: "%DEST%"
echo [NOTE] This may take a few minutes. Please wait...
echo ----------------------------------------------------

:: Run DISM and stream the live output directly to the console
dism /online /export-driver /destination:"%DEST%" /logpath:"%DEST%\dism.log"

echo ----------------------------------------------------
if %ERRORLEVEL% EQU 0 (
    echo [INFO] Log saved to: "%DEST%\dism.log"
) else (
    echo [ERROR] DISM failed with exit code %ERRORLEVEL%.
    echo [INFO] Check the log at "%DEST%\dism.log" for details.
)

pause
endlocal
