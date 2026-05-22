@echo off
:: Must be run elevated to work
setlocal
set "SRC=%~dp0"
set "FLD=$WinpeDriver$"

:: Ensure the output folder exists and create if necessary
if not exist "%SRC%%FLD%" mkdir "%SRC%%FLD%"

echo Exporting drivers using DISM...
dism /online /export-driver /destination:"%SRC%%FLD%" /logpath:"%SRC%%FLD%\dism-export.log"

endlocal
