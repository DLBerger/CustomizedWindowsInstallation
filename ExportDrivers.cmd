@echo off
:: Must be run elevated to work
setlocal
set "SRC=%~dp0"
set "FLD=$WinpeDriver$"

:: Ensure the output folder exists and create if necessary
if not exist "%SRC%%FLD%" mkdir "%SRC%%FLD%"

echo Export drivers
pnputil /export-driver * "%SRC%%FLD%"
endlocal
