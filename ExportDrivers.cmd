@echo off
:: Must be run elevated to work
setlocal
set "SRC=%~dp0"
set "FLD=$WinpeDriver$"

echo Import drivers
pnputil /export-driver * "%SRC%%FLD%"
endlocal
