@echo off
:: Must be run elevated to work
setlocal
set "SRC=%~dp0"
set "FLD=$WinpeDriver$"

echo Importing drivers using DISM...
dism /online /Add-Driver /Driver:"%SRC%%FLD%" /Recurse /ForceUnsigned /logpath:"%SRC%%FLD%\dism-import.log"

endlocal
