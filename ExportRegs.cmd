@echo off
setlocal ENABLEDELAYEDEXPANSION

:: Get full path to this .cmd file
set "SELF=%~f0"

:: Derive matching .ps1 file
set "PS1=%SELF:.cmd=.ps1%"

if not exist "%PS1%" (
    echo ERROR: PowerShell script not found: %PS1%
    exit /b 1
)

:: Pass all arguments exactly as received
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %*

:: Let us see what happened
pause

endlocal
exit /b %ERRORLEVEL%
