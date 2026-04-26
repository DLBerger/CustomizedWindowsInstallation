@echo off
setlocal enabledelayedexpansion

:: Get the full path of this .cmd file
set "SELF=%~f0"

:: Strip extension and append .ps1
set "PS1=%SELF:.cmd=.ps1%"

if not exist "%PS1%" (
    echo PowerShell script not found: %PS1%
    pause
    exit /b 1
)

:: Build argument string to pass through (no leading spaces)
set "ARGS="
:loop
if "%~1"=="" goto afterargs
if defined ARGS (
    set "ARGS=%ARGS% %~1"
) else (
    set "ARGS=%~1"
)
shift
goto loop

:afterargs

:: Run PowerShell with execution policy bypass
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %ARGS%

endlocal
exit /b %ERRORLEVEL%

