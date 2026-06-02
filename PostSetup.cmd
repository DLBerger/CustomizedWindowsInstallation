@echo off
setlocal enabledelayedexpansion
set "SRC=%~dp0"

set "REGCMD=%SRC%\InstallRegs.cmd"
set "NETCMD=%SRC%\Update.NET.cmd"
set "EDGECMD=%SRC%\UpdateEdge.cmd"

:: Abort if required files are missing
for %%F in ("%REGCMD%" "%NETCMD%" %EDGECMD%) do (
    if not exist %%F (
        echo ERROR: Required file not found: %%F
        exit /b 1
    )
)

echo Importing registry files
call "%REGCMD%"

echo Updating .NET
call "%NETCMD%"

echo Updating Microsoft Edge
call "%EDGECMD%"

echo Updating with winget
winget update --all --silent

echo Updating Windows Store apps
store updates --apply

echo.
echo Updates complete
pause

:end
endlocal
