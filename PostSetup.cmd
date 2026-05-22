@echo off
setlocal enabledelayedexpansion
set "SRC=%~dp0"

set "REGCMD=%SRC%\InstallRegs.cmd"
set "DRVCMD=%SRC%\InstallDrivers.cmd"

:: Abort if required files are missing
for %%F in ("%REGCMD%" "%DRVCMD%") do (
    if not exist %%F (
        echo ERROR: Required file not found: %%F
        exit /b 1
    )
)

echo Importing registry files
call "%REGCMD%"

echo Installing drivers
call "%DRVCMD%"

goto end

:: Apply KBs in the correct order

echo Installing updates from %SRC%\KBs\NET

:: Install EXE installers
for %%F in ("%SRC%\KBs\NET\*.exe") do (
    echo Installing EXE %%F
    "%%F" /quiet /norestart
)

:: Install MSI installers
for %%F in ("%SRC%\KBs\NET\*.msi") do (
    echo Installing MSI %%F
    msiexec.exe /i "%%F" /quiet /norestart
)

:: Run CMD/BAT scripts
for %%F in ("%SRC%\KBs\NET\*.cmd") do (
    echo Running CMD %%F
    call "%%F"
)
for %%F in ("%SRC%\KBs\NET\*.bat") do (
    echo Running BAT %%F
    call "%%F"
)
echo Installing updates from %SRC%\KBs\MISC

:: Install EXE installers
for %%F in ("%SRC%\KBs\MISC\*.exe") do (
    echo Installing EXE %%F
    "%%F" /quiet /norestart
)

:: Install MSI installers
for %%F in ("%SRC%\KBs\MISC\*.msi") do (
    echo Installing MSI %%F
    msiexec.exe /i "%%F" /quiet /norestart
)

:: Run CMD/BAT scripts
for %%F in ("%SRC%\KBs\MISC\*.cmd") do (
    echo Running CMD %%F
    call "%%F"
)
for %%F in ("%SRC%\KBs\MISC\*.bat") do (
    echo Running BAT %%F
    call "%%F"
)

:end
endlocal
