:: SetupComplete.cmd
:: This file must be placed on installation media in: sources\$OEM$\$$\Setup\Scripts

@echo off
setlocal enabledelayedexpansion

set "DRVNAME=$WinpeDriver$"
set "REGCMD=InstallRegs.cmd"

:: Search all drives for both REGCMD and DRVNAME
for %%D in (C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
    if not defined OSDRIVE if exist "%%D:\Windows\System32\config\SYSTEM" (
        set "OSDRIVE=%%D:"
    )
    if not defined DRV if exist "%%D:\%DRVNAME%" (
        set "DRV=%%D:\%DRVNAME%"
    )
    if defined DRV if not defined RC if exist "%%D:\%REGCMD%" (
        set "RC=%%D:\%REGCMD%"
    )
    :: Done when we find all 3
    if defined OSDRIVE if defined DRV if defined RC (
        goto found
    )
)
:fail
exit /b 1

:found

echo Running %RC%
call "%RC%"

:: Inject the drivers
dism /Image:"%OSDRIVE%\" /Add-Driver /Driver:"%DRV%" /Recurse /ForceUnsigned /logpath:"%DRV%\dism.log"

endlocal
exit /b %ERRORLEVEL%