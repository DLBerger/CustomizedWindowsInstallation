:: SetupComplete.cmd
:: This file must be placed on installation media in: sources\$OEM$\$$\Setup\Scripts

@echo off
setlocal enabledelayedexpansion

set "PSNAME=PostSetup.cmd"
set "DRVNAME=$WinpeDriver$"

:: Search all drives for both PSNAME and DRVNAME
for %%D in (C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
    if not defined PS if exist "%%D:\%PSNAME%" (
        set "PS=%%D:\%PSNAME%"
    )
    if not defined DRV if exist "%%D:\%DRVNAME%" (
        set "DRV=%%D:\%DRVNAME%"
    )
    if defined PS if defined DRV (
        goto runps
    )
)

:runps
if defined PS (
    echo Running PostSetup.cmd from "%PS%"
    call "%PS%"
) else (
    echo PostSetup.cmd not found on any drive
)

endlocal
exit /b 0
