@echo off
setlocal enabledelayedexpansion

set LOG=%SystemRoot%\Setup\Scripts\SetupComplete.log
echo [%DATE% %TIME%] SetupComplete starting... > "%LOG%"

set BASE=%~dp0

:: Apply .NET updates
for %%F in ("%BASE%..\..\..\Updates\NET\*.msu") do (
    wusa.exe "%%F" /quiet /norestart >> "%LOG%" 2>&1
)

:: Apply OS updates
for %%F in ("%BASE%..\..\..\Updates\OSCU\*.msu") do (
    wusa.exe "%%F" /quiet /norestart >> "%LOG%" 2>&1
)

:: Import registry
for %%F in ("%BASE%..\..\..\Registry\*.reg") do (
    reg.exe import "%%F" >> "%LOG%" 2>&1
)

:: Install drivers
if exist "%SystemDrive%\Install Drivers.cmd" (
    call "%SystemDrive%\Install Drivers.cmd" >> "%LOG%" 2>&1
)

echo [%DATE% %TIME%] SetupComplete finished. >> "%LOG%"
exit /b 0
