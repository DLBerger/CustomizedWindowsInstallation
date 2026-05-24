@echo off
setlocal
set "SRC=%~dp0"
set "FLD=Registry"

echo Import registry files
for %%F in ("%SRC%%FLD%\*.reg") do (
    reg.exe import "%%F"
)
endlocal