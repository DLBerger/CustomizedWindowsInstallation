@echo off
setlocal enabledelayedexpansion

echo Updating Microsoft Edge to the latest stable build...

:: Set local variables for clean handling
set "EDGE_MSI=%TEMP%\MicrosoftEdgeEnterpriseX64.msi"

:: Clean up any lingering files from previous failed runs
if exist "%EDGE_MSI%" del /f /q "%EDGE_MSI%"

:: 1. Download using static fwlink and -L flag to follow Microsoft's download server redirects
::    --retry 3: retry on transient failures (network drops, 503s)
::    --max-time 120: abort if the connection stalls for more than 2 minutes
::    --fail: make curl exit non-zero on HTTP 4xx/5xx responses
curl -s -L --retry 3 --max-time 120 --fail -o "%EDGE_MSI%" "https://go.microsoft.com/fwlink/?LinkID=2093437"

if %ERRORLEVEL% NEQ 0 (
    echo ERROR: curl failed with exit code %ERRORLEVEL%. Check internet connection.
    goto :EndEdgeUpdate
)

:: 2. Check if file exists AND verify it isn't an empty 0-byte file
if exist "%EDGE_MSI%" (
    
    :: Call a local block to check file size safely
    call :CheckFileSize "%EDGE_MSI%"
    
) else (
    echo ERROR: Failed to download the Edge MSI installer. Check internet connection.
)

:: Skip past the subroutine block so it doesn't run twice
goto :EndEdgeUpdate

:CheckFileSize
:: %~z1 expands the first argument passed to this subroutine to its file size in bytes
if "%~z1"=="0" (
    echo ERROR: Curl pulled down an empty 0-byte file! Installation aborted.
    del /f /q %1
) else (
    echo [OK] Verified valid payload download (%~z1 bytes). Installing...
    start /wait "" msiexec.exe /i %1 /qn /norestart
    del /f /q %1
    echo Edge update process complete.
)
goto :eof

:EndEdgeUpdate

echo Edge update process complete
pause

endlocal
