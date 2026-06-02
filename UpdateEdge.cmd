@echo off
setlocal enabledelayedexpansion

echo Edge update process start
:: 1. Download the latest Stable 64-bit Edge Enterprise MSI
curl -sL -o "%TEMP%\MicrosoftEdgeEnterpriseX64.msi" "https://msedge.sf.dl.delivery.mp.microsoft.com/filestreamingservice/files/microsoft/edge/MicrosoftEdgeEnterpriseX64.msi"

:: 2. Force install/upgrade the MSI completely silently
::    If it's already completely up-to-date, msiexec will safely do a quick repair or skip without errors.
if exist "%TEMP%\MicrosoftEdgeEnterpriseX64.msi" (
    start /wait "" msiexec.exe /i "%TEMP%\MicrosoftEdgeEnterpriseX64.msi" /qn /norestart
    del /f /q "%TEMP%\MicrosoftEdgeEnterpriseX64.msi"
)

echo Edge update process complete
pause

endlocal
