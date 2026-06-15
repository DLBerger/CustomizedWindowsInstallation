<#

.SYNOPSIS
Installs missing drivers by matching unknown device Hardware IDs against a folder of INF files.

.DESCRIPTION
This script scans the system for devices missing drivers, extracts their Hardware/Compatible IDs,
and selectively installs only the required INF files from the specified folder using pnputil.
#>

[CmdletBinding()]
param(
    [string]$Folder = ".\Drivers"
)

# If the user passed a relative path (like ".\Drivers"), resolve it relative to the script location
if (-not ([System.IO.Path]::IsPathRooted($Folder))) {
    $Folder = Join-Path $PSScriptRoot $Folder
}

if (-not (Test-Path $Folder)) {
    New-Item -ItemType Directory -Path $Folder -Force | Out-Null
}

$Folder = (Resolve-Path -LiteralPath $Folder -ErrorAction Continue).ProviderPath
Write-Host "Resolved driver folder: $Folder" -ForegroundColor Green

function Test-HwIdMatchesInfContent {
    param(
        [string]$HwId,
        [string]$InfContent
    )

    if ([string]::IsNullOrWhiteSpace($HwId) -or [string]::IsNullOrWhiteSpace($InfContent)) {
        return $false
    }

    # Full ID match (rare but valid when INF lists the complete hardware ID)
    if ($InfContent.IndexOf($HwId, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
        return $true
    }

    # INF entries are usually shorter prefixes of the device-reported ID
    # (e.g. INF has PCI\VEN_8086&DEV_15F3, device reports ...&SUBSYS_...&REV_...)
    $parts = $HwId -split '&'
    $prefix = ''
    foreach ($part in $parts) {
        if ($prefix) {
            $prefix = "$($prefix)&$($part)"
        } else {
            $prefix = $part
        }

        if ($InfContent.IndexOf($prefix, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
            return $true
        }
    }

    return $false
}

function Get-DriverStoreOriginalInfNames {
    $names = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    $enumOutput = & pnputil.exe /enum-drivers 2>&1
    if (-not $enumOutput) {
        return $names
    }

    $joined = $enumOutput -join "`n"
    $blocks = [regex]::Split($joined, "^\s*$", [System.Text.RegularExpressions.RegexOptions]::Multiline)

    foreach ($block in $blocks) {
        if ([string]::IsNullOrWhiteSpace($block)) { continue }
        foreach ($line in ($block -split "`n")) {
            $trim = $line.Trim()
            if ($trim -match 'Original Name\s*:\s*(.+)') {
                $origFile = [System.IO.Path]::GetFileName($matches[1].Trim())
                if (-not [string]::IsNullOrWhiteSpace($origFile)) {
                    $names.Add($origFile) | Out-Null
                }
                break
            }
        }
    }

    return $names
}

Write-Host "Searching Device Manager for devices missing drivers..." -ForegroundColor Cyan

# 2. Find devices missing drivers 
# Problem 28 is standard for "The drivers for this device are not installed."
$MissingDevices = Get-PnpDevice | Where-Object { 
    $_.Problem -eq 28 -or 
    $_.Status -eq 'Error' -or 
    $_.Status -eq 'Unknown' 
}

if (-not $MissingDevices) {
    Write-Host "All devices currently have drivers installed. Nothing to do!" -ForegroundColor Green
    Exit 0
}

Write-Host "Found $( $MissingDevices.Count ) device(s) missing drivers. Extracting Hardware IDs..." -ForegroundColor Cyan

# 3. Extract Hardware and Compatible IDs for the missing devices
$MissingHwIds = @()

foreach ($Device in $MissingDevices) {
    $HwIdProp = Get-PnpDeviceProperty -InstanceId $Device.InstanceId -KeyName 'DEVPKEY_Device_HardwareIds' -ErrorAction SilentlyContinue
    $CompIdProp = Get-PnpDeviceProperty -InstanceId $Device.InstanceId -KeyName 'DEVPKEY_Device_CompatibleIds' -ErrorAction SilentlyContinue
    
    if ($HwIdProp.Data) {
        foreach ($id in @($HwIdProp.Data)) {
            if (-not [string]::IsNullOrWhiteSpace($id)) { $MissingHwIds += $id }
        }
    }
    if ($CompIdProp.Data) {
        foreach ($id in @($CompIdProp.Data)) {
            if (-not [string]::IsNullOrWhiteSpace($id)) { $MissingHwIds += $id }
        }
    }
}

# Clean up the list (remove empties and duplicates)
$MissingHwIds = $MissingHwIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique

if (-not $MissingHwIds) {
    Write-Warning "Could not retrieve Hardware IDs for the missing devices."
    Exit 1
}

Write-Host "Searching INF files in '$Folder' for matching IDs..." -ForegroundColor Cyan

# 4. Scan all INF files in the target folder for matches
$AllInfs = Get-ChildItem -Path $Folder -Recurse -File |
    Where-Object { $_.Extension -ieq '.inf' }
$MatchedInfs = @()

foreach ($Inf in $AllInfs) {
    try {
        $InfContent = Get-Content -LiteralPath $Inf.FullName -Raw -ErrorAction Stop
    } catch {
        Write-Warning "Skipping unreadable path: $($Inf.FullName) ($($_.Exception.Message))"
        continue
    }

    if ([string]::IsNullOrWhiteSpace($InfContent)) {
        continue
    }
    
    foreach ($HwId in $MissingHwIds) {
        if (Test-HwIdMatchesInfContent -HwId $HwId -InfContent $InfContent) {
            $MatchedInfs += $Inf.FullName
            break # Match found, stop checking this INF against other IDs and move to the next INF
        }
    }
}

$MatchedInfs = $MatchedInfs | Sort-Object -Unique

if (-not $MatchedInfs) {
    Write-Warning "No matching drivers were found in the provided folder."
    Exit 0
}

Write-Host "Checking driver store for packages already present..." -ForegroundColor Cyan
$DriverStoreInfNames = Get-DriverStoreOriginalInfNames

$ToInstall = @()
foreach ($InfPath in $MatchedInfs) {
    $infName = [System.IO.Path]::GetFileName($InfPath)
    if ($DriverStoreInfNames.Contains($infName)) {
        Write-Host "Skipping (already in driver store): $InfPath" -ForegroundColor DarkGray
        continue
    }
    $ToInstall += $InfPath
}

if (-not $ToInstall) {
    Write-Host "All matching drivers are already in the driver store. Nothing to install." -ForegroundColor Green
    Exit 0
}

Write-Host "Found $( $ToInstall.Count ) matching driver INF(s) to install (skipped $( $MatchedInfs.Count - $ToInstall.Count ) already present)..." -ForegroundColor Magenta
Write-Host "------------------------------------------------------" -ForegroundColor Gray

# 5. Install only the matched drivers using pnputil
foreach ($InfPath in $ToInstall) {
    Write-Host ">>> Installing: $InfPath" -ForegroundColor Yellow
    
    # Run pnputil to add and install the specific driver
    $pnputilArgs = @("/add-driver", "`"$InfPath`"", "/install")
    & pnputil.exe @pnputilArgs | Out-Host
    
    Write-Host "------------------------------------------------------" -ForegroundColor Gray
}

Write-Host "Driver installation phase complete!" -ForegroundColor Green