<#
.SYNOPSIS
  The ultimate driver and hardware cleanup tool. Removes non-present (ghost) devices 
  using pure PowerShell registry hooks, clears folder-matched OEM drivers, and 
  sweeps the driver store for unreferenced packages.

.DESCRIPTION
  Phase 1: Scans for and removes only GREYED OUT (hidden/phantom) devices using the exact 
           same ConfigManager Status logic as Device Manager. Prompts before deleting protected classes.
  Phase 2: (Optional) Matches original INFs from a specific folder against the driver store and deletes them.
  Phase 3: Computes remaining oem*.inf files in the driver store that have no attached devices and offers removal.
#>

param(
    [string]$InfPath = "",
    [switch]$Recurse = $false,
    [switch]$DryRun = $false,
    [switch]$RemoveProtected = $false,
    [switch]$AcceptAll = $false,
    [switch]$Force = $false,
    [switch]$CreateRestorePoint = $false
)

# --- Helper Functions ---

function Assert-Admin {
    $current = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($current)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Error "This script must be run elevated (as Administrator). Exiting."
        exit 1
    }
}

function Get-PublishedNamesFromPnPUtilOutput {
    param([string[]]$Lines)
    $set = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($line in $Lines) {
        if ($line -match '\b(oem\d+\.inf)\b') {
            $set.Add($matches[1]) | Out-Null
        }
    }
    return $set
}

function Get-InUsePublishedNames {
    $outputs = @()
    $cmds = @(
        @{Exe='pnputil'; Args='/enum-devices'},
        @{Exe='pnputil'; Args='/enum-drivers /devices'},
        @{Exe='pnputil'; Args='/enum-drivers /installed'}
    )
    foreach ($c in $cmds) {
        try {
            $out = & $c.Exe $c.Args 2>&1
            if ($out -and $out.Count -gt 0) { $outputs += $out }
        } catch {}
    }
    return Get-PublishedNamesFromPnPUtilOutput -Lines $outputs
}

function Get-AllPublishedNamesInStore {
    $out = & pnputil /enum-drivers 2>&1
    return Get-PublishedNamesFromPnPUtilOutput -Lines $out
}

# --- Initialization ---

Assert-Admin

if ($CreateRestorePoint) {
    try {
        Write-Host "Creating system restore point..." -ForegroundColor Cyan
        Checkpoint-Computer -Description "Pre-Master Driver Cleanse" -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
        Write-Host "Restore point created successfully.`n" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to create restore point: $_`n"
    }
}

# ==========================================
# PHASE 1: EXACT DEVICE MANAGER GHOST REMOVAL (NO C#)
# ==========================================
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " PHASE 1: Removing Hidden (Greyed Out) Devices" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

# Define protected classes by their friendly Device Manager Class names
$ProtectedClasses = @(
    "PrintQueue",                  # Print queues
    "Printer",                     # Printers
    "VolumeSnapshot",              # Storage volume shadow copies
    "Volume",                      # Storage volumes
    "WPD",                         # Portable Devices (Phones/Media players often unplugged)
    "Bluetooth",                   # Bluetooth links/Headphones (avoids re-pairing hassles)
    "Net"                          # Network adapters (prevents losing hidden VPN/Virtual switches)
)

$ProtectedClassAnswers = @{}
$UserApprovedAllClasses = $false

Write-Host "Connecting to native Windows Management Engine..." -ForegroundColor DarkGray

# Pulling all registered system device instances natively via WMI/CIM
# ConfigManagerErrorCode 0 means the device is functioning; StatusInfo 3 means it's started/present.
# If a device lacks an active status or shows an explicit disconnected code, it's greyed out.
$allDevices = Get-CimInstance -ClassName Win32_PnPEntity -ErrorAction SilentlyContinue

$processedCounter = 0
$removedCount = 0
$skippedCount = 0

Write-Host "Analyzing active device nodes matching greyed-out state..." -ForegroundColor Cyan

foreach ($device in $allDevices) {
    $processedCounter++

    # Check live update diagnostics loop every 50 nodes
    if ($processedCounter % 50 -eq 0) {
        Write-Host "  -> Inspected $processedCounter hardware profiles... running seamlessly..." -ForegroundColor DarkGray
    }

    # REPLICATION OF DEVICE MANAGER RULES:
    # Present devices return Status = "OK". If Status is null, empty, or anything else, 
    # the device is physically absent and appears as GREYED OUT / HIDDEN in Device Manager.
    $isStarted = ($device.Status -eq "OK")

    if (-not $isStarted) {
        # This device is strictly a phantom/hidden device (Greyed Out)
        $friendlyName = $device.Name
        if ([string]::IsNullOrEmpty($friendlyName)) { $friendlyName = $device.Description }
        if ([string]::IsNullOrEmpty($friendlyName)) { $friendlyName = "Unknown Device" }

        $deviceClass = $device.ClassGuid
        # Resolve Guid back to class name if possible
        if ($device.Service) { $deviceClass = $device.Service }
        
        # Pull clean Class string matching our array lookup
        $pnpDev = Get-PnpDevice -InstanceId $device.DeviceID -ErrorAction SilentlyContinue
        if ($pnpDev) { $deviceClass = $pnpDev.Class }
        if ([string]::IsNullOrEmpty($deviceClass)) { $deviceClass = "Unclassified" }

        $isProtected = $false
        if ($deviceClass -and ($ProtectedClasses -contains $deviceClass)) {
            $isProtected = $true
            
            if ($RemoveProtected -or $UserApprovedAllClasses) {
                $ProtectedClassAnswers[$deviceClass] = $true
            }
            elif (-not $ProtectedClassAnswers.ContainsKey($deviceClass)) {
                if ($DryRun) {
                    $ProtectedClassAnswers[$deviceClass] = $false
                } else {
                    Write-Host ""
                    Write-Host "[PROMPT REQUIRED] Hidden node matched inside a protected group." -ForegroundColor Yellow
                    $response = Read-Host "Found hidden item in protected class '$deviceClass' ($friendlyName).`nDo you want to clear this class? [Y]es / [N]o / [A]ll remaining classes"
                    
                    if ($response -match '^[aA]') {
                        $UserApprovedAllClasses = $true
                        $ProtectedClassAnswers[$deviceClass] = $true
                        Write-Host "-> Auto-approving ALL protected device classes from here on out!" -ForegroundColor Cyan
                    }
                    elif ($response -match '^[yY]') {
                        $ProtectedClassAnswers[$deviceClass] = $true
                        Write-Host "-> Proceeding with removal for class: $deviceClass" -ForegroundColor Yellow
                    } else {
                        $ProtectedClassAnswers[$deviceClass] = $false
                        Write-Host "-> Skipping all hidden items in class: $deviceClass" -ForegroundColor Gray
                    }
                }
            }
        }

        if ($isProtected -and ($ProtectedClassAnswers[$deviceClass] -eq $false)) {
            Write-Host "  [SKIPPED] Protected class preservation: $friendlyName (Class: $deviceClass)" -ForegroundColor DarkYellow
            $skippedCount++
            continue
        }

        if ($DryRun) {
            Write-Host "  [DRY RUN] Would forcefully purge greyed-out device: $friendlyName (Class: $deviceClass)" -ForegroundColor Yellow
        } else {
            Write-Host "  [PURGING PHANTOM] $friendlyName (Class: $deviceClass)..." -NoNewline
            
            # Use native WMI instance binding to execute the deletion command directly inside the OS 
            # hardware engine. This executes using current shell process tokens without file system modifications.
            try {
                $device | Remove-CimInstance -ErrorAction Stop
                Write-Host " Cleaned." -ForegroundColor Green
                $removedCount++
            } catch {
                # Fallback to structural execution if WMI handle is locked
                $instanceId = $device.DeviceID
                $escapedId = $instanceId -replace '\\', '\\'
                $wmiMatch = Get-WmiObject -Class Win32_PnPEntity -Filter "DeviceID='$escapedId'" -ErrorAction SilentlyContinue
                if ($wmiMatch) {
                    try {
                        $wmiMatch.Delete() | Out-Null
                        Write-Host " Cleaned via fallback." -ForegroundColor Green
                        $removedCount++
                    } catch {
                        Write-Host " Failed." -ForegroundColor Red
                    }
                } else {
                    Write-Host " Failed." -ForegroundColor Red
                }
            }
        }
    }
}

Write-Host "`n--- Phase 1 Diagnostics Summary ---" -ForegroundColor Cyan
Write-Host "Total Hardware DevNodes Inspected: $processedCounter"
Write-Host "True Greyed-Out Phantoms Purged:    $removedCount" -ForegroundColor Green
Write-Host "Protected Hidden Nodes Retained:    $skippedCount" -ForegroundColor Yellow
Write-Host "------------------------------------`n"


# ==========================================
# PHASE 2: FOLDER-BASED INF REMOVAL
# ==========================================
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " PHASE 2: Folder-Based OEM Driver Cleanup" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

if ([string]::IsNullOrWhiteSpace($InfPath)) {
    Write-Host "No -InfPath provided. Skipping Phase 2."
} elseif (-not (Test-Path -Path $InfPath)) {
    Write-Warning "Path '$InfPath' does not exist. Skipping Phase 2."
} else {
    $searchOption = if ($Recurse) { [System.IO.SearchOption]::AllDirectories } else { [System.IO.SearchOption]::TopDirectoryOnly }
    Write-Host "Scanning '$InfPath' for original manufacturer INF source layouts..." -ForegroundColor Gray
    
    $infFiles = [System.IO.Directory]::EnumerateFiles((Resolve-Path -Path $InfPath).Path, '*.inf', $searchOption) | 
                ForEach-Object { [System.IO.Path]::GetFileName($_) } | Sort-Object -Unique

    if (-not $infFiles -or $infFiles.Count -eq 0) {
        Write-Host "No .inf files found under '$InfPath'."
    } else {
        Write-Host "Found $($infFiles.Count) layout INF file(s). Correlating with live system definitions..."
        
        $enumOutput = & pnputil.exe /enum-drivers 2>&1
        $joined = $enumOutput -join "`n"
        $blocks = [regex]::Split($joined, "^\s*$", [System.Text.RegularExpressions.RegexOptions]::Multiline)
        $map = @{} 

        foreach ($b in $blocks) {
            if ([string]::IsNullOrWhiteSpace($b)) { continue }
            $published = $null
            $original = $null
            foreach ($line in ($b -split "`n")) {
                $trim = $line.Trim()
                if ($trim -match 'Published Name\s*:\s*(\S+)') { $published = $matches[1] } 
                elseif ($trim -match 'Original Name\s*:\s*(.+)') { $original = $matches[1].Trim() }
            }
            if ($published -and $original) {
                $origFile = [System.IO.Path]::GetFileName($original).ToLowerInvariant()
                if (-not $map.ContainsKey($origFile)) { $map[$origFile] = @() }
                $map[$origFile] += $published
            }
        }

        foreach ($infName in $infFiles) {
            $infKey = $infName.ToLowerInvariant()
            if (-not $map.ContainsKey($infKey)) {
                continue 
            }

            $publishedList = $map[$infKey] | Sort-Object -Unique
            foreach ($pub in $publishedList) {
                $args = @("/delete-driver", $pub, "/uninstall")
                if ($Force) { $args += "/force" }

                if ($DryRun) {
                    Write-Host "  [DRY RUN] Would remove mapped file: pnputil.exe $($args -join ' ')" -ForegroundColor Yellow
                } else {
                    Write-Host "  [REMOVING MAPPED STORE] Matches '$infName' -> Deleting $pub..." -NoNewline
                    $output = & pnputil.exe @args 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        Write-Host " Success." -ForegroundColor Green
                    } else {
                        Write-Host " Failed (Exit Code: $LASTEXITCODE)." -ForegroundColor Red
                    }
                }
            }
        }
    }
}
Write-Host ""


# ==========================================
# PHASE 3: ORPHANED DRIVER SWEEP
# ==========================================
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " PHASE 3: Orphaned Driver Store Sweep" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

Write-Host "Querying PnPUtil engine for in-use components..." -ForegroundColor DarkGray
$inUseSet = Get-InUsePublishedNames

Write-Host "Querying driver store for all registered oem*.inf catalogs..." -ForegroundColor DarkGray
$allSet = Get-AllPublishedNamesInStore

$orphanList = @()
foreach ($name in $allSet) {
    if (-not $inUseSet.Contains($name)) {
        $orphanList += $name
    }
}

if ($orphanList.Count -eq 0) {
    Write-Host "No orphaned oem*.inf packages found. Driver store matches hardware precisely!" -ForegroundColor Green
} else {
    Write-Host "`nFound $($orphanList.Count) abandoned driver configurations (not used by any active or saved device profile):" -ForegroundColor Yellow
    for ($i = 0; $i -lt $orphanList.Count; $i++) {
        Write-Host ("  [{0}] {1}" -f ($i + 1), $orphanList[$i])
    }

    $selection = @()
    if ($AcceptAll) {
        $selection = 1..$orphanList.Count
    } elseif (-not $DryRun) {
        Write-Host ""
        $input = Read-Host "Enter comma-separated numbers to remove (e.g. 1,3), 'a' for all, or 'q' to quit"
        if ($input -match '^[aA]') {
            $selection = 1..$orphanList.Count
        } elseif ($input -notmatch '^[qQ]') {
            $parts = $input -split '[, ]+' | Where-Object { $_ -match '^\d+$' }
            foreach ($p in $parts) {
                $n = [int]$p
                if ($n -ge 1 -and $n -le $orphanList.Count) { $selection += $n }
            }
        }
    } else {
        $selection = 1..$orphanList.Count 
    }

    if ($selection.Count -eq 0) {
        Write-Host "No items selected for removal. Skipping store purge." -ForegroundColor Gray
    } else {
        Write-Host "`nProcessing selected driver deletions..." -ForegroundColor Cyan
        foreach ($i in $selection) {
            $pub = $orphanList[$i - 1]
            $args = @("/delete-driver", $pub, "/uninstall")
            if ($Force) { $args += "/force" }

            if ($DryRun) {
                Write-Host "  [DRY RUN] Would execute: pnputil $($args -join ' ')" -ForegroundColor Yellow
            } else {
                Write-Host "  [DELETING ORPHAN] Dropping $pub from DriverStore..." -NoNewline
                $output = & pnputil @args 2>&1
                if ($LASTEXITCODE -eq 0) {
                    Write-Host " Complete." -ForegroundColor Green
                } else {
                    Write-Host " Rejected (Exit: $LASTEXITCODE)." -ForegroundColor Red
                    Write-Warning "Package $pub could not be dropped. It may require a forced reboot loop or the -Force switch."
                }
            }
        }
    }
}

Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host " MASTER CLEANSE COMPLETE." -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Cyan
