<#
.SYNOPSIS
  The ultimate driver cleanup tool. Removes all non-present (ghost) devices, 
  clears folder-matched OEM drivers, and sweeps the driver store for unreferenced packages.

.DESCRIPTION
  Phase 1: Scans for and removes ALL non-present devices across all classes (Registry cleanup).
  Phase 2: (Optional) Matches original INFs from a specific folder against the driver store and deletes them.
  Phase 3: Computes remaining oem*.inf files in the driver store that have no attached devices and offers removal.

.PARAMETER InfPath
  (Optional) Path to a folder containing *.inf files to explicitly remove. 

.PARAMETER Recurse
  If present, searches the InfPath subfolders recursively.

.PARAMETER DryRun
  If specified, prints the commands/removals that would be executed without changing the system.

.PARAMETER AcceptAll
  If specified, automatically selects all orphaned packages for removal in Phase 3.

.PARAMETER Force
  If specified, adds the /force flag to the delete-driver commands.

.PARAMETER CreateRestorePoint
  If specified, attempts to create a system restore point before beginning the cleanse.

.EXAMPLE
  .\CleanupDrivers.ps1 -DryRun

.EXAMPLE
  .\CleanupDrivers.ps1 -AcceptAll -Force -CreateRestorePoint

.EXAMPLE
  .\CleanupDrivers.ps1 -InfPath "C:\OldDrivers" -Recurse -AcceptAll
#>

param(
    [string]$InfPath = "",
    [switch]$Recurse = $false,
    [switch]$DryRun = $false,
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
            if ($out -and $out.Count -gt 0) {
                $outputs += $out
            }
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
# PHASE 1: GHOST DEVICE REMOVAL (ALL CLASSES)
# ==========================================
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " PHASE 1: Removing Non-Present Devices" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

# Fetch all Plug and Play devices and filter for those not present
$ghostDevices = Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object { $_.Present -eq $false }

if (-not $ghostDevices -or $ghostDevices.Count -eq 0) {
    Write-Host "No ghost devices found across any class."
} else {
    Write-Host "Found $($ghostDevices.Count) non-present device(s) across all classes."
    
    foreach ($device in $ghostDevices) {
        $friendlyName = if ($device.FriendlyName) { $device.FriendlyName } else { "Unknown Device" }
        $instanceId = $device.InstanceId

        if ($DryRun) {
            Write-Host "[DRY RUN] Would remove ghost device: $friendlyName"
        } else {
            Write-Host "Removing: $friendlyName..."
            $output = & pnputil.exe /remove-device $instanceId 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Host "  -> Successfully removed." -ForegroundColor Green
            } else {
                Write-Warning "  -> Failed to remove. It may be a protected software device."
            }
        }
    }
}
Write-Host ""


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
    $infFiles = [System.IO.Directory]::EnumerateFiles((Resolve-Path -Path $InfPath).Path, '*.inf', $searchOption) | 
                ForEach-Object { [System.IO.Path]::GetFileName($_) } | Sort-Object -Unique

    if (-not $infFiles -or $infFiles.Count -eq 0) {
        Write-Host "No .inf files found under '$InfPath'."
    } else {
        Write-Host "Found $($infFiles.Count) INF file(s) in folder. Mapping to driver store..."
        
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
            if (-not $map.ContainsKey($infKey)) { continue }

            $publishedList = $map[$infKey] | Sort-Object -Unique
            foreach ($pub in $publishedList) {
                $args = @("/delete-driver", $pub, "/uninstall")
                if ($Force) { $args += "/force" }

                if ($DryRun) {
                    Write-Host "[DRY RUN] Would run: pnputil.exe $($args -join ' ')"
                } else {
                    Write-Host "Running: pnputil.exe $($args -join ' ')"
                    $output = & pnputil.exe @args 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        Write-Host "  -> Successfully removed '$pub' (matched to '$infName')." -ForegroundColor Green
                    } else {
                        Write-Warning "  -> Failed for '$pub'. Exit Code: $LASTEXITCODE"
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

Write-Host "Gathering in-use published names..."
$inUseSet = Get-InUsePublishedNames

Write-Host "Gathering all published names in the driver store..."
$allSet = Get-AllPublishedNamesInStore

$orphanList = @()
foreach ($name in $allSet) {
    if (-not $inUseSet.Contains($name)) {
        $orphanList += $name
    }
}

if ($orphanList.Count -eq 0) {
    Write-Host "No orphaned oem*.inf packages found. Driver store is clean!" -ForegroundColor Green
} else {
    Write-Host "`nFound the following orphaned driver packages (not referenced by ANY active device):`n" -ForegroundColor Yellow
    for ($i = 0; $i -lt $orphanList.Count; $i++) {
        Write-Host ("[{0}] {1}" -f ($i + 1), $orphanList[$i])
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
        # DryRun with no AcceptAll just assumes all for display purposes
        $selection = 1..$orphanList.Count 
    }

    foreach ($i in $selection) {
        $pub = $orphanList[$i - 1]
        $args = @("/delete-driver", $pub, "/uninstall")
        if ($Force) { $args += "/force" }

        if ($DryRun) {
            Write-Host "[DRY RUN] Would execute: pnputil $($args -join ' ')"
        } else {
            Write-Host "`nExecuting: pnputil $($args -join ' ')"
            & pnputil @args
            if ($LASTEXITCODE -eq 0) {
                Write-Host "  -> Removed $pub successfully." -ForegroundColor Green
            } else {
                Write-Warning "  -> Failed to remove $pub. Exit code: $LASTEXITCODE"
            }
        }
    }
}

Write-Host "`nMaster Cleanse Complete." -ForegroundColor Cyan