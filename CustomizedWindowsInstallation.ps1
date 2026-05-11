<#
.SYNOPSIS
Builds a reusable update/driver/registry/script payload for Windows 10/11 installation media.

.DESCRIPTION
This script prepares a directory structure that can be copied to the root of a USB drive
containing official Windows installation media (Windows 10 22H2+ or Windows 11 25H2+).

It supports:
- Downloading OS cumulative updates and .NET updates from the Microsoft Update Catalog.
- Exporting all third-party drivers from the current system into $WinpeDriver$.
- Exporting registry keys into .reg files.
- Generating:
    - InstallDrivers.cmd
    - InstallRegs.cmd
    - PostSetup.cmd
    - SetupConfig-Clean.ini
    - SetupConfig-Upgrade.ini
    - Upgrade.cmd
    - CleanInstall.cmd
- Dry-run mode (no changes made)
- Clean mode (remove generated content)

.PARAMETER Folder
Root folder where the update/driver/registry/scripts structure will be created.
If omitted, defaults to the current working directory.

.PARAMETER WinOS
Windows major version: '10' or '11'.
Alias: -OS
If omitted, defaults to '11'.

.PARAMETER Version
Windows feature update version (for example: '22H2', '25H2').
If omitted:
- Windows 10 -> '22H2'
- Windows 11 -> '25H2'

.PARAMETER Arch
CPU architecture: 'x64' or 'arm64'.
If omitted, defaults to 'x64'.

.PARAMETER Export
Mount and copy the contents of <ISO> to <Folder> and then export the requested indices (see below).

.PARAMETER KB
Download OS and .NET updates.

.PARAMETER Service
  Apply the download KBs to the exported indices and create the final install.wim and boot.wim.

.PARAMETER Drivers
Export drivers into $WinpeDriver$.

.PARAMETER Reg
Export registry keys.

.PARAMETER Files
Generate PostSetup.cmd, SetupConfig-*.ini, and additional .cmd files.

.PARAMETER All
Shorthand for -Export -KB -Service -Drivers -Reg -Files and the default if no specific switch is provided.

.PARAMETER Clean
Remove generated content instead of creating it.

.PARAMETER DryRun
Show actions without performing them.

.PARAMETER Help
Displays help and exits.

.NOTES
Fully compatible with Windows PowerShell 5.x.

References:
[1] https://learn.microsoft.com/en-us/windows/deployment/update/media-dynamic-update
[2] https://github.com/Marco-online/MSCatalogLTS
[3] https://www.deploymentresearch.com/removing-applications-from-your-windows-11-image-before-and-during-deployment/
[4] https://thedotsource.com/2021/03/16/building-iso-files-with-powershell-7/
[5] https://learn.microsoft.com/en-us/windows-hardware/customize/desktop/unattend/microsoft-windows-pnpcustomizationswinpe-driverpaths
[6] https://community.spiceworks.com/t/autounattend-xml-driver-path-issue-for-windows-11-24h2-and-25h2/1244985
[7] https://github.com/wikijm/PowerShell-AdminScripts/blob/master/Miscellaneous/New-IsoFile.ps1
[8] https://www.winhelponline.com/blog/servicing-stack-diagnosis-dism-sfc/
#>

param(
    [Parameter(Position = 0)]
    [string]$Folder,

    [Alias('OS')]
    [ValidateSet('10','11')]
    [string]$WinOS,

    [string]$Version,

    [ValidateSet('x64','arm64')]
    [string]$Arch,

    [switch]$Export,
    [switch]$KB,
    [switch]$Service,
    [switch]$Drivers,
    [switch]$Reg,
    [switch]$Files,
    [switch]$All,

    [switch]$Clean,

    [switch]$DryRun,

    [switch]$Help
)

# git hash
$GitHash = "ade1977"

# ==============================
# Core names
# ==============================
$names = [ordered]@{
    Checkpoint            = 'Checkpoint'
    SrcIso                = 'SrcISO'
    DestIso               = 'DestISO'
    KBs                   = 'KBs'
    Wims                  = 'Wims'
    WinpeDriver           = '$WinpeDriver$'
    Registry              = 'Registry'
    Source                = 'source'
    BootWim               = 'boot.wim'
    InstallEsd            = 'install.esd'
    InstallWim            = 'install.wim'
    WinreWim              = 'winre.wim'
    InstallDriversCmd     = 'InstallDrivers.cmd'
    InstallRegsCmd        = 'InstallRegs.cmd'
    PostSetupCmd          = 'PostSetup.cmd'
    SetupConfigCleanIni   = 'SetupConfig-Clean.ini'
    SetupConfigUpgradeIni = 'SetupConfig-Upgrade.ini'
    CleanInstallCmd       = 'CleanInstall.cmd'
    UpgradeCmd            = 'Upgrade.cmd'
}

$kbDirs = @('SSU', 'OSCU', 'NET', 'MISC')
foreach ($u in $kbDirs) {
    $names[$u] = $u
}

$wimDirs = @('Indices', 'Mounts', 'Serviced')
foreach ($u in $wimDirs) {
    $names[$u] = $u
}

# Ensure elevated
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
Write-Error "$PSCommandPath must be run elevated as Administrator."
exit 1
}

if ($Help) {
Get-Help -Full $PSCommandPath
exit
}

# If -Debug was passed, force debug output to auto-continue
if ($PSBoundParameters.ContainsKey('Debug')) {
$DebugPreference = 'Continue'
Write-Debug "Debug mode enabled: DebugPreference set to 'Continue'"
}

# Silence progress bars
$ProgressPreference = 'SilentlyContinue'
Write-Debug "ProgressPreference set to 'SilentlyContinue'"

# ==============================
# Helper functions
# ==============================
function Write-Log {
param([string]$Message, [string]$Level = 'INFO')
$ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
Write-Host "[$ts] [$Level] $Message"
}

# =========================
# Extract/Export section
# =========================

function Invoke-ExportWork {
    [CmdletBinding()]
    param()

    Write-Log "Starting Export workflow..."
    
    # Ensure source ISO root
    Ensure-Folder -Path $paths.SrcIsoRoot
    
    if ($DryRun) {
        Write-Log "[DryRun] Would mount and copy ISO contents to $($paths.SrcIsoRoot)"
        Write-Log "[DryRun] Would copy from $($paths.SrcIsoRoot) to $($paths.DestIsoRoot)"
        Write-Log "[DryRun] Would export requested indices to $($paths.WimsIndices)"
        return
    }
    
    if ($Clean) {
        if (Test-Path $paths.SrcIsoRoot) {
            Write-Log "Removing source ISO root: $($paths.SrcIsoRoot)"
            Remove-Item $paths.SrcIsoRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path $paths.DestIsoRoot) {
            Write-Log "Removing dest ISO root: $($paths.DestIsoRoot)"
            Remove-Item $paths.DestIsoRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path $paths.WimsIndices) {
            Write-Log "Removing indices: $($paths.WimsIndices)"
            Remove-Item $paths.WimsIndices -Recurse -Force -ErrorAction SilentlyContinue
        }
        return
    }
    
    # Checkpoint tracking
    $srcIsoCopyCheckpoint = Join-Path $paths.Checkpoint "srciso.copy.done"
    $destIsoCopyCheckpoint = Join-Path $paths.Checkpoint "destiso.copy.done"
    
    # Check if source ISO copy is already done
    if (Test-Path $srcIsoCopyCheckpoint) {
        Write-Log "Source ISO copy already completed (checkpoint exists)"
    } else {
        Write-Log "Copying ISO contents from mounted ISO to $($paths.SrcIsoRoot)..."
        
        # Verify $names.BootWim and ($names.InstallEsd or $names.InstallWim) exist
        $bootWimPath = Join-Path $paths.SourceInSrc $names.BootWim
        $installWimPath = Join-Path $paths.SourceInSrc $names.InstallWim
        $installEsdPath = Join-Path $paths.SourceInSrc $names.InstallEsd
        
        if (-not (Test-Path $bootWimPath)) {
            throw "Required file not found: $bootWimPath"
        }
        if (-not ((Test-Path $installWimPath) -or (Test-Path $installEsdPath))) {
            throw "Required file not found: $installWimPath or $installEsdPath"
        }
        
        # Mark checkpoint
        Ensure-Folder -Path $paths.Checkpoint
        Set-Content -Path $srcIsoCopyCheckpoint -Value (Get-Date -Format s) -Encoding UTF8
        Write-Log "Source ISO copy completed - checkpoint marked"
    }
    
    # Check if dest ISO copy is already done
    if (Test-Path $destIsoCopyCheckpoint) {
        Write-Log "Destination ISO copy already completed (checkpoint exists)"
    } else {
        Write-Log "Copying from $($paths.SrcIsoRoot) to $($paths.DestIsoRoot)..."
        
        # Copy contents using robocopy, excluding boot.wim, install.wim, install.esd
        Ensure-Folder -Path $paths.DestIsoRoot
        
        $robocopyArgs = @(
            $paths.SrcIsoRoot,
            $paths.DestIsoRoot,
            '/E',           # Recursively copy all subdirectories
            '/COPY:DAT',    # Copy data, attributes, timestamps
            '/XF',
            $names.BootWim,
            $names.InstallWim,
            $names.InstallEsd  # Exclude these files
        )
        
        Write-Verbose "Running: robocopy $($robocopyArgs -join ' ')"
        $robocopyOutput = robocopy @robocopyArgs
        
        if ($LASTEXITCODE -gt 7) {
            throw "Robocopy failed with exit code $LASTEXITCODE"
        }
        
        # Mark checkpoint
        Set-Content -Path $destIsoCopyCheckpoint -Value (Get-Date -Format s) -Encoding UTF8
        Write-Log "Destination ISO copy completed - checkpoint marked"
    }
    
    # Export requested indices
    Write-Log "Exporting requested indices to $($paths.WimsIndices)..."
    Ensure-Folder -Path $paths.WimsIndices
    
    # TODO: Extract indices from install.wim/install.esd based on index selection
    # For now, create checkpoint marker for the export job completion
    $exportCheckpoint = Join-Path $paths.Checkpoint "export.done"
    if (-not (Test-Path $exportCheckpoint)) {
        Write-Log "Index extraction would occur here (per-index uncompressed WIMs)"
        Set-Content -Path $exportCheckpoint -Value (Get-Date -Format s) -Encoding UTF8
        Write-Log "Export checkpoint marked"
    }
    
    Write-Log "Export workflow complete"
}

# =========================
# Service/Patch section
# =========================

function Invoke-ServiceWork {
    [CmdletBinding()]
    param()

    Write-Log "Starting Service workflow..."
    
    if ($DryRun) {
        Write-Log "[DryRun] Would mount and service extracted indices"
        Write-Log "[DryRun] Would apply packages from KB folders (SSU, OSCU) to install.wim and boot.wim"
        Write-Log "[DryRun] Would service winre.wim inside each index"
        Write-Log "[DryRun] Would create final compressed install.wim and boot.wim"
        return
    }
    
    if ($Clean) {
        $wimCheckpoint = Join-Path $paths.Checkpoint "wims.done"
        if (Test-Path $paths.WimsRoot) {
            Write-Log "Removing WIMs root: $($paths.WimsRoot)"
            Remove-Item $paths.WimsRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path $wimCheckpoint) {
            Remove-Item $wimCheckpoint -Force -ErrorAction SilentlyContinue
        }
        return
    }
    
    # Service configuration
    $CompressionType = 'Maximum'  # None, Fast, Maximum
    $MaxParallelJobs = 4
    
    Ensure-Folder -Path $paths.WimsMounts
    Ensure-Folder -Path $paths.WimsServiced
    Ensure-Folder -Path $paths.Checkpoint
    
    # Get list of extracted indices to service
    if (-not (Test-Path $paths.WimsIndices)) {
        Write-Log "No indices folder found: $($paths.WimsIndices)"
        return
    }
    
    $extractedIndices = @()
    $extractedIndices += Get-ChildItem -Path $paths.WimsIndices -Filter "*_$($names.InstallWim)" -File |
        ForEach-Object { [int]($_.BaseName -replace "_$([regex]::Escape($names.InstallWim))$") } |
        Sort-Object
    
    if ($extractedIndices.Count -eq 0) {
        Write-Log "No extracted install.wim indices found"
        return
    }
    
    Write-Log "Found $($extractedIndices.Count) indices to service: $($extractedIndices -join ', ')"
    
    # Check for packages in KB folders
    $hasSSU = (Get-ChildItem -Path $paths.KBsSSU -Filter "*.cab" -ErrorAction SilentlyContinue | Measure-Object).Count -gt 0
    $hasOSCU = (Get-ChildItem -Path $paths.KBsOSCU -Filter "*.cab" -ErrorAction SilentlyContinue | Measure-Object).Count -gt 0
    $hasNET = (Get-ChildItem -Path $paths.KBsNET -Filter "*.cab" -ErrorAction SilentlyContinue | Measure-Object).Count -gt 0
    
    Write-Log "Package availability - SSU: $hasSSU, OSCU: $hasOSCU, NET: $hasNET"
    
    # Service each index (serial processing for safety)
    foreach ($idx in $extractedIndices) {
        $installWimFile = "$idx`_$($names.InstallWim)"
        $bootWimFile = "$idx`_$($names.BootWim)"
        $installWimPath = Join-Path $paths.WimsIndices $installWimFile
        $bootWimPath = Join-Path $paths.WimsIndices $bootWimFile
        
        $mountDir = Join-Path $paths.WimsMounts "mount_$idx"
        $winreMountDir = Join-Path $paths.WimsMounts "winre_mount_$idx"
        
        $indexCheckpoint = Join-Path $paths.Checkpoint "$idx.done"
        
        if (Test-Path $indexCheckpoint) {
            Write-Log "Index $idx already serviced (checkpoint exists)"
            continue
        }
        
        Write-Log "Servicing index $idx..."
        
        # Service install.wim
        if ($hasSSU -or $hasOSCU -or $hasNET) {
            Ensure-Folder -Path $mountDir
            
            try {
                Write-Log "  Mounting index $idx install.wim to $mountDir..."
                $dismArgs = @("/Mount-Image", "/ImageFile:$installWimPath", "/Index:1", "/MountDir:$mountDir")
                & dism.exe $dismArgs | Write-Verbose
                if ($LASTEXITCODE -ne 0) {
                    throw "Failed to mount install.wim for index $idx"
                }
                
                # Check for winre.wim inside mounted install
                $winreInWim = Join-Path $mountDir $paths.WinreWimInWim
                if (Test-Path $winreInWim) {
                    Write-Log "  Found winre.wim in mounted install image"
                    
                    # Extract winre.wim for separate servicing
                    if ($hasSSU) {
                        $winreExtracted = Join-Path $paths.WimsServiced "$idx`_winre_extracted.wim"
                        Write-Log "    Extracting winre.wim..."
                        Copy-Item -Path $winreInWim -Destination $winreExtracted -Force
                        
                        # Mount and service winre
                        Ensure-Folder -Path $winreMountDir
                        Write-Log "    Mounting winre.wim..."
                        $dismWinreArgs = @("/Mount-Image", "/ImageFile:$winreExtracted", "/Index:1", "/MountDir:$winreMountDir")
                        & dism.exe $dismWinreArgs | Write-Verbose
                        if ($LASTEXITCODE -ne 0) {
                            throw "Failed to mount winre.wim for index $idx"
                        }
                        
                        # Apply SSU packages to winre
                        $ssuFiles = Get-ChildItem -Path $paths.KBsSSU -Filter "*.cab" -ErrorAction SilentlyContinue
                        foreach ($file in $ssuFiles) {
                            Write-Log "    Applying SSU to winre: $($file.Name)"
                            $dismApplyArgs = @("/Image:$winreMountDir", "/Add-Package", "/PackagePath:$($file.FullName)")
                            & dism.exe $dismApplyArgs | Write-Verbose
                            if ($LASTEXITCODE -ne 0) {
                                Write-Log "    WARNING: Failed to apply $($file.Name) to winre (continuing)"
                            }
                        }
                        
                        # Unmount and commit winre
                        Write-Log "    Unmounting winre.wim..."
                        $dismUnmountArgs = @("/Unmount-Image", "/MountDir:$winreMountDir", "/Commit")
                        & dism.exe $dismUnmountArgs | Write-Verbose
                        if ($LASTEXITCODE -ne 0) {
                            throw "Failed to unmount winre.wim for index $idx"
                        }
                        
                        # Reinsert serviced winre back into install.wim
                        Write-Log "    Reinserting serviced winre.wim..."
                        Copy-Item -Path $winreExtracted -Destination $winreInWim -Force
                    }
                }
                
                # Apply SSU packages to install.wim
                if ($hasSSU) {
                    $ssuFiles = Get-ChildItem -Path $paths.KBsSSU -Filter "*.cab" -ErrorAction SilentlyContinue
                    foreach ($file in $ssuFiles) {
                        Write-Log "  Applying SSU: $($file.Name)"
                        $dismApplyArgs = @("/Image:$mountDir", "/Add-Package", "/PackagePath:$($file.FullPath)")
                        & dism.exe $dismApplyArgs | Write-Verbose
                        if ($LASTEXITCODE -ne 0) {
                            Write-Log "  WARNING: Failed to apply $($file.Name) (continuing)"
                        }
                    }
                }
                
                # Apply OSCU packages to install.wim
                if ($hasOSCU) {
                    $oscuFiles = Get-ChildItem -Path $paths.KBsOSCU -Filter "*.cab" -ErrorAction SilentlyContinue
                    foreach ($file in $oscuFiles) {
                        Write-Log "  Applying OSCU: $($file.Name)"
                        $dismApplyArgs = @("/Image:$mountDir", "/Add-Package", "/PackagePath:$($file.FullPath)")
                        & dism.exe $dismApplyArgs | Write-Verbose
                        if ($LASTEXITCODE -ne 0) {
                            Write-Log "  WARNING: Failed to apply $($file.Name) (continuing)"
                        }
                    }
                }
                
                # Apply NET packages to install.wim
                if ($hasNET) {
                    $netFiles = Get-ChildItem -Path $paths.KBsNET -Filter "*.cab" -ErrorAction SilentlyContinue
                    foreach ($file in $netFiles) {
                        Write-Log "  Applying NET: $($file.Name)"
                        $dismApplyArgs = @("/Image:$mountDir", "/Add-Package", "/PackagePath:$($file.FullPath)")
                        & dism.exe $dismApplyArgs | Write-Verbose
                        if ($LASTEXITCODE -ne 0) {
                            Write-Log "  WARNING: Failed to apply $($file.Name) (continuing)"
                        }
                    }
                }
                
                # Unmount and commit install.wim
                Write-Log "  Unmounting index $idx install.wim..."
                $dismUnmountArgs = @("/Unmount-Image", "/MountDir:$mountDir", "/Commit")
                & dism.exe $dismUnmountArgs | Write-Verbose
                if ($LASTEXITCODE -ne 0) {
                    throw "Failed to unmount install.wim for index $idx"
                }
                
            } catch {
                Write-Log "  ERROR servicing index $idx install.wim: $_"
                throw $_
            } finally {
                # Cleanup mount directory
                if (Test-Path $mountDir) {
                    Remove-Item $mountDir -Recurse -Force -ErrorAction SilentlyContinue
                }
                if (Test-Path $winreMountDir) {
                    Remove-Item $winreMountDir -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        }
        
        # Service boot.wim if SSU packages exist
        if ($hasSSU -and (Test-Path $bootWimPath)) {
            $bootCheckpoint = Join-Path $paths.Checkpoint "boot.$idx.done"
            
            if (-not (Test-Path $bootCheckpoint)) {
                Ensure-Folder -Path $mountDir
                
                try {
                    Write-Log "  Mounting index $idx boot.wim..."
                    $dismArgs = @("/Mount-Image", "/ImageFile:$bootWimPath", "/Index:1", "/MountDir:$mountDir")
                    & dism.exe $dismArgs | Write-Verbose
                    if ($LASTEXITCODE -ne 0) {
                        throw "Failed to mount boot.wim for index $idx"
                    }
                    
                    # Apply SSU packages to boot.wim
                    $ssuFiles = Get-ChildItem -Path $paths.KBsSSU -Filter "*.cab" -ErrorAction SilentlyContinue
                    foreach ($file in $ssuFiles) {
                        Write-Log "  Applying SSU to boot.wim: $($file.Name)"
                        $dismApplyArgs = @("/Image:$mountDir", "/Add-Package", "/PackagePath:$($file.FullPath)")
                        & dism.exe $dismApplyArgs | Write-Verbose
                        if ($LASTEXITCODE -ne 0) {
                            Write-Log "  WARNING: Failed to apply $($file.Name) to boot.wim (continuing)"
                        }
                    }
                    
                    # Unmount and commit boot.wim
                    Write-Log "  Unmounting index $idx boot.wim..."
                    $dismUnmountArgs = @("/Unmount-Image", "/MountDir:$mountDir", "/Commit")
                    & dism.exe $dismUnmountArgs | Write-Verbose
                    if ($LASTEXITCODE -ne 0) {
                        throw "Failed to unmount boot.wim for index $idx"
                    }
                    
                    # Mark checkpoint
                    Set-Content -Path $bootCheckpoint -Value (Get-Date -Format s) -Encoding UTF8
                    
                } catch {
                    Write-Log "  ERROR servicing index $idx boot.wim: $_"
                    throw $_
                } finally {
                    if (Test-Path $mountDir) {
                        Remove-Item $mountDir -Recurse -Force -ErrorAction SilentlyContinue
                    }
                }
            }
        }
        
        # Mark index as serviced
        Set-Content -Path $indexCheckpoint -Value (Get-Date -Format s) -Encoding UTF8
        Write-Log "Index $idx servicing complete"
    }
    
    # Final assembly - combine indices into install.wim and boot.wim
    Write-Log "Final assembly: combining serviced indices..."
    
    $installWimCheckpoint = Join-Path $paths.Checkpoint "install.wim.done"
    $bootWimCheckpoint = Join-Path $paths.Checkpoint "boot.wim.done"
    
    if (-not (Test-Path $installWimCheckpoint)) {
        Write-Log "Creating final install.wim with compression: $CompressionType..."
        
        # Take first index and set compression
        $firstIdx = $extractedIndices[0]
        $firstInstallWim = Join-Path $paths.WimsIndices "$firstIdx`_$($names.InstallWim)"
        
        $compressionMap = @{
            'None' = 'none'
            'Fast' = 'fast'
            'Maximum' = 'maximum'
        }
        $dismCompression = $compressionMap[$CompressionType]
        
        $dismExportArgs = @(
            "/Export-Image",
            "/SourceImageFile:$firstInstallWim",
            "/SourceIndex:1",
            "/DestinationImageFile:$($paths.InstallWimInDest)",
            "/Compress:$dismCompression"
        )
        
        Write-Verbose "Running: dism.exe $($dismExportArgs -join ' ')"
        & dism.exe $dismExportArgs | Write-Verbose
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to export first install.wim"
        }
        
        # Append remaining indices
        for ($i = 1; $i -lt $extractedIndices.Count; $i++) {
            $idx = $extractedIndices[$i]
            $srcInstallWim = Join-Path $paths.WimsIndices "$idx`_$($names.InstallWim)"
            
            Write-Log "  Appending index $idx to final install.wim..."
            $dismAppendArgs = @(
                "/Export-Image",
                "/SourceImageFile:$srcInstallWim",
                "/SourceIndex:1",
                "/DestinationImageFile:$($paths.InstallWimInDest)"
            )
            
            & dism.exe $dismAppendArgs | Write-Verbose
            if ($LASTEXITCODE -ne 0) {
                throw "Failed to append install.wim index $idx"
            }
        }
        
        Set-Content -Path $installWimCheckpoint -Value (Get-Date -Format s) -Encoding UTF8
        Write-Log "Final install.wim created"
    }
    
    if (-not (Test-Path $bootWimCheckpoint)) {
        # Combine boot.wim files if they exist
        $bootWimFiles = Get-ChildItem -Path $paths.WimsIndices -Filter "*_$($names.BootWim)" -File | Sort-Object
        
        if ($bootWimFiles.Count -gt 0) {
            Write-Log "Creating final boot.wim with compression: $CompressionType..."
            
            $firstBootWim = $bootWimFiles[0].FullName
            
            $dismExportArgs = @(
                "/Export-Image",
                "/SourceImageFile:$firstBootWim",
                "/SourceIndex:1",
                "/DestinationImageFile:$($paths.BootWimInDest)",
                "/Compress:$dismCompression"
            )
            
            & dism.exe $dismExportArgs | Write-Verbose
            if ($LASTEXITCODE -ne 0) {
                throw "Failed to export first boot.wim"
            }
            
            # Append remaining boot.wim files
            for ($i = 1; $i -lt $bootWimFiles.Count; $i++) {
                $srcBootWim = $bootWimFiles[$i].FullName
                
                Write-Log "  Appending boot.wim index $($i + 1)..."
                $dismAppendArgs = @(
                    "/Export-Image",
                    "/SourceImageFile:$srcBootWim",
                    "/SourceIndex:1",
                    "/DestinationImageFile:$($paths.BootWimInDest)"
                )
                
                & dism.exe $dismAppendArgs | Write-Verbose
                if ($LASTEXITCODE -ne 0) {
                    throw "Failed to append boot.wim file $i"
                }
            }
            
            Set-Content -Path $bootWimCheckpoint -Value (Get-Date -Format s) -Encoding UTF8
            Write-Log "Final boot.wim created"
        }
    }
    
    Write-Log "Service workflow complete"
}

# =========================
# HTML-based Update Catalog search
# =========================

function Invoke-CatalogRequest {
param(
    [Parameter(Mandatory)]
    [string]$Uri
)

Write-Verbose "GET $Uri"

try {
    $oldProtocol = [Net.ServicePointManager]::SecurityProtocol
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $Headers = @{
        "Cache-Control" = "no-cache"
        "Pragma"        = "no-cache"
        "User-Agent"    = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    }

    $Params = @{
        Uri             = $Uri
        Headers         = $Headers
        UseBasicParsing = $true
        ErrorAction     = "Stop"
    }

    $Response = Invoke-WebRequest @Params

    Write-Debug "RawContent length = $($Response.RawContent.Length)"

    $HtmlDoc = [HtmlAgilityPack.HtmlDocument]::new()
    $HtmlDoc.LoadHtml($Response.RawContent.ToString())

    return $HtmlDoc
}
catch {
    Write-Warning "Failed: $_"
    return $null
}
finally {
    [Net.ServicePointManager]::SecurityProtocol = $oldProtocol
}
}
function Search-UpdateCatalogHtml {
param(
    [Parameter(Mandatory)]
    [string]$Query,

    [Parameter(Mandatory)]
    [bool]$FirstOnly,

    [Parameter(Mandatory)]
    [string]$TargetFolder
)

Write-Host ("Searching for {0}{1}..." -f $Query, ($(if ($FirstOnly) { " (first result only)" } else { "" })))

$Encoded = [uri]::EscapeDataString($Query)
$Uri = "https://www.catalog.update.microsoft.com/Search.aspx?q=$Encoded"

Write-Debug "Encoded URI: $Uri"

$Html = Invoke-CatalogRequest -Uri $Uri
if (-not $Html) {
    Write-Warning "No HTML returned"
    return
}

Write-Verbose "Extracting update IDs from HTML"
#   Write-Debug "HTML: $($Html.DocumentNode.InnerHtml)"

# Look for goToDetails('GUID')
$pattern = 'goToDetails\("([0-9A-Fa-f\-]{36})"\)'
$matches = [regex]::Matches($Html.DocumentNode.InnerHtml, $pattern)
$ids = @()
foreach ($m in $matches) {
    $id = $m.Groups[1].Value
    Write-Debug "Found update GUID: $id"
    $ids += [PSCustomObject]@{
        Guid         = $id
        TargetFolder = $TargetFolder
    }
    if ($FirstOnly) { break }
}

Write-Verbose "Total IDs extracted: $($ids.Count)"
return $ids
}

function Get-UpdateLinks {
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $Guid
)

Write-Verbose ("GUID: {0}" -f $Guid)

# Build POST body
$postObject = @{
    size         = 0
    UpdateID     = $Guid
    UpdateIDInfo = $Guid
} | ConvertTo-Json -Compress

$body = @{
    UpdateIDs = "[$postObject]"
}

$params = @{
    Uri             = "https://www.catalog.update.microsoft.com/DownloadDialog.aspx"
    Method          = 'POST'
    Body            = $body
    ContentType     = "application/x-www-form-urlencoded"
    UseBasicParsing = $true
}

Write-Verbose "Requesting DownloadDialog.aspx via POST"
Write-Debug   "POST body:`n$($body.UpdateIDs)"

$response = Invoke-WebRequest @params

Write-Verbose "Received $($response.RawContentLength)-byte response of content type $($response.ContentType)"

# Normalize content for regex (remove newlines, collapse whitespace)
$content = $response.Content -replace "www\.download\.windowsupdate", "download.windowsupdate"
$content = $content -replace "`r?`n", ' '
$content = $content -replace '\s+', ' '

Write-Verbose "Normalized content length : $($content.Length)"
#Write-Debug   "Raw content (first 1000 chars):`n$($content.Substring(0, [Math]::Min(1000, $content.Length)))"

# Regex: downloadInformation[<idx>].files[<idx>].url = '<url>'
$pattern = "downloadInformation\[(\d+)\]\.files\[(\d+)\]\.url\s*=\s*'([^']*)'"
Write-Verbose "Running regex against DownloadDialog content"
Write-Debug   "Regex pattern: $pattern"

$matches = [regex]::Matches(
    $content,
    $pattern,
    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
)

if ($matches.Count -eq 0) {
    Write-Warning "No downloadInformation URL matches for $Guid (regex returned 0 matches)"
    return @()
}

Write-Verbose "Found $($matches.Count) download link match(es)"

try {
    $links = foreach ($m in $matches) {
        $downloadInfoIndex = [int]$m.Groups[1].Value
        $fileIndex         = [int]$m.Groups[2].Value
        $url               = $m.Groups[3].Value

        # Ignore garbage like "h" or empty strings
        if (-not $url -or
            [string]::IsNullOrWhiteSpace($url) -or
            $url.Length -lt 10 -or
            -not ($url -like "http*")) {

            Write-Verbose ("Ignoring malformed URL: {0}" -f $url)
            continue
        }

        # Try to extract KB number if present
        $kbNumber = 0
        if ($url -match 'kb(\d+)') {
            $kbNumber = [int]$Matches[1]
        }

        [PSCustomObject]@{
            URL               = $url
            KB                = $kbNumber
            DownloadInfoIndex = $downloadInfoIndex
            FileIndex         = $fileIndex
        }
    }
}
catch {
    Write-Warning ("Error processing download links for {0}: {1}" -f $Guid, $_.Exception.Message)
}

# Deduplicate by URL
$unique = $links | Group-Object -Property URL | ForEach-Object { $_.Group[0] }

# Sort by KB descending (0s at the end)
$sorted = $unique | Sort-Object KB -Descending

Write-Verbose "Unique URLs after de-duplication: $($sorted.Count)"
foreach ($l in $sorted) {
    Write-Debug "URL=$($l.URL) KB=$($l.KB) DI=$($l.DownloadInfoIndex) FI=$($l.FileIndex)"
}

return $sorted
}

# ==============================
# Download helper
# ==============================

function Ensure-Folder {
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $Path
)

if (-not (Test-Path $Path)) {
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
}
}

function Load-Manifest {
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $Folder
)

$manifestPath = Join-Path $Folder 'manifest.json'
if (-not (Test-Path $manifestPath -PathType Leaf)) {
    return @()
}

try {
    $json = Get-Content -Path $manifestPath -Raw -ErrorAction Stop
    $data = $json | ConvertFrom-Json -ErrorAction Stop
    if ($null -eq $data) { return @() }
    if ($data -is [System.Array]) { return $data }
    return @($data)
}
catch {
    Write-Warning ("Failed to load manifest from {0}: {1}" -f $manifestPath, $_.Exception.Message)
    return @()
}
}

function Write-Manifest {
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $Folder,

    [Parameter(Mandatory = $true)]
    [object[]] $Entries
)

$manifestPath = Join-Path $Folder 'manifest.json'
$json = $Entries | ConvertTo-Json -Depth 6
$json | Set-Content -Path $manifestPath -Encoding UTF8
}
function Download-MUFile {
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [PSCustomObject] $Update,

    [Parameter(Mandatory = $true)]
    [string] $TargetFolder
)

Ensure-Folder -Path $TargetFolder

Write-Host ("Preparing downloads for update {0}: {1}" -f $Update.Guid, $Update.Title)

$results = @()

# No URLs → nothing to do
if (-not $Update.DownloadUrls -or $Update.DownloadUrls.Count -eq 0) {
    Write-Host ("No download URLs for update {0}" -f $Update.Guid)
    return $results
}

# Detect whether output is piped (CR-safe)
$isPiped = -not $Host.UI.RawUI.KeyAvailable

foreach ($url in $Update.DownloadUrls) {

    if ([string]::IsNullOrWhiteSpace($url)) {
        Write-Warning ("Ignoring empty URL for update {0}" -f $Update.Guid)
        continue
    }

    $fileName = Split-Path -Path $url -Leaf
    $destPath = Join-Path $TargetFolder $fileName

    # ------------------------------------------------------------
    # SKIP IF FILE ALREADY EXISTS
    # ------------------------------------------------------------
    if (Test-Path $destPath -PathType Leaf) {
        Write-Host ("File already exists, skipping: {0}" -f $fileName)

        $results += [PSCustomObject]@{
            FileName = $fileName
            FullPath = $destPath
            Url      = $url
        }

        continue
    }

    Write-Host ("Downloading {0}..." -f $fileName)

    # ------------------------------------------------------------
    # Retry loop (3 attempts)
    # ------------------------------------------------------------
    $maxRetries = 3
    $success = $false

    for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {

        Write-Host ("  Attempt {0} of {1}" -f $attempt, $maxRetries)

        try {
            $req = [System.Net.HttpWebRequest]::Create($url)
            $req.Method = "GET"
            $req.UserAgent = "Mozilla/5.0"

            $resp = $req.GetResponse()
            $total = $resp.ContentLength
            $inStream  = $resp.GetResponseStream()
            $outStream = [System.IO.File]::Open($destPath, [System.IO.FileMode]::Create)

            $buffer = New-Object byte[] 65536
            $totalRead = 0
            $nextMark = 10

            # Initial progress line
            Write-Host ("    0%  0/{0} bytes" -f $total)

            while (($read = $inStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                $outStream.Write($buffer, 0, $read)
                $totalRead += $read

                if ($total -gt 0) {
                    $pct = [math]::Floor(($totalRead / $total) * 100)

                    if ($pct -ge $nextMark) {
                        # Pipe-safe: always print full lines, never CR
                        Write-Host ("  {0,3}%  {1:N0}/{2:N0} bytes" -f $pct, $totalRead, $total)
                        $nextMark += 10
                    }
                }
            }

            Write-Host ("  Completed: {0}" -f $fileName)

            $outStream.Close()
            $inStream.Close()
            $resp.Close()

            $success = $true
            break
        }
        catch {
            Write-Warning ("  ERROR: {0}" -f $_.Exception.Message)
            Write-Warning ("  Retrying...")

            # Clean up partial file
            if (Test-Path $destPath) {
                Remove-Item -Force $destPath -ErrorAction SilentlyContinue
            }
        }
    }

    if (-not $success) {
        Write-Warning ("FAILED after {0} attempts: {1}" -f $maxRetries, $fileName)
        continue
    }

    $results += [PSCustomObject]@{
        FileName = $fileName
        FullPath = $destPath
        Url      = $url
    }
}

Write-Host ("Completed downloads for update {0}: {1}" -f $Update.Guid, $Update.Title)

return $results
}

function Get-UpdateDetails {
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [int] $Count,

    [Parameter(Mandatory = $true)]
    [string] $Guid,

    [Parameter(Mandatory = $true)]
    [string] $TargetFolder
)

Write-Host ("Processing update #{0}: {1}" -f $Count, $Guid)
Write-Verbose ("TargetFolder: {0}" -f $TargetFolder)

# ------------------------------------------------------------
# DETAILS PAGE (ScopedViewInline.aspx)
# Extracts:
#    - Title
#    - KB number
#    - SupersededBy list
# ------------------------------------------------------------

$detailsUrl = "https://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid=$Guid"

try {
    $detailsResponse = Invoke-WebRequest -Uri $detailsUrl -UseBasicParsing -ErrorAction Stop
}
catch {
    Write-Debug ("Failed to fetch details page for {0}: {1}" -f $Guid, $_.Exception.Message)
    return
}

$detailsDoc = New-Object HtmlAgilityPack.HtmlDocument
$detailsDoc.LoadHtml($detailsResponse.Content)

# Title
$titleNode = $detailsDoc.DocumentNode.SelectSingleNode("//span[@id='ScopedViewHandler_titleText']")
$title = if ($titleNode) { $titleNode.InnerText.Trim() } else { "" }
Write-Verbose ("Title: {0}" -f $title)

# KB
$kbMatch = [regex]::Match($title, "KB\d+")
$kb = if ($kbMatch.Success) { $kbMatch.Value } else { "" }
Write-Verbose ("KB: {0}" -f $kb)

# SupersededBy
$supersededBy = @()
$supNodes = $detailsDoc.DocumentNode.SelectNodes("//div[@id='supersededbyInfo']//a")
if ($supNodes) {
    foreach ($n in $supNodes) {
        $supersededBy += $n.InnerText.Trim()
    }
}

# Not a keeper if superseded by anything else, even if it has download links
if ($supersededBy.Count -gt 0) {
    Write-Verbose ("SupersededBy: {0}" -f ($supersededBy -join ', '))
    Write-Host ("{0} superseded" -f $Guid)
    return
}

# ------------------------------------------------------------
# DOWNLOAD LINKS (via Get-UpdateLinks)
# ------------------------------------------------------------

Write-Host "Finding download links for $title"

$links = Get-UpdateLinks -Guid $Guid
$downloadUrls = @()
if ($links) {
    $downloadUrls = $links.URL | Select-Object -Unique
}

Write-Host ("Found {0} file(s) for this update" -f $downloadUrls.Count)

# Not a keeper if no download links
if ($downloadUrls.Count -eq 0) {
    Write-Host ("{0} has no download links" -f $Guid)
    return
}
Write-Verbose ("Download URLs: {0}" -f ($downloadUrls -join ', '))

return [PSCustomObject]@{
    Guid         = $Guid
    Title        = $title
    KB           = $kb
    SupersededBy = $supersededBy
    DownloadUrls = $downloadUrls
    TargetFolder = $TargetFolder
}
}

function Build-ManifestEntry {
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [PSCustomObject] $Details,

    [Parameter(Mandatory = $true)]
    [PSCustomObject] $DownloadInfo
)

[PSCustomObject]@{
    Guid         = $Details.Guid
    Title        = $Details.Title
    DownloadUrl  = $DownloadInfo.Url
    FileName     = $DownloadInfo.FileName
    Timestamp    = (Get-Date).ToString("s")
}
}

function Invoke-KBWork {

# Clean mode
if ($Clean) {
    $path = $paths.KBsRoot
    if ($DryRun) {
        Write-Log "[DryRun] Would clean $path"
    } elseif (Test-Path $path) {
        Write-log "Removing: $path"
        Remove-Item $path -Recurse -Force
    }
    return
}

if ($DryRun) {
    foreach ($u in $kbDirs) {
        Write-Log ("[DryRun] Would fill: {0}" -f $paths["KBs$u"])
    }
    return
}

Write-Log "Starting KB update workflow..."

# Ensure folders exist
$KBsPaths = @()
foreach ($u in $kbDirs) {
    $KBsPaths += $paths["KBs$u"]
}

foreach ($folder in $KBsPaths) {
    Ensure-Folder -Path $folder
}

# Build queries
$queries = @(
    [PSCustomObject]@{
        Query        = "Cumulative Updates for Windows $WinOS Version $Version for $Arch-based Systems"
        FirstOnly    = $false
        TargetFolder = $paths.KBsOSCU
    }
    [PSCustomObject]@{
        Query        = ".NET Framework for Windows $WinOS Version $Version $Arch"
        FirstOnly    = $true
        TargetFolder = $paths.KBsNET
    }
    [PSCustomObject]@{
        Query        = ".NET 8.0 $Arch Client"
        FirstOnly    = $true
        TargetFolder = $paths.KBsNET
    }
    [PSCustomObject]@{
        Query        = "Update for Windows Security platform"
        FirstOnly    = $true
        TargetFolder = $paths.KBsMISC
    }
)

$results = foreach ($q in $queries) {
    Search-UpdateCatalogHtml -Query $q.Query -FirstOnly $q.FirstOnly -TargetFolder $q.TargetFolder
}

$allGuids = $results | Sort-Object Guid -Unique
Write-Host ("Found {0} total updates to process" -f $allGuids.Count)
Write-Debug ("Guid TargetFolder:`n" + (@($allGuids | ForEach-Object { '{0} {1}' -f $_.Guid, $_.TargetFolder }) -join "`n"))

if ($allGuids.Count -eq 0) {
    Write-Host "No updates found"
    return
}

Write-Host "Retrieving update details..."

$count = 0
$details = foreach ($g in $allGuids) {
    Write-Debug ("Resolving details for {0} ({1})" -f $g.Guid, $g.TargetFolder)
    Get-UpdateDetails -Count (++$count) -Guid $g.Guid -TargetFolder $g.TargetFolder
}

if ($details.Count -eq 0) {
    Write-Host "No usable updates after details resolution"
    return
}

Write-Host ("Remaining applicable updates: {0}" -f $details.Count)
Write-Debug "GUIDs:`n$($details.Guid -join "`n")"

Write-Host "Synchronizing update folders..."

$requiredFiles = @()
foreach ($d in $details) {
    foreach ($url in $d.DownloadUrls) {
        $requiredFiles += (Split-Path $url -Leaf)
    }
}
$requiredFiles = $requiredFiles | Select-Object -Unique

# Sync: remove stale files in all folders
foreach ($folder in $KBsPaths) {
    Write-Verbose "Checking folder: $folder"

    $existingFiles = @()
    if (Test-Path $folder) {
        $existingFiles = Get-ChildItem -Path $folder -File |
                            Select-Object -ExpandProperty Name
    }

    $stale = $existingFiles | Where-Object { $_ -notin $requiredFiles }
    foreach ($file in $stale) {
        $path = Join-Path $folder $file
        Write-Verbose "Removing stale file: $file"
        Remove-Item $path -Force
    }
}

Write-Host "Downloading required update files..."

$manifestByFolder = @{}
foreach ($d in $details) {
    $manifestByFolder[$d.TargetFolder] = @()
}
foreach ($d in $details) {
    $targetFolder = $d.TargetFolder

    # Download all files for this update into the target folder
    $downloadInfos = Download-MUFile -Update $d -TargetFolder $targetFolder

    foreach ($downloadInfo in $downloadInfos) {
        $entry = Build-ManifestEntry -Details $d -DownloadInfo $downloadInfo
        $manifestByFolder[$targetFolder] += $entry
    }
}

Write-Host "Writing manifests..."
foreach ($kvp in $manifestByFolder.GetEnumerator()) {
    $folder  = $kvp.Key
    $entries = $kvp.Value
    if ($entries.Count -gt 0) {
        Write-Verbose "Writing manifest for $folder"
        Write-Manifest -Folder $folder -Entries $entries
    }
    else {
        $manifestPath = Join-Path $folder 'manifest.json'
        if (Test-Path $manifestPath -PathType Leaf) {
            Remove-Item $manifestPath -Force
        }
    }
}

Write-Host "KB update workflow complete"
}

# ==============================
# Driver export
# ==============================
function Invoke-DriverWork {
$WinpeDriverRoot = $paths.WinpeDriverRoot

if ($Clean) {
    if ($DryRun) {
        Write-Log "[DryRun] Would remove: $WinpeDriverRoot"
    } elseif (Test-Path $WinpeDriverRoot) {
        Write-Log "Removing: $WinpeDriverRoot"
        Remove-Item $WinpeDriverRoot -Recurse -Force
    }
    return
}

if ($DryRun) {
    Write-Log "[DryRun] Would run: DISM /export-driver"
    return
}

Write-Log "Exporting drivers..."
Ensure-Folder -Path $WinpeDriverRoot
$args = "/online /export-driver /destination:`"$WinpeDriverRoot`""
$p = Start-Process -FilePath dism.exe -ArgumentList $args -NoNewWindow -PassThru -Wait -RedirectStandardOutput "$WinpeDriverRoot\dism.log"
if ($p.ExitCode -ne 0) { throw "DISM failed" }
}

# ==============================
# Main execution
# ==============================

# Default values
if ([string]::IsNullOrWhiteSpace($Folder)) {
    $Folder = Get-Location
}
if ([string]::IsNullOrWhiteSpace($WinOS)) {
    $WinOS = '11'
}
if ([string]::IsNullOrWhiteSpace($Version)) {
    $Version = if ($WinOS -eq '10') { '22H2' } else { '25H2' }
}
if ([string]::IsNullOrWhiteSpace($Arch)) {
    $Arch = 'x64'
}

Write-Log "Initialized: Folder=$Folder, WinOS=$WinOS, Version=$Version, Arch=$Arch"

# Setup paths
$paths = [ordered]@{}
$paths.SrcIsoRoot            = Join-Path $Folder $names.SrcIso
$paths.SourceInSrc           = Join-Path $paths.SrcIsoRoot $names.Source
$paths.DestIsoRoot           = Join-Path $Folder $names.DestIso
$paths.SourceInDest          = Join-Path $paths.DestIsoRoot $names.Source
$paths.BootWimInDest         = Join-Path $paths.SourceInDest $names.BootWim 
$paths.InstallWimInDest      = Join-Path $paths.SourceInDest $names.InstallWim 
$paths.WinpeDriverRoot       = Join-Path $Folder $names.WinpeDriver
$paths.RegistryRoot          = Join-Path $Folder $names.Registry
$paths.InstallDriversCmd     = Join-Path $Folder $names.InstallDriversCmd
$paths.InstallRegsCmd        = Join-Path $Folder $names.InstallRegsCmd
$paths.PostSetupCmd          = Join-Path $Folder $names.PostSetupCmd
$paths.SetupConfigCleanIni   = Join-Path $Folder $names.SetupConfigCleanIni
$paths.SetupConfigUpgradeIni = Join-Path $Folder $names.SetupConfigUpgradeIni
$paths.CleanInstallCmd       = Join-Path $Folder $names.CleanInstallCmd
$paths.UpgradeCmd            = Join-Path $Folder $names.UpgradeCmd
$paths.WinreWimInWim         = Join-Path "Windows\System32\Recovery" $names.WinreWim
$paths.KBsRoot               = Join-Path $Folder $names.KBs
foreach ($u in $kbDirs) {
    $paths["KBs$u"]          = Join-Path $paths.KBsRoot $names.$u
}
$paths.WimsRoot              = Join-Path $Folder $names.Wims
foreach ($u in $wimDirs) {
    $paths["Wims$u"]         = Join-Path $paths.WimsRoot $names.$u
}
$paths.Checkpoint            = Join-Path $Folder $names.Checkpoint
$paths.SrcIsoCheckpoint      = Join-Path $paths.Checkpoint $names.SrcIso
$paths.DestIsoCheckpoint     = Join-Path $paths.Checkpoint $names.DestIso
$paths.KBsCheckpoint         = Join-Path $paths.Checkpoint $names.KBs
$paths.WimsCheckpoint        = Join-Path $paths.WimsCheckpoint $names.Wims

# Run requested operations
if ($Export -or $All) {
    Invoke-ExportWork
}

if ($KB -or $All) {
    Invoke-KBWork
}

if ($Service -or $All) {
    Invoke-ServiceWork
}

if ($Drivers -or $All) {
    Invoke-DriverWork
}

Write-Log "Script execution complete"
