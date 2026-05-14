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

.PARAMETER ShowIndices
Shows available image indices (index and name) and exits.

.PARAMETER Home
Select editions whose normalized label matches "Home" exactly.

.PARAMETER Pro
Select editions whose normalized label matches "Pro" exactly.

.PARAMETER Indices
Comma-separated selector string supporting:
- numbers: 6
- ranges: 3-6, 7-*
- exact labels: "Education N"
- wildcard labels: "*Home*", "* N*"
- regex labels: "re:^Education( N)?$"

.PARAMETER ISO
Explicit path to source ISO.
If omitted, the script discovers the single .iso file in <Folder>.
If more than one .iso is present an error is raised; use this parameter to disambiguate.

.PARAMETER DestISO
Explicit path to destination ISO.
If omitted, the source ISO path is reused with the extension changed to _KBs.iso.

.PARAMETER ShowIndices
Print the available image indices (index number and name) from the source ISO and exit.

.PARAMETER CreateISO
Reuse an existing work folder from a prior run.
Without explicit index selection (-Home/-Pro/-Indices) no rebuild/service/DU actions are performed.

.PARAMETER UseADK
Prefer ADK dism.exe and oscdimg.exe when available.

.PARAMETER UseSystem
Force system dism.exe and PATH oscdimg.exe.

.PARAMETER dism
Explicit path to dism.exe.

.PARAMETER oscdimg
Explicit path to oscdimg.exe.

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

    [Parameter(Position = 1)]
    [string]$ISO,

    [Parameter(Position = 2)]
    [string]$DestISO,

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

    [switch]$ShowIndices,

    [Alias('Home')]
    [switch]$SelectHome,

    [Alias('Pro')]
    [switch]$SelectPro,

    [string]$Indices,

    [switch]$CreateISO,

    [switch]$UseADK,

    [switch]$UseSystem,

    [string]$dism,

    [string]$oscdimg,

    [switch]$Clean,

    [switch]$DryRun,

    [switch]$Help
)

# git hash
$GitHash = "84a4b09"

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
    Sources               = 'sources'
    BootWim               = 'boot.wim'
    InstallEsd            = 'install.esd'
    InstallWim            = 'install.wim'
    WinreWim              = 'winre.wim'
    BootFileBIOS          = 'boot\etfsboot.com'
    BootFileUEFI          = 'efi\microsoft\boot\efisys.bin'
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

# ==============================
# Tool discovery
# ==============================

function Resolve-DismExe {
    [CmdletBinding()]
    param(
        [string]$ExplicitPath,
        [switch]$PreferADK,
        [switch]$ForceSystem
    )

    Write-Debug "Resolve-DismExe: ExplicitPath='$ExplicitPath' PreferADK=$PreferADK ForceSystem=$ForceSystem"

    if ($ExplicitPath) {
        if (-not (Test-Path $ExplicitPath)) {
            throw "Explicit dism path not found: $ExplicitPath"
        }
        Write-Verbose "Using explicit dism: $ExplicitPath"
        return $ExplicitPath
    }

    $adkRoots = @(
        "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools",
        "${env:ProgramFiles}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools"
    )
    $adkArches = @('amd64', 'arm64', 'x86')

    $adkDism = $null
    foreach ($root in $adkRoots) {
        foreach ($a in $adkArches) {
            $candidate = Join-Path $root "$a\DISM\dism.exe"
            if (Test-Path $candidate) { $adkDism = $candidate; break }
        }
        if ($adkDism) { break }
    }

    $systemDism = Join-Path $env:SystemRoot 'System32\dism.exe'

    if ($ForceSystem) {
        if (Test-Path $systemDism) {
            Write-Verbose "Using system dism (forced): $systemDism"
            return $systemDism
        }
        throw "System dism.exe not found at: $systemDism"
    }

    if ($PreferADK -and $adkDism) {
        Write-Verbose "Using ADK dism (preferred): $adkDism"
        return $adkDism
    }

    if ($adkDism) {
        Write-Verbose "Using ADK dism (auto-discovered): $adkDism"
        return $adkDism
    }

    if (Test-Path $systemDism) {
        Write-Verbose "Using system dism (fallback): $systemDism"
        return $systemDism
    }

    throw "dism.exe not found. Install Windows ADK or specify -dism."
}

function Resolve-OscdimgExe {
    [CmdletBinding()]
    param(
        [string]$ExplicitPath,
        [switch]$PreferADK,
        [switch]$ForceSystem
    )

    Write-Debug "Resolve-OscdimgExe: ExplicitPath='$ExplicitPath' PreferADK=$PreferADK ForceSystem=$ForceSystem"

    if ($ExplicitPath) {
        if (-not (Test-Path $ExplicitPath)) {
            throw "Explicit oscdimg path not found: $ExplicitPath"
        }
        Write-Verbose "Using explicit oscdimg: $ExplicitPath"
        return $ExplicitPath
    }

    $adkRoots = @(
        "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools",
        "${env:ProgramFiles}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools"
    )
    $adkArches = @('amd64', 'arm64', 'x86')

    $adkOscdimg = $null
    foreach ($root in $adkRoots) {
        foreach ($a in $adkArches) {
            $candidate = Join-Path $root "$a\Oscdimg\oscdimg.exe"
            if (Test-Path $candidate) { $adkOscdimg = $candidate; break }
        }
        if ($adkOscdimg) { break }
    }

    if ($ForceSystem) {
        $pathCmd = Get-Command 'oscdimg.exe' -ErrorAction SilentlyContinue
        if ($pathCmd) {
            Write-Verbose "Using PATH oscdimg (forced): $($pathCmd.Source)"
            return $pathCmd.Source
        }
        Write-Warning "oscdimg.exe not found in PATH"
        return $null
    }

    if ($PreferADK -and $adkOscdimg) {
        Write-Verbose "Using ADK oscdimg (preferred): $adkOscdimg"
        return $adkOscdimg
    }

    if ($adkOscdimg) {
        Write-Verbose "Using ADK oscdimg (auto-discovered): $adkOscdimg"
        return $adkOscdimg
    }

    $pathCmd = Get-Command 'oscdimg.exe' -ErrorAction SilentlyContinue
    if ($pathCmd) {
        Write-Verbose "Using PATH oscdimg: $($pathCmd.Source)"
        return $pathCmd.Source
    }

    Write-Warning "oscdimg.exe not found. ISO creation will not be available."
    return $null
}

# ==============================
# ISO / WIM introspection
# ==============================

function Get-WimImageList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$WimPath
    )

    Write-Debug "Get-WimImageList: WimPath='$WimPath'"
    Write-Verbose "Reading image list from: $WimPath"

    $output = & $dismExe /Get-WimInfo "/WimFile:$WimPath" 2>&1

    if ($LASTEXITCODE -ne 0) {
        Write-Warning "DISM /Get-WimInfo failed (exit $LASTEXITCODE) for: $WimPath"
        return @()
    }

    $images      = [System.Collections.Generic.List[object]]::new()
    $currentIdx  = $null
    $currentName = $null

    foreach ($line in $output) {
        Write-Debug "  WimInfo> $line"
        if ($line -match '^\s*Index\s*:\s*(\d+)') {
            if ($null -ne $currentIdx) {
                $images.Add([PSCustomObject]@{ Index = $currentIdx; Name = $currentName })
            }
            $currentIdx  = [int]$Matches[1]
            $currentName = ''
        }
        elseif ($null -ne $currentIdx -and $line -match '^\s*Name\s*:\s*(.+)') {
            $currentName = $Matches[1].Trim()
        }
    }
    if ($null -ne $currentIdx) {
        $images.Add([PSCustomObject]@{ Index = $currentIdx; Name = $currentName })
    }

    Write-Verbose "Found $($images.Count) image(s) in WIM"
    return $images.ToArray()
}

function Get-ISOMetadataFromWim {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$WimPath
    )

    Write-Debug "Get-ISOMetadataFromWim: WimPath='$WimPath'"

    $output = & $dismExe /Get-WimInfo "/WimFile:$WimPath" /Index:1 2>&1

    $buildNumber = 0
    $archStr     = 'x64'

    foreach ($line in $output) {
        if ($line -match '^\s*Version\s*:\s*\d+\.\d+\.(\d+)\.') {
            $buildNumber = [int]$Matches[1]
        }
        if ($line -match '^\s*Architecture\s*:\s*(.+)') {
            $archStr = $Matches[1].Trim()
        }
    }

    Write-Debug "  build=$buildNumber arch='$archStr'"

    $detectedWinOS = if ($buildNumber -ge 22000) { '11' } else { '10' }

    $detectedVersion = switch ($buildNumber) {
        { $_ -ge 26200 } { '25H2'; break }
        { $_ -ge 26100 } { '24H2'; break }
        { $_ -ge 22631 } { '23H2'; break }
        { $_ -ge 22621 } { '22H2'; break }
        { $_ -ge 22000 } { '21H2'; break }
        { $_ -ge 19045 } { '22H2'; break }
        { $_ -ge 19044 } { '21H2'; break }
        { $_ -ge 19043 } { '21H1'; break }
        { $_ -ge 19042 } { '20H2'; break }
        { $_ -ge 19041 } { '2004'; break }
        default           { if ($detectedWinOS -eq '11') { '25H2' } else { '22H2' }; break }
    }

    $detectedArch = switch -Wildcard ($archStr.ToLower()) {
        '*arm64*' { 'arm64' }
        '*amd64*' { 'x64'   }
        '*x64*'   { 'x64'   }
        default   { 'x64'   }
    }

    return [PSCustomObject]@{
        WinOS   = $detectedWinOS
        Version = $detectedVersion
        Arch    = $detectedArch
        Build   = $buildNumber
    }
}

# ==============================
# Index selection
# ==============================

function Resolve-IndexSelection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$AllImages,
        [switch]$SelectHome,
        [switch]$SelectPro,
        [string]$IndicesStr
    )

    Write-Debug "Resolve-IndexSelection: Home=$SelectHome Pro=$SelectPro Indices='$IndicesStr' TotalImages=$($AllImages.Count)"

    function Get-NormalizedLabel([string]$name) {
        ($name -replace '^Windows\s+(10|11)\s+', '').Trim()
    }

    $anyExplicit = $SelectHome -or $SelectPro -or $IndicesStr

    if (-not $anyExplicit) {
        Write-Verbose "No explicit index selection; returning all $($AllImages.Count) indices"
        return $AllImages
    }

    $selected = [System.Collections.Generic.List[object]]::new()

    if ($SelectHome) {
        Write-Verbose "Selecting 'Home' editions"
        $AllImages | Where-Object { (Get-NormalizedLabel $_.Name) -eq 'Home' } | ForEach-Object { $selected.Add($_) }
    }

    if ($SelectPro) {
        Write-Verbose "Selecting 'Pro' editions"
        $AllImages | Where-Object { (Get-NormalizedLabel $_.Name) -eq 'Pro' } | ForEach-Object { $selected.Add($_) }
    }

    if ($IndicesStr) {
        $tokens = $IndicesStr -split '\s*,\s*'
        foreach ($token in $tokens) {
            $token = $token.Trim().Trim('"').Trim("'")
            Write-Verbose "  Processing token: '$token'"

            if ($token -match '^(\d+)-(\*|\d+)$') {
                $from = [int]$Matches[1]
                $to   = if ($Matches[2] -eq '*') { [int]::MaxValue } else { [int]$Matches[2] }
                Write-Debug "    Range $from-$to"
                $AllImages | Where-Object { $_.Index -ge $from -and $_.Index -le $to } | ForEach-Object { $selected.Add($_) }
            }
            elseif ($token -match '^\d+$') {
                Write-Debug "    Single index $token"
                $AllImages | Where-Object { $_.Index -eq [int]$token } | ForEach-Object { $selected.Add($_) }
            }
            elseif ($token -match '^re:(.+)$') {
                $pattern = $Matches[1]
                Write-Debug "    Regex '$pattern'"
                $AllImages | Where-Object { $_.Name -match $pattern } | ForEach-Object { $selected.Add($_) }
            }
            elseif ($token -match '[*?]') {
                Write-Debug "    Wildcard '$token'"
                $AllImages | Where-Object { $_.Name -like $token } | ForEach-Object { $selected.Add($_) }
            }
            else {
                Write-Debug "    Exact label '$token'"
                $AllImages | Where-Object { $_.Name -eq $token } | ForEach-Object { $selected.Add($_) }
            }
        }
    }

    $result = @($selected | Sort-Object Index -Unique)
    Write-Verbose "Index selection resolved to $($result.Count) index/indices: $($result.Index -join ', ')"
    return $result
}

# ==============================
# Hardlink tree copy
# ==============================

function New-HardLinkTree {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Source,
        [Parameter(Mandatory)]
        [string]$Destination,
        [string[]]$ExcludeFileNames = @()
    )

    Write-Verbose "New-HardLinkTree: '$Source' -> '$Destination'"
    Write-Debug   "New-HardLinkTree: ExcludeFileNames=[$($ExcludeFileNames -join ', ')]"

    Ensure-Folder -Path $Destination

    $allFiles = @(Get-ChildItem -Path $Source -Recurse -File -ErrorAction SilentlyContinue)
    $total    = $allFiles.Count
    $done     = 0
    $lastPct  = -1

    Write-Output ("Hardlinking {0} files: '{1}' -> '{2}'" -f $total, $Source, $Destination)

    foreach ($file in $allFiles) {
        $done++
        $pct = [math]::Floor(($done / [math]::Max($total, 1)) * 100)
        if ($pct -ge ($lastPct + 10)) {
            Write-Output ("  {0,3}%  {1}/{2} files" -f $pct, $done, $total)
            $lastPct = $pct - ($pct % 10)
        }

        if ($file.Name -in $ExcludeFileNames) {
            Write-Debug "  Skip (excluded): $($file.Name)"
            continue
        }

        $relPath  = $file.FullName.Substring($Source.TrimEnd('\', '/').Length).TrimStart('\', '/')
        $destPath = Join-Path $Destination $relPath
        $destDir  = Split-Path $destPath -Parent

        Ensure-Folder -Path $destDir

        if (Test-Path $destPath) {
            Write-Debug "  Already exists: $relPath"
            continue
        }

        Write-Debug "  Hardlink: $relPath"
        try {
            New-Item -ItemType HardLink -Path $destPath -Value $file.FullName -Force -ErrorAction Stop | Out-Null
        }
        catch {
            Write-Warning "Hardlink failed for '$relPath'; falling back to copy: $_"
            Copy-Item -Path $file.FullName -Destination $destPath -Force
        }
    }

    Write-Output ("  Hardlink tree complete: {0} files processed" -f $done)
}

# =========================
# Extract/Export section
# =========================

function Invoke-ExportWork {
    [CmdletBinding()]
    param()

    if ($DryRun) {
        Write-Output "[DryRun] Would mount ISO: $ISO"
        Write-Output "[DryRun] Would validate $($names.BootFileBIOS) in $ISO"
        Write-Output "[DryRun] Would validate $($names.BootFileUEFI) in $ISO"
        Write-Output "[DryRun] Would validate $($paths.BootWimInIso) in $ISO"
        Write-Output "[DryRun] Would validate $($paths.InstallWimInIso) or $($paths.InstallEsdInIso) in $ISO"
        Write-Output "[DryRun] Would robocopy tree -> $($paths.SrcIsoRoot)"
        Write-Output "[DryRun] Would hardlink-copy $($paths.SrcIsoRoot) -> $($paths.DestIsoRoot) (excluding $($names.BootWim), $($names.InstallWim), $($names.InstallEsd))"
        Write-Output "[DryRun] Would export $($SelectedIndices.Count) indices to $($paths.WimsIndices)"
        return
    }

    if ($Clean) {
        Write-Debug "Invoke-ExportWork: Clean mode"
        foreach ($cleanPath in @($paths.SrcIsoRoot, $paths.DestIsoRoot, $paths.WimsRoot)) {
            if (Test-Path $cleanPath) {
                Write-Output "Removing: $cleanPath"
                Remove-Item $cleanPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
        foreach ($ckptName in @('srciso.copy.done', 'destiso.copy.done', 'export.done')) {
            $ckpt = Join-Path $paths.Checkpoint $ckptName
            if (Test-Path $ckpt) { Remove-Item $ckpt -Force }
        }
        return
    }

    Write-Output  "Starting Export workflow..."
    Write-Verbose "Invoke-ExportWork: ISO='$ISO' SrcIsoRoot='$($paths.SrcIsoRoot)'"
    Write-Debug   "Invoke-ExportWork: SourcesInSrc='$($paths.SourcesInSrc)' DestIsoRoot='$($paths.DestIsoRoot)' WimsIndices='$($paths.WimsIndices)'"

    Ensure-Folder -Path $paths.Checkpoint
    Ensure-Folder -Path $paths.SrcIsoRoot

    # -----------------------------------------------------------------------
    # Step 1: Mount the source ISO and robocopy its contents to SrcIsoRoot
    # -----------------------------------------------------------------------
    $srcIsoCopyCheckpoint = Join-Path $paths.Checkpoint "srciso.copy.done"

    if (Test-Path $srcIsoCopyCheckpoint) {
        Write-Output "Source ISO copy already completed (checkpoint: srciso.copy.done)"
        Write-Debug  "srciso.copy.done timestamp: $(Get-Content $srcIsoCopyCheckpoint -Raw -ErrorAction SilentlyContinue)"
    } else {
        if (-not $ISO -or -not (Test-Path $ISO)) {
            throw "Source ISO not found or not specified. Use -ISO to point to your Windows .iso file."
        }

        Write-Output  "Mounting ISO: $ISO"
        $diskImage = Mount-DiskImage -ImagePath $ISO -PassThru -ErrorAction Stop

        try {
            # Wait for the volume to actually appear
            $vol = $null
            $retryCount = 0
            while ($null -eq $vol -and $retryCount -lt 5) {
                $vol = $diskImage | Get-Volume -ErrorAction SilentlyContinue
                if ($null -eq $vol) { 
                    Start-Sleep -Seconds 1 
                    $retryCount++
                }
            }

            if ($null -eq $vol) { throw "Timeout waiting for ISO volume to initialize." }

            # Create a 'safe' path for Test-Path
            $driveLetterRaw = $vol.DriveLetter + ":"      # e.g., "D:"
            $driveLetter    = $driveLetterRaw + "\"       # e.g., "D:\" (for Test-Path)

            Write-Output "ISO mounted at: $driveLetter"

            # ---- Validate required files in ISO ----
            $missing = @()
            if (-not (Test-Path (Join-Path $driveLetter $names.BootFileBIOS))) { $missing += $names.BootFileBIOS }
            if (-not (Test-Path (Join-Path $driveLetter $names.BootFileUEFI))) { $missing += $names.BootFileUEFI }
            if (-not (Test-Path (Join-Path $driveLetter $paths.BootWimInIso))) { $missing += $paths.BootWimInIso }
            if (-not ((Test-Path (Join-Path $driveLetter $paths.InstallWimInIso)) -or (Test-Path (Join-Path $driveLetter $paths.InstallEsdInIso)))) {
                $missing += "$($paths.InstallWimInIso) or $($paths.InstallEsdInIso)"
            }

            if ($missing.Count -gt 0) {
                throw "Source ISO validation failed. Missing: $($missing -join ', ')"
            }
            Write-Output "Source ISO validation passed"

            # Copy the ISO root tree to SrcIsoRoot
            Write-Output "Copying ISO root tree to $($paths.SrcIsoRoot)..."

            $roboRootArgs = @(
                $driveLetterRaw,
                $paths.SrcIsoRoot,
                '/E',
                '/R:2', '/W:1', '/NP', '/NDL', '/NC'
            )
            Write-Verbose "robocopy $($roboRootArgs -join ' ')"
            robocopy @roboRootArgs

            # Robocopy exit codes 0-7 are all 'Success' variants
            if ($LASTEXITCODE -ge 8) {
                throw "robocopy failed with exit code $LASTEXITCODE"
            }
        } finally {
            if (Get-DiskImage -ImagePath $ISO | Where-Object { $_.Attached }) {
                Write-Output "Unmounting ISO..."
                Dismount-DiskImage -ImagePath $ISO -ErrorAction SilentlyContinue | Out-Null
            }
        }
        Set-Content -Path $srcIsoCopyCheckpoint -Value (Get-Date -Format s) -Encoding UTF8
        Write-Output "Source ISO copy complete (checkpoint: srciso.copy.done)"
    }

    # Sanity check: ensure boot/install images are present in SrcIsoRoot after copy
    $bootSrc = if (Test-Path $paths.BootWimInSrc) {
        $paths.BootWimInSrc
    } else { 
        throw "Boot image $paths.BootWimInSrc not found in source after copy"
    }
    $installSrc = if (Test-Path $paths.InstallWimInSrc) {
        $paths.InstallWimInSrc
    } elseif (Test-Path $paths.InstallEsdInSrc) {
        $paths.InstallEsdInSrc
    } else { 
        throw "Install image $paths.InstallWimInSrc or $paths.InstallEsdInSrc not found in source after copy"
    }
    Write-Verbose "Existing boot image in source: $bootSrc"
    Write-Verbose "Existing install image in source: $installSrc"

    # -----------------------------------------------------------------------
    # Step 2: Hardlink-copy SrcIsoRoot -> DestIsoRoot, excluding wim/esd files
    # -----------------------------------------------------------------------
    $destIsoCopyCheckpoint = Join-Path $paths.Checkpoint "destiso.copy.done"

    if (Test-Path $destIsoCopyCheckpoint) {
        Write-Output "Destination ISO copy already completed (checkpoint: destiso.copy.done)"
        Write-Debug  "destiso.copy.done timestamp: $(Get-Content $destIsoCopyCheckpoint -Raw -ErrorAction SilentlyContinue)"
    } else {
        Write-Output "Hardlink-copying $($paths.SrcIsoRoot) -> $($paths.DestIsoRoot) (excluding boot/install images)..."
        Ensure-Folder -Path $paths.DestIsoRoot

        $excludeNames = @($names.BootWim, $names.InstallWim, $names.InstallEsd)
        Write-Verbose "Excluding: $($excludeNames -join ', ')"

        New-HardLinkTree -Source $paths.SrcIsoRoot -Destination $paths.DestIsoRoot -ExcludeFileNames $excludeNames

        Set-Content -Path $destIsoCopyCheckpoint -Value (Get-Date -Format s) -Encoding UTF8
        Write-Output "Destination ISO copy complete (checkpoint: destiso.copy.done)"
    }

    # -----------------------------------------------------------------------
    # Step 3: Export selected indices to per-index uncompressed WIMs
    # -----------------------------------------------------------------------
    # If index selection couldn't be resolved before Export (no ISO accessible at startup),
    # resolve it now from the freshly-copied source WIM.
    if ($SelectedIndices.Count -eq 0) {
        Write-Verbose "SelectedIndices is empty; resolving from SrcISO now..."
        $lateImages = @(Get-WimImageList -WimPath $installSrc)
        $SelectedIndices = @(Resolve-IndexSelection -AllImages $lateImages -SelectHome:$Home -SelectPro:$Pro -IndicesStr $Indices)
        Write-Verbose "Late index resolution: $($SelectedIndices.Count) index/indices selected"
    }

    Write-Output  "Exporting $($SelectedIndices.Count) selected index/indices to $($paths.WimsIndices)..."
    Write-Verbose "Selected indices: $($SelectedIndices.Index -join ', ')"
    Ensure-Folder -Path $paths.WimsRoot
    Ensure-Folder -Path $paths.WimsIndices

    # Cache boot.wim image list (used for per-index boot export)
    $bootImageList = @(Get-WimImageList -WimPath $bootSrc)
    Write-Debug "$names.BootWim has $($bootImageList.Count) image(s)"

    foreach ($img in $SelectedIndices) {
        $idx      = $img.Index
        $imgName  = $img.Name
        Write-Output "  [Index $idx] Exporting: $imgName"

        # -- Export install image --
        $installDest    = Join-Path $paths.WimsIndices ("{0}_{1}" -f $idx, $names.InstallWim)
        $installExtCkpt = "$installDest.extracted"

        if (Test-Path $installExtCkpt) {
            Write-Output "    $names.InstallWim index $idx already extracted (checkpoint exists)"
        } else {
            Write-Output "    Exporting install image index $idx -> $(Split-Path $installDest -Leaf)"
            $dismInstArgs = @(
                '/Export-Image',
                "/SourceImageFile:$installSrc",
                "/SourceIndex:$idx",
                "/DestinationImageFile:$installDest",
                '/Compress:None'
            )
            Write-Verbose "    dism $($dismInstArgs -join ' ')"
            $dismOut = & $dismExe $dismInstArgs 2>&1
            if ($DebugPreference -eq 'Continue') { $dismOut | ForEach-Object { Write-Debug "    DISM> $_" } }
            if ($LASTEXITCODE -ne 0) {
                throw "DISM export failed for install index $idx (exit $LASTEXITCODE)"
            }
            Set-Content -Path $installExtCkpt -Value (Get-Date -Format s) -Encoding UTF8
            Write-Output "    install.wim index $idx extracted (checkpoint: $(Split-Path $installExtCkpt -Leaf))"
        }

        # -- Export boot image (use boot.wim index 2 = Windows Setup PE, fallback to 1) --
        $bootDest    = Join-Path $paths.WimsIndices ("{0}_{1}" -f $idx, $names.BootWim)
        $bootExtCkpt = "$bootDest.extracted"

        if (Test-Path $bootExtCkpt) {
            Write-Output "    $names.BootWim index $idx already extracted (checkpoint exists)"
        } else {
            # Determine which boot.wim index to use (Windows Setup PE = index 2 by convention)
            $bootIdx = 2
            if (-not ($bootImageList | Where-Object { $_.Index -eq $bootIdx })) {
                $bootIdx = 1
            }
            if (-not ($bootImageList | Where-Object { $_.Index -eq $bootIdx })) {
                Write-Warning "    $names.BootWim has no usable index for $idx; skipping boot export"
                continue
            }

            Write-Output "    Exporting $names.BootWim (source index $bootIdx) -> $(Split-Path $bootDest -Leaf)"
            $dismBootArgs = @(
                '/Export-Image',
                "/SourceImageFile:$bootSrc",
                "/SourceIndex:$bootIdx",
                "/DestinationImageFile:$bootDest",
                '/Compress:None'
            )
            Write-Verbose "    dism $($dismInstArgs -join ' ')"
            $dismBootOut = & $dismExe $dismBootArgs 2>&1
            if ($DebugPreference -eq 'Continue') { $dismBootOut | ForEach-Object { Write-Debug "    DISM> $_" } }
            if ($LASTEXITCODE -ne 0) {
                Write-Warning "    DISM boot export failed for index $idx (exit $LASTEXITCODE); skipping"
            } else {
                Set-Content -Path $bootExtCkpt -Value (Get-Date -Format s) -Encoding UTF8
                Write-Output "    $names.BootWim index $idx extracted (checkpoint: $(Split-Path $bootExtCkpt -Leaf))"
            }
        }
    }

    $exportCheckpoint = Join-Path $paths.Checkpoint "export.done"
    Set-Content -Path $exportCheckpoint -Value (Get-Date -Format s) -Encoding UTF8
    Write-Output "Export workflow complete (checkpoint: export.done)"
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

    Write-Output ("Searching for {0}{1}..." -f $Query, ($(if ($FirstOnly) { " (first result only)" } else { "" })))

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

    Write-Output ("Preparing downloads for update {0}: {1}" -f $Update.Guid, $Update.Title)

    $results = @()

    # No URLs → nothing to do
    if (-not $Update.DownloadUrls -or $Update.DownloadUrls.Count -eq 0) {
        Write-Output ("No download URLs for update {0}" -f $Update.Guid)
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
            Write-Output ("File already exists, skipping: {0}" -f $fileName)

            $results += [PSCustomObject]@{
                FileName = $fileName
                FullPath = $destPath
                Url      = $url
            }

            continue
        }

        Write-Output ("Downloading {0}..." -f $fileName)

        # ------------------------------------------------------------
        # Retry loop (3 attempts)
        # ------------------------------------------------------------
        $maxRetries = 3
        $success = $false

        for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {

            Write-Output ("  Attempt {0} of {1}" -f $attempt, $maxRetries)

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
                Write-Output ("    0%  0/{0} bytes" -f $total)

                while (($read = $inStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                    $outStream.Write($buffer, 0, $read)
                    $totalRead += $read

                    if ($total -gt 0) {
                        $pct = [math]::Floor(($totalRead / $total) * 100)

                        if ($pct -ge $nextMark) {
                            # Pipe-safe: always print full lines, never CR
                            Write-Output ("  {0,3}%  {1:N0}/{2:N0} bytes" -f $pct, $totalRead, $total)
                            $nextMark += 10
                        }
                    }
                }

                Write-Output ("  Completed: {0}" -f $fileName)

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

    Write-Output ("Completed downloads for update {0}: {1}" -f $Update.Guid, $Update.Title)

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

    Write-Output  ("Processing update #{0}: {1}" -f $Count, $Guid)
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
        Write-Output ("{0} superseded" -f $Guid)
        return
    }

    # ------------------------------------------------------------
    # DOWNLOAD LINKS (via Get-UpdateLinks)
    # ------------------------------------------------------------

    Write-Output "Finding download links for $title"

    $links = Get-UpdateLinks -Guid $Guid
    $downloadUrls = @()
    if ($links) {
        $downloadUrls = $links.URL | Select-Object -Unique
    }

    Write-Output ("Found {0} file(s) for this update" -f $downloadUrls.Count)

    # Not a keeper if no download links
    if ($downloadUrls.Count -eq 0) {
        Write-Output ("{0} has no download links" -f $Guid)
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
            Write-Output "[DryRun] Would clean $path"
        } elseif (Test-Path $path) {
            Write-Output "Removing: $path"
            Remove-Item $path -Recurse -Force
        }
        return
    }

    if ($DryRun) {
        foreach ($u in $kbDirs) {
            Write-Output ("[DryRun] Would fill: {0}" -f $paths["KBs$u"])
        }
        return
    }

    Write-Output "Starting KB update workflow..."

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
    Write-Output ("Found {0} total updates to process" -f $allGuids.Count)
    Write-Debug  ("Guid TargetFolder:`n" + (@($allGuids | ForEach-Object { '{0} {1}' -f $_.Guid, $_.TargetFolder }) -join "`n"))

    if ($allGuids.Count -eq 0) {
        Write-Output "No updates found"
        return
    }

    Write-Output "Retrieving update details..."

    $count = 0
    $details = foreach ($g in $allGuids) {
        Write-Debug ("Resolving details for {0} ({1})" -f $g.Guid, $g.TargetFolder)
        Get-UpdateDetails -Count (++$count) -Guid $g.Guid -TargetFolder $g.TargetFolder
    }

    if ($details.Count -eq 0) {
        Write-Output "No usable updates after details resolution"
        return
    }

    Write-Output ("Remaining applicable updates: {0}" -f $details.Count)
    Write-Debug   "GUIDs:`n$($details.Guid -join "`n")"

    Write-Output "Synchronizing update folders..."

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

    Write-Output "Downloading required update files..."

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

    Write-Output "Writing manifests..."
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

    Write-Output "KB update workflow complete"
}

# =========================
# Service/Patch section
# =========================

function Invoke-ServiceWork {
    [CmdletBinding()]
    param()

    if ($DryRun) {
        Write-Output "[DryRun] Would service extracted indices in $($paths.WimsIndices)"
        Write-Output "[DryRun] Would apply SSU packages from : $($paths.KBsSSU)"
        Write-Output "[DryRun] Would apply LCU packages from : $($paths.KBsOSCU)"
        Write-Output "[DryRun] Would service winre.wim inside each index's install.wim"
        Write-Output "[DryRun] Would assemble final install.wim -> $($paths.InstallWimInDest)"
        Write-Output "[DryRun] Would assemble final boot.wim   -> $($paths.BootWimInDest)"
        return
    }

    if ($Clean) {
        Write-Debug "Invoke-ServiceWork: Clean mode"
        if (Test-Path $paths.WimsRoot) {
            Write-Output "Removing WIMs root: $($paths.WimsRoot)"
            Remove-Item $paths.WimsRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        # Remove service checkpoints (per-index and final assembly)
        if (Test-Path $paths.Checkpoint) {
            Get-ChildItem $paths.Checkpoint -Filter "*.done" -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '^\d+_|^install\.wim\.done$|^boot\.wim\.done$' } |
                ForEach-Object {
                    Write-Verbose "Removing service checkpoint: $($_.Name)"
                    Remove-Item $_.FullName -Force
                }
            Get-ChildItem $paths.Checkpoint -Filter "*.extracted" -ErrorAction SilentlyContinue |
                ForEach-Object {
                    Write-Verbose "Removing extract checkpoint: $($_.Name)"
                    Remove-Item $_.FullName -Force
                }
        }
        return
    }

    Write-Output  "Starting Service workflow..."
    Write-Verbose "Invoke-ServiceWork: WimsIndices='$($paths.WimsIndices)'"
    Write-Debug   "Invoke-ServiceWork: KBsSSU='$($paths.KBsSSU)' KBsOSCU='$($paths.KBsOSCU)' WimsMounts='$($paths.WimsMounts)'"

    $CompressionType = 'Maximum'  # None, Fast, Maximum

    Ensure-Folder -Path $paths.WimsMounts
    Ensure-Folder -Path $paths.WimsServiced
    Ensure-Folder -Path $paths.Checkpoint

    if (-not (Test-Path $paths.WimsIndices)) {
        Write-Output "No indices folder found: $($paths.WimsIndices). Run with -Export first."
        return
    }

    # Discover extracted install.wim files and derive index numbers
    $indexFiles = @(Get-ChildItem -Path $paths.WimsIndices -Filter "*_$($names.InstallWim)" -File -ErrorAction SilentlyContinue)
    if ($indexFiles.Count -eq 0) {
        Write-Output "No extracted $($names.InstallWim) files found in $($paths.WimsIndices). Run with -Export first."
        return
    }

    $extractedIndices = $indexFiles |
        ForEach-Object { [int](($_.BaseName -split '_')[0]) } |
        Sort-Object
    Write-Output "Found $($extractedIndices.Count) extracted indices: $($extractedIndices -join ', ')"

    # Gather available packages (.msu and .cab)
    $ssuFiles = @(
        Get-ChildItem -Path $paths.KBsSSU -Include '*.msu', '*.cab' -Recurse -ErrorAction SilentlyContinue
    )
    $lcuFiles = @(
        Get-ChildItem -Path $paths.KBsOSCU -Include '*.msu', '*.cab' -Recurse -ErrorAction SilentlyContinue
    )
    $hasSSU = $ssuFiles.Count -gt 0
    $hasLCU = $lcuFiles.Count -gt 0

    Write-Output  "Packages available - SSU: $hasSSU ($($ssuFiles.Count) files), LCU: $hasLCU ($($lcuFiles.Count) files)"
    Write-Verbose "SSU : $($ssuFiles.Name -join ', ')"
    Write-Verbose "LCU : $($lcuFiles.Name -join ', ')"

    # -----------------------------------------------------------------------
    # Per-index servicing
    # -----------------------------------------------------------------------
    foreach ($idx in $extractedIndices) {
        Write-Output "--- Index $idx ---"
        Write-Verbose "Processing index $idx"

        $installWimPath = Join-Path $paths.WimsIndices ("{0}_{1}" -f $idx, $names.InstallWim)
        $bootWimPath    = Join-Path $paths.WimsIndices ("{0}_{1}" -f $idx, $names.BootWim)

        $installDoneChkpt = Join-Path $paths.Checkpoint ("{0}_{1}.done"      -f $idx, $names.InstallWim)
        $bootDoneChkpt    = Join-Path $paths.Checkpoint ("{0}_{1}.done"      -f $idx, $names.BootWim)
        $winreDoneChkpt   = Join-Path $paths.Checkpoint ("{0}_{1}.done"      -f $idx, $names.WinreWim)
        $winreExtChkpt    = Join-Path $paths.Checkpoint ("{0}_{1}.extracted" -f $idx, $names.WinreWim)

        $mountDir      = Join-Path $paths.WimsMounts ("mount_{0}"      -f $idx)
        $winreMountDir = Join-Path $paths.WimsMounts ("winremount_{0}" -f $idx)
        $winreWimPath  = Join-Path $paths.WimsIndices ("{0}_{1}"       -f $idx, $names.WinreWim)

        # ---- Service install.wim ----
        if (Test-Path $installDoneChkpt) {
            Write-Output "  $($names.InstallWim) index $idx already serviced (checkpoint exists)"
        } else {
            if ($hasSSU -or $hasLCU) {
                Write-Output "  Servicing $($names.InstallWim) for index $idx..."
                Ensure-Folder -Path $mountDir

                try {
                    Write-Output  "  Mounting $installWimPath -> $mountDir"
                    Write-Verbose "  dism /Mount-Image /Index:1"
                    $outMnt = & $dismExe /Mount-Image "/ImageFile:$installWimPath" /Index:1 "/MountDir:$mountDir" 2>&1
                    if ($DebugPreference -eq 'Continue') { $outMnt | ForEach-Object { Write-Debug "  DISM> $_" } }
                    if ($LASTEXITCODE -ne 0) { throw "DISM mount $($names.InstallWim) failed for index $idx (exit $LASTEXITCODE)" }

                    # ---- Service winre.wim inside install.wim ----
                    if ($hasSSU) {
                        $winreInMount = Join-Path $mountDir $paths.WinreWimInWim
                        Write-Debug "  Checking for $($names.WinreWim) at: $winreInMount"

                        if (Test-Path $winreInMount) {
                            if (Test-Path $winreDoneChkpt) {
                                Write-Output "  $($names.WinreWim) index $idx already serviced (checkpoint exists)"
                            } else {
                                # Extract winre.wim
                                if (-not (Test-Path $winreExtChkpt)) {
                                    Write-Output "  Extracting $($names.WinreWim) from mounted install image..."
                                    Copy-Item -Path $winreInMount -Destination $winreWimPath -Force
                                    Set-Content -Path $winreExtChkpt -Value (Get-Date -Format s) -Encoding UTF8
                                    Write-Output "  $($names.WinreWim) extracted (checkpoint: $($idx)_$($names.WinreWim).extracted)"
                                } else {
                                    Write-Output "  $($names.WinreWim) already extracted (checkpoint exists)"
                                }

                                # Mount winre.wim
                                Ensure-Folder -Path $winreMountDir
                                Write-Output  "  Mounting $($names.WinreWim) -> $winreMountDir"
                                Write-Verbose "  dism /Mount-Image winre"
                                $outWrMnt = & $dismExe /Mount-Image "/ImageFile:$winreWimPath" /Index:1 "/MountDir:$winreMountDir" 2>&1
                                if ($DebugPreference -eq 'Continue') { $outWrMnt | ForEach-Object { Write-Debug "  DISM> $_" } }
                                if ($LASTEXITCODE -ne 0) { throw "DISM mount $($names.WinreWim) failed for index $idx (exit $LASTEXITCODE)" }

                                # Apply SSU packages to winre
                                foreach ($pkg in $ssuFiles) {
                                    Write-Output  "  Applying SSU to $($names.WinreWim): $($pkg.Name)"
                                    Write-Verbose "  dism /Add-Package winre <- $($pkg.Name)"
                                    $outWrPkg = & $dismExe /Add-Package "/Image:$winreMountDir" "/PackagePath:$($pkg.FullName)" 2>&1
                                    if ($DebugPreference -eq 'Continue') { $outWrPkg | ForEach-Object { Write-Debug "  DISM> $_" } }
                                    if ($LASTEXITCODE -ne 0) {
                                        Write-Warning "  DISM SSU->$($names.WinreWim) failed: $($pkg.Name) index $idx (exit $LASTEXITCODE); continuing"
                                    }
                                }

                                # Unmount and commit winre
                                Write-Output "  Unmounting $($names.WinreWim) (commit)..."
                                $outWrUnm = & $dismExe /Unmount-Image "/MountDir:$winreMountDir" /Commit 2>&1
                                if ($DebugPreference -eq 'Continue') { $outWrUnm | ForEach-Object { Write-Debug "  DISM> $_" } }
                                if ($LASTEXITCODE -ne 0) { throw "DISM unmount $($names.WinreWim) failed for index $idx (exit $LASTEXITCODE)" }

                                # Reinsert serviced winre.wim back into mounted install.wim
                                Write-Output "  Reinserting serviced $($names.WinreWim) into install image..."
                                Copy-Item -Path $winreWimPath -Destination $winreInMount -Force
                                Set-Content -Path $winreDoneChkpt -Value (Get-Date -Format s) -Encoding UTF8
                                Write-Output "  $($names.WinreWim) serviced and reinserted (checkpoint: $($idx)_$($names.WinreWim).done)"
                            }
                        } else {
                            Write-Verbose "  $($names.WinreWim) not found in mounted install image at $winreInMount; skipping winre servicing"
                        }

                        # Apply SSU packages to install.wim
                        foreach ($pkg in $ssuFiles) {
                            Write-Output  "  Applying SSU to $($names.InstallWim): $($pkg.Name)"
                            Write-Verbose "  dism /Add-Package install <- $($pkg.Name)"
                            $outInstPkg = & $dismExe /Add-Package "/Image:$mountDir" "/PackagePath:$($pkg.FullName)" 2>&1
                            if ($DebugPreference -eq 'Continue') { $outInstPkg | ForEach-Object { Write-Debug "  DISM> $_" } }
                            if ($LASTEXITCODE -ne 0) {
                                Write-Warning "  DISM SSU->$($names.InstallWim) failed: $($pkg.Name) index $idx (exit $LASTEXITCODE); continuing"
                            }
                        }
                    }

                    # Apply LCU (OSCU) packages to install.wim
                    if ($hasLCU) {
                        foreach ($pkg in $lcuFiles) {
                            Write-Output  "  Applying LCU to $($names.InstallWim): $($pkg.Name)"
                            Write-Verbose "  dism /Add-Package install (LCU) <- $($pkg.Name)"
                            $outLcuPkg = & $dismExe /Add-Package "/Image:$mountDir" "/PackagePath:$($pkg.FullName)" 2>&1
                            if ($DebugPreference -eq 'Continue') { $outLcuPkg | ForEach-Object { Write-Debug "  DISM> $_" } }
                            if ($LASTEXITCODE -ne 0) {
                                Write-Warning "  DISM LCU->$($names.InstallWim) failed: $($pkg.Name) index $idx (exit $LASTEXITCODE); continuing"
                            }
                        }
                    }

                    # Unmount and commit install.wim
                    Write-Output "  Unmounting $($names.InstallWim) index $idx (commit)..."
                    $outInstUnm = & $dismExe /Unmount-Image "/MountDir:$mountDir" /Commit 2>&1
                    if ($DebugPreference -eq 'Continue') { $outInstUnm | ForEach-Object { Write-Debug "  DISM> $_" } }
                    if ($LASTEXITCODE -ne 0) { throw "DISM unmount $($names.InstallWim) failed for index $idx (exit $LASTEXITCODE)" }

                    Set-Content -Path $installDoneChkpt -Value (Get-Date -Format s) -Encoding UTF8
                    Write-Output "  $($names.InstallWim) index $idx serviced (checkpoint: $($idx)_$($names.InstallWim).done)"

                } catch {
                    Write-Output "  ERROR servicing $($names.InstallWim) index $idx`: $_" 'ERROR'
                    if (Test-Path $mountDir) {
                        Write-Output "  Discarding mounted $($names.InstallWim)..."
                        & $dismExe /Unmount-Image "/MountDir:$mountDir" /Discard 2>&1 | Out-Null
                    }
                    if (Test-Path $winreMountDir) {
                        Write-Output "  Discarding mounted $($names.WinreWim)..."
                        & $dismExe /Unmount-Image "/MountDir:$winreMountDir" /Discard 2>&1 | Out-Null
                    }
                    throw
                } finally {
                    if (Test-Path $mountDir)      { Remove-Item $mountDir      -Recurse -Force -ErrorAction SilentlyContinue }
                    if (Test-Path $winreMountDir) { Remove-Item $winreMountDir -Recurse -Force -ErrorAction SilentlyContinue }
                }
            } else {
                Write-Output "  No SSU or LCU packages present; skipping $($names.InstallWim) servicing for index $idx"
                Set-Content -Path $installDoneChkpt -Value (Get-Date -Format s) -Encoding UTF8
            }
        }

        # ---- Service boot.wim ----
        if (-not (Test-Path $bootWimPath)) {
            Write-Verbose "  No $($names.BootWim) found for index $idx; skipping boot servicing"
        } elseif (Test-Path $bootDoneChkpt) {
            Write-Output "  $($names.BootWim) index $idx already serviced (checkpoint exists)"
        } elseif (-not $hasSSU) {
            Write-Output "  No SSU packages present; skipping $($names.BootWim) servicing for index $idx"
            Set-Content -Path $bootDoneChkpt -Value (Get-Date -Format s) -Encoding UTF8
        } else {
            Write-Output "  Servicing $($names.BootWim) for index $idx..."
            Ensure-Folder -Path $mountDir

            try {
                Write-Output  "  Mounting $bootWimPath -> $mountDir"
                Write-Verbose "  dism /Mount-Image boot /Index:1"
                $outBootMnt = & $dismExe /Mount-Image "/ImageFile:$bootWimPath" /Index:1 "/MountDir:$mountDir" 2>&1
                if ($DebugPreference -eq 'Continue') { $outBootMnt | ForEach-Object { Write-Debug "  DISM> $_" } }
                if ($LASTEXITCODE -ne 0) { throw "DISM mount $($names.BootWim) failed for index $idx (exit $LASTEXITCODE)" }

                foreach ($pkg in $ssuFiles) {
                    Write-Output  "  Applying SSU to $($names.BootWim): $($pkg.Name)"
                    Write-Verbose "  dism /Add-Package boot <- $($pkg.Name)"
                    $outBootPkg = & $dismExe /Add-Package "/Image:$mountDir" "/PackagePath:$($pkg.FullName)" 2>&1
                    if ($DebugPreference -eq 'Continue') { $outBootPkg | ForEach-Object { Write-Debug "  DISM> $_" } }
                    if ($LASTEXITCODE -ne 0) {
                        Write-Warning "  DISM SSU->$($names.BootWim) failed: $($pkg.Name) index $idx (exit $LASTEXITCODE); continuing"
                    }
                }

                Write-Output "  Unmounting $($names.BootWim) index $idx (commit)..."
                $outBootUnm = & $dismExe /Unmount-Image "/MountDir:$mountDir" /Commit 2>&1
                if ($DebugPreference -eq 'Continue') { $outBootUnm | ForEach-Object { Write-Debug "  DISM> $_" } }
                if ($LASTEXITCODE -ne 0) { throw "DISM unmount $($names.BootWim) failed for index $idx (exit $LASTEXITCODE)" }

                Set-Content -Path $bootDoneChkpt -Value (Get-Date -Format s) -Encoding UTF8
                Write-Output "  $($names.BootWim) index $idx serviced (checkpoint: $($idx)_$($names.BootWim).done)"

            } catch {
                Write-Output "  ERROR servicing $($names.BootWim) index $idx`: $_" 'ERROR'
                if (Test-Path $mountDir) {
                    & $dismExe /Unmount-Image "/MountDir:$mountDir" /Discard 2>&1 | Out-Null
                }
                throw
            } finally {
                if (Test-Path $mountDir) { Remove-Item $mountDir -Recurse -Force -ErrorAction SilentlyContinue }
            }
        }
    }

    # -----------------------------------------------------------------------
    # Final assembly (serial — compression can be slow)
    # -----------------------------------------------------------------------
    Write-Output "Final assembly: combining serviced indices (compression: $CompressionType)..."

    $compressionMap  = @{ 'None' = 'none'; 'Fast' = 'fast'; 'Maximum' = 'maximum' }
    $dismCompression = $compressionMap[$CompressionType]

    # -- Assemble final install.wim --
    $installWimDoneChkpt = Join-Path $paths.Checkpoint ("{0}.done" -f $names.InstallWim)
    if (Test-Path $installWimDoneChkpt) {
        Write-Output "Final $($names.InstallWim) already assembled (checkpoint exists)"
    } else {
        Ensure-Folder -Path (Split-Path $paths.InstallWimInDest -Parent)

        $sortedInstallWims = @(
            $extractedIndices |
                ForEach-Object { Join-Path $paths.WimsIndices ("{0}_{1}" -f $_, $names.InstallWim) } |
                Where-Object   { Test-Path $_ }
        )

        if ($sortedInstallWims.Count -eq 0) {
            Write-Warning "No $($names.InstallWim) files found to assemble"
        } else {
            $firstWim = $sortedInstallWims[0]
            Write-Output  "Creating final $($names.InstallWim) from: $(Split-Path $firstWim -Leaf)"
            Write-Verbose "dism /Export-Image /Compress:$dismCompression -> $($paths.InstallWimInDest)"

            $outFirst = & $dismExe /Export-Image "/SourceImageFile:$firstWim" /SourceIndex:1 `
                         "/DestinationImageFile:$($paths.InstallWimInDest)" "/Compress:$dismCompression" 2>&1
            if ($DebugPreference -eq 'Continue') { $outFirst | ForEach-Object { Write-Debug "DISM> $_" } }
            if ($LASTEXITCODE -ne 0) { throw "DISM export (first $($names.InstallWim)) failed (exit $LASTEXITCODE)" }

            for ($i = 1; $i -lt $sortedInstallWims.Count; $i++) {
                $srcWim = $sortedInstallWims[$i]
                $idxNum = $extractedIndices[$i]
                Write-Output  "  Appending $($names.InstallWim) index $idxNum..."
                Write-Verbose "  dism /Export-Image append index $idxNum"

                $outApp = & $dismExe /Export-Image "/SourceImageFile:$srcWim" /SourceIndex:1 `
                           "/DestinationImageFile:$($paths.InstallWimInDest)" 2>&1
                if ($DebugPreference -eq 'Continue') { $outApp | ForEach-Object { Write-Debug "DISM> $_" } }
                if ($LASTEXITCODE -ne 0) { throw "DISM append $($names.InstallWim) index $idxNum failed (exit $LASTEXITCODE)" }
            }

            Set-Content -Path $installWimDoneChkpt -Value (Get-Date -Format s) -Encoding UTF8
            Write-Output "Final $($names.InstallWim) assembled (checkpoint: $($names.InstallWim).done)"
        }
    }

    # -- Assemble final boot.wim --
    $bootWimDoneChkpt = Join-Path $paths.Checkpoint ("{0}.done" -f $names.BootWim)
    if (Test-Path $bootWimDoneChkpt) {
        Write-Output "Final $($names.BootWim) already assembled (checkpoint exists)"
    } else {
        Ensure-Folder -Path (Split-Path $paths.BootWimInDest -Parent)

        $sortedBootWims = @(
            $extractedIndices |
                ForEach-Object { Join-Path $paths.WimsIndices ("{0}_{1}" -f $_, $names.BootWim) } |
                Where-Object   { Test-Path $_ }
        )

        if ($sortedBootWims.Count -eq 0) {
            Write-Warning "No $($names.BootWim) files found to assemble"
        } else {
            $firstBoot = $sortedBootWims[0]
            Write-Output  "Creating final $($names.BootWim) from: $(Split-Path $firstBoot -Leaf)"
            Write-Verbose "dism /Export-Image /Compress:$dismCompression -> $($paths.BootWimInDest)"

            $outFirstB = & $dismExe /Export-Image "/SourceImageFile:$firstBoot" /SourceIndex:1 `
                          "/DestinationImageFile:$($paths.BootWimInDest)" "/Compress:$dismCompression" 2>&1
            if ($DebugPreference -eq 'Continue') { $outFirstB | ForEach-Object { Write-Debug "DISM> $_" } }
            if ($LASTEXITCODE -ne 0) { throw "DISM export (first $($names.BootWim)) failed (exit $LASTEXITCODE)" }

            for ($i = 1; $i -lt $sortedBootWims.Count; $i++) {
                $srcBoot = $sortedBootWims[$i]
                $idxNum  = $extractedIndices[$i]
                Write-Output "  Appending $($names.BootWim) index $idxNum..."

                $outAppB = & $dismExe /Export-Image "/SourceImageFile:$srcBoot" /SourceIndex:1 `
                            "/DestinationImageFile:$($paths.BootWimInDest)" 2>&1
                if ($DebugPreference -eq 'Continue') { $outAppB | ForEach-Object { Write-Debug "DISM> $_" } }
                if ($LASTEXITCODE -ne 0) { throw "DISM append $($names.BootWim) index $idxNum failed (exit $LASTEXITCODE)" }
            }

            Set-Content -Path $bootWimDoneChkpt -Value (Get-Date -Format s) -Encoding UTF8
            Write-Output "Final $($names.BootWim) assembled (checkpoint: $($names.BootWim).done)"
        }
    }

    Write-Output "Service workflow complete"
}

# ==============================
# Driver export
# ==============================
function Invoke-DriverWork {
    $WinpeDriverRoot = $paths.WinpeDriverRoot

    if ($Clean) {
        if ($DryRun) {
            Write-Output "[DryRun] Would remove: $WinpeDriverRoot"
        } elseif (Test-Path $WinpeDriverRoot) {
            Write-Output "Removing: $WinpeDriverRoot"
            Remove-Item $WinpeDriverRoot -Recurse -Force
        }
        return
    }

    if ($DryRun) {
        Write-Output "[DryRun] Would run: DISM /export-driver"
        return
    }

    Write-Output  "Exporting drivers..."
    Write-Verbose "Invoke-DriverWork: WinpeDriverRoot='$WinpeDriverRoot'"
    Ensure-Folder -Path $WinpeDriverRoot
    $driverArgs = "/online /export-driver /destination:`"$WinpeDriverRoot`""
    Write-Debug "dism $driverArgs"
    $p = Start-Process -FilePath $dismExe -ArgumentList $driverArgs -NoNewWindow -PassThru -Wait -RedirectStandardOutput "$WinpeDriverRoot\dism.log"
    if ($p.ExitCode -ne 0) { throw "DISM /export-driver failed (exit $($p.ExitCode))" }
}

# ==============================
# Registry export
# ==============================
function Invoke-RegWork {

    <#
    #
    # Example tables (embedded like original template)
    #
    $RegistryAddModify = @(
        @{
            Key    = 'HKLM\SOFTWARE\MyCompany'
            Values = @(
                @('SettingA'),
                @('SettingB','X')
            )
        },
        @{
            Key    = 'HKCU\Software\MyCompany'
            Values = @(
                @()                     # export entire key (dominates)
            )
        }
    )

    $RegistryRemove = @(
        @{
            Key    = 'HKLM\SOFTWARE\MyCompany'
            Values = @(
                @('OldValue')
            )
        },
        @{
            Key    = 'HKCU\Software\MyCompany'
            Values = @(
                @()                     # delete entire key (dominates)
            )
        }
    )
    #>

    $RegistryAddModify = @(
        @{
            Key    = 'HKEY_CLASSES_ROOT\AllFilesystemObjects\shell\Windows.ShowFileExtensions'
            Values = @()
        },
        @{
            Key    = 'HKEY_CLASSES_ROOT\Directory\Background\shell\Windows.ShowFileExtensions'
            Values = @()
        },
        @{
            Key    = 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
            Values = @('LaunchTo',
                       'Start_IrisRecommendations',
                       'ShowTaskViewButton',
                       'HideFileExt',
                       'SeparateProcess')
        },
        @{
            Key    = 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState'
            Values = @('FullPath')
        },
        @{
            Key    = 'HKEY_LOCAL_MACHINE\Software\Microsoft\WindowsUpdate\UX\Settings'
            Values = @('AllowMUUpdateService')
        },
        @{
            Key    = 'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\FileSystem'
            Values = @('LongPathsEnabled')
        },
        @{
            Key    = 'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power'
            Values = @('HibernateEnabled')
        },
        @{
            Key    = 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings'
            Values = @('TaskbarEndTask')
        }
    )

    $RegistryRemove = @(
    )

    $RegistryRoot = $paths.RegistryRoot

    #
    # CLEAN MODE
    #
    if ($Clean) {
        if ($DryRun) {
            Write-Output "[DryRun] Would remove: $RegistryRoot"
        } elseif (Test-Path $RegistryRoot) {
            Write-Output "Removing: $RegistryRoot"
            Remove-Item $RegistryRoot -Recurse -Force
        }
        return
    }

    #
    # PREP OUTPUT FOLDER
    #
    if (-not $DryRun) {
        Ensure-Folder -Path $RegistryRoot
    }

    #
    # HELPERS
    #
    function RegSafeName([string]$key) {
        ($key -replace '[\\/:*?"<>|]', '_') + '.reg'
    }

    function RegExportEntireKey([string]$key, [string]$dest) {
        if ($DryRun) {
            Write-Output "[DryRun] Would export ENTIRE key: $key -> $dest"
        } else {
            Write-Output "Export ENTIRE key: $key -> $dest"
            reg.exe export "$key" "$dest" /y | Out-Null
        }
    }

    function RegExportSpecificValues([string]$key, [string]$dest, [object[]]$groups) {

        # Flatten values, skipping null/empty
        $allValues = @()
        foreach ($g in $groups) {
            if (-not $g) { continue }
            foreach ($v in $g) {
                if ($null -ne $v -and $v -ne '') { $allValues += $v }
            }
        }
        $allValues = $allValues | Sort-Object -Unique

        if (-not $allValues -or $allValues.Count -eq 0) {
            Write-Output "No specific values requested for $key; skipping."
            return
        }

        if ($DryRun) {
            Write-Output "[DryRun] Would export values [$($allValues -join ', ')] from $key -> $dest"
            return
        }

        Write-Output "Export specific values [$($allValues -join ', ')] from $key -> $dest"

        $query = reg.exe query "$key" /v * 2>$null
        if (-not $query) {
            Write-Output "WARNING: No data returned for $key"
            return
        }

        $out = New-Object System.Collections.Generic.List[string]
        $out.Add("Windows Registry Editor Version 5.00")
        $out.Add("")

        $header = "[" + $key + "]"
        $out.Add($header)

        foreach ($line in $query) {

            if ($line -match '^\s+([^\s]+)\s+REG_([A-Z0-9_]+)\s+(.*)$') {

                $valName = $matches[1]
                $type    = $matches[2]
                $data    = $matches[3]

                if ($allValues -contains $valName) {

                    switch ($type) {
                        "SZ" {
                            $regLine = '"' + $valName + '"="' + $data + '"'
                        }
                        "DWORD" {
                            $regLine = '"' + $valName + '"=dword:' + ("{0:x8}" -f [int]$data)
                        }
                        default {
                            $regLine = '"' + $valName + '"="' + $data + '"'
                        }
                    }

                    $out.Add($regLine)
                }
            }
        }

        $out.Add("")
        $out -join "`r`n" | Set-Content -Path $dest -Encoding Unicode
    }

    function RegAppendDelete([string]$dest, [string]$key, [string[]]$values) {

        if ($DryRun) {
            if (-not $values -or $values.Count -eq 0) {
                Write-Output "[DryRun] Would delete ENTIRE key: $key"
            } else {
                Write-Output "[DryRun] Would delete values [$($values -join ', ')] from $key"
            }
            return
        }

        Write-Output "Appending delete instructions for $key -> $dest"

        $out = New-Object System.Collections.Generic.List[string]

        if (Test-Path $dest) {
            $existing = Get-Content $dest -Raw
            foreach ($l in ($existing -split "`r?`n")) { $out.Add($l) }
        } else {
            $out.Add("Windows Registry Editor Version 5.00")
            $out.Add("")
        }

        if (-not $values -or $values.Count -eq 0) {
            $line = "[-" + $key + "]"
            $out.Add($line)
            $out.Add("")
        } else {
            $header = "[" + $key + "]"
            $out.Add($header)

            foreach ($v in $values) {
                $line = '"' + $v + '"=-'
                $out.Add($line)
            }

            $out.Add("")
        }

        $out -join "`r`n" | Set-Content -Path $dest -Encoding Unicode
    }

    #
    # PROCESS ADD/MODIFY
    #
    foreach ($entry in $RegistryAddModify) {

        $key    = $entry.Key
        $groups = $entry.Values

        $safe = "AddModify_" + (RegSafeName $key)
        $dest = Join-Path $RegistryRoot $safe

        # entire key if:
        # - Values is null/empty, OR
        # - any inner list is null/empty
        $hasEntire = $false
        if (-not $groups -or $groups.Count -eq 0) {
            $hasEntire = $true
        } else {
            foreach ($g in $groups) {
                if (-not $g -or ($g -is [System.Array] -and $g.Count -eq 0)) {
                    $hasEntire = $true
                    break
                }
            }
        }

        if ($hasEntire) {
            if ($DryRun) {
                Write-Output "[DryRun] Would export ENTIRE key: $key -> $dest"
            } else {
                RegExportEntireKey $key $dest
            }
        } else {
            RegExportSpecificValues $key $dest $groups
        }
    }

    #
    # PROCESS REMOVE
    #
    foreach ($entry in $RegistryRemove) {

        $key    = $entry.Key
        $groups = $entry.Values

        $safe = "Remove_" + (RegSafeName $key)
        $dest = Join-Path $RegistryRoot $safe

        $hasEntire = $false
        if (-not $groups -or $groups.Count -eq 0) {
            $hasEntire = $true
        } else {
            foreach ($g in $groups) {
                if (-not $g -or ($g -is [System.Array] -and $g.Count -eq 0)) {
                    $hasEntire = $true
                    break
                }
            }
        }

        if ($hasEntire) {
            RegAppendDelete $dest $key @()
        } else {
            $allValues = @()
            foreach ($g in $groups) {
                if (-not $g) { continue }
                foreach ($v in $g) {
                    if ($null -ne $v -and $v -ne '') { $allValues += $v }
                }
            }
            $allValues = $allValues | Sort-Object -Unique

            RegAppendDelete $dest $key $allValues
        }
    }
}

# ==============================
# InstallDrivers.cmd
# ==============================
function Write-InstallDriversCmd {

    $path = $paths.InstallDriversCmd

    $template = @'
@echo off
setlocal
set "SRC=%~dp0"
:: Must be run elevated to work

echo Import drivers
pnputil /add-driver "%SRC%\{0}\*.inf" /subdirs /install
endlocal
'@

    if ($Clean) {
        if ($DryRun) {
            Write-Output "[DryRun] Would remove: $path"
        } elseif (Test-Path $path) {
            Write-Output "Removing: $path"
            Remove-Item $path -Force
        }
        return
    }

    if ($DryRun) {
        Write-Output "[DryRun] Would write: $path"
    } else {
        Write-Output "Writing: $path"
        Ensure-Folder (Split-Path $path -Parent)

        $content = $template -f $names.WinpeDriver
        Set-Content -LiteralPath $path -Value $content -Encoding ASCII
    }
}

# ==============================
# InstallRegs.cmd
# ==============================
function Write-InstallRegsCmd {

    $path = $paths.InstallRegsCmd

    $template = @'
@echo off
setlocal
set "SRC=%~dp0"

echo Import registry files
for %%F in ("%SRC%\{0}\*.reg") do (
    reg.exe import "%%F"
)
endlocal
'@

    if ($Clean) {
        if ($DryRun) {
            Write-Output "[DryRun] Would remove: $path"
        } elseif (Test-Path $path) {
            Write-Output "Removing: $path"
            Remove-Item $path -Force
        }
        return
    }

    if ($DryRun) {
        Write-Output "[DryRun] Would write: $path"
    } else {
        Write-Output "Writing: $path"
        Ensure-Folder (Split-Path $path -Parent)

        $osContent = ""
        foreach ($u in $kbDirs) {
            $osContent += $osTemplate -f $names.KBs, $names.$u
        }
        $content = $template -f $names.Registry
        Set-Content -LiteralPath $path -Value $content -Encoding ASCII
    }
}

# ==============================
# PostSetup.cmd
# ==============================
function Write-PostSetupCmd {

    $path = $paths.PostSetupCmd

    $template = @'
@echo off
setlocal enabledelayedexpansion
set "SRC=%~dp0"

:: Apply KBs in the correct order

{0}
endlocal
'@

    $osTemplate = @'
echo Installing updates from %SRC%\{0}\{1}

:: Install EXE installers
for %%F in ("%SRC%\{0}\{1}\*.exe") do (
    echo Installing EXE %%F
    "%%F" /quiet /norestart
)

:: Install MSI installers
for %%F in ("%SRC%\{0}\{1}\*.msi") do (
    echo Installing MSI %%F
    msiexec.exe /i "%%F" /quiet /norestart
)

:: Run CMD/BAT scripts
for %%F in ("%SRC%\{0}\{1}\*.cmd") do (
    echo Running CMD %%F
    call "%%F"
)
for %%F in ("%SRC%\{0}\{1}\*.bat") do (
    echo Running BAT %%F
    call "%%F"
)

'@

    if ($Clean) {
        if ($DryRun) {
            Write-Output "[DryRun] Would remove: $path"
        } elseif (Test-Path $path) {
            Write-Output "Removing: $path"
            Remove-Item $path -Force
        }
        return
    }

    if ($DryRun) {
        Write-Output "[DryRun] Would write: $path"
    } else {
        Write-Output "Writing: $path"
        Ensure-Folder (Split-Path $path -Parent)

        $osContent = ""
        foreach ($u in $kbDirs) {
            $osContent += $osTemplate -f $names.KBs, $names.$u
        }
        $content = $template -f $osContent, $names.Registry
        Set-Content -LiteralPath $path -Value $content -Encoding ASCII
    }
}

# ==============================
# SetupConfig files
# ==============================
function Write-SetupConfigFiles {
    $cleanPath   = $paths.SetupConfigCleanIni
    $upgradePath = $paths.SetupConfigUpgradeIni

<#
    $cleanTemplate = @'
# Clean installation configuration
[SetupConfig]

# Perform a clean installation
Auto=Clean

# Disable Dynamic Update (no online updates or drivers)
DynamicUpdate=Disable

# Prevent Setup from injecting drivers automatically
InstallDrivers=Off

# Show the full Out-of-Box Experience (OOBE)
ShowOOBE=Full

# Disable Setup telemetry
Telemetry=Disable

'@

    $upgradeTemplate = @'
# Upgrade installation configuration
[SetupConfig]

# Perform an in-place upgrade
Auto=Upgrade

# Disable Dynamic Update (no online updates or drivers)
DynamicUpdate=Disable

# Prevent Setup from injecting drivers automatically
InstallDrivers=Off

# Do not show the Out-of-Box Experience (OOBE)
ShowOOBE=None

# Disable Setup telemetry
Telemetry=Disable

'@
#>
    $cleanTemplate = @'
[SetupConfig]
Auto=Clean
DynamicUpdate=Disable
Telemetry=Disable
'@

    $upgradeTemplate = @'
[SetupConfig]
Auto=Upgrade
DynamicUpdate=Disable
Telemetry=Disable
'@


    if ($Clean) {
        if ($DryRun) {
            Write-Output "[DryRun] Would remove: $cleanPath"
            Write-Output "[DryRun] Would remove: $upgradePath"
        } else {
            if (Test-Path $cleanPath) {
                Write-Output "Removing: $cleanPath"
                Remove-Item $cleanPath -Force
            }
            if (Test-Path $upgradePath) {
                Write-Output "Removing: $upgradePath"
                Remove-Item $upgradePath -Force
            }
        }
        return
    }

    if ($DryRun) {
        Write-Output "[DryRun] Would write: $cleanPath"
        Write-Output "[DryRun] Would write: $upgradePath"
    } else {
        Write-Output "Writing: $cleanPath"
        Set-Content -LiteralPath $cleanPath   -Value $cleanTemplate   -Encoding ASCII
        Write-Output "Writing: $upgradePath"
        Set-Content -LiteralPath $upgradePath -Value $upgradeTemplate -Encoding ASCII
    }
}

# ==============================
# Setup CMD Files
# ==============================
function Write-SetupCmdFiles {
    $cleanPath   = $paths.CleanInstallCmd
    $upgradePath = $paths.UpgradeCmd

    $cleanTemplate = @'
@echo off
setlocal
set "SRC=%~dp0"
echo WARNING: This will start a CLEAN install (wipe-and-load) when run from within Windows.
echo Close all apps and ensure you have backups.
echo.
"%SRC%setup.exe" /auto clean /eula accept /configfile "%SRC%{0}"
endlocal
'@

    $upgradeTemplate = @'
@echo off
setlocal
set "SRC=%~dp0"
echo Running in-place upgrade from: %SRC%
"%SRC%setup.exe" /auto upgrade /eula accept /configfile "%SRC%{0}"
endlocal
'@

    if ($Clean) {
        if ($DryRun) {
            Write-Output "[DryRun] Would remove: $cleanPath"
            Write-Output "[DryRun] Would remove: $upgradePath"
        } else {
            if (Test-Path $cleanPath) {
                Write-Output "Removing: $cleanPath"
                Remove-Item $cleanPath -Force
            }
            if (Test-Path $upgradePath) {
                Write-Output "Removing: $upgradePath"
                Remove-Item $upgradePath -Force
            }
        }
        return
    }

    if ($DryRun) {
        Write-Output "[DryRun] Would write: $cleanPath"
        Write-Output "[DryRun] Would write: $upgradePath"
    } else {
        # Fill in the actual name for the files in the template
        $cleanContent   = $cleanTemplate -f $names.SetupConfigCleanIni
        $upgradeContent = $upgradeTemplate -f $names.SetupConfigUpgradeIni

        Write-Output "Writing: $cleanPath"
        Set-Content -LiteralPath $cleanPath   -Value $cleanContent   -Encoding ASCII
        Write-Output "Writing: $upgradePath"
        Set-Content -LiteralPath $upgradePath -Value $upgradeContent -Encoding ASCII
    }
}

# ==============================
# Create the Destination ISO
# ==============================
function Invoke-CreateISOWork {
    if ($Clean) {
        if ($DryRun) {
            Write-Output "[DryRun] Would remove: $DestISO"
        } else {
            if (Test-Path $DestISO) {
                Write-Output "Removing: $DestISO"
                Remove-Item $DestISO -Force
            }
        }
        return
    }

    if ($DryRun) {
        Write-Output "[DryRun] Would write: $DestISO"
        return
    }

    Write-Output "Starting Create ISO workflow..."

    $etfs = $paths.BIOSInDest
    $efis = $paths.UEFIInDest
    # Sanity check for boot files before invoking oscdimg
    if (-not (Test-Path $etfs)) { throw "Missing BIOS boot file: $etfs" }
    if (-not (Test-Path $efis)) { throw "Missing UEFI boot file: $efis" }

    $IsoVolumeLabel = "Win$($OS)_$($Version)_$($Arch)_KBs"

    $bootdata = "2#p0,e,b$etfs#pEF,e,b$efis"
    $oscdimgfsArgs = @(
        "-m",                   # Ignore size limits
        "-o",                   # Optimize storage by encoding duplicate files only once
        "-u2",                  # Use UTF-8 encoding for file names (allows for long file names and Unicode characters) 
        "-udfver102",           # Use UDF 1.02 filesystem version (max compatibility, required for some boot scenarios)
        "-l$($IsoVolumeLabel)", # Set volume label
        "-bootdata:$bootdata",  # Define multi-boot configuration for BIOS and UEFI
        $paths.DestIsoRoot,
        $DestISO)

    Write-Output  "Building ISO: $DestISO"
    Write-Verbose "oscdimg $($oscdimgfsArgs -join ' ')"
    $oscdimgOut = & $oscdimgExe $oscdimgfsArgs 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "OSCDIMG failed to build ISO $DestISO (exit $LASTEXITCODE)"
    }
}

# Real work starts here

# Apply folder default
if (-not $Folder) { $Folder = (Get-Location).ProviderPath }

# ==============================
# Resolve working folder
# ==============================
$Folder = (Resolve-Path -LiteralPath $Folder).ProviderPath
Write-Verbose "Resolved working folder: $Folder"

# ==============================
# Resolve DISM and oscdimg
# ==============================
Write-Verbose "Resolving tool paths..."
$dismExe    = Resolve-DismExe    -ExplicitPath $dism    -PreferADK:$UseADK -ForceSystem:$UseSystem
$oscdimgExe = Resolve-OscdimgExe -ExplicitPath $oscdimg -PreferADK:$UseADK -ForceSystem:$UseSystem
Write-Output "dism    : $dismExe"
if ($oscdimgExe) {
    Write-Output "oscdimg : $oscdimgExe"
} else {
    Write-Output "oscdimg : not found (ISO creation unavailable)"
}

# ==============================
# Core paths (requires $Folder)
# ==============================
$paths = [ordered]@{}
$paths.BootWimInIso          = Join-Path $names.Sources $names.BootWim
$paths.InstallEsdInIso       = Join-Path $names.Sources $names.InstallEsd
$paths.InstallWimInIso       = Join-Path $names.Sources $names.InstallWim
$paths.SrcIsoRoot            = Join-Path $Folder $names.SrcIso
$paths.BIOSInSrc             = Join-Path $paths.SrcIsoRoot $names.BootFileBIOS
$paths.UEFIInSrc             = Join-Path $paths.SrcIsoRoot $names.BootFileUEFI
$paths.SourcesInSrc          = Join-Path $paths.SrcIsoRoot $names.Sources
$paths.BootWimInSrc          = Join-Path $paths.SourcesInSrc $names.BootWim
$paths.InstallEsdInSrc       = Join-Path $paths.SourcesInSrc $names.InstallEsd
$paths.InstallWimInSrc       = Join-Path $paths.SourcesInSrc $names.InstallWim
$paths.DestIsoRoot           = Join-Path $Folder $names.DestIso
$paths.SourcesInDest         = Join-Path $paths.DestIsoRoot $names.Sources
$paths.BootWimInDest         = Join-Path $paths.SourcesInDest $names.BootWim
$paths.InstallWimInDest      = Join-Path $paths.SourcesInDest $names.InstallWim
$paths.BIOSInDest            = Join-Path $paths.DestIsoRoot $names.BootFileBIOS
$paths.UEFIInDest            = Join-Path $paths.DestIsoRoot $names.BootFileUEFI
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
$paths.WimsCheckpoint        = Join-Path $paths.Checkpoint $names.Wims

# ==============================
# Resolve source ISO
# ==============================
if (-not $ISO) {
    Write-Verbose "No -ISO specified; searching for *.iso in: $Folder"
    $isoFiles = @(Get-ChildItem -Path $Folder -Filter '*.iso' -File -ErrorAction SilentlyContinue)
    if ($isoFiles.Count -eq 0) {
        # ISO is only required when -Export (or -All / default) is in effect
        $needsISO = $Export -or (-not ($KB -or $Service -or $Drivers -or $Reg -or $Files))
        if ($needsISO -and -not $Clean -and -not $DryRun) {
            Write-Error "No .iso file found in: $Folder`nPlace the Windows ISO there or use -ISO to specify its path."
            exit 1
        }
        Write-Verbose "No ISO found; continuing (ISO not required for selected operations)"
        $ISO = $null
    } elseif ($isoFiles.Count -gt 1) {
        Write-Error ("Multiple .iso files found in: $Folder`n  {0}`nUse -ISO to specify which one to use." -f ($isoFiles.FullName -join "`n  "))
        exit 1
    } else {
        $ISO = $isoFiles[0].FullName
        Write-Output "Auto-discovered ISO: $ISO"
    }
}

if ($ISO -and (Test-Path $ISO)) {
    $ISO = (Resolve-Path -LiteralPath $ISO).ProviderPath
    Write-Verbose "Resolved ISO path: $ISO"
}

# ==============================
# Resolve destination ISO
# ==============================
if (-not $DestISO -and $ISO) {
    $DestISO = $ISO -replace '\.iso$', '_KBs.iso'
    Write-Verbose "Auto-derived DestISO: $DestISO"
}

# ==============================
# Read ISO / WIM metadata for WinOS / Version / Arch and index list
# ==============================
$allImages       = @()
$isoMetaResolved = $false

# Prefer already-copied SrcISO when resuming
$srcIsoCopyDone = Join-Path $paths.Checkpoint "srciso.copy.done"
if ((-not $DryRun -or -not $Clean) -and (Test-Path $srcIsoCopyDone) -and (Test-Path $paths.SourcesInSrc)) {
    $existingWim = if (Test-Path $paths.InstallWimInSrc) {
        $paths.InstallWimInSrc
    } elseif (Test-Path $paths.InstallEsdInSrc) {
        $paths.InstallEsdInSrc
    } else { $null }

    if ($existingWim) {
        Write-Verbose "Reading metadata from existing SrcISO: $existingWim"
        try {
            $allImages = @(Get-WimImageList -WimPath $existingWim)
            if (-not $WinOS -or -not $Version -or -not $Arch) {
                $meta = Get-ISOMetadataFromWim -WimPath $existingWim
                if (-not $WinOS)   { $WinOS   = $meta.WinOS;   Write-Verbose "WinOS auto-detected   : $WinOS" }
                if (-not $Version) { $Version = $meta.Version; Write-Verbose "Version auto-detected : $Version" }
                if (-not $Arch)    { $Arch    = $meta.Arch;    Write-Verbose "Arch auto-detected    : $Arch" }
            }
            $isoMetaResolved = $true
        } catch {
            Write-Warning "Could not read metadata from SrcISO: $_"
        }
    }
}

# Mount the ISO briefly if we still need metadata
if (-not $isoMetaResolved -and $ISO -and (Test-Path $ISO) -and -not $DryRun -and -not $Clean) {
    Write-Verbose "Mounting ISO for metadata: $ISO"
    $metaDiskImg = Mount-DiskImage -ImagePath $ISO -PassThru -ErrorAction SilentlyContinue
    if ($metaDiskImg) {
        try {
            $metaVol   = $metaDiskImg | Get-Volume
            $metaDrive = $metaVol.DriveLetter + ':\'
            $metaWim   = if (Test-Path (Join-Path $metaDrive "sources\$($names.InstallWim)")) {
                             Join-Path $metaDrive "sources\$($names.InstallWim)"
                         } elseif (Test-Path (Join-Path $metaDrive "sources\$($names.InstallEsd)")) {
                             Join-Path $metaDrive "sources\$($names.InstallEsd)"
                         } else { $null }

            if ($metaWim) {
                Write-Verbose "Reading metadata from mounted ISO: $metaWim"
                $allImages = @(Get-WimImageList -WimPath $metaWim)
                if (-not $WinOS -or -not $Version -or -not $Arch) {
                    $meta = Get-ISOMetadataFromWim -WimPath $metaWim
                    if (-not $WinOS)   { $WinOS   = $meta.WinOS;   Write-Verbose "WinOS auto-detected   : $WinOS" }
                    if (-not $Version) { $Version = $meta.Version; Write-Verbose "Version auto-detected : $Version" }
                    if (-not $Arch)    { $Arch    = $meta.Arch;    Write-Verbose "Arch auto-detected    : $Arch" }
                }
                $isoMetaResolved = $true
            }
        } finally {
            Dismount-DiskImage -ImagePath $ISO -ErrorAction SilentlyContinue | Out-Null
        }
    }
}

# Hard defaults for anything still unresolved
if (-not $WinOS)   { $WinOS   = '11' }
if (-not $Arch)    { $Arch    = 'x64' }
if (-not $Version) { $Version = if ($WinOS -eq '10') { '22H2' } else { '25H2' } }

# ==============================
# ShowIndices
# ==============================
if ($ShowIndices) {
    if ($allImages.Count -eq 0) {
        Write-Error "Cannot show indices: no ISO is accessible and SrcISO has not been populated yet. Run -Export first, or provide -ISO."
        exit 1
    }
    Write-Output "`nAvailable images in $($names.InstallWim):`n"
    Write-Output ("{0,6}  {1}" -f 'Index', 'Name')
    Write-Output ("{0,6}  {1}" -f '------', '----')
    foreach ($img in $allImages) {
        Write-Output ("{0,6}  {1}" -f $img.Index, $img.Name)
    }
    Write-Output ""
    exit 0
}

# ==============================
# Resolve index selection
# ==============================
$SelectedIndices = @()
if ($allImages.Count -gt 0) {
    $SelectedIndices = @(Resolve-IndexSelection -AllImages $allImages -SelectHome:$SelectHome -SelectPro:$SelectPro -IndicesStr $Indices)
} else {
    Write-Verbose "Image list unavailable yet; index selection deferred until Export"
    # Will be populated during Invoke-ExportWork from the already-mounted ISO
    $SelectedIndices = @()
}

# ==============================
# Determine work modes
# ==============================
$workSwitches = @()
if ($Export)    { $workSwitches += 'Export' }
if ($KB)        { $workSwitches += 'KB' }
if ($Service)   { $workSwitches += 'Service' }
if ($Drivers)   { $workSwitches += 'Drivers' }
if ($Reg)       { $workSwitches += 'Reg' }
if ($Files)     { $workSwitches += 'Files' }
if ($CreateISO) { $workSwitches += 'CreateISO' }

if (-not $workSwitches) {
    $Export    = $true
    $KB        = $true
    $Service   = $true
    $Drivers   = $true
    $Reg       = $true
    $Files     = $true
    $CreateISO = $true
    $workSwitches = @('All')
}

Write-Output "Target profile: Windows $WinOS $Version $Arch"
Write-Output "Root folder   : $Folder"
Write-Output "ISO           : $(if ($ISO) { $ISO } else { '(none)' })"
Write-Output "DestISO       : $(if ($DestISO) { $DestISO } else { '(none)' })"
Write-Output "Selected idx  : $(if ($SelectedIndices.Count -gt 0) { $SelectedIndices.Index -join ', ' } else { 'all (determined at export time)' })"
Write-Output "Mode          : $($workSwitches -join ', ')"
if ($Clean)  { Write-Output "Clean mode    : Enabled" }
if ($DryRun) { Write-Output "Dry-run mode  : Enabled" }

if ($KB) { # Only KB workflow needs HTML parsing, so we delay this until now
    # --- HtmlAgilityPack bootstrap (PS 5.x SAFE) ---------------------------------
    $HtmlAgilityPackDll = 'HtmlAgilityPack.dll'
    $hapDll = Join-Path $Folder $HtmlAgilityPackDll

    if ($Clean) {
        if ($DryRun) {
            Write-Output "[DryRun] Would remove: $hapDll"
        } elseif (Test-Path $hapDLL) {
            Write-Output "Removing: $hapDLL"
            Remove-Item $hapDLL -Force
        }
    }
    elseif ($DryRun) {
        if (-not (Test-Path $hapDll)) {
            Write-Output "[DryRun] Would download: $HtmlAgilityPackDll"
        }   
    } else {
        if (-not (Test-Path $hapDll)) {
            Write-Output "HtmlAgilityPack.dll not found - downloading..."

            $nugetUrl   = "https://www.nuget.org/api/v2/package/HtmlAgilityPack"
            $tmpNupkg   = Join-Path $PSScriptRoot "HtmlAgilityPack.nupkg"
            $extractDir = Join-Path $PSScriptRoot "HtmlAgilityPack_Extract"

            # Clean old extraction folder if it exists
            if (Test-Path $extractDir) {
                Remove-Item $extractDir -Recurse -Force
            }

            # --- Download using .NET WebClient (PS 5.x safe) ---
            Write-Verbose "Downloading via WebClient..."
            $wc = New-Object System.Net.WebClient
            $wc.DownloadFile($nugetUrl, $tmpNupkg)

            # --- Extract using .NET ZipFile (PS 5.x safe) ---
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [System.IO.Compression.ZipFile]::ExtractToDirectory($tmpNupkg, $extractDir)

            # Prefer netstandard2.0, fallback to net45
            $candidatePaths = @(
                (Join-Path $extractDir "lib\netstandard2.0\$HtmlAgilityPackDll"),
                (Join-Path $extractDir "lib\net45\$HtmlAgilityPackDll")
            )

            $sourceDll = $candidatePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
            if (-not $sourceDll) {
                throw "$HtmlAgilityPackDll not found inside NuGet package"
            }

            Copy-Item -Path $sourceDll -Destination $hapDll -Force
            Write-Verbose "$HtmlAgilityPackDll copied to: $hapDll"

            # Cleanup: remove extraction folder + nupkg
            Remove-Item $extractDir -Recurse -Force
            Remove-Item $tmpNupkg -Force
        }

        # --- Load the DLL (PS 5.x safe) ---
        $hapLoaded = $false
        try {
            [void][HtmlAgilityPack.HtmlDocument]
            $hapLoaded = $true
        } catch {}

        if (-not $hapLoaded) {
            Write-Verbose "Loading HtmlAgilityPack from: $hapDll"
            Add-Type -Path $hapDll
            Write-Debug "HtmlAgilityPack successfully loaded"
        }
    }
}


# ==============================
# Main orchestration
# ==============================

if ($Export)  { Invoke-ExportWork }
if ($KB)      { Invoke-KBWork }
if ($Service) { Invoke-ServiceWork }
if ($Drivers) { Invoke-DriverWork }
if ($Reg)     { Invoke-RegWork }
if ($Files) {
    Write-InstallDriversCmd
    Write-InstallRegsCmd
    Write-PostSetupCmd
    Write-SetupConfigFiles
    Write-SetupCmdFiles
}
if ($CreateISO) { Invoke-CreateISOWork }

Write-Output "Completed"
