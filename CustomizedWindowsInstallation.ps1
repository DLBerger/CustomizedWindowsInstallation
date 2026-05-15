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

.PARAMETER Extract
Mount the source ISO and extract its full content tree to <Folder>\SrcISO\Content\.
Alias: -ExtractISO

.PARAMETER Export
Export selected indices from SrcISO\Content\ into per-index uncompressed WIMs under Wims\Indices\.

.PARAMETER KB
Download OS and .NET updates.

.PARAMETER Service
Apply downloaded KBs to the exported indices and produce final install.wim and boot.wim in Wims\Final\.

.PARAMETER Drivers
Export drivers into $WinpeDriver$.

.PARAMETER Reg
Export registry keys.

.PARAMETER Files
Generate PostSetup.cmd, SetupConfig-*.ini, and additional .cmd files.

.PARAMETER Prep
Hardlink-copy SrcISO\Content\ to DestISO\Content\, then place the final WIMs from Wims\Final\.
Alias: -PrepDestISO

.PARAMETER CreateISO
Create the final .iso from DestISO\Content\ using oscdimg.

.PARAMETER All
Shorthand for -Extract -Export -KB -Service -Drivers -Reg -Files -Prep -CreateISO.
Default when no specific switch is provided.

.PARAMETER Most
Same as -All without -CreateISO.

.PARAMETER ShowIndices
Print available image indices from the source ISO (or cached metadata) and exit.

.PARAMETER Home
Select editions whose normalized label matches "Home" exactly.

.PARAMETER Pro
Select editions whose normalized label matches "Pro" exactly.

.PARAMETER Indices
Comma-separated index selector supporting:
- Numbers:        6
- Ranges:         3-6, 7-*
- Exact labels:   "Education N"
- Wildcard labels: "*Home*", "* N*"
- Regex labels:   "re:^Education( N)?$"

.PARAMETER ISO
Explicit path to source ISO.
If omitted, the script discovers the single .iso file in <Folder>.
If more than one .iso is present an error is raised; use this parameter to disambiguate.

.PARAMETER DestISO
Explicit path to destination ISO.
If omitted, the source ISO path is reused with the extension changed to .bundled.iso.

.PARAMETER UpdateISO
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

    [string]$ISO,

    [string]$DestISO,

    [Alias('OS')]
    [ValidateSet('10','11')]
    [string]$WinOS,

    [string]$Version,

    [ValidateSet('x64','arm64')]
    [string]$Arch,

    [Alias('ExtractISO')]
    [switch]$Extract,

    [switch]$Export,
    [switch]$KB,
    [switch]$Service,
    [switch]$Drivers,
    [switch]$Reg,
    [switch]$Files,

    [Alias('PrepDestISO')]
    [switch]$Prep,

    [switch]$CreateISO,

    [switch]$All,

    [switch]$Most,

    [switch]$ShowIndices,

    [switch]$Home,

    [switch]$Pro,

    [string]$Indices,

    [switch]$UpdateISO,

    [switch]$UseADK,

    [switch]$UseSystem,

    [string]$dism,

    [string]$oscdimg,

    [switch]$Clean,

    [switch]$DryRun,

    [switch]$Help
)

# git hash
$GitHash = "9537dce"

# ==============================
# Core names
# ==============================
$names = [ordered]@{
    SrcIso                = 'SrcISO'
    DestIso               = 'DestISO'
    KBs                   = 'KBs'
    Wims                  = 'Wims'
    WinpeDriver           = '$WinpeDriver$'
    Registry              = 'Registry'
    Content               = 'Content'
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

$wimDirs = @('Indices', 'Mounts', 'Serviced', 'Final')
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

function Read-JsonFile {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path -PathType Leaf)) { return $null }
    try   { Get-Content $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop }
    catch { Write-Warning "Failed to read JSON '$Path': $_"; $null }
}

function Write-JsonFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][object]$Data
    )
    Ensure-Folder -Path (Split-Path $Path -Parent)
    $Data | ConvertTo-Json -Depth 6 | Set-Content -Path $Path -Encoding UTF8
}

function Run-App {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Exe,
        [string[]]$AppArgs = @(),
        [switch]$ShowOutput
    )

    Write-Debug "Run-App: '$Exe' $($AppArgs -join ' ')"

    $psi                        = [System.Diagnostics.ProcessStartInfo]::new($Exe)
    $psi.Arguments              = $AppArgs -join ' '
    $psi.UseShellExecute        = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.CreateNoWindow         = $true

    $proc = [System.Diagnostics.Process]::Start($psi)
    $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
    $stderrTask = $proc.StandardError.ReadToEndAsync()
    $proc.WaitForExit()
    [void][System.Threading.Tasks.Task]::WaitAll($stdoutTask, $stderrTask)

    $seen    = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    $outList = [System.Collections.Generic.List[string]]::new()

    foreach ($src in @($stdoutTask.Result, $stderrTask.Result)) {
        foreach ($line in ($src -split "`r?`n")) {
            if ($line.Length -gt 0 -and $seen.Add($line)) {
                $outList.Add($line)
            }
        }
    }

    if ($ShowOutput -or $DebugPreference -eq 'Continue') {
        $outList | ForEach-Object { Write-Debug "  OUT> $_" }
    }

    return [PSCustomObject]@{
        ExitCode = $proc.ExitCode
        Output   = $outList.ToArray()
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
        [string]$WimPath,
        [Parameter(Mandatory)]
        [string]$DismExe
    )

    Write-Debug "Get-WimImageList: WimPath='$WimPath'"
    Write-Verbose "Reading image list from: $WimPath"

    $output = & $DismExe /Get-WimInfo "/WimFile:$WimPath" 2>&1

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
        [string]$WimPath,
        [Parameter(Mandatory)]
        [string]$DismExe
    )

    Write-Debug "Get-ISOMetadataFromWim: WimPath='$WimPath'"

    $output = & $DismExe /Get-WimInfo "/WimFile:$WimPath" /Index:1 2>&1

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

# =========================
# Extract / Export / Prep / CreateISO section
# =========================

function Invoke-ExtractISO {
    [CmdletBinding()]
    param()

    Write-Verbose "Invoke-ExtractISO: ISO='$ISO' SrcIsoContent='$($paths.SrcIsoContent)'"
    Write-Log "Starting ExtractISO workflow..."

    $extractJson = Join-Path $paths.SrcIsoRoot "extract.json"

    if ($DryRun) {
        Write-Log "[DryRun] Would mount ISO: $ISO"
        Write-Log "[DryRun] Would robocopy entire ISO tree -> $($paths.SrcIsoContent)"
        Write-Log "[DryRun] Would validate $($names.BootWim) and ($($names.InstallWim) or $($names.InstallEsd)) in $($paths.SourcesInSrc)"
        Write-Log "[DryRun] Would write extract.json to $extractJson"
        return
    }

    if ($Clean) {
        Write-Debug "Invoke-ExtractISO: Clean mode"
        if (Test-Path $paths.SrcIsoRoot) {
            Write-Log "Removing: $($paths.SrcIsoRoot)"
            Remove-Item $paths.SrcIsoRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        return
    }

    if (-not $ISO -or -not (Test-Path $ISO)) {
        throw "Source ISO not found or not specified. Use -ISO to specify the Windows .iso file."
    }

    # Check existing extract.json: skip if same ISO path, clean+re-extract if different
    $existingJson = Read-JsonFile -Path $extractJson
    if ($existingJson) {
        if ($existingJson.ISOPath -eq $ISO) {
            Write-Log "ExtractISO already done for this ISO (extract.json match)"
            Write-Debug "extract.json: ISOPath='$($existingJson.ISOPath)' Date='$($existingJson.Date)'"
            return
        }
        Write-Log "ISO path changed (was '$($existingJson.ISOPath)'); cleaning SrcIsoRoot..."
        Remove-Item $paths.SrcIsoRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    Ensure-Folder -Path $paths.SrcIsoContent

    Write-Log "Mounting ISO: $ISO"
    $diskImage = Mount-DiskImage -ImagePath $ISO -PassThru -ErrorAction Stop
    try {
        $vol         = $diskImage | Get-Volume
        $driveLetter = $vol.DriveLetter + ':\'
        Write-Log "ISO mounted at: $driveLetter"
        Write-Debug "Volume: DriveLetter='$($vol.DriveLetter)' FileSystem='$($vol.FileSystem)'"

        Write-Log "Copying entire ISO tree -> $($paths.SrcIsoContent)..."
        $roboArgs = @($driveLetter, $paths.SrcIsoContent, '/E', '/R:2', '/W:1', '/NP', '/NDL', '/NC')
        Write-Verbose "robocopy $($roboArgs -join ' ')"
        $rc = (Run-App -Exe 'robocopy.exe' -AppArgs $roboArgs).ExitCode
        if ($rc -gt 7) { throw "robocopy failed (exit $rc)" }
        Write-Debug "robocopy exit $rc"

    } finally {
        Write-Log "Unmounting ISO..."
        Dismount-DiskImage -ImagePath $ISO -ErrorAction SilentlyContinue | Out-Null
    }

    # Validate required files
    Write-Log "Validating required files in $($paths.SourcesInSrc)..."
    $missing = [System.Collections.Generic.List[string]]::new()
    if (-not (Test-Path $paths.BootWimInSrc)) {
        $missing.Add("$($names.BootWim)  (expected: $($paths.BootWimInSrc))")
    }
    if (-not (Test-Path $paths.InstallWimInSrc) -and -not (Test-Path $paths.InstallEsdInSrc)) {
        $missing.Add("$($names.InstallWim) or $($names.InstallEsd)  (expected: $($paths.InstallWimInSrc) / $($paths.InstallEsdInSrc))")
    }
    if ($missing.Count -gt 0) {
        throw "Source ISO validation failed. Missing required file(s):`n  - $($missing -join "`n  - ")"
    }

    Write-Log "Source ISO validation passed"
    Write-JsonFile -Path $extractJson -Data @{ ISOPath = $ISO; Date = (Get-Date -Format s) }
    Write-Log "ExtractISO complete (extract.json written)"
}

function Invoke-Export {
    [CmdletBinding()]
    param()

    Write-Verbose "Invoke-Export: SourcesInSrc='$($paths.SourcesInSrc)' WimsIndices='$($paths.WimsIndices)'"
    Write-Log "Starting Export workflow..."

    $extractJson    = Join-Path $paths.SrcIsoRoot "extract.json"
    $metadataJson   = Join-Path $paths.WimsIndices "wim-metadata.json"

    if ($DryRun) {
        Write-Log "[DryRun] Would collect WIM metadata from: $($paths.SourcesInSrc)"
        Write-Log "[DryRun] Would export $($SelectedIndices.Count) indices to $($paths.WimsIndices)"
        return
    }

    if ($Clean) {
        Write-Debug "Invoke-Export: Clean mode"
        if (Test-Path $paths.WimsRoot) {
            Write-Log "Removing: $($paths.WimsRoot)"
            Remove-Item $paths.WimsRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        return
    }

    $extractMeta = Read-JsonFile -Path $extractJson
    if (-not $extractMeta) {
        throw "extract.json not found at '$extractJson'. Run -Extract first."
    }
    $extractDate = [datetime]::Parse($extractMeta.Date)
    Write-Verbose "Extract date: $extractDate"

    Ensure-Folder -Path $paths.WimsRoot
    Ensure-Folder -Path $paths.WimsIndices

    # Collect WIM metadata and save to wim-metadata.json
    $installSrc = if (Test-Path $paths.InstallWimInSrc) { $paths.InstallWimInSrc } else { $paths.InstallEsdInSrc }
    $bootSrc    = $paths.BootWimInSrc

    Write-Log "Collecting WIM metadata..."
    Write-Verbose "install source: $installSrc"
    Write-Verbose "boot source   : $bootSrc"

    $installImages = @(Get-WimImageList -WimPath $installSrc -DismExe $dismExe)
    $bootImages    = @(Get-WimImageList -WimPath $bootSrc    -DismExe $dismExe)
    $installMeta   = Get-ISOMetadataFromWim -WimPath $installSrc -DismExe $dismExe

    $wimMeta = @{
        ISOPath       = $extractMeta.ISOPath
        CollectedDate = (Get-Date -Format s)
        WinOS         = $installMeta.WinOS
        Version       = $installMeta.Version
        Arch          = $installMeta.Arch
        Build         = $installMeta.Build
        InstallImages = @($installImages | ForEach-Object { @{ Index = $_.Index; Name = $_.Name } })
        BootImages    = @($bootImages    | ForEach-Object { @{ Index = $_.Index; Name = $_.Name } })
    }
    Write-JsonFile -Path $metadataJson -Data $wimMeta
    Write-Log "WIM metadata saved: $metadataJson"
    Write-Debug "InstallImages: $($installImages.Count)  BootImages: $($bootImages.Count)"

    # Resolve index selection if deferred
    if ($SelectedIndices.Count -eq 0) {
        Write-Verbose "SelectedIndices empty; resolving from collected metadata..."
        $SelectedIndices = @(Resolve-IndexSelection -AllImages $installImages -SelectHome:$Home -SelectPro:$Pro -IndicesStr $Indices)
        Write-Verbose "Late index resolution: $($SelectedIndices.Count) index/indices"
    }

    Write-Log "Exporting $($SelectedIndices.Count) index/indices..."
    Write-Verbose "Selected indices: $($SelectedIndices.Index -join ', ')"

    # Determine boot.wim source index (index 2 = Windows Setup PE, fallback 1)
    $bootSrcIdx = if ($bootImages | Where-Object { $_.Index -eq 2 }) { 2 } else { 1 }

    foreach ($img in $SelectedIndices) {
        $idx     = $img.Index
        $imgName = $img.Name
        Write-Log "  [Index $idx] $imgName"
        Write-Debug "  Index=$idx Name='$imgName'"

        # -- Export install image --
        $installDest = Join-Path $paths.WimsIndices ("{0}_{1}" -f $idx, $names.InstallWim)
        $installJson = "$installDest.json"
        $existInstall = Read-JsonFile -Path $installJson
        $needInstall  = (-not $existInstall) -or ([datetime]::Parse($existInstall.ExportDate) -le $extractDate)

        if (-not $needInstall) {
            Write-Log "    install.wim index $idx already exported (JSON: $($existInstall.ExportDate))"
        } else {
            Write-Log "    Exporting install.wim index $idx -> $(Split-Path $installDest -Leaf)"
            Write-Verbose "    dism /Export-Image /SourceIndex:$idx /Compress:None"
            $r = Run-App -Exe $dismExe -AppArgs @('/Export-Image', "/SourceImageFile:$installSrc",
                "/SourceIndex:$idx", "/DestinationImageFile:$installDest", '/Compress:None')
            if ($DebugPreference -eq 'Continue') { $r.Output | ForEach-Object { Write-Debug "    DISM> $_" } }
            if ($r.ExitCode -ne 0) { throw "DISM export failed for install index $idx (exit $($r.ExitCode))" }
            Write-JsonFile -Path $installJson -Data @{ Index = $idx; Name = $imgName; ExportDate = (Get-Date -Format s) }
            Write-Log "    install.wim index $idx exported"
        }

        # -- Export boot image --
        $bootDest = Join-Path $paths.WimsIndices ("{0}_{1}" -f $idx, $names.BootWim)
        $bootJson = "$bootDest.json"
        $existBoot = Read-JsonFile -Path $bootJson
        $needBoot  = (-not $existBoot) -or ([datetime]::Parse($existBoot.ExportDate) -le $extractDate)

        if (-not $needBoot) {
            Write-Log "    boot.wim index $idx already exported (JSON: $($existBoot.ExportDate))"
        } else {
            Write-Log "    Exporting boot.wim (source idx $bootSrcIdx) -> $(Split-Path $bootDest -Leaf)"
            Write-Verbose "    dism /Export-Image boot.wim /SourceIndex:$bootSrcIdx /Compress:None"
            $rb = Run-App -Exe $dismExe -AppArgs @('/Export-Image', "/SourceImageFile:$bootSrc",
                "/SourceIndex:$bootSrcIdx", "/DestinationImageFile:$bootDest", '/Compress:None')
            if ($DebugPreference -eq 'Continue') { $rb.Output | ForEach-Object { Write-Debug "    DISM> $_" } }
            if ($rb.ExitCode -ne 0) {
                Write-Warning "    DISM boot export failed (exit $($rb.ExitCode)); skipping"
            } else {
                Write-JsonFile -Path $bootJson -Data @{ Index = $idx; SourceBootIndex = $bootSrcIdx; ExportDate = (Get-Date -Format s) }
                Write-Log "    boot.wim index $idx exported"
            }
        }
    }

    Write-Log "Export workflow complete"
}

function Invoke-PrepDestISO {
    [CmdletBinding()]
    param()

    Write-Verbose "Invoke-PrepDestISO: SrcIsoContent='$($paths.SrcIsoContent)' DestIsoContent='$($paths.DestIsoContent)'"
    Write-Log "Starting PrepDestISO workflow..."

    $prepJson     = Join-Path $paths.DestIsoRoot "prep.json"
    $extractJson  = Join-Path $paths.SrcIsoRoot  "extract.json"
    $finalInstall = Join-Path $paths.WimsFinal $names.InstallWim
    $finalBoot    = Join-Path $paths.WimsFinal $names.BootWim
    $finalJson    = Join-Path $paths.WimsFinal "final.json"

    if ($DryRun) {
        Write-Log "[DryRun] Would hardlink-copy $($paths.SrcIsoContent) -> $($paths.DestIsoContent)"
        Write-Log "[DryRun] Would copy $finalInstall -> $($paths.InstallWimInDest)"
        Write-Log "[DryRun] Would copy $finalBoot    -> $($paths.BootWimInDest)"
        Write-Log "[DryRun] Would write prep.json to $prepJson"
        return
    }

    if ($Clean) {
        Write-Debug "Invoke-PrepDestISO: Clean mode"
        if (Test-Path $paths.DestIsoRoot) {
            Write-Log "Removing: $($paths.DestIsoRoot)"
            Remove-Item $paths.DestIsoRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        return
    }

    $extractMeta = Read-JsonFile -Path $extractJson
    if (-not $extractMeta) {
        throw "extract.json not found at '$extractJson'. Run -Extract first."
    }
    $extractDate = [datetime]::Parse($extractMeta.Date)

    $finalMeta = Read-JsonFile -Path $finalJson
    $prep      = Read-JsonFile -Path $prepJson

    # Step A: Hardlink-copy SrcIsoContent -> DestIsoContent (excluding install/boot WIMs)
    $needHardlink = (-not $prep -or -not $prep.HardlinkDate) -or
                    ([datetime]::Parse($prep.HardlinkDate) -le $extractDate)

    if ($needHardlink) {
        Write-Log "Hardlink-copying $($paths.SrcIsoContent) -> $($paths.DestIsoContent) (excluding install/boot images)..."
        if (Test-Path $paths.DestIsoRoot) {
            Write-Log "Cleaning existing DestIsoRoot before hardlink rebuild..."
            Remove-Item $paths.DestIsoRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        Ensure-Folder -Path $paths.DestIsoContent

        $excludeNames = @($names.BootWim, $names.InstallWim, $names.InstallEsd)
        Write-Verbose "Excluding: $($excludeNames -join ', ')"

        $allFiles = @(Get-ChildItem -Path $paths.SrcIsoContent -Recurse -File -ErrorAction SilentlyContinue)
        $total    = $allFiles.Count
        $done     = 0
        $lastPct  = -1

        Write-Host ("Hardlinking {0} files: '{1}' -> '{2}'" -f $total, $paths.SrcIsoContent, $paths.DestIsoContent)

        foreach ($file in $allFiles) {
            $done++
            $pct = [math]::Floor(($done / [math]::Max($total, 1)) * 100)
            if ($pct -ge ($lastPct + 10)) {
                Write-Host ("  {0,3}%  {1}/{2} files" -f $pct, $done, $total)
                $lastPct = $pct - ($pct % 10)
            }
            if ($file.Name -in $excludeNames) {
                Write-Debug "  Skip (excluded): $($file.Name)"
                continue
            }
            $relPath  = $file.FullName.Substring($paths.SrcIsoContent.TrimEnd('\').Length).TrimStart('\')
            $destPath = Join-Path $paths.DestIsoContent $relPath
            Ensure-Folder -Path (Split-Path $destPath -Parent)
            if (-not (Test-Path $destPath)) {
                try {
                    New-Item -ItemType HardLink -Path $destPath -Value $file.FullName -Force -ErrorAction Stop | Out-Null
                } catch {
                    Write-Warning "Hardlink failed for '$relPath'; copying: $_"
                    Copy-Item -Path $file.FullName -Destination $destPath -Force
                }
            }
        }
        Write-Host ("  Hardlink tree complete: {0} files processed" -f $done)

        $now  = Get-Date -Format s
        $prep = @{ HardlinkDate = $now }
        Write-JsonFile -Path $prepJson -Data $prep
        Write-Log "DestIsoContent hardlink-copy complete"
    } else {
        Write-Log "DestIsoContent hardlink-copy already done (prep.json: $($prep.HardlinkDate))"
    }

    $prep = Read-JsonFile -Path $prepJson
    if (-not $prep) { $prep = @{} }

    # Step B: Copy final install.wim -> DestIsoContent\sources\install.wim
    $finalInstallDate = if ($finalMeta -and $finalMeta.InstallWimDate) { [datetime]::Parse($finalMeta.InstallWimDate) } else { [datetime]::MinValue }
    $destInstallDate  = if ($prep.InstallWimDate) { [datetime]::Parse($prep.InstallWimDate) } else { [datetime]::MinValue }
    $needInstall      = (Test-Path $finalInstall) -and ($destInstallDate -le $finalInstallDate)

    if ($needInstall) {
        Write-Log "Copying final $($names.InstallWim) -> $($paths.InstallWimInDest)..."
        Ensure-Folder -Path (Split-Path $paths.InstallWimInDest -Parent)
        Copy-Item -Path $finalInstall -Destination $paths.InstallWimInDest -Force
        $prep['InstallWimDate'] = (Get-Date -Format s)
        Write-JsonFile -Path $prepJson -Data $prep
        Write-Log "  $($names.InstallWim) placed"
    } elseif (Test-Path $paths.InstallWimInDest) {
        Write-Log "  $($names.InstallWim) already current (prep.json: $($prep.InstallWimDate))"
    } else {
        Write-Warning "Final $($names.InstallWim) not found at: $finalInstall (run -Service first)"
    }

    $prep = Read-JsonFile -Path $prepJson
    if (-not $prep) { $prep = @{} }

    # Step C: Copy final boot.wim -> DestIsoContent\sources\boot.wim
    $finalBootDate = if ($finalMeta -and $finalMeta.BootWimDate) { [datetime]::Parse($finalMeta.BootWimDate) } else { [datetime]::MinValue }
    $destBootDate  = if ($prep.BootWimDate) { [datetime]::Parse($prep.BootWimDate) } else { [datetime]::MinValue }
    $needBoot      = (Test-Path $finalBoot) -and ($destBootDate -le $finalBootDate)

    if ($needBoot) {
        Write-Log "Copying final $($names.BootWim) -> $($paths.BootWimInDest)..."
        Ensure-Folder -Path (Split-Path $paths.BootWimInDest -Parent)
        Copy-Item -Path $finalBoot -Destination $paths.BootWimInDest -Force
        $prep['BootWimDate'] = (Get-Date -Format s)
        Write-JsonFile -Path $prepJson -Data $prep
        Write-Log "  $($names.BootWim) placed"
    } elseif (Test-Path $paths.BootWimInDest) {
        Write-Log "  $($names.BootWim) already current (prep.json: $($prep.BootWimDate))"
    } else {
        Write-Warning "Final $($names.BootWim) not found at: $finalBoot (run -Service first)"
    }

    Write-Log "PrepDestISO workflow complete"
}

function Invoke-CreateISO {
    [CmdletBinding()]
    param()

    Write-Verbose "Invoke-CreateISO: DestIsoContent='$($paths.DestIsoContent)' DestISO='$DestISO'"
    Write-Log "Starting CreateISO workflow..."

    $prepJson = Join-Path $paths.DestIsoRoot "prep.json"

    if ($DryRun) {
        Write-Log "[DryRun] Would create ISO: $DestISO"
        Write-Log "[DryRun]   from: $($paths.DestIsoContent)"
        Write-Log "[DryRun]   BIOS boot: $($paths.BIOSInDest)"
        Write-Log "[DryRun]   UEFI boot: $($paths.UEFIInDest)"
        return
    }

    if ($Clean) {
        Write-Debug "Invoke-CreateISO: Clean mode"
        if ($DestISO -and (Test-Path $DestISO)) {
            Write-Log "Removing: $DestISO"
            Remove-Item $DestISO -Force -ErrorAction SilentlyContinue
        }
        return
    }

    $prepMeta = Read-JsonFile -Path $prepJson
    if (-not $prepMeta) {
        throw "prep.json not found at '$prepJson'. Run -Prep first to prepare the destination ISO content."
    }

    if (-not $oscdimgExe) {
        throw "oscdimg.exe not found. Install Windows ADK or specify -oscdimg."
    }
    if (-not $DestISO) {
        throw "DestISO path is not set. Specify -DestISO or ensure -ISO is provided."
    }
    if (-not (Test-Path $paths.DestIsoContent)) {
        throw "DestIsoContent folder does not exist: $($paths.DestIsoContent). Run -Prep first."
    }

    Write-Log "Creating ISO: $DestISO"
    Write-Verbose "Source   : $($paths.DestIsoContent)"
    Write-Verbose "BIOS boot: $($paths.BIOSInDest)"
    Write-Verbose "UEFI boot: $($paths.UEFIInDest)"

    # Build oscdimg command: dual-boot (BIOS + UEFI) per standard Windows ISO format
    $oscdimgArgs = @(
        '-m',             # ignore maximum size limit
        '-o',             # optimize storage (single-instance files)
        '-u2',            # UDF file system
        '-udfver102',     # UDF 1.02 compatibility
        "-bootdata:2#p0,e,b`"$($paths.BIOSInDest)`"#pEF,e,b`"$($paths.UEFIInDest)`"",
        '-h',             # include hidden files
        '-l',             # label – use ISO filename without extension
        (Split-Path $DestISO -Leaf) -replace '\..*$',
        $paths.DestIsoContent,
        $DestISO
    )

    Write-Verbose "oscdimg $($oscdimgArgs -join ' ')"
    $r = Run-App -Exe $oscdimgExe -AppArgs $oscdimgArgs -ShowOutput:($DebugPreference -eq 'Continue')
    if ($r.ExitCode -ne 0) {
        throw "oscdimg failed (exit $($r.ExitCode))"
    }

    Write-Log "ISO created: $DestISO"
    Write-Log "CreateISO workflow complete"
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

# =========================
# Service/Patch section
# =========================

function Invoke-ServiceWork {
    [CmdletBinding()]
    param()

    Write-Verbose "Invoke-ServiceWork: WimsIndices='$($paths.WimsIndices)' WimsFinal='$($paths.WimsFinal)'"
    Write-Debug   "Invoke-ServiceWork: KBsSSU='$($paths.KBsSSU)' KBsOSCU='$($paths.KBsOSCU)' WimsMounts='$($paths.WimsMounts)'"
    Write-Log "Starting Service workflow..."

    $finalJson = Join-Path $paths.WimsFinal "final.json"

    if ($DryRun) {
        Write-Log "[DryRun] Would service extracted indices in $($paths.WimsIndices)"
        Write-Log "[DryRun] Would apply SSU from : $($paths.KBsSSU)"
        Write-Log "[DryRun] Would apply LCU from : $($paths.KBsOSCU)"
        Write-Log "[DryRun] Would assemble final $($names.InstallWim) -> $($paths.WimsFinal)"
        Write-Log "[DryRun] Would assemble final $($names.BootWim)    -> $($paths.WimsFinal)"
        return
    }

    if ($Clean) {
        Write-Debug "Invoke-ServiceWork: Clean mode"
        if (Test-Path $paths.WimsRoot) {
            Write-Log "Removing: $($paths.WimsRoot)"
            Remove-Item $paths.WimsRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        return
    }

    $CompressionType = 'Maximum'  # None, Fast, Maximum

    Ensure-Folder -Path $paths.WimsMounts
    Ensure-Folder -Path $paths.WimsServiced

    if (-not (Test-Path $paths.WimsIndices)) {
        Write-Log "No indices folder found: $($paths.WimsIndices). Run with -Export first."
        return
    }

    $indexFiles = @(Get-ChildItem -Path $paths.WimsIndices -Filter "*_$($names.InstallWim)" -File -ErrorAction SilentlyContinue)
    if ($indexFiles.Count -eq 0) {
        Write-Log "No extracted $($names.InstallWim) files found. Run with -Export first."
        return
    }

    $extractedIndices = $indexFiles |
        ForEach-Object { [int](($_.BaseName -split '_')[0]) } |
        Sort-Object
    Write-Log "Found $($extractedIndices.Count) extracted indices: $($extractedIndices -join ', ')"

    $ssuFiles = @(Get-ChildItem -Path $paths.KBsSSU  -Include '*.msu','*.cab' -Recurse -ErrorAction SilentlyContinue)
    $lcuFiles = @(Get-ChildItem -Path $paths.KBsOSCU -Include '*.msu','*.cab' -Recurse -ErrorAction SilentlyContinue)
    $hasSSU = $ssuFiles.Count -gt 0
    $hasLCU = $lcuFiles.Count -gt 0

    Write-Log "Packages - SSU: $hasSSU ($($ssuFiles.Count)), LCU: $hasLCU ($($lcuFiles.Count))"
    Write-Verbose "SSU : $($ssuFiles.Name -join ', ')"
    Write-Verbose "LCU : $($lcuFiles.Name -join ', ')"

    foreach ($idx in $extractedIndices) {
        Write-Log "--- Index $idx ---"

        $installWimPath = Join-Path $paths.WimsIndices ("{0}_{1}" -f $idx, $names.InstallWim)
        $bootWimPath    = Join-Path $paths.WimsIndices ("{0}_{1}" -f $idx, $names.BootWim)
        $installSvcJson = (Join-Path $paths.WimsIndices ("{0}_{1}.serviced.json" -f $idx, $names.InstallWim))
        $bootSvcJson    = (Join-Path $paths.WimsIndices ("{0}_{1}.serviced.json" -f $idx, $names.BootWim))
        $winreSvcJson   = (Join-Path $paths.WimsIndices ("{0}_{1}.serviced.json" -f $idx, $names.WinreWim))
        $winreExtJson   = (Join-Path $paths.WimsIndices ("{0}_{1}.extracted.json" -f $idx, $names.WinreWim))
        $mountDir       = Join-Path $paths.WimsMounts ("mount_{0}"      -f $idx)
        $winreMountDir  = Join-Path $paths.WimsMounts ("winremount_{0}" -f $idx)
        $winreWimPath   = Join-Path $paths.WimsIndices ("{0}_{1}"       -f $idx, $names.WinreWim)

        # ---- Service install.wim ----
        $installSvcMeta = Read-JsonFile -Path $installSvcJson
        if ($installSvcMeta) {
            Write-Log "  $($names.InstallWim) index $idx already serviced ($($installSvcMeta.ServicedDate))"
        } elseif ($hasSSU -or $hasLCU) {
            Write-Log "  Servicing $($names.InstallWim) for index $idx..."
            Ensure-Folder -Path $mountDir

            try {
                Write-Log "  Mounting $installWimPath -> $mountDir"
                $r = Run-App -Exe $dismExe -AppArgs @('/Mount-Image', "/ImageFile:$installWimPath", '/Index:1', "/MountDir:$mountDir")
                if ($DebugPreference -eq 'Continue') { $r.Output | ForEach-Object { Write-Debug "  DISM> $_" } }
                if ($r.ExitCode -ne 0) { throw "DISM mount $($names.InstallWim) failed for index $idx (exit $($r.ExitCode))" }

                if ($hasSSU) {
                    $winreInMount = Join-Path $mountDir $paths.WinreWimInWim

                    if (Test-Path $winreInMount) {
                        $winreSvcMeta = Read-JsonFile -Path $winreSvcJson
                        if ($winreSvcMeta) {
                            Write-Log "  $($names.WinreWim) index $idx already serviced ($($winreSvcMeta.ServicedDate))"
                        } else {
                            if (-not (Read-JsonFile -Path $winreExtJson)) {
                                Write-Log "  Extracting $($names.WinreWim)..."
                                Copy-Item -Path $winreInMount -Destination $winreWimPath -Force
                                Write-JsonFile -Path $winreExtJson -Data @{ Index = $idx; ExtractedDate = (Get-Date -Format s) }
                            }

                            Ensure-Folder -Path $winreMountDir
                            Write-Log "  Mounting $($names.WinreWim) -> $winreMountDir"
                            $rw = Run-App -Exe $dismExe -AppArgs @('/Mount-Image', "/ImageFile:$winreWimPath", '/Index:1', "/MountDir:$winreMountDir")
                            if ($DebugPreference -eq 'Continue') { $rw.Output | ForEach-Object { Write-Debug "  DISM> $_" } }
                            if ($rw.ExitCode -ne 0) { throw "DISM mount $($names.WinreWim) failed for index $idx (exit $($rw.ExitCode))" }

                            foreach ($pkg in $ssuFiles) {
                                Write-Log "  Applying SSU to $($names.WinreWim): $($pkg.Name)"
                                $rp = Run-App -Exe $dismExe -AppArgs @('/Add-Package', "/Image:$winreMountDir", "/PackagePath:$($pkg.FullName)")
                                if ($DebugPreference -eq 'Continue') { $rp.Output | ForEach-Object { Write-Debug "  DISM> $_" } }
                                if ($rp.ExitCode -ne 0) { Write-Warning "  SSU->winre failed: $($pkg.Name) (exit $($rp.ExitCode)); continuing" }
                            }

                            Write-Log "  Unmounting $($names.WinreWim) (commit)..."
                            $ru = Run-App -Exe $dismExe -AppArgs @('/Unmount-Image', "/MountDir:$winreMountDir", '/Commit')
                            if ($DebugPreference -eq 'Continue') { $ru.Output | ForEach-Object { Write-Debug "  DISM> $_" } }
                            if ($ru.ExitCode -ne 0) { throw "DISM unmount $($names.WinreWim) failed for index $idx (exit $($ru.ExitCode))" }

                            Copy-Item -Path $winreWimPath -Destination $winreInMount -Force
                            Write-JsonFile -Path $winreSvcJson -Data @{ Index = $idx; ServicedDate = (Get-Date -Format s) }
                            Write-Log "  $($names.WinreWim) serviced and reinserted"
                        }
                    } else {
                        Write-Verbose "  $($names.WinreWim) not found in mounted install; skipping winre servicing"
                    }

                    foreach ($pkg in $ssuFiles) {
                        Write-Log "  Applying SSU to $($names.InstallWim): $($pkg.Name)"
                        $ri = Run-App -Exe $dismExe -AppArgs @('/Add-Package', "/Image:$mountDir", "/PackagePath:$($pkg.FullName)")
                        if ($DebugPreference -eq 'Continue') { $ri.Output | ForEach-Object { Write-Debug "  DISM> $_" } }
                        if ($ri.ExitCode -ne 0) { Write-Warning "  SSU->install failed: $($pkg.Name) (exit $($ri.ExitCode)); continuing" }
                    }
                }

                if ($hasLCU) {
                    foreach ($pkg in $lcuFiles) {
                        Write-Log "  Applying LCU to $($names.InstallWim): $($pkg.Name)"
                        $rl = Run-App -Exe $dismExe -AppArgs @('/Add-Package', "/Image:$mountDir", "/PackagePath:$($pkg.FullName)")
                        if ($DebugPreference -eq 'Continue') { $rl.Output | ForEach-Object { Write-Debug "  DISM> $_" } }
                        if ($rl.ExitCode -ne 0) { Write-Warning "  LCU->install failed: $($pkg.Name) (exit $($rl.ExitCode)); continuing" }
                    }
                }

                Write-Log "  Unmounting $($names.InstallWim) index $idx (commit)..."
                $rc = Run-App -Exe $dismExe -AppArgs @('/Unmount-Image', "/MountDir:$mountDir", '/Commit')
                if ($DebugPreference -eq 'Continue') { $rc.Output | ForEach-Object { Write-Debug "  DISM> $_" } }
                if ($rc.ExitCode -ne 0) { throw "DISM unmount $($names.InstallWim) failed for index $idx (exit $($rc.ExitCode))" }

                Write-JsonFile -Path $installSvcJson -Data @{ Index = $idx; ServicedDate = (Get-Date -Format s) }
                Write-Log "  $($names.InstallWim) index $idx serviced"

            } catch {
                Write-Log "  ERROR servicing $($names.InstallWim) index $idx`: $_" 'ERROR'
                if (Test-Path $mountDir)      { Run-App -Exe $dismExe -AppArgs @('/Unmount-Image', "/MountDir:$mountDir",      '/Discard') | Out-Null }
                if (Test-Path $winreMountDir) { Run-App -Exe $dismExe -AppArgs @('/Unmount-Image', "/MountDir:$winreMountDir", '/Discard') | Out-Null }
                throw
            } finally {
                if (Test-Path $mountDir)      { Remove-Item $mountDir      -Recurse -Force -ErrorAction SilentlyContinue }
                if (Test-Path $winreMountDir) { Remove-Item $winreMountDir -Recurse -Force -ErrorAction SilentlyContinue }
            }
        } else {
            Write-Log "  No SSU/LCU packages; marking $($names.InstallWim) index $idx as done"
            Write-JsonFile -Path $installSvcJson -Data @{ Index = $idx; ServicedDate = (Get-Date -Format s) }
        }

        # ---- Service boot.wim ----
        if (-not (Test-Path $bootWimPath)) {
            Write-Verbose "  No $($names.BootWim) for index $idx; skipping"
        } else {
            $bootSvcMeta = Read-JsonFile -Path $bootSvcJson
            if ($bootSvcMeta) {
                Write-Log "  $($names.BootWim) index $idx already serviced ($($bootSvcMeta.ServicedDate))"
            } elseif (-not $hasSSU) {
                Write-Log "  No SSU packages; marking $($names.BootWim) index $idx as done"
                Write-JsonFile -Path $bootSvcJson -Data @{ Index = $idx; ServicedDate = (Get-Date -Format s) }
            } else {
                Write-Log "  Servicing $($names.BootWim) for index $idx..."
                Ensure-Folder -Path $mountDir
                try {
                    Write-Log "  Mounting $bootWimPath -> $mountDir"
                    $rb = Run-App -Exe $dismExe -AppArgs @('/Mount-Image', "/ImageFile:$bootWimPath", '/Index:1', "/MountDir:$mountDir")
                    if ($DebugPreference -eq 'Continue') { $rb.Output | ForEach-Object { Write-Debug "  DISM> $_" } }
                    if ($rb.ExitCode -ne 0) { throw "DISM mount $($names.BootWim) failed for index $idx (exit $($rb.ExitCode))" }

                    foreach ($pkg in $ssuFiles) {
                        Write-Log "  Applying SSU to $($names.BootWim): $($pkg.Name)"
                        $rp = Run-App -Exe $dismExe -AppArgs @('/Add-Package', "/Image:$mountDir", "/PackagePath:$($pkg.FullName)")
                        if ($DebugPreference -eq 'Continue') { $rp.Output | ForEach-Object { Write-Debug "  DISM> $_" } }
                        if ($rp.ExitCode -ne 0) { Write-Warning "  SSU->boot failed: $($pkg.Name) (exit $($rp.ExitCode)); continuing" }
                    }

                    Write-Log "  Unmounting $($names.BootWim) index $idx (commit)..."
                    $rc = Run-App -Exe $dismExe -AppArgs @('/Unmount-Image', "/MountDir:$mountDir", '/Commit')
                    if ($DebugPreference -eq 'Continue') { $rc.Output | ForEach-Object { Write-Debug "  DISM> $_" } }
                    if ($rc.ExitCode -ne 0) { throw "DISM unmount $($names.BootWim) failed for index $idx (exit $($rc.ExitCode))" }

                    Write-JsonFile -Path $bootSvcJson -Data @{ Index = $idx; ServicedDate = (Get-Date -Format s) }
                    Write-Log "  $($names.BootWim) index $idx serviced"
                } catch {
                    Write-Log "  ERROR servicing $($names.BootWim) index $idx`: $_" 'ERROR'
                    if (Test-Path $mountDir) { Run-App -Exe $dismExe -AppArgs @('/Unmount-Image', "/MountDir:$mountDir", '/Discard') | Out-Null }
                    throw
                } finally {
                    if (Test-Path $mountDir) { Remove-Item $mountDir -Recurse -Force -ErrorAction SilentlyContinue }
                }
            }
        }
    }

    # -----------------------------------------------------------------------
    # Final assembly → WimsFinal
    # -----------------------------------------------------------------------
    Write-Log "Final assembly (compression: $CompressionType) -> $($paths.WimsFinal)..."
    Ensure-Folder -Path $paths.WimsFinal

    $compressionMap  = @{ 'None' = 'none'; 'Fast' = 'fast'; 'Maximum' = 'maximum' }
    $dismCompression = $compressionMap[$CompressionType]
    $finalMeta       = Read-JsonFile -Path $finalJson
    if (-not $finalMeta) { $finalMeta = @{} }

    # -- Assemble final install.wim --
    $finalInstallPath = Join-Path $paths.WimsFinal $names.InstallWim
    if ($finalMeta.InstallWimDate) {
        Write-Log "Final $($names.InstallWim) already assembled ($($finalMeta.InstallWimDate))"
    } else {
        $sortedInstall = @($extractedIndices | ForEach-Object { Join-Path $paths.WimsIndices ("{0}_{1}" -f $_, $names.InstallWim) } | Where-Object { Test-Path $_ })
        if ($sortedInstall.Count -eq 0) {
            Write-Warning "No $($names.InstallWim) files found to assemble"
        } else {
            Write-Log "Creating final $($names.InstallWim) from $(Split-Path $sortedInstall[0] -Leaf)..."
            $r0 = Run-App -Exe $dismExe -AppArgs @('/Export-Image', "/SourceImageFile:$($sortedInstall[0])", '/SourceIndex:1', "/DestinationImageFile:$finalInstallPath", "/Compress:$dismCompression")
            if ($DebugPreference -eq 'Continue') { $r0.Output | ForEach-Object { Write-Debug "DISM> $_" } }
            if ($r0.ExitCode -ne 0) { throw "DISM export (first $($names.InstallWim)) failed (exit $($r0.ExitCode))" }
            for ($i = 1; $i -lt $sortedInstall.Count; $i++) {
                $idxNum = $extractedIndices[$i]
                Write-Log "  Appending $($names.InstallWim) index $idxNum..."
                $ra = Run-App -Exe $dismExe -AppArgs @('/Export-Image', "/SourceImageFile:$($sortedInstall[$i])", '/SourceIndex:1', "/DestinationImageFile:$finalInstallPath")
                if ($DebugPreference -eq 'Continue') { $ra.Output | ForEach-Object { Write-Debug "DISM> $_" } }
                if ($ra.ExitCode -ne 0) { throw "DISM append $($names.InstallWim) index $idxNum failed (exit $($ra.ExitCode))" }
            }
            $finalMeta['InstallWimDate'] = (Get-Date -Format s)
            Write-JsonFile -Path $finalJson -Data $finalMeta
            Write-Log "Final $($names.InstallWim) assembled"
        }
    }

    # -- Assemble final boot.wim --
    $finalBootPath = Join-Path $paths.WimsFinal $names.BootWim
    if ($finalMeta.BootWimDate) {
        Write-Log "Final $($names.BootWim) already assembled ($($finalMeta.BootWimDate))"
    } else {
        $sortedBoot = @($extractedIndices | ForEach-Object { Join-Path $paths.WimsIndices ("{0}_{1}" -f $_, $names.BootWim) } | Where-Object { Test-Path $_ })
        if ($sortedBoot.Count -eq 0) {
            Write-Warning "No $($names.BootWim) files found to assemble"
        } else {
            Write-Log "Creating final $($names.BootWim) from $(Split-Path $sortedBoot[0] -Leaf)..."
            $rb0 = Run-App -Exe $dismExe -AppArgs @('/Export-Image', "/SourceImageFile:$($sortedBoot[0])", '/SourceIndex:1', "/DestinationImageFile:$finalBootPath", "/Compress:$dismCompression")
            if ($DebugPreference -eq 'Continue') { $rb0.Output | ForEach-Object { Write-Debug "DISM> $_" } }
            if ($rb0.ExitCode -ne 0) { throw "DISM export (first $($names.BootWim)) failed (exit $($rb0.ExitCode))" }
            for ($i = 1; $i -lt $sortedBoot.Count; $i++) {
                $idxNum = $extractedIndices[$i]
                Write-Log "  Appending $($names.BootWim) index $idxNum..."
                $rba = Run-App -Exe $dismExe -AppArgs @('/Export-Image', "/SourceImageFile:$($sortedBoot[$i])", '/SourceIndex:1', "/DestinationImageFile:$finalBootPath")
                if ($DebugPreference -eq 'Continue') { $rba.Output | ForEach-Object { Write-Debug "DISM> $_" } }
                if ($rba.ExitCode -ne 0) { throw "DISM append $($names.BootWim) index $idxNum failed (exit $($rba.ExitCode))" }
            }
            $finalMeta['BootWimDate'] = (Get-Date -Format s)
            Write-JsonFile -Path $finalJson -Data $finalMeta
            Write-Log "Final $($names.BootWim) assembled"
        }
    }

    Write-Log "Service workflow complete"
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
            Write-Log "[DryRun] Would remove: $RegistryRoot"
        } elseif (Test-Path $RegistryRoot) {
            Write-Log "Removing: $RegistryRoot"
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
            Write-Log "[DryRun] Would export ENTIRE key: $key -> $dest"
        } else {
            Write-Log "Export ENTIRE key: $key -> $dest"
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
            Write-Log "No specific values requested for $key; skipping."
            return
        }

        if ($DryRun) {
            Write-Log "[DryRun] Would export values [$($allValues -join ', ')] from $key -> $dest"
            return
        }

        Write-Log "Export specific values [$($allValues -join ', ')] from $key -> $dest"

        $query = reg.exe query "$key" /v * 2>$null
        if (-not $query) {
            Write-Log "WARNING: No data returned for $key"
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
                Write-Log "[DryRun] Would delete ENTIRE key: $key"
            } else {
                Write-Log "[DryRun] Would delete values [$($values -join ', ')] from $key"
            }
            return
        }

        Write-Log "Appending delete instructions for $key -> $dest"

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
                Write-Log "[DryRun] Would export ENTIRE key: $key -> $dest"
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
            Write-Log "[DryRun] Would remove: $path"
        } elseif (Test-Path $path) {
            Write-Log "Removing: $path"
            Remove-Item $path -Force
        }
        return
    }

    if ($DryRun) {
        Write-Log "[DryRun] Would write: $path"
    } else {
        Write-Log "Writing: $path"
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
            Write-Log "[DryRun] Would remove: $path"
        } elseif (Test-Path $path) {
            Write-Log "Removing: $path"
            Remove-Item $path -Force
        }
        return
    }

    if ($DryRun) {
        Write-Log "[DryRun] Would write: $path"
    } else {
        Write-Log "Writing: $path"
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
            Write-Log "[DryRun] Would remove: $path"
        } elseif (Test-Path $path) {
            Write-Log "Removing: $path"
            Remove-Item $path -Force
        }
        return
    }

    if ($DryRun) {
        Write-Log "[DryRun] Would write: $path"
    } else {
        Write-Log "Writing: $path"
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
            Write-Log "[DryRun] Would remove: $cleanPath"
            Write-Log "[DryRun] Would remove: $upgradePath"
        } else {
            if (Test-Path $cleanPath) {
                Write-Log "Removing: $cleanPath"
                Remove-Item $cleanPath -Force
            }
            if (Test-Path $upgradePath) {
                Write-Log "Removing: $upgradePath"
                Remove-Item $upgradePath -Force
            }
        }
        return
    }

    if ($DryRun) {
        Write-Log "[DryRun] Would write: $cleanPath"
        Write-Log "[DryRun] Would write: $upgradePath"
    } else {
        Write-Log "Writing: $cleanPath"
        Set-Content -LiteralPath $cleanPath   -Value $cleanTemplate   -Encoding ASCII
        Write-Log "Writing: $upgradePath"
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
            Write-Log "[DryRun] Would remove: $cleanPath"
            Write-Log "[DryRun] Would remove: $upgradePath"
        } else {
            if (Test-Path $cleanPath) {
                Write-Log "Removing: $cleanPath"
                Remove-Item $cleanPath -Force
            }
            if (Test-Path $upgradePath) {
                Write-Log "Removing: $upgradePath"
                Remove-Item $upgradePath -Force
            }
        }
        return
    }

    if ($DryRun) {
        Write-Log "[DryRun] Would write: $cleanPath"
        Write-Log "[DryRun] Would write: $upgradePath"
    } else {
        # Fill in the actual name for the files in the template
        $cleanContent   = $cleanTemplate -f $names.SetupConfigCleanIni
        $upgradeContent = $upgradeTemplate -f $names.SetupConfigUpgradeIni

        Write-Log "Writing: $cleanPath"
        Set-Content -LiteralPath $cleanPath   -Value $cleanContent   -Encoding ASCII
        Write-Log "Writing: $upgradePath"
        Set-Content -LiteralPath $upgradePath -Value $upgradeContent -Encoding ASCII
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
Write-Log "dism    : $dismExe"
if ($oscdimgExe) {
    Write-Log "oscdimg : $oscdimgExe"
} else {
    Write-Log "oscdimg : not found (ISO creation unavailable)"
}

# ==============================
# Core paths  (requires $Folder)
# ==============================
$paths = [ordered]@{}
$paths.BootWimInIso          = Join-Path $names.Sources $names.BootWim
$paths.InstallEsdInIso       = Join-Path $names.Sources $names.InstallEsd
$paths.InstallWimInIso       = Join-Path $names.Sources $names.InstallWim
$paths.SrcIsoRoot            = Join-Path $Folder $names.SrcIso
$paths.SrcIsoContent         = Join-Path $paths.SrcIsoRoot $names.Content
$paths.BIOSInSrc             = Join-Path $paths.SrcIsoContent $names.BootFileBIOS
$paths.UEFIInSrc             = Join-Path $paths.SrcIsoContent $names.BootFileUEFI
$paths.SourcesInSrc          = Join-Path $paths.SrcIsoContent $names.Sources
$paths.BootWimInSrc          = Join-Path $paths.SourcesInSrc $names.BootWim
$paths.InstallEsdInSrc       = Join-Path $paths.SourcesInSrc $names.InstallEsd
$paths.InstallWimInSrc       = Join-Path $paths.SourcesInSrc $names.InstallWim
$paths.DestIsoRoot           = Join-Path $Folder $names.DestIso
$paths.DestIsoContent        = Join-Path $paths.DestIsoRoot $names.Content
$paths.BIOSInDest            = Join-Path $paths.DestIsoContent $names.BootFileBIOS
$paths.UEFIInDest            = Join-Path $paths.DestIsoContent $names.BootFileUEFI
$paths.SourcesInDest         = Join-Path $paths.DestIsoContent $names.Sources
$paths.BootWimInDest         = Join-Path $paths.SourcesInDest $names.BootWim
$paths.InstallWimInDest      = Join-Path $paths.SourcesInDest $names.InstallWim
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

# ==============================
# Resolve source ISO
# ==============================
if (-not $ISO) {
    Write-Verbose "No -ISO specified; searching for *.iso in: $Folder"
    $isoFiles = @(Get-ChildItem -Path $Folder -Filter '*.iso' -File -ErrorAction SilentlyContinue)
    if ($isoFiles.Count -eq 0) {
        $needsISO = $Extract -or $Export -or (-not ($KB -or $Service -or $Drivers -or $Reg -or $Files -or $Prep -or $CreateISO))
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
        Write-Log "Auto-discovered ISO: $ISO"
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
    $DestISO = $ISO -replace '\.iso$', '.bundled.iso'
    Write-Verbose "Auto-derived DestISO: $DestISO"
}

# ==============================
# Read ISO / WIM metadata for WinOS / Version / Arch and index list
# Priority: 1) cached wim-metadata.json  2) SrcIsoContent  3) mount ISO
# ==============================
$allImages       = @()
$isoMetaResolved = $false

function Load-WimMetadata {
    param([string]$WimPath)
    $imgs = @(Get-WimImageList -WimPath $WimPath -DismExe $dismExe)
    if ($imgs.Count -gt 0) {
        $m = Get-ISOMetadataFromWim -WimPath $WimPath -DismExe $dismExe
        if (-not $WinOS)   { $script:WinOS   = $m.WinOS;   Write-Verbose "WinOS   auto: $($m.WinOS)" }
        if (-not $Version) { $script:Version = $m.Version; Write-Verbose "Version auto: $($m.Version)" }
        if (-not $Arch)    { $script:Arch    = $m.Arch;    Write-Verbose "Arch    auto: $($m.Arch)" }
        $script:isoMetaResolved = $true
        return $imgs
    }
    return @()
}

# 1) Prefer cached wim-metadata.json if it matches the current ISO
$metadataJson = Join-Path $paths.WimsIndices "wim-metadata.json"
$wimMeta      = Read-JsonFile -Path $metadataJson
if ($wimMeta -and ($wimMeta.ISOPath -eq $ISO -or -not $ISO)) {
    Write-Verbose "Loading index list from cached wim-metadata.json"
    $allImages = @($wimMeta.InstallImages | ForEach-Object { [PSCustomObject]@{ Index = [int]$_.Index; Name = $_.Name } })
    if (-not $WinOS)   { $WinOS   = $wimMeta.WinOS }
    if (-not $Version) { $Version = $wimMeta.Version }
    if (-not $Arch)    { $Arch    = $wimMeta.Arch }
    if ($allImages.Count -gt 0) { $isoMetaResolved = $true }
}

# 2) Fall back to reading from SrcIsoContent on disk
if (-not $isoMetaResolved -and (Test-Path $paths.SourcesInSrc)) {
    $existingWim = if (Test-Path $paths.InstallWimInSrc) { $paths.InstallWimInSrc }
                   elseif (Test-Path $paths.InstallEsdInSrc) { $paths.InstallEsdInSrc }
                   else { $null }
    if ($existingWim) {
        Write-Verbose "Reading metadata from SrcIsoContent: $existingWim"
        try { $allImages = @(Load-WimMetadata -WimPath $existingWim) } catch { Write-Warning "SrcISO read failed: $_" }
    }
}

# 3) Mount the ISO briefly if still needed
if (-not $isoMetaResolved -and $ISO -and (Test-Path $ISO) -and -not $DryRun -and -not $Clean) {
    Write-Verbose "Mounting ISO for metadata: $ISO"
    $metaDiskImg = Mount-DiskImage -ImagePath $ISO -PassThru -ErrorAction SilentlyContinue
    if ($metaDiskImg) {
        try {
            $metaDrive = ($metaDiskImg | Get-Volume).DriveLetter + ':\'
            $metaWim = if (Test-Path "$metaDrive$($names.Sources)\$($names.InstallWim)") {
                "$metaDrive$($names.Sources)\$($names.InstallWim)"
            } elseif (Test-Path "$metaDrive$($names.Sources)\$($names.InstallEsd)") {
                "$metaDrive$($names.Sources)\$($names.InstallEsd)"
            } else { $null }
            if ($metaWim) {
                Write-Verbose "Reading metadata from mounted ISO: $metaWim"
                try { $allImages = @(Load-WimMetadata -WimPath $metaWim) } catch { Write-Warning "ISO metadata read failed: $_" }
            }
        } finally { Dismount-DiskImage -ImagePath $ISO -ErrorAction SilentlyContinue | Out-Null }
    }
}

# Hard defaults
if (-not $WinOS)   { $WinOS   = '11' }
if (-not $Arch)    { $Arch    = 'x64' }
if (-not $Version) { $Version = if ($WinOS -eq '10') { '22H2' } else { '25H2' } }

# ==============================
# ShowIndices
# ==============================
if ($ShowIndices) {
    if ($allImages.Count -eq 0) {
        Write-Error "Cannot show indices: no metadata available. Run -Extract or -Export first, or provide -ISO."
        exit 1
    }
    $metaSrc = if ($wimMeta -and $allImages.Count -gt 0) { "cached (wim-metadata.json)" }
               elseif (Test-Path $paths.SourcesInSrc)    { "SrcIsoContent" }
               else                                       { "mounted ISO" }
    Write-Host "`nAvailable images in $($names.InstallWim) [$metaSrc]:`n"
    Write-Host ("{0,6}  {1}" -f 'Index', 'Name')
    Write-Host ("{0,6}  {1}" -f '------', '----')
    foreach ($img in $allImages) { Write-Host ("{0,6}  {1}" -f $img.Index, $img.Name) }
    Write-Host ""
    exit 0
}

# ==============================
# Determine work modes
# ==============================
$workSwitches = @()
if ($Extract)   { $workSwitches += 'Extract' }
if ($Export)    { $workSwitches += 'Export' }
if ($KB)        { $workSwitches += 'KB' }
if ($Service)   { $workSwitches += 'Service' }
if ($Drivers)   { $workSwitches += 'Drivers' }
if ($Reg)       { $workSwitches += 'Reg' }
if ($Files)     { $workSwitches += 'Files' }
if ($Prep)      { $workSwitches += 'Prep' }
if ($CreateISO) { $workSwitches += 'CreateISO' }

if ($All -or $Most -or (-not $workSwitches)) {
    $Extract  = $true
    $Export   = $true
    $KB       = $true
    $Service  = $true
    $Drivers  = $true
    $Reg      = $true
    $Files    = $true
    $Prep     = $true
    if ($Most) {
        $workSwitches = @('Most')
    } else {
        $CreateISO = $true
        $All       = $true
        $workSwitches = @('All')
    }
}

# ==============================
# UpdateISO: suppress Extract/Export/Service/KB unless indices given
# ==============================
if ($UpdateISO) {
    $hasExplicitIndices = $Home -or $Pro -or $Indices
    if (-not $hasExplicitIndices) {
        Write-Log "UpdateISO: no explicit index selection; skipping Extract / Export / KB / Service"
        $Extract = $false; $Export = $false; $KB = $false; $Service = $false
        $workSwitches = @($workSwitches | Where-Object { $_ -notin @('Extract','Export','KB','Service','All','Most') })
        if (-not $workSwitches) { $workSwitches = @('UpdateISO') }
    } else {
        Write-Log "UpdateISO: explicit indices provided; Extract / Export / Service will run"
    }
}

# ==============================
# Resolve index selection
# ==============================
$SelectedIndices = @()
if ($allImages.Count -gt 0) {
    $SelectedIndices = @(Resolve-IndexSelection -AllImages $allImages -SelectHome:$Home -SelectPro:$Pro -IndicesStr $Indices)
} else {
    Write-Verbose "Image list not yet available; index selection deferred until Invoke-Export"
}

Write-Log "Target profile: Windows $WinOS $Version $Arch"
Write-Log "Root folder   : $Folder"
Write-Log "ISO           : $(if ($ISO) { $ISO } else { '(none)' })"
Write-Log "DestISO       : $(if ($DestISO) { $DestISO } else { '(none)' })"
Write-Log "Selected idx  : $(if ($SelectedIndices.Count -gt 0) { $SelectedIndices.Index -join ', ' } else { 'all (deferred)' })"
Write-Log "Mode          : $($workSwitches -join ', ')"
if ($Clean)     { Write-Log "Clean mode    : Enabled" "WARN" }
if ($DryRun)    { Write-Log "Dry-run mode  : Enabled" "WARN" }
if ($UpdateISO) { Write-Log "UpdateISO     : Enabled" "WARN" }

if ($KB) { # Only KB workflow needs HTML parsing, so we delay this until now
    # --- HtmlAgilityPack bootstrap (PS 5.x SAFE) ---------------------------------
    $HtmlAgilityPackDll = 'HtmlAgilityPack.dll'
    $hapDll = Join-Path $Folder $HtmlAgilityPackDll

    if ($Clean) {
        if ($DryRun) {
            Write-Log "[DryRun] Would remove: $hapDll"
        } elseif (Test-Path $hapDLL) {
            Write-Log "Removing: $hapDLL"
            Remove-Item $hapDLL -Force
        }
    }
    elseif ($DryRun) {
        if (-not (Test-Path $hapDll)) {
            Write-Log "[DryRun] Would download: $HtmlAgilityPackDll"
        }   
    } else {
        if (-not (Test-Path $hapDll)) {
            Write-Log "HtmlAgilityPack.dll not found - downloading..."

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

if ($Extract)   { Invoke-ExtractISO }
if ($Export)    { Invoke-Export }
if ($KB)        { Invoke-KBWork }
if ($Service)   { Invoke-ServiceWork }
if ($Drivers)   { Invoke-DriverWork }
if ($Reg)       { Invoke-RegWork }
if ($Files) {
    Write-InstallDriversCmd
    Write-InstallRegsCmd
    Write-PostSetupCmd
    Write-SetupConfigFiles
    Write-SetupCmdFiles
}
if ($Prep)      { Invoke-PrepDestISO }
if ($CreateISO) { Invoke-CreateISO }

Write-Log "Completed"
