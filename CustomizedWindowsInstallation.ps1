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

.PARAMETER ExportWims
Export selected indices from SrcISO\Content\ into per-index uncompressed WIMs under Wims\Indices\.
Alias: -Export

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

    [Alias('ExtractISO')]
    [switch]$Extract,

    [Alias('ExportWims')]
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

    [Alias('Home')]
    [switch]$SelectHome,

    [Alias('Pro')]
    [switch]$SelectPro,

    [string]$Indices,

    [switch]$UseADK,

    [switch]$UseSystem,

    [string]$dism,

    [string]$oscdimg,

    [switch]$Clean,

    [switch]$DryRun,

    [switch]$Help
)

# git hash
$GitHash = "956c07d"

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

function Clean-File {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    if ($Clean) {
        if ($DryRun) {
            Write-Output "[DryRun] Would remove file  : $Path"
        } elseif (Test-Path $path) {
            Write-Output "Removing file  : $($Path)"
            Remove-Item $Path -Force -ErrorAction SilentlyContinue
        }
    }
}

function Clean-Folder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    if ($Clean) {
        if ($DryRun) {
            Write-Output "[DryRun] Would remove folder: $Path"
        } elseif (Test-Path $Path) {
            Write-Output "Removing folder: $($Path)"
            Remove-Item $Path -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Read-JsonFile {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path -PathType Leaf)) { return $null }
    try   { Get-Content $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop }
    catch { Write-Warning "Failed to read JSON '$Path': $_"; $null }
}

function Write-JsonFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][object]$Data
    )
    Ensure-Folder -Path (Split-Path $Path -Parent)
    $Data | ConvertTo-Json -Depth 6 | Set-Content -Path $Path -Encoding UTF8
}

function Run-App {
    param(
        [Parameter(Mandatory)]
        [string]$Exe,
        [Parameter(Mandatory)]
        [string[]]$Arguments
    )
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Exe
    $psi.Arguments = ($Arguments -join " ")
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    Write-Verbose "$($psi.FileName) $($psi.Arguments)"

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $allOutput = @()

    $proc.Start() | Out-Null

    $stdout = $proc.StandardOutput
    $stderr = $proc.StandardError

    # Track which lines have already been emitted to prevent stdout/stderr duplicates
    $seenLines       = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    # Track progress intervals to keep logs clean
    $lastLoggedValue = -1

    while (-not $proc.HasExited -or -not $stdout.EndOfStream -or -not $stderr.EndOfStream) {
        while (-not $stdout.EndOfStream) {
            $line = $stdout.ReadLine()
            # ONLY skip if the stream is literally sending nothing because it's idling
            if ($null -eq $line) { continue }

            # If it's a valid, intentional blank line printed by the app, let it pass through safely
            if ([string]::IsNullOrWhiteSpace($line)) {
                Write-Output ""
                $allOutput += ""
                continue
            }

            $isProgressLine = $false
            $logLine = $line

            # ----------------------------------------------------
            # DISM Engine
            # ----------------------------------------------------
            if ($Exe -match 'dism\.exe$') {
                if ($line -match '\[=+?\s+(\d+(?:\.\d+)?)%\s+=+?\]') {
                    $isProgressLine = $true
                    $percent = [math]::Round([double]$Matches[1])
                    if ($percent % 10 -eq 0 -and $percent -ne $lastLoggedValue) {
                        $logLine = "   Progress: $percent%"
                        $lastLoggedValue = $percent
                        $isProgressLine = $false # Allow this specific milestone line to be output
                    }
                }
            }
            # ----------------------------------------------------
            # Robocopy Engine
            # ----------------------------------------------------
            elseif ($Exe -match 'robocopy\.exe$') {
                if ($line -match '(\d+(?:\.\d+)?)%') {
                    $isProgressLine = $true
                    $percent = [math]::Round([double]$Matches[1])
                    if ($percent % 25 -eq 0 -and $percent -ne $lastLoggedValue) {
                        $logLine = "   Progress: $percent%"
                        $lastLoggedValue = $percent
                        $isProgressLine = $false
                    }
                }
            }
            # ----------------------------------------------------
            # Oscdimg Engine
            # ----------------------------------------------------
            elseif ($Exe -match 'oscdimg\.exe$') {
                if ($line -match '(\d+)%\s+complete') {
                    $isProgressLine = $true
                    $percent = [int]$Matches[1]
                    if ($percent % 10 -eq 0 -and $percent -ne $lastLoggedValue) {
                        $logLine = "   Progress: $percent%"
                        $lastLoggedValue = $percent
                        $isProgressLine = $false
                    }
                }
            }

            # Output and store the line if it isn't intermediate progress spam
            if (-not $isProgressLine) {
                $seenLines.Add($logLine) | Out-Null
                Write-Output $logLine
                $allOutput += $logLine
            }
        }

        while (-not $stderr.EndOfStream) {
            $line = $stderr.ReadLine()
            if ($null -ne $line -and -not [string]::IsNullOrWhiteSpace($line)) {
                # Only output stderr lines that were NOT already written from stdout
                if ($seenLines.Add($line)) {
                    Write-Output $line
                    $allOutput += $line
                }
            }
        }
        Start-Sleep -Milliseconds 50
    }
    $proc.WaitForExit()
    $rc = $proc.ExitCode
    $global:LASTEXITCODE = $rc
    return ,$allOutput
}

# ==============================
# Tool discovery
# ==============================

function Find-ADKTool {
    # Locate an ADK tool (e.g. dism.exe, oscdimg.exe) using this priority:
    #   1. Explicit path supplied by the caller
    #   2. Windows ADK installation (preferred when -PreferADK or -UseADK)
    #   3. System32 / PATH
    # Returns the full path on success, $null on failure.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ToolName,       # filename, e.g. 'dism.exe'

        [Parameter(Mandatory)]
        [string]$ADKSubfolder,   # subfolder under each arch dir, e.g. 'DISM' or 'Oscdimg'

        [string]$ExplicitPath,   # value of -dism / -oscdimg parameter
        [switch]$PreferADK,      # -UseADK
        [switch]$ForceSystem     # -UseSystem
    )

    Write-Debug "Find-ADKTool: '$ToolName' ExplicitPath='$ExplicitPath' PreferADK=$PreferADK ForceSystem=$ForceSystem"

    # 1. Explicit override
    if ($ExplicitPath) {
        if (-not (Test-Path $ExplicitPath)) {
            Write-Warning "Explicit path not found for $ToolName`: $ExplicitPath"
            return $null
        }
        Write-Verbose "Using explicit $ToolName`: $ExplicitPath"
        return $ExplicitPath
    }

    # Search ADK installations
    $adkRoots = @(
        "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools",
        "${env:ProgramFiles}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools"
    )
    $adkArches = @('amd64', 'arm64', 'x86')

    $adkPath = $null
    foreach ($root in $adkRoots) {
        foreach ($arch in $adkArches) {
            $candidate = Join-Path $root "$arch\$ADKSubfolder\$ToolName"
            if (Test-Path $candidate) { $adkPath = $candidate; break }
        }
        if ($adkPath) { break }
    }

    # System / PATH fallback: try System32 first, then PATH
    $systemPath = Join-Path $env:SystemRoot "System32\$ToolName"
    if (-not (Test-Path $systemPath)) {
        $cmd = Get-Command $ToolName -ErrorAction SilentlyContinue
        $systemPath = if ($cmd) { $cmd.Source } else { $null }
    }

    # 2. Apply priority
    if ($ForceSystem) {
        if ($systemPath) {
            Write-Verbose "Using system $ToolName (forced): $systemPath"
            return $systemPath
        }
        Write-Warning "$ToolName not found in System32 or PATH"
        return $null
    }

    if ($PreferADK -and $adkPath) {
        Write-Verbose "Using ADK $ToolName (preferred): $adkPath"
        return $adkPath
    }

    if ($adkPath) {
        Write-Verbose "Using ADK $ToolName (auto-discovered): $adkPath"
        return $adkPath
    }

    if ($systemPath) {
        Write-Verbose "Using system $ToolName (fallback): $systemPath"
        return $systemPath
    }

    Write-Warning "$ToolName not found. Install Windows ADK or specify the path explicitly."
    return $null
}

# ==============================
# ISO / WIM introspection
# ==============================

function Get-WimMetadata {
    # Reads a WIM file and returns both the full image list and OS details
    # (WinOS, Version, Arch, Build) in a single object, making two DISM calls:
    #   /Get-WimInfo            — to enumerate all images (Index + Name)
    #   /Get-WimInfo /Index:1   — to get Version and Architecture from index 1
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$WimPath
    )

    Write-Debug "Get-WimMetadata: WimPath='$WimPath'"

    # --- Call 1: enumerate all images ---
    $listOutput = Run-App $dismExe @('/Get-WimInfo', "/WimFile:$WimPath")
    $images     = [System.Collections.Generic.List[object]]::new()

    if ($LASTEXITCODE -eq 0) {
        $currentIdx  = $null
        $currentName = $null
        foreach ($line in $listOutput) {
            Write-Debug "  WimInfo> $line"
            if ($line -match '^\s*Index\s*:\s*(\d+)') {
                if ($null -ne $currentIdx) {
                    $images.Add([PSCustomObject]@{ Index = $currentIdx; Name = $currentName })
                }
                $currentIdx  = [int]$Matches[1]
                $currentName = ''
            } elseif ($null -ne $currentIdx -and $line -match '^\s*Name\s*:\s*(.+)') {
                $currentName = $Matches[1].Trim()
            }
        }
        if ($null -ne $currentIdx) {
            $images.Add([PSCustomObject]@{ Index = $currentIdx; Name = $currentName })
        }
        Write-Verbose "Found $($images.Count) image(s) in: $WimPath"
    } else {
        Write-Warning "DISM /Get-WimInfo failed (exit $LASTEXITCODE) for: $WimPath"
    }

    # --- Call 2: OS details from index 1 ---
    $buildNumber = 0
    $archStr     = 'x64'
    $detailOutput = Run-App $dismExe @('/Get-WimInfo', "/WimFile:$WimPath", '/Index:1')
    foreach ($line in $detailOutput) {
        if ($line -match '^\s*Version\s*:\s*\d+\.\d+\.(\d+)\.') { $buildNumber = [int]$Matches[1] }
        if ($line -match '^\s*Architecture\s*:\s*(.+)')         { $archStr     = $Matches[1].Trim() }
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
        default          { if ($detectedWinOS -eq '11') { '25H2' } else { '22H2' }; break }
    }

    $detectedArch = switch -Wildcard ($archStr.ToLower()) {
        '*arm64*' { 'arm64' }
        '*amd64*' { 'x64'   }
        '*x64*'   { 'x64'   }
        default   { 'x64'   }
    }

    return [PSCustomObject]@{
        Images  = $images.ToArray()
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
# Extract section
# =========================

function Invoke-ExtractISO {
    [CmdletBinding()]
    param()

    if ($Clean) {
        Clean-Folder $paths.SrcIsoRoot
        return
    }

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

    Write-Output "Starting ExtractISO workflow..."
    Write-Verbose "Invoke-ExtractISO: ISO='$ISO' SrcIsoContent='$($paths.SrcIsoContent)'"

    if (-not $ISO -or -not (Test-Path $ISO)) {
        Write-Warning "Source ISO not found or not specified. Use -ISO to point to your Windows .iso file."
        return
    }

    # Checkpoint: skip if same ISO was already extracted; clean and re-extract if ISO changed
    $extractJson  = Join-Path $paths.SrcIsoRoot "extract.json"
    $existingJson = Read-JsonFile -Path $extractJson
    if ($existingJson) {
        if ($existingJson.ISOPath -eq $ISO) {
            Write-Output "ExtractISO already done for this ISO (extract.json matches)"
            Write-Debug  "extract.json: ISOPath='$($existingJson.ISOPath)' Date='$($existingJson.Date)'"
            return
        }
        Write-Output "ISO path changed (was '$($existingJson.ISOPath)'); cleaning SrcIsoRoot and re-extracting..."
        Remove-Item $paths.SrcIsoRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    Ensure-Folder -Path $paths.SrcIsoContent

    Write-Output "Mounting ISO: $ISO"
    $diskImage = Mount-DiskImage -ImagePath $ISO -PassThru -ErrorAction Stop

    try {
        # Wait for the volume to actually appear
        $vol = $null
        for ($r = 0; $r -lt 5 -and $null -eq $vol; $r++) {
            $vol = $diskImage | Get-Volume -ErrorAction SilentlyContinue
            if ($null -eq $vol) { Start-Sleep -Seconds 1 }
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
        if (-not ((Test-Path (Join-Path $driveLetter $paths.InstallWimInIso)) -or
                  (Test-Path (Join-Path $driveLetter $paths.InstallEsdInIso)))) {
            $missing += "$($paths.InstallWimInIso) or $($paths.InstallEsdInIso)"
        }
        if ($missing.Count -gt 0) {
            throw "Source ISO validation failed. Missing required file(s): $($missing -join ', ')"
        }
        Write-Output "Source ISO validation passed"

        # Copy the ISO tree to SrcIsoContent
        Write-Output "Copying ISO tree -> $($paths.SrcIsoContent)..."
        $roboArgs = @(
            $driveLetterRaw,
            $paths.SrcIsoContent,
            '/E',    # copy subdirectories, including Empty ones
            '/R:2',  # retry twice
            '/W:1',  # wait 1 second between retries
            '/NC',   # No Class - don't show file classes (e.g., "New File", "Same File"), just show the file names
            '/TEE'   # output to console and log file
        )
        Run-App 'robocopy.exe' $roboArgs
        # Robocopy exit codes 0-7 are all 'Success' variants
        if ($LASTEXITCODE -ge 8) { throw "robocopy failed with exit code $LASTEXITCODE" }

    } catch {
        Write-Output "ERROR during ISO extraction: $_"
        # Clean up the partial SrcIsoContent so a re-run starts fresh
        if (Test-Path $paths.SrcIsoContent) {
            Remove-Item $paths.SrcIsoContent -Recurse -Force -ErrorAction SilentlyContinue
        }
        return
    } finally {
        if (Get-DiskImage -ImagePath $ISO | Where-Object { $_.Attached }) {
            Write-Output "Unmounting ISO..."
            Dismount-DiskImage -ImagePath $ISO -ErrorAction SilentlyContinue | Out-Null
        }
    }

    Write-JsonFile -Path $extractJson -Data @{ ISOPath = $ISO; Date = (Get-Date -Format s) }
    Write-Output "ExtractISO complete (extract.json written)"
}

# =========================
# Extract section
# =========================
function Invoke-Export {
    [CmdletBinding()]
    param()

    if ($Clean) {
        Clean-Folder $paths.WimsRoot
        return
    }

    if ($DryRun) {
        Write-Output "[DryRun] Would collect WIM metadata from SrcIsoContent"
        Write-Output "[DryRun] Would export $($SelectedIndices.Count) indices to $($paths.WimsIndices)"
        return
    }

    Write-Output "Starting Export workflow..."
    Write-Verbose "Invoke-Export: SourcesInSrc='$($paths.SourcesInSrc)' WimsIndices='$($paths.WimsIndices)'"

    $extractJson  = Join-Path $paths.SrcIsoRoot "extract.json"
    $metadataJson = Join-Path $paths.WimsIndices "wim-metadata.json"

    Ensure-Folder -Path $paths.WimsRoot
    Ensure-Folder -Path $paths.WimsIndices

    $extractMeta = Read-JsonFile -Path $extractJson
    if (-not $extractMeta) {
        Write-Warning "extract.json not found at '$extractJson'. Run -Extract first."
        return
    }
    $extractDate = [datetime]::Parse($extractMeta.Date)

    # Locate source WIMs
    $installSrc = if (Test-Path $paths.InstallWimInSrc) { $paths.InstallWimInSrc }
                  elseif (Test-Path $paths.InstallEsdInSrc) { $paths.InstallEsdInSrc }
                  else { $null }
    $bootSrc    = if (Test-Path $paths.BootWimInSrc) { $paths.BootWimInSrc } else { $null }

    if (-not $installSrc) {
        Write-Warning "Install image not found in $($paths.SourcesInSrc). Run -Extract first."
        return
    }
    if (-not $bootSrc) {
        Write-Warning "Boot image not found in $($paths.SourcesInSrc). Run -Extract first."
        return
    }

    Write-Verbose "install source: $installSrc"
    Write-Verbose "boot source   : $bootSrc"

    # Collect WIM metadata and write wim-metadata.json
    Write-Output "Collecting WIM metadata..."
    $installMeta = Get-WimMetadata -WimPath $installSrc
    $bootImages  = (Get-WimMetadata -WimPath $bootSrc).Images

    Write-JsonFile -Path $metadataJson -Data @{
        ISOPath       = $extractMeta.ISOPath
        CollectedDate = (Get-Date -Format s)
        WinOS         = $installMeta.WinOS
        Version       = $installMeta.Version
        Arch          = $installMeta.Arch
        Build         = $installMeta.Build
        InstallImages = @($installMeta.Images | ForEach-Object { @{ Index = $_.Index; Name = $_.Name } })
        BootImages    = @($bootImages          | ForEach-Object { @{ Index = $_.Index; Name = $_.Name } })
    }
    Write-Output "WIM metadata saved ($($installMeta.Images.Count) install image(s), $($bootImages.Count) boot image(s))"

    # Resolve index selection if not already set
    if ($SelectedIndices.Count -eq 0) {
        Write-Verbose "SelectedIndices empty; resolving from collected metadata..."
        $SelectedIndices = @(Resolve-IndexSelection -AllImages $installMeta.Images -SelectHome:$SelectHome -SelectPro:$SelectPro -IndicesStr $Indices)
        Write-Verbose "Resolved $($SelectedIndices.Count) index/indices"
    }

    Write-Output "Exporting $($SelectedIndices.Count) index/indices..."
    Write-Verbose "Selected: $($SelectedIndices.Index -join ', ')"

    $bootSrcIdx = if ($bootImages | Where-Object { $_.Index -eq 2 }) { 2 } else { 1 }

    foreach ($img in $SelectedIndices) {
        $idx     = $img.Index
        $imgName = $img.Name
        Write-Output "  [Index $idx] $imgName"

        # -- Export install image --
        $installDest = Join-Path $paths.WimsIndices ("{0}_{1}" -f $idx, $names.InstallWim)
        $installJson = "$installDest.json"
        $existInstall = Read-JsonFile -Path $installJson
        $needInstall  = (-not $existInstall) -or ([datetime]::Parse($existInstall.ExportDate) -le $extractDate)

        if (-not $needInstall) {
            Write-Output "    $($names.InstallWim) index $idx already exported ($($existInstall.ExportDate))"
        } else {
            Write-Output "    Exporting $($names.InstallWim) index $idx..."
            Run-App $dismExe @('/Export-Image', "/SourceImageFile:$installSrc", "/SourceIndex:$idx",
                               "/DestinationImageFile:$installDest", '/Compress:None')
            if ($LASTEXITCODE -ne 0) {
                Write-Warning "    DISM export failed for $($names.InstallWim) index $idx (exit $LASTEXITCODE) — skipping this index"
                continue
            }
            Write-JsonFile -Path $installJson -Data @{ Index = $idx; Name = $imgName; ExportDate = (Get-Date -Format s) }
            Write-Output "    $($names.InstallWim) index $idx exported"
        }

        # -- Export boot image --
        $bootDest = Join-Path $paths.WimsIndices ("{0}_{1}" -f $idx, $names.BootWim)
        $bootJson = "$bootDest.json"
        $existBoot = Read-JsonFile -Path $bootJson
        $needBoot  = (-not $existBoot) -or ([datetime]::Parse($existBoot.ExportDate) -le $extractDate)

        if (-not $needBoot) {
            Write-Output "    $($names.BootWim) index $idx already exported ($($existBoot.ExportDate))"
        } else {
            Write-Output "    Exporting $($names.BootWim) (src idx $bootSrcIdx) for index $idx..."
            Run-App $dismExe @('/Export-Image', "/SourceImageFile:$bootSrc", "/SourceIndex:$bootSrcIdx",
                               "/DestinationImageFile:$bootDest", '/Compress:None')
            if ($LASTEXITCODE -ne 0) {
                Write-Warning "    DISM export failed for $($names.BootWim) index $idx (exit $LASTEXITCODE) — skipping boot image for this index"
            } else {
                Write-JsonFile -Path $bootJson -Data @{ Index = $idx; SourceBootIndex = $bootSrcIdx; ExportDate = (Get-Date -Format s) }
                Write-Output "    $($names.BootWim) index $idx exported"
            }
        }
    }

    Write-Output "Export workflow complete"
}

# =========================
# KBs section
# =========================

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

    if ($Clean) {
        Clean-Folder $paths.KBsRoot
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
# Service section
# =========================

function Invoke-ServiceWork {
    [CmdletBinding()]
    param()

    if ($Clean) {
        Clean-Folder $paths.WimsMounts
        Clean-Folder $paths.WimsServiced
        return
    }

    if ($DryRun) {
        Write-Output "[DryRun] Would service extracted indices in $($paths.WimsIndices)"
        Write-Output "[DryRun] Would apply SSU packages from : $($paths.KBsSSU)"
        Write-Output "[DryRun] Would apply LCU packages from : $($paths.KBsOSCU)"
        Write-Output "[DryRun] Would service winre.wim inside each index's install.wim"
        Write-Output "[DryRun] Would assemble final install.wim -> $($paths.InstallWimInDest)"
        Write-Output "[DryRun] Would assemble final boot.wim   -> $($paths.BootWimInDest)"
        return
    }

    Write-Output  "Starting Service workflow..."
    Write-Verbose "Invoke-ServiceWork: WimsIndices='$($paths.WimsIndices)' WimsFinal='$($paths.WimsFinal)'"
    Write-Debug   "Invoke-ServiceWork: KBsSSU='$($paths.KBsSSU)' KBsOSCU='$($paths.KBsOSCU)' WimsMounts='$($paths.WimsMounts)'"

    $CompressionType = 'Maximum'  # None, Fast, Maximum

    Ensure-Folder -Path $paths.WimsMounts
    Ensure-Folder -Path $paths.WimsServiced

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

        $installDoneChkpt = Join-Path $paths.WimsIndices ("{0}_{1}.serviced.json" -f $idx, $names.InstallWim)
        $bootDoneChkpt    = Join-Path $paths.WimsIndices ("{0}_{1}.serviced.json" -f $idx, $names.BootWim)
        $winreDoneChkpt   = Join-Path $paths.WimsIndices ("{0}_{1}.serviced.json" -f $idx, $names.WinreWim)
        $winreExtChkpt    = Join-Path $paths.WimsIndices ("{0}_{1}.extracted.json" -f $idx, $names.WinreWim)

        $mountDir      = Join-Path $paths.WimsMounts ("mount_{0}"      -f $idx)
        $winreMountDir = Join-Path $paths.WimsMounts ("winremount_{0}" -f $idx)
        $winreWimPath  = Join-Path $paths.WimsIndices ("{0}_{1}"       -f $idx, $names.WinreWim)

        # ---- Service install.wim ----
        if (Read-JsonFile -Path $installDoneChkpt) {
            Write-Output "  $($names.InstallWim) index $idx already serviced"
        } else {
            if ($hasSSU -or $hasLCU) {
                Write-Output "  Servicing $($names.InstallWim) for index $idx..."
                Ensure-Folder -Path $mountDir

                try {
                    Write-Output  "  Mounting $installWimPath -> $mountDir"
                    Run-App $dismExe @("/Mount-Image /ImageFile:$installWimPath /Index:1 /MountDir:$mountDir")
                    if ($LASTEXITCODE -ne 0) { throw "DISM mount $($names.InstallWim) failed for index $idx (exit $LASTEXITCODE)" }

                    # ---- Service winre.wim inside install.wim ----
                    if ($hasSSU) {
                        $winreInMount = Join-Path $mountDir $paths.WinreWimInWim
                        Write-Debug "  Checking for $($names.WinreWim) at: $winreInMount"

                        if (Test-Path $winreInMount) {
                            if (Read-JsonFile -Path $winreDoneChkpt) {
                                Write-Output "  $($names.WinreWim) index $idx already serviced"
                            } else {
                                # Extract winre.wim
                                if (-not (Read-JsonFile -Path $winreExtChkpt)) {
                                    Write-Output "  Extracting $($names.WinreWim) from mounted install image..."
                                    Copy-Item -Path $winreInMount -Destination $winreWimPath -Force
                                    Write-JsonFile -Path $winreExtChkpt -Data @{ Index = $idx; ExtractedDate = (Get-Date -Format s) }
                                    Write-Output "  $($names.WinreWim) extracted"
                                } else {
                                    Write-Output "  $($names.WinreWim) already extracted"
                                }

                                # Mount winre.wim
                                Ensure-Folder -Path $winreMountDir
                                Write-Output  "  Mounting $($names.WinreWim) -> $winreMountDir"
                                Run-App $dismExe @("/Mount-Image /ImageFile:$winreWimPath /Index:1 /MountDir:$winreMountDir")
                                if ($LASTEXITCODE -ne 0) { throw "DISM mount $($names.WinreWim) failed for index $idx (exit $LASTEXITCODE)" }

                                # Apply SSU packages to winre
                                foreach ($pkg in $ssuFiles) {
                                    Write-Output  "  Applying SSU to $($names.WinreWim): $($pkg.Name)"
                                    Run-App $dismExe @("/Add-Package /Image:$winreMountDir /PackagePath:$($pkg.FullName)")
                                    if ($LASTEXITCODE -ne 0) {
                                        Write-Warning "  DISM SSU->$($names.WinreWim) failed: $($pkg.Name) index $idx (exit $LASTEXITCODE); continuing"
                                    }
                                }

                                # Unmount and commit winre
                                Write-Output "  Unmounting $($names.WinreWim) (commit)..."
                                Run-App $dismExe @("/Unmount-Image /MountDir:$winreMountDir /Commit")
                                if ($LASTEXITCODE -ne 0) { throw "DISM unmount $($names.WinreWim) failed for index $idx (exit $LASTEXITCODE)" }

                                # Reinsert serviced winre.wim back into mounted install.wim
                                Write-Output "  Reinserting serviced $($names.WinreWim) into install image..."
                                Copy-Item -Path $winreWimPath -Destination $winreInMount -Force
                                Write-JsonFile -Path $winreDoneChkpt -Data @{ Index = $idx; ServicedDate = (Get-Date -Format s) }
                                Write-Output "  $($names.WinreWim) serviced and reinserted"
                            }
                        } else {
                            Write-Verbose "  $($names.WinreWim) not found in mounted install image at $winreInMount; skipping winre servicing"
                        }

                        # Apply SSU packages to install.wim
                        foreach ($pkg in $ssuFiles) {
                            Write-Output  "  Applying SSU to $($names.InstallWim): $($pkg.Name)"
                            Run-App $dismExe @("/Add-Package /Image:$mountDir /PackagePath:$($pkg.FullName)")
                            if ($LASTEXITCODE -ne 0) {
                                Write-Warning "  DISM SSU->$($names.InstallWim) failed: $($pkg.Name) index $idx (exit $LASTEXITCODE); continuing"
                            }
                        }
                    }

                    # Apply LCU (OSCU) packages to install.wim
                    if ($hasLCU) {
                        foreach ($pkg in $lcuFiles) {
                            Write-Output  "  Applying LCU to $($names.InstallWim): $($pkg.Name)"
                            Run-App $dismExe @("/Add-Package /Image:$mountDir /PackagePath:$($pkg.FullName)")
                            if ($LASTEXITCODE -ne 0) {
                                Write-Warning "  DISM LCU->$($names.InstallWim) failed: $($pkg.Name) index $idx (exit $LASTEXITCODE); continuing"
                            }
                        }
                    }

                    # Unmount and commit install.wim
                    Write-Output "  Unmounting $($names.InstallWim) index $idx (commit)..."
                    Run-App $dismExe @("/Unmount-Image /MountDir:$mountDir /Commit")
                    if ($LASTEXITCODE -ne 0) { throw "DISM unmount $($names.InstallWim) failed for index $idx (exit $LASTEXITCODE)" }

                    Write-JsonFile -Path $installDoneChkpt -Data @{ Index = $idx; ServicedDate = (Get-Date -Format s) }
                    Write-Output "  $($names.InstallWim) index $idx serviced"

                } catch {
                    Write-Output "  ERROR servicing $($names.InstallWim) index $idx`: $_"
                    if (Test-Path $mountDir) {
                        Write-Output "  Discarding mounted $($names.InstallWim)..."
                        Run-App $dismExe @("/Unmount-Image /MountDir:$mountDir /Discard") | Out-Null
                    }
                    if (Test-Path $winreMountDir) {
                        Write-Output "  Discarding mounted $($names.WinreWim)..."
                        Run-App $dismExe @("/Unmount-Image /MountDir:$winreMountDir /Discard") | Out-Null
                    }
                    throw
                } finally {
                    if (Test-Path $mountDir)      { Remove-Item $mountDir      -Recurse -Force -ErrorAction SilentlyContinue }
                    if (Test-Path $winreMountDir) { Remove-Item $winreMountDir -Recurse -Force -ErrorAction SilentlyContinue }
                }
            } else {
                Write-Output "  No SSU/LCU packages; marking $($names.InstallWim) index $idx as done"
                Write-JsonFile -Path $installDoneChkpt -Data @{ Index = $idx; ServicedDate = (Get-Date -Format s) }
            }
        }

        # ---- Service boot.wim ----
        if (-not (Test-Path $bootWimPath)) {
            Write-Verbose "  No $($names.BootWim) for index $idx; skipping boot servicing"
        } elseif (Read-JsonFile -Path $bootDoneChkpt) {
            Write-Output "  $($names.BootWim) index $idx already serviced"
        } elseif (-not $hasSSU) {
            Write-Output "  No SSU packages; marking $($names.BootWim) index $idx as done"
            Write-JsonFile -Path $bootDoneChkpt -Data @{ Index = $idx; ServicedDate = (Get-Date -Format s) }
        } else {
            Write-Output "  Servicing $($names.BootWim) for index $idx..."
            Ensure-Folder -Path $mountDir

            try {
                Write-Output  "  Mounting $bootWimPath -> $mountDir"
                Run-App $dismExe @("/Mount-Image", "/ImageFile:$bootWimPath", "/Index:1", "/MountDir:$mountDir")
                if ($LASTEXITCODE -ne 0) { throw "DISM mount $($names.BootWim) failed for index $idx (exit $LASTEXITCODE)" }

                foreach ($pkg in $ssuFiles) {
                    Write-Output  "  Applying SSU to $($names.BootWim): $($pkg.Name)"
                    Run-App $dismExe @("/Add-Package", "/Image:$mountDir", "/PackagePath:$($pkg.FullName)")
                    if ($LASTEXITCODE -ne 0) {
                        Write-Warning "  DISM SSU->$($names.BootWim) failed: $($pkg.Name) index $idx (exit $LASTEXITCODE); continuing"
                    }
                }

                Write-Output "  Unmounting $($names.BootWim) index $idx (commit)..."
                Run-App $dismExe @("/Unmount-Image /MountDir:$mountDir /Commit")
                if ($LASTEXITCODE -ne 0) { throw "DISM unmount $($names.BootWim) failed for index $idx (exit $LASTEXITCODE)" }

                Write-JsonFile -Path $bootDoneChkpt -Data @{ Index = $idx; ServicedDate = (Get-Date -Format s) }
                Write-Output "  $($names.BootWim) index $idx serviced"

            } catch {
                Write-Output "  ERROR servicing $($names.BootWim) index $idx`: $_"
                if (Test-Path $mountDir) {
                    Run-App $dismExe @("/Unmount-Image /MountDir:$mountDir /Discard") | Out-Null
                }
                throw
            } finally {
                if (Test-Path $mountDir) { Remove-Item $mountDir -Recurse -Force -ErrorAction SilentlyContinue }
            }
        }
    }

    # -----------------------------------------------------------------------
    # Final assembly (serial compression can be slow)
    # -----------------------------------------------------------------------
    Write-Output "Final assembly: combining serviced indices (compression: $CompressionType) -> $($paths.WimsFinal)..."
    Ensure-Folder -Path $paths.WimsFinal

    $compressionMap  = @{ 'None' = 'none'; 'Fast' = 'fast'; 'Maximum' = 'maximum' }
    $dismCompression = $compressionMap[$CompressionType]
    $finalJson    = Join-Path $paths.WimsFinal "final.json"
    $finalMeta    = Read-JsonFile -Path $finalJson
    if (-not $finalMeta) { $finalMeta = @{} }

    # Assemble one WIM (install or boot) by exporting all per-index source files
    # into a single compressed destination.  DISM creates the file on the first
    # call and appends on subsequent ones — the arguments are identical either
    # way, so a single loop starting at 0 handles both cases.
    function Invoke-AssembleWim {
        param(
            [string]   $WimLabel,    # display name for messages
            [string]   $DestPath,    # final output file
            [string[]] $Sources,     # per-index source WIM files (sorted)
            [int[]]    $Indices,     # corresponding index numbers (for messages)
            [string]   $Compression, # dism compress value (none/fast/maximum)
            [string]   $DateKey      # key to stamp in $finalMeta on success
        )

        if ($Sources.Count -eq 0) { Write-Warning "No $WimLabel files found to assemble"; return }

        # Individual WIMs always have a SourceIndex of 1 regardless of their original index
        # You must have the compression or dism will corrupt the file
        $baseArgs = @('/Export-Image', "/DestinationImageFile:$DestPath", '/SourceIndex:1', "/Compress:$Compression")
        try {
            Write-Output "Assembling final $WimLabel ($($Sources.Count) index/indices)..."
            for ($i = 0; $i -lt $Sources.Count; $i++) {
                Write-Output "  Index $($Indices[$i])..."
                Run-App $dismExe ($baseArgs + @("/SourceImageFile:$($Sources[$i])"))
                if ($LASTEXITCODE -ne 0) { throw "DISM failed on $WimLabel index $($Indices[$i]) (exit $LASTEXITCODE)" }
            }
            $finalMeta[$DateKey] = (Get-Date -Format s)
            Write-JsonFile -Path $finalJson -Data $finalMeta
            Write-Output "Final $WimLabel assembled"
        } catch {
            Write-Output "ERROR assembling final $WimLabel`: $_"
            if (Test-Path $DestPath) { Remove-Item $DestPath -Force -ErrorAction SilentlyContinue }
        }
    }

    # -- Assemble final install.wim --
    $finalInstallPath = Join-Path $paths.WimsFinal $names.InstallWim
    if ($finalMeta.InstallWimDate) {
        Write-Output "Final $($names.InstallWim) already assembled ($($finalMeta.InstallWimDate))"
    } else {
        $sortedInstallWims = @($extractedIndices |
            ForEach-Object { Join-Path $paths.WimsIndices ("{0}_{1}" -f $_, $names.InstallWim) } |
            Where-Object   { Test-Path $_ })
        Invoke-AssembleWim -WimLabel $names.InstallWim -DestPath $finalInstallPath `
                           -Sources $sortedInstallWims -Indices $extractedIndices `
                           -Compression $dismCompression -DateKey 'InstallWimDate'
    }

    # -- Assemble final boot.wim --
    $finalBootPath = Join-Path $paths.WimsFinal $names.BootWim
    if ($finalMeta.BootWimDate) {
        Write-Output "Final $($names.BootWim) already assembled ($($finalMeta.BootWimDate))"
    } else {
        $sortedBootWims = @($extractedIndices |
            ForEach-Object { Join-Path $paths.WimsIndices ("{0}_{1}" -f $_, $names.BootWim) } |
            Where-Object   { Test-Path $_ })
        Invoke-AssembleWim -WimLabel $names.BootWim -DestPath $finalBootPath `
                           -Sources $sortedBootWims -Indices $extractedIndices `
                           -Compression $dismCompression -DateKey 'BootWimDate'
    }

    Write-Output "Service workflow complete"
}

# ==============================
# Driver export
# ==============================
function Invoke-DriverWork {
    $WinpeDriverRoot = $paths.WinpeDriverRoot

    if ($Clean) {
        Clean-Folder $WinpeDriverRoot
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
    if ($p.ExitCode -ne 0) {
        Write-Output "ERROR: DISM /export-driver failed (exit $($p.ExitCode)), check $WinpeDriverRoot\dism.log"
    }
}

# ==============================
# Registry export
# ==============================
function Invoke-RegWork {

    # An empty list in the Values means work on the entire key
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

    if ($Clean) {
        Clean-Folder $RegistryRoot
        return
    }

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
        Clean-File $path
    } elseif ($DryRun) {
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
        Clean-File $path
    } elseif ($DryRun) {
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
        Clean-File $path
    } elseif ($DryRun) {
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
        Clean-File $cleanPath
        Clean-File $upgradePath
    } elseif ($DryRun) {
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
        Clean-File $cleanPath
        Clean-File $upgradePath
    } elseif ($DryRun) {
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
# Prep for the Destination ISO
# ==============================

function Invoke-PrepDestISO {
    [CmdletBinding()]
    param()

    if ($Clean) {
        Clean-Folder $paths.DestIsoRoot
        return
    }

    if ($DryRun) {
        Write-Output "[DryRun] Would hardlink-copy $($paths.SrcIsoContent) -> $($paths.DestIsoContent)"
        Write-Output "[DryRun] Would copy $finalInstall -> $($paths.InstallWimInDest)"
        Write-Output "[DryRun] Would copy $finalBoot    -> $($paths.BootWimInDest)"
        return
    }

    Write-Output "Starting PrepDestISO workflow..."
    Write-Verbose "Invoke-PrepDestISO: SrcIsoContent='$($paths.SrcIsoContent)' DestIsoContent='$($paths.DestIsoContent)'"
    $prepJson   = Join-Path $paths.DestIsoRoot "prep.json"
    $extractJson = Join-Path $paths.SrcIsoRoot "extract.json"
    $finalJson   = Join-Path $paths.WimsFinal "final.json"
    $finalInstall = Join-Path $paths.WimsFinal $names.InstallWim
    $finalBoot    = Join-Path $paths.WimsFinal $names.BootWim

    $extractMeta = Read-JsonFile -Path $extractJson
    if (-not $extractMeta) {
        Write-Warning "extract.json not found. Run -Extract first."
        return
    }
    $extractDate = [datetime]::Parse($extractMeta.Date)

    $finalMeta = Read-JsonFile -Path $finalJson
    $prep      = Read-JsonFile -Path $prepJson

    # ---- Step A: Hardlink-copy SrcIsoContent -> DestIsoContent ----
    $needHardlink = (-not $prep -or -not $prep.HardlinkDate) -or
                    ([datetime]::Parse($prep.HardlinkDate) -le $extractDate)

    if ($needHardlink) {
        Write-Output "Hardlink-copying $($paths.SrcIsoContent) -> $($paths.DestIsoContent) (excluding install/boot images)..."
        if (Test-Path $paths.DestIsoRoot) {
            Write-Output "Cleaning existing DestIsoRoot..."
            Remove-Item $paths.DestIsoRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        Ensure-Folder -Path $paths.DestIsoContent

        $excludeNames = @($names.BootWim, $names.InstallWim, $names.InstallEsd)
        $allFiles = @(Get-ChildItem -Path $paths.SrcIsoContent -Recurse -File -ErrorAction SilentlyContinue)
        $total = $allFiles.Count; $done = 0; $lastPct = -1

        Write-Output "  Hardlinking $total files..."
        foreach ($file in $allFiles) {
            $done++
            $pct = [math]::Floor(($done / [math]::Max($total, 1)) * 100)
            if ($pct -ge ($lastPct + 10)) {
                Write-Output ("  {0,3}%  {1}/{2} files" -f $pct, $done, $total)
                $lastPct = $pct - ($pct % 10)
            }
            if ($file.Name -in $excludeNames) { continue }
            $rel  = $file.FullName.Substring($paths.SrcIsoContent.TrimEnd('\').Length).TrimStart('\')
            $dest = Join-Path $paths.DestIsoContent $rel
            Ensure-Folder -Path (Split-Path $dest -Parent)
            if (-not (Test-Path $dest)) {
                try {
                    New-Item -ItemType HardLink -Path $dest -Value $file.FullName -Force -ErrorAction Stop | Out-Null
                } catch {
                    Write-Warning "Hardlink failed for '$rel'; copying: $_"
                    Copy-Item -Path $file.FullName -Destination $dest -Force
                }
            }
        }
        Write-Output "  Hardlink tree complete"
        $prep = @{ HardlinkDate = (Get-Date -Format s) }
        Write-JsonFile -Path $prepJson -Data $prep
    } else {
        Write-Output "DestIsoContent hardlink-copy already current (prep.json: $($prep.HardlinkDate))"
    }

    $prep = Read-JsonFile -Path $prepJson
    if (-not $prep) { $prep = @{} }

    # ---- Step B: Copy final install.wim ----
    $finalInstDate = if ($finalMeta -and $finalMeta.InstallWimDate) { [datetime]::Parse($finalMeta.InstallWimDate) } else { [datetime]::MinValue }
    $destInstDate  = if ($prep.InstallWimDate) { [datetime]::Parse($prep.InstallWimDate) } else { [datetime]::MinValue }

    if ((Test-Path $finalInstall) -and ($destInstDate -le $finalInstDate)) {
        Write-Output "Copying $($names.InstallWim) -> $($paths.InstallWimInDest)..."
        Ensure-Folder -Path (Split-Path $paths.InstallWimInDest -Parent)
        Copy-Item -Path $finalInstall -Destination $paths.InstallWimInDest -Force
        $prep['InstallWimDate'] = (Get-Date -Format s)
        Write-JsonFile -Path $prepJson -Data $prep
    } elseif (Test-Path $paths.InstallWimInDest) {
        Write-Output "$($names.InstallWim) already current (prep.json: $($prep.InstallWimDate))"
    } else {
        Write-Warning "Final $($names.InstallWim) not found at: $finalInstall (run -Service first)"
    }

    $prep = Read-JsonFile -Path $prepJson
    if (-not $prep) { $prep = @{} }

    # ---- Step C: Copy final boot.wim ----
    $finalBootDate = if ($finalMeta -and $finalMeta.BootWimDate) { [datetime]::Parse($finalMeta.BootWimDate) } else { [datetime]::MinValue }
    $destBootDate  = if ($prep.BootWimDate) { [datetime]::Parse($prep.BootWimDate) } else { [datetime]::MinValue }

    if ((Test-Path $finalBoot) -and ($destBootDate -le $finalBootDate)) {
        Write-Output "Copying $($names.BootWim) -> $($paths.BootWimInDest)..."
        Ensure-Folder -Path (Split-Path $paths.BootWimInDest -Parent)
        Copy-Item -Path $finalBoot -Destination $paths.BootWimInDest -Force
        $prep['BootWimDate'] = (Get-Date -Format s)
        Write-JsonFile -Path $prepJson -Data $prep
    } elseif (Test-Path $paths.BootWimInDest) {
        Write-Output "$($names.BootWim) already current (prep.json: $($prep.BootWimDate))"
    } else {
        Write-Warning "Final $($names.BootWim) not found at: $finalBoot (run -Service first)"
    }

    Write-Output "PrepDestISO workflow complete"
}

# ==============================
# Create the Destination ISO
# ==============================
function Invoke-CreateISOWork {
    if ($Clean) {
        Clean-File $DestISO
        return
    }

    if ($DryRun) {
        Write-Output "[DryRun] Would create ISO: $DestISO"
        Write-Output "[DryRun]   from: $($paths.DestIsoContent)"
        return
    }

    Write-Output "Starting CreateISO workflow..."
    Write-Verbose "Invoke-CreateISOWork: DestIsoContent='$($paths.DestIsoContent)' DestISO='$DestISO'"
    $prepJson = Join-Path $paths.DestIsoRoot "prep.json"

    # Depends on prep.json — fail if DestISO content has not been prepared
    $prepMeta = Read-JsonFile -Path $prepJson
    if (-not $prepMeta) {
        Write-Warning "prep.json not found at '$prepJson'. Run -Prep first to prepare the destination ISO content."
        return
    }

    if (-not $oscdimgExe) {
        Write-Warning "oscdimg.exe not found. Install Windows ADK or specify -oscdimg."
        return
    }
    if (-not $DestISO) {
        Write-Warning "DestISO path is not set. Specify -DestISO or ensure -ISO is provided."
        return
    }

    # Sanity check for boot files before invoking oscdimg
    $etfs = $paths.BIOSInDest
    $efis = $paths.UEFIInDest
    if (-not (Test-Path $etfs)) {
        Write-Warning "Missing BIOS boot file: $etfs"
        return
    }
    if (-not (Test-Path $efis)) {
        Write-Warning "Missing UEFI boot file: $efis"
        return
    }

    $IsoVolumeLabel = "Win$($WinOS)_$($Version)_$($Arch)_KBs"
    $bootdata = "2#p0,e,b$etfs#pEF,e,b$efis"

    $oscdimgArgs = @(
        "-m",                    # Ignore size limits
        "-o",                    # Optimize storage by encoding duplicate files only once
        "-u2",                   # Use UTF-8 encoding for file names (allows for long file names and Unicode characters)
        "-udfver102",            # Use UDF 1.02 filesystem version (max compatibility, required for some boot scenarios)
        "-l$IsoVolumeLabel",     # Set volume label
        "-bootdata:$bootdata",   # Define multi-boot configuration for BIOS and UEFI
        $paths.DestIsoContent,
        $DestISO
    )

    Write-Output "Building ISO: $DestISO"
    Run-App $oscdimgExe $oscdimgArgs
    if ($LASTEXITCODE -ne 0) {
        Write-Output "ERROR: oscdimg failed to build ISO (exit $LASTEXITCODE)"
        return
    }
    Write-Output "Created ISO: $DestISO"
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
$dismExe    = Find-ADKTool -ToolName 'dism.exe'    -ADKSubfolder 'DISM'    -ExplicitPath $dism    -PreferADK:$UseADK -ForceSystem:$UseSystem
$oscdimgExe = Find-ADKTool -ToolName 'oscdimg.exe' -ADKSubfolder 'Oscdimg' -ExplicitPath $oscdimg -PreferADK:$UseADK -ForceSystem:$UseSystem

if (-not $dismExe) {
    Write-Output "ERROR: dism.exe is required but was not found."
    Write-Output "       Install the Windows ADK or use -dism to specify its path."
    exit 1
}
Write-Output "dism    : $dismExe"
if ($oscdimgExe) {
    Write-Output "oscdimg : $oscdimgExe"
} else {
    Write-Output "oscdimg : not found (ISO creation unavailable; -CreateISO will fail)"
}

# ==============================
# Core paths (requires $Folder)
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

if (-not $DryRun -and -not $Clean) {
    # ==============================
    # Read ISO / WIM metadata for WinOS / Version / Arch and index list
    # Priority: 1) wim-metadata.json  2) SrcIsoContent WIMs  3) mount ISO
    # ==============================
    $allImages       = @()
    $isoMetaResolved = $false
    $metaSrc         = $null

    # Helper: apply Get-WimMetadata result to the script-scope variables
    function Apply-WimMetadata {
        param([object]$Meta)
        $script:allImages = $Meta.Images
        if (-not $script:WinOS)   { $script:WinOS   = $Meta.WinOS }
        if (-not $script:Version) { $script:Version = $Meta.Version }
        if (-not $script:Arch)    { $script:Arch    = $Meta.Arch }
    }

    # 1) Prefer cached wim-metadata.json if it matches the current ISO
    $metadataJson = Join-Path $paths.WimsIndices "wim-metadata.json"
    $wimMeta      = Read-JsonFile -Path $metadataJson
    if ($wimMeta -and ($wimMeta.ISOPath -eq $ISO -or -not $ISO)) {
        Write-Verbose "Loading metadata from wim-metadata.json"
        $allImages = @($wimMeta.InstallImages | ForEach-Object { [PSCustomObject]@{ Index = [int]$_.Index; Name = $_.Name } })
        if (-not $WinOS)   { $WinOS   = $wimMeta.WinOS }
        if (-not $Version) { $Version = $wimMeta.Version }
        if (-not $Arch)    { $Arch    = $wimMeta.Arch }
        if ($allImages.Count -gt 0) { $isoMetaResolved = $true; $metaSrc = "wim-metadata.json" }
    }

    # 2) Fall back to SrcIsoContent on disk
    if (-not $isoMetaResolved -and (Test-Path $paths.SourcesInSrc)) {
        $existingWim = if (Test-Path $paths.InstallWimInSrc) { $paths.InstallWimInSrc }
                       elseif (Test-Path $paths.InstallEsdInSrc) { $paths.InstallEsdInSrc }
                       else { $null }
        if ($existingWim) {
            Write-Verbose "Reading metadata from SrcIsoContent: $existingWim"
            try {
                Apply-WimMetadata (Get-WimMetadata -WimPath $existingWim)
                if ($allImages.Count -gt 0) { $isoMetaResolved = $true; $metaSrc = "SrcIsoContent" }
            } catch { Write-Warning "Could not read metadata from SrcIsoContent: $_" }
        }
    }

    # 3) Mount the ISO briefly if still needed
    if (-not $isoMetaResolved -and $ISO -and (Test-Path $ISO)) {
        Write-Verbose "Mounting ISO for metadata: $ISO"
        $metaDiskImg = Mount-DiskImage -ImagePath $ISO -PassThru -ErrorAction SilentlyContinue
        if ($metaDiskImg) {
            try {
                $metaDrive = ($metaDiskImg | Get-Volume).DriveLetter + ':\'
                $metaWim   = if (Test-Path "$metaDrive$($names.Sources)\$($names.InstallWim)") {
                                 "$metaDrive$($names.Sources)\$($names.InstallWim)"
                             } elseif (Test-Path "$metaDrive$($names.Sources)\$($names.InstallEsd)") {
                                 "$metaDrive$($names.Sources)\$($names.InstallEsd)"
                             } else { $null }
                if ($metaWim) {
                    Write-Verbose "Reading metadata from mounted ISO"
                    Apply-WimMetadata (Get-WimMetadata -WimPath $metaWim)
                    if ($allImages.Count -gt 0) { $isoMetaResolved = $true; $metaSrc = "mounted ISO" }
                }
            } finally { Dismount-DiskImage -ImagePath $ISO -ErrorAction SilentlyContinue | Out-Null }
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
            Write-Error "Cannot show indices: no metadata available. Run -Extract or -Export first, or provide -ISO."
            exit 1
        }
        Write-Output "`nAvailable images in $($names.InstallWim) [source: $metaSrc]:`n"
        Write-Output ("{0,6}  {1}" -f 'Index', 'Name')
        Write-Output ("{0,6}  {1}" -f '------', '----')
        foreach ($img in $allImages) { Write-Output ("{0,6}  {1}" -f $img.Index, $img.Name) }
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
        Write-Verbose "Image list unavailable yet; index selection deferred until -Export"
        $SelectedIndices = @()
    }
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
    $Extract = $true
    $Export  = $true
    $KB      = $true
    $Service = $true
    $Drivers = $true
    $Reg     = $true
    $Files   = $true
    $Prep    = $true
    if ($Most) {
        $workSwitches = @('Most')
    } else {
        $CreateISO = $true
        $All       = $true
        $workSwitches = @('All')
    }
}

Write-Output "Target profile : Windows $WinOS $Version $Arch"
Write-Output "Root folder    : $Folder"
Write-Output "ISO            : $(if ($ISO) { $ISO } else { '(none)' })"
Write-Output "DestISO        : $(if ($DestISO) { $DestISO } else { '(none)' })"
Write-Output "Selected idx   : $(if ($SelectedIndices.Count -gt 0) { $SelectedIndices.Index -join ', ' } else { 'all (determined at export time)' })"
Write-Output "Mode           : $($workSwitches -join ', ')"
if ($Clean)  { Write-Output "Clean mode     : Enabled" }
if ($DryRun) { Write-Output "Dry-run mode   : Enabled" }

if ($KB) { # Only KB workflow needs HTML parsing, so we delay this until now
    # --- HtmlAgilityPack bootstrap (PS 5.x SAFE) ---------------------------------
    $HtmlAgilityPackDll = 'HtmlAgilityPack.dll'
    $hapDll = Join-Path $Folder $HtmlAgilityPackDll

    if ($Clean) {
        Clean-File $hapDll
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
                Write-Output "ERROR: $HtmlAgilityPackDll not found inside the downloaded NuGet package. KB downloads will not work."
                return
            }

            Copy-Item -Path $sourceDll -Destination $hapDll -Force
            Write-Verbose "$HtmlAgilityPackDll copied to: $hapDll"

            # Cleanup: remove extraction folder + nupkg
            Remove-Item $extractDir -Recurse -Force
            Remove-Item $tmpNupkg -Force
        }

        # --- Load the DLL (PS 5.x safe) ---
        # Load via byte array instead of file path so .NET does not hold a
        # file lock on the DLL — this allows -Clean to delete or overwrite it
        # even in the same PowerShell session.
        $hapLoaded = $false
        try {
            [void][HtmlAgilityPack.HtmlDocument]
            $hapLoaded = $true
        } catch {}

        if (-not $hapLoaded) {
            Write-Verbose "Loading HtmlAgilityPack from bytes: $hapDll"
            $hapBytes = [System.IO.File]::ReadAllBytes($hapDll)
            [void][System.Reflection.Assembly]::Load($hapBytes)
            Write-Debug "HtmlAgilityPack loaded from byte array (file lock avoided)"
        }
    }
}


# ==============================
# Main orchestration
# ==============================

try {
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
    if ($CreateISO) { Invoke-CreateISOWork }

    Write-Output "Completed"
} catch {
    Write-Output ""
    Write-Output "ERROR: $_"
    Write-Output "       Run the script again once the issue is resolved; completed steps will be skipped."
    exit 1
} finally {
    # -----------------------------------------------------------------------
    # Cleanup: release any resources that may have been left open if the
    # script was interrupted (Ctrl+C, early fatal error, etc.)
    # -----------------------------------------------------------------------

    # 1. Dismount the source ISO if it is still attached as a virtual drive.
    if ($ISO -and (Test-Path $ISO)) {
        try {
            $img = Get-DiskImage -ImagePath $ISO -ErrorAction SilentlyContinue
            if ($img -and $img.Attached) {
                Write-Output "Cleanup: dismounting ISO..."
                Dismount-DiskImage -ImagePath $ISO -ErrorAction SilentlyContinue | Out-Null
            }
        } catch {}
    }

    # 2. Discard any DISM-mounted WIM images that the servicing loop left open.
    #    Each active mount shows up as a non-empty subdirectory under WimsMounts.
    if ($dismExe -and $paths -and $paths.WimsMounts -and (Test-Path $paths.WimsMounts)) {
        $leftoverDirs = @(Get-ChildItem -Path $paths.WimsMounts -Directory -ErrorAction SilentlyContinue |
                          Where-Object { (Get-ChildItem $_.FullName -ErrorAction SilentlyContinue).Count -gt 0 })
        foreach ($mountDir in $leftoverDirs) {
            Write-Output "Cleanup: discarding leftover DISM mount at $($mountDir.FullName)..."
            Run-App $dismExe @('/Unmount-Image', "/MountDir:$($mountDir.FullName)", '/Discard') | Out-Null
            Remove-Item $mountDir.FullName -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}
