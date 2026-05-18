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
Copy various .cmd, .ps1, and .ini files with transformations to customize them for the current folder structure and configuration.

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

.PARAMETER Usage
Displays help and exits.

.NOTES
Fully compatible with Windows PowerShell 5.x.
Never use Write-Output, always use Write-Host or it breaks having normal output in a function that returns anything
DO NOT assign directly from a foreach loop to it will lose any kind of Write-Host, Write-Verbose, or Write-Debug from inside the loop.
Instead, accumulate results in a list and output the list at the end of the loop.

.LINK
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

    [switch]$Help,
    [switch]$Usage
)

# git hash
$GitHash = "8c8355f"

# Leadin to get ':' to line up in output. Write-xxxx (&$LeadIn "dism" "$dismExe")
$LeadIn = { param($Label, $Value) '{0,-20}: {1}' -f $Label, $Value }

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
    ExportDriversCmd      = 'ExportDrivers.cmd'
    InstallDriversCmd     = 'InstallDrivers.cmd'
    ExportRegsCmd         = 'ExportRegs.cmd'
    ExportRegsPs1         = 'ExportRegs.ps1'
    InstallRegsCmd        = 'InstallRegs.cmd'
    SetupConfigCleanIni   = 'SetupConfig-Clean.ini'
    SetupConfigUpgradeIni = 'SetupConfig-Upgrade.ini'
    CleanInstallCmd       = 'CleanInstall.cmd'
    UpgradeCmd            = 'Upgrade.cmd'
}

$kbDirs = @('SSU', 'OSCU', 'NET', 'MISC')
foreach ($u in $kbDirs) {
    $names[$u] = $u
}

$wimDirs = @('Indices', 'Mounts', 'Serviced', 'Final', 'Scratch', 'Logs')
foreach ($u in $wimDirs) {
    $names[$u] = $u
}

# ==============================
# RequiredFiles and RequiredTransforms
# ==============================
$names.RequiredFiles = @(
    $names.InstallDriversCmd,
    $names.InstallRegsCmd,
    $names.ExportDriversCmd,
    $names.ExportRegsCmd,
    $names.ExportRegsPs1,
    $names.SetupConfigCleanIni,
    $names.SetupConfigUpgradeIni,
    $names.CleanInstallCmd,
    $names.UpgradeCmd
)

# Each entry is: TargetFile, List of (SearchPattern, Replacement) pairs to apply to the target file before copying to the destination.
# **** The SearchPattern needs to be a regex to match the line to replace with the Replacement ****
$names.RequiredTransforms = @(
    @($names.ExportDriversCmd, @(
        @('set "FLD=$WinpeDriver$"', $names.WinpeDriver)
    )),
    @($names.InstallDriversCmd, @(
        @('set "FLD=$WinpeDriver$"', $names.WinpeDriver)
    )),
    @($names.ExportRegsPs1, @(
        @('set "FLD=Registry"', $names.Registry)
    )),
    @($names.InstallRegsCmd, @(
        @("$RegistryRoot = 'Registry'", $names.Registry)
    ))
)

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
function Protect-Token([string]$s) {
  if (-not $s) { return "unknown" }
  $t = $s -replace '[^\w\.-]+','_'
  $t = $t -replace '_+','_'
  return $t.Trim('_')
}

function Show-Usage {
  $name = Split-Path -Leaf $PSCommandPath
  Write-Host ""
  Write-Host "$name ($GitHash)" -ForegroundColor Cyan
  Write-Host "Usage:" -ForegroundColor Cyan
  Write-Host "  & $name [<Folder>] [-ISO <path>] [-DestISO <path>] [-OS <version>] [-Version <version>] [-Arch <arch>] [<Index Selections>] [<Work Selections>] [<More Options>]" -ForegroundColor Cyan
  Write-Host "  & $name [-ISO <path>] -ShowIndices'" -ForegroundColor Cyan
  Write-Host ""
  Write-Host "Key Options:" -ForegroundColor Cyan
  Write-Host "  <Folder>   Work directory (default: current directory)" -ForegroundColor Gray
  Write-Host "  <SrcISO>   Path to Windows Install ISO (overrides auto-detect) or you can use -ISO/SrcISO" -ForegroundColor Gray
  Write-Host "  <DestISO>  Path to output Windows Install ISO (overrides using <SrcISO>_KBs.iso) or you can use -DestISO" -ForegroundColor Gray
  Write-Host ""
  Write-Host "Index Selections:" -ForegroundColor Cyan
  Write-Host "  -Home            Select editions whose normalized label matches `"Home`" exactly." -ForegroundColor Gray
  Write-Host "  -Pro             Select editions whose normalized label matches `"Pro`" exactly." -ForegroundColor Gray
  Write-Host "  -Indices <spec>  Select editions based on a comma-separated selector string." -ForegroundColor Gray
  Write-Host "     Where <spec> can include:" -ForegroundColor Gray
  Write-Host "       - numbers and ranges: 6, 3-6, 7-*" -ForegroundColor Gray
  Write-Host "       - exact labels: `"Education N`"" -ForegroundColor Gray
  Write-Host "       - wildcard labels: `"*Home*`", `"* N*`"" -ForegroundColor Gray
  Write-Host "       - regex labels: `"re:^Education( N)?$`"" -ForegroundColor Gray
  Write-Host ""
  Write-Host "Work Selections:" -ForegroundColor Cyan
  Write-Host "  -Extract    Mount the source ISO and extract its full content tree." -ForegroundColor Gray
  Write-Host "  -Export     Export selected indices from the install.wim and boot.wim." -ForegroundColor Gray
  Write-Host "  -KB         Download OS and .NET updates." -ForegroundColor Gray
  Write-Host "  -Service    Apply downloaded KBs to the exported indices and produce final install.wim and boot.wim." -ForegroundColor Gray
  Write-Host "  -Files      Copy various .cmd, .ps1, and .ini files to the root of DestISO." -ForegroundColor Gray
  Write-Host "  -Prep       Prepare the destination ISO work area." -ForegroundColor Gray
  Write-Host "  -CreateISO  Create the DestISO using oscdimg." -ForegroundColor Gray
  Write-Host "  -All        Shorthand for all work selections." -ForegroundColor Gray
  Write-Host "  -Most       Shorthand for all work selections except -CreateISO." -ForegroundColor Gray
  Write-Host ""
  Write-Host "More Options:" -ForegroundColor Cyan
  Write-Host "  -ShowIndices     Print available image indices from the source ISO (or cached metadata) and exit." -ForegroundColor Gray
  Write-Host "  -UseADK          Prefer ADK dism.exe and oscdimg.exe when available." -ForegroundColor Gray
  Write-Host "  -UseSystem       Force system dism.exe and oscdimg.exe when available." -ForegroundColor Gray
  Write-Host "  -dism <path>     Explicit path to dism.exe." -ForegroundColor Gray
  Write-Host "  -oscdimg <path>  Explicit path to oscdimg.exe." -ForegroundColor Gray
  Write-Host "  -Clean           Overrides all other options combine with <Work Selections> to narrow the selection." -ForegroundColor Gray
  Write-Host "  -DryRun          Perform a dry run without making any changes." -ForegroundColor Gray
  Write-Host "  -Debug           Enable debug output." -ForegroundColor Gray
  Write-Host "  -Verbose         Enable verbose output." -ForegroundColor Gray
  Write-Host "  -Help            Display help information." -ForegroundColor Gray
  Write-Host "  -Usage           Display usage information." -ForegroundColor Gray
  Write-Host ""
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

function Remove-Folder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    if (Test-Path $path) {
        Remove-Item $Path -Recurse -Force -ErrorAction SilentlyContinue
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
            Write-Host "[DryRun] Would remove file  : $Path"
        } elseif (Test-Path $path) {
            Write-Host "Removing file  : $($Path)"
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
            Write-Host "[DryRun] Would remove folder: $Path"
        } elseif (Test-Path $Path) {
            Write-Host "Removing folder: $($Path)"
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
    if ([string]::IsNullOrEmpty($Path)) {
        throw "Write-JsonFile called with an empty Path (caller: $(Get-PSCallStack | Select-Object -Skip 1 -First 1 | ForEach-Object { "$($_.Command) line $($_.ScriptLineNumber)" }))"
    }
    $parent = Split-Path $Path -Parent
    if (-not [string]::IsNullOrEmpty($parent)) {
        Ensure-Folder -Path $parent
    }
    $Data | ConvertTo-Json -Depth 6 | Set-Content -Path $Path -Encoding UTF8
}

function Run-Dism {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string[]]$ArgumentList,

        [Parameter(Mandatory = $false)]
        [switch]$Capture,

        [Parameter(Mandatory = $false)]
        [int]$Indent = 0
    )

    # In .NET Framework (PS 5.x), ProcessStartInfo expects arguments as a single space-separated string
    $combinedArgs = $ArgumentList -join ' '

    Write-Verbose "Executing: $dismExe $combinedArgs"

    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfo.FileName               = $dismExe
    $processInfo.Arguments              = $combinedArgs
    $processInfo.RedirectStandardOutput = $true
    $processInfo.RedirectStandardError  = $true
    $processInfo.UseShellExecute        = $false
    $processInfo.CreateNoWindow         = $true

    # Use the system's default OEM code page to ensure special console characters don't break regex parsing
    $processInfo.StandardOutputEncoding = [System.Text.Encoding]::Default

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $processInfo

    # Use a generic list to accumulate strings if capturing internally
    if ($Capture) { $outputCollection = New-Object System.Collections.Generic.List[string] }

    $lastReportedPercent = -1
    try {
        $null = $process.Start()

        # StreamReader.ReadLine() treats carriage returns (\r) as line breaks.
        # This is ideal because DISM uses \r to overwrite its text progress bar in place.
        while (-not $process.StandardOutput.EndOfStream) {
            $line = $process.StandardOutput.ReadLine()
            if ($null -eq $line) { continue }

            # Always send raw output to Stream 4 (Verbose) if anyone is listening
            if ($line.Trim()) { 
                Write-Verbose ((' ' * $Indent) + $($line.Trim()))
            }

            if ($Capture) {
                # Accumulate the raw lines internally instead of looking for progress
                $outputCollection.Add($line)
            } else {
                # Regex looks for numbers immediately preceding a percentage sign (e.g., 20.0%)
                if ($line -match '(\d+(?:\.\d+)?)%') {
                    $currentPercent = [math]::Floor([double]$Matches[1])
                    $percentBucket  = [math]::Floor($currentPercent / 10) * 10
                    
                    if ($percentBucket -gt $lastReportedPercent -and $percentBucket -le 100) {
                        Write-Host ((' ' * $Indent) + "Progress: $percentBucket%")
                        $lastReportedPercent = $percentBucket
                    }
                }
            }
        }

        try { $process.WaitForExit() }       catch {}
        try { $rc = [int]$process.ExitCode } catch { $rc = 0 }
        Write-Debug "$($process.StartInfo.FileName) $($process.StartInfo.Arguments) exited with code $rc"

        # If capturing was requested, output the entire text array down Stream 1 - SPECIAL CASE
        if ($Capture) {
            Write-Output $outputCollection.ToArray()
        }

        # Callers that need the exit status check $LASTEXITCODE directly.
        $global:LASTEXITCODE = $rc
    } finally {
        # CRITICAL CLEANUP: If the user forces a hard halt using Ctrl+C, 
        # this block intercepts the abort and violently closes the dism.exe thread.
        if ($process) {
            if (-not $process.HasExited) {
                $process.Kill()
            }
            $process.Dispose()
        }
    }
}

# ==============================
# Report-Missing helper
# ==============================
function Report-Missing {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Required,

        [Parameter(Mandatory = $false)]
        [string[]]$AtLeastOne = @()
    )

    # Check Required files and build missing list
    $missingFiles = @()
    foreach ($file in $Required) {
        if (-not (Test-Path $file)) {
            $missingFiles += $file
        }
    }

    # Report missing files
    if ($missingFiles.Count -gt 0) {
        Write-Host "Missing file(s): $($missingFiles -join ', ')"
    }

    # Check AtLeastOne requirement
    if ($AtLeastOne.Count -gt 0) {
        $foundAny = $false
        foreach ($file in $AtLeastOne) {
            if (Test-Path $file) {
                $foundAny = $true
                break
            }
        }

        if (-not $foundAny) {
            Write-Host "At least 1 required: $($AtLeastOne -join ', ')"
            return $true
        }
    }

    # Return $true if any required files were missing
    return ($missingFiles.Count -gt 0)
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
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$WimPath
    )

    # Reads a WIM file and returns both the full image list and OS details
    # (WinOS, Version, Arch, Build) in a single object, making two DISM calls:
    #   /Get-WimInfo            to enumerate all images (Index + Name)
    #   /Get-WimInfo /Index:1   to get Version and Architecture from index 1
    Write-Debug "Get-WimMetadata: WimPath='$WimPath'"

    # --- Call 1: enumerate all images ---
    ($listOutput = Run-Dism @("/Get-WimInfo", "`"/WimFile:$WimPath`"") -Capture)
    $images      = [System.Collections.Generic.List[object]]::new()

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
    $buildNumber   = 0
    $archStr       = 'x64'
    ($detailOutput = Run-Dism @("/Get-WimInfo", "`"/WimFile:$WimPath`"", "/Index:1") -Capture)
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
        Write-Host "[DryRun] Would mount ISO: $ISO"
        Write-Host "[DryRun] Would validate $($names.BootFileBIOS) in $ISO"
        Write-Host "[DryRun] Would validate $($names.BootFileUEFI) in $ISO"
        Write-Host "[DryRun] Would validate $($paths.BootWimInIso) in $ISO"
        Write-Host "[DryRun] Would validate $($paths.InstallWimInIso) or $($paths.InstallEsdInIso) in $ISO"
        Write-Host "[DryRun] Would copy tree -> $($paths.SrcIsoRoot)"
        Write-Host "[DryRun] Would hardlink-copy $($paths.SrcIsoRoot) -> $($paths.DestIsoRoot) (excluding $($names.BootWim), $($names.InstallWim), $($names.InstallEsd))"
        Write-Host "[DryRun] Would export $($SelectedIndices.Count) indices to $($paths.WimsIndices)"
        return
    }

    Write-Host "Starting ExtractISO workflow..."
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
            Write-Host "ExtractISO already done for this ISO (extract.json matches)"
            Write-Debug  "extract.json: ISOPath='$($existingJson.ISOPath)' Date='$($existingJson.Date)'"
            return
        }
        Write-Host "ISO path changed (was '$($existingJson.ISOPath)'); cleaning SrcIsoRoot and re-extracting..."
        Remove-Folder $paths.SrcIsoRoot
    }

    Ensure-Folder $paths.SrcIsoContent

    Write-Host "Mounting ISO: $ISO"
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
        Write-Host "ISO mounted at: $driveLetter"

        # ---- Validate required files in ISO ----
        if (Report-Missing -Required @(
            (Join-Path $driveLetter $names.BootFileBIOS),
            (Join-Path $driveLetter $names.BootFileUEFI),
            (Join-Path $driveLetter $paths.BootWimInIso)
        ) -AtLeastOne @(
            (Join-Path $driveLetter $paths.InstallWimInIso),
            (Join-Path $driveLetter $paths.InstallEsdInIso)
        )) {
            throw "Source ISO validation failed. See above for missing file details."
        }
        Write-Host "Source ISO validation passed"

        # Copy the ISO tree to SrcIsoContent

        # Standardize paths to ensure safe string replacement for relative paths
        $sourceBase = $driveLetterRaw.TrimEnd('\') + '\'
        $destBase   = $paths.SrcIsoContent

        Write-Host "Copying ISO tree -> $destBase..."

        # Pre-gather items and calculate total payload size using 64-bit integers
        $allItems = Get-ChildItem -Path $sourceBase -Recurse
        $totalBytes = [int64]0
        foreach ($item in $allItems) {
            if (-not $item.PSIsContainer) { $totalBytes += $item.Length }
        }
        if ($totalBytes -eq 0) { $totalBytes = 1 } # Avoid divide-by-zero on empty sources

        $copiedBytes = [int64]0
        $lastReportedPercent = -1
        $bufferSize = 4194304 # 4MB chunk buffers for optimal performance

        foreach ($item in $allItems) {
            # Isolate the relative path (e.g., 'boot\bcd' instead of 'D:\boot\bcd')
            $relativePath = $item.FullName.Substring($sourceBase.Length)
            $targetPath   = Join-Path -Path $destBase -ChildPath $relativePath

            if ($item.PSIsContainer) {
                # Outputs to Stream 4 (Verbose). Captured by *>&1
                Write-Verbose "Folder: $relativePath"
                if (-not (Test-Path -Path $targetPath)) {
                    $null = New-Item -Path $targetPath -ItemType Directory -Force
                }
            } else {
                # Outputs to Stream 4 (Verbose). Captured by *>&1
                Write-Verbose "File:   $relativePath"
                
                # Mimic Robocopy /R:2 /W:1 (1 initial attempt + 2 retries, 1 sec wait)
                $retries = 2
                $success = $false
                
                while (-not $success) {
                    $sourceStream = $null
                    $destStream   = $null
                    $bytesWrittenThisAttempt = [int64]0
                    $aborted = $true # Default to true; proven false only on complete file success
                    
                    try {
                        # Ensure parent directory exists before streaming
                        $parentDir = Split-Path -Path $targetPath -Parent
                        if (-not (Test-Path -Path $parentDir)) {
                            $null = New-Item -Path $parentDir -ItemType Directory -Force
                        }

                        # Open low-level .NET file streams
                        $sourceStream = [System.IO.File]::OpenRead($item.FullName)
                        $destStream   = [System.IO.File]::Create($targetPath)
                        $buffer       = New-Object Byte[] $bufferSize

                        # Read and write in 4MB blocks
                        while (($bytesRead = $sourceStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                            $destStream.Write($buffer, 0, $bytesRead)
                            
                            $bytesWrittenThisAttempt += $bytesRead
                            $copiedBytes             += $bytesRead

                            # Calculate progress dynamically mid-file execution
                            $currentPercent = [math]::Floor(($copiedBytes / $totalBytes) * 100)
                            $percentBucket  = [math]::Floor($currentPercent / 10) * 10
                            
                            if ($percentBucket -gt $lastReportedPercent -and $percentBucket -le 100) {
                                # Outputs to Stream 1 (Success). Captured directly by tee
                                Write-Host "Copy progress: $percentBucket%"
                                $lastReportedPercent = $percentBucket
                            }
                        }

                        $aborted = $false # Entire file copied without interruption
                        $success = $true
                    } catch {
                        # Rollback global counter for what we wrote during this failed attempt
                        $copiedBytes -= $bytesWrittenThisAttempt

                        if ($retries -gt 0) {
                            Start-Sleep -Seconds 1
                            $retries--
                        } else {
                            throw "Native stream copy failed for '$($item.FullName)' to '$targetPath' after retries: $_"
                        }
                    } finally {
                        # CRITICAL: This block executes even if the pipeline is halted via Ctrl+C
                        if ($sourceStream) { $sourceStream.Close(); $sourceStream.Dispose() }
                        if ($destStream)   { $destStream.Close(); $destStream.Dispose() }
                        
                        # If an error occurred or user issued Ctrl+C, delete the incomplete file
                        if ($aborted -and (Test-Path -Path $targetPath)) {
                            Remove-Item -Path $targetPath -Force -ErrorAction SilentlyContinue
                        }
                    }
                }
            }
        }
    } catch {
        Write-Host "ERROR during ISO extraction: $_"
        # Clean up the partial SrcIsoContent so a re-run starts fresh
        Remove-Folder $paths.SrcIsoContent
        return
    } finally {
        if (Get-DiskImage -ImagePath $ISO | Where-Object { $_.Attached }) {
            Write-Host "Unmounting ISO..."
            Dismount-DiskImage -ImagePath $ISO -ErrorAction SilentlyContinue | Out-Null
        }
    }

    Write-JsonFile -Path $extractJson -Data @{ ISOPath = $ISO; Date = (Get-Date -Format s) }
    Write-Host "ExtractISO complete (extract.json written)"
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
        Write-Host "[DryRun] Would collect WIM metadata from SrcIsoContent"
        Write-Host "[DryRun] Would export $($SelectedIndices.Count) indices to $($paths.WimsIndices)"
        return
    }

    Write-Host "Starting Export workflow..."
    Write-Verbose "Invoke-Export: SourcesInSrc='$($paths.SourcesInSrc)' WimsIndices='$($paths.WimsIndices)'"

    Ensure-Folder $paths.WimsRoot
    Ensure-Folder $paths.WimsIndices

    $extractJson  = Join-Path $paths.SrcIsoRoot "extract.json"
    $metadataJson = Join-Path $paths.WimsIndices "wim-metadata.json"

    $extractMeta = Read-JsonFile -Path $extractJson
    if (-not $extractMeta) {
        Write-Warning "$extractJson not found. Run -Extract first."
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
    Write-Host "Collecting WIM metadata..."
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
    Write-Host "WIM metadata saved ($($installMeta.Images.Count) install image(s), $($bootImages.Count) boot image(s))"

    # Resolve index selection if not already set
    if ($SelectedIndices.Count -eq 0) {
        Write-Verbose "SelectedIndices empty; resolving from collected metadata..."
        $SelectedIndices = @(Resolve-IndexSelection -AllImages $installMeta.Images -SelectHome:$SelectHome -SelectPro:$SelectPro -IndicesStr $Indices)
        Write-Verbose "Resolved $($SelectedIndices.Count) index/indices"
    }

    Write-Host "Exporting $($SelectedIndices.Count) index/indices..."
    Write-Verbose "Selected: $($SelectedIndices.Index -join ', ')"

    $bootSrcIdx = if ($bootImages | Where-Object { $_.Index -eq 2 }) { 2 } else { 1 }

    # Cache these locally since they're used repeatedly in the loop
    $installWimName = $names.InstallWim
    $bootWimName    = $names.BootWim

    foreach ($img in $SelectedIndices) {
        $idx     = $img.Index
        $imgName = $img.Name
        Write-Host "  [Index $idx] $imgName"

        # -- Export install image --
        $installDest = Join-Path $paths.WimsIndices ("{0}_{1}" -f $idx, $installWimName)
        $installJson = $installDest + ".json"
        $existInstall = Read-JsonFile -Path $installJson
        $needInstall  = (-not $existInstall) -or ([datetime]::Parse($existInstall.ExportDate) -le $extractDate)

        if (-not $needInstall) {
            Write-Host "    $installWimName index $idx already exported ($($existInstall.ExportDate))"
        } else {
            Write-Host "    Exporting $installWimName index $idx..."
            Run-Dism @("/Export-Image", "`"/SourceImageFile:$installSrc`"", "`"/SourceIndex:$idx`"",
                       "`"/DestinationImageFile:$installDest`"", "/Compress:None", "/CheckIntegrity") -Indent 4
            $rc = $LASTEXITCODE
            if ($rc -ne 0) {
                Write-Warning "    DISM export failed for $installWimName index $idx (exit $rc)."
                Write-Warning "    Try running this script again. Skipping export for this index for now."
                return
            }
            Write-JsonFile -Path $installJson -Data @{ Index = $idx; Name = $imgName; ExportDate = (Get-Date -Format s) }
            Write-Host "    $installWimName index $idx exported"
        }

        # -- Export boot image --
        $bootDest = Join-Path $paths.WimsIndices ("{0}_{1}" -f $idx, $bootWimName)
        $bootJson = $bootDest + ".json"
        $existBoot = Read-JsonFile -Path $bootJson
        $needBoot  = (-not $existBoot) -or ([datetime]::Parse($existBoot.ExportDate) -le $extractDate)

        if (-not $needBoot) {
            Write-Host "    $bootWimName index $idx already exported ($($existBoot.ExportDate))"
        } else {
            Write-Host "    Exporting $bootWimName (src idx $bootSrcIdx) for index $idx..."
            Run-Dism @("/Export-Image", "`"/SourceImageFile:$bootSrc`"", "`"/SourceIndex:$bootSrcIdx`"",
                       "`"/DestinationImageFile:$bootDest`"", "/Compress:None", "/CheckIntegrity") -Indent 4
            $rc = $LASTEXITCODE
            if ($rc -ne 0) {
                Write-Warning "    DISM export failed for $bootWimName index $idx (exit $rc)."
                Write-Warning "    Try running this script again. Skipping export for this index for now."
                return
            } else {
                Write-JsonFile -Path $bootJson -Data @{ Index = $idx; SourceBootIndex = $bootSrcIdx; ExportDate = (Get-Date -Format s) }
                Write-Host "    $bootWimName index $idx exported"
            }
        }
    }

    Write-Host "Export workflow complete"
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

    Write-Verbose ("Searching for {0}{1}..." -f $Query, $(if ($FirstOnly) { " (first result only)" } else { "" }))

    $Encoded = [uri]::EscapeDataString($Query)
    $Uri = "https://www.catalog.update.microsoft.com/Search.aspx?q=$Encoded"

    Write-Debug "Encoded URI: $Uri"

    $Html = Invoke-CatalogRequest -Uri $Uri
    if (-not $Html) {
        Write-Warning "No HTML returned from $Uri"
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
    Write-Debug  ('{0,3} {1,-36} TargetFolder:' -f "ID:", "Guid:")
    for ($i = 0; $i -lt $ids.Count; $i++) {
        Write-Debug ('{0,3} {1,-36} {2}' -f $($i + 1), $($ids[$i].Guid), $($ids[$i].TargetFolder))
    }
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

    $links = @()
    try {
        foreach ($m in $matches) {
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

            $links += [PSCustomObject]@{
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

    Write-Host ("Preparing downloads for update {0}: {1}" -f $Update.Guid, $Update.Title)

    Ensure-Folder $TargetFolder

    $results = @()

    # No URLs, then nothing to do
    if (-not $Update.DownloadUrls -or $Update.DownloadUrls.Count -eq 0) {
        Write-Host ("No download URLs for update {0}" -f $Update.Guid)
        return $results
    }

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
                Write-Host ("  {0,3}% {1,12:N0}/{2:N0} bytes" -f 0, 0, $total)

                while (($read = $inStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                    $outStream.Write($buffer, 0, $read)
                    $totalRead += $read

                    if ($total -gt 0) {
                        $pct = [math]::Floor(($totalRead / $total) * 100)

                        if ($pct -ge $nextMark) {
                            # Pipe-safe: always print full lines, never CR
                            Write-Host ("  {0,3}%  {1,12:N0}/{2:N0} bytes" -f $pct, $totalRead, $total)
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
                    Remove-Item $destPath -Force -ErrorAction SilentlyContinue
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
        return [System.Management.Automation.Internal.AutomationNull]::Value
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
        return [System.Management.Automation.Internal.AutomationNull]::Value
    }

    # ------------------------------------------------------------
    # DOWNLOAD LINKS (via Get-UpdateLinks)
    # ------------------------------------------------------------

    Write-Host "Finding download links for $title"

    # Need to use () and crazy ForEach-Object to see the Write-xxxx output
    $links = (Get-UpdateLinks -Guid $Guid | ForEach-Object { $_ })
    $downloadUrls = @()
    if ($links) {
        $downloadUrls = $links.URL | Select-Object -Unique
    }

    Write-Host ("Found {0} file(s) for this update" -f $downloadUrls.Count)

    # Not a keeper if no download links
    if ($downloadUrls.Count -eq 0) {
        Write-Host ("{0} has no download links" -f $Guid)
        return [System.Management.Automation.Internal.AutomationNull]::Value
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
            Write-Host ("[DryRun] Would fill: {0}" -f $paths["KBs$u"])
        }
        return
    }

    Write-Host "Starting KB update workflow..."

    # Ensure folders exist
    $KBsPaths = @()
    foreach ($u in $kbDirs) {
        $KBsPaths += $paths["KBs$u"]
    }

    foreach ($folder in $KBsPaths) {
        Ensure-Folder $folder
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

    $results = @()
    foreach ($q in $queries) {
        $results += Search-UpdateCatalogHtml -Query $q.Query -FirstOnly $q.FirstOnly -TargetFolder $q.TargetFolder
    }

    $allGuids = $results | Sort-Object Guid -Unique
    Write-Host ("Found {0} total updates to process" -f $allGuids.Count)
    Write-Debug  (('{0,-36} TargetFolder:{1}' -f "Guid:", "`n") + (@($allGuids | ForEach-Object { '{0} {1}' -f $_.Guid, $_.TargetFolder }) -join "`n"))

    if ($allGuids.Count -eq 0) {
        Write-Host "No updates found"
        return
    }

    Write-Host "Retrieving update details..."

    $count   = 0
    $details = @()
    foreach ($g in $allGuids) {
        Write-Debug ("Resolving details for {0} ({1})" -f $g.Guid, $g.TargetFolder)
        # 1. Run the function in a pipeline so Write-Host is visible
        $detail = Get-UpdateDetails -Count (++$count) -Guid $g.Guid -TargetFolder $g.TargetFolder | ForEach-Object { $_ }

        # 2. Append the returned object separately
        if ($detail -ne $null -and
            $detail -ne [System.Management.Automation.Internal.AutomationNull]::Value) {

            $details += $detail
        }
    }


    if ($details.Count -eq 0) {
        Write-Host "No usable updates after details resolution"
        return
    }

    Write-Host ("Remaining applicable updates: {0}" -f $details.Count)
    foreach ($d in $details) {
        Write-Debug ("Update: {0}`n  Title: {1}`n  KB: {2}`n  URLs: {3}" -f $d.Guid, $d.Title, $d.KB, ($d.DownloadUrls -join ', '))
    }

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
            Remove-Item $path -Force -ErrorAction SilentlyContinue
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
                Remove-Item $manifestPath -Force -ErrorAction SilentlyContinue
            }
        }
    }

    Write-Host "KB update workflow complete"
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
        Write-Host "[DryRun] Would service extracted indices in $($paths.WimsIndices)"
        Write-Host "[DryRun] Would apply SSU packages from : $($paths.KBsSSU)"
        Write-Host "[DryRun] Would apply LCU packages from : $($paths.KBsOSCU)"
        Write-Host "[DryRun] Would service winre.wim inside each index's install.wim"
        Write-Host "[DryRun] Would assemble final install.wim -> $($paths.InstallWimInDest)"
        Write-Host "[DryRun] Would assemble final boot.wim   -> $($paths.BootWimInDest)"
        return
    }

    Write-Host  "Starting Service workflow..."
    Write-Verbose "Invoke-ServiceWork: WimsIndices='$($paths.WimsIndices)' WimsFinal='$($paths.WimsFinal)'"
    Write-Debug   "Invoke-ServiceWork: KBsSSU='$($paths.KBsSSU)' KBsOSCU='$($paths.KBsOSCU)' WimsMounts='$($paths.WimsMounts)'"

    $CompressionType = 'Maximum'  # None, Fast, Maximum

    Remove-Folder $paths.WimsMounts
    Ensure-Folder $paths.WimsMounts
    Ensure-Folder $paths.WimsServiced
    Ensure-Folder $paths.WimsScratch
    Ensure-Folder $paths.WimsLogs

    if (-not (Test-Path $paths.WimsIndices)) {
        Write-Host "No indices folder found: $($paths.WimsIndices). Run with -Export first."
        return
    }

    # Discover extracted install.wim files and derive index numbers
    $indexFiles = @(Get-ChildItem -Path $paths.WimsIndices -Filter "*_$($names.InstallWim)" -File -ErrorAction SilentlyContinue)
    if ($indexFiles.Count -eq 0) {
        Write-Host "No extracted $($names.InstallWim) files found in $($paths.WimsIndices). Run with -Export first."
        return
    }

    $extractedIndices = $indexFiles |
        ForEach-Object { [int](($_.BaseName -split '_')[0]) } |
        Sort-Object
    Write-Host "Found $($extractedIndices.Count) extracted indices: $($extractedIndices -join ', ')"

    # Gather available packages (.msu and .cab)
    $ssuFiles = @(
        Get-ChildItem -Path $paths.KBsSSU -Include '*.msu', '*.cab' -Recurse -ErrorAction SilentlyContinue
    )
    $lcuFiles = @(
        Get-ChildItem -Path $paths.KBsOSCU -Include '*.msu', '*.cab' -Recurse -ErrorAction SilentlyContinue
    )
    $hasSSU = $ssuFiles.Count -gt 0
    $hasLCU = $lcuFiles.Count -gt 0

    Write-Host  "Packages available - SSU: $hasSSU ($($ssuFiles.Count) files), LCU: $hasLCU ($($lcuFiles.Count) files)"
    Write-Verbose "SSU : $($ssuFiles.Name -join ', ')"
    Write-Verbose "LCU : $($lcuFiles.Name -join ', ')"

    # -----------------------------------------------------------------------
    # Per-index servicing
    # -----------------------------------------------------------------------
    foreach ($idx in $extractedIndices) {
        Write-Host "--- Index $idx ---"
        Write-Verbose "Processing index $idx"

        $installWimPath   = Join-Path $paths.WimsIndices ("{0}_{1}" -f $idx, $names.InstallWim)
        $bootWimPath      = Join-Path $paths.WimsIndices ("{0}_{1}" -f $idx, $names.BootWim)

        $installDoneChkpt = Join-Path $paths.WimsIndices ("{0}_{1}.serviced.json" -f $idx, $names.InstallWim)
        $bootDoneChkpt    = Join-Path $paths.WimsIndices ("{0}_{1}.serviced.json" -f $idx, $names.BootWim)
        $winreExtChkpt    = Join-Path $paths.WimsIndices ("{0}_{1}.extracted.json" -f $idx, $names.WinreWim)
        $winreDoneChkpt   = Join-Path $paths.WimsIndices ("{0}_{1}.serviced.json" -f $idx, $names.WinreWim)

        $scratchDir       = $paths.WimsScratch
        $installMountDir  = Join-Path $paths.WimsMounts  ("mount_{0}_{1}"     -f $idx, $names.InstallWim)
        $installMountLog  = Join-Path $paths.WimsLogs    ("mount_{0}_{1}.log" -f $idx, $names.InstallWim)
        $installLog       = Join-Path $paths.WimsLogs    ("{0}_{1}.log"       -f $idx, $names.InstallWim)
        $winreMountDir    = Join-Path $paths.WimsMounts  ("mount_{0}_{1}"     -f $idx, $names.WinreWim)
        $winreMountLog    = Join-Path $paths.WimsLogs    ("mount_{0}_{1}.log" -f $idx, $names.WinreWim)
        $winreWimPath     = Join-Path $paths.WimsIndices ("{0}_{1}"           -f $idx, $names.WinreWim)
        $winreLog         = Join-Path $paths.WimsLogs    ("{0}_{1}.log"       -f $idx, $names.WinreWim)
        $bootMountDir     = Join-Path $paths.WimsMounts  ("mount_{0}_{1}"     -f $idx, $names.BootWim)
        $bootMountLog     = Join-Path $paths.WimsLogs    ("mount_{0}_{1}.log" -f $idx, $names.BootWim)
        $bootLog          = Join-Path $paths.WimsLogs    ("{0}_{1}.log"       -f $idx, $names.BootWim)

        # ---- Service install.wim ----
        if (Read-JsonFile -Path $installDoneChkpt) {
            Write-Host "  $($names.InstallWim) index $idx already serviced"
        } else {
            if ($hasSSU -or $hasLCU) {
                Write-Host "  Servicing $($names.InstallWim) for index $idx..."
                Ensure-Folder $installMountDir

                try {
                    Write-Host  "  Mounting $installWimPath -> $installMountDir"
                    Run-Dism @("/Mount-Image", "`"/ImageFile:$installWimPath`"", "/Index:1", "`"/MountDir:$installMountDir`"", "`"/ScratchDir:$scratchDir`"", "`"/LogPath:$installMountLog`"") -Indent 2
                    if ($LASTEXITCODE -ne 0) { throw "DISM mount $($names.InstallWim) failed for index $idx (exit $LASTEXITCODE)" }

                    # ---- Service winre.wim inside install.wim ----
                    if ($hasSSU) {
                        $winreInMount = Join-Path $installMountDir $paths.WinreWimInWim
                        Write-Debug "  Checking for $($names.WinreWim) at: $winreInMount"

                        if (Test-Path $winreInMount) {
                            if (Read-JsonFile -Path $winreDoneChkpt) {
                                Write-Host "  $($names.WinreWim) index $idx already serviced"
                            } else {
                                # Extract winre.wim
                                if (-not (Read-JsonFile -Path $winreExtChkpt)) {
                                    Write-Host "  Extracting $($names.WinreWim) from mounted install image..."
                                    Copy-Item -Path $winreInMount -Destination $winreWimPath -Force
                                    Write-JsonFile -Path $winreExtChkpt -Data @{ Index = $idx; ExtractedDate = (Get-Date -Format s) }
                                    Write-Host "  $($names.WinreWim) extracted"
                                } else {
                                    Write-Host "  $($names.WinreWim) already extracted"
                                }

                                # Mount winre.wim
                                Ensure-Folder $winreMountDir
                                Write-Host  "  Mounting $($names.WinreWim) -> $winreMountDir"
                                Run-Dism @("/Mount-Image", "`"/ImageFile:$winreWimPath`"", "/Index:1", "`"/MountDir:$winreMountDir`"", "`"/ScratchDir:$scratchDir`"", "`"/LogPath:$winreMountLog`"") -Indent 2
                                if ($LASTEXITCODE -ne 0) { throw "DISM mount $($names.WinreWim) failed for index $idx (exit $LASTEXITCODE)" }

                                # Apply SSU packages to winre
                                foreach ($pkg in $ssuFiles) {
                                    Write-Host  "  Applying SSU to $($names.WinreWim): $($pkg.Name)"
                                    $leaf   = Split-Path $pkg.FullName -Leaf
                                    $pkgLog = $winreLog.Replace(".log", ("_{0}.log" -f (Protect-Token $leaf)))
                                    Run-Dism @("/Add-Package", "`"/Image:$winreMountDir`"", "`"/PackagePath:$($pkg.FullName)`"", "/ScratchDir:$scratchDir`"", "`"/LogPath:$pkgLog`"") -Indent 2
                                    if ($LASTEXITCODE -ne 0) {
                                        Write-Warning "  DISM SSU->$($names.WinreWim) failed: $($pkg.Name) index $idx (exit $LASTEXITCODE)"
                                    }
                                }

                                # Unmount and commit winre
                                Write-Host "  Unmounting $($names.WinreWim) (commit)..."
                                Run-Dism @("/Unmount-Image", "`"/MountDir:$winreMountDir`"", "/Commit") -Indent 2
                                if ($LASTEXITCODE -ne 0) { throw "DISM unmount $($names.WinreWim) failed for index $idx (exit $LASTEXITCODE)" }

                                # Reinsert serviced winre.wim back into mounted install.wim
                                Write-Host "  Reinserting serviced $($names.WinreWim) into install image..."
                                Copy-Item -Path $winreWimPath -Destination $winreInMount -Force
                                Write-JsonFile -Path $winreDoneChkpt -Data @{ Index = $idx; ServicedDate = (Get-Date -Format s) }
                                Write-Host "  $($names.WinreWim) serviced and reinserted"
                            }
                        } else {
                            Write-Verbose "  $($names.WinreWim) not found in mounted install image at $winreInMount; skipping winre servicing"
                        }

                        # Apply SSU packages to install.wim
                        foreach ($pkg in $ssuFiles) {
                            Write-Host  "  Applying SSU to $($names.InstallWim): $($pkg.Name)"
                            $leaf   = Split-Path $pkg.FullName -Leaf
                            $pkgLog = $installLog.Replace(".log", ("_{0}.log" -f (Protect-Token $leaf)))
                            Run-Dism @("/Add-Package", "`"/Image:$installMountDir`"", "`"/PackagePath:$($pkg.FullName)`"", "/ScratchDir:$scratchDir`"", "`"/LogPath:$pkgLog`"") -Indent 2
                            if ($LASTEXITCODE -ne 0) {
                                Write-Warning "  DISM SSU->$($names.InstallWim) failed: $($pkg.Name) index $idx (exit $LASTEXITCODE)"
                            }
                        }
                    }

                    # Apply LCU (OSCU) packages to install.wim
                    if ($hasLCU) {
                        foreach ($pkg in $lcuFiles) {
                            Write-Host  "  Applying LCU to $($names.InstallWim): $($pkg.Name)"
                            $leaf   = Split-Path $pkg.FullName -Leaf
                            $pkgLog = $installLog.Replace(".log", ("_{0}.log" -f (Protect-Token $leaf)))
                            Run-Dism @("/Add-Package", "`"/Image:$installMountDir`"", "`"/PackagePath:$($pkg.FullName)`"", "/ScratchDir:$scratchDir`"", "`"/LogPath:$pkgLog`"") -Indent 2
                            if ($LASTEXITCODE -ne 0) {
                                Write-Warning "  DISM LCU->$($names.InstallWim) failed: $($pkg.Name) index $idx (exit $LASTEXITCODE)"
                            }
                        }
                    }

                    # Unmount and commit install.wim
                    Write-Host "  Unmounting $($names.InstallWim) index $idx (commit)..."
                    Run-Dism @("/Unmount-Image", "`"/MountDir:$installMountDir`"", "/Commit") -Indent 2
                    if ($LASTEXITCODE -ne 0) { throw "DISM unmount $($names.InstallWim) failed for index $idx (exit $LASTEXITCODE)" }

                    Write-JsonFile -Path $installDoneChkpt -Data @{ Index = $idx; ServicedDate = (Get-Date -Format s) }
                    Write-Host "  $($names.InstallWim) index $idx serviced"

                } catch {
                    Write-Host "  ERROR servicing $($names.InstallWim) index $idx`: $_"
                    if (Test-Path $winreMountDir) {
                        Write-Host "  Discarding mounted $($names.WinreWim)..."
                        Run-Dism @("/Unmount-Image", "`"/MountDir:$winreMountDir`"", "/Discard") -Indent 2
                    }
                    if (Test-Path $installMountDir) {
                        Write-Host "  Discarding mounted $($names.InstallWim)..."
                        Run-Dism @("/Unmount-Image", "`"/MountDir:$installMountDir`"", "/Discard") -Indent 2
                    }
                    throw
                } finally {
                    Remove-Folder $winreMountDir
                    Remove-Folder $installMountDir
                }
            } else {
                Write-Host "  No SSU/LCU packages; marking $($names.InstallWim) index $idx as done"
                Write-JsonFile -Path $installDoneChkpt -Data @{ Index = $idx; ServicedDate = (Get-Date -Format s) }
            }
        }

        # ---- Service boot.wim ----
        if (-not (Test-Path $bootWimPath)) {
            Write-Verbose "  No $($names.BootWim) for index $idx; skipping boot servicing"
        } elseif (Read-JsonFile -Path $bootDoneChkpt) {
            Write-Host "  $($names.BootWim) index $idx already serviced"
        } elseif (-not $hasSSU) {
            Write-Host "  No SSU packages; marking $($names.BootWim) index $idx as done"
            Write-JsonFile -Path $bootDoneChkpt -Data @{ Index = $idx; ServicedDate = (Get-Date -Format s) }
        } else {
            Write-Host "  Servicing $($names.BootWim) for index $idx..."
            Ensure-Folder $bootMountDir

            try {
                Write-Host  "  Mounting $bootWimPath -> $bootMountDir"
                Run-Dism @("/Mount-Image", "`"/ImageFile:$bootWimPath`"", "/Index:1", "`"/MountDir:$bootMountDir`"", "`"/ScratchDir:$scratchDir`"", "`"/LogPath:$bootMountLog`"") -Indent 2
                if ($LASTEXITCODE -ne 0) { throw "DISM mount $($names.BootWim) failed for index $idx (exit $LASTEXITCODE)" }

                foreach ($pkg in $ssuFiles) {
                    Write-Host  "  Applying SSU to $($names.BootWim): $($pkg.Name)"
                    $leaf   = Split-Path $pkg.FullName -Leaf
                    $pkgLog = $bootLog.Replace(".log", ("_{0}.log" -f (Protect-Token $leaf)))
                    Run-Dism @("/Add-Package", "`"/Image:$bootMountDir`"", "`"/PackagePath:$($pkg.FullName)`"", "/ScratchDir:$scratchDir`"", "`"/LogPath:$pkgLog`"") -Indent 2
                    if ($LASTEXITCODE -ne 0) {
                        Write-Warning "  DISM SSU->$($names.BootWim) failed: $($pkg.Name) index $idx (exit $LASTEXITCODE)"
                    }
                }

                Write-Host "  Unmounting $($names.BootWim) index $idx (commit)..."
                Run-Dism @("/Unmount-Image", "`"/MountDir:$bootMountDir`"", "/Commit") -Indent 2
                if ($LASTEXITCODE -ne 0) { throw "DISM unmount $($names.BootWim) failed for index $idx (exit $LASTEXITCODE)" }

                Write-JsonFile -Path $bootDoneChkpt -Data @{ Index = $idx; ServicedDate = (Get-Date -Format s) }
                Write-Host "  $($names.BootWim) index $idx serviced"

            } catch {
                Write-Host "  ERROR servicing $($names.BootWim) index $idx`: $_"
                if (Test-Path $bootMountDir) {
                    Run-Dism @("/Unmount-Image", "`"/MountDir:$bootMountDir`"", "/Commit") -Indent 2
                }
                throw
            } finally {
                Remove-Folder $bootMountDir
            }
        }
    }

    # -----------------------------------------------------------------------
    # Final assembly (serial compression can be slow)
    # -----------------------------------------------------------------------
    Write-Host "Final assembly: combining serviced indices (compression: $CompressionType) -> $($paths.WimsFinal)..."
    Ensure-Folder $paths.WimsFinal

    $compressionMap  = @{ 'None' = 'none'; 'Fast' = 'fast'; 'Maximum' = 'max' }
    $dismCompression = $compressionMap[$CompressionType]
    $finalJson    = Join-Path $paths.WimsFinal "final.json"
    $finalMeta    = Read-JsonFile -Path $finalJson
    if (-not $finalMeta) { $finalMeta = @{} }

    # Assemble one WIM (install or boot) by exporting all per-index source files
    # into a single compressed destination.  DISM creates the file on the first
    # call and appends on subsequent ones.  The arguments are identical either
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
        $baseArgs = @('/Export-Image', "`"/DestinationImageFile:$DestPath`"", '/SourceIndex:1', "`"/Compress:$Compression`"", "/CheckIntegrity")
        try {
            Write-Host "Assembling final $WimLabel ($($Sources.Count) index/indices)..."
            for ($i = 0; $i -lt $Sources.Count; $i++) {
                Write-Host "  Index $($Indices[$i])..."
                Run-Dism ($baseArgs + @("`"/SourceImageFile:$($Sources[$i])`"")) -Indent 2
                if ($LASTEXITCODE -ne 0) { throw "DISM failed on $WimLabel index $($Indices[$i]) (exit $LASTEXITCODE)" }
            }
            $finalMeta[$DateKey] = (Get-Date -Format s)
            Write-JsonFile -Path $finalJson -Data $finalMeta
            Write-Host "Final $WimLabel assembled"
        } catch {
            Write-Host "ERROR assembling final $WimLabel`: $_"
            Remove-Folder $DestPath
        }
    }

    # -- Assemble final install.wim --
    $finalInstallPath = Join-Path $paths.WimsFinal $names.InstallWim
    if ($finalMeta.InstallWimDate) {
        Write-Host "Final $($names.InstallWim) already assembled ($($finalMeta.InstallWimDate))"
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
        Write-Host "Final $($names.BootWim) already assembled ($($finalMeta.BootWimDate))"
    } else {
        $sortedBootWims = @($extractedIndices |
            ForEach-Object { Join-Path $paths.WimsIndices ("{0}_{1}" -f $_, $names.BootWim) } |
            Where-Object   { Test-Path $_ })
        Invoke-AssembleWim -WimLabel $names.BootWim -DestPath $finalBootPath `
                           -Sources $sortedBootWims -Indices $extractedIndices `
                           -Compression $dismCompression -DateKey 'BootWimDate'
    }

    Write-Host "Service workflow complete"
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
        Write-Host "[DryRun] Would run: DISM /export-driver"
        return
    }

    Write-Host  "Exporting drivers..."
    Write-Verbose "Invoke-DriverWork: WinpeDriverRoot='$WinpeDriverRoot'"
    Ensure-Folder $WinpeDriverRoot
    $driverArgs = "/online /export-driver /destination:`"$WinpeDriverRoot`""
    Write-Debug "dism $driverArgs"
    $p = Start-Process -FilePath $dismExe -ArgumentList $driverArgs -NoNewWindow -PassThru -Wait -RedirectStandardOutput "$WinpeDriverRoot\dism.log"
    if ($p.ExitCode -ne 0) {
        Write-Host "ERROR: DISM /export-driver failed (exit $($p.ExitCode)), check $WinpeDriverRoot\dism.log"
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
        Ensure-Folder $RegistryRoot
    }

    #
    # HELPERS
    #
    function RegSafeName([string]$key) {
        ($key -replace '[\\/:*?"<>|]', '_') + '.reg'
    }

    function RegExportEntireKey([string]$key, [string]$dest) {
        if ($DryRun) {
            Write-Host "[DryRun] Would export ENTIRE key: $key -> $dest"
        } else {
            Write-Host "Export ENTIRE key: $key -> $dest"
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
            Write-Host "No specific values requested for $key; skipping."
            return
        }

        if ($DryRun) {
            Write-Host "[DryRun] Would export values [$($allValues -join ', ')] from $key -> $dest"
            return
        }

        Write-Host "Export specific values [$($allValues -join ', ')] from $key -> $dest"

        $query = reg.exe query "$key" /v * 2>$null
        if (-not $query) {
            Write-Host "WARNING: No data returned for $key"
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
                Write-Host "[DryRun] Would delete ENTIRE key: $key"
            } else {
                Write-Host "[DryRun] Would delete values [$($values -join ', ')] from $key"
            }
            return
        }

        Write-Host "Appending delete instructions for $key -> $dest"

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
                Write-Host "[DryRun] Would export ENTIRE key: $key -> $dest"
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
        Write-Host "[DryRun] Would write: $path"
    } else {
        Write-Host "Writing: $path"
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
        Write-Host "[DryRun] Would write: $path"
    } else {
        Write-Host "Writing: $path"
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
        Write-Host "[DryRun] Would write: $cleanPath"
        Write-Host "[DryRun] Would write: $upgradePath"
    } else {
        Write-Host "Writing: $cleanPath"
        Set-Content -LiteralPath $cleanPath   -Value $cleanTemplate   -Encoding ASCII
        Write-Host "Writing: $upgradePath"
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
        Write-Host "[DryRun] Would write: $cleanPath"
        Write-Host "[DryRun] Would write: $upgradePath"
    } else {
        # Fill in the actual name for the files in the template
        $cleanContent   = $cleanTemplate -f $names.SetupConfigCleanIni
        $upgradeContent = $upgradeTemplate -f $names.SetupConfigUpgradeIni

        Write-Host "Writing: $cleanPath"
        Set-Content -LiteralPath $cleanPath   -Value $cleanContent   -Encoding ASCII
        Write-Host "Writing: $upgradePath"
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
        Write-Host "[DryRun] Would hardlink-copy $($paths.SrcIsoContent) -> $($paths.DestIsoContent)"
        Write-Host "[DryRun] Would copy $finalInstall -> $($paths.InstallWimInDest)"
        Write-Host "[DryRun] Would copy $finalBoot    -> $($paths.BootWimInDest)"
        return
    }

    Write-Host "Starting PrepDestISO workflow..."
    Write-Verbose "Invoke-PrepDestISO: SrcIsoContent='$($paths.SrcIsoContent)' DestIsoContent='$($paths.DestIsoContent)'"
    $prepJson   = Join-Path $paths.DestIsoRoot "prep.json"
    $extractJson = Join-Path $paths.SrcIsoRoot "extract.json"
    $finalJson   = Join-Path $paths.WimsFinal "final.json"
    $finalInstall = Join-Path $paths.WimsFinal $names.InstallWim
    $finalBoot    = Join-Path $paths.WimsFinal $names.BootWim

    $extractMeta = Read-JsonFile -Path $extractJson
    if (-not $extractMeta) {
        Write-Warning "$extractJson not found. Run -Extract first."
        return
    }
    $extractDate = [datetime]::Parse($extractMeta.Date)

    $finalMeta = Read-JsonFile -Path $finalJson
    $prep      = Read-JsonFile -Path $prepJson

    # ---- Step A: Hardlink-copy SrcIsoContent -> DestIsoContent ----
    $needHardlink = (-not $prep -or -not $prep.HardlinkDate) -or
                    ([datetime]::Parse($prep.HardlinkDate) -le $extractDate)

    if ($needHardlink) {
        Write-Host "Hardlink-copying $($paths.SrcIsoContent) -> $($paths.DestIsoContent) (excluding install/boot images)..."
        if (Test-Path $paths.DestIsoRoot) {
            Write-Host "Removing existing DestIsoRoot..."
            Remove-Folder $paths.DestIsoRoot
        }
        Ensure-Folder $paths.DestIsoContent

        $excludeNames = @($names.BootWim, $names.InstallWim, $names.InstallEsd)
        $allFiles = @(Get-ChildItem -Path $paths.SrcIsoContent -Recurse -File -ErrorAction SilentlyContinue)
        $total = $allFiles.Count; $done = 0; $lastPct = -1

        Write-Host "  Hardlinking $total files..."
        foreach ($file in $allFiles) {
            $done++
            $pct = [math]::Floor(($done / [math]::Max($total, 1)) * 100)
            if ($pct -ge ($lastPct + 10)) {
                Write-Host ("  {0,3}%  {1}/{2} files" -f $pct, $done, $total)
                $lastPct = $pct - ($pct % 10)
            }
            if ($file.Name -in $excludeNames) { continue }
            $rel  = $file.FullName.Substring($paths.SrcIsoContent.TrimEnd('\').Length).TrimStart('\')
            $dest = Join-Path $paths.DestIsoContent $rel
            Ensure-Folder (Split-Path $dest -Parent)
            if (-not (Test-Path $dest)) {
                try {
                    New-Item -ItemType HardLink -Path $dest -Value $file.FullName -Force -ErrorAction Stop | Out-Null
                } catch {
                    Write-Warning "Hardlink failed for '$rel'; copying: $_"
                    Copy-Item -Path $file.FullName -Destination $dest -Force
                }
            }
        }
        Write-Host "  Hardlink tree complete"
        $prep = @{ HardlinkDate = (Get-Date -Format s) }
        Write-JsonFile -Path $prepJson -Data $prep
    } else {
        Write-Host "DestIsoContent hardlink-copy already current (prep.json: $($prep.HardlinkDate))"
    }

    $prep = Read-JsonFile -Path $prepJson
    if (-not $prep) { $prep = @{} }

    # ---- Step B: Copy final install.wim ----
    $finalInstDate = if ($finalMeta -and $finalMeta.InstallWimDate) { [datetime]::Parse($finalMeta.InstallWimDate) } else { [datetime]::MinValue }
    $destInstDate  = if ($prep.InstallWimDate) { [datetime]::Parse($prep.InstallWimDate) } else { [datetime]::MinValue }

    if ((Test-Path $finalInstall) -and ($destInstDate -le $finalInstDate)) {
        Write-Host "Copying $($names.InstallWim) -> $($paths.InstallWimInDest)..."
        Ensure-Folder (Split-Path $paths.InstallWimInDest -Parent)
        Copy-Item -Path $finalInstall -Destination $paths.InstallWimInDest -Force
        $prep['InstallWimDate'] = (Get-Date -Format s)
        Write-JsonFile -Path $prepJson -Data $prep
    } elseif (Test-Path $paths.InstallWimInDest) {
        Write-Host "$($names.InstallWim) already current (prep.json: $($prep.InstallWimDate))"
    } else {
        Write-Warning "Final $($names.InstallWim) not found at: $finalInstall (run -Service first)"
    }

    $prep = Read-JsonFile -Path $prepJson
    if (-not $prep) { $prep = @{} }

    # ---- Step C: Copy final boot.wim ----
    $finalBootDate = if ($finalMeta -and $finalMeta.BootWimDate) { [datetime]::Parse($finalMeta.BootWimDate) } else { [datetime]::MinValue }
    $destBootDate  = if ($prep.BootWimDate) { [datetime]::Parse($prep.BootWimDate) } else { [datetime]::MinValue }

    if ((Test-Path $finalBoot) -and ($destBootDate -le $finalBootDate)) {
        Write-Host "Copying $($names.BootWim) -> $($paths.BootWimInDest)..."
        Ensure-Folder (Split-Path $paths.BootWimInDest -Parent)
        Copy-Item -Path $finalBoot -Destination $paths.BootWimInDest -Force
        $prep['BootWimDate'] = (Get-Date -Format s)
        Write-JsonFile -Path $prepJson -Data $prep
    } elseif (Test-Path $paths.BootWimInDest) {
        Write-Host "$($names.BootWim) already current (prep.json: $($prep.BootWimDate))"
    } else {
        Write-Warning "Final $($names.BootWim) not found at: $finalBoot (run -Service first)"
    }

    Write-Host "PrepDestISO workflow complete"
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
        Write-Host "[DryRun] Would create ISO: $DestISO"
        Write-Host "[DryRun]   from: $($paths.DestIsoContent)"
        return
    }

    Write-Host "Starting CreateISO workflow..."
    Write-Verbose "Invoke-CreateISOWork: DestIsoContent='$($paths.DestIsoContent)' DestISO='$DestISO'"
    $prepJson = Join-Path $paths.DestIsoRoot "prep.json"

    # Depends on prep.json, fail if DestISO content has not been prepared
    $prepMeta = Read-JsonFile -Path $prepJson
    if (-not $prepMeta) {
        Write-Warning "$prepJson not found. Run -Prep first to prepare the destination ISO content."
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
    if (Report-Missing -Required @($paths.BIOSInDest, $paths.UEFIInDest)) {
        Write-Warning "Boot files are missing from the destination ISO content. Run -Prep first to prepare the destination ISO."
        return
    }

    $IsoVolumeLabel = "Win$($WinOS)_$($Version)_$($Arch)_KBs"
    $bootdata = "2#p0,e,b$etfs#pEF,e,b$efis"

    $oscdimgArgsLiteral = @"
-m                  # Ignore size limits
-o                  # Optimize duplicate files
-u2                 # UTF-8 filenames
-udfver102          # UDF 1.02 for max compatibility
-l$IsoVolumeLabel   # Volume label
-bootdata:$bootdata # BIOS+UEFI boot entries
"$($paths.DestIsoContent)"
"$DestISO"
"@

    Write-Host "Building ISO: $DestISO"
    Write-Verbose ("& {0} --% {1}" -f $oscdimgExe, ($oscdimgArgsLiteral -replace '\r?\n',' '))

    & $oscdimgExe --% $oscdimgArgsLiteral
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERROR: oscdimg failed to build ISO (exit $LASTEXITCODE)"
        return
    }
    Write-Host "Created ISO: $DestISO"
}

# Real work starts here
if ($Usage) {
    Show-Usage
    exit
}

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
    Write-Host "ERROR: dism.exe is required but was not found."
    Write-Host "       Install the Windows ADK or use -dism to specify its path."
    exit 1
}
Write-Host (&$LeadIn "dism" "$dismExe")
if ($oscdimgExe) {
    Write-Host (&$LeadIn "oscdimg" "$oscdimgExe")
} else {
    Write-Host (&$LeadIn "oscdimg" "not found (ISO creation unavailable; -CreateISO will fail)")
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
        Write-Host (&$LeadIn "Auto-discovered ISO" "$ISO")
    }
}

if ($ISO -and (Test-Path $ISO)) {
    $ISO = (Resolve-Path -LiteralPath $ISO).ProviderPath
    Write-Verbose (&$LeadIn "Resolved ISO path" "$ISO")
}

# ==============================
# Resolve destination ISO
# ==============================
if (-not $DestISO -and $ISO) {
    $DestISO = $ISO -replace '\.iso$', '_KBs.iso'
    Write-Verbose (&$LeadIn "Auto-derived DestISO" "$DestISO")
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
        Write-Host "`nAvailable images in $($names.InstallWim) [source: $metaSrc]:`n"
        Write-Host ("{0,6}  {1}" -f 'Index', 'Name')
        Write-Host ("{0,6}  {1}" -f '------', '----')
        foreach ($img in $allImages) { Write-Host ("{0,6}  {1}" -f $img.Index, $img.Name) }
        Write-Host ""
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

Write-Host (&$LeadIn "Auto-derived DestISO" "$DestISO")
Write-Host (&$LeadIn "Target profile" "Windows $WinOS $Version $Arch")
Write-Host (&$LeadIn "Root folder" "$Folder")
Write-Host (&$LeadIn "ISO" "$(if ($ISO) { $ISO } else { '(none)' })")
Write-Host (&$LeadIn "DestISO" "$(if ($DestISO) { $DestISO } else { '(none)' })")
Write-Host (&$LeadIn "Selected indices" "$(if ($SelectedIndices.Count -gt 0) { $SelectedIndices.Index -join ', ' } else { 'all (determined at export time)' })")
Write-Host (&$LeadIn "Mode" "$($workSwitches -join ', ')")
if ($Clean)  { Write-Host (&$LeadIn "Clean mode" "Enabled") }
if ($DryRun) { Write-Host (&$LeadIn "Dry-run mode" "Enabled") }

if ($KB) { # Only KB workflow needs HTML parsing, so we delay this until now
    # --- HtmlAgilityPack bootstrap (PS 5.x SAFE) ---------------------------------
    $HtmlAgilityPackDll = 'HtmlAgilityPack.dll'
    $hapDll = Join-Path $Folder $HtmlAgilityPackDll

    if ($Clean) {
        Clean-File $hapDll
    }
    elseif ($DryRun) {
        if (-not (Test-Path $hapDll)) {
            Write-Host "[DryRun] Would download: $HtmlAgilityPackDll"
        }   
    } else {
        if (-not (Test-Path $hapDll)) {
            Write-Host "HtmlAgilityPack.dll not found - downloading..."

            $nugetUrl   = "https://www.nuget.org/api/v2/package/HtmlAgilityPack"
            $tmpNupkg   = Join-Path $PSScriptRoot "HtmlAgilityPack.nupkg"
            $extractDir = Join-Path $PSScriptRoot "HtmlAgilityPack_Extract"

            # Clean old extraction folder if it exists
            Remove-Folder $extractDir

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
                Write-Host "ERROR: $HtmlAgilityPackDll not found inside the downloaded NuGet package. KB downloads will not work."
                return
            }

            Copy-Item -Path $sourceDll -Destination $hapDll -Force
            Write-Verbose "$HtmlAgilityPackDll copied to: $hapDll"

            # Cleanup: remove extraction folder + nupkg
            Remove-Folder $extractDir
            Remove-Item $tmpNupkg -Force
        }

        # --- Load the DLL (PS 5.x safe) ---
        # Load via byte array instead of file path so .NET does not hold a
        # file lock on the DLL. This allows -Clean to delete it.
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
        Write-SetupConfigFiles
        Write-SetupCmdFiles
    }
    if ($Prep)      { Invoke-PrepDestISO }
    if ($CreateISO) { Invoke-CreateISOWork }

    Write-Host "Completed"
} catch {
    Write-Host ""
    Write-Host "ERROR: $_"
    Write-Host "       Run the script again once the issue is resolved; completed steps will be skipped."
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
                Write-Host "Cleanup: dismounting ISO..."
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
            Write-Host "Cleanup: discarding leftover DISM mount at $($mountDir.FullName)..."
            Run-Dism @('/Unmount-Image', "`"/MountDir:$($mountDir.FullName)`"", "/Discard")
            Remove-Folder $mountDir.FullName
        }
    }
}
