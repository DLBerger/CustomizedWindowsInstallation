# Remember: variable names are CASE INSENSITIVE
<#
.SYNOPSIS
Builds a reusable update/driver/registry/script payload for Windows 10/11 installation media

.DESCRIPTION
This script prepares a directory structure that can be copied to the root of a USB drive
containing official Windows installation media (Windows 10 22H2+ or Windows 11 25H2+)

It supports:
- Downloading OS cumulative updates and .NET updates from the Microsoft Update Catalog
- Exporting all third-party drivers from the current system into $WinPEDriver$
- Exporting registry keys into .reg files
- Dry-run mode (no changes made)
- Clean mode (remove generated content)

.PARAMETER Folder
Root folder where the update/driver/registry/scripts structure will be created
If omitted, defaults to the current working directory

.PARAMETER WinOS
Windows major version: '10' or '11'
Alias: -OS
If omitted, defaults to '11'

.PARAMETER Version
Windows feature update version (for example: '22H2', '25H2')
If omitted:
- Windows 10 -> '22H2'
- Windows 11 -> '25H2'

.PARAMETER Arch
CPU architecture: 'x64' or 'arm64'
If omitted, defaults to 'x64'

.PARAMETER Extract
Mount the source ISO and extract its full content tree to <Folder>\SrcISO\Content\
Alias: -ExtractISO

.PARAMETER ExportWims
Export selected indices from SrcISO\Content\ into per-index uncompressed WIMs under Wims\Indices\
Alias: -Export

.PARAMETER KB
Download OS and .NET updates

.PARAMETER Service
Apply downloaded KBs to the exported indices Wims\Serviced\

.PARAMETER Final
Produce final install.wim and boot.wim in Wims\Final\ from the Wims in Wims\Serviced\

.PARAMETER Prep
Hardlink-copy SrcISO\Content\ to DestISO\Content\, then place the final WIMs from Wims\Final\
Alias: -PrepDestISO

.PARAMETER Files
Copy various .cmd, .ps1, and .ini files with transformations to customize them for the current folder structure and configuration

.PARAMETER CreateISO
Create the final .iso from DestISO\Content\ using oscdimg

.PARAMETER All
Shorthand for -Extract -Export -KB -Service -Final -Prep -Files -CreateISO
Default when no specific switch is provided

.PARAMETER Most
Same as -All without -CreateISO

.PARAMETER ShowIndices
Print available image indices from the source ISO (or cached metadata) and exit

.PARAMETER Home
Select editions whose normalized label matches "Home" exactly

.PARAMETER Pro
Select editions whose normalized label matches "Pro" exactly

.PARAMETER Indices
Comma-separated selector string supporting:
- numbers: 6
- ranges: 3-6, 7-*
- exact labels: "Education N"
- wildcard labels: "*Home*", "* N*"
- regex labels: "re:^Education( N)?$"

.PARAMETER ISO
Alias: -SrcISO
Explicit path to source ISO
If omitted, the script discovers the single .iso file in <Folder>
If more than one .iso is present an error is raised, use this parameter to disambiguate

.PARAMETER DestISO
Explicit path to destination ISO
If omitted, the source ISO path is reused with the extension changed to _KBs.iso

.PARAMETER DestISOVolumeLabel
Explicit volume label for the destination ISO
If omitted, the volume label is derived from the metadata of the boot.wim in the destination ISO
Example: Win11_25H2_x64_26200_6584

.PARAMETER UseADK
Prefer ADK dism.exe and oscdimg.exe when available

.PARAMETER UseSystem
Force system dism.exe and PATH oscdimg.exe

.PARAMETER dism
Explicit path to dism.exe

.PARAMETER oscdimg
Explicit path to oscdimg.exe

.PARAMETER Clean
Remove generated content instead of creating it

.PARAMETER DryRun
Show actions without performing them

.PARAMETER Help
Displays help and exits

.PARAMETER Usage
Displays help and exits

.NOTES
Fully compatible with Windows PowerShell 5.x
Never use Write-Output, always use Write-Host or it breaks having normal output in a function that returns anything
DO NOT assign directly from a foreach loop to it will lose any kind of Write-Host, Write-Verbose, or Write-Debug from inside the loop
Instead, accumulate results in a list and output the list at the end of the loop

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

[CmdletBinding()]
param( # No positional parameters as they are broken in PowerShell 5.x
    [string]$Folder = '.\',   # Default is the current working directory

    [Alias('SrcISO')]
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

    [Alias('ExportWims')]
    [switch]$Export,

    [switch]$KB,
    [switch]$Service,
    [switch]$Final,
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

    [string]$DestISOVolumeLabel,

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
$GitHash = "4946154"

# Leadin to get ':' to line up in output. Write-xxxx (&$LeadIn "dism" "$dismExe")
$LeadIn = { param($Label, $Value) '{0,-20}: {1}' -f $Label, $Value }

# Retry count for IO operations (file copy, download, etc.)
$MaxIORetries = 3

# Chunk buffers for optimal performance
$BufferSize = 64MB

# Calculate the current percentage progress bucket
$ProgressPrecentage = 10 # Report progress in 10% increments
$PercentWidth       = 3  # Always 0–100
$Bucket = { param([int64]$CopiedBytes, [int64]$TotalBytes)
            $(if   ($CopiedBytes -eq 0) { -1 }
              else                      { [math]::Floor(([math]::Floor(($CopiedBytes / $TotalBytes) * 100)) / $ProgressPrecentage) * $ProgressPrecentage })
         }

# Hex converter to string.  Write-xxxx ($(& $Hex $rc))
$Hex = { param([int]$Code) ('0x{0:X8}' -f ($Code -band 0xFFFFFFFF)) }

# Final compression type
$DismCompression = 'max' # none, fast, max

# Auto-generated DestISO?
$AutoGeneratedDestISO = $false

# Core names begin
$names = [ordered]@{
    SrcIso                 = 'SrcISO'
    DestIso                = 'DestISO'
    DefaultDestISO         = '_KBs'
    KBs                    = 'KBs'
    Wims                   = 'Wims'
    WinPEDriver            = '$WinPEDriver$'
    Registry               = 'Registry'
    Content                = 'Content'
    Sources                = 'sources'
    BootWim                = 'boot.wim'
    InstallEsd             = 'install.esd'
    InstallWim             = 'install.wim'
    WinreWim               = 'winre.wim'
    BootFileBIOS           = 'boot\etfsboot.com'
    BootFileUEFI           = 'efi\microsoft\boot\efisys.bin'
    SetupExe               = 'setup.exe'
    ExportDriversCmd       = 'ExportDrivers.cmd'
    InstallDriversCmd      = 'InstallDrivers.cmd'
    ExportRegsCmd          = 'ExportRegs.cmd'
    ExportRegsPs1          = 'ExportRegs.ps1'
    InstallRegsCmd         = 'InstallRegs.cmd'
    UpdateNETCmd           = 'Update.NET.cmd'
    UpdateNETPs1           = 'Update.NET.ps1'
    WindowsInstallationCmd = 'WindowsInstallation.cmd'
    PostSetupCmd           = 'PostSetup.cmd'
    MetadataJson           = 'wim-metadata.json'
    ManifestJson           = 'manifest.json'
    ExtractJson            = 'extract.json'
    Unknown                = 'unknown'
}

$kbDirs = @('SSU', 'OSCU', 'NET', 'MISC')
foreach ($u in $kbDirs) {
    $names[$u] = $u
}

$wimDirs = @('Indices', 'Mounts', 'Serviced', 'Final', 'Scratch', 'Logs')
foreach ($u in $wimDirs) {
    $names[$u] = $u
}
# Core names end

# Ensure elevated
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "$PSCommandPath must be run elevated as Administrator"
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

function Resolve-FullPath {
    param([string]$Path)

    # Absolute path? Normalize and return
    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    # Relative path? Resolve against the shell's working directory
    return [System.IO.Path]::GetFullPath((Join-Path $PWD.ProviderPath $Path))
}

function FolderRelName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ($Path -like "$Folder*") {
        $len = $Folder.Length

        if ($Path.Length -gt ($len + 2)) {
            if ($Path[$len] -eq '\') {
                return $Path.Substring($len + 1)
            }
        }
    }

    return $Path
}

function Protect-Token([string]$s) {
  if (-not $s) { return $names.Unknown }
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
  Write-Host "  -Home            Select editions whose normalized label matches `"Home`" exactly" -ForegroundColor Gray
  Write-Host "  -Pro             Select editions whose normalized label matches `"Pro`" exactly" -ForegroundColor Gray
  Write-Host "  -Indices <spec>  Select editions based on a comma-separated selector string" -ForegroundColor Gray
  Write-Host "     Where <spec> can include:" -ForegroundColor Gray
  Write-Host "       - numbers and ranges: 6, 3-6, 7-*" -ForegroundColor Gray
  Write-Host "       - exact labels: `"Education N`"" -ForegroundColor Gray
  Write-Host "       - wildcard labels: `"*Home*`", `"* N*`"" -ForegroundColor Gray
  Write-Host "       - regex labels: `"re:^Education( N)?$`"" -ForegroundColor Gray
  Write-Host ""
  Write-Host "Work Selections:" -ForegroundColor Cyan
  Write-Host "  -Extract    Mount the source ISO and extract its full content tree" -ForegroundColor Gray
  Write-Host "  -Export     Export selected indices from the install.wim and boot.wim" -ForegroundColor Gray
  Write-Host "  -KB         Download OS and .NET updates" -ForegroundColor Gray
  Write-Host "  -Service    Apply downloaded KBs to the exported indices and produce final install.wim and boot.wim" -ForegroundColor Gray
  Write-Host "  -Final      Produce final install.wim and boot.wim in Wims\Final\ from the Wims in Wims\Serviced\" -ForegroundColor Gray
  Write-Host "  -Prep       Prepare the destination ISO work area" -ForegroundColor Gray
  Write-Host "  -Files      Copy various .cmd and .ps1 files to the root of DestISO" -ForegroundColor Gray
  Write-Host "  -CreateISO  Create the DestISO using oscdimg" -ForegroundColor Gray
  Write-Host "  -All        Shorthand for all work selections" -ForegroundColor Gray
  Write-Host "  -Most       Shorthand for all work selections except -CreateISO" -ForegroundColor Gray
  Write-Host ""
  Write-Host "More Options:" -ForegroundColor Cyan
  Write-Host "  -ShowIndices     Print available image indices from the source ISO (or cached metadata) and exit" -ForegroundColor Gray
  Write-Host "  -DestISOVolumeLabel <label> Explicit volume label for the destination ISO" -ForegroundColor Gray
  Write-Host "     If omitted, the volume label is derived from the metadata of the boot.wim in the destination ISO." -ForegroundColor Gray
  Write-Host "     Example: Win11_25H2_x64_26200_6584" -ForegroundColor Gray
  Write-Host "  -UseADK          Prefer ADK dism.exe and oscdimg.exe when available" -ForegroundColor Gray
  Write-Host "  -UseSystem       Force system dism.exe and oscdimg.exe when available" -ForegroundColor Gray
  Write-Host "  -dism <path>     Explicit path to dism.exe" -ForegroundColor Gray
  Write-Host "  -oscdimg <path>  Explicit path to oscdimg.exe" -ForegroundColor Gray
  Write-Host "  -Clean           Overrides all other options combine with <Work Selections> to narrow the selection" -ForegroundColor Gray
  Write-Host "  -DryRun          Perform a dry run without making any changes" -ForegroundColor Gray
  Write-Host "  -Debug           Enable debug output" -ForegroundColor Gray
  Write-Host "  -Verbose         Enable verbose output" -ForegroundColor Gray
  Write-Host "  -Help            Display help information" -ForegroundColor Gray
  Write-Host "  -Usage           Display usage information" -ForegroundColor Gray
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
        $relPath = $(FolderRelName $Path)
        if ($DryRun) {
            Write-Host "[DryRun] Would remove file  : $relPath"
        } elseif (Test-Path $Path) {
            Write-Host "Removing file  : $relPath"
            Remove-Item $Path -Force -ErrorAction SilentlyContinue
        }
    }
}

function Clean-Folder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ($Clean) {
        $relPath = $(FolderRelName $Path)
        if ($DryRun) {
            Write-Host "[DryRun] Would remove folder: $relPath"
        } elseif (Test-Path $Path) {
            Write-Host "Removing folder: $relPath"
            Remove-Item $Path -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Stream-FileCopy {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [string]$SourcePath,

        [Parameter(Position = 1)]
        [string]$DestinationPath,

        [int64]$CopiedBytes = 0,
        [int64]$TotalBytes = 0,
        [switch]$ShowSourceOnly,
        [switch]$NoTitle,
        [switch]$NoProgress,
        [switch]$NoComplete
    )

    # -----------------------------
    # VALIDATION
    # -----------------------------
    if (-not (Test-Path -Path $SourcePath -PathType Leaf)) {
        throw "Stream-FileCopy source file does not exist: $SourcePath"
    }

    $destParent = Split-Path -Path $DestinationPath -Parent
    if (-not (Test-Path -Path $destParent)) {
        $null = New-Item -Path $destParent -ItemType Directory -Force
    }

    # -----------------------------
    # INITIALIZATION
    # -----------------------------
    $fileBytes = [int64](Get-Item -LiteralPath $SourcePath).Length
    if ($TotalBytes -eq 0) { $TotalBytes = $fileBytes } # current file size
    if ($TotalBytes -eq 0) { $TotalBytes = 1 }          # avoid divide-by-zero

    # Build dynamic format string for progress output, aligning byte counts to the right based on the total size
    $maxBytesWidth = $TotalBytes.ToString("N0").Length
    $fmt           = "Progress: {0,$($PercentWidth)}%  {1,$($maxBytesWidth):N0}/{2:N0} bytes"

    $lastReportedPercent = (&$Bucket $CopiedBytes $TotalBytes)

    # Progress?
    $doProgress = ((-not $NoProgress) -and                                                   # No output wanted
                   (
                     ($fileBytes -gt $BufferSize) -or                                        # Large file
                     (-not (($CopiedBytes -eq 0) -and ($fileBytes -eq $TotalBytes))) -or     # Not a single file
                     (($CopiedBytes -ne 0) -and ($CopiedBytes + $fileBytes) -ge $TotalBytes) # Last file of a series of files
                   )
                  )

    # Robocopy-like retry behavior
    $retries = $MaxIORetries
    $success = $false

    if (-not $NoTitle) { Write-Host ("Copying {0}{1}" -f (FolderRelName $SourcePath), $(if ($ShowSourceOnly) { "" } else { " to $(FolderRelName $DestinationPath)" })) }
    while (-not $success) {
        $sourceStream = $null
        $destStream   = $null
        $bytesWrittenThisAttempt = [int64]0
        $aborted = $true

        try {
            # -----------------------------
            # OPEN STREAMS
            # -----------------------------
            $sourceStream = [System.IO.File]::OpenRead($SourcePath)
            $destStream   = [System.IO.File]::Create($DestinationPath)
            $buffer       = New-Object byte[] $BufferSize

            # -----------------------------
            # STREAM COPY LOOP
            # -----------------------------
            # Initial line
            while (($bytesRead = $sourceStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                $destStream.Write($buffer, 0, $bytesRead)

                $bytesWrittenThisAttempt += $bytesRead
                $CopiedBytes             += $bytesRead

                # Progress calculation
                $percentBucket = (&$Bucket $CopiedBytes $TotalBytes)
                if ($percentBucket -gt $lastReportedPercent) {
                    if ($doProgress) { Write-Host ($fmt -f $percentBucket, $CopiedBytes, $TotalBytes) }
                    $lastReportedPercent = $percentBucket
                }
            }

            $aborted = $false
            $success = $true
        }
        catch {
            # Roll back global counter for this failed attempt
            $CopiedBytes -= $bytesWrittenThisAttempt

            if ($retries -gt 0) {
                Write-Warning ("Copy failed: {0}" -f $_.Exception.Message)
                Write-Warning ("Retrying...")
                Start-Sleep -Seconds 1
                $retries--
            }
            else {
                throw "Stream-FileCopy failed after retries"
            }
        }
        finally {
            if ($sourceStream) { $sourceStream.Close(); $sourceStream.Dispose() }
            if ($destStream)   { $destStream.Close();   $destStream.Dispose() }

            # Delete incomplete file if aborted
            if ($aborted -and (Test-Path -Path $DestinationPath)) {
                Remove-Item -Path $DestinationPath -Force -ErrorAction SilentlyContinue
            }
        }
    }
    if (-not $NoComplete) { Write-Host "Copy complete" }
}

function Stream-FolderCopy {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [string]$SourcePath,

        [Parameter(Position = 1)]
        [string]$DestinationPath,

        [switch]$ShowSourceOnly,
        [switch]$NoTitle,
        [switch]$NoProgress,
        [switch]$NoComplete
    )

    if (-not (Test-Path -Path $SourcePath)) {
        throw "Stream-FolderCopy source folder does not exist: $SourcePath"
    }
    try {
        if (-not $NoTitle) { Write-Host ("Copying {0}{1}" -f (FolderRelName $SourcePath), $(if ($ShowSourceOnly) { "" } else { " to $(FolderRelName $DestinationPath)" })) }

        # Pre-gather items and calculate total payload size using 64-bit integers
        $allItems = Get-ChildItem -Path $SourcePath -Recurse
        $TotalBytes = [int64]0
        foreach ($item in $allItems) {
            if (-not $item.PSIsContainer) { $TotalBytes += $item.Length }
        }
        if ($TotalBytes -eq 0) { $TotalBytes = 1 } # Avoid divide-by-zero on empty sources
        $CopiedBytes = [int64]0

        foreach ($item in $allItems) {
            # Isolate the relative path (e.g., 'boot\bcd' instead of 'D:\boot\bcd')
            $relativePath = $item.FullName.Substring($SourcePath.Length)
            $targetPath   = Join-Path -Path $DestinationPath -ChildPath $relativePath

            if ($item.PSIsContainer) {
                # Outputs to Stream 4 (Verbose). Captured by *>&1
                Write-Verbose "Folder: $relativePath"
                if (-not (Test-Path -Path $targetPath)) {
                    $null = New-Item -Path $targetPath -ItemType Directory -Force
                }
            } else {
                # Outputs to Stream 4 (Verbose). Captured by *>&1
                Write-Verbose "File:   $relativePath"
                Stream-FileCopy -SourcePath $item.FullName -DestinationPath $targetPath -CopiedBytes $CopiedBytes -TotalBytes $TotalBytes `
                                -ShowSourceOnly:$ShowSourceOnly -NoTitle:$NoTitle -NoProgress:$NoProgress -NoComplete
                $CopiedBytes += $item.Length
            }
        }
        if (-not $NoComplete) { Write-Host "Copy complete" }
    } catch {
        Remove-Folder $DestinationPath
        throw "ERROR during folder copy: $_"
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

        # StreamReader.ReadLine() treats carriage returns (\r) as line breaks
        # This is ideal because DISM uses \r to overwrite its text progress bar in place
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
        Write-Debug "$($process.StartInfo.FileName) $($process.StartInfo.Arguments) exited with code $(& $Hex $rc)"

        # If capturing was requested, output the entire text array down Stream 1 - SPECIAL CASE
        if ($Capture) {
            Write-Output $outputCollection.ToArray()
        }

<# This would be nice but we will never see this currently as it's a PS 5 limitation
        if ($rc -ne 0) {
            # Propagate the Ctrl-C exit code as a special case
            $hex = $(& $Hex $rc)
            if ($hex -eq '0xC000013A') {
                $msg = "$dismExe was terminated by Ctrl-C (Exit Code: $hex)"
                Write-Host $msg
                throw $msg
            }
        }
 #>

        # Callers that need the exit status check $LASTEXITCODE directly
        $global:LASTEXITCODE = $rc
    } finally {
        # CRITICAL CLEANUP: If the user forces a hard halt using Ctrl+C, 
        # this block intercepts the abort and violently closes the dism.exe thread
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
    # Locate an ADK tool (e.g., dism.exe, oscdimg.exe) using this priority:
    #   1. Explicit path supplied by the caller
    #   2. Windows ADK installation (preferred when -PreferADK or -UseADK)
    #   3. System32 / PATH
    # Returns the full path on success, $null on failure
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ToolName,       # filename, e.g., 'dism.exe'

        [Parameter(Mandatory)]
        [string]$ADKSubfolder,   # subfolder under each arch dir, e.g., 'DISM' or 'Oscdimg'

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

    Write-Warning "$ToolName not found, install Windows ADK or specify the path explicitly"
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
        Write-Warning "DISM /Get-WimInfo failed (exit $(& $Hex $LASTEXITCODE)) for: $WimPath"
    }

    # --- Call 2: OS details from index 1 ---
    $buildNumber      = 0
    $servicePackBuild = 0
    $archStr          = 'x64'
    ($detailOutput    = Run-Dism @("/Get-WimInfo", "`"/WimFile:$WimPath`"", "/Index:1") -Capture)
    $count = 0
    foreach ($line in $detailOutput) {
        $count++
        if ($count -le 3) { continue } # skip the first 3 lines
        if ($line -match '^\s*Version\s*:\s*\d+\.\d+\.(\d+)') { $buildNumber      = [int]$Matches[1] }
        if ($line -match '^\s*ServicePack Build\s*:\s*(\d+)') { $servicePackBuild = [int]$Matches[1] }
        if ($line -match '^\s*Architecture\s*:\s*(.+)')       { $archStr          = $Matches[1].Trim() }
    }
    Write-Debug "  build=$buildNumber servicePackBuild=$servicePackBuild arch=$archStr"

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
        Images           = $images.ToArray()
        WinOS            = $detectedWinOS
        Version          = $detectedVersion
        Arch             = $detectedArch
        Build            = $buildNumber
        ServicePackBuild = $servicePackBuild
    }
}

# ==============================
# Index selection
# ==============================

function Resolve-IndexSelection {
    Write-Debug "Resolve-IndexSelection: Home=$SelectHome Pro=$SelectPro Indices='$Indices' TotalImages=$($InstallImages.Count)"

    function Get-NormalizedLabel([string]$name) {
        ($name -replace '^Windows\s+(10|11)\s+', '').Trim()
    }

    $anyExplicit = $SelectHome -or $SelectPro -or $Indices

    if (-not $anyExplicit) {
        Write-Verbose "No explicit index selection, returning all $($InstallImages.Count) indices"
        return $InstallImages
    }

    $selected = [System.Collections.Generic.List[object]]::new()

    if ($SelectHome) {
        Write-Verbose "Selecting 'Home' editions"
        $InstallImages | Where-Object { (Get-NormalizedLabel $_.Name) -eq 'Home' } | ForEach-Object { $selected.Add($_) }
    }

    if ($SelectPro) {
        Write-Verbose "Selecting 'Pro' editions"
        $InstallImages | Where-Object { (Get-NormalizedLabel $_.Name) -eq 'Pro' } | ForEach-Object { $selected.Add($_) }
    }

    if ($Indices) {
        $tokens = $Indices -split '\s*,\s*'
        foreach ($token in $tokens) {
            $token = $token.Trim().Trim('"').Trim("'")
            Write-Verbose "  Processing token: '$token'"

            if ($token -match '^(\d+)-(\*|\d+)$') {
                $from = [int]$Matches[1]
                $to   = if ($Matches[2] -eq '*') { [int]::MaxValue } else { [int]$Matches[2] }
                Write-Debug "    Range $from-$to"
                $InstallImages | Where-Object { $_.Index -ge $from -and $_.Index -le $to } | ForEach-Object { $selected.Add($_) }
            }
            elseif ($token -match '^\d+$') {
                Write-Debug "    Single index $token"
                $InstallImages | Where-Object { $_.Index -eq [int]$token } | ForEach-Object { $selected.Add($_) }
            }
            elseif ($token -match '^re:(.+)$') {
                $pattern = $Matches[1]
                Write-Debug "    Regex '$pattern'"
                $InstallImages | Where-Object { $_.Name -match $pattern } | ForEach-Object { $selected.Add($_) }
            }
            elseif ($token -match '[*?]') {
                Write-Debug "    Wildcard '$token'"
                $InstallImages | Where-Object { $_.Name -like $token } | ForEach-Object { $selected.Add($_) }
            }
            else {
                Write-Debug "    Exact label '$token'"
                $InstallImages | Where-Object { $_.Name -eq $token } | ForEach-Object { $selected.Add($_) }
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
        Write-Host "[DryRun] Would export $($InstalledIndices.Count) indices to $($paths.WimsIndices)"
        return
    }

    Write-Host "Starting ExtractISO workflow..."
    Write-Verbose "Invoke-ExtractISO: ISO='$ISO' SrcIsoContent='$($paths.SrcIsoContent)'"

    if (-not $ISO -or -not (Test-Path $ISO)) {
        Write-Warning "Source ISO not found or not specified, use -ISO to point to your Windows .iso file"
        return
    }

    # Checkpoint: skip if same ISO was already extracted, clean and re-extract if ISO changed
    $extractJson  = $paths.ExtractJson
    $existingJson = Read-JsonFile -Path $extractJson
    if ($existingJson) {
        if ($existingJson.ISOPath -eq $ISO) {
            Write-Host  "ExtractISO already done for this ISO ($extractJson) matches)"
            Write-Debug "$($extractJson): ISOPath='$($existingJson.ISOPath)' Date='$($existingJson.Date)'"
            return
        }
        Write-Host "ISO path changed (was '$($existingJson.ISOPath)'), cleaning SrcIsoRoot and re-extracting..."
        Remove-Folder $paths.SrcIsoRoot
    }

    Ensure-Folder $paths.SrcIsoContent

    Write-Host "Mounting ISO: $ISO"
    $diskImage = Mount-DiskImage -ImagePath $ISO -PassThru -ErrorAction SilentlyContinue
    try {
        # Wait for the volume to actually appear
        $vol = $null
        for ($r = 0; $r -lt 5 -and $null -eq $vol; $r++) {
            $vol = $diskImage | Get-Volume -ErrorAction SilentlyContinue
            if ($null -eq $vol) { Start-Sleep -Seconds 1 }
        }
        if ($null -eq $vol) { throw "Timeout waiting for ISO volume to initialize" }

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
            throw "Source ISO validation failed, see above for missing file details"
        }
        Write-Host "Source ISO validation passed"

        # Copy the ISO tree to SrcIsoContent

        # Standardize paths to ensure safe string replacement for relative paths
        $sourceBase = $driveLetterRaw.TrimEnd('\') + '\'
        $destBase   = $paths.SrcIsoContent

        # Copy it and Stream-FolderCopy will give you a running commentary
        Stream-FolderCopy $sourceBase $destBase -ShowSourceOnly
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
    Write-Host "ExtractISO complete ($extractJson written)"
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
        Write-Host "[DryRun] Would export $($names.InstallWim) $($InstalledIndices.Count) indices to $($paths.WimsIndices)"
        Write-Host "[DryRun] Would export $($names.BootWim) $($BootImages.Count) indices to $($paths.WimsIndices)"
        return
    }

    Write-Host "Starting Export workflow..."
    Write-Verbose "Invoke-Export: SourcesInSrc='$($paths.SourcesInSrc)' WimsIndices='$($paths.WimsIndices)'"

    Ensure-Folder $paths.WimsRoot
    Ensure-Folder $paths.WimsIndices

    $extractJson = $paths.ExtractJson
    $extractMeta = Read-JsonFile -Path $extractJson
    if (-not $extractMeta) {
        Write-Warning "$extractJson not found, run -Extract first"
        return
    }
    $extractDate = [datetime]::Parse($extractMeta.Date)

    # Use MetadataJson if nothing has changed
    $metadataJson = $paths.MetadataJson
    $wimMeta      = Read-JsonFile -Path $metadataJson
    if (-not $wimMeta -or ($wimMeta.ISOPath -ne $ISO)) {
        # Locate source WIMs
        $installSrc = if     (Test-Path $paths.InstallWimInSrc) { $paths.InstallWimInSrc }
                      elseif (Test-Path $paths.InstallEsdInSrc) { $paths.InstallEsdInSrc }
                      else   { $null }
        $bootSrc    = if     (Test-Path $paths.BootWimInSrc) { $paths.BootWimInSrc }
                      else   { $null }

        if (-not $installSrc) {
            Write-Warning "Install image not found in $($paths.SourcesInSrc), run -Extract first"
            return
        }
        if (-not $bootSrc) {
            Write-Warning "Boot image not found in $($paths.SourcesInSrc), run -Extract first"
            return
        }

        Write-Verbose "install source: $installSrc"
        Write-Verbose "boot source   : $bootSrc"

        # Collect WIM metadata and write MetadataJson
        Write-Host "Collecting WIM metadata..."
        $installMeta = Get-WimMetadata -WimPath $installSrc
        $bootMeta    = Get-WimMetadata -WimPath $bootSrc

        # Reset our globals
        $InstallImages = $installMeta.Images
        $BootImages    = $bootMeta.Images

        Write-JsonFile -Path $MetadataJson -Data @{
            ISOPath          = $extractMeta.ISOPath
            CollectedDate    = (Get-Date -Format s)
            WinOS            = $installMeta.WinOS
            Version          = $installMeta.Version
            Arch             = $installMeta.Arch
            Build            = $installMeta.Build
            ServicePackBuild = $installMeta.ServicePackBuild
            InstallImages    = @($InstallImages | ForEach-Object { @{ Index = $_.Index; Name = $_.Name } })
            BootImages       = @($BootImages    | ForEach-Object { @{ Index = $_.Index; Name = $_.Name } })
        }
        Write-Host "WIM metadata saved ($($InstallImages.Count) install image(s), $($BootImages.Count) boot image(s))"
    } else {
        Write-Host "Using existing WIM metadata..."

        # Reset our globals
        $InstallImages = $wimMeta.InstallImages
        $BootImages    = $wimMeta.BootImages
    }

    # Resolve anything that hasn't been
    $InstallIndices = Resolve-IndexSelection

    # Give a recap of what we found and what we're about to export
    Write-Host "$($names.InstallWim) all indices: $($InstallImages.Index -join ', ') indices to export: $($InstallIndices.Index -join ', ')"
    Write-Host "$($names.BootWim)    all indices: $($BootImages.Index -join ', ')"

    function Export-WimImage {
        param(
            [PSCustomObject] $img,     # object with Index and Name properties
            [string]         $WimName, # install or boot
            [string]         $Src      # source WIM file path
        )
        $idx     = $img.Index
        $imgName = $img.Name
        Write-Host "  [Index $idx] $imgName"

        $installDest = Join-Path $paths.WimsIndices ("{0}_{1}"      -f $idx, $WimName)
        $installJson = Join-Path $paths.WimsIndices ("{0}_{1}.json" -f $idx, $WimName)
        $existInstall = Read-JsonFile -Path $installJson
        $needInstall  = (-not $existInstall) -or ([datetime]::Parse($existInstall.ExportDate) -le $extractDate)

        if (-not $needInstall) {
            Write-Host "    $WimName index $idx already exported ($($existInstall.ExportDate))"
        } else {
            Write-Host "    Exporting $WimName index $idx..."
            Run-Dism @("/Export-Image", "`"/SourceImageFile:$Src`"", "`"/SourceIndex:$idx`"",
                       "`"/DestinationImageFile:$installDest`"", "/Compress:None", "/CheckIntegrity") -Indent 4
            $rc = $LASTEXITCODE
            if ($rc -ne 0) {
                Write-Warning "    DISM export failed for $WimName index $idx (exit $(& $Hex $rc))"
                Write-Warning "    Try running this script again, skipping export for this index for now"
                return
            }
            Write-JsonFile -Path $installJson -Data @{ Index = $idx; Name = $imgName; ExportDate = (Get-Date -Format s) }
            Write-Host "    $WimName index $idx exported"
        }
    }

    # Export each selected install image
    foreach ($img in $InstallIndices) {
        Export-WimImage -img $img -WimName $names.InstallWim -Src $installSrc
    }
    # Export each boot image
    foreach ($img in $BootImages) {
        Export-WimImage -img $img -WimName $names.BootWim -Src $bootSrc
    }

    Write-Host "Export workflow complete"
}

# =========================
# KBs section
# =========================

# =========================
# HTML-based Update Catalog search
# =========================

<# Copilot's HTMHtmlAgilityPack.HtmlDocument free suggestions

1. Make Invoke-CatalogRequest return raw HTML (string)
$Response = Invoke-WebRequest @Params

Write-Debug "RawContent length = $($Response.RawContent.Length)"

$content = $Response.Content
Write-Debug "Content length = $($content.Length)"
return $content


2. Update Search-UpdateCatalogHtml to use the raw HTML string
$html = Invoke-CatalogRequest -Uri $Uri
if (-not $html) {
    Write-Warning "No HTML returned from $Uri"
    return
}

Write-Verbose "Extracting update IDs from HTML"

$pattern = 'goToDetails\("([0-9A-Fa-f\-]{36})"\)'
$matches = [regex]::Matches($html, $pattern)

3. Add a helper: Get-CatalogSupersededBy (tight, aligned to your HTML)
function Get-CatalogSupersededBy {
    param([string]$Html)

    $section = ($Html -split 'id="supersededbyInfo"')[1]
    if (-not $section) { return @() }

    $pattern = '<a[^>]*>([^<]+)</a>'

    $matches = [regex]::Matches(
        $section,
        $pattern,
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    return $matches |
        ForEach-Object { $_.Groups[1].Value.Trim() } |
        Where-Object { $_ -ne "" }
}

4. Convert Get-UpdateDetails to pure regex (no HtmlAgilityPack)
$detailsResponse = Invoke-WebRequest -Uri $detailsUrl -UseBasicParsing -ErrorAction Stop
$html = $detailsResponse.Content

# Title
$title = ""
$titleMatch = [regex]::Match(
    $html,
    '<span[^>]*id="ScopedViewHandler_titleText"[^>]*>(.*?)</span>',
    [System.Text.RegularExpressions.RegexOptions]::Singleline
)
if ($titleMatch.Success) {
    $title = $titleMatch.Groups[1].Value.Trim()
}
Write-Verbose ("Title: {0}" -f $title)

# KB
$kb = ""
$kbMatch = [regex]::Match($title, "KB\d+")
if ($kbMatch.Success) { $kb = $kbMatch.Value }
Write-Verbose ("KB: {0}" -f $kb)

# SupersededBy
$supersededBy = Get-CatalogSupersededBy $html

#>

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
        [int]$MaxResults,

        [Parameter(Mandatory)]
        [string]$KBIndex
    )

    Write-Verbose ("Searching for {0}{1}..." -f $Query, $(if ($MaxResults -ne 0) { " (max results: $MaxResults)" } else { "" }))

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
    $pattern   = 'goToDetails\("([0-9A-Fa-f\-]{36})"\)'
    $matchList = [regex]::Matches($Html.DocumentNode.InnerHtml, $pattern)
    $ids = @()
    $count = 0
    foreach ($m in $matchList) {
        $id = $m.Groups[1].Value
        Write-Debug "Found update GUID: $id"
        $ids += [PSCustomObject]@{
            Guid    = $id
            KBIndex = $KBIndex
        }
        $count++
        if (($MaxResults -ne 0) -and ($count -ge $MaxResults)) { break }
    }

    Write-Verbose "Total IDs extracted: $($ids.Count)"
    if ($ids.Count -ne 0) {
        Write-Debug  ('{0,3} {1,-36} KBIndex:' -f "ID:", "Guid:")
        for ($i = 0; $i -lt $ids.Count; $i++) {
            Write-Debug ('{0,3} {1,-36} {2}' -f $($i + 1), $($ids[$i].Guid), $($ids[$i].KBIndex))
        }
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

    $matchList = [regex]::Matches(
        $content,
        $pattern,
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    if ($matchList.Count -eq 0) {
        Write-Warning "No downloadInformation URL matches for $Guid (regex returned 0 matches)"
        return @()
    }

    Write-Verbose "Found $($matchList.Count) download link match(es)"

    $links = @()
    try {
        foreach ($m in $matchList) {
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

    $manifestPath = Join-Path $Folder $names.ManifestJson
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

    $manifestPath = Join-Path $Folder $names.ManifestJson
    $json = $Entries | ConvertTo-Json -Depth 6
    $json | Set-Content -Path $manifestPath -Encoding UTF8
}

function Download-MUFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject] $Update,

        [Parameter(Mandatory = $true)]
        [string] $KBIndex
    )

    Write-Host ("Preparing downloads for update {0}: {1}" -f $Update.Guid, $Update.Title)

    $TargetFolder = $paths["KBs$KBIndex"]
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
        # Retry loop
        # ------------------------------------------------------------
        $success = $false

        for ($attempt = 1; $attempt -le $MaxIORetries; $attempt++) {

            Write-Host ("  Attempt {0} of {1}" -f $attempt, $MaxIORetries)

            try {
                $req = [System.Net.HttpWebRequest]::Create($url)
                $req.Method = "GET"
                $req.UserAgent = "Mozilla/5.0"

                $resp = $req.GetResponse()
                $total = $resp.ContentLength
                $inStream  = $resp.GetResponseStream()
                $outStream = [System.IO.File]::Open($destPath, [System.IO.FileMode]::Create)

                $buffer = New-Object byte[] $BufferSize
                $totalRead = 0
                $nextMark = $ProgressPercentage

                # Compute max widths
                $maxBytesWidth = $total.ToString("N0").Length

                # Build dynamic format string
                $fmt = " {0,$($PercentWidth)}%  {1,$($maxBytesWidth):N0}/{2:N0} bytes"

                # Initial progress line
                Write-Host ($fmt -f 0, 0, $total)

                while (($read = $inStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                    $outStream.Write($buffer, 0, $read)
                    $totalRead += $read

                    if ($total -gt 0) {
                        $pct = [math]::Floor(($totalRead / $total) * 100)

                        if ($pct -ge $nextMark) {
                            Write-Host ($fmt -f $pct, $totalRead, $total)
                            $nextMark += $ProgressPrecentage
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
            Write-Warning ("FAILED after {0} attempts: {1}" -f $MaxIORetries, $fileName)
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
        [string] $KBIndex
    )

    Write-Host ("Processing update #{0}: {1}" -f $Count, $Guid)
    Write-Verbose ("KBIndex: {0}" -f $KBIndex)
    $TargetFolder = $paths["KBs$KBIndex"]
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
    $supersededByPreview = @()
    $supNodes = $detailsDoc.DocumentNode.SelectNodes("//div[@id='supersededbyInfo']//a")
    if ($supNodes) {
        foreach ($n in $supNodes) {
            $supersededByKB = $n.InnerText.Trim()
            if ($supersededByKB -imatch 'preview') {
                $supersededByPreview += $supersededByKB
            } else {
                $supersededBy += $supersededByKB
            }
        }
    }
    Write-Verbose ("SupersededBy: {0}" -f ($supersededBy -join ', '))
    Write-Verbose ("SupersededByPreview: {0}" -f ($supersededByPreview -join ', '))

    # Not a keeper if superseded by anything else, even if it has download links
    if ($supersededBy.Count -gt 0) {
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
        KBIndex      = $KBIndex
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
            Query        = "Critical Updates Windows $WinOS $Version $Arch"
            MaxResults   = 0
            KBIndex      = 'SSU'
        }
        [PSCustomObject]@{
            Query        = "Security Updates Windows $WinOS $Version $Arch"
            MaxResults   = 0
            KBIndex      = 'OSCU'
        }
        [PSCustomObject]@{
            Query        = ".NET Security Updates Windows $WinOS $Version $Arch"
            MaxResults   = 0
            KBIndex      = 'NET'
        }
        [PSCustomObject]@{
            Query        = ".NET 8.0 $Arch Client"
            MaxResults   = 0
            KBIndex      = 'NET'
        }
        [PSCustomObject]@{
            Query        = ".NET 9.0 $Arch Client"
            MaxResults   = 0
            KBIndex      = 'NET'
        }
        [PSCustomObject]@{
            Query        = ".NET 10.0 $Arch Client"
            MaxResults   = 0
            KBIndex      = 'NET'
        }
        [PSCustomObject]@{
            Query        = ".NET 11.0 $Arch Client"
            MaxResults   = 0
            KBIndex      = 'NET'
        }
        [PSCustomObject]@{
            Query        = ".NET 12.0 $Arch Client"
            MaxResults   = 0
            KBIndex      = 'NET'
        }
        [PSCustomObject]@{
            Query        = "Definition Updates Windows Security platform"
            MaxResults   = 0
            KBIndex      = 'MISC'
        }
    )

    $results = @()
    foreach ($q in $queries) {
        $results += Search-UpdateCatalogHtml -Query $q.Query -MaxResults $q.MaxResults -KBIndex $q.KBIndex
    }

    $allGuids = $results | Sort-Object Guid -Unique
    Write-Host ("Found {0} total updates to process" -f $allGuids.Count)
    Write-Debug  ('{0,3} {1,-36} KBIndex:' -f "ID:", "Guid:")
    for ($i = 0; $i -lt $allGuids.Count; $i++) {
        Write-Debug ('{0,3} {1,-36} {2}' -f $($i + 1), $($allGuids[$i].Guid), $($allGuids[$i].KBIndex))
    }

    if ($allGuids.Count -eq 0) {
        Write-Host "No updates found"
        return
    }

    Write-Host "Retrieving update details..."

    $count   = 0
    $details = @()
    foreach ($g in $allGuids) {
        Write-Debug ("Resolving details for {0} ({1})" -f $g.Guid, $g.KBIndex)
        # 1. Run the function in a pipeline so Write-Host is visible
        $detail = Get-UpdateDetails -Count (++$count) -Guid $g.Guid -KBIndex $g.KBIndex | ForEach-Object { $_ }

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
        Write-Debug ("Update: {0}`n  Title: {1}`n  KB: {2}`n  KBIndex: {3}`n  URLs: {4}" -f $d.Guid, $d.Title, $d.KB, $d.KBIndex, ($d.DownloadUrls -join ', '))
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
            $existingFiles = Get-ChildItem -Path $folder -File | Select-Object -ExpandProperty Name
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
        $manifestByFolder[$d.KBIndex] = @()
    }
    foreach ($d in $details) {
        $targetFolder = $paths["KBs$d.KBIndex"]

        # Download all files for this update into the target folder
        $downloadInfos = Download-MUFile -Update $d -KBIndex $d.KBIndex

        foreach ($downloadInfo in $downloadInfos) {
            $entry = Build-ManifestEntry -Details $d -DownloadInfo $downloadInfo
            $manifestByFolder[$d.KBIndex] += $entry
        }
    }

    Write-Host "Writing manifests..."
    foreach ($kvp in $manifestByFolder.GetEnumerator()) {
        $folder  = $paths["KBs$($kvp.Key)"]
        $entries = $kvp.Value
        if ($entries.Count -gt 0) {
            Write-Verbose "Writing manifest for $folder"
            Write-Manifest -Folder $folder -Entries $entries
        }
        else {
            $manifestPath = Join-Path $folder $names.ManifestJson
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

function Test-MissingIndices {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $RunNext
    )
<#
    $m = ""
    if (-not (Test-Path $paths.WimsIndices)) {
        $m = "No indices folder found at $($paths.WimsIndices)"
    } else {
        $hasInstallIndices = ($InstallIndices -and $InstallIndices.Count -gt 0)
        $hasBootImages = ($BootImages -and $BootImages.Count -gt 0)
        if (-not ($hasInstallIndices -and $hasBootImages)) {
            $s = ""
            if (-not $hasInstallIndices) { $s += "InstallIndices" }
            if (-not $hasInstallIndices -and -not $hasBootImages) { $s += " and " }
            if (-not $hasBootImages)     { $s += "BootImages" }
            $m = "No $s defined"
        }
    }

    # Anyone set an error message?
    if (-not [string]::IsNullOrWhiteSpace($m)) {
        # Something is missing we're done
        throw ("{0}, re-run with -$RunNext" -f $m)
    }
#>
}

function Invoke-ServiceWork {
    [CmdletBinding()]
    param()

    if ($Clean) {
        Clean-Folder $paths.WimsLogs
        Clean-Folder $paths.WimsMounts
        Clean-Folder $paths.WimsScratch
        Clean-Folder $paths.WimsServiced
        return
    }

    if ($DryRun) {
        Write-Host "[DryRun] Would service extracted indices in $($paths.WimsIndices)"
        Write-Host "[DryRun] Would apply SSU packages from : $($paths.KBsSSU)"
        Write-Host "[DryRun] Would apply LCU packages from : $($paths.KBsOSCU)"
        Write-Host "[DryRun] Would service winre.wim inside each index's install.wim"
        return
    }

    Write-Host  "Starting Service workflow..."

    # Sanity check our indices first
    Test-MissingIndices "Service"
    # Resolve anything that hasn't been
    $InstallIndices = Resolve-IndexSelection

    Write-Verbose "Invoke-ServiceWork: WimsIndices='$($paths.WimsIndices)' WimsFinal='$($paths.WimsFinal)'"
    Write-Debug   "Invoke-ServiceWork: KBsSSU='$($paths.KBsSSU)' KBsOSCU='$($paths.KBsOSCU)' WimsMounts='$($paths.WimsMounts)'"

    Remove-Folder $paths.WimsLogs
    Remove-Folder $paths.WimsMounts
    Remove-Folder $paths.WimsScratch
    Remove-Folder $paths.WimsServiced
    Ensure-Folder $paths.WimsLogs
    Ensure-Folder $paths.WimsMounts
    Ensure-Folder $paths.WimsScratch
    Ensure-Folder $paths.WimsServiced

    $scratchDir = $paths.WimsScratch

    # Gather available packages (.msu and .cab)
    $ssuFiles = @(
        Get-ChildItem -Path $paths.KBsSSU -Include '*.msu', '*.cab' -Recurse -ErrorAction SilentlyContinue
    )
    $lcuFiles = @(
        Get-ChildItem -Path $paths.KBsOSCU -Include '*.msu', '*.cab' -Recurse -ErrorAction SilentlyContinue
    )

    Write-Host  "Packages available - SSU: ($($ssuFiles.Count) files), LCU: ($($lcuFiles.Count) files)"
    Write-Verbose "SSU : $($ssuFiles.Name -join ', ')"
    Write-Verbose "LCU : $($lcuFiles.Name -join ', ')"

    # -----------------------------------------------------------------------
    # Internal helper: Add-Packages
    # -----------------------------------------------------------------------
    function Add-Packages {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string] $IdxName,

            [Parameter(Mandatory = $true)]
            [string] $MountDir,

            [Parameter(Mandatory = $true)]
            [string] $Src,

            [Parameter(Mandatory = $true)]
            [AllowNull()]
            [AllowEmptyString()]
            [string] $Dest,

            [Parameter(Mandatory = $true)]
            [bool]   $Mount,

            [Parameter(Mandatory = $true)]
            [bool]   $Unmount,

            [Parameter(Mandatory = $true)]
            [string] $PkgName,

            [Parameter(Mandatory = $true)]
            [AllowNull()]
            [AllowEmptyCollection()]
            [System.IO.FileInfo[]] $Pkgs
        )

        Write-Verbose "Add-Packages: IdxName='$IdxName' MountDir='$MountDir' Src='$Src' Dest='$Dest' Mount=$Mount Unmount=$Unmount PkgName='$PkgName' Pkgs=$($Pkgs.Count)"

        if ([string]::IsNullOrWhiteSpace($Src) -or -not (Test-Path $Src)) {
            throw "Add-Packages: Source '$Src' is required but missing or empty"
        }

        # This copy is key as we need to get the file into Servicing was we don't corrupt the original extracted file in case we need to retry servicing after a failure
        # If Dest is empty or the same as Src, we'll just service in-place
        $effectiveDest = $Dest
        if ([string]::IsNullOrWhiteSpace($Dest)) {
            $effectiveDest = $Src
            Write-Verbose "Add-Packages: Dest empty, using Src as Dest ('$effectiveDest')"
        } elseif ($Src -ne $Dest) {
            # Stream-FileCopy will tell us what it's doing
            Stream-FileCopy -SourcePath $Src -DestinationPath $Dest
            $effectiveDest = $Dest
        }

        # Collect for debugging/logging
        $relMountDir      = FolderRelName $MountDir
        $relEffectiveDest = FolderRelName $effectiveDest

        # We can bail early if there are no packages to apply if there are no packages to apply on we're mounting and dismounting
        $hasPkgs = $Pkgs -and $Pkgs.Count -gt 0
        if (-not $hasPkgs -and $Mount -and $Unmount) {
            Write-Verbose "Add-Packages: No $PkgName files to apply and mounting/unmounting so nothing to do"
            return
        }

        try {
            if ($Mount) {
                Ensure-Folder $MountDir
                $leaf     = Split-Path -Path $effectiveDest -Leaf
                $safeLeaf = Protect-Token $leaf
                $mountLog = Join-Path $paths.WimsLogs ("mount_{0}_{1}.log" -f $idxName, $safeLeaf)

                # Clean the scratch directory to be safe
                Remove-Folder $scratchDir
                Ensure-Folder $scratchDir

                Write-Host "  Mounting $relEffectiveDest -> $relMountDir"
                Run-Dism @(
                    "/Mount-Image",
                    "`"/ImageFile:$effectiveDest`"",
                    "/Index:1",
                    "`"/MountDir:$MountDir`"",
                    "`"/ScratchDir:$scratchDir`"",
                    "`"/LogPath:$mountLog`""
                ) -Indent 2
                if ($LASTEXITCODE -ne 0) {
                    throw "DISM mount failed for '$relEffectiveDest' (exit $(& $Hex $LASTEXITCODE))"
                }
            }

            if ($hasPkgs) {
                foreach ($pkg in $Pkgs) {
                    $leaf   = Split-Path $pkg.FullName -Leaf
                    $safe   = Protect-Token $leaf
                    $pkgLog = Join-Path $paths.WimsLogs ("pkg_{0}_{1}_{2}.log" -f $idxName, (Protect-Token (Split-Path $MountDir -Leaf)), $safe)

                    # Clean the scratch directory before each package application to avoid DISM errors
                    Remove-Folder $scratchDir
                    Ensure-Folder $scratchDir

                    Write-Host "  Applying $PkgName file to $($idxName): $($pkg.Name)"
                    Run-Dism @(
                        "/Add-Package",
                        "`"/Image:$MountDir`"",
                        "`"/PackagePath:$($pkg.FullName)`"",
                        "`"/ScratchDir:$scratchDir`"",
                        "`"/LogPath:$pkgLog`""
                    ) -Indent 2
                    if ($LASTEXITCODE -ne 0) {
                        Write-Warning "  Add-Package failed for $($pkg.Name) on $idxName (exit $(& $Hex $LASTEXITCODE)) - continuing"
                    }
                }
            } else {
                Write-Host "  No $PkgName files supplied for $idxName so skipping package application"
            }

            if ($Unmount) {
                Write-Host "  Unmounting $idxName at $relMountDir (commit)..."
                Run-Dism @(
                    "/Unmount-Image",
                    "`"/MountDir:$MountDir`"",
                    "/Commit"
                ) -Indent 2
                if ($LASTEXITCODE -ne 0) {
                    throw "DISM unmount (commit) failed for '$relMountDir' (exit $(& $Hex $LASTEXITCODE))"
                }
            }
        }
        catch {
            Write-Host "  ERROR in Add-Packages for $($idxName)"
            if (Test-Path $MountDir) {
                Write-Host "  Discarding mounted image at $($MountDir)..."
                Run-Dism @(
                    "/Unmount-Image",
                    "`"/MountDir:$MountDir`"",
                    "/Discard"
                ) -Indent 2
            }
            throw
        }
    }

    # -----------------------------------------------------------------------
    # Internal helper: Service-Index
    # -----------------------------------------------------------------------
    function Service-Index {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]        $WimName,

            [Parameter(Mandatory = $true)]
            [PSCustomObject]$Img,

            [Parameter(Mandatory = $true)]
            [bool]          $InstallType
        )

        $idx     = $Img.Index
        $idxName = ("{0}_{1}" -f $idx, $WimName)
        $imgName = $Img.Name

        Write-Host    "  [Index $idx] $imgName"
        Write-Verbose "Service-Index: WimName='$WimName' Index=$idx IdxName='$IdxName' Name='$imgName' InstallType=$InstallType"

        $imgJson   = Join-Path $paths.WimsServiced ("{0}.json" -f $idxName)
        $imgExists = Read-JsonFile -Path $imgJson

        $imgNeeded = $true
        if ($imgExists -and $imgExists.ServicedDate) {
            try {
                $lastServiced = [datetime]::Parse($imgExists.ServicedDate)
                $imgNeeded    = $lastServiced -le $extractDate
            } catch {
                Write-Verbose "Service-Index: Failed to parse ServicedDate for $WimName index $($idx), will re-service"
                $imgNeeded = $true
            }
        }

        if (-not $imgNeeded) {
            Write-Host "    $WimName index $idx already serviced ($($imgExists.ServicedDate))"
            return
        }

        $mountDir = Join-Path $paths.WimsMounts   ("mount_{0}" -f $idxName)
        $wimPath  = Join-Path $paths.WimsIndices  $idxName
        $srvPath  = Join-Path $paths.WimsServiced $idxName

        Write-Debug "Service-Index: mountDir='$mountDir' wimPath='$wimPath' srvPath='$srvPath'"

        try {
            if ($InstallType) {
                # Copy extracted to serviced, mount and service SSUs (may be empty) on install.wim
                Add-Packages -IdxName $idxName -MountDir $mountDir -Src $wimPath -Dest $srvPath -Mount $true -Unmount $false -PkgName "SSU" -Pkgs $ssuFiles

                # Mount embedded WinRE, service with SSUs, and dismount
                $winrePath      = Join-Path $mountDir $paths.WinreWimInWim
                $safeParent     = (Protect-Token (Split-Path $mountDir -Leaf))
                $safeLeaf       = (Protect-Token $paths.WinreWimInWim)
                $winrePathMount = Join-Path $paths.WimsMounts ("mount_{0}_{1}_{2}" -f $idxName, $safeParent, $safeLeaf)
                Add-Packages -IdxName $names.WinreWim -MountDir $winrePathMount -Src $winrePath -Dest "" -Mount $true -Unmount $true -PkgName "SSU" -Pkgs $ssuFiles

                # Service LCUs on install.wim and finally dismount
                Add-Packages -IdxName $IdxName -MountDir $mountDir -Src $srvPath -Dest "" -Mount $false -Unmount $true -PkgName "LCU" -Pkgs $lcuFiles
            } else {
                # Boot WIM: mount, service with SSUs (may be empty), and dismount
                Add-Packages -Idx $IdxName -MountDir $mountDir -Src $wimPath -Dest $srvPath -Mount $true -Unmount $true -PkgName "SSU" -Pkgs $ssuFiles
            }

            Write-JsonFile -Path $imgJson -Data @{ Index = $idx; Name = $imgName; ServicedDate = (Get-Date -Format s) }
            Write-Host "  $WimName index $idx serviced"
        }
        catch {
            Write-Host "  ERROR servicing $WimName index $($idx)"
            throw
        }
        finally {
            if (Test-Path $mountDir) {
                Write-Verbose "Service-Index: Cleaning mount directory $mountDir"
                Remove-Folder $mountDir
            }
        }
    }

    # -----------------------------------------------------------------------
    # Per-index servicing (using Service-Index)
    # -----------------------------------------------------------------------

    # Clean the scratch directory
    Remove-Folder $scratchDir
    Ensure-Folder $scratchDir

    Write-Host "Servicing install indices..."
    foreach ($img in $InstallIndices) {
        Service-Index -WimName $names.InstallWim -Img $img -InstallType $true
    }

    Write-Host "Servicing boot indices..."
    foreach ($img in $BootImages) {
        Service-Index -WimName $names.BootWim -Img $img -InstallType $false
    }

    # Clean the scratch directory
    Remove-Folder $scratchDir

    Write-Host "Service workflow complete"
}

function Invoke-FinalAssembly {
    [CmdletBinding()]
    param()

    if ($Clean) {
        Clean-Folder $paths.WimsFinal
        return
    }

    if ($DryRun) {
        Write-Host "[DryRun] Would assemble final $($names.InstallWim) -> $($paths.InstallWimInDest)"
        Write-Host "[DryRun] Would assemble final $($names.BootWim)    -> $($paths.BootWimInDest)"
        return
    }

    Write-Host  "Starting Final Assembly workflow..."

    # Sanity check our indices first
    Test-MissingIndices "Final"

    # Resolve anything that hasn't been
    $InstallIndices = Resolve-IndexSelection

    Write-Verbose "Invoke-FinalAssembly: WimsServiced='$($paths.WimsServiced)' WimsFinal='$($paths.WimsFinal)'"

    Ensure-Folder $paths.WimsServiced
    Ensure-Folder $paths.WimsFinal

    function Find-AnyOldServiced {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory=$true)]
            [System.Object[]]$Images,

            [Parameter(Mandatory=$true)]
            [string]$WimLabel
        )

        $finalJson = Join-Path $paths.WimsFinal ("{0}.json" -f $WimLabel)
        if (-not (Test-Path $finalJson)) {
            Write-Host "Final JSON for $WimLabel not found, rebuilding"
            return $true
        }

        try {
            $finalMeta = Read-JsonFile -Path $finalJson
            $finalDate = [datetime]::Parse($finalMeta.Date)
        }
        catch {
            Write-Host "Final JSON for $WimLabel invalid, rebuilding"
            return $true
        }

        $need = $false
        foreach ($img in $Images) {
            $idx = $img.Index
            $srv = Join-Path $paths.WimsServiced ("{0}_{1}" -f $idx, $WimLabel)
            if (-not (Test-Path $srv)) {
                throw "Missing serviced $WimLabel index $idx ($($img.Name)), run -Service again"
            }

            $srvJson = Join-Path $paths.WimsServiced ("{0}_{1}.json" -f $idx, $WimLabel)
            try {
                $meta = Read-JsonFile -Path $srvJson
                $sd = [datetime]::Parse($meta.ServicedDate)
                if ($sd -gt $finalDate) {
                    Write-Host "$WimLabel index $idx ($($img.Name)) is newer than final"
                    $need = $true
                }
            } catch {
                throw "$WimLabel index $idx ($($img.Name)) missing metadata or invalid, run -Service again"
            }
        }

        return $need
    }

    function Build-FinalImage {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory=$true)]
            [System.Object[]]$Images,

            [Parameter(Mandatory=$true)]
            [string]$WimLabel,

            [Parameter(Mandatory=$true)]
            [string]$DestPath
        )

        Remove-Folder $DestPath

        foreach ($img in $Images) {
            $idx = $img.Index

            $srvJson = Join-Path $paths.WimsServiced ("{0}_{1}.json" -f $idx, $WimLabel)
            $meta    = Read-JsonFile -Path $srvJson
            if (-not $meta -or -not $meta.ServicedDate) {
                Write-Host "$WimLabel index $idx ($($img.Name)) missing metadata, run -Service again"
                return
            }

            $src = Join-Path $paths.WimsServiced ("{0}_{1}" -f $idx, $WimLabel)
            if (-not (Test-Path $src)) {
                Write-Host "Missing serviced source for $WimLabel index $idx"
                return
            }

            $baseArgs = @(
                '/Export-Image',
                "`"/DestinationImageFile:$DestPath`"",
                '/SourceIndex:1',
                "`"/Compress:$DismCompression`"",
                '/CheckIntegrity'
            )

            Write-Host "  Adding $WimLabel index $idx from $(FolderRelName $src)"
            Run-Dism ($baseArgs + @("`"/SourceImageFile:$src`"")) -Indent 2
            if ($LASTEXITCODE -ne 0) {
                Write-Host "ERROR exporting $WimLabel index $idx"
                return
            }
        }

        $finalJson = Join-Path $paths.WimsFinal ("{0}.json" -f $WimLabel)
        Write-JsonFile -Path $finalJson -Data @{ Date = (Get-Date -Format s) }
    }

    if (Find-AnyOldServiced $InstallIndices $names.InstallWim) {
        Write-Host "Rebuilding final $($names.InstallWim)..."
        Build-FinalImage $InstallIndices $names.InstallWim $paths.InstallWimInFinal
    } else {
        Write-Host "Final $($names.InstallWim) is up-to-date"
    }

    if (Find-AnyOldServiced $BootImages $names.BootWim) {
        Write-Host "Rebuilding final $($names.BootWim)..."
        Build-FinalImage $BootImages $names.BootWim $paths.BootWimInFinal
    } else {
        Write-Host "Final $($names.BootWim) is up-to-date"
    }

    Write-Host "Final workflow complete"
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
        Write-Host "[DryRun] Would hardlink-copy $(FolderRelName $paths.SrcIsoContent) -> $(FolderRelName $paths.DestIsoContent)"
        Write-Host "[DryRun] Would copy $(FolderRelName $paths.InstallWimInFinal) -> $(FolderRelName $paths.InstallWimInDest)"
        Write-Host "[DryRun] Would copy $(FolderRelName $paths.BootWimInFinal)    -> $(FolderRelName $paths.BootWimInDest)"
        return
    }

    Write-Host "Starting PrepDestISO workflow..."
    Write-Verbose "Invoke-PrepDestISO: SrcIsoContent='$($paths.SrcIsoContent)' DestIsoContent='$($paths.DestIsoContent)'"

    $installWim          = $names.InstallWim
    $bootWim             = $names.bootWim
    $extractJson         = $paths.ExtractJson
    $finalInstallWimJson = Join-Path $paths.WimsFinal   ("{0}.json" -f $installWim)
    $finalBootWimJson    = Join-Path $paths.WimsFinal   ("{0}.json" -f $bootWim)
    $prepHardlinkJson    = Join-Path $paths.DestIsoRoot 'hardlink.json'
    $prepInstallWimJson  = Join-Path $paths.DestIsoRoot ("{0}.json" -f $installWim)
    $prepBootWimJson     = Join-Path $paths.DestIsoRoot ("{0}.json" -f $bootWim)

    function Fetch-Date {
        [CmdletBinding()]
        param(
            [string]$JsonPath, 
            [string]$Step = ""
        )
        try {
            $meta = Read-JsonFile -Path $JsonPath
            $date = [datetime]::Parse($meta.Date)
        }
        catch {
            if ([string]::IsNullOrEmpty($Path)) {
                # Any date is never than this
                return [datetime]::MinValue
            } else {
                throw "$(FolderRelName $JsonPath) not found or invalid, run -$(Step) first"
            }
          
        }
        return $date
    }

    # Check our prerequites
    # 1. Source folder for the ISO
    $extractDate = Fetch-Date $extractJson 'Extract'
    
    # 2. Missing files
    if (Report-Missing -Required @($finalInstallWimJson, $finalBootWimJson, $paths.InstallWimInFinal, $paths.BootWimInFinal)) {
        throw "See above for missing file details, run -Final first"
    }

    # 3. Finalized install.wim and boot.wim
    $finalInstallWimDate = Fetch-Date $finalInstallWimJson 'Final'
    $finalBootWimDate    = Fetch-Date $finalBootWimJson 'Final'

    # ---- Step A: Hardlink-copy SrcIsoContent -> DestIsoContent ----
    $prepHardlinkDate = Fetch-Date $prepHardlinkJson
    if ($prepHardlinkDate -le $extractDate) {
        Write-Host "Hardlink-copying $(FolderRelName $paths.SrcIsoContent) -> $(FolderRelName $paths.DestIsoContent) (excluding install/boot images)..."

        if (Test-Path $paths.DestIsoRoot) {
            Write-Host "Removing existing DestIsoRoot..."
            Remove-Folder $paths.DestIsoRoot
            Ensure-Folder $paths.DestIsoRoot
        }
        Ensure-Folder $paths.DestIsoContent

        # ------------------------------------------------------------
        # 1. Create all directories first (including empty ones)
        # ------------------------------------------------------------
        Write-Host "  Creating directories including empty ones..."
        $allDirs = @(Get-ChildItem -Path $paths.SrcIsoContent -Recurse -Directory -ErrorAction SilentlyContinue)
        foreach ($dir in $allDirs) {
            $rel  = $dir.FullName.Substring($paths.SrcIsoContent.TrimEnd('\').Length).TrimStart('\')
            $dest = Join-Path $paths.DestIsoContent $rel
            Ensure-Folder $dest
        }

        # ------------------------------------------------------------
        # 2. Hardlink all files (excluding boot/install images)
        # ------------------------------------------------------------
        $excludeNames = @($names.BootWim, $names.InstallWim, $names.InstallEsd)
        $allFiles = @(Get-ChildItem -Path $paths.SrcIsoContent -Recurse -File -ErrorAction SilentlyContinue)
        $total = $allFiles.Count
        $done = 0
        $lastPct = -1

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
                    Write-Warning "Hardlink failed for '$rel', copying: $_"
                    Copy-Item -Path $file.FullName -Destination $dest -Force
                }
            }
        }
        Write-JsonFile -Path $prepHardlinkJson -Data @{ Date = (Get-Date -Format s) }
        Write-Host "  Hardlink tree complete"
    } else {
        Write-Host "DestIsoContent hardlink-copy already current ($(FolderRelName $prepHardlinkJson): $prepHardlinkDate)"
    }

    # ---- Step B: Copy final install.wim ----
    $prepInstallWimDate = Fetch-Date $prepInstallWimJson
    if ($prepInstallWimDate -le $finalInstallWimDate) {
        # Stream-FileCopy will tell us what it is doing
        Stream-FileCopy $paths.InstallWimInFinal $paths.InstallWimInDest
        Write-JsonFile -Path $prepInstallWimJson -Data @{ Date = (Get-Date -Format s) }
    } else {
        Write-Host "$($names.InstallWim) already current ($(FolderRelName $prepInstallWimJson): $prepInstallWimDate)"
    }

    # ---- Step C: Copy final boot.wim ----
    $prepBootWimDate = Fetch-Date $prepBootWimJson
    if ($prepBootWimDate -le $finalBootWimDate) {
        # Stream-FileCopy will tell us what it is doing
        Stream-FileCopy $paths.BootWimInFinal $paths.BootWimInDest
        Write-JsonFile -Path $prepBootWimJson -Data @{ Date = (Get-Date -Format s) }
    } else {
        Write-Host "$($names.BootWim) already current ($(Resolve-IndexSelection $prepBootWimJson): $prepBootWimDate)"
    }

    Write-Host "PrepDestISO workflow complete"
}

# ==============================
# File work for the destination ISO
# ==============================
function Invoke-FilesWork {
    [CmdletBinding()]
    param()

    Write-Host "Starting Files workflow..."

    # Lists
    $destFolders = $(
        $paths.WinPEDriverRoot
        $paths.RegistryRoot
    )

    $requiredFiles = $(
        $names.ExportDriversCmd
        $names.InstallDriversCmd
        $names.ExportRegsCmd
        $names.ExportRegsPs1
        $names.InstallRegsCmd
        $names.UpdateNETCmd
        $names.UpdateNETPs1
        $names.WindowsInstallationCmd
        $names.PostSetupCmd
    )

    # Copy these KBs folders, if they exist
    $maybeKBs = $(
        $names.NET
        $names.MISC
    )

<#
    # Each entry is: TargetFile, List of (SearchPattern, Replacement) pairs to apply to the target file before copying to the destination
    # **** The SearchPattern needs to be a regex to match the line to replace with the Replacement ****
    $transforms = @(
        @($names.ExportDriversCmd, @(
            @('set "FLD=$WinPEDriver$"', $names.WinPEDriver)
        )),
        @($names.InstallDriversCmd, @(
            @('set "FLD=$WinPEDriver$"', $names.WinPEDriver)
        )),
        @($names.ExportRegsPs1, @(
            @('set "FLD=Registry"', $names.Registry)
        )),
        @($names.InstallRegsCmd, @(
            @("$RegistryRoot = 'Registry'", $names.Registry)
        ))
    )

    #
    # Transforms block
    #
    # These transforms rewrite destination paths for special cases
    #
    $transforms = @(
        @{ Match = $names.SetupConfigCleanIni;   To = $paths.SetupConfigIni }
        @{ Match = $names.SetupConfigUpgradeIni; To = $paths.SetupConfigIni }
    )
#>

    # Build jobs list
    $jobs = @()

    foreach ($f in $destFolders) {
        $jobs += [ordered]@{ Action = 'Folder'; From = $f; To = $null }
    }

    foreach ($f in $requiredFiles) {
        $jobs += [ordered]@{ Action = 'FileCopy'; From = (Join-Path $Folder $f); To = (Join-Path $paths.DestIsoContent $f) }
    }

    $addedEnsure = $false
    foreach ($kb in $maybeKBs) {
        $src = Join-Path $paths.KBsRoot $kb
        if (Test-Path $src) {
            if (-not $addedEnsure) {
                $addedEnsure = $true
                $jobs += [ordered]@{ Action = 'Folder'; From = $paths.KBsInDest; To = $null }
            }
            $dst = Join-Path $paths.KBsInDest $kb
            $jobs += [ordered]@{ Action = 'SubFolderCopy'; From = $src; To = $dst }
        }
    }

    # Debug dump
    if ($DebugPreference -eq 'Continue') {
        Write-Debug "Jobs list:"
        $w = ($jobs.Action | Measure-Object -Maximum -Property Length).Maximum
        foreach ($j in $jobs) {
            $a = $j.Action.PadLeft($w)
            Write-Debug ("  {0} {1} {2}" -f $a, $j.From, $j.To)
        }
    }

    # Missing-file check
    $missing = $jobs |
        Where-Object   { $_.Action -like '*FileCopy' } |
        ForEach-Object { $_.From }

    if (Report-Missing -Required $missing) {
        throw "Boot files are missing from $Folder"
    }

    # Execute each job
    foreach ($j in $jobs) {
        switch ($j.Action) {
            'Folder' {
                if     ($Clean)  { Clean-Folder $j.From }
                elseif ($DryRun) { Write-Host "[DryRun] Would create: $(FolderRelName $j.From)" }
                else             { Ensure-Folder $j.From }
            }
            'FileCopy' {
                if     ($Clean)  { Clean-File $j.To }
                elseif ($DryRun) { Write-Host "[DryRun] Would write : $(FolderRelName $j.To)" }
                else             { Stream-FileCopy -SourcePath $j.From -DestinationPath $j.To -NoComplete }
            }
            'SubFileCopy' {
                if     ($Clean)  { <# Parent folder already deleted #>}
                elseif ($DryRun) { Write-Host "[DryRun] Would write : $(FolderRelName $j.To)" }
                else             { Stream-FileCopy -SourcePath $j.From -DestinationPath $j.To -NoComplete }
            }
            'SubFolderCopy' {
                if     ($Clean)  { <# Parent folder already deleted #>}
                elseif ($DryRun) { Write-Host "[DryRun] Would copy  : $(FolderRelName $j.From) to $(FolderRelName $j.To)" }
                else             { if (Test-Path $j.To) { Remove-Item $j.To -Recurse -Force -ErrorAction SilentlyContinue }  # Clean out any old files
                                   Stream-FolderCopy -SourcePath $j.From -DestinationPath $j.To -NoComplete }
            }
        }
    }


    Write-Host "Files workflow complete"
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
        Write-Host "[DryRun]             from: $(FolderRelName $paths.DestIsoContent)"
        return
    }

    Write-Host "Starting CreateISO workflow..."
    Write-Verbose ("Invoke-CreateISOWork: DestIsoContent='{0}' DestISO='{1}'" -f $paths.DestIsoContent, $DestISO)

    if (-not $oscdimgExe) {
        Write-Warning "oscdimg.exe not found, install Windows ADK or specify -oscdimg"
        return
    }
    if (-not $DestISO) {
        Write-Warning "DestISO path is not set, use -DestISO or -ISO is provided"
        return
    }
    
    # Needed for the boot data
    $etfs = $paths.BIOSInDest
    $efis = $paths.UEFIInDest

    # Required boot files
    $required = @(
        $etfs
        $efis
        $paths.SetupExeInDest
        $paths.SourcesInDest
        $paths.BootWimInDest
        $paths.InstallWimInDest
    )
    if (Report-Missing -Required $required) {
        throw "Boot files are missing from the destination ISO content, run -Prep first to prepare the destination ISO"
    }

    if ($AutoGeneratedDestISO -or -not $DestISOVolumeLabel) {
        # Get the metadata from the boot.wim as it is always the smallest file to read
        $metaBoot               = Get-WimMetadata -WimPath $paths.BootWimInDest
        $buildServicePack       = "_{0}_{1}" -f $metaBoot.Build, $metaBoot.ServicePackBuild
        if ($AutoGeneratedDestISO) {
            $DestISO            = $DestISO -replace ([Regex]::Escape($names.DefaultDestISO) + '\.iso$'), ("$($buildServicePack).iso")
        }
        if (-not $DestISOVolumeLabel) {
            $DestISOVolumeLabel = "Win$($metaBoot.WinOS)_$($metaBoot.Version)_$($metaBoot.Arch)$($buildServicePack)"
        }
    }
    $bootdata = "2#p0,e,b$etfs#pEF,e,b$efis"

    # Build argument list for oscdimg
    $oscdimgArgs = @(
        "-o"                        # Optimize duplicate files
        "-u2"                       # UTF-8 filenames
        "-udfver102"                # UDF 1.02 for compatibility
        "-l`"$DestISOVolumeLabel`"" # Volume label
        "-bootdata:$bootdata"       # BIOS and UEFI boot entries
        $paths.DestIsoContent       # Source folder
        $DestISO                    # Output ISO
    )

    Write-Host "Building ISO: $DestISO"
    Write-Verbose ("& {0} {1}" -f $oscdimgExe, ($oscdimgArgs -join ' '))

    # Build process start info
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $oscdimgExe
    $psi.Arguments = ($oscdimgArgs -join " ")
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.CreateNoWindow = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $proc.Start() | Out-Null

    # Track last printed line to avoid duplicates
    $script:lastLine = ""

    # Track next 10 percent mark
    $script:nextMark = [int]10

    # Unified line handler
    function Handle-Line {
        param([string]$line)

        if (-not $line)                 { return }
        if ($line -eq $script:lastLine) { return }

        $script:lastLine = $line

        # Case 1: Mixed line with prefix + percent, just output the prefix
        if ($line -match "^(\S.*\S)\s+(\d+)% complete$") {
            Write-Host $Matches[1]
            return
        }

        # Case 2: Pure progress line
        if ($line -match "^\s*(\d+)% complete$") {
            $pct = [int]$Matches[1]
            if ($pct -lt $script:nextMark) { return }
            Write-Host ("Progress: {0}%" -f $pct)
            $script:nextMark += 10
            return
        }

        # Case 3: Normal output (optimization lines, copyright, etc.)
        Write-Host $line
    }

    # Real-time unified non-blocking read loop
    while (-not $proc.HasExited) {

        # stdout available?
        if ($proc.StandardOutput.Peek() -ne -1) {
            Handle-Line ($proc.StandardOutput.ReadLine())
        }

        # stderr available?
        if ($proc.StandardError.Peek() -ne -1) {
            Handle-Line ($proc.StandardError.ReadLine())
        }

        Start-Sleep -Milliseconds 20
    }

    # Flush remaining lines after exit
    while ($proc.StandardOutput.Peek() -ne -1) {
        Handle-Line ($proc.StandardOutput.ReadLine())
    }
    while ($proc.StandardError.Peek() -ne -1) {
        Handle-Line ($proc.StandardError.ReadLine())
    }

    $proc.WaitForExit()

    if ($proc.ExitCode -ne 0) {
        Write-Host ("ERROR: oscdimg failed to build ISO (exit {0})" -f (& $Hex $proc.ExitCode))
        return
    }

    Write-Host "Created ISO: $DestISO"
}

# Real work starts here
if ($Usage) {
    Show-Usage
    exit
}

# ==============================
# Resolve working folder
# ==============================
Write-Verbose "Working folder: $Folder"
$Folder = Resolve-FullPath $Folder
Write-Verbose "Resolved working folder: $Folder"

# ==============================
# Resolve DISM and oscdimg
# ==============================
Write-Verbose "Resolving tool paths..."
$dismExe    = Find-ADKTool -ToolName 'dism.exe'    -ADKSubfolder 'DISM'    -ExplicitPath $dism    -PreferADK:$UseADK -ForceSystem:$UseSystem
$oscdimgExe = Find-ADKTool -ToolName 'oscdimg.exe' -ADKSubfolder 'Oscdimg' -ExplicitPath $oscdimg -PreferADK:$UseADK -ForceSystem:$UseSystem

if (-not $dismExe) {
    Write-Host "ERROR: dism.exe is required but was not found"
    Write-Host "       Install the Windows ADK or use -dism to specify its path"
    exit 1
}
Write-Host (&$LeadIn "dism" "$dismExe")
if ($oscdimgExe) {
    Write-Host (&$LeadIn "oscdimg" "$oscdimgExe")
} else {
    Write-Host (&$LeadIn "oscdimg" "not found (ISO creation unavailable, -CreateISO will fail)")
}

# Core paths begin
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
$paths.SetupExeInDest        = Join-Path $paths.DestIsoContent $names.SetupExe
$paths.SourcesInDest         = Join-Path $paths.DestIsoContent $names.Sources
$paths.BootWimInDest         = Join-Path $paths.SourcesInDest $names.BootWim
$paths.InstallWimInDest      = Join-Path $paths.SourcesInDest $names.InstallWim
$paths.WinPEDriverRoot       = Join-Path $paths.DestIsoContent $names.WinPEDriver
$paths.RegistryRoot          = Join-Path $paths.DestIsoContent $names.Registry
$paths.WinreWimInWim         = Join-Path 'Windows\System32\Recovery' $names.WinreWim
$paths.KBsRoot               = Join-Path $Folder $names.KBs
foreach ($u in $kbDirs) {
    $paths["KBs$u"]          = Join-Path $paths.KBsRoot $names.$u
}
$paths.KBsInDest             = Join-Path $paths.DestIsoContent $names.KBs
foreach ($u in $kbDirs) {
    $paths["KBsInDest$u"]    = Join-Path $paths.KBsInDest $names.$u
}
$paths.WimsRoot              = Join-Path $Folder $names.Wims
foreach ($u in $wimDirs) {
    $paths["Wims$u"]         = Join-Path $paths.WimsRoot $names.$u
}
$paths.ExtractJson           = Join-Path $paths.SrcIsoRoot  $names.ExtractJson
$paths.MetadataJson          = Join-Path $paths.WimsIndices $names.MetadataJson
$paths.InstallWimInFinal     = Join-Path $paths.WimsFinal   $names.InstallWim
$paths.BootWimInFinal        = Join-Path $paths.WimsFinal   $names.BootWim
# Core paths end

# ==============================
# Resolve source ISO
# ==============================
if (-not $ISO) {
    Write-Verbose "No -ISO specified, searching for *.iso in: $Folder"
    $isoFiles = @(Get-ChildItem -Path $Folder -Filter '*.iso' -File -ErrorAction SilentlyContinue)
    if ($isoFiles.Count -eq 0) {
        $needsISO = $Extract -or $Export -or (-not ($KB -or $Service -or $Drivers -or $Reg -or $Files -or $Prep -or $CreateISO))
        if ($needsISO -and -not $Clean -and -not $DryRun) {
            Write-Error "No .iso file found in: $Folder`nPlace the Windows ISO there or use -ISO to specify its path"
            exit 1
        }
        Write-Verbose "No ISO found, continuing (ISO not required for selected operations)"
        $ISO = $null
    } elseif ($isoFiles.Count -gt 1) {
        Write-Error ("Multiple .iso files found in: $Folder`n  {0}`nUse -ISO to specify which one to use" -f ($isoFiles.FullName -join "`n  "))
        exit 1
    } else {
        $ISO = $isoFiles[0].FullName
        Write-Host (&$LeadIn "Auto-discovered ISO" "$ISO")
    }
}

if ($ISO -and (Test-Path $ISO)) {
    $ISO = Resolve-FullPath $ISO
    Write-Verbose (&$LeadIn "Resolved ISO path" "$ISO")
}

# ==============================
# Resolve destination ISO
# ==============================
if (-not $DestISO -and $ISO) {
    $DestISO = $ISO -replace '\.iso$', ($names.DefaultDestISO + '.iso')
    $AutoGeneratedDestISO = $true
    Write-Host (&$LeadIn "Auto-derived DestISO" "$DestISO")
}
$DestISO = Resolve-FullPath $DestISO

# ==============================
# Read ISO / WIM metadata for WinOS / Version / Arch and index list
# Priority: 1) MetadataJson  2) SrcIsoContent WIMs  3) Mount ISO
# ==============================
$InstallImages   = @()
$BootImages      = @()
$IsoMetaResolved = $false
$MetaSrc         = 'To be determined at export time'

    # Helper: apply WimMetadata result to the script-scope variables
    function Apply-WimMetadata {
        param([object]$Meta)
        if (-not $script:WinOS)   { $script:WinOS   = $Meta.WinOS }
        if (-not $script:Version) { $script:Version = $Meta.Version }
        if (-not $script:Arch)    { $script:Arch    = $Meta.Arch }
    }
    function Try-WimMetadata {
        param([string]$InstallWimPath, [string]$InstallEsdPath, [string]$BootWimPath, [string]$From)
        $installWim = if     (Test-Path $InstallWimPath) { $InstallWimPath }
                      elseif (Test-Path $InstallEsdPath) { $InstallEsdPath }
                      else   { $null }
        $bootWim    = if     (Test-Path $BootWimPath)    { $BootWimPath }
                      else   { $null }
        if ($installWim -and $bootWim) {
            Write-Verbose "Reading metadata from $($From): $InstallWimPath and $BootWimPath"
            try {
                $metaInstall = Get-WimMetadata -WimPath $InstallWimPath
                $metaBoot    = Get-WimMetadata -WimPath $BootWimPath
                if ($metaInstall.InstallImages.Count -gt 0) {
                    $script:InstallImages = @($metaInstall.InstallImages | ForEach-Object { [PSCustomObject]@{ Index = [int]$_.Index; Name = $_.Name } })
                    Apply-WimMetadata $metaInstall
                }
                if ($metaBoot.BootImages.Count -gt 0) {
                    $script:BootImages = @($metaBoot.BootImages | ForEach-Object { [PSCustomObject]@{ Index = [int]$_.Index; Name = $_.Name } })
                }
            } catch {
                Write-Warning "Failed to read WIM metadata from '$From': $_"
            }
            if (($script:InstallImages.Count -gt 0) -and ($script:BootImages.Count -gt 0)) { $script:IsoMetaResolved = $true; $script:MetaSrc = $From }
        }
    }

# 1) Prefer cached MetadataJson if it matches the current ISO
$metadataJson = $paths.MetadataJson
$wimMeta      = Read-JsonFile -Path $metadataJson
if ($wimMeta -and (-not $ISO -or $wimMeta.ISOPath -eq $ISO)) {
    Write-Verbose "Loading metadata from $metadataJson"
    $InstallImages = @($wimMeta.InstallImages | ForEach-Object { [PSCustomObject]@{ Index = [int]$_.Index; Name = $_.Name } })
    $BootImages    = @($wimMeta.BootImages    | ForEach-Object { [PSCustomObject]@{ Index = [int]$_.Index; Name = $_.Name } })
    Apply-WimMetadata $wimMeta
    if (($InstallImages.Count -gt 0) -and ($BootImages.Count -gt 0)) { $isoMetaResolved = $true; $MetaSrc = $names.MetadataJson }
}

if (-not $DryRun -and -not $Clean) {
    # 2) Fall back to SrcIsoContent on disk
    if (-not $isoMetaResolved -and (Test-Path $paths.SourcesInSrc)) {
        Try-WimMetadata -InstallWimPath $paths.InstallWimInSrc `
                        -InstallEsdPath $paths.InstallEsdInSrc `
                        -BootWimPath    $paths.BootWimInSrc `
                        -From           "SrcIsoContent"
    }

    # 3) Mount the ISO briefly if still needed
    if (-not $isoMetaResolved -and $ISO -and (Test-Path $ISO)) {
        Write-Verbose "Mounting ISO for metadata: $ISO"
        $metaDiskImg = Mount-DiskImage -ImagePath $ISO -PassThru -ErrorAction SilentlyContinue
        try {
            # Wait for the volume to actually appear
            $vol = $null
            for ($r = 0; $r -lt 5 -and $null -eq $vol; $r++) {
                $vol = $metaDiskImg | Get-Volume -ErrorAction SilentlyContinue
                if ($null -eq $vol) { Start-Sleep -Seconds 1 }
            }
            if ($null -eq $vol) { throw "Timeout waiting for ISO volume to initialize" }

            $metaDrive = $vol.DriveLetter + ':\'
            Try-WimMetadata -InstallWimPath "$metaDrive$($paths.InstallWimInIso)" `
                            -InstallEsdPath "$metaDrive$($paths.InstallEsdInIso)" `
                            -BootWimPath    "$metaDrive$($paths.BootWimInIso)" `
                            -From           "Mounted ISO"
        } finally {
            if ($metaDiskImg -and $metaDiskImg.Attached) {
                Write-Verbose "Dismounting ISO..."
                Dismount-DiskImage -ImagePath $ISO -ErrorAction SilentlyContinue | Out-Null
            }
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
    if ($InstallImages.Count -eq 0) {
        Write-Error "Cannot show indices: no metadata available, run -Extract first, or provide -ISO"
        exit 1
    }
    Write-Host "`nAvailable images in $($names.InstallWim) [source: $MetaSrc]:`n"
    Write-Host ("{0,6}  {1}" -f 'Index', 'Name')
    Write-Host ("{0,6}  {1}" -f '------', '----')
    foreach ($img in $InstallImages) { Write-Host ("{0,6}  {1}" -f $img.Index, $img.Name) }
    Write-Host ""
    exit 0
}

# ==============================
# Resolve index selection
# ==============================
$InstallIndices = @()
if ($InstallImages.Count -gt 0) {
    $InstallIndices = Resolve-IndexSelection
} else {
    Write-Verbose "Image list unavailable yet, index selection deferred until -Export"
}

# ==============================
# Determine work modes
# ==============================
$workSwitches = @()
if ($Extract)   { $workSwitches += 'Extract' }
if ($Export)    { $workSwitches += 'Export' }
if ($KB)        { $workSwitches += 'KB' }
if ($Service)   { $workSwitches += 'Service' }
if ($Final)     { $workSwitches += 'Final' }
if ($Prep)      { $workSwitches += 'Prep' }
if ($Files)     { $workSwitches += 'Files' }
if ($Drivers)   { $workSwitches += 'Drivers' }
if ($Reg)       { $workSwitches += 'Reg' }
if ($CreateISO) { $workSwitches += 'CreateISO' }

if ($All -or $Most -or (-not $workSwitches)) {
    $Extract = $true
    $Export  = $true
    $KB      = $true
    $Service = $true
    $Final   = $true
    $Prep    = $true
    $Files   = $true
    $Drivers = $true
    $Reg     = $true
    if ($Most) {
        $workSwitches = @('Most')
    } else {
        $CreateISO = $true
        $All       = $true
        $workSwitches = @('All')
    }
}

Write-Host (&$LeadIn "Target profile" "Windows $WinOS $Version $Arch")
Write-Host (&$LeadIn "Working folder" "$Folder")
Write-Host (&$LeadIn "ISO" "$(if ($ISO) { $ISO } else { '(none)' })")
Write-Host (&$LeadIn "DestISO" "$(if ($DestISO) { $DestISO } else { '(none)' })")
Write-Host (&$LeadIn "Selected indices" "$(if ($InstallIndices.Count -gt 0) { $InstallIndices.Index -join ', ' } else { 'all' }) [source: $MetaSrc]")
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
                Write-Host "ERROR: $HtmlAgilityPackDll not found inside the downloaded NuGet package, KB downloads will not work"
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
        # file lock on the DLL. This allows -Clean to delete it
        # even in the same PowerShell session
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
    if ($Final)     { Invoke-FinalAssembly }
    if ($Prep)      { Invoke-PrepDestISO }
    if ($Files)     { Invoke-FilesWork }
    if ($CreateISO) { Invoke-CreateISOWork }

    Write-Host "Completed"
} catch {
    Write-Host ""
    Write-Host "ERROR: $_"
    Write-Host "       Run the script again once the issue is resolved, completed steps will be skipped"
    exit 1
} finally {
    # -----------------------------------------------------------------------
    # Cleanup: release any resources that may have been left open if the
    # script was interrupted (Ctrl+C, early fatal error, etc.)
    # -----------------------------------------------------------------------

    # 1. Dismount the source ISO if it is still attached as a virtual drive
    if ($ISO -and (Test-Path $ISO)) {
        try {
            $img = Get-DiskImage -ImagePath $ISO -ErrorAction SilentlyContinue
            if ($img -and $img.Attached) {
                Write-Host "Cleanup: dismounting ISO..."
                Dismount-DiskImage -ImagePath $ISO -ErrorAction SilentlyContinue | Out-Null
            }
        } catch {}
    }

    # 2. Discard any DISM-mounted WIM images that the servicing loop left open
    #    Each active mount shows up as a non-empty subdirectory under WimsMounts
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
