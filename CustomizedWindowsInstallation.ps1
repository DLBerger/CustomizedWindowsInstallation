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
    - Install Drivers.cmd
    - SetupComplete.cmd
    - SetupConfig-Clean.ini
    - SetupConfig-Upgrade.ini
- Dry-run mode (no changes made)
- Refresh mode (re-download updates even if present)
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

.PARAMETER All
Enables all work modes (KB, Drivers, Reg).

.PARAMETER KB
Download OS and .NET updates.

.PARAMETER Drivers
Export drivers into $WinpeDriver$.

.PARAMETER Reg
Export registry keys.

.PARAMETER Clean
Remove generated content instead of creating it.

.PARAMETER DryRun
Show actions without performing them.

.PARAMETER Refresh
Re-download updates even if files already exist.

.PARAMETER Help
Displays help and exits.

.NOTES
- Fully compatible with Windows PowerShell 5.x.
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

    [switch]$All,
    [switch]$KB,
    [switch]$Drivers,
    [switch]$Reg,
    [switch]$Clean,

    [switch]$DryRun,
    [switch]$Refresh,

    [switch]$Help
)

# ==============================
# Apply defaults AFTER param()
# ==============================

if ($Help) {
    Get-Help -Full $PSCommandPath
    exit
}

# If -Debug was passed, force debug output to auto-continue
if ($PSBoundParameters.ContainsKey('Debug')) {
    $DebugPreference = 'Continue'
    Write-Debug "Debug mode enabled: DebugPreference set to 'Continue'"
}

if (-not $Folder) { $Folder = (Get-Location).ProviderPath }
if (-not $WinOS) { $WinOS = '11' }
if (-not $Arch)  { $Arch  = 'x64' }

if (-not $Version) {
    if ($WinOS -eq '10') { $Version = '22H2' }
    else                 { $Version = '25H2' }
}

# --- HtmlAgilityPack bootstrap (PS 5.x SAFE) ---------------------------------
# --- HtmlAgilityPack bootstrap (local temp folder + cleanup) -----------------
$hapDll = Join-Path $PSScriptRoot "HtmlAgilityPack.dll"

if (-not (Test-Path $hapDll)) {
    Write-Host "HtmlAgilityPack.dll not found; downloading..." -ForegroundColor Cyan

    $nugetUrl   = "https://www.nuget.org/api/v2/package/HtmlAgilityPack"
    $tmpNupkg   = Join-Path $PSScriptRoot "HtmlAgilityPack.nupkg"
    $extractDir = Join-Path $PSScriptRoot "HtmlAgilityPack_Extract"

    # Clean old extraction folder if it exists
    if (Test-Path $extractDir) {
        Remove-Item $extractDir -Recurse -Force
    }

    # --- Download using .NET WebClient (PS 5.x safe) ---
    Write-Verbose "[HAP] Downloading via WebClient..."
    $wc = New-Object System.Net.WebClient
    $wc.DownloadFile($nugetUrl, $tmpNupkg)

    # --- Extract using .NET ZipFile (PS 5.x safe) ---
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($tmpNupkg, $extractDir)

    # Prefer netstandard2.0, fallback to net45
    $candidatePaths = @(
        (Join-Path $extractDir "lib\netstandard2.0\HtmlAgilityPack.dll"),
        (Join-Path $extractDir "lib\net45\HtmlAgilityPack.dll")
    )

    $sourceDll = $candidatePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $sourceDll) {
        throw "HtmlAgilityPack.dll not found inside NuGet package."
    }

    Copy-Item -Path $sourceDll -Destination $hapDll -Force
    Write-Verbose "[HAP] HtmlAgilityPack.dll copied to: $hapDll"

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
    Write-Verbose "[HAP] Loading HtmlAgilityPack from: $hapDll"
    Add-Type -Path $hapDll
    Write-Debug "[HAP] HtmlAgilityPack successfully loaded."
}
# ---------------------------------------------------------------------------

# ==============================
# Helper: Write-Log
# ==============================
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "[$ts] [$Level] $Message"
}

# ==============================
# Resolve working folder
# ==============================
$Folder = (Resolve-Path -LiteralPath $Folder).ProviderPath
Write-Verbose "Resolved working folder: $Folder"

# ==============================
# Determine work modes
# ==============================
$workSwitches = @()
if ($All)     { $workSwitches += 'All' }
if ($KB)      { $workSwitches += 'KB' }
if ($Drivers) { $workSwitches += 'Drivers' }
if ($Reg)     { $workSwitches += 'Reg' }

if (-not $workSwitches) {
    $All = $true
    $KB = $true
    $Drivers = $true
    $Reg = $true
    Write-Verbose "Defaulting to -All"
}

Write-Log "Target profile: Windows $WinOS $Version $Arch"
Write-Log "Root folder   : $Folder"
Write-Log "Mode          : $($workSwitches -join ', ')"
if ($Clean)  { Write-Log "Clean mode    : Enabled" "WARN" }
if ($DryRun) { Write-Log "Dry-run mode  : Enabled" "WARN" }
if ($Refresh){ Write-Log "Refresh mode  : Enabled" }

# ==============================
# Core paths
# ==============================
$paths = [ordered]@{
    UpdatesRoot     = Join-Path $Folder 'Updates'
    UpdatesOSCU     = Join-Path $Folder 'Updates\OSCU'
    UpdatesNET      = Join-Path $Folder 'Updates\NET'
    WinpeDriverRoot = Join-Path $Folder '$WinpeDriver$'
    RegistryRoot    = Join-Path $Folder 'Registry'
    ScriptsRoot     = Join-Path $Folder 'Scripts'
}

if (-not $Clean -and -not $DryRun) {
    foreach ($p in $paths.Values) {
        if (-not (Test-Path $p)) {
            Write-Verbose "Creating folder: $p"
            New-Item -ItemType Directory -Path $p -Force | Out-Null
        }
    }
}

# =========================
# HTML-based Update Catalog search
# =========================

function Invoke-CatalogRequest {
    param(
        [Parameter(Mandatory)]
        [string]$Uri
    )

    Write-Verbose "[CatalogRequest] GET $Uri"

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

        Write-Debug "[CatalogRequest] RawContent length = $($Response.RawContent.Length)"

        $HtmlDoc = [HtmlAgilityPack.HtmlDocument]::new()
        $HtmlDoc.LoadHtml($Response.RawContent.ToString())

        return $HtmlDoc
    }
    catch {
        Write-Warning "[CatalogRequest] Failed: $_"
        return $null
    }
    finally {
        [Net.ServicePointManager]::SecurityProtocol = $oldProtocol
    }
}
function Parse-CatalogSearchResults {
    param(
        [Parameter(Mandatory)]
        [HtmlAgilityPack.HtmlDocument]$Html
    )

    Write-Verbose "[CatalogParse] Extracting update IDs from HTML"

    $ids = @()

    # Look for goToDetails('GUID')
    $pattern = "goToDetails\('([0-9a-fA-F-]{36})'\)"

    $matches = [regex]::Matches($Html.DocumentNode.InnerHtml, $pattern)

    foreach ($m in $matches) {
        $id = $m.Groups[1].Value
        Write-Debug "[CatalogParse] Found update ID: $id"
        $ids += $id
    }

    Write-Verbose "[CatalogParse] Total IDs extracted: $($ids.Count)"
    return $ids
}
function Search-UpdateCatalogHtml {
    param(
        [Parameter(Mandatory)]
        [string]$Query
    )

    Write-Verbose "[CatalogSearch] Searching Update Catalog (HTML mode)"
    Write-Verbose "[CatalogSearch] Query: $Query"

    $Encoded = [uri]::EscapeDataString($Query)
    $Uri = "https://www.catalog.update.microsoft.com/Search.aspx?q=$Encoded"

    Write-Debug "[CatalogSearch] Encoded URI: $Uri"

    $Html = Invoke-CatalogRequest -Uri $Uri
    if (-not $Html) {
        Write-Warning "[CatalogSearch] No HTML returned"
        return @()
    }

    return Parse-CatalogSearchResults -Html $Html
}

# ==============================
# Download helper
# ==============================
function Download-MUFile {
    param(
        [string]$Url,
        [string]$DestinationFolder,
        [switch]$Force
    )

    $fileName = Split-Path -Path $Url -Leaf
    $destPath = Join-Path $DestinationFolder $fileName

    if (-not $Force -and (Test-Path $destPath)) {
        Write-Verbose "Already exists: $fileName"
        return $destPath
    }

    if ($DryRun) {
        Write-Log "[DryRun] Would download: $fileName"
        return $destPath
    }

    Write-Log "Downloading: $fileName"
    Invoke-WebRequest -Uri $Url -OutFile $destPath
    return $destPath
}

# ==============================
# KB Work
# ==============================
function Invoke-KBWork {
    Write-Verbose "=== KB Work ==="

    # Build queries
    $osQuery  = "Cumulative Updates for Windows 11 Version 25H2 for x64-based Systems"
    $netQuery = ".NET Framework for Windows 11 Version 25H2 x64"

    Write-Verbose "[KB] OS Query:  $osQuery"
    Write-Verbose "[KB] .NET Query: $netQuery"

    # Search OS updates
    Write-Verbose "[KB] Searching OS updates..."
    $osGuids = Search-UpdateCatalogHtml -Query $osQuery

    # Search .NET updates
    Write-Verbose "[KB] Searching .NET updates..."
    $netGuids = Search-UpdateCatalogHtml -Query $netQuery

    # Combine
    $allGuids = $osGuids + $netGuids

    Write-Verbose "[KB] Total updates found: $($allGuids.Count)"
    Write-Debug   "[KB] GUID list:`n$($allGuids -join "`n")"

    Write-Verbose "[KB] KB work complete."
    return $allGuids
}

# ==============================
# Driver export
# ==============================
function Invoke-DriverWork {
    param([string]$WinpeDriverRoot, [switch]$Clean)

    if ($Clean) {
        if ($DryRun) {
            Write-Log "[DryRun] Would remove: $WinpeDriverRoot"
        } else {
            if (Test-Path $WinpeDriverRoot) { Remove-Item $WinpeDriverRoot -Recurse -Force }
        }
        return
    }

    if (-not $DryRun -and -not (Test-Path $WinpeDriverRoot)) {
        New-Item -ItemType Directory -Path $WinpeDriverRoot -Force | Out-Null
    }

    if ($DryRun) {
        Write-Log "[DryRun] Would run DISM export-driver"
        return
    }

    Write-Log "Exporting drivers..."
    $args = "/online /export-driver /destination:`"$WinpeDriverRoot`""
    $p = Start-Process -FilePath dism.exe -ArgumentList $args -NoNewWindow -PassThru -Wait -RedirectStandardOutput "$WinpeDriverRoot\dism.log"
    if ($p.ExitCode -ne 0) { throw "DISM failed." }
}

# ==============================
# Registry export
# ==============================
function Invoke-RegWork {
    param([string]$RegistryRoot, [switch]$Clean)

    if ($Clean) {
        if ($DryRun) {
            Write-Log "[DryRun] Would remove: $RegistryRoot"
        } else {
            if (Test-Path $RegistryRoot) { Remove-Item $RegistryRoot -Recurse -Force }
        }
        return
    }

    if (-not $DryRun -and -not (Test-Path $RegistryRoot)) {
        New-Item -ItemType Directory -Path $RegistryRoot -Force | Out-Null
    }

    $keys = @(
        'HKLM\SOFTWARE\MyCompany',
        'HKCU\Software\MyCompany'
    )

    foreach ($key in $keys) {
        $safe = ($key -replace '[\\/:*?"<>|]', '_') + '.reg'
        $dest = Join-Path $RegistryRoot $safe

        if ($DryRun) {
            Write-Log "[DryRun] Would export: $key -> $dest"
        } else {
            reg.exe export "$key" "$dest" /y | Out-Null
        }
    }
}

# ==============================
# Install Drivers.cmd
# ==============================
function Write-InstallDriversScript {
    param([string]$RootFolder)

    $path = Join-Path $RootFolder 'Install Drivers.cmd'
    $content = @'
@echo off
setlocal enabledelayedexpansion

set LOG=%SystemRoot%\Temp\InstallDrivers.log
echo [%DATE% %TIME%] Starting driver installation... > "%LOG%"

set DRIVERROOT=%~dp0$WinpeDriver$
if not exist "%DRIVERROOT%" (
    echo $WinpeDriver$ folder not found at "%DRIVERROOT%". >> "%LOG%"
    exit /b 0
)

pnputil /add-driver "%DRIVERROOT%\*.inf" /subdirs /install >> "%LOG%" 2>&1

exit /b %ERRORLEVEL%
'@

    if ($DryRun) {
        Write-Log "[DryRun] Would write: $path"
    } else {
        Set-Content -LiteralPath $path -Value $content -Encoding ASCII
    }
}

# ==============================
# SetupComplete.cmd
# ==============================
function Write-SetupCompleteScript {
    param([string]$ScriptsRoot)

    $path = Join-Path $ScriptsRoot 'SetupComplete.cmd'
    $content = @'
@echo off
setlocal enabledelayedexpansion

set LOG=%SystemRoot%\Setup\Scripts\SetupComplete.log
echo [%DATE% %TIME%] SetupComplete starting... > "%LOG%"

set BASE=%~dp0

:: Apply .NET updates
for %%F in ("%BASE%..\..\..\Updates\NET\*.msu") do (
    wusa.exe "%%F" /quiet /norestart >> "%LOG%" 2>&1
)

:: Apply OS updates
for %%F in ("%BASE%..\..\..\Updates\OSCU\*.msu") do (
    wusa.exe "%%F" /quiet /norestart >> "%LOG%" 2>&1
)

:: Import registry
for %%F in ("%BASE%..\..\..\Registry\*.reg") do (
    reg.exe import "%%F" >> "%LOG%" 2>&1
)

:: Install drivers
if exist "%SystemDrive%\Install Drivers.cmd" (
    call "%SystemDrive%\Install Drivers.cmd" >> "%LOG%" 2>&1
)

echo [%DATE% %TIME%] SetupComplete finished. >> "%LOG%"
exit /b 0
'@

    if ($DryRun) {
        Write-Log "[DryRun] Would write: $path"
    } else {
        if (-not (Test-Path $ScriptsRoot)) {
            New-Item -ItemType Directory -Path $ScriptsRoot -Force | Out-Null
        }
        Set-Content -LiteralPath $path -Value $content -Encoding ASCII
    }
}

# ==============================
# SetupConfig files
# ==============================
function Write-SetupConfigFiles {
    param([string]$RootFolder)

    $cleanPath   = Join-Path $RootFolder 'SetupConfig-Clean.ini'
    $upgradePath = Join-Path $RootFolder 'SetupConfig-Upgrade.ini'

    $clean = @'
[SetupConfig]
Auto=Clean
DynamicUpdate=Enable
Telemetry=Disable
'@

    $upgrade = @'
[SetupConfig]
Auto=Upgrade
DynamicUpdate=Enable
Telemetry=Disable
'@

    if ($DryRun) {
        Write-Log "[DryRun] Would write: $cleanPath"
        Write-Log "[DryRun] Would write: $upgradePath"
    } else {
        Set-Content -LiteralPath $cleanPath   -Value $clean   -Encoding ASCII
        Set-Content -LiteralPath $upgradePath -Value $upgrade -Encoding ASCII
    }
}

# ==============================
# Main orchestration
# ==============================
if ($KB) {
    Invoke-KBWork -WinOS $WinOS -Version $Version -Arch $Arch `
        -UpdatesOSCU $paths.UpdatesOSCU -UpdatesNET $paths.UpdatesNET `
        -Clean:$Clean -Refresh:$Refresh
}

if ($Drivers) {
    Invoke-DriverWork -WinpeDriverRoot $paths.WinpeDriverRoot -Clean:$Clean
}

if ($Reg) {
    Invoke-RegWork -RegistryRoot $paths.RegistryRoot -Clean:$Clean
}

if (-not ($Clean -and -not ($KB -or $Drivers -or $Reg))) {
    Write-InstallDriversScript -RootFolder $Folder
    Write-SetupCompleteScript -ScriptsRoot $paths.ScriptsRoot
    Write-SetupConfigFiles -RootFolder $Folder
}

Write-Log "Completed."
Write-Verbose "Script execution finished."
