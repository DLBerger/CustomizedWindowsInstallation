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

.PARAMETER KB
Download OS and .NET updates.

.PARAMETER Drivers
Export drivers into $WinpeDriver$.

.PARAMETER Reg
Export registry keys.

.PARAMETER Files
Generate Install Drivers.cmd, SetupComplete.cmd, and SetupConfig-*.ini files.

.PARAMETER All
Shorthand for -KB -Drivers -Reg -Files and the default if no specific switch is provided.

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

    [switch]$KB,
    [switch]$Drivers,
    [switch]$Reg,
    [switch]$Files,
    [switch]$All,

    [switch]$Clean,

    [switch]$DryRun,

    [switch]$Help
)

# git hash
$GitHash = "c98df56"

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
function Parse-CatalogSearchResults {
    param(
        [Parameter(Mandatory)]
        [HtmlAgilityPack.HtmlDocument]$Html
    )

    Write-Verbose "Extracting update IDs from HTML"
#   Write-Debug "HTML: $($Html.DocumentNode.InnerHtml)"

    $ids = @()

    # Look for goToDetails('GUID')
    $pattern = 'goToDetails\("([0-9A-Fa-f\-]{36})"\)'
    $matches = [regex]::Matches($Html.DocumentNode.InnerHtml, $pattern)
    foreach ($m in $matches) {
        $id = $m.Groups[1].Value
        Write-Debug "Found update ID: $id"
        $ids += $id
    }

    Write-Verbose "Total IDs extracted: $($ids.Count)"
    return $ids
}
function Search-UpdateCatalogHtml {
    param(
        [Parameter(Mandatory)]
        [string]$Query
    )

    Write-Verbose "Searching Update Catalog (HTML mode)"
    Write-Verbose "Query: $Query"

    $Encoded = [uri]::EscapeDataString($Query)
    $Uri = "https://www.catalog.update.microsoft.com/Search.aspx?q=$Encoded"

    Write-Debug "Encoded URI: $Uri"

    $Html = Invoke-CatalogRequest -Uri $Uri
    if (-not $Html) {
        Write-Warning "No HTML returned"
        return @()
    }

    return Parse-CatalogSearchResults -Html $Html
}

function Get-UpdateLinks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Guid
    )

    Write-Verbose ("GUID: {0}" -f $Guid)
    Write-Verbose ("Requesting DownloadDialog.aspx via POST")

    # Build POST body
    $postObject = @{
        size        = 0
        UpdateID    = $Guid
        UpdateIDInfo= $Guid
    } | ConvertTo-Json -Compress

    $body = @{
        UpdateIDs = "[$postObject]"
    }

    $params = @{
        Uri         = "https://www.catalog.update.microsoft.com/DownloadDialog.aspx"
        Method      = 'POST'
        Body        = $body
        ContentType = "application/x-www-form-urlencoded"
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
    Write-Debug   "Raw content (first 400 chars):`n$($content.Substring(0, [Math]::Min(400, $content.Length)))"

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

    $links = foreach ($m in $matches) {
        $downloadInfoIndex = [int]$m.Groups[1].Value
        $fileIndex         = [int]$m.Groups[2].Value
        $url               = $m.Groups[3].Value

        # Ignore garbage like "h" or empty strings
        if ([string]::IsNullOrWhiteSpace($url) -or
            $url.Length -lt 10 -or
            -not ($url -like "http*")) {

            Write-Warning ("Ignoring malformed URL: {0}" -f $url)
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
        Write-Warning ("Failed to load manifest from {0}: {1}" -f $manifestPath, $($_.Exception.Message))
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
        [pscustomobject] $Update,

        [Parameter(Mandatory = $true)]
        [string] $TargetFolder
    )

    Ensure-Folder -Path $TargetFolder

    Write-Host ("Preparing downloads for update {0}: {1}" -f $Update.Guid, $Update.Title)

    $results = @()

    # No URLs → nothing to do
    if (-not $Update.DownloadUrls -or $Update.DownloadUrls.Count -eq 0) {
        Write-Host ("No download URLs for update {0}" -f $Update.Guid)
        return @()
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

            $results += [pscustomobject]@{
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

        $results += [pscustomobject]@{
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
        [string] $Guid
    )

    Write-Host ("Processing update #{0}: {1}" -f $Count, $Guid)

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
        Write-Warning ("Failed to fetch details page for {0}: {1}" -f $Guid, $_.Exception.Message)
        return $null
    }

    $detailsDoc = New-Object HtmlAgilityPack.HtmlDocument
    $detailsDoc.LoadHtml($detailsResponse.Content)

    # Title
    $titleNode = $detailsDoc.DocumentNode.SelectSingleNode("//span[@id='ScopedViewHandler_titleText']")
    $title = if ($titleNode) { $titleNode.InnerText.Trim() } else { "" }

    # KB
    $kbMatch = [regex]::Match($title, "KB\d+")
    $kb = if ($kbMatch.Success) { $kbMatch.Value } else { "" }

    # SupersededBy
    $supersededBy = @()
    $supNodes = $detailsDoc.DocumentNode.SelectNodes("//div[@id='supersededbyInfo']//a")
    if ($supNodes) {
        foreach ($n in $supNodes) {
            $supersededBy += $n.InnerText.Trim()
        }
    }

    # ------------------------------------------------------------
    # 2. DOWNLOAD LINKS (via Get-UpdateLinks)
    # ------------------------------------------------------------

    Write-Host "Finding download links for $title"

    $links = Get-UpdateLinks -Guid $Guid
    $downloadUrls = @()
    if ($links) {
        $downloadUrls = $links.URL | Select-Object -Unique
    }

    Write-Host ("Found {0} file(s) for this update" -f $downloadUrls.Count)

    Write-Verbose ("Title: {0}" -f $title)
    Write-Verbose ("KB: {0}" -f $kb)
    if ($supersededBy.Count -gt 0) {
        Write-Verbose ("SupersededBy: {0}" -f ($supersededBy -join ', '))
    }
    Write-Verbose ("URLs: {0}" -f $downloadUrls.Count)

    return [pscustomobject]@{
        Guid           = $Guid
        Title          = $title
        KB             = $kb
        SupersededBy   = $supersededBy
        DownloadUrls   = $downloadUrls
    }
}

function Build-ManifestEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject] $Details,

        [Parameter(Mandatory = $true)]
        [pscustomobject] $DownloadInfo
    )

    [pscustomobject]@{
        Guid           = $Details.Guid
        Title          = $Details.Title
        DownloadUrl    = $DownloadInfo.Url
        FileName       = $DownloadInfo.FileName
        SupersededBy   = $Details.SupersededBy
        Timestamp      = (Get-Date).ToString("s")
    }
}

function Invoke-KBWork {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $WinOS,

        [Parameter(Mandatory = $true)]
        [string] $Version,

        [Parameter(Mandatory = $true)]
        [string] $Arch,

        [Parameter(Mandatory = $true)]
        [string] $UpdatesOSCU,

        [Parameter(Mandatory = $true)]
        [string] $UpdatesNET,

        [Parameter()]
        [switch] $Clean
    )

    Write-Log "Starting KB update workflow..."

    # Derive SSU folder from OSCU parent
    $rootFolder = Split-Path $UpdatesOSCU -Parent
    if (-not $rootFolder) { $rootFolder = $UpdatesOSCU }
    $UpdatesSSU = Join-Path $rootFolder 'UpdatesSSU'

    # Ensure folders
    foreach ($folder in @($UpdatesOSCU, $UpdatesNET, $UpdatesSSU)) {
        Ensure-Folder -Path $folder
    }

    # Clean mode
    if ($Clean) {
        if ($DryRun) {
            Write-Log "[DryRun] Would clean update folders: $UpdatesOSCU, $UpdatesNET, $UpdatesSSU"
        } else {
            Write-log "Cleaning update folders"
            foreach ($folder in @($UpdatesOSCU, $UpdatesNET, $UpdatesSSU)) {
                if (Test-Path $folder) {
                    Get-ChildItem -Path $folder -File -ErrorAction SilentlyContinue |
                        Remove-Item -Force
                }
            }
        }
        return
    }

    if ($DryRun) {
        Write-Log "[DryRun] Would fill update folders: $UpdatesOSCU, $UpdatesNET, $UpdatesSSU"
        return
    }

    # Build queries
    $osQuery  = "Cumulative Updates for Windows $WinOS Version $Version for $Arch-based Systems"
    $netQuery = ".NET Framework for Windows $WinOS Version $Version $Arch"

    Write-Host "Searching for OS updates..."
    Write-Verbose "OS Query:  $osQuery"
    $osGuids = Search-UpdateCatalogHtml -Query $osQuery
    Write-Host "Searching for .NET updates..."
    Write-Verbose ".NET Query: $netQuery"
    $netGuids = Search-UpdateCatalogHtml -Query $netQuery

    $allGuids = ($osGuids + $netGuids) | Select-Object -Unique
    Write-Host ("Found {0} total updates to process" -f $allGuids.Count)
    Write-Debug " GUID list:`n$($allGuids -join "`n")"

    if (-not $allGuids -or $allGuids.Count -eq 0) {
        Write-Host "No updates found."
        return @()
    }

    Write-Host "Retrieving update details..."

    $details = @()
    $count = 0
    foreach ($g in $allGuids) {
        $count++
        try {
            $d = Get-UpdateDetails -Count $count -Guid $g
            if ($d -and $d.DownloadUrls -and $d.DownloadUrls.Count -gt 0) {
                $details += $d
            }
            else {
                Write-Verbose "No download URLs for $g"
            }
        }
        catch {
            Write-Warning ("Failed to resolve details for {0}: {1}" -f $g, $($_.Exception.Message))
        }
    }

    if ($details.Count -eq 0) {
        Write-Host "No usable updates after details resolution."
        return @()
    }

    # Supersedence filtering
    $supersededSet = @{}
    foreach ($d in $details) {
        foreach ($s in $d.SupersededBy) {
            $supersededSet[$d.Guid] = $true
        }
    }

    $effective = $details | Where-Object { -not $supersededSet.ContainsKey($_.Guid) }

    Write-Host ("Remaining updates after supersedence filtering: {0}" -f $effective.Count)
    Write-Debug "Effective GUIDs:`n$($effective.Guid -join "`n")"

    if ($effective.Count -eq 0) {
        Write-Host "No effective updates after supersedence filtering."
        return @()
    }

    Write-Host "Synchronizing update folders..."

    $requiredFiles = @()
    foreach ($d in $effective) {
        foreach ($url in $d.DownloadUrls) {
            $requiredFiles += (Split-Path $url -Leaf)
        }
    }
    $requiredFiles = $requiredFiles | Select-Object -Unique

    # Sync: remove stale files in all folders
    foreach ($folder in @($UpdatesOSCU, $UpdatesNET, $UpdatesSSU)) {
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

    $manifestByFolder = @{
        $UpdatesOSCU = @()
        $UpdatesNET  = @()
        $UpdatesSSU  = @()
    }

    foreach ($d in $effective) {
        $targetFolder = switch -Regex ($d.Title) {
                'Servicing Stack Update' { $UpdatesSSU }
                '\.NET'                  { $UpdatesNET }
                default                  { $UpdatesOSCU }
            }

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

    Write-Host "KB update workflow complete."
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
            Write-Log "Removing: $WinpeDriverRoot"
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
            Write-Log "Removing: $RegistryRoot"
            if (Test-Path $RegistryRoot) { Remove-Item $RegistryRoot -Recurse -Force }
        }
        return
    }

    if (-not $DryRun -and -not (Test-Path $RegistryRoot)) {
        New-Item -ItemType Directory -Path $RegistryRoot -Force | Out-Null
    }

    Write-Log "Exporting Registry keys..."
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
            Write-Log "Export: $key -> $dest"
            reg.exe export "$key" "$dest" /y | Out-Null
        }
    }
}

# ==============================
# Install Drivers.cmd
# ==============================
function Write-InstallDriversScript {
    param([string]$RootFolder, [switch]$Clean)

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

    if ($Clean) {
        if ($DryRun) {
            Write-Log "[DryRun] Would remove: $path"
        } else {
            Write-Log "Removing: $path"
            if (Test-Path $path) { Remove-Item $path -Force }
        }
        return
    }

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
    param([string]$ScriptsRoot, [switch]$Clean)

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

    if ($Clean) {
        if ($DryRun) {
            Write-Log "[DryRun] Would remove: $path"
        } else {
            Write-Log "Removing: $path"
            if (Test-Path $path) { Remove-Item $path -Force }
        }
        return
    }

    if ($DryRun) {
        Write-Log "[DryRun] Would write: $path"
    } else {
        Write-Log "Writing: $path"
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
    param([string]$RootFolder, [switch]$Clean)

    $cleanPath   = Join-Path $RootFolder 'SetupConfig-Clean.ini'
    $upgradePath = Join-Path $RootFolder 'SetupConfig-Upgrade.ini'

    $cleanini = @'
[SetupConfig]
Auto=Clean
DynamicUpdate=Enable
Telemetry=Disable
'@

    $upgradeini = @'
[SetupConfig]
Auto=Upgrade
DynamicUpdate=Enable
Telemetry=Disable
'@

    if ($Clean) {
        if ($DryRun) {
            Write-Log "[DryRun] Would remove: $cleanPath"
            Write-Log "[DryRun] Would remove: $upgradePath"
        } else {
            Write-Log "Removing: $cleanPath"
            if (Test-Path $cleanPath) { Remove-Item $cleanPath -Force }
            Write-Log "Removing: $upgradePath"
            if (Test-Path $upgradePath) { Remove-Item $upgradePath -Force }
        }
        return
    }

    if ($DryRun) {
        Write-Log "[DryRun] Would write: $cleanPath"
        Write-Log "[DryRun] Would write: $upgradePath"
    } else {
        Write-Log "Writing: $cleanPath"
        Set-Content -LiteralPath $cleanPath   -Value $cleanini   -Encoding ASCII
        Write-Log "Writing: $upgradePath"
        Set-Content -LiteralPath $upgradePath -Value $upgradeini -Encoding ASCII
    }
}

# Real work starts here

# Apply defaults
if (-not $Folder) { $Folder = (Get-Location).ProviderPath }
if (-not $WinOS) { $WinOS = '11' }
if (-not $Arch)  { $Arch  = 'x64' }

if (-not $Version) {
    if ($WinOS -eq '10') { $Version = '22H2' }
    else                 { $Version = '25H2' }
}

# --- HtmlAgilityPack bootstrap (PS 5.x SAFE) ---------------------------------
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
    Write-Verbose "Downloading via WebClient..."
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
    Write-Verbose "HtmlAgilityPack.dll copied to: $hapDll"

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
    Write-Debug "HtmlAgilityPack successfully loaded."
}
# ---------------------------------------------------------------------------

# ==============================
# Resolve working folder
# ==============================
$Folder = (Resolve-Path -LiteralPath $Folder).ProviderPath
Write-Verbose "Resolved working folder: $Folder"

# ==============================
# Determine work modes
# ==============================
$workSwitches = @()
if ($KB)      { $workSwitches += 'KB' }
if ($Drivers) { $workSwitches += 'Drivers' }
if ($Reg)     { $workSwitches += 'Reg' }
if ($Files)   { $workSwitches += 'Files' }

if (-not $workSwitches) {
    # All if no specific components selected
    $KB = $true
    $Drivers = $true
    $Reg = $true
    $Files = $true
    $workSwitches = @('All')
}

Write-Log "Target profile: Windows $WinOS $Version $Arch"
Write-Log "Root folder   : $Folder"
Write-Log "Mode          : $($workSwitches -join ', ')"
if ($Clean)  { Write-Log "Clean mode    : Enabled" "WARN" }
if ($DryRun) { Write-Log "Dry-run mode  : Enabled" "WARN" }

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

# ==============================
# Main orchestration
# ==============================
if ($KB) {
    Invoke-KBWork -WinOS $WinOS -Version $Version -Arch $Arch `
        -UpdatesOSCU $paths.UpdatesOSCU -UpdatesNET $paths.UpdatesNET `
        -Clean:$Clean
}

if ($Drivers) {
    Invoke-DriverWork -WinpeDriverRoot $paths.WinpeDriverRoot -Clean:$Clean
}

if ($Reg) {
    Invoke-RegWork -RegistryRoot $paths.RegistryRoot -Clean:$Clean
}

if ($Files) {
    Write-Log "Scripts and config files..."
    Write-InstallDriversScript -RootFolder $Folder -Clean:$Clean
    Write-SetupCompleteScript -ScriptsRoot $paths.ScriptsRoot -Clean:$Clean
    Write-SetupConfigFiles -RootFolder $Folder -Clean:$Clean
}

Write-Log "Completed."
