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
$GitHash = "68df975"

# ==============================
# Core names
# ==============================
$names = [ordered]@{
    Iso                   = 'ISO'
    KBs                   = 'KBs'
    Wims                  = 'Wims'
    WinpeDriver           = '$WinpeDriver$'
    Registry              = 'Registry'
    InstallEsd            = 'install.esd'
    InstallWim            = 'install.wim'
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

$wimDirs = @('Temp', 'Mounts', 'Captures')
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
# WIM export
# =========================

function Invoke-ExportWork {
    Write-Log "WIM export not implemented yet"
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
# ISO generation
# ==============================
function Invoke-IsoWork {
    Write-Log "ISO generation not implemented yet"
<#
# -------------------- Configuration --------------------
$OriginalInstallWim = 'C:\images\install.wim'
$WorkRoot           = 'C:\images\working'          # temp working root
$TempWimFolder      = $paths.WimsTemp
$MountRoot          = $paths.WimsMounts
$CaptureRoot        = $paths.WimsCaptures
$FinalInstallWim    = Join-Path $WorkRoot 'final_install.wim'
$IndicesToProcess   = 1,6
$CompressionType    = 'Maximum'                    # None, Fast, Maximum
$Packages           = @('C:\updates\kb1.cab','C:\updates\kb2.msu')  # list of CAB/MSU paths
$DriverFolders      = @('C:\drivers\intel','C:\drivers\others')     # driver folders (Recurse)
$MinFreeSpaceBytes  = 50GB                         # minimum free space required (adjust)
$LogFile            = Join-Path $WorkRoot 'service.log'
# -------------------------------------------------------

# Helper: convert friendly size to bytes
function Convert-ToBytes($size) {
    if ($size -is [string]) {
        $s = $size.ToUpper().Trim()
        if ($s -match '(\d+)\s*GB') { return [int64]$matches[1] * 1GB }
        if ($s -match '(\d+)\s*MB') { return [int64]$matches[1] * 1MB }
        return [int64]$s
    }
    return [int64]$size
}

# Logging
function Log($msg) {
    $t = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "$t`t$msg"
    $line | Tee-Object -FilePath $LogFile -Append
}

# Prepare folders
New-Item -Path $WorkRoot -ItemType Directory -Force | Out-Null
New-Item -Path $TempWimFolder -ItemType Directory -Force | Out-Null
New-Item -Path $MountRoot -ItemType Directory -Force | Out-Null
New-Item -Path $CaptureRoot -ItemType Directory -Force | Out-Null
if (Test-Path $LogFile) { Remove-Item $LogFile -Force }

Log "Starting servicing workflow"
Log "Source WIM: $OriginalInstallWim"
Log "Indices: $($IndicesToProcess -join ',')"

# Validate source WIM
if (-not (Test-Path $OriginalInstallWim)) {
    Log "Source install.wim not found: $OriginalInstallWim"
    Write-Error "Source install.wim not found: $OriginalInstallWim"
    exit 1
}

# Check available disk space on WorkRoot drive
$drive = Get-Item $WorkRoot
$free = (Get-PSDrive -Name $drive.PSDrive.Name).Free
$minBytes = Convert-ToBytes $MinFreeSpaceBytes
if ($free -lt $minBytes) {
    Log "Insufficient free space on drive $($drive.PSDrive.Name): $free bytes free, require at least $minBytes"
    Write-Error "Insufficient free space on drive $($drive.PSDrive.Name). Free: $free bytes. Required: $minBytes bytes."
    exit 1
}
Log "Free space check OK: $free bytes available"

# Function to run dism.exe fallback
function Run-DismFallback {
    param($ArgsArray)
    $argLine = $ArgsArray -join ' '
    Log "Running dism.exe $argLine"
    $proc = Start-Process -FilePath 'dism.exe' -ArgumentList $ArgsArray -Wait -NoNewWindow -PassThru
    if ($proc.ExitCode -ne 0) {
        throw "dism.exe failed with exit code $($proc.ExitCode) for args: $argLine"
    }
}

# Export indices to per-index uncompressed WIMs in parallel
$exportJobs = @()
foreach ($idx in $IndicesToProcess) {
    $destWim = Join-Path $TempWimFolder ("{0}.wim" -f $idx)
    $exportJobs += Start-Job -Name "ExportIndex$idx" -ArgumentList $OriginalInstallWim, $idx, $destWim, $hasExportCmdlet -ScriptBlock {
        param($srcWim, $index, $destWim, $useExportCmdlet)
        try {
            $args = @('/Export-Image', "/SourceImageFile:$srcWim", "/SourceIndex:$index", "/DestinationImageFile:$destWim", '/Compress:None')
            Start-Process -FilePath 'dism.exe' -ArgumentList $args -Wait -NoNewWindow -PassThru | Out-Null
            "Exported index $index to $destWim (dism.exe)"
            }
        } catch {
            throw "Export failed for index $index : $($_.Exception.Message)"
        }
    }
}
Log "Started export jobs: $($exportJobs | ForEach-Object { $_.Name } -join ', ')"

# Wait and collect export results
Wait-Job -Job $exportJobs
$exportErrors = @()
foreach ($j in $exportJobs) {
    $state = $j.State
    $out = Receive-Job -Job $j -ErrorAction SilentlyContinue
    if ($state -ne 'Completed') {
        $exportErrors += ,@{ Job = $j.Name; State = $state; Output = $out }
    } else {
        Log ("Job {0} completed: {1}" -f $j.Name, ($out -join '; '))
    }
    Remove-Job -Job $j -Force -ErrorAction SilentlyContinue
}
if ($exportErrors.Count -gt 0) {
    $exportErrors | ForEach-Object { Log ("Export error: {0} {1}" -f $_.Job, $_.Output) }
    Write-Error "One or more exports failed. See log $LogFile"
    exit 1
}

# For each exported WIM, mount, service, copy to capture dir, then unmount. Do this in parallel jobs.
$serviceJobs = @()
foreach ($idx in $IndicesToProcess) {
    $srcWim = Join-Path $TempWimFolder ("{0}.wim" -f $idx)
    $mountDir = Join-Path $MountRoot ("mount_{0}" -f $idx)
    $captureDir = Join-Path $CaptureRoot ("capture_{0}" -f $idx)
    New-Item -Path $mountDir -ItemType Directory -Force | Out-Null
    New-Item -Path $captureDir -ItemType Directory -Force | Out-Null

    $serviceJobs += Start-Job -Name "ServiceIndex$idx" -ArgumentList $srcWim, $idx, $mountDir, $captureDir, $hasMountCmdlet, $hasAddPkgCmdlet, $hasAddDrvCmdlet, $Packages, $DriverFolders -ScriptBlock {
        param($srcWim, $index, $mountDir, $captureDir, $useMountCmdlet, $useAddPkgCmdlet, $useAddDrvCmdlet, $packages, $driverFolders)

        function LocalLog($m) { $t=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); "$t`t[$index] $m" | Out-File -FilePath (Join-Path $mountDir 'job.log') -Append }

        try {
            LocalLog "Starting service job for index $index"

            # Mount image
            $args = @('/Mount-Image', "/ImageFile:$srcWim", "/Index:1", "/MountDir:$mountDir")
            Start-Process -FilePath 'dism.exe' -ArgumentList $args -Wait -NoNewWindow -PassThru | Out-Null
            LocalLog "Mounted $srcWim to $mountDir (dism.exe)"

            # Apply packages
            foreach ($pkg in $packages) {
                if (-not (Test-Path $pkg)) { LocalLog "Package not found: $pkg"; continue }
                $args = @('/Image:' + $mountDir, '/Add-Package', "/PackagePath:$pkg")
                Start-Process -FilePath 'dism.exe' -ArgumentList $args -Wait -NoNewWindow -PassThru | Out-Null
                LocalLog "Added package $pkg (dism.exe)"
            }

            # Add drivers
            foreach ($drv in $driverFolders) {
                if (-not (Test-Path $drv)) { LocalLog "Driver folder not found: $drv"; continue }
                $args = @('/Image:' + $mountDir, '/Add-Driver', "/Driver:$drv", '/Recurse')
                Start-Process -FilePath 'dism.exe' -ArgumentList $args -Wait -NoNewWindow -PassThru | Out-Null
                LocalLog "Added drivers from $drv (dism.exe)"
            }

            # Commit and unmount
            $args = @('/Unmount-Image', "/MountDir:$mountDir", '/Commit')
            Start-Process -FilePath 'dism.exe' -ArgumentList $args -Wait -NoNewWindow -PassThru | Out-Null
            LocalLog "Dismounted and committed $mountDir (dism.exe)"

            # Capture the image tree to captureDir (use image capture via dism or robocopy copy of mounted tree)
            # We will expand the WIM by mounting then copying the mounted tree; since we unmounted above, remount read-only to copy
            $args = @('/Mount-Image', "/ImageFile:$srcWim", "/Index:1", "/MountDir:$mountDir")
            Start-Process -FilePath 'dism.exe' -ArgumentList $args -Wait -NoNewWindow -PassThru | Out-Null
            robocopy $mountDir $captureDir /MIR /NFL /NDL /NJH /NJS | Out-Null
            $args = @('/Unmount-Image', "/MountDir:$mountDir", '/Discard')
            Start-Process -FilePath 'dism.exe' -ArgumentList $args -Wait -NoNewWindow -PassThru | Out-Null
            LocalLog "Copied mounted tree to $captureDir and dismounted (dism.exe)"

            # Clean up the temporary per-index WIM
            Remove-Item -Path $srcWim -Force -ErrorAction SilentlyContinue
            LocalLog "Removed temporary WIM $srcWim"

            "Service job for index $index completed successfully"
        } catch {
            LocalLog "ERROR: $_"
            throw "Service job failed for index $index : $($_.Exception.Message)"
        } finally {
            # ensure mount dir is removed if empty
            if (Test-Path $mountDir) {
                try { Remove-Item -Path $mountDir -Recurse -Force -ErrorAction SilentlyContinue } catch {}
            }
        }
    }
}

Log "Started service jobs: $($serviceJobs | ForEach-Object { $_.Name } -join ', ')"

# Wait for service jobs
Wait-Job -Job $serviceJobs
$serviceErrors = @()
foreach ($j in $serviceJobs) {
    $state = $j.State
    $out = Receive-Job -Job $j -ErrorAction SilentlyContinue
    if ($state -ne 'Completed') {
        $serviceErrors += ,@{ Job = $j.Name; State = $state; Output = $out }
        Log ("Service job {0} failed: {1}" -f $j.Name, ($out -join '; '))
    } else {
        Log ("Service job {0} completed: {1}" -f $j.Name, ($out -join '; '))
    }
    Remove-Job -Job $j -Force -ErrorAction SilentlyContinue
}
if ($serviceErrors.Count -gt 0) {
    Write-Error "One or more service jobs failed. See log $LogFile"
    exit 1
}

# Build final compressed install.wim
# Capture first capture dir into final_install.wim, then append others
$captures = Get-ChildItem -Path $CaptureRoot -Directory | Sort-Object Name
if ($captures.Count -eq 0) {
    Log "No capture directories found in $CaptureRoot"
    Write-Error "No capture directories found"
    exit 1
}

# Capture first
$first = $captures[0].FullName
$args = @('/Capture-Image', "/ImageFile:$FinalInstallWim", "/CaptureDir:$first", "/Name:ServicedImage1", "/Compress:Maximum")
Run-DismFallback $args
Log "Captured $first to $FinalInstallWim (dism.exe)"

# Append remaining captures
$counter = 2
foreach ($cap in $captures | Select-Object -Skip 1) {
    $name = "ServicedImage$counter"
    $args = @('/Append-Image', "/ImageFile:$FinalInstallWim", "/CaptureDir:$($cap.FullName)", "/Name:$name")
    Run-DismFallback $args
    Log "Appended $($cap.FullName) as $name to $FinalInstallWim"
    $counter++
}

Log "Final install.wim created at $FinalInstallWim"

# Cleanup capture dirs if desired (commented out)
# Remove-Item -Path $CaptureRoot -Recurse -Force

Log "Workflow completed successfully"
Write-Output "Completed successfully. Final WIM: $FinalInstallWim"

#>
}

# ==============================
# Registry export
# ==============================
function Invoke-RegWork {

    $RegistryRoot = $paths.RegistryRoot
    if ($Clean) {
        if ($DryRun) {
            Write-Log "[DryRun] Would remove: $RegistryRoot"
        } elseif (Test-Path $RegistryRoot) {
            Write-Log "Removing: $RegistryRoot"
            Remove-Item $RegistryRoot -Recurse -Force
        }
        return
    }

    if (-not $DryRun) {
        Write-Log "Exporting Registry keys..."
        Ensure-Folder -Path $RegistryRoot
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
            Write-Log "Export: $key -> $dest"
            reg.exe export "$key" "$dest" /y | Out-Null
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

# Apply defaults
if (-not $Folder) { $Folder = (Get-Location).ProviderPath }
if (-not $WinOS) { $WinOS = '11' }
if (-not $Arch)  { $Arch  = 'x64' }

if (-not $Version) {
    if ($WinOS -eq '10') { $Version = '22H2' }
    else                 { $Version = '25H2' }
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
if ($Export)  { $workSwitches += 'Export' }
if ($KB)      { $workSwitches += 'KB' }
if ($Drivers) { $workSwitches += 'Drivers' }
if ($Reg)     { $workSwitches += 'Reg' }
if ($Service) { $workSwitches += 'Service' }
if ($Files)   { $workSwitches += 'Files' }

if (-not $workSwitches) {
    # All if no specific components selected
    $Export = $true
    $KB = $true
    $Drivers = $true
    $Reg = $true
    $Service = $true
    $Files = $true
    $workSwitches = @('All')
}

Write-Log "Target profile: Windows $WinOS $Version $Arch"
Write-Log "Root folder   : $Folder"
Write-Log "Mode          : $($workSwitches -join ', ')"
if ($Clean)  { Write-Log "Clean mode    : Enabled" "WARN" }
if ($DryRun) { Write-Log "Dry-run mode  : Enabled" "WARN" }

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
# Core paths
# ==============================
$paths = [ordered]@{}
$paths.IsoRoot               = Join-Path $Folder $names.Iso
$paths.WinpeDriverRoot       = Join-Path $Folder $names.WinpeDriver
$paths.RegistryRoot          = Join-Path $Folder $names.Registry
$paths.InstallDriversCmd     = Join-Path $Folder $names.InstallDriversCmd
$paths.InstallRegsCmd        = Join-Path $Folder $names.InstallRegsCmd
$paths.PostSetupCmd          = Join-Path $Folder $names.PostSetupCmd
$paths.SetupConfigCleanIni   = Join-Path $Folder $names.SetupConfigCleanIni
$paths.SetupConfigUpgradeIni = Join-Path $Folder $names.SetupConfigUpgradeIni
$paths.CleanInstallCmd       = Join-Path $Folder $names.CleanInstallCmd
$paths.UpgradeCmd            = Join-Path $Folder $names.UpgradeCmd
$paths.KBsRoot               = Join-Path $Folder $names.KBs
foreach ($u in $kbDirs) {
    $paths["KBs$u"]          = Join-Path $paths.KBsRoot $names.$u
}
$paths.WimsRoot              = Join-Path $Folder $names.Wims
foreach ($u in $wimDirs) {
    $paths["Wims$u"]         = Join-Path $paths.WimsRoot $names.$u
}

# ==============================
# Main orchestration
# ==============================

if ($Export) { Invoke-ExportWork }
if ($KB) { Invoke-KBWork }
if ($Drivers) { Invoke-DriverWork }
if ($Reg) { Invoke-RegWork }
if ($Service) { Invoke-IsoWork }

if ($Files) {
    Write-InstallDriversCmd
    Write-InstallRegsCmd
    Write-PostSetupCmd
    Write-SetupConfigFiles
    Write-SetupCmdFiles
}

Write-Log "Completed"
