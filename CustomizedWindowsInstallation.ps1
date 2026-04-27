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

    [switch]$All,
    [switch]$KB,
    [switch]$Drivers,
    [switch]$Reg,
    [switch]$Clean,

    [switch]$DryRun,

    [switch]$Help
)

# git hash
$GitHash = "2d1cfa3"

if ($Help) {
    Get-Help -Full $PSCommandPath
    exit
}


# ==============================
# Helper: Write-Log
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
#   Write-Debug "[CatalogParse] HTML: $($Html.DocumentNode.InnerHtml)"

    $ids = @()

    # Look for goToDetails('GUID')
    $pattern = 'goToDetails\("([0-9A-Fa-f\-]{36})"\)'
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

function Get-UpdateLinks {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $Guid
    )

    Write-Verbose "[Details] [Get-UpdateLinks] GUID           : $Guid"

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

    Write-Verbose "[Details] [Get-UpdateLinks] Requesting DownloadDialog.aspx via POST"
    Write-Debug   "[Details] [Get-UpdateLinks] POST body:`n$($body.UpdateIDs)"

    $response = Invoke-WebRequest @params

    Write-Verbose "[Details] [Get-UpdateLinks] received $($response.RawContentLength)-byte response of content type $($response.ContentType)"

    # Normalize content for regex (remove newlines, collapse whitespace)
    $content = $response.Content -replace "www\.download\.windowsupdate", "download.windowsupdate"
    $content = $content -replace "`r?`n", ' '
    $content = $content -replace '\s+', ' '

    Write-Verbose "[Details] [Get-UpdateLinks] Normalized content length : $($content.Length)"
    Write-Debug   "[Details] [Get-UpdateLinks] Raw content (first 400 chars):`n$($content.Substring(0, [Math]::Min(400, $content.Length)))"

    # Regex: downloadInformation[<idx>].files[<idx>].url = '<url>'
    $pattern = "downloadInformation\[(\d+)\]\.files\[(\d+)\]\.url\s*=\s*'([^']*)'"
    Write-Verbose "[Details] [Get-UpdateLinks] Running regex against DownloadDialog content"
    Write-Debug   "[Details] [Get-UpdateLinks] Regex pattern: $pattern"

    $matches = [regex]::Matches(
        $content,
        $pattern,
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    if ($matches.Count -eq 0) {
        Write-Warning "[Details] [Get-UpdateLinks] No downloadInformation URL matches for $Guid (regex returned 0 matches)"
        return @()
    }

    Write-Verbose "[Details] [Get-UpdateLinks] Found $($matches.Count) download link match(es)"

    $links = foreach ($m in $matches) {
        $downloadInfoIndex = [int]$m.Groups[1].Value
        $fileIndex         = [int]$m.Groups[2].Value
        $url               = $m.Groups[3].Value

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

    Write-Verbose "[Details] [Get-UpdateLinks] Unique URLs after de-duplication: $($sorted.Count)"
    foreach ($l in $sorted) {
        Write-Debug "[Details] [Get-UpdateLinks] URL=$($l.URL) KB=$($l.KB) DI=$($l.DownloadInfoIndex) FI=$($l.FileIndex)"
    }

    return $sorted
}

function Get-UpdateDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $Guid
    )

    Write-Verbose "[Details] GUID : $Guid"

    # Use POST-based DownloadDialog parser
    $links = Get-UpdateLinks -Guid $Guid

    if (-not $links -or $links.Count -eq 0) {
        Write-Warning "[Details] No download URLs found for $Guid (Get-UpdateLinks returned 0 items)"
        return @()
    }

    Write-Verbose "[Details] Total download URLs for $Guid : $($links.Count)"

    # Project into a simple, stable shape that the rest of the pipeline can consume
    # (adjust only if the downstream code expects extra fields)
    $results = foreach ($link in $links) {
        [PSCustomObject]@{
            Guid              = $Guid
            Url               = $link.URL
            KB                = $link.KB
            DownloadInfoIndex = $link.DownloadInfoIndex
            FileIndex         = $link.FileIndex
        }
    }

    foreach ($r in $results) {
        Write-Debug "[Details] Resolved: GUID=$($r.Guid) KB=$($r.KB) URL=$($r.Url)"
    }

    return $results
}

# ==============================
# Download helper
# ==============================

function Compute-Sha256 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    if (-not (Test-Path $Path -PathType Leaf)) {
        throw "File not found for SHA-256: $Path"
    }

    $stream = [System.IO.File]::OpenRead($Path)
    try {
        $sha = [System.Security.Cryptography.SHA256]::Create()
        $hashBytes = $sha.ComputeHash($stream)
        ($hashBytes | ForEach-Object { $_.ToString("x2") }) -join ''
    }
    finally {
        $stream.Dispose()
    }
}

function Invoke-WithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock] $ScriptBlock,

        [int] $MaxAttempts = 3,
        [int] $DelaySeconds = 3
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return & $ScriptBlock
        }
        catch {
            if ($attempt -ge $MaxAttempts) {
                throw
            }
            Write-Warning "Attempt $attempt failed: $($_.Exception.Message). Retrying in $DelaySeconds seconds..."
            Start-Sleep -Seconds $DelaySeconds
        }
    }
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

    foreach ($url in $Update.Urls) {
        $fileName = Split-Path -Path $url -Leaf
        $destPath = Join-Path $TargetFolder $fileName

        Invoke-WithRetry {
            Invoke-WebRequest -Uri $url -OutFile $destPath -UseBasicParsing
        } | Out-Null

        $hash = Compute-Sha256 -Path $destPath

        # Your existing Build-ManifestEntry call here, e.g.:
        # Build-ManifestEntry -Update $Update -FilePath $destPath -Sha256 $hash
    }
}

function Get-UpdateDetails {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $true,
            Position = 0
        )]
        [string] $Guid
    )

    Write-Verbose "[Details] GUID : $Guid"

    # Build POST body
    $post = @{ size = 0; UpdateID = $Guid; UpdateIDInfo = $Guid } | ConvertTo-Json -Compress
    $body = @{ UpdateIDs = "[$post]" }

    $params = @{
        Uri         = "https://www.catalog.update.microsoft.com/DownloadDialog.aspx"
        Method      = 'POST'
        Body        = $body
        ContentType = "application/x-www-form-urlencoded"
        UseBasicParsing = $true
    }

    Write-Verbose "[Details] Requesting DownloadDialog.aspx via POST"
    Write-Debug   "[Details] POST body: $($body.UpdateIDs)"

    $response = Invoke-WebRequest @params

    if (-not $response -or -not $response.Content) {
        Write-Warning "[Details] Empty response for GUID $Guid"
        return $null
    }

    $content = $response.Content -replace "www.download.windowsupdate", "download.windowsupdate"

    Write-Verbose "[Details] Raw content length : $($content.Length)"
    Write-Debug   "[Details] Raw content (first 400 chars): $($content.Substring(0, [Math]::Min(400, $content.Length)))"

    # Regex: capture ALL downloadInformation[x].files[y].url
    $regex = "downloadInformation\[(\d+)\]\.files\[(\d+)\]\.url\s*=\s*'([^']*)'"

    Write-Verbose "[Details] Running regex against DownloadDialog content"
    Write-Debug   "[Details] Regex pattern: $regex"

    $matches = [regex]::Matches($content, $regex)

    if ($matches.Count -eq 0) {
        Write-Warning "[Details] No download URLs found for $Guid (regex returned 0 matches)"
        return $null
    }

    Write-Verbose "[Details] Found $($matches.Count) download link match(es) for $Guid"

    $linkObjects = foreach ($m in $matches) {
        $downloadInfoIndex = [int]$m.Groups[1].Value
        $fileIndex         = [int]$m.Groups[2].Value
        $url               = $m.Groups[3].Value.Trim()

        Write-Debug "[Details] Match: downloadInformation[$downloadInfoIndex].files[$fileIndex].url = $url"

        # Try to extract KB number from URL (if present)
        $kbNumber = 0
        if ($url -match 'kb(\d+)') {
            $kbNumber = [int]$Matches[1]
        }

        [PSCustomObject]@{
            Url               = $url
            KB                = $kbNumber
            DownloadInfoIndex = $downloadInfoIndex
            FileIndex         = $fileIndex
            Guid              = $Guid
        }
    }

    # De-duplicate by URL
    $unique = $linkObjects |
        Group-Object -Property Url |
        ForEach-Object { $_.Group[0] }

    Write-Verbose "[Details] Unique URLs after de-duplication: $($unique.Count)"

    # Sort by KB descending so callers can pick "best" if they only want one
    $sorted = $unique | Sort-Object KB -Descending

    return $sorted
}

function Get-TargetFolderForUpdate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject] $Details,

        [Parameter(Mandatory = $true)]
        [string] $UpdatesOSCU,

        [Parameter(Mandatory = $true)]
        [string] $UpdatesNET,

        [Parameter(Mandatory = $true)]
        [string] $UpdatesSSU
    )

    switch -Regex ($Details.Classification) {
        'Servicing Stack Update' { return $UpdatesSSU }
        '\.NET'                  { return $UpdatesNET }
        default                  { return $UpdatesOSCU }
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
        Guid          = $Details.Guid
        Title         = $Details.Title
        Classification= $Details.Classification
        DownloadUrl   = $DownloadInfo.Url
        FileName      = $DownloadInfo.FileName
        Sha256        = $DownloadInfo.Sha256
        SupersededBy  = $Details.SupersededBy
        Timestamp     = (Get-Date).ToString("s")
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

    Write-Verbose "=== KB Work ==="

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
        Write-Verbose "[KB] Clean mode: removing all existing update files and manifests..."
        foreach ($folder in @($UpdatesOSCU, $UpdatesNET, $UpdatesSSU)) {
            if (Test-Path $folder) {
                Get-ChildItem -Path $folder -File -ErrorAction SilentlyContinue |
                    Remove-Item -Force
                $manifestPath = Join-Path $folder 'manifest.json'
                if (Test-Path $manifestPath -PathType Leaf) {
                    Remove-Item $manifestPath -Force
                }
            }
        }
    }

    # Build queries
    $osQuery  = "Cumulative Updates for Windows $WinOS Version $Version for $Arch-based Systems"
    $netQuery = ".NET Framework for Windows $WinOS Version $Version $Arch"

    Write-Verbose "[KB] OS Query:  $osQuery"
    Write-Verbose "[KB] .NET Query: $netQuery"

    # Search OS updates
    Write-Verbose "[KB] Searching OS updates..."
    $osGuids = Search-UpdateCatalogHtml -Query $osQuery

    # Search .NET updates
    Write-Verbose "[KB] Searching .NET updates..."
    $netGuids = Search-UpdateCatalogHtml -Query $netQuery

    $allGuids = ($osGuids + $netGuids) | Select-Object -Unique
    Write-Verbose "[KB] Total GUIDs found: $($allGuids.Count)"
    Write-Debug   "[KB] GUID list:`n$($allGuids -join "`n")"

    if (-not $allGuids -or $allGuids.Count -eq 0) {
        Write-Verbose "[KB] No updates found."
        return @()
    }

    # Resolve details
    $details = @()
    foreach ($g in $allGuids) {
        try {
            $d = Get-UpdateDetails -Guid $g
            if ($d.DownloadUrls.Count -gt 0) {
                $details += $d
            }
            else {
                Write-Verbose "[KB] No download URLs for $g"
            }
        }
        catch {
            Write-Warning ("Failed to resolve details for {0}: {1}" -f $g, $($_.Exception.Message))
        }
    }

    if ($details.Count -eq 0) {
        Write-Verbose "[KB] No usable updates after details resolution."
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

    Write-Verbose "[KB] Effective (non-superseded) updates: $($effective.Count)"
    Write-Debug   "[KB] Effective GUIDs:`n$($effective.Guid -join "`n")"

    if ($effective.Count -eq 0) {
        Write-Verbose "[KB] No effective updates after supersedence filtering."
        return @()
    }

    # Build required file list (leaf names)
    $requiredFiles = @()
    foreach ($d in $effective) {
        foreach ($url in $d.DownloadUrls) {
            $requiredFiles += (Split-Path $url -Leaf)
        }
    }
    $requiredFiles = $requiredFiles | Select-Object -Unique

    # Sync: remove stale files in all folders
    foreach ($folder in @($UpdatesOSCU, $UpdatesNET, $UpdatesSSU)) {
        Write-Verbose "[Sync] Checking folder: $folder"

        $existingFiles = @()
        if (Test-Path $folder) {
            $existingFiles = Get-ChildItem -Path $folder -File |
                             Select-Object -ExpandProperty Name
        }

        $stale = $existingFiles | Where-Object { $_ -notin $requiredFiles }
        foreach ($file in $stale) {
            $path = Join-Path $folder $file
            Write-Verbose "[Sync] Removing stale file: $file"
            Remove-Item $path -Force
        }
    }

    # Download + manifest building per folder
    $manifestByFolder = @{
        $UpdatesOSCU = @()
        $UpdatesNET  = @()
        $UpdatesSSU  = @()
    }

    $results = @()

    foreach ($d in $effective) {
        $targetFolder = Get-TargetFolderForUpdate -Details $d -UpdatesOSCU $UpdatesOSCU -UpdatesNET $UpdatesNET -UpdatesSSU $UpdatesSSU

        foreach ($url in $d.DownloadUrls) {
            $fileName = Split-Path $url -Leaf
            $destPath = Join-Path $targetFolder $fileName

            if (Test-Path $destPath -PathType Leaf) {
                Write-Verbose "[Sync] Already present: $fileName"
                $sha = Compute-Sha256 -Path $destPath
                $downloadInfo = [pscustomobject]@{
                    FileName = $fileName
                    FullPath = $destPath
                    Sha256   = $sha
                    Url      = $url
                }
            }
            else {
                $downloadInfo = Download-MUFile -Url $url -DestinationFolder $targetFolder
            }

            $entry = Build-ManifestEntry -Details $d -DownloadInfo $downloadInfo
            $manifestByFolder[$targetFolder] += $entry
            $results += $entry
        }
    }

    # Write manifests
    foreach ($kvp in $manifestByFolder.GetEnumerator()) {
        $folder = $kvp.Key
        $entries = $kvp.Value
        if ($entries.Count -gt 0) {
            Write-Verbose "[Manifest] Writing manifest for $folder"
            Write-Manifest -Folder $folder -Entries $entries
        }
        else {
            # If no entries, remove stale manifest if present
            $manifestPath = Join-Path $folder 'manifest.json'
            if (Test-Path $manifestPath -PathType Leaf) {
                Remove-Item $manifestPath -Force
            }
        }
    }

    Write-Verbose "[KB] KB work complete."
    return $results
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

# Real work starts here

# If -Debug was passed, force debug output to auto-continue
if ($PSBoundParameters.ContainsKey('Debug')) {
    $DebugPreference = 'Continue'
    Write-Debug "Debug mode enabled: DebugPreference set to 'Continue'"
}

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

if (-not ($Clean -and -not ($KB -or $Drivers -or $Reg))) {
    Write-InstallDriversScript -RootFolder $Folder
    Write-SetupCompleteScript -ScriptsRoot $paths.ScriptsRoot
    Write-SetupConfigFiles -RootFolder $Folder
}

Write-Log "Completed."
Write-Verbose "Script execution finished."
