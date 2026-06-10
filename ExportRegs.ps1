param([string]$Folder = '.\Registry')


# An empty list in the Values means work on the entire key. '*' and '?' are supported as wildcards.
# Try to keep in alphabetical order.
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
        Key    = 'HKEY_CURRENT_USER\Control Panel\NotifyIconSettings\*'
        Values = @('IsPromoted')
    },
    @{
        Key    = 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings'
        Values = @('TaskbarEndTask')
    },
    @{
        Key    = 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
        Values = @(
                   'HideFileExt',
                   'LaunchTo',
                   'SeparateProcess',
                   'ShowTaskViewButton',
                   'Start_IrisRecommendations',
                   'UseCompactMode'
                   )
    },
    @{
        Key    = 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState'
        Values = @('FullPath')
    },
    @{
        Key    = 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Start'
        Values = @('AllAppsViewMode')
    },
    @{
        Key    = 'HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\Explorer'
        Values = @('HideRecommendedSection')
    },
    @{
        Key    = 'HKEY_LOCAL_MACHINE\Software\Microsoft\WindowsUpdate\UX\Settings'
        Values = @('AllowMUUpdateService')
    },
    @{
        Key    = 'HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\FileSystem'
        Values = @('LongPathsEnabled')
    },
    @{
        Key    = 'HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Power'
        Values = @('HibernateEnabled')
    },
    @{
        Key    = 'HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\Explorer'
        Values = @('HideRecommendedSection')
    },
    @{
        Key    = 'HKEY_LOCAL_MACHINE\Software\Microsoft\PolicyManager\current\device\Start'
        Values = @('HideRecommendedSection')
    }
)

$RegistryRemove = @(
)

#
# HELPERS
#
function RegSafeName([string]$key) {
    ($key -replace '[\\/:*?"<>|]', '_') + '.reg'
}

function Convert-PsPathToHKey([string]$psPath) {
    if ($psPath -match 'Registry::(.+)$') {
        return $matches[1]
    }
    return $psPath
}

function Expand-RegistryKeyPatternSegments([string]$psPath, [string[]]$segments) {
    if (-not $segments -or $segments.Count -eq 0) {
        return @(Convert-PsPathToHKey $psPath)
    }

    $segment = $segments[0]
    $tail = @()
    if ($segments.Count -gt 1) {
        $tail = $segments[1..($segments.Count - 1)]
    }

    $found = @()
    if ($segment -match '[\*\?]') {
        if (Test-Path -LiteralPath $psPath) {
            Get-ChildItem -LiteralPath $psPath -ErrorAction SilentlyContinue |
                Where-Object { $_.PSChildName -like $segment } |
                ForEach-Object {
                    $found += Expand-RegistryKeyPatternSegments $_.PSPath $tail
                }
        }
    } else {
        $next = Join-Path $psPath $segment
        if (Test-Path -LiteralPath $next) {
            $found += Expand-RegistryKeyPatternSegments $next $tail
        }
    }

    return $found
}

function Expand-RegistryKeyPattern([string]$keyPattern) {
    if ($keyPattern -notmatch '[\*\?]') {
        return @($keyPattern)
    }

    $parts = $keyPattern -split '\\', 2
    $hive = $parts[0]
    $rest = $parts[1]

    $psRoot = $null
    switch ($hive) {
        'HKEY_CURRENT_USER'   { $psRoot = 'HKCU:' }
        'HKEY_LOCAL_MACHINE'  { $psRoot = 'HKLM:' }
        'HKEY_CLASSES_ROOT'   { $psRoot = 'HKCR:' }
        'HKEY_USERS'          { $psRoot = 'HKU:' }
        'HKEY_CURRENT_CONFIG' { $psRoot = 'HKCC:' }
    }

    if (-not $psRoot) {
        Write-Output "WARNING: Unknown hive in key pattern: $keyPattern"
        return @()
    }

    $segments = $rest -split '\\'
    return Expand-RegistryKeyPatternSegments $psRoot $segments
}

function Format-RegValueLine([string]$valName, [string]$type, [string]$data) {
    switch ($type) {
        "SZ" {
            return '"' + $valName + '"="' + $data + '"'
        }
        "DWORD" {
            return '"' + $valName + '"=dword:' + ("{0:x8}" -f [int]$data)
        }
        default {
            return '"' + $valName + '"="' + $data + '"'
        }
    }
}

function RegExportEntireKey([string]$keyPattern, [string]$dest) {
    $keys = @(Expand-RegistryKeyPattern $keyPattern)
    if (-not $keys -or $keys.Count -eq 0) {
        Write-Output "WARNING: No keys matched pattern: $keyPattern"
        return
    }

    if ($keys.Count -eq 1) {
        Write-Output "Export ENTIRE key: $($keys[0]) -> $dest"
        reg.exe export "$($keys[0])" "$dest" /y | Out-Null
        return
    }

    Write-Output "Export ENTIRE key: $($keys.Count) key(s) matching $keyPattern -> $dest"

    $out = New-Object System.Collections.Generic.List[string]
    $out.Add("Windows Registry Editor Version 5.00")
    $out.Add("")

    $temp = [System.IO.Path]::GetTempFileName()
    foreach ($key in ($keys | Sort-Object -Unique)) {
        Write-Output "  Exporting $key"
        reg.exe export "$key" "$temp" /y | Out-Null
        if (-not (Test-Path $temp)) {
            continue
        }

        $started = $false
        foreach ($line in (Get-Content $temp)) {
            if (-not $started) {
                if ($line -match '^\[') {
                    $started = $true
                    $out.Add($line)
                }
            } else {
                $out.Add($line)
            }
        }
    }

    Remove-Item $temp -Force -ErrorAction SilentlyContinue

    if ($out.Count -gt 2) {
        $out -join "`r`n" | Set-Content -Path $dest -Encoding Unicode
    }
}

function RegExportSpecificValues([string]$keyPattern, [string]$dest, [object[]]$groups) {

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
        Write-Output "No specific values requested for $keyPattern; skipping."
        return
    }

    $keys = @(Expand-RegistryKeyPattern $keyPattern)
    if (-not $keys -or $keys.Count -eq 0) {
        Write-Output "WARNING: No keys matched pattern: $keyPattern"
        return
    }

    Write-Output "Export specific values [$($allValues -join ', ')] from $($keys.Count) key(s) matching $keyPattern -> $dest"

    $out = New-Object System.Collections.Generic.List[string]
    $out.Add("Windows Registry Editor Version 5.00")
    $out.Add("")

    foreach ($key in ($keys | Sort-Object -Unique)) {

        $query = reg.exe query "$key" /v * 2>$null
        if (-not $query) {
            Write-Output "WARNING: No data returned for $key"
            continue
        }

        $sectionLines = New-Object System.Collections.Generic.List[string]
        $sectionLines.Add("[" + $key + "]")

        foreach ($line in $query) {

            if ($line -match '^\s+([^\s]+)\s+REG_([A-Z0-9_]+)\s+(.*)$') {

                $valName = $matches[1]
                $type    = $matches[2]
                $data    = $matches[3]

                if ($allValues -contains $valName) {
                    $sectionLines.Add((Format-RegValueLine $valName $type $data))
                }
            }
        }

        if ($sectionLines.Count -gt 1) {
            foreach ($l in $sectionLines) { $out.Add($l) }
            $out.Add("")
        }
    }

    if ($out.Count -le 2) {
        Write-Output "WARNING: No matching values found for pattern: $keyPattern"
        return
    }

    $out -join "`r`n" | Set-Content -Path $dest -Encoding Unicode
}

function RegAppendDelete([string]$dest, [string]$key, [string[]]$values) {

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

# ==============================
# Resolve working folder and make sure it exists
# ==============================

# If the user passed a relative path (like ".\Registry"), resolve it relative to the script location
if (-not ([System.IO.Path]::IsPathRooted($Folder))) {
    $Folder = Join-Path $PSScriptRoot $Folder
}

if (-not (Test-Path $Folder)) {
    New-Item -ItemType Directory -Path $Folder -Force | Out-Null
}

$Folder = (Resolve-Path -LiteralPath $Folder -ErrorAction Continue).ProviderPath
Write-Output "Resolved working folder: $Folder"

#
# PROCESS ADD/MODIFY
#
foreach ($entry in $RegistryAddModify) {

    $key    = $entry.Key
    $groups = $entry.Values

    $safe = "AddModify_" + (RegSafeName $key)
    $dest = Join-Path $Folder $safe

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
        RegExportEntireKey $key $dest
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
    $dest = Join-Path $Folder $safe

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

    $keys = @(Expand-RegistryKeyPattern $key)
    if (-not $keys -or $keys.Count -eq 0) {
        Write-Output "WARNING: No keys matched pattern: $key"
        continue
    }

    if ($hasEntire) {
        foreach ($matchedKey in ($keys | Sort-Object -Unique)) {
            RegAppendDelete $dest $matchedKey @()
        }
    } else {
        $allValues = @()
        foreach ($g in $groups) {
            if (-not $g) { continue }
            foreach ($v in $g) {
                if ($null -ne $v -and $v -ne '') { $allValues += $v }
            }
        }
        $allValues = $allValues | Sort-Object -Unique

        foreach ($matchedKey in ($keys | Sort-Object -Unique)) {
            RegAppendDelete $dest $matchedKey $allValues
        }
    }
}
