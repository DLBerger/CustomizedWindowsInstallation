param([string]$Folder = '.\Registry')


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

#
# HELPERS
#
function RegSafeName([string]$key) {
    ($key -replace '[\\/:*?"<>|]', '_') + '.reg'
}

function RegExportEntireKey([string]$key, [string]$dest) {
    Write-Output "Export ENTIRE key: $key -> $dest"
    reg.exe export "$key" "$dest" /y | Out-Null
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
if (-not (Test-Path $Folder)) { New-Item -ItemType Directory -Path $Folder -Force | Out-Null }
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
