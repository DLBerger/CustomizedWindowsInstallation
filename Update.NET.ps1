param(
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

# Root folder containing manifest.json and payloads
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$root       = Join-Path $scriptRoot 'KBs\Net'
$manifest   = Join-Path $root 'manifest.json'

if (-not (Test-Path $manifest)) {
    Write-Error "manifest.json not found at $manifest"
}

# Load manifest
$items = Get-Content $manifest -Raw | ConvertFrom-Json

# Detect installed runtimes and SDKs
$runtimesX64 = @()
$runtimesX86 = @()
$sdkVersions = @()
$hasHosting  = $false

try { $runtimesX64 = & dotnet.exe --list-runtimes 2>$null } catch {}

$dotnetX86 = 'C:\Program Files (x86)\dotnet\dotnet.exe'
if (Test-Path $dotnetX86) {
    try { $runtimesX86 = & $dotnetX86 --list-runtimes 2>$null } catch {}
}

try {
    $sdkVersions = (& dotnet.exe --list-sdks 2>$null) |
        ForEach-Object { ($_ -split '\s+')[0] }
} catch {}

# Hosting bundle detection
$hostingKey = 'HKLM:\SOFTWARE\Microsoft\IIS Extensions\IIS AspNetCore Module V2'
if (Test-Path $hostingKey) { $hasHosting = $true }

# Helper to check runtime presence
function Test-HasRuntime {
    param(
        [string]$Flavor,
        [string]$Arch
    )
    $src = if ($Arch -eq 'x86') { $runtimesX86 } else { $runtimesX64 }
    if (-not $src) { return $false }
    return $src -match [regex]::Escape($Flavor)
}

Write-Host ""
Write-Host "Evaluating applicable .NET updates..."
Write-Host ""

$applied = @()

foreach ($item in $items) {
    $fileName = $item.FileName
    if (-not $fileName) { continue }

    $fullPath = Join-Path $root $fileName
    if (-not (Test-Path $fullPath)) {
        Write-Warning "File listed in manifest but not found: $fileName"
        continue
    }

    $name = $fileName.ToLowerInvariant()
    $shouldInstall = $false

    # Desktop runtime
    if ($name -like '*windowsdesktop-runtime*') {
        if ($name -like '*-win-x64*' -and (Test-HasRuntime 'Microsoft.WindowsDesktop.App' 'x64')) { $shouldInstall = $true }
        if ($name -like '*-win-x86*' -and (Test-HasRuntime 'Microsoft.WindowsDesktop.App' 'x86')) { $shouldInstall = $true }
    }

    # Core runtime
    elseif ($name -like '*dotnet-runtime*') {
        if ($name -like '*-win-x64*' -and (Test-HasRuntime 'Microsoft.NETCore.App' 'x64')) { $shouldInstall = $true }
        if ($name -like '*-win-x86*' -and (Test-HasRuntime 'Microsoft.NETCore.App' 'x86')) { $shouldInstall = $true }
    }

    # ASP.NET Core runtime
    elseif ($name -like '*aspnetcore-runtime*') {
        if ($name -like '*-win-x64*' -and (Test-HasRuntime 'Microsoft.AspNetCore.App' 'x64')) { $shouldInstall = $true }
        if ($name -like '*-win-x86*' -and (Test-HasRuntime 'Microsoft.AspNetCore.App' 'x86')) { $shouldInstall = $true }
    }

    # SDKs
    elseif ($name -like '*dotnet-sdk*') {
        foreach ($v in $sdkVersions) {
            if ($name -like "*dotnet-sdk-$v-win-*") {
                $shouldInstall = $true
                break
            }
        }
    }

    # Hosting bundle
    elseif ($name -like '*dotnet-hosting*') {
        if ($hasHosting) { $shouldInstall = $true }
    }

    # .NET Framework MSU (always applicable on Win11)
    elseif ($name -like '*ndp481*') {
        $shouldInstall = $true
    }

    if ($shouldInstall) {
        if ($DryRun) {
            Write-Host "[DryRun] Would install: $fileName"
        }
        else {
            Write-Host "Installing: $fileName"

            if ($name.EndsWith('.msu')) {
                # MSU requires wusa.exe
                Start-Process wusa.exe -ArgumentList "`"$fullPath`" /quiet /norestart" -Wait
            }
            elseif ($name.EndsWith('.cab')) {
                # Determine if it's a driver CAB or update CAB
                $isDriver = $false
                $tempCab = Join-Path $env:TEMP "cabtest_$([guid]::NewGuid().ToString())"
                New-Item -ItemType Directory -Path $tempCab -Force | Out-Null

                try {
                    expand.exe "$fullPath" -F:* "$tempCab" 2>$null | Out-Null
                    if (Get-ChildItem "$tempCab" -Filter *.inf -ErrorAction SilentlyContinue) {
                        $isDriver = $true
                    }
                } catch {}

                Remove-Item $tempCab -Recurse -Force -ErrorAction SilentlyContinue

                if ($isDriver) {
                    # Driver CAB → pnputil
                    Start-Process pnputil.exe -ArgumentList "/add-driver `"$fullPath`" /install" -Wait
                }
                else {
                    # Update CAB → DISM
                    Start-Process dism.exe -ArgumentList "/online /add-package /packagepath:`"$fullPath`" /quiet /norestart" -Wait
                }
            }
            else {
                # EXE installer
                & $fullPath /quiet /norestart 2>&1 | Out-Null
            }
        }

        $applied += $fileName
    }
}

if (-not $applied) {
    Write-Host ""
    Write-Host "No applicable .NET updates found for this system."
} else {
    Write-Host ""
    if ($DryRun) {
        Write-Host "DryRun complete. These updates would be applied:"
    } else {
        Write-Host "Applied:"
    }
    $applied | ForEach-Object { Write-Host "  $_" }
}