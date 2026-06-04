param(
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

# Root folder containing manifest.json and payloads
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$root       = Join-Path $scriptRoot 'KBs\NET'
$manifest   = Join-Path $root 'manifest.json'

if (-not (Test-Path $manifest)) {
    Write-Error "manifest.json not found at $manifest"
}

Write-Host ">>> Loading manifest database..." -ForegroundColor Cyan
$items = Get-Content $manifest -Raw | ConvertFrom-Json

# Detect installed runtimes and SDKs
Write-Host ">>> Scanning local system for installed .NET environments..." -ForegroundColor Cyan
$runtimesX64 = @()
$runtimesX86 = @()
$sdksX64     = @()
$sdksX86     = @()
$hasHosting  = $false

try { $runtimesX64 = & dotnet.exe --list-runtimes 2>$null } catch {}
try { $sdksX64     = & dotnet.exe --list-sdks 2>$null } catch {}

$dotnetX86 = 'C:\Program Files (x86)\dotnet\dotnet.exe'
if (Test-Path $dotnetX86) {
    try { $runtimesX86 = & $dotnetX86 --list-runtimes 2>$null } catch {}
    try { $sdksX86     = & $dotnetX86 --list-sdks 2>$null } catch {}
}

# Hosting bundle detection
$hostingKey = 'HKLM:\SOFTWARE\Microsoft\IIS Extensions\IIS AspNetCore Module V2'
if (Test-Path $hostingKey) { 
    $hasHosting = $true 
    Write-Host "    Found IIS ASP.NET Core Module V2 registration." -ForegroundColor Gray
}

# Helper to check runtime presence and get maximum installed version
function Get-MaxInstalledRuntimeVersion {
    param(
        [string]$Flavor,
        [string]$Arch
    )
    $src = if ($Arch -eq 'x86') { $runtimesX86 } else { $runtimesX64 }
    if (-not $src) { return $null }
    
    # Parse out versions matching the flavor (e.g., "Microsoft.NETCore.App [8.0.4]")
    $versions = $src | Where-Object { $_ -match [regex]::Escape($Flavor) } | ForEach-Object {
        if ($_ -match '(\d+\.\d+\.\d+)') { [version]$Matches[1] }
    }
    if (-not $versions) { return $null }
    return ($versions | Measure-Object -Maximum).Maximum
}

# Helper to get maximum installed SDK major/minor matching version
function Get-MaxInstalledSdkVersion {
    param(
        [string]$Arch
    )
    $src = if ($Arch -eq 'x86') { $sdksX86 } else { $sdksX64 }
    if (-not $src) { return $null }
    
    $versions = $src | ForEach-Object {
        if ($_ -match '^(\d+\.\d+\.\d+)') { [version]$Matches[1] }
    }
    if (-not $versions) { return $null }
    return ($versions | Measure-Object -Maximum).Maximum
}

Write-Host ""
Write-Host "=========================================================" -ForegroundColor Yellow
Write-Host " Evaluating applicable .NET updates..." -ForegroundColor Yellow
Write-Host "=========================================================" -ForegroundColor Yellow
Write-Host ""

$applied = @()

foreach ($item in $items) {
    $fileName = $item.FileName
    if (-not $fileName) { continue }

    $fullPath = Join-Path $root $fileName
    if (-not (Test-Path $fullPath)) {
        Write-Warning "File listed in manifest but not found in storage directory: $fileName"
        continue
    }

    $name = $fileName.ToLowerInvariant()
    $shouldInstall = $false
    $skipReason = ""

    # Pull target file metadata version for smart comparison if it's an executable
    $fileVersion = $null
    if ($name.EndsWith('.exe')) {
        try {
            $fileVersionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($fullPath)
            if ($fileVersionInfo.ProductVersion -match '^(\d+\.\d+\.\d+)') {
                $fileVersion = [version]$Matches[1]
            }
        } catch {
            Write-Verbose "Could not extract version from $fileName"
        }
    }

    # Determine Architecture context
    $archContext = if ($name -like '*-x86*' -or $name -like '*-win-x86*') { 'x86' } else { 'x64' }

    # 1. Desktop Runtime
    if ($name -like '*windowsdesktop-runtime*') {
        $maxLocal = Get-MaxInstalledRuntimeVersion 'Microsoft.WindowsDesktop.App' $archContext
        if ($null -ne $maxLocal) {
            if ($fileVersion -and $maxLocal -ge $fileVersion) {
                $skipReason = "Current or newer version ($maxLocal) already installed on system"
            } else { $shouldInstall = $true }
        } else {
            $skipReason = "Base WindowsDesktop Runtime ($archContext) is not active on this machine"
        }
    }

    # 2. Core Runtime
    elseif ($name -like '*dotnet-runtime*') {
        $maxLocal = Get-MaxInstalledRuntimeVersion 'Microsoft.NETCore.App' $archContext
        if ($null -ne $maxLocal) {
            if ($fileVersion -and $maxLocal -ge $fileVersion) {
                $skipReason = "Current or newer version ($maxLocal) already installed on system"
            } else { $shouldInstall = $true }
        } else {
            $skipReason = "Base .NET Core Runtime ($archContext) is not active on this machine"
        }
    }

    # 3. ASP.NET Core Runtime
    elseif ($name -like '*aspnetcore-runtime*') {
        $maxLocal = Get-MaxInstalledRuntimeVersion 'Microsoft.AspNetCore.App' $archContext
        if ($null -ne $maxLocal) {
            if ($fileVersion -and $maxLocal -ge $fileVersion) {
                $skipReason = "Current or newer version ($maxLocal) already installed on system"
            } else { $shouldInstall = $true }
        } else {
            $skipReason = "Base ASP.NET Core Runtime ($archContext) is not active on this machine"
        }
    }

    # 4. SDKs
    elseif ($name -like '*dotnet-sdk*') {
        $maxLocal = Get-MaxInstalledSdkVersion $archContext
        if ($null -ne $maxLocal) {
            # Check major/minor version boundary alignment
            if ($fileVersion) {
                if ($maxLocal -ge $fileVersion) {
                    $skipReason = "Current or newer SDK version ($maxLocal) already installed"
                } elseif ($maxLocal.Major -eq $fileVersion.Major) {
                    # Same major band but older system patch, clear for upgrade
                    $shouldInstall = $true
                } else {
                    $skipReason = "SDK version belongs to a different generation band"
                }
            } else { $shouldInstall = $true }
        } else {
            $skipReason = "No existing .NET SDK installation detected for architecture $archContext"
        }
    }

    # 5. Hosting Bundle
    elseif ($name -like '*dotnet-hosting*') {
        if ($hasHosting) {
            # Compare using the ASP.NET Core engine version bundled inside the installer
            $maxLocal = Get-MaxInstalledRuntimeVersion 'Microsoft.AspNetCore.App' 'x64'
            if ($fileVersion -and $maxLocal -and $maxLocal -ge $fileVersion) {
                $skipReason = "Hosting bundle components are already current ($maxLocal)"
            } else { $shouldInstall = $true }
        } else {
            $skipReason = "IIS Environment / AspNetCore Module V2 is not enabled on host machine"
        }
    }

    # 6. .NET Framework MSU (Windows 11 Servicing)
    elseif ($name -like '*ndp481*') {
        # Check if KB payload is already tracked inside WUSA hotfix catalog
        if ($name -match 'kb(\d+)') {
            $kbId = $Matches[1]
            if (Get-HotFix -Id "KB$kbId" -ErrorAction SilentlyContinue) {
                $skipReason = "Windows Hotfix KB$kbId is already active on this operating system"
            } else { $shouldInstall = $true }
        } else {
            $shouldInstall = $true
        }
    }

    # Process Actions
    if ($shouldInstall) {
        if ($DryRun) {
            Write-Host "[DryRun] Would upgrade/install target: $fileName" -ForegroundColor DarkYellow
        }
        else {
            Write-Host ">>> Processing patch execution for: $fileName" -ForegroundColor White

            if ($name.EndsWith('.msu')) {
                Write-Host "    Executing Windows Update Standalone Installer..." -ForegroundColor Gray
                Start-Process wusa.exe -ArgumentList "`"$fullPath`" /quiet /norestart" -Wait
            }
            elseif ($name.EndsWith('.cab')) {
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
                    Write-Host "    Injecting Driver package components..." -ForegroundColor Gray
                    Start-Process pnputil.exe -ArgumentList "/add-driver `"$fullPath`" /install" -Wait
                }
                else {
                    Write-Host "    Injecting Operating System Update package via DISM deployment engine..." -ForegroundColor Gray
                    Start-Process dism.exe -ArgumentList "/online /add-package /packagepath:`"$fullPath`" /quiet /norestart" -Wait
                }
            }
            else {
                Write-Host "    Spawning background silent package deployment thread..." -ForegroundColor Gray
                # Using Start-Process instead of direct invocation to protect execution pipelines from hanging on background mutexes
                Start-Process -FilePath $fullPath -ArgumentList "/quiet /norestart" -Wait -NoNewWindow
            }
            
            Write-Host "    Done processing: $fileName" -ForegroundColor Green
        }
        $applied += $fileName
    } else {
        if ($skipReason) {
            Write-Host "[-] Skipping: $fileName" -ForegroundColor Gray
            Write-Host "    Reason: $skipReason" -ForegroundColor DarkGray
        }
    }
}

Write-Host ""
Write-Host "=========================================================" -ForegroundColor Yellow
Write-Host " Processing Summary" -ForegroundColor Yellow
Write-Host "=========================================================" -ForegroundColor Yellow

if (-not $applied) {
    Write-Host "All components match or exceed manifest targets. No operations needed." -ForegroundColor Green
} else {
    if ($DryRun) {
        Write-Host "DryRun evaluation phase finalized. The following assets would be updated:" -ForegroundColor Yellow
    } else {
        Write-Host "Successfully analyzed and processed the following changes:" -ForegroundColor Green
    }
    $applied | ForEach-Object { Write-Host "  [+] $_" -ForegroundColor Cyan }
}
Write-Host ""
