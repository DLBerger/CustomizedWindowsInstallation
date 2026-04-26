# update-git-hash.ps1 (CI-compatible)

# Use GitHub Actions SHA if running in GitHub Actions, otherwise use local git hash
$commitHash = if ($env:GITHUB_SHA) { $env:GITHUB_SHA.Substring(0,7) } else { git rev-parse --short HEAD }

$targetFile = "CustomizedWindowsInstallation.ps1"

$lines = Get-Content $targetFile
$found = $false
$newLines = $lines | ForEach-Object {
    if (-not $found -and $_ -match '^\$GitHash\s*=\s*".*"$') {
        $found = $true
        '$GitHash = "' + $commitHash + '"'
    } else {
        $_
    }
}

if ($found) {
    Set-Content -Path $targetFile -Value $newLines
}
