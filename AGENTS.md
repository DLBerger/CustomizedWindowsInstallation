# AGENTS.md

## Cursor Cloud specific instructions

### About this codebase

This is a **Windows PowerShell 5.1** toolkit (`CustomizedWindowsInstallation.ps1`, ~3,280 lines) that customizes Windows 10/11 installation media. It downloads KB updates from the Microsoft Update Catalog, services WIM images with DISM, exports drivers, applies registry tweaks, and creates a final bootable ISO.

The codebase has **no package manager, no build system, and no Docker/container setup**. All files are standalone `.ps1` and `.cmd` scripts.

### Development on Linux (Cloud Agent VM)

The scripts **cannot execute end-to-end on Linux** because they depend on Windows-only tools (`dism.exe`, `oscdimg.exe`) and Windows APIs (e.g., `WindowsPrincipal` admin check). Development on Linux is limited to:

- **Linting**: `pwsh -NoProfile -Command "Invoke-ScriptAnalyzer -Path '<script>.ps1' -Severity Error,Warning"`
- **Syntax validation**: `pwsh -NoProfile -Command "[System.Management.Automation.Language.Parser]::ParseFile('<script>.ps1', [ref]\$null, [ref]\$null)"`
- **Usage/Help output**: `pwsh -NoProfile -Command '& "./CustomizedWindowsInstallation.ps1" -Usage'` (the `WindowsPrincipal` error on line 264 is expected on Linux and non-blocking)
- **Code editing and git operations**

### Coding rules

- **Strictly target PowerShell 5.1**: no ternary (`? :`), no null-coalescing (`??`), no pipeline chain operators (`&&`, `||`), no `Parallel` in loops.
- **Always use `Write-Host`** for user-facing output (never `Write-Output`), per the SOW and script notes.
- **Use `"$($var)"`** when concatenating variables with special characters.
- The `.cursor/rules/powershell-51.mdc` file has additional mandatory standards.

### Lint expectations

- **0 errors** is the baseline for all scripts.
- **Warnings** like `PSAvoidUsingWriteHost` (159 in the main script) are **intentional** and expected per the SOW — the script explicitly requires `Write-Host` for pipeline-safe output.

### Key scripts

| Script | Purpose |
|---|---|
| `CustomizedWindowsInstallation.ps1` | Main toolkit (~3,280 lines) |
| `ExportRegs.ps1` | Registry key export utility |
| `Update.NET.ps1` | .NET runtime update utility |
| `.git-scripts/update-git-hash-*.ps1` | Git commit hash embedding (local + GitHub Actions) |

### Git hooks

The repo uses a post-commit hook (`.git-scripts/update-git-hash-local.ps1`) to embed the git hash into `$GitHash` in the main script. This only works on Windows with PowerShell 5.1; on Linux it is not active.
