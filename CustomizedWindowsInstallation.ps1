<#
.SYNOPSIS
Creates a bundled Windows installation ISO by extracting an input ISO, optionally rebuilding install.wim from selected editions, refreshing the media with Dynamic Update packages (LCU, Setup DU, SafeOS DU, and related prerequisites), and generating a new bootable ISO.

.DESCRIPTION
This script operates on a folder that contains a single Windows ISO (excluding *.bundled.iso). It extracts the ISO into a stable work directory, stages the installation media tree, optionally rebuilds ISO\sources\install.wim from selected WIM indices, refreshes the media using Dynamic Update packages, and builds a new bootable ISO.

Dynamic Update (DU) alignment goal:
- Windows Setup normally contacts Microsoft endpoints early in a feature update or media-based install to acquire Dynamic Update packages, then applies those updates to installation media. These packages can include updates to Setup binaries, SafeOS/WinRE, servicing stack requirements, the latest cumulative update (LCU), and applicable drivers intended for DU.
- In environments where devices should not (or cannot) download these during setup, DU packages can be acquired from Microsoft Update Catalog and applied to the image prior to running Setup.
- This script aims to pre-stage those DU packages into the media so the resulting ISO behaves like a current, self-contained installation source with minimal additional downloads at install/upgrade time.

DU package acquisition:
- If DU-related MSU packages are missing (or if -UpdateMSUs is specified), the script uses MSCatalogLTS to search the Microsoft Update Catalog and download the appropriate packages into a `msus\<category>` directory tree located in the same folder as the source ISO. MSCatalogLTS provides commands for searching and downloading updates from the Microsoft Update Catalog.
- The downloaded packages are saved in `<isoDir>\msus\<category>\` (e.g., msus\SSU\,msus\LCU\, msus\SafeOS\, msus\SetupDU\) so they are reusable across runs and can be applied to the staged media.
- For LCUs: if multiple cumulative updates are found for the detected build, all are downloaded (in oldest-to-newest order) to support checkpoint cumulative update chains; otherwise just the latest is used.
-           if the OnlyLatestLCU option is given, only the latest applicable LCU is downloaded and applied, without checkpoint updates.
- For Setup DU, SafeOS DU, and SSU: the latest applicable package is selected, preferring the same month as the latest LCU.

DU package application targets:
Properly updating installation media involves operating on multiple target images. Microsoft identifies the primary targets as:
- WinPE (boot.wim): used to install/deploy/repair Windows.
- WinRE (winre.wim): recovery environment used for offline repair; based on WinPE.
- Windows OS image(s) (install.wim): one or more Windows editions stored in \sources\install.wim.
- The full media tree: Setup.exe and supporting media files.

This script refreshes the media by applying the DU package types Microsoft documents for Windows installation media:
- Latest Cumulative Update (LCU) (and prerequisites/checkpoints when applicable).
- Setup Dynamic Update (Setup DU): updates setup binaries/files used for feature updates and installs.
- Safe OS Dynamic Update (SafeOS DU): updates the safe operating system used for the recovery environment (WinRE).
- Servicing stack requirements: modern LCUs often embed the servicing stack; separate servicing stack packages may exist only when required.

Checkpoint cumulative updates:
- When the catalog search for the detected build number returns multiple LCU entries, all are downloaded in oldest-to-newest order to ensure the full checkpoint chain is available. The existing KB-ordered application logic applies them in the correct sequence.
- When acquiring DU packages, Microsoft guidance also recommends ensuring DU packages correspond to the same month as the latest cumulative update; if a DU package is not available for that month, use the most recently published version.

Drivers on media:
- The script creates a special folder at the root of the staged ISO named "$WinpeDriver$". Windows Setup can scan this folder for driver INF files during installation.
- Place INF-based drivers (subfolders allowed) under \$WinpeDriver$ in the final media.

SetupConfig + convenience launchers + driver installer:
- The script writes two SetupConfig files into the ISO root:
  - SetupConfig-Upgrade.ini (for in-place upgrades)
  - SetupConfig-Clean.ini (for clean installs)
- The script writes three launcher batch files into the ISO root:
  - Upgrade.cmd: runs setup.exe /auto upgrade and passes SetupConfig-Upgrade.ini via /ConfigFile
  - Clean install.cmd: runs setup.exe /auto clean and passes SetupConfig-Clean.ini via /ConfigFile
  - Install Drivers.cmd: installs drivers from $WinpeDriver$ (if present) using pnputil; intended to be run after the initial installation has completed
- SetupConfig is applied only when setup.exe is launched with /ConfigFile <path>. Microsoft documents that when running setup from media/ISO, you must include /ConfigFile to use SetupConfig.ini.
- /Auto {Clean | Upgrade} controls the automated setup mode.

Index selection:
- If no selection is provided, behavior depends on -UpdateISO:
  - Without -UpdateISO: defaults to ALL indices.
  - With -UpdateISO: defaults to EMPTY selection unless indices are explicitly specified.
- Explicit selection can be made using:
  - -Home, -Pro
  - -Indices with numbers, ranges, labels, wildcard labels (* and ?), or regex labels (re:<pattern>).

UpdateISO behavior:
- -UpdateISO reuses an existing work folder from a prior run.
- If -UpdateISO is specified and NO explicit index selection is provided (-Home/-Pro/-Indices):
  - The script does not rebuild ISO\sources\install.wim
  - The script does not service/refresh the images
  - The script does not apply DU/MSU packages
- To force rebuild (and subsequent servicing/refresh), explicitly specify indices (for example: -Pro or -Indices 6,8,10).

UpdateMSUs behavior:
- -UpdateMSUs forces the DU/MSU download logic via MSCatalogLTS, even if MSU files already exist in the msus subdirectory.
- Without -UpdateMSUs, download occurs only when DU/MSU packages are missing (none present in the msus subdirectory alongside the ISO).

MSU directory layout:
- MSU/CAB packages are downloaded into <isoDir>\msus\<category>\ subdirectories.
- Category subdirectories: SSU, LCU, SafeOS, SetupDU.
- Application targets per category:
  - install.wim (each index): SSU (prerequisites) -> LCU (checkpoint chain)
  - winre.wim (inside install.wim): SafeOS
  - boot.wim (all WinPE indices): SafeOS
  - root of ISO: SetupDU

DryRun behavior:
- With -DryRun, the script completes PREP actions needed to stage the work tree and then prints what would happen for post-PREP actions.

References:
[1] https://learn.microsoft.com/en-us/windows/deployment/update/media-dynamic-update
[2] https://github.com/Marco-online/MSCatalogLTS
[3] https://www.deploymentresearch.com/removing-applications-from-your-windows-11-image-before-and-during-deployment/
[4] https://thedotsource.com/2021/03/16/building-iso-files-with-powershell-7/
[5] https://learn.microsoft.com/en-us/windows-hardware/customize/desktop/unattend/microsoft-windows-pnpcustomizationswinpe-driverpaths
[6] https://community.spiceworks.com/t/autounattend-xml-driver-path-issue-for-windows-11-24h2-and-25h2/1244985
[7] https://github.com/wikijm/PowerShell-AdminScripts/blob/master/Miscellaneous/New-IsoFile.ps1
[8] https://www.winhelponline.com/blog/servicing-stack-diagnosis-dism-sfc/

.PARAMETER Folder
Optional. Folder to process. If omitted, the current directory is used.

.PARAMETER Debug
Optional. Enable debugging output and passed to any tools executed.

.PARAMETER Verbose
Optional. Enable verbose output and passed to any tools executed.

.PARAMETER OS
Optional. 11 is the default. 10 or 11.

.PARAMETER Version
Optional. OS dependent. Defaults are 22H2 for 10 and 25H2 for 11.

.PARAMETER Arch
Optional. OS dependent. Default is x64, but 11 supports x64 or arm64.

.PARAMETER All
Optional. Default if others are not specifically given. Performs all actions.

.PARAMETER KB
Optional. Downloads all MS Catalog updates to the specified (OS Version Arch).

.PARAMETER Drivers
Optional. Uses DISM to export all the current system drivers and later injected as required during the installation.

.PARAMETER Reg
Optional. Capture specific registry keys from the current system to be applied during the installation.

.PARAMETER Clean
-All is the default, otherwise cleans all generated output for the option(s)

.EXAMPLE
PS> .\CustomizedWindowsInstallation.ps1
Generates a customize set for files and folders for Windows 11 25H2 x64 with all the required KBs as well as the drivers and specific registry settings from the current system.

.EXAMPLE
PS> .\CustomizedWindowsInstallation.ps1 -KB
Generates or updates folders for Windows 11 25H2 x64 with all the required KBs.

.EXAMPLE
PS> .\CustomizedWindowsInstallation.ps1 D:\temp -Clean
Completely removes all generated files and folders from D:\temp.

PS> .\CustomizedWindowsInstallation.ps1 -Clean -KB
Completely removes all KB-related files and folders from the current directory.

.NOTES
- Dynamic Update packages can be acquired from Microsoft Update Catalog and applied to installation media prior to running Setup.
- Microsoft documents the DU package categories (LCU, Setup DU, SafeOS DU, servicing stack requirements) and the image targets involved in updating installation media (WinRE, OS image, WinPE, and the media tree).
- Starting with Windows 11, version 24H2, checkpoint cumulative updates might be required as prerequisites for the latest LCU.
- The "$WinpeDriver$" folder at the root of installation media can be used to provide drivers that Setup scans during installation.
#>

# ==============================
$script:Name = "CustomizedWindowsInstallation.ps1"
# ==============================

# ==============================
# git information
# ==============================
$GitHash = "7300909"

# ==============================
# Script identity
# ==============================
$script:ScriptPath = $PSCommandPath
if (-not $script:ScriptPath) { $script:ScriptPath = $MyInvocation.MyCommand.Path }

