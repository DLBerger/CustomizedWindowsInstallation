Create a complete PowerShell 5.x master script that implements the following:

Required external programs:
  - dism.exe - needs to be auto-discovered from the ADK or system or overridden with on of -UseADK, -UseSystem, or -dism
  - oscdimg.exe - needs to be auto-discovered from the ADK or system or overridden with on of -UseADK, -UseSystem, or -oscdimg
  - robocopy.exe - always use the system routine

Parameters:
  - Folder
    Optional. Work folder. If omitted, the current directory is used.

  - Help
    Displays help and exits.

  - DryRun
    Show actions without performing them.

  - Clean
    Remove generated content instead of creating it.
    This is a modifier to whatever additional parameters are given.
    Its order in the argument list is not important.

  - ISO
    Explicit path to source ISO.
    If not provided find the single .iso file in <Folder>.
    If more than one .iso present, give an error, and suggest using this parameter.

  - DestISO
    Explicit path to destination ISO.
    If not provided, use <ISO>, substituting .bundled.iso for .iso.

  - WinOS or OS
    Windows major version: '10' or '11'.
    If omitted, determined from the contents of <ISO>.

  - Version
    Windows feature update version (for example: '22H2', '25H2').
    If omitted, determined from the contents of <ISO>.

  - Arch
    CPU architecture: 'x64' or 'arm64'.
    If omitted, determined from the contents of <ISO>.

  - Export
    Mount and copy the contents of <ISO> to <Folder> and then export the requested indices (see below).

  - KB
    Download OS and .NET updates based on <WinOS>, <Version>, and <Arch>.

  - Service
    Apply the download KBs to the exported indices and create the final install.wim and boot.wim.

  - Drivers
    Export drivers into $WinpeDriver$.

  - Reg
    Export registry keys.

  - Files
    Generate .cmd and .ini files.

  - All
    Shorthand for -Export -KB -Service -Drivers -Reg -Files and the default if no specific switch is provided.

  - Home
    Select editions whose normalized label matches "Home" exactly.

  - Pro
    Select editions whose normalized label matches "Pro" exactly.

  - Indices
    Comma-separated selector string supporting:
      numbers: 6
      ranges: 3-6, 7-*
      exact labels: "Education N"
      wildcard labels: "*Home*", "* N*"
      regex labels: "re:^Education( N)?$"

  - ShowIndices
    Shows available image indices (index and name) and exits.

  - UpdateISO
    Reuses an existing work folder. If used without explicit indices, no rebuild/servicing/DU actions occur.

  - UseADK
    Prefer ADK DISM and oscdimg tools when available.

  - UseSystem
    Force system DISM and PATH oscdimg.

  - dism
    Explicit path to dism.exe.

  - oscdimg
    Explicit path to oscdimg.exe.

Examples:
  - .\CustomizedWindowsInstallation.ps1
    Set <Folder> to the current directory.
    Set <ISO> to the single .iso in <Folder>.
    Create a working folder in <Folder>.
    Perform every action and finally new .bundled.iso back into <Folder>.

  - .\CustomizedWindowsInstallation.ps1 D:\temp
    Set <Folder> to D:\temp.
    Set <ISO> to the single .iso in <Folder>.
    Create a working folder in <Folder>.
    Perform every action and finally new .bundled.iso back into <Folder>.

  - .\CustomizedWindowsInstallation.ps1 -Clean
    Cleans everything

  - .\CustomizedWindowsInstallation.ps1 -KB -Clean
  - .\CustomizedWindowsInstallation.ps1 -Clean -KB
    Cleans all KB-related file and folders
    Placement of the -Clean in not important as a single -Clean changes the action from a create to a remove.

  - .\CustomizedWindowsInstallation.ps1 -UpdateISO
    Reuses the existing work folder and performs no rebuild/servicing/DU actions (no indices were specified).

  - .\CustomizedWindowsInstallation.ps1 -UpdateISO -Pro
    Reuses the existing work folder and forces rebuild/servicing/DU refresh using the Pro edition selection.

  - .\CustomizedWindowsInstallation.ps1 -KB
    Forces DU/MSU downloads into the working folder

  - .\CustomizedWindowsInstallation.ps1 -Indices "* N*"
    Selects all N editions (quote required due to space).

  - .\CustomizedWindowsInstallation.ps1 -ShowIndices
    Shows install.wim indices and exits.

Index selection:
  - Explicit selection can be made using:
      -Home, -Pro
      -Indices with numbers, ranges, labels, wildcard labels (* and ?), or regex labels (re:<pattern>).
  - If no selection is provided, behavior depends on -UpdateISO:
      With -UpdateISO: defaults to EMPTY selection unless indices are explicitly specified.
      Without -UpdateISO: defaults to ALL indices.

Primary assumptions:
  - Temporary files:
      Written under <Folder> and retained.
      Deleted during relevant cleaning operations.
  - Checkpointing and restartability:
      Every major step writes a small .done file in the per-index work folder: export.done, boot.<index>.done, winre.<index>.done, <index>.done, kb.done, drivers.done, reg.done, files.done 
      On resume the script skips steps with checkpoints.
  - Logging and safety:
      Elevated check, disk-space validation, per-job logs, robust mount/unmount handling, conservative cleanup so interrupted runs remain resumable.
      Use dism for wim servicing.
      Use robocopy for copying trees.
      Use oscdimg for final creation of the ISO.
      Provide clear error messages and nonzero exit codes on failure.
  - Concurrency:
      Allow configurable max parallel jobs
  - Provide a clear configuration header at the top (paths, indices, packages, driver root, WinPE subfolder name, BootWimMode=PerIndex, ServiceWinREInPlace flag, compression, concurrency, MinFreeSpaceBytes, logging path) and usage notes showing how to resume after Ctrl-C, how to run Clean, and a short validation checklist (test modified boot.wim in VM, verify WinRE presence, verify final install.wim contents).
  - Include inline comments and per-index job-local logs. Keep intermediate artifacts unless Clean is run. Ensure idempotence and deterministic behavior.
  - Task context:
      PowerShell 5.x
      Lots of fast storage available.
      User will supply a single driver root folder produced by dism /export-driver.
      The winre.wim to service is located at Windows\System32\Recovery\winre.wim inside each index's wim.
  - Naming conventions for files and paths we may want to change easily:
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

Suggested Export processing:
  - Ensure there is a folder under <Folder> to copy the .iso contents to.
  - Mount and copy the contents of the source .iso under <Folder> and mark iso.copy.done.
  - Verify boot.wim and (install.esd or install.wim) exits in the iso contents, if not fail.
  - Ensure there is a folder under <Folder> to copy the original .wims to.
  - Check if boot.wim.move.done, if not:
      Move boot.wim for iso contents folder to original .wims folder and mark boot.wim.move.done.
  - Check if install.esd.move.done, if not:
      If install.esd exists in iso contents folder move to to original .wims folder.
      Mark install.esd.move.done.
  - Check if install.wim.move.done, if not:
      If install.wim exists in iso contents folder then move to original .wims folder
      else check if install.esd.move.done, if not fail with a consistency error
        Convert install.esd to install.wim
      Mark boot.wim.move.done.
   Export the requested indices in parallel from boot.wim and install.wim into per-index uncompressed WIMs named <index>_boot.wim and <index>_install.wim into a wims folder.
     Check each <index>_boot.wim.extracted and <index>_install.wim.extracted and only extract if they don't exist.
     At the end of each wim extraction mark <index>_boot.wim.extracted or <index>_install.wim.extracted, respectively.

Suggested KB processing:
  - TBD - (Just accept what is already there in the existing script)

Suggested Servive processing:
  - For each index (parallel jobs):
      Check for <index>_install.done, if not:
        Mount <index>_install.wim to a per-index mount folder.
        Check for <index>_winre.done, if not:
          Locate Windows\System32\Recovery\winre.wim inside the mounted tree. If present, extract it to <index>_winre.wim and mark checkpoint <index>_winre.extracted.
          Apply KB package file to the mounted install image with dism.
          Unmount and commit the winre image.
          Reinsert the serviced <index>_winre.wim back into the mounted install image at Windows\System32\Recovery\winre.wim and mark <index>_winre.done.
        Apply KB package file to the mounted install image with dism.
        Unmount and commit the install image and mark <index>_install.done.
      Check for <index>_boot.done, if not:
        Mount <index>_boot.wim to a per-index mount folder.
        Apply KB package file to the mounted install image with dism.
        Unmount and commit the install image and mark <index>_boot.done.
  - Final assembly:
      Do the following steps serially as the final compression is going to be slow if Maximum is selected.
      Check for install.wim.done, if not: 
        Take the first <index>_install.wim and copy to the final compressed install.wim with configurable compression (None, Fast, Maximum).
        Append all the remaining <index>_install.wim to the final install.wim
        After the final install.wim is finished, mark install.wim.done
      Check for boot.wim.done, if not: 
        Take the first <index>_boot.wim and copy to the final compressed boot.wim with configurable compression (None, Fast, Maximum).
        Append all the remaining <index>_boot.wim to the final boot.wim.
        After the final boot.wim is finished, mark boot.wim.done
  - Final step:
      Check for install.wim.servining.done, if not: 
        Copy the final install.wim back into its iso contents place and mark install.wim.servicing.done
      Check for boot.wim.servining.done, if not: 
        Copy the final boot.wim back into its iso contents place and mark boot.wim.servicing.done

Deliverable: a single PowerShell 5.x script file content and brief usage notes. Include examples of checkpoint files and resume commands.

*) WinPE driver selection rules: prefer newest DriverVer when multiple INFs; log selected INFs and their classes; selection is deterministic and configurable.


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

Drivers on media:
- The script creates a special folder at the root of the staged ISO named "$WinpeDriver$". Windows Setup can scan this folder for driver INF files during installation.
- INF-based drivers (subfolders allowed) under \$WinpeDriver$ will be in the final media.

SetupConfig + convenience launchers + driver installer:
- The script writes two SetupConfig files into the ISO root:
  - SetupConfig-Upgrade.ini (for in-place upgrades)
  - SetupConfig-Clean.ini (for clean installs)
- The script writes three launcher batch files into the ISO root:
  - Upgrade.cmd: runs setup.exe /auto upgrade and passes SetupConfig-Upgrade.ini via /ConfigFile
  - CleanInstall.cmd: runs setup.exe /auto clean and passes SetupConfig-Clean.ini via /ConfigFile
  - InstallDrivers.cmd: installs drivers from $WinpeDriver$ using pnputil; intended to be run after the initial installation has completed

UpdateISO behavior:
- -UpdateISO reuses an existing work folder from a prior run.
- If -UpdateISO is specified and NO explicit index selection is provided (-Home/-Pro/-Indices):
  - The script does not rebuild ISO\sources\install.wim
  - The script does not service/refresh the images
  - The script does not apply DU/MSU packages
- To force rebuild (and subsequent servicing/refresh), explicitly specify indices (for example: -Pro or -Indices 6,8,10).

UpdateKBs behavior:
- -UpdateKBs forces the DU/MSU download logic, even if MSU files already exist in the msus subdirectory.
- Without -UpdateKBs, download occurs only when DU/MSU packages are missing (none present in the msus subdirectory alongside the ISO).

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

Normal, Verbose, and Debug output behavior:
- All output must be pipeline-able as the output will often be piped to tee to be captured.
- Don't duplicate any output from a previous normal, verbose, or debug output.
- Progress bars:
    Normal externally run programs like robocopy and dism output is not required, but the progress bar or some type of progress bar is.
    No blue powershell progress bars.
    When robocopy is being used or files are being downloaded, a progress bar showing changes in 10% increments is preferred.
- Normal operating output behavior with neither -Verbose or -Debug is to give a description of what is happening so the user can see a steady stream of information
    State information for major iterators, especially time consuming ones.
    The user needs to see the system working.
- With -Verbose, in addition to the normal output, the parameter lists to important functions need to be output.
    More detailed state information for important iterators needs to be output.
- With -Debug, in addition to the verbose output:
    Parameter lists for all functions need to be output.
    Fields within iterators need to be displayed as compactly as possible.
    All output from externally run programs needs to be visible like robocopy and dism

Add these references:
[1] https://learn.microsoft.com/en-us/windows/deployment/update/media-dynamic-update
[2] https://github.com/Marco-online/MSCatalogLTS
[3] https://www.deploymentresearch.com/removing-applications-from-your-windows-11-image-before-and-during-deployment/
[4] https://thedotsource.com/2021/03/16/building-iso-files-with-powershell-7/
[5] https://learn.microsoft.com/en-us/windows-hardware/customize/desktop/unattend/microsoft-windows-pnpcustomizationswinpe-driverpaths
[6] https://community.spiceworks.com/t/autounattend-xml-driver-path-issue-for-windows-11-24h2-and-25h2/1244985
[7] https://github.com/wikijm/PowerShell-AdminScripts/blob/master/Miscellaneous/New-IsoFile.ps1
[8] https://www.winhelponline.com/blog/servicing-stack-diagnosis-dism-sfc/