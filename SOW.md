# CustomizedWindowsInstallation.ps1 — Statement of Work

PowerShell **5.1** master script that customizes Windows 10/11 installation media: extract a source ISO, export WIM indices, download updates from the Microsoft Update Catalog, service images with DISM, assemble final `install.wim` / `boot.wim`, stage a destination tree, copy helper scripts, and build a bootable ISO with `oscdimg.exe`.

**Deliverable:** `CustomizedWindowsInstallation.ps1` (single file) plus companion `.cmd` / `.ps1` files in the repo root that `-Files` copies into the ISO root.

---

Required external programs:
  - dism.exe
    WIM mount, export, and package apply. Required. Auto-discovered from the ADK or system, or overridden with -dism.
    Use -UseADK, -UseSystem, or an explicit path to control discovery.

  - oscdimg.exe
    Final ISO creation (-CreateISO). Optional until ISO build. Same discovery rules as dism (-oscdimg, -UseADK, -UseSystem).

  File and folder trees are copied in-script (Stream-FileCopy, Stream-FolderCopy) or hardlinked during Prep.

  The script must be run elevated as Administrator.

  HtmlAgilityPack.dll is downloaded into <Folder> when -KB runs (NuGet package, loaded from a byte array to avoid file locks).

---

Script-level configuration (top of file; not parameters):
  - $GitHash
    Embedded by git hook. Shown in the Show-Usage banner.

  - $LeadIn
    Scriptblock formatting status lines as '{0,-20}: {1}'.

  - $MaxIORetries (default 3)
    Retries for Stream-FileCopy and Download-MUFile.

  - $BufferSize (default 64MB)
    Stream copy and download buffer size.

  - $ProgressPrecentage (default 10)
    Report progress in 10% increments.

  - $PercentWidth (default 3)
    Column width for percent in progress lines (fits 0–100%).

  - $Bucket, $Hex
    Scriptblocks for progress bucketing and hex exit codes.

  - $DismCompression (default 'max')
    Passed to DISM /Compress: on final export (none, fast, max).

  - $names
    Ordered hashtable of logical folder and file names (see Naming and paths below).

  - $kbDirs
    SSU, OSCU, NET, MISC — KB download subfolders under KBs\.

  - $wimDirs
    Indices, Mounts, Serviced, Final, Scratch, Logs — WIM work subfolders under Wims\.

---

Parameters:
  Positional parameters (in order): <Folder>, <ISO>, <DestISO>.

  - Folder
    Optional. Work folder. Default is the current directory (`.\`), resolved to a full path.
    Positional parameter 0.

  - ISO
    Alias: SrcISO.
    Optional. Explicit path to the source ISO.
    If not provided, find the single .iso file in <Folder>.
    If more than one .iso is present, error and suggest using this parameter.
    Positional parameter 1.

  - DestISO
    Optional. Explicit path to the destination ISO file.
    If not provided and <ISO> is set, use <ISO> with the extension changed to _KBs.iso.
    Positional parameter 2.

  - Help
    Displays full help (Get-Help -Full) and exits.

  - Usage
    Displays the short usage banner (Show-Usage) and exits.

  - DryRun
    Show actions without performing them.

  - Clean
    Remove generated content instead of creating it.
    This is a modifier for whichever work parameters are also given.
    Its order in the argument list is not important.

  - WinOS
    Alias: OS.
    Windows major version: '10' or '11'.
    If omitted, determined from WIM/ISO metadata when available; otherwise defaults to '11'.

  - Version
    Windows feature update version (for example: '22H2', '25H2').
    If omitted, determined from WIM build when available:
      Windows 10 -> '22H2'
      Windows 11 -> '25H2'

  - Arch
    CPU architecture: 'x64' or 'arm64'.
    If omitted, determined from WIM metadata when available; otherwise defaults to 'x64'.

  - Extract
    Alias: ExtractISO.
    Mount the source ISO and extract its full content tree to <Folder>\SrcISO\Content\.

  - Export
    Alias: ExportWims.
    Export the selected install.wim indices (and all boot.wim indices) into per-index
    uncompressed WIMs under <Folder>\Wims\Indices\.

  - KB
    Download OS and .NET updates from the Microsoft Update Catalog into <Folder>\KBs\.

  - Service
    Apply downloaded packages (SSU from KBs\SSU, LCU from KBs\OSCU) to the exported indices
    and write serviced WIMs to <Folder>\Wims\Serviced\.

  - Final
    Build compressed install.wim and boot.wim under <Folder>\Wims\Final\ from the serviced indices.

  - Prep
    Alias: PrepDestISO.
    Hardlink-copy SrcISO\Content to DestISO\Content (excluding install/boot images), then copy
    the final WIMs from Wims\Final into DestISO\Content\sources\.

  - Files
    Create $WinPEDriver$ and Registry folders on the destination ISO and copy helper
    .cmd, .ps1, and .ini files (and optional KBs\NET, KBs\MISC trees) into DestISO\Content\.

  - CreateISO
    Build the final .iso file from <DestISO>\Content using oscdimg.exe.

  - All
    Shorthand for -Extract -Export -KB -Service -Final -Prep -Files -CreateISO.
    Default when no specific work switch is provided.

  - Most
    Same as -All except without -CreateISO.

  - SelectHome
    Alias: Home.
    Select editions whose normalized label matches "Home" exactly.

  - SelectPro
    Alias: Pro.
    Select editions whose normalized label matches "Pro" exactly.

  - Indices
    Comma-separated selector string supporting:
      numbers: 6
      ranges: 3-6, 7-*
      exact labels: "Education N"
      wildcard labels: "*Home*", "* N*"
      regex labels: "re:^Education( N)?$"

  - ShowIndices
    Shows available install.wim indices (index and name) and exits.

  - UseADK
    Prefer ADK DISM and oscdimg tools when available.

  - UseSystem
    Force system DISM and PATH oscdimg.

  - dism
    Explicit path to dism.exe.

  - oscdimg
    Explicit path to oscdimg.exe.

  - Verbose
    Standard PowerShell verbose stream (parameter lists for important functions, more detail).

  - Debug
    Standard PowerShell debug stream; also sets $DebugPreference to Continue at startup.

Examples:
  - .\CustomizedWindowsInstallation.ps1
    Set <Folder> to the current directory.
    Set <ISO> to the single .iso in <Folder>.
    Run every work step and finally write <ISO>_KBs.iso.

  - .\CustomizedWindowsInstallation.ps1 D:\temp
    Set <Folder> to D:\temp.
    Set <ISO> to the single .iso in <Folder>.
    Run every work step and finally write the derived _KBs.iso.

  - .\CustomizedWindowsInstallation.ps1 -Clean
    Cleans outputs for every work step that -All would run (when no narrower work switch is given).

  - .\CustomizedWindowsInstallation.ps1 -KB -Clean
  - .\CustomizedWindowsInstallation.ps1 -Clean -KB
    Cleans the KBs tree (and HtmlAgilityPack.dll). Placement of -Clean is not important.

  - .\CustomizedWindowsInstallation.ps1 -KB
    Downloads catalog updates into the working folder only.

  - .\CustomizedWindowsInstallation.ps1 -Indices "* N*"
    Selects all N editions (quote required due to space).

  - .\CustomizedWindowsInstallation.ps1 -ShowIndices
    Shows install.wim indices and exits.

  - .\CustomizedWindowsInstallation.ps1 -Most
    Full pipeline except ISO creation.

  - .\CustomizedWindowsInstallation.ps1 -CreateISO
    Builds <DestISO> from an existing DestISO\Content tree (run -Prep first if needed).

Index selection:
  - Explicit selection can be made using:
      -Home, -Pro
      -Indices with numbers, ranges, labels, wildcard labels (* and ?), or regex labels (re:<pattern>).
  - If no selection is provided, all install.wim indices are selected.

---

## Naming and paths

### `$names` (logical keys → folder/file names)

```powershell
$names = [ordered]@{
    SrcIso                = 'SrcISO'
    DestIso               = 'DestISO'
    KBs                   = 'KBs'
    Wims                  = 'Wims'
    WinPEDriver           = '$WinPEDriver$'
    Registry              = 'Registry'
    Content               = 'Content'          # ISO root inside SrcISO/DestISO
    Sources               = 'sources'
    BootWim               = 'boot.wim'
    InstallEsd            = 'install.esd'
    InstallWim            = 'install.wim'
    WinreWim              = 'winre.wim'
    BootFileBIOS          = 'boot\etfsboot.com'
    BootFileUEFI          = 'efi\microsoft\boot\efisys.bin'
    SetupExe              = 'setup.exe'
    ExportDriversCmd      = 'ExportDrivers.cmd'
    InstallDriversCmd     = 'InstallDrivers.cmd'
    ExportRegsCmd         = 'ExportRegs.cmd'
    ExportRegsPs1         = 'ExportRegs.ps1'
    InstallRegsCmd        = 'InstallRegs.cmd'
    SetupConfigCleanIni   = 'SetupConfig-Clean.ini'
    SetupConfigUpgradeIni = 'SetupConfig-Upgrade.ini'
    CleanInstallCmd       = 'CleanInstall.cmd'
    UpgradeCmd            = 'Upgrade.cmd'
    UpdateNETCmd          = 'Update.NET.cmd'
    UpdateNETPs1          = 'Update.NET.ps1'
    MetadataJson          = 'wim-metadata.json'
    ManifestJson          = 'manifest.json'
    ExtractJson           = 'extract.json'
    Unknown               = 'unknown'
}
# Plus SSU, OSCU, NET, MISC and Wims subdirs: Indices, Mounts, Serviced, Final, Scratch, Logs
```

### Resolved layout under `<Folder>`

```
<Folder>\
  HtmlAgilityPack.dll          # KB workflow only
  SrcISO\
    extract.json                 # { ISOPath, Date } — extract checkpoint
    Content\                     # Full ISO tree (sources\install.wim, boot.wim, …)
  DestISO\
    hardlink.json                # Prep hardlink step date
    install.wim.json, boot.wim.json
    Content\                     # Staged ISO root (hardlinks + final WIMs + helpers)
      sources\install.wim, boot.wim
      $WinPEDriver$, Registry\, KBs\NET, KBs\MISC (if present)
      *.cmd, *.ini, Update.NET.*
  KBs\
    SSU\                         # Servicing reads here (often empty unless populated manually)
    OSCU\                        # Cumulative updates (catalog search → here)
    NET\
    MISC\
    manifest.json per folder with downloads
  Wims\
    Indices\
      wim-metadata.json
      <index>_install.wim, <index>_install.wim.json
      <index>_boot.wim, <index>_boot.wim.json
    Mounts\                      # mount_<index>_install.wim, etc.
    Serviced\                    # Serviced WIMs + <index>_<wim>.json
    Final\
      install.wim, boot.wim
      install.wim.json, boot.wim.json
    Scratch\
    Logs\
  <DestISO path>.iso             # Default: <ISO>_KBs.iso
```

**WinRE path inside mounted install image:** `Windows\System32\Recovery\winre.wim` (`$paths.WinreWimInWim`).

---

Checkpointing and resume:
  Resume uses JSON timestamps

  - SrcISO\extract.json (Invoke-ExtractISO)
    Skip when ISOPath matches the current -ISO.

  - Wims\Indices\<n>_install.wim.json (Invoke-Export)
    Skip when ExportDate is newer than the extract date.

  - Wims\Indices\wim-metadata.json (Invoke-Export)
    Skip when valid and ISOPath matches.

  - Wims\Serviced\<n>_install.wim.json (Invoke-ServiceWork)
    Skip when ServicedDate is newer than the extract date.

  - Wims\Final\install.wim.json (Invoke-FinalAssembly)
    Skip when all serviced indices are older than the final assembly date.

  - DestISO\hardlink.json, install.wim.json, boot.wim.json (Invoke-PrepDestISO)
    Skip when dates are newer than upstream steps.

  On failure, the finally block dismounts any attached ISO and discards leftover DISM mounts under Wims\Mounts\.

---

## Pipeline orchestration (main block)

Order when multiple switches are set:

1. `Invoke-ExtractISO`
2. `Invoke-Export`
3. `Invoke-KBWork` (+ HtmlAgilityPack bootstrap)
4. `Invoke-ServiceWork`
5. `Invoke-FinalAssembly`
6. `Invoke-PrepDestISO`
7. `Invoke-FilesWork`
8. `Invoke-CreateISOWork`

Metadata for `WinOS` / `Version` / `Arch` / indices: `wim-metadata.json` → `SrcISO\Content\sources\` WIMs → brief ISO mount.

---

## Workflows (implementation summary)

### Extract (`Invoke-ExtractISO`)

- **Clean:** `Clean-Folder` `SrcISO` root.
- **DryRun:** Print mount/copy/validate intentions.
- Mount `-ISO`, validate `boot\etfsboot.com`, `efi\microsoft\boot\efisys.bin`, `sources\boot.wim`, and `install.wim` or `install.esd`.
- `Stream-FolderCopy` ISO drive → `SrcISO\Content`.
- Write `extract.json`; skip if same `ISOPath`.

### Export (`Invoke-Export`)

- **Clean:** `Wims` root.
- Requires `extract.json`; builds/refreshes `wim-metadata.json` via `Get-WimMetadata`.
- `Resolve-IndexSelection` → export each selected install index and **all** boot indices with DISM `/Export-Image` `/Compress:None` to `Wims\Indices\<index>_<wim>`.
- Per-export JSON: `{ Index, Name, ExportDate }`.

### KB (`Invoke-KBWork`)

Catalog queries (HtmlAgilityPack + Invoke-WebRequest):
  - "Cumulative Updates for Windows <WinOS> Version <Version> for <Arch>-based Systems"
    FirstOnly: false. Target: KBs\OSCU.

  - ".NET Framework for Windows <WinOS> Version <Version> <Arch>"
    FirstOnly: true. Target: KBs\NET.

  - ".NET 8.0 <Arch> Client"
    FirstOnly: true. Target: KBs\NET.

  - "Update for Windows Security platform"
    FirstOnly: true. Target: KBs\MISC.

  Flow: Search-UpdateCatalogHtml → dedupe GUIDs → Get-UpdateDetails (drop superseded, no URLs)
    → sync stale files → Download-MUFile → manifest.json per folder.

  KBs\SSU is used by Service but is not populated by catalog search in current code.

### Service (`Invoke-ServiceWork`)

- **Clean:** `Wims\Logs`, `Mounts`, `Scratch`, `Serviced`.
- Packages: `*.msu`, `*.cab` under `KBs\SSU` (SSU) and `KBs\OSCU` (LCU).
- Per install index (`Service-Index` `-InstallType $true`):
  1. Copy `Indices\<n>_install.wim` → `Serviced\`, mount index 1.
  2. Mount embedded `winre.wim`, apply SSU packages, unmount/commit WinRE in-place.
  3. Apply LCU packages, unmount/commit install WIM.
- Per boot index: mount, apply SSU, unmount/commit.
- Internal `Add-Packages`: optional copy via `Stream-FileCopy`, DISM mount/add-package/unmount with per-op logs under `Wims\Logs`.

### Final (`Invoke-FinalAssembly`)

- **Clean:** `Wims\Final`.
- `Build-FinalImage`: for each image (install indices sorted by index, all boot indices), DISM `/Export-Image` from `Serviced\<n>_<wim>` into `Wims\Final\<wim>` with `/Compress:$DismCompression`, appending indices 2..N.
- Writes `Wims\Final\install.wim.json` / `boot.wim.json` with `{ Date }`.

### Prep (`Invoke-PrepDestISO`)

- **Clean:** `DestISO` root.
- Hardlink all files from `SrcISO\Content` → `DestISO\Content` except `boot.wim`, `install.wim`, `install.esd` (fallback to copy on hardlink failure).
- `Stream-FileCopy` final WIMs from `Wims\Final` → `DestISO\Content\sources\`.

### Files (`Invoke-FilesWork`)

- Ensures `DestISO\Content\$WinPEDriver$` and `Registry`.
- Copies from **repo root** (`$Folder`): driver/registry/setup scripts, `Update.NET.*`.
- Optionally copies `KBs\NET` and `KBs\MISC` into `DestISO\Content\KBs\` if they exist.
- Does **not** run export drivers/regs; only stages files.

### Create ISO (`Invoke-CreateISOWork`)

- Requires `DestISO\Content` with boot files and `setup.exe`.
- `oscdimg -o -u2 -udfver102 -l<label> -bootdata:2#p0,e,b<etfsboot>#pEF,e,b<efisys> <Content> <DestISO>`.
- Progress: parse `% complete` lines, emit every 10%.

---

## Function reference (for reimplementation)

Each entry lists **parameters**, **returns**, **side effects**, and **key logic**. All user-visible text uses **`Write-Host`** (not `Write-Output`), except `Run-Dism -Capture` which uses `Write-Output` for captured lines.

### `Resolve-FullPath` — `[string]$Path` → string

Resolve relative paths against `$PWD`; absolute paths via `[IO.Path]::GetFullPath`.

### `FolderRelName` — `[string]$Path` (mandatory) → string

If path is under script-scope `$Folder`, return substring after folder prefix; else return full path. Used in log messages.

### `Protect-Token` — `[string]$s` → string

Sanitize strings for filenames: non `[word.-]` → `_`, collapse underscores, trim. Empty → `$names.Unknown`.

### `Show-Usage`

No parameters. Print colored usage summary (includes `$GitHash`). Does not list every parameter (use `-Help` for full help).

### `Ensure-Folder` — `[string]$Path`

Create directory if missing (`New-Item -Force`).

### `Remove-Folder` — `[string]$Path`

`Remove-Item -Recurse -Force` if exists.

### `Clean-File` / `Clean-Folder` — `[string]$Path`

If script-scope `$Clean`: DryRun message or delete file/folder; paths shown via `FolderRelName`.

### `Stream-FileCopy`

Parameters: SourcePath, DestinationPath (file copy); CopiedBytes, TotalBytes (multi-file progress);
ShowSourceOnly, NoTitle, NoProgress, NoComplete (output control).

64 MB buffered copy; progress when file is at least $BufferSize or part of a multi-file total;
$MaxIORetries with rollback of partial destination on failure.

### `Stream-FolderCopy` — `SourcePath`, `DestinationPath`, same switches

Recurse source; sum file sizes; per-file `Stream-FileCopy`; on error remove destination folder and throw.

### `Read-JsonFile` — `[string]$Path` → object or `$null`

`Get-Content -Raw | ConvertFrom-Json`; warning on failure.

### `Write-JsonFile` — `[string]$Path`, `[object]$Data`

Ensure parent dir; `ConvertTo-Json -Depth 6` UTF8.

### `Run-Dism` — `[string[]]$ArgumentList`, `[switch]$Capture`, `[int]$Indent`

Start `$dismExe` with redirected stdout/stderr, OEM encoding. Without `-Capture`: parse `%` in DISM output, emit 10% bucket progress via `Write-Host`. With `-Capture`: accumulate lines, `Write-Output` array at end. Sets `$global:LASTEXITCODE`. Kill process in `finally` if still running.

### `Report-Missing` — `[string[]]$Required`, `[string[]]$AtLeastOne`

Return `$true` if any required path missing or none of `AtLeastOne` exist; print which.

### `Find-ADKTool` — `ToolName`, `ADKSubfolder`, `ExplicitPath`, `PreferADK`, `ForceSystem` → string or `$null`

Priority: explicit path → (ForceSystem ? system : ADK if PreferADK or ADK found) → system fallback. ADK roots: `Program Files (x86)\Windows Kits\10\...\Deployment Tools\{amd64|arm64|x86}\<subfolder>\<tool>`.

### `Get-WimMetadata` — `[string]$WimPath` → PSCustomObject

DISM `/Get-WimInfo` (all indices) + `/Index:1` for build/arch. Returns `{ Images, WinOS, Version, Arch, Build }`. Build→version table includes 26200→25H2, 26100→24H2, etc.

### `Resolve-IndexSelection` → array of `{ Index, Name }`

Uses script-scope `$InstallImages`, `$SelectHome`, `$SelectPro`, `$Indices`. Union of Home/Pro/token matches; `Sort-Object Index -Unique`. No explicit flags → all images.

### `Invoke-ExtractISO`

See [Extract](#extract-invoke-extractiso). Uses `$ISO`, `$paths`, `$Clean`, `$DryRun`.

### `Invoke-Export`

See [Export](#export-invoke-export). Nested `Export-WimImage` writes per-index JSON checkpoints.

### `Invoke-CatalogRequest` — `[string]$Uri` → HtmlAgilityPack.HtmlDocument or `$null`

TLS 1.2, browser User-Agent, `Invoke-WebRequest` + load `RawContent` into HTML doc.

### `Get-CatalogSupersededBy` — `[string]$Html` → string[]

Regex extract links from `id="supersededbyInfo"` section (helper; primary path uses HtmlAgilityPack in `Get-UpdateDetails`).

### `Search-UpdateCatalogHtml` — `Query`, `FirstOnly`, `TargetFolder` → `{ Guid, TargetFolder }[]`

Search catalog URL; regex `goToDetails("GUID")` on HTML.

### `Get-UpdateLinks` — `[string]$Guid` → link objects

POST `DownloadDialog.aspx`; regex `downloadInformation[n].files[m].url`; dedupe; sort by KB number descending.

### `Load-Manifest` / `Write-Manifest` — `Folder`, (`Entries`)

Read/write `manifest.json` array in a KB folder.

### `Download-MUFile` — `Update` (with `DownloadUrls`, `Guid`, `Title`), `TargetFolder`

Skip existing files; else `HttpWebRequest` streaming download with 10% progress; retries `$MaxIORetries`.

### `Get-UpdateDetails` — `Count`, `Guid`, `TargetFolder` → detail object or automation null

Fetch details page; skip if superseded or no URLs; attach sorted download URLs.

### `Build-ManifestEntry` — `Details`, `DownloadInfo`

Single manifest row: Guid, Title, DownloadUrl, FileName, Timestamp.

### `Invoke-KBWork`

See [KB](#kb-invoke-kbwork).

### `Test-MissingIndices`

Throw if `Wims\Indices` missing or `$InstallIndices` / `$BootImages` empty (message: run `-Export -Clean` first).

### `Invoke-ServiceWork`

See [Service](#service-invoke-servicework). Nested `Add-Packages`, `Service-Index`. Uses `$extractDate` from `extract.json` for staleness (script-scope; set during export path).

### `Invoke-FinalAssembly`

See [Final](#final-invoke-finalassembly). Nested `Find-AnyOldServiced`, `Build-FinalImage`.

### `Invoke-PrepDestISO`

See [Prep](#prep-invoke-prepdestiso). Nested `Fetch-Date` reads JSON `Date` field.

### `Invoke-FilesWork`

See [Files](#files-invoke-fileswork). Job list: folders, file copies from `$Folder`, optional KB subfolders.

### `Invoke-CreateISOWork`

See [Create ISO](#create-iso-invoke-createisowork). Nested `Handle-Line` for oscdimg stdout/stderr.

### Main-only helpers

- **`Apply-WimMetadata` / `Try-WimMetadata`:** Populate `$WinOS`, `$Version`, `$Arch`, `$InstallImages`, `$BootImages` from JSON or live WIMs.
- **`$workSwitches` / `-All` / `-Most`:** Expand default pipeline (see Parameters above).

---

Dynamic Update concepts (servicing design):
  Microsoft media dynamic updates (reference [1]):

  - Servicing stack / SSU (KBs\SSU)
    Applied to install.wim (including WinRE inside the mount) and boot.wim.

  - Latest cumulative / LCU (KBs\OSCU)
    Applied to install.wim after the SSU and WinRE servicing steps.

  - .NET and misc (KBs\NET, KBs\MISC)
    Staged on the ISO via -Files; not DISM-applied during Service.

  The $WinPEDriver$ folder at the ISO root is scanned by Setup during installation ([5]).
  InstallDrivers.cmd uses pnputil after setup completes.

  SetupConfig-Upgrade.ini and SetupConfig-Clean.ini are used with Upgrade.cmd and CleanInstall.cmd
  (/ConfigFile passed to setup.exe).

---

## Output, verbosity, and progress

- **Normal:** Steady `Write-Host` status; DISM and oscdimg progress at 10% buckets; copy and download progress via Stream-FileCopy and Download-MUFile at 10% increments; no default PowerShell blue progress bars (`$ProgressPreference = 'SilentlyContinue'`).
- **Verbose:** Parameter echoes for important calls; DISM lines to verbose stream.
- **Debug:** All function entry context; DISM raw lines; iterator details.
- **Pipeline:** Do not rely on `Write-Output` for user messages; avoid assigning directly from `foreach` when the loop body uses `Write-Host` (use list accumulation pattern noted in script header).

Progress format for copies/downloads: right-aligned byte counts using width derived from total size; percent width `$PercentWidth` (3 → fits 100%).

---

Clean behavior (by step):
  - -Extract -Clean
    Removes SrcISO\.

  - -Export -Clean
    Removes Wims\.

  - -KB -Clean
    Removes KBs\ and HtmlAgilityPack.dll (via Clean-File).

  - -Service -Clean
    Removes Wims\Logs, Wims\Mounts, Wims\Scratch, and Wims\Serviced.

  - -Final -Clean
    Removes Wims\Final.

  - -Prep -Clean
    Removes DestISO\.

  - -Files -Clean
    Removes DestISO helper paths via Clean-File / Clean-Folder in the job loop.

  - -CreateISO -Clean
    Removes the output ISO file at DestISO.

---

## DryRun behavior

Each `Invoke-*` prints `[DryRun] Would …` for its actions and returns without mutation. KB path may print HtmlAgilityPack download intent only.

---

## Validation checklist (operator)

1. Run `-ShowIndices` or inspect `wim-metadata.json` after export.
2. Confirm `KBs\OSCU` has packages before `-Service`.
3. Test `Wims\Final\boot.wim` in VM if boot behavior matters.
4. Verify WinRE servicing path under mounted install image when SSU packages exist.
5. `-Prep` then inspect `DestISO\Content\sources\install.wim` indices.
6. `-CreateISO` and boot test UEFI/BIOS as needed.

---

## References

1. https://learn.microsoft.com/en-us/windows/deployment/update/media-dynamic-update  
2. https://github.com/Marco-online/MSCatalogLTS  
3. https://www.deploymentresearch.com/removing-applications-from-your-windows-11-image-before-and-during-deployment/  
4. https://thedotsource.com/2021/03/16/building-iso-files-with-powershell-7/  
5. https://learn.microsoft.com/en-us/windows-hardware/customize/desktop/unattend/microsoft-windows-pnpcustomizationswinpe-driverpaths  
6. https://community.spiceworks.com/t/autounattend-xml-driver-path-issue-for-windows-11-24h2-and-25h2/1244985  
7. https://github.com/wikijm/PowerShell-AdminScripts/blob/master/Miscellaneous/New-IsoFile.ps1  
8. https://www.winhelponline.com/blog/servicing-stack-diagnosis-dism-sfc/

---

## Known gaps vs. ideal design (documented for agents)

1. **`-Drivers` / `-Reg`:** Referenced in `-All` block but not in `param()`; no in-script export workflow—only file copy via `-Files`.
2. **`KBs\SSU`:** Not filled by `-KB`; Service still applies any files placed there manually.
3. **`Try-WimMetadata`:** Uses `.InstallImages` / `.BootImages` on `Get-WimMetadata` result, but that function returns `.Images` only—live ISO metadata path may not populate indices until export.
