<#

.SYNOPSIS
Installs only the best-matching drivers for devices that are missing drivers.

.DESCRIPTION
Finds present devices with CM problem code 28 (drivers not installed), then uses
Windows SetupAPI to rank compatible INF packages under -Folder without staging
them into the Driver Store. Only the single best INF selected for each device
is added/installed via pnputil.

.PARAMETER Folder
Root folder containing driver packages (searched recursively). Relative paths
are resolved from the script directory.

#>

[CmdletBinding()]
param(
    [string]$Folder = ".\Drivers"
)

$ErrorActionPreference = 'Stop'

# Resolve relative paths from the script location
if (-not ([System.IO.Path]::IsPathRooted($Folder))) {
    $Folder = Join-Path $PSScriptRoot $Folder
}

if (-not (Test-Path -LiteralPath $Folder -PathType Container)) {
    Write-Error "Driver folder not found: $Folder"
    Exit 1
}

$Folder = (Resolve-Path -LiteralPath $Folder).ProviderPath
Write-Host "Resolved driver folder: $Folder" -ForegroundColor Green

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)
if (-not $isAdmin) {
    Write-Error "This script must be run as Administrator."
    Exit 1
}

$setupApiCode = @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace SetupApiDriverHelperV3
{
    [StructLayout(LayoutKind.Sequential)]
    public struct SP_DEVINFO_DATA
    {
        public int cbSize;
        public Guid ClassGuid;
        public int DevInst;
        public IntPtr Reserved;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct SP_DEVINSTALL_PARAMS
    {
        public int cbSize;
        public int Flags;
        public int FlagsEx;
        public IntPtr hwndParent;
        public IntPtr InstallMsgHandler;
        public IntPtr InstallMsgHandlerContext;
        public IntPtr FileQueue;
        public IntPtr ClassInstallReserved;
        public int Reserved;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string DriverPath;
    }

    // Pack=4 matches setupapi.h pshpack1/pack usage for this struct on both x86/x64.
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode, Pack = 4)]
    public struct SP_DRVINFO_DATA
    {
        public int cbSize;
        public int DriverType;
        public IntPtr Reserved;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string Description;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string MfgName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string ProviderName;
        public System.Runtime.InteropServices.ComTypes.FILETIME DriverDate;
        public ulong DriverVersion;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct SP_DRVINSTALL_PARAMS
    {
        public int cbSize;
        public uint Rank;
        public int Flags;
        public IntPtr PrivateData;
        public int Reserved;
    }

    // Default packing (8 on x64) — Pack=4 misaligns FILETIME and causes error 1784.
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct SP_DRVINFO_DETAIL_DATA
    {
        public int cbSize;
        public System.Runtime.InteropServices.ComTypes.FILETIME InfDate;
        public int CompatIDsOffset;
        public int CompatIDsLength;
        public IntPtr Reserved;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string SectionName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string InfFileName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string DrvDescription;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 1)]
        public string HardwareID;
    }

    public class DriverMatch
    {
        public string InstanceId;
        public string InfPath;
        public string Description;
        public string Provider;
        public string Manufacturer;
        public uint Rank;
        public string Error;
    }

    public static class Native
    {
        private const int DIGCF_PRESENT = 0x00000002;
        private const int DIGCF_ALLCLASSES = 0x00000004;
        private const int SPDIT_COMPATDRIVER = 2;
        private const int DIF_SELECTBESTCOMPATDRV = 0x00000017;
        private const int DI_FLAGSEX_ALLOWEXCLUDEDDRVS = 0x00000800;
        private const int DI_FLAGSEX_RECURSIVESEARCH = unchecked((int)0x40000000);
        private const int ERROR_INSUFFICIENT_BUFFER = 122;
        private const int ERROR_DI_BAD_PATH = unchecked((int)0xE000020B);
        private const int ERROR_NO_COMPAT_DRIVERS = unchecked((int)0xE0000203);
        // SetupAPI: "There are no compatible drivers for this device."
        private const int ERROR_NO_COMPATIBLE_DRIVERS = unchecked((int)0xE0000228);
        private static readonly IntPtr INVALID_HANDLE_VALUE = new IntPtr(-1);

        [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr SetupDiGetClassDevs(
            IntPtr ClassGuid,
            string Enumerator,
            IntPtr hwndParent,
            int Flags);

        [DllImport("setupapi.dll", SetLastError = true)]
        private static extern bool SetupDiDestroyDeviceInfoList(IntPtr DeviceInfoSet);

        [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SetupDiOpenDeviceInfo(
            IntPtr DeviceInfoSet,
            string DeviceInstanceId,
            IntPtr hwndParent,
            int OpenFlags,
            ref SP_DEVINFO_DATA DeviceInfoData);

        [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SetupDiGetDeviceInstallParams(
            IntPtr DeviceInfoSet,
            ref SP_DEVINFO_DATA DeviceInfoData,
            ref SP_DEVINSTALL_PARAMS DeviceInstallParams);

        [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SetupDiSetDeviceInstallParams(
            IntPtr DeviceInfoSet,
            ref SP_DEVINFO_DATA DeviceInfoData,
            ref SP_DEVINSTALL_PARAMS DeviceInstallParams);

        [DllImport("setupapi.dll", SetLastError = true)]
        private static extern bool SetupDiBuildDriverInfoList(
            IntPtr DeviceInfoSet,
            ref SP_DEVINFO_DATA DeviceInfoData,
            int DriverType);

        [DllImport("setupapi.dll", SetLastError = true)]
        private static extern bool SetupDiDestroyDriverInfoList(
            IntPtr DeviceInfoSet,
            ref SP_DEVINFO_DATA DeviceInfoData,
            int DriverType);

        [DllImport("setupapi.dll", SetLastError = true)]
        private static extern bool SetupDiCallClassInstaller(
            int InstallFunction,
            IntPtr DeviceInfoSet,
            ref SP_DEVINFO_DATA DeviceInfoData);

        [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SetupDiGetSelectedDriver(
            IntPtr DeviceInfoSet,
            ref SP_DEVINFO_DATA DeviceInfoData,
            ref SP_DRVINFO_DATA DriverInfoData);

        [DllImport("setupapi.dll", SetLastError = true)]
        private static extern bool SetupDiGetDriverInstallParams(
            IntPtr DeviceInfoSet,
            ref SP_DEVINFO_DATA DeviceInfoData,
            ref SP_DRVINFO_DATA DriverInfoData,
            ref SP_DRVINSTALL_PARAMS DriverInstallParams);

        [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SetupDiGetDriverInfoDetail(
            IntPtr DeviceInfoSet,
            ref SP_DEVINFO_DATA DeviceInfoData,
            ref SP_DRVINFO_DATA DriverInfoData,
            IntPtr DriverInfoDetailData,
            int DriverInfoDetailDataSize,
            out int RequiredSize);

        private static string FormatWin32Error(int error)
        {
            return error.ToString() + " (0x" + ((uint)error).ToString("X8") + ")";
        }

        private static string TryGetInfPath(IntPtr detailPtr)
        {
            IntPtr nameOffset = Marshal.OffsetOf(typeof(SP_DRVINFO_DETAIL_DATA), "InfFileName");
            IntPtr namePtr = new IntPtr(detailPtr.ToInt64() + nameOffset.ToInt64());
            return Marshal.PtrToStringUni(namePtr);
        }

        public static DriverMatch[] SelectBestDrivers(string[] instanceIds, string driverFolder)
        {
            if (instanceIds == null)
            {
                throw new ArgumentNullException("instanceIds");
            }
            if (string.IsNullOrWhiteSpace(driverFolder))
            {
                throw new ArgumentException("driverFolder is required.");
            }

            List<DriverMatch> results = new List<DriverMatch>();
            IntPtr deviceInfoSet = SetupDiGetClassDevs(
                IntPtr.Zero,
                null,
                IntPtr.Zero,
                DIGCF_PRESENT | DIGCF_ALLCLASSES);

            if (deviceInfoSet == IntPtr.Zero || deviceInfoSet == INVALID_HANDLE_VALUE)
            {
                throw new InvalidOperationException(
                    "SetupDiGetClassDevs failed: " + FormatWin32Error(Marshal.GetLastWin32Error()));
            }

            try
            {
                for (int i = 0; i < instanceIds.Length; i++)
                {
                    string instanceId = instanceIds[i];
                    DriverMatch match = new DriverMatch();
                    match.InstanceId = instanceId;
                    results.Add(match);

                    if (string.IsNullOrWhiteSpace(instanceId))
                    {
                        match.Error = "Empty InstanceId.";
                        continue;
                    }

                    SP_DEVINFO_DATA deviceInfoData = new SP_DEVINFO_DATA();
                    deviceInfoData.cbSize = Marshal.SizeOf(typeof(SP_DEVINFO_DATA));

                    if (!SetupDiOpenDeviceInfo(deviceInfoSet, instanceId, IntPtr.Zero, 0, ref deviceInfoData))
                    {
                        match.Error = "SetupDiOpenDeviceInfo failed: " + FormatWin32Error(Marshal.GetLastWin32Error());
                        continue;
                    }

                    bool builtList = false;
                    try
                    {
                        SP_DEVINSTALL_PARAMS installParams = new SP_DEVINSTALL_PARAMS();
                        installParams.cbSize = Marshal.SizeOf(typeof(SP_DEVINSTALL_PARAMS));

                        if (!SetupDiGetDeviceInstallParams(deviceInfoSet, ref deviceInfoData, ref installParams))
                        {
                            match.Error = "SetupDiGetDeviceInstallParams failed: " + FormatWin32Error(Marshal.GetLastWin32Error());
                            continue;
                        }

                        installParams.DriverPath = driverFolder;
                        installParams.FlagsEx = installParams.FlagsEx |
                            DI_FLAGSEX_ALLOWEXCLUDEDDRVS |
                            DI_FLAGSEX_RECURSIVESEARCH;

                        if (!SetupDiSetDeviceInstallParams(deviceInfoSet, ref deviceInfoData, ref installParams))
                        {
                            match.Error = "SetupDiSetDeviceInstallParams failed: " + FormatWin32Error(Marshal.GetLastWin32Error());
                            continue;
                        }

                        if (!SetupDiBuildDriverInfoList(deviceInfoSet, ref deviceInfoData, SPDIT_COMPATDRIVER))
                        {
                            match.Error = "SetupDiBuildDriverInfoList failed: " + FormatWin32Error(Marshal.GetLastWin32Error());
                            continue;
                        }
                        builtList = true;

                        if (!SetupDiCallClassInstaller(DIF_SELECTBESTCOMPATDRV, deviceInfoSet, ref deviceInfoData))
                        {
                            int err = Marshal.GetLastWin32Error();
                            if (err == ERROR_DI_BAD_PATH ||
                                err == ERROR_NO_COMPAT_DRIVERS ||
                                err == ERROR_NO_COMPATIBLE_DRIVERS)
                            {
                                match.Error = "No compatible function driver found in the folder (SetupAPI " +
                                    FormatWin32Error(err) + ").";
                            }
                            else
                            {
                                match.Error = "DIF_SELECTBESTCOMPATDRV failed: " + FormatWin32Error(err);
                            }
                            continue;
                        }

                        SP_DRVINFO_DATA driverInfoData = new SP_DRVINFO_DATA();
                        driverInfoData.cbSize = Marshal.SizeOf(typeof(SP_DRVINFO_DATA));

                        if (!SetupDiGetSelectedDriver(deviceInfoSet, ref deviceInfoData, ref driverInfoData))
                        {
                            match.Error = "SetupDiGetSelectedDriver failed: " + FormatWin32Error(Marshal.GetLastWin32Error());
                            continue;
                        }

                        match.Description = driverInfoData.Description;
                        match.Provider = driverInfoData.ProviderName;
                        match.Manufacturer = driverInfoData.MfgName;

                        SP_DRVINSTALL_PARAMS driverInstallParams = new SP_DRVINSTALL_PARAMS();
                        driverInstallParams.cbSize = Marshal.SizeOf(typeof(SP_DRVINSTALL_PARAMS));
                        if (SetupDiGetDriverInstallParams(
                            deviceInfoSet,
                            ref deviceInfoData,
                            ref driverInfoData,
                            ref driverInstallParams))
                        {
                            match.Rank = driverInstallParams.Rank;
                        }

                        int structSize = Marshal.SizeOf(typeof(SP_DRVINFO_DETAIL_DATA));
                        int requiredSize = 0;

                        // Pass 1: ask Windows for the full buffer size (includes Hardware ID list).
                        SetupDiGetDriverInfoDetail(
                            deviceInfoSet,
                            ref deviceInfoData,
                            ref driverInfoData,
                            IntPtr.Zero,
                            0,
                            out requiredSize);

                        int sizeErr = Marshal.GetLastWin32Error();
                        if (requiredSize < structSize)
                        {
                            // Some hosts return 1784 instead of 122 on the probe call; fall back.
                            requiredSize = structSize + 4096;
                        }
                        else if (sizeErr != ERROR_INSUFFICIENT_BUFFER && sizeErr != 0 && requiredSize == 0)
                        {
                            match.Error = "SetupDiGetDriverInfoDetail size probe failed: " + FormatWin32Error(sizeErr);
                            continue;
                        }

                        IntPtr detailPtr = Marshal.AllocHGlobal(requiredSize);
                        try
                        {
                            for (int b = 0; b < requiredSize; b++)
                            {
                                Marshal.WriteByte(detailPtr, b, 0);
                            }

                            // cbSize must be sizeof(SP_DRVINFO_DETAIL_DATA), NOT the allocated buffer size.
                            Marshal.WriteInt32(detailPtr, 0, structSize);

                            int requiredSize2 = 0;
                            if (!SetupDiGetDriverInfoDetail(
                                deviceInfoSet,
                                ref deviceInfoData,
                                ref driverInfoData,
                                detailPtr,
                                requiredSize,
                                out requiredSize2))
                            {
                                match.Error = "SetupDiGetDriverInfoDetail failed: " +
                                    FormatWin32Error(Marshal.GetLastWin32Error()) +
                                    " (cbSize=" + structSize + ", buf=" + requiredSize + ")";
                                continue;
                            }

                            match.InfPath = TryGetInfPath(detailPtr);

                            if (string.IsNullOrWhiteSpace(match.InfPath))
                            {
                                match.Error = "Selected driver did not return an INF path.";
                            }
                        }
                        finally
                        {
                            Marshal.FreeHGlobal(detailPtr);
                        }
                    }
                    finally
                    {
                        if (builtList)
                        {
                            SetupDiDestroyDriverInfoList(deviceInfoSet, ref deviceInfoData, SPDIT_COMPATDRIVER);
                        }
                    }
                }
            }
            finally
            {
                SetupDiDestroyDeviceInfoList(deviceInfoSet);
            }

            return results.ToArray();
        }
    }
}
"@

if (-not ('SetupApiDriverHelperV3.Native' -as [type])) {
    Add-Type -TypeDefinition $setupApiCode -Language CSharp -ErrorAction Stop
}

Write-Host "Searching Device Manager for devices missing drivers (problem 28)..." -ForegroundColor Cyan

$MissingDevices = @(Get-PnpDevice -PresentOnly -ErrorAction Stop | Where-Object {
    $_.Problem -eq 28
})

if ($MissingDevices.Count -eq 0) {
    Write-Host "All present devices currently have drivers installed. Nothing to do!" -ForegroundColor Green
    Exit 0
}

Write-Host "Found $($MissingDevices.Count) device(s) missing drivers." -ForegroundColor Cyan
foreach ($device in $MissingDevices) {
    Write-Host ("  - {0}  [{1}]" -f $device.FriendlyName, $device.InstanceId) -ForegroundColor DarkCyan
}

Write-Host "Asking SetupAPI to select the best INF from '$Folder' for each device..." -ForegroundColor Cyan

$instanceIds = @($MissingDevices | ForEach-Object { $_.InstanceId })
$selections = [SetupApiDriverHelperV3.Native]::SelectBestDrivers($instanceIds, $Folder)

$toInstall = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
$selectedCount = 0

foreach ($selection in $selections) {
    if ($selection.Error) {
        Write-Warning ("No driver selected for {0}: {1}" -f $selection.InstanceId, $selection.Error)
        continue
    }

    $selectedCount++
    Write-Host ("Best match for {0}:" -f $selection.InstanceId) -ForegroundColor Magenta
    Write-Host ("  Description : {0}" -f $selection.Description)
    Write-Host ("  Provider    : {0}" -f $selection.Provider)
    Write-Host ("  Manufacturer: {0}" -f $selection.Manufacturer)
    Write-Host ("  Rank        : 0x{0:X8}" -f $selection.Rank)
    Write-Host ("  INF         : {0}" -f $selection.InfPath)

    if (-not (Test-Path -LiteralPath $selection.InfPath -PathType Leaf)) {
        Write-Warning ("Selected INF does not exist: {0}" -f $selection.InfPath)
        continue
    }

    [void]$toInstall.Add($selection.InfPath)
}

if ($toInstall.Count -eq 0) {
    Write-Warning "No matching drivers were found in the provided folder."
    Exit 0
}

Write-Host ("Installing {0} unique driver package(s) for {1} device(s)..." -f $toInstall.Count, $selectedCount) -ForegroundColor Magenta
Write-Host "------------------------------------------------------" -ForegroundColor Gray

foreach ($infPath in ($toInstall | Sort-Object)) {
    Write-Host ">>> Installing: $infPath" -ForegroundColor Yellow
    & pnputil.exe /add-driver $infPath /install | Out-Host
    Write-Host "------------------------------------------------------" -ForegroundColor Gray
}

Write-Host "Rescanning devices..." -ForegroundColor Cyan
& pnputil.exe /scan-devices | Out-Host

Write-Host "Verifying remaining missing drivers..." -ForegroundColor Cyan
$stillMissing = @(Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue | Where-Object {
    $_.Problem -eq 28
})

if ($stillMissing.Count -eq 0) {
    Write-Host "All previously missing devices now have drivers installed." -ForegroundColor Green
}
else {
    Write-Host ("{0} device(s) still report missing drivers:" -f $stillMissing.Count) -ForegroundColor Yellow
    foreach ($device in $stillMissing) {
        Write-Host ("  - {0}  [{1}]" -f $device.FriendlyName, $device.InstanceId) -ForegroundColor Yellow
    }
}

Write-Host "Driver installation phase complete!" -ForegroundColor Green
