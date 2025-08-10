# COMPREHENSIVE PC HEALTH CHECK SCRIPT - FULL VERSION
# Run as Administrator for complete analysis

# Self-elevate if not admin
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Requesting Administrator privileges..." -ForegroundColor Yellow
    Start-Process PowerShell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Setup
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$reportDir = Join-Path $env:USERPROFILE "Desktop\PC_Health_Report_$timestamp"
New-Item -Path $reportDir -ItemType Directory -Force | Out-Null

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  COMPREHENSIVE PC HEALTH CHECK - FULL SCAN" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "Report folder: $reportDir" -ForegroundColor Yellow

# Helper function with detailed output
function Save-Data {
    param($Name, $ScriptBlock, $FileName)
    Write-Host "Collecting $Name..." -ForegroundColor Cyan -NoNewline
    try {
        $startTime = Get-Date
        $data = & $ScriptBlock
        $endTime = Get-Date
        $duration = ($endTime - $startTime).TotalSeconds
        
        if ($data) {
            $data | Format-List * | Out-File -FilePath (Join-Path $reportDir $FileName) -Encoding UTF8
            Write-Host " COMPLETE ($([math]::Round($duration,1))s)" -ForegroundColor Green
        } else {
            "No data available" | Out-File -FilePath (Join-Path $reportDir $FileName) -Encoding UTF8
            Write-Host " NO DATA" -ForegroundColor Yellow
        }
    } catch {
        "ERROR: $_" | Out-File -FilePath (Join-Path $reportDir $FileName) -Encoding UTF8
        Write-Host " FAILED: $_" -ForegroundColor Red
    }
}

# ===== BASIC SYSTEM INFORMATION =====
Write-Host "`n[1/25] BASIC SYSTEM INFORMATION" -ForegroundColor Magenta

Save-Data "System Information (systeminfo)" {
    systeminfo
} "01_SystemInfo_Command.txt"

Save-Data "Computer Information" {
    Get-ComputerInfo
} "02_ComputerInfo.txt"

Save-Data "Operating System Details" {
    Get-CimInstance Win32_OperatingSystem
} "03_OperatingSystem.txt"

Save-Data "Computer System" {
    Get-CimInstance Win32_ComputerSystem
} "04_ComputerSystem.txt"

# ===== DETAILED HARDWARE ANALYSIS =====
Write-Host "`n[2/25] DETAILED HARDWARE ANALYSIS" -ForegroundColor Magenta

Save-Data "CPU/Processor Information" {
    Get-CimInstance Win32_Processor
} "05_CPU_Processor.txt"

Save-Data "Physical Memory (RAM) Details" {
    Get-CimInstance Win32_PhysicalMemory
} "06_PhysicalMemory.txt"

Save-Data "Memory Array Information" {
    Get-CimInstance Win32_PhysicalMemoryArray
} "07_MemoryArray.txt"

Save-Data "Motherboard Information" {
    Get-CimInstance Win32_BaseBoard
} "08_Motherboard.txt"

Save-Data "BIOS Information" {
    Get-CimInstance Win32_BIOS
} "09_BIOS.txt"

Save-Data "System Enclosure/Chassis" {
    Get-CimInstance Win32_SystemEnclosure
} "10_SystemEnclosure.txt"

Save-Data "Graphics/Video Controllers" {
    Get-CimInstance Win32_VideoController
} "11_VideoController.txt"

Save-Data "Sound Devices" {
    Get-CimInstance Win32_SoundDevice
} "12_SoundDevices.txt"

# ===== NETWORK HARDWARE =====
Write-Host "`n[3/25] NETWORK HARDWARE" -ForegroundColor Magenta

Save-Data "Network Adapters" {
    Get-CimInstance Win32_NetworkAdapter
} "13_NetworkAdapters.txt"

Save-Data "Network Adapter Configuration" {
    Get-CimInstance Win32_NetworkAdapterConfiguration
} "14_NetworkConfig.txt"

Save-Data "Active Network Adapters" {
    Get-NetAdapter -ErrorAction SilentlyContinue
} "15_ActiveNetworkAdapters.txt"

# ===== COMPREHENSIVE STORAGE ANALYSIS =====
Write-Host "`n[4/25] COMPREHENSIVE STORAGE ANALYSIS" -ForegroundColor Magenta

Save-Data "Physical Disks (Windows Storage)" {
    Get-PhysicalDisk -ErrorAction SilentlyContinue
} "16_PhysicalDisks.txt"

Save-Data "Storage Pools" {
    Get-StoragePool -ErrorAction SilentlyContinue
} "17_StoragePools.txt"

Save-Data "Virtual Disks" {
    Get-VirtualDisk -ErrorAction SilentlyContinue
} "18_VirtualDisks.txt"

Save-Data "Disk Drives (WMI)" {
    Get-CimInstance Win32_DiskDrive
} "19_DiskDrives_WMI.txt"

Save-Data "Logical Disks" {
    Get-CimInstance Win32_LogicalDisk
} "20_LogicalDisks.txt"

Save-Data "Disk Partitions" {
    Get-CimInstance Win32_DiskPartition
} "21_DiskPartitions.txt"

Save-Data "Volume Information" {
    Get-CimInstance Win32_Volume
} "22_Volumes.txt"

# ===== SMART DATA AND DRIVE HEALTH =====
Write-Host "`n[5/25] SMART DATA AND DRIVE HEALTH" -ForegroundColor Magenta

Save-Data "SMART Failure Predict Status" {
    Get-CimInstance -Namespace root\wmi -ClassName MSStorageDriver_FailurePredictStatus -ErrorAction SilentlyContinue
} "23_SMART_FailurePredict.txt"

Save-Data "SMART Failure Predict Data" {
    Get-CimInstance -Namespace root\wmi -ClassName MSStorageDriver_FailurePredictData -ErrorAction SilentlyContinue
} "24_SMART_Data.txt"

Save-Data "SMART Failure Predict Thresholds" {
    Get-CimInstance -Namespace root\wmi -ClassName MSStorageDriver_FailurePredictThresholds -ErrorAction SilentlyContinue
} "25_SMART_Thresholds.txt"

Save-Data "Disk Drive Performance" {
    Get-CimInstance Win32_PerfRawData_PerfDisk_PhysicalDisk -ErrorAction SilentlyContinue
} "26_DiskPerformance.txt"

# ===== USB AND EXTERNAL DEVICES =====
Write-Host "`n[6/25] USB AND EXTERNAL DEVICES" -ForegroundColor Magenta

Save-Data "USB Controllers" {
    Get-CimInstance Win32_USBController
} "27_USBControllers.txt"

Save-Data "USB Hub Information" {
    Get-CimInstance Win32_USBHub
} "28_USBHubs.txt"

Save-Data "PnP Devices" {
    Get-CimInstance Win32_PnPEntity
} "29_PnPDevices.txt"

Save-Data "System Devices" {
    Get-CimInstance Win32_SystemDevices -ErrorAction SilentlyContinue
} "30_SystemDevices.txt"

# ===== PERFORMANCE COUNTERS =====
Write-Host "`n[7/25] PERFORMANCE ANALYSIS" -ForegroundColor Magenta

Save-Data "CPU Performance Counters" {
    Write-Output "=== Current CPU Usage (5 samples) ==="
    Get-Counter '\Processor(_Total)\% Processor Time' -SampleInterval 1 -MaxSamples 5 -ErrorAction SilentlyContinue
    
    Write-Output "`n=== CPU Queue Length ==="
    Get-Counter '\System\Processor Queue Length' -MaxSamples 3 -ErrorAction SilentlyContinue
    
    Write-Output "`n=== Interrupt Time ==="
    Get-Counter '\Processor(_Total)\% Interrupt Time' -MaxSamples 3 -ErrorAction SilentlyContinue
} "31_CPU_Performance.txt"

Save-Data "Memory Performance Counters" {
    Write-Output "=== Available Memory ==="
    Get-Counter '\Memory\Available MBytes' -MaxSamples 3 -ErrorAction SilentlyContinue
    
    Write-Output "`n=== Page Faults ==="
    Get-Counter '\Memory\Page Faults/sec' -MaxSamples 3 -ErrorAction SilentlyContinue
    
    Write-Output "`n=== Pages per Second ==="
    Get-Counter '\Memory\Pages/sec' -MaxSamples 3 -ErrorAction SilentlyContinue
    
    Write-Output "`n=== Pool Usage ==="
    Get-Counter '\Memory\Pool Paged Bytes', '\Memory\Pool Nonpaged Bytes' -MaxSamples 1 -ErrorAction SilentlyContinue
} "32_Memory_Performance.txt"

Save-Data "Disk Performance Counters" {
    Write-Output "=== Disk Queue Length ==="
    Get-Counter '\PhysicalDisk(_Total)\Current Disk Queue Length' -MaxSamples 3 -ErrorAction SilentlyContinue
    
    Write-Output "`n=== Disk Transfers ==="
    Get-Counter '\PhysicalDisk(_Total)\Disk Transfers/sec' -MaxSamples 3 -ErrorAction SilentlyContinue
    
    Write-Output "`n=== Disk Read/Write ==="
    Get-Counter '\PhysicalDisk(_Total)\Disk Reads/sec', '\PhysicalDisk(_Total)\Disk Writes/sec' -MaxSamples 3 -ErrorAction SilentlyContinue
} "33_Disk_Performance.txt"

# ===== POWER AND THERMAL =====
Write-Host "`n[8/25] POWER AND THERMAL" -ForegroundColor Magenta

Save-Data "Power Configuration" {
    powercfg /query
} "34_PowerConfig.txt"

Save-Data "Battery Information" {
    Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
} "35_Battery.txt"

Save-Data "Portable Battery" {
    Get-CimInstance Win32_PortableBattery -ErrorAction SilentlyContinue
} "36_PortableBattery.txt"

Save-Data "Temperature Zones" {
    Get-CimInstance -Namespace root\wmi -ClassName MSAcpi_ThermalZoneTemperature -ErrorAction SilentlyContinue
} "37_ThermalZones.txt"

Save-Data "Power Supply" {
    Get-CimInstance Win32_PowerSupply -ErrorAction SilentlyContinue
} "38_PowerSupply.txt"

# ===== COMPREHENSIVE EVENT LOGS =====
Write-Host "`n[9/25] COMPREHENSIVE EVENT LOGS" -ForegroundColor Magenta

Save-Data "System Errors (Last 7 Days)" {
    $startDate = (Get-Date).AddDays(-7)
    Get-WinEvent -FilterHashtable @{LogName='System'; Level=1,2; StartTime=$startDate} -MaxEvents 500 -ErrorAction SilentlyContinue
} "39_SystemErrors_7Days.txt"

Save-Data "System Warnings (Last 7 Days)" {
    $startDate = (Get-Date).AddDays(-7)
    Get-WinEvent -FilterHashtable @{LogName='System'; Level=3; StartTime=$startDate} -MaxEvents 200 -ErrorAction SilentlyContinue
} "40_SystemWarnings_7Days.txt"

Save-Data "Application Errors (Last 7 Days)" {
    $startDate = (Get-Date).AddDays(-7)
    Get-WinEvent -FilterHashtable @{LogName='Application'; Level=1,2; StartTime=$startDate} -MaxEvents 200 -ErrorAction SilentlyContinue
} "41_ApplicationErrors_7Days.txt"

Save-Data "Hardware Events" {
    Get-WinEvent -FilterHashtable @{LogName='System'; ProviderName='Microsoft-Windows-Kernel-Power'} -MaxEvents 100 -ErrorAction SilentlyContinue
} "42_HardwareEvents.txt"

Save-Data "WHEA Hardware Error Events" {
    Get-WinEvent -FilterHashtable @{LogName='System'; ProviderName='Microsoft-Windows-WHEA-Logger'} -MaxEvents 100 -ErrorAction SilentlyContinue
} "43_WHEA_Events.txt"

Save-Data "Disk Events" {
    Get-WinEvent -FilterHashtable @{LogName='System'; ProviderName='disk'} -MaxEvents 100 -ErrorAction SilentlyContinue
} "44_DiskEvents.txt"

# ===== SOFTWARE AND DRIVERS =====
Write-Host "`n[10/25] SOFTWARE AND DRIVERS" -ForegroundColor Magenta

Save-Data "Installed Programs (Registry)" {
    $regPaths = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    
    $allPrograms = @()
    foreach ($path in $regPaths) {
        $programs = Get-ItemProperty $path -ErrorAction SilentlyContinue | Where-Object DisplayName
        $allPrograms += $programs
    }
    $allPrograms | Sort-Object DisplayName
} "45_InstalledPrograms.txt"

Save-Data "System Drivers" {
    Get-CimInstance Win32_PnPSignedDriver
} "46_SystemDrivers.txt"

Save-Data "Running Services" {
    Get-Service
} "47_Services.txt"

Save-Data "Running Processes" {
    Get-Process | Sort-Object CPU -Descending
} "48_Processes.txt"

Save-Data "Startup Programs" {
    Get-CimInstance Win32_StartupCommand
} "49_StartupPrograms.txt"

Save-Data "Windows Features" {
    Get-WindowsOptionalFeature -Online -ErrorAction SilentlyContinue
} "50_WindowsFeatures.txt"

# ===== WINDOWS UPDATES =====
Write-Host "`n[11/25] WINDOWS UPDATES" -ForegroundColor Magenta

Save-Data "Installed Hotfixes" {
    Get-HotFix | Sort-Object InstalledOn -Descending
} "51_HotFixes.txt"

Save-Data "Windows Update Log" {
    Get-WindowsUpdateLog -ErrorAction SilentlyContinue
} "52_WindowsUpdateLog.txt"

# ===== SECURITY AND FIREWALL =====
Write-Host "`n[12/25] SECURITY AND FIREWALL" -ForegroundColor Magenta

Save-Data "Windows Defender Status" {
    Get-MpPreference -ErrorAction SilentlyContinue
} "53_DefenderStatus.txt"

Save-Data "Firewall Rules" {
    Get-NetFirewallRule -ErrorAction SilentlyContinue
} "54_FirewallRules.txt"

Save-Data "User Accounts" {
    Get-LocalUser -ErrorAction SilentlyContinue
} "55_UserAccounts.txt"

# ===== REGISTRY INFORMATION =====
Write-Host "`n[13/25] REGISTRY INFORMATION" -ForegroundColor Magenta

Save-Data "System Registry Info" {
    Write-Output "=== Windows Version ==="
    Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue
    
    Write-Output "`n=== Hardware Info ==="
    Get-ItemProperty "HKLM:\HARDWARE\DESCRIPTION\System\CentralProcessor\0" -ErrorAction SilentlyContinue
    
    Write-Output "`n=== System Policies ==="
    Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -ErrorAction SilentlyContinue
} "56_RegistryInfo.txt"

# ===== SCHEDULED TASKS =====
Write-Host "`n[14/25] SCHEDULED TASKS" -ForegroundColor Magenta

Save-Data "Scheduled Tasks" {
    Get-ScheduledTask -ErrorAction SilentlyContinue
} "57_ScheduledTasks.txt"

# ===== ENVIRONMENT INFORMATION =====
Write-Host "`n[15/25] ENVIRONMENT INFORMATION" -ForegroundColor Magenta

Save-Data "Environment Variables" {
    Get-ChildItem Env:
} "58_EnvironmentVariables.txt"

Save-Data "System PATH" {
    $env:PATH -split ';'
} "59_SystemPATH.txt"

# ===== DETAILED CALCULATIONS =====
Write-Host "`n[16/25] DETAILED SYSTEM CALCULATIONS" -ForegroundColor Magenta

Save-Data "Memory Analysis" {
    $os = Get-CimInstance Win32_OperatingSystem
    $cs = Get-CimInstance Win32_ComputerSystem
    
    Write-Output "=== MEMORY ANALYSIS ==="
    Write-Output "Total Physical Memory: $([math]::Round($cs.TotalPhysicalMemory / 1GB, 2)) GB"
    Write-Output "Available Physical Memory: $([math]::Round($os.FreePhysicalMemory / 1MB, 2)) GB"
    Write-Output "Total Virtual Memory: $([math]::Round($os.TotalVirtualMemorySize / 1MB, 2)) GB"
    Write-Output "Available Virtual Memory: $([math]::Round($os.FreeVirtualMemory / 1MB, 2)) GB"
    Write-Output "Memory Usage: $([math]::Round((($cs.TotalPhysicalMemory - ($os.FreePhysicalMemory * 1024)) / $cs.TotalPhysicalMemory) * 100, 2))%"
    
    $pageFile = Get-CimInstance Win32_PageFileUsage -ErrorAction SilentlyContinue
    if ($pageFile) {
        Write-Output "Page File Size: $([math]::Round($pageFile.AllocatedBaseSize / 1024, 2)) GB"
        Write-Output "Page File Usage: $([math]::Round($pageFile.CurrentUsage / 1024, 2)) GB"
    }
} "60_MemoryAnalysis.txt"

Save-Data "Storage Analysis" {
    $disks = Get-CimInstance Win32_LogicalDisk
    Write-Output "=== STORAGE ANALYSIS ==="
    foreach ($disk in $disks) {
        if ($disk.Size) {
            $percentFree = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 2)
            $sizeGB = [math]::Round($disk.Size / 1GB, 2)
            $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
            $usedGB = $sizeGB - $freeGB
            
            Write-Output "Drive $($disk.DeviceID)"
            Write-Output "  Total Size: $sizeGB GB"
            Write-Output "  Used Space: $usedGB GB"
            Write-Output "  Free Space: $freeGB GB"
            Write-Output "  Percent Free: $percentFree%"
            Write-Output "  File System: $($disk.FileSystem)"
            Write-Output ""
        }
    }
} "61_StorageAnalysis.txt"

# ===== EXTERNAL DIAGNOSTICS =====
Write-Host "`n[17/25] EXTERNAL DIAGNOSTICS" -ForegroundColor Magenta

# DXDiag
Write-Host "Running DXDiag (DirectX Diagnostics)..." -ForegroundColor Cyan -NoNewline
try {
    $dxdiagPath = Join-Path $reportDir "62_DXDiag.txt"
    Start-Process dxdiag -ArgumentList "/t `"$dxdiagPath`"" -Wait -NoNewWindow
    Write-Host " COMPLETE" -ForegroundColor Green
} catch {
    Write-Host " FAILED" -ForegroundColor Red
    "DXDiag failed: $_" | Out-File -FilePath (Join-Path $reportDir "62_DXDiag.txt")
}

# MSInfo32
Write-Host "Running MSInfo32 (System Information)..." -ForegroundColor Cyan -NoNewline
try {
    $msinfoPath = Join-Path $reportDir "63_MSInfo32.nfo"
    Start-Process msinfo32 -ArgumentList "/nfo `"$msinfoPath`"" -Wait -NoNewWindow
    Write-Host " COMPLETE" -ForegroundColor Green
} catch {
    Write-Host " FAILED" -ForegroundColor Red
    "MSInfo32 failed: $_" | Out-File -FilePath (Join-Path $reportDir "63_MSInfo32_Error.txt")
}

# System File Checker
Write-Host "Running System File Checker..." -ForegroundColor Cyan -NoNewline
try {
    $sfcResult = sfc /scannow 2>&1
    $sfcResult | Out-File -FilePath (Join-Path $reportDir "64_SFC_Results.txt")
    Write-Host " COMPLETE" -ForegroundColor Green
} catch {
    Write-Host " FAILED" -ForegroundColor Red
    "SFC failed: $_" | Out-File -FilePath (Join-Path $reportDir "64_SFC_Results.txt")
}

# ===== FINAL SYSTEM STATE =====
Write-Host "`n[18/25] FINAL SYSTEM STATE" -ForegroundColor Magenta

Save-Data "Current System Uptime" {
    $bootTime = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
    $uptime = (Get-Date) - $bootTime
    
    Write-Output "=== SYSTEM UPTIME ==="
    Write-Output "Last Boot Time: $bootTime"
    Write-Output "Current Time: $(Get-Date)"
    Write-Output "Uptime: $($uptime.Days) days, $($uptime.Hours) hours, $($uptime.Minutes) minutes"
    Write-Output "Total Uptime Hours: $([math]::Round($uptime.TotalHours, 2))"
} "65_SystemUptime.txt"

Save-Data "Final Performance Snapshot" {
    Write-Output "=== FINAL PERFORMANCE SNAPSHOT ==="
    Write-Output "Timestamp: $(Get-Date)"
    
    # CPU
    $cpu = Get-Counter '\Processor(_Total)\% Processor Time' -MaxSamples 1 -ErrorAction SilentlyContinue
    if ($cpu) {
        Write-Output "CPU Usage: $([math]::Round($cpu.CounterSamples[0].CookedValue, 2))%"
    }
    
    # Memory
    $mem = Get-Counter '\Memory\Available MBytes' -MaxSamples 1 -ErrorAction SilentlyContinue
    if ($mem) {
        $totalMem = (Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB
        $availMem = $mem.CounterSamples[0].CookedValue / 1024
        Write-Output "Available Memory: $([math]::Round($availMem, 2)) GB"
        Write-Output "Memory Usage: $([math]::Round((($totalMem - $availMem) / $totalMem) * 100, 2))%"
    }
    
    # Disk
    $disk = Get-Counter '\PhysicalDisk(_Total)\Current Disk Queue Length' -MaxSamples 1 -ErrorAction SilentlyContinue
    if ($disk) {
        Write-Output "Disk Queue Length: $([math]::Round($disk.CounterSamples[0].CookedValue, 2))"
    }
} "66_FinalSnapshot.txt"

# ===== CREATE COMPREHENSIVE SUMMARY =====
Write-Host "`n[19/25] CREATING COMPREHENSIVE SUMMARY" -ForegroundColor Magenta

$summary = @"
===============================================
    COMPREHENSIVE PC HEALTH CHECK REPORT
===============================================
Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Computer: $env:COMPUTERNAME
User: $env:USERNAME
Report Location: $reportDir

SCAN OVERVIEW:
- Total Files Generated: $(Get-ChildItem $reportDir -File | Measure-Object).Count
- Scan Duration: Started at $((Get-Date).ToString())
- Administrator Rights: $((([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)))

CRITICAL FILES TO REVIEW:
1. System Errors: 39_SystemErrors_7Days.txt, 41_ApplicationErrors_7Days.txt
2. Hardware Health: 23_SMART_FailurePredict.txt, 42_HardwareEvents.txt, 43_WHEA_Events.txt
3. Performance: 31_CPU_Performance.txt, 32_Memory_Performance.txt, 33_Disk_Performance.txt
4. Storage Health: 16_PhysicalDisks.txt, 20_LogicalDisks.txt, 61_StorageAnalysis.txt
5. Memory Analysis: 60_MemoryAnalysis.txt
6. Temperature: 37_ThermalZones.txt
7. Power: 34_PowerConfig.txt, 35_Battery.txt

HARDWARE SUMMARY:
- CPU: See 05_CPU_Processor.txt
- Memory: See 06_PhysicalMemory.txt and 60_MemoryAnalysis.txt
- Storage: See 61_StorageAnalysis.txt
- Graphics: See 11_VideoController.txt
- Network: See 13_NetworkAdapters.txt

SOFTWARE SUMMARY:
- Installed Programs: See 45_InstalledPrograms.txt
- Services: See 47_Services.txt
- Drivers: See 46_SystemDrivers.txt
- Updates: See 51_HotFixes.txt

EXTERNAL DIAGNOSTICS:
- DirectX: See 62_DXDiag.txt
- System Info: See 63_MSInfo32.nfo
- System Files: See 64_SFC_Results.txt

PERFORMANCE METRICS:
- Current State: See 66_FinalSnapshot.txt
- Historical: See performance counter files (31-33)

RECOMMENDATIONS:
1. Review all ERROR entries in event log files
2. Check SMART status for drive health warnings
3. Monitor temperature readings if available
4. Verify critical services are running
5. Review memory usage patterns
6. Check for hardware error events (WHEA logs)
7. Ensure Windows updates are current

NEXT STEPS:
- Run this scan monthly to track changes
- Compare results over time for trend analysis
- Address any critical errors found
- Monitor performance degradation
- Keep baseline copies for comparison

For detailed analysis, review individual files in numerical order.
Critical issues will be found in error logs and SMART data files.

Report Generation Complete: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
===============================================
"@

$summary | Out-File -FilePath (Join-Path $reportDir "00_COMPREHENSIVE_SUMMARY.txt") -Encoding UTF8

# ===== CREATE ZIP ARCHIVE =====
Write-Host "`n[20/25] CREATING ZIP ARCHIVE" -ForegroundColor Magenta

try {
    $zipPath = Join-Path $env:USERPROFILE "Desktop\PC_Health_FULL_Report_$timestamp.zip"
    Write-Host "Creating ZIP archive..." -ForegroundColor Cyan -NoNewline
    Compress-Archive -Path "$reportDir\*" -DestinationPath $zipPath -Force
    Write-Host " COMPLETE" -ForegroundColor Green
    Write-Host "ZIP Location: $zipPath" -ForegroundColor Yellow
} catch {
    Write-Host " FAILED: $_" -ForegroundColor Red
    Write-Host "Individual files available in: $reportDir" -ForegroundColor Yellow
}

# ===== COMPLETION =====
Write-Host "`n=============================================" -ForegroundColor Green
Write-Host "    COMPREHENSIVE PC HEALTH SCAN COMPLETE!" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green
Write-Host "Total Files Created: $((Get-ChildItem $reportDir -File | Measure-Object).Count)" -ForegroundColor Cyan
Write-Host "Report Folder: $reportDir" -ForegroundColor Yellow
if (Test-Path $zipPath) {
    Write-Host "ZIP Archive: $zipPath" -ForegroundColor Yellow
}
Write-Host "`nKey Files to Check First:" -ForegroundColor Cyan
Write-Host "- 00_COMPREHENSIVE_SUMMARY.txt (This file!)" -ForegroundColor White
Write-Host "- 39_SystemErrors_7Days.txt" -ForegroundColor White
Write-Host "- 23_SMART_FailurePredict.txt" -ForegroundColor White
Write-Host "- 60_MemoryAnalysis.txt" -ForegroundColor White
Write-Host "- 61_StorageAnalysis.txt" -ForegroundColor White

# Open the report folder
Start-Process explorer.exe -ArgumentList $reportDir

Write-Host "`nOpening report folder..." -ForegroundColor Green
Write-Host "Upload the ZIP file for detailed analysis!" -ForegroundColor Yellow
Write-Host "`nPress any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")