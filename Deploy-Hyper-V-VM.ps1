<#

.SYNOPSIS
Interactive Hyper-V VM Deployment Script

.DESCRIPTION
- This script creates a new Hyper-V virtual machine with user-specified parameters

.EXAMPLE
.\Deploy-Hyper-V-VM.ps1

.NOTES
    Author: Yoennis Olmo
    Version: v1.0
    Release Date: 06-01-2026

#>

#Requires -RunAsAdministrator

# --- Helper Functions ---

Add-Type -AssemblyName System.Windows.Forms

# Opens a FolderBrowserDialog on a dedicated STA runspace thread.
# Using a dedicated runspace with ApartmentState.STA works on both
# PS5.1 (.NET Framework) and PS7 (.NET Core / .NET 5+) without any
# C# compilation, which avoids assembly-reference differences between
# the two runtimes.
function Get-FolderPath {
    param([string]$Description, [string]$InitialPath = 'C:\')
    $startPath = if (Test-Path $InitialPath) { $InitialPath } else { 'C:\' }

    $ps = [System.Management.Automation.PowerShell]::Create()
    $null = $ps.AddScript({
        param($desc, $path)
        Add-Type -AssemblyName System.Windows.Forms
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description         = $desc
        $dlg.SelectedPath        = $path
        $dlg.ShowNewFolderButton = $true
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $dlg.SelectedPath
        }
    }).AddArgument($Description).AddArgument($startPath)

    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = [System.Threading.ApartmentState]::STA
    $rs.Open()
    $ps.Runspace = $rs
    $result = $ps.Invoke() | Select-Object -First 1
    $rs.Close()
    $ps.Dispose()
    return $result
}

# Opens an OpenFileDialog for ISO selection on a dedicated STA runspace thread.
function Get-ISOPath {
    $ps = [System.Management.Automation.PowerShell]::Create()
    $null = $ps.AddScript({
        Add-Type -AssemblyName System.Windows.Forms
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title            = 'Select ISO File'
        $dlg.Filter           = 'ISO Files (*.iso)|*.iso|All Files (*.*)|*.*'
        $dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $dlg.FileName
        }
    })

    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = [System.Threading.ApartmentState]::STA
    $rs.Open()
    $ps.Runspace = $rs
    $result = $ps.Invoke() | Select-Object -First 1
    $rs.Close()
    $ps.Dispose()
    return $result
}

# --- Pre-Flight ---

# --- Step 1: Check if Hyper-V is installed ---
$hyperVFeature = Get-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V-All -ErrorAction SilentlyContinue

if ($hyperVFeature -and $hyperVFeature.State -eq 'Enabled') {
    $hyperVAlreadyInstalled = $true
} else {
    Write-Host "Hyper-V is not installed. Installing now (this requires Administrator privileges)..." -ForegroundColor Yellow
    $hyperVAlreadyInstalled = $false
    try {
        Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V-All -Verbose -ErrorAction Stop
        Write-Host "Hyper-V installed successfully." -ForegroundColor Green
        Write-Host ""
        Write-Host "*** A restart is required to complete the Hyper-V installation. ***" -ForegroundColor Yellow
        Write-Host "Re-run this script after restarting to continue VM creation." -ForegroundColor Yellow
        Write-Host ""
        $restartNow = Read-Host "Restart the computer now? (Y/N)"
        if ($restartNow -eq "Y" -or $restartNow -eq "y") {
            Write-Host "Restarting in 10 seconds. Press Ctrl+C to cancel." -ForegroundColor Cyan
            Start-Sleep -Seconds 10
            Restart-Computer -Force
        } else {
            Write-Host "Please restart manually and re-run this script." -ForegroundColor Yellow
            exit
        }
    } catch {
        Write-Host "Failed to install Hyper-V: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Please install Hyper-V manually and re-run this script." -ForegroundColor Red
        exit
    }
}

# --- Step 2: Determine $VMPath and $VHDBasePath ---
$VMPath = $null
$VHDBasePath = $null

if ($hyperVAlreadyInstalled) {
    try {
        $vmHost = Get-VMHost -ErrorAction Stop
        $existingVMPath  = $vmHost.VirtualMachinePath
        $existingVHDPath = $vmHost.VirtualHardDiskPath

        $isWindowsDefault = ($existingVMPath  -match '\\Users\\' -or $existingVMPath  -match '\\ProgramData\\') -or
                            ($existingVHDPath -match '\\Users\\' -or $existingVHDPath -match '\\ProgramData\\')

        if (!$isWindowsDefault) {
            $VMPath      = $existingVMPath
            $VHDBasePath = $existingVHDPath
        }
    } catch {
        Write-Host "Warning: Could not read VMHost paths: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

if ($null -eq $VMPath) {
    Write-Host "Opening folder browser for VM Configuration path..." -ForegroundColor Cyan
    $chosen  = Get-FolderPath -Description "Select VM Configuration Folder" -InitialPath 'C:\'
    $VMPath  = if ($chosen) { $chosen } else { 'C:\HYPER-V' }
    Write-Host "VM Config Path: $VMPath" -ForegroundColor Green

    Write-Host "Opening folder browser for Virtual Hard Disk path..." -ForegroundColor Cyan
    $chosen      = Get-FolderPath -Description "Select Virtual Hard Disk Folder" -InitialPath $VMPath
    $VHDBasePath = if ($chosen) { $chosen } else { $VMPath }
    Write-Host "VHD Base Path: $VHDBasePath" -ForegroundColor Green
}

# --- Step 3: Create both directories ---
foreach ($dir in @($VMPath, $VHDBasePath)) {
    if (!(Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

# --- Step 4: Configure VMHost ---
try {
    $vmHost = Get-VMHost -ErrorAction Stop
    $needsConfig = ($vmHost.VirtualHardDiskPath -ne $VHDBasePath) -or
                   ($vmHost.VirtualMachinePath  -ne $VMPath) -or
                   (-not $vmHost.EnableEnhancedSessionMode)

    if ($needsConfig) {
        Write-Host "Applying Hyper-V host settings..." -ForegroundColor Cyan
        Set-VMHost -VirtualHardDiskPath $VHDBasePath
        Set-VMHost -VirtualMachinePath  $VMPath
        Set-VMHost -EnableEnhancedSessionMode $true
        Write-Host "Hyper-V host configuration applied." -ForegroundColor Green
    }
} catch {
    Write-Host "Warning: Could not verify/apply VMHost settings: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "This may happen immediately after first install - continue after a restart." -ForegroundColor Yellow
}

Write-Host ""

# Prompt for VM parameters
$VMName = Read-Host "Enter VM Name"
while ([string]::IsNullOrWhiteSpace($VMName)) {
    Write-Host "VM Name cannot be empty!" -ForegroundColor Red
    $VMName = Read-Host "Enter VM Name"
}

if (Get-VM -Name $VMName -ErrorAction SilentlyContinue) {
    Write-Host "A VM named '$VMName' already exists. Please choose a different name." -ForegroundColor Red
    exit
}

$VHDSizeInput = Read-Host "Enter VHD Size in GB (e.g., 128)"
while (![int]::TryParse($VHDSizeInput, [ref]$null) -or [int]$VHDSizeInput -le 0) {
    Write-Host "Please enter a valid positive number for VHD size!" -ForegroundColor Red
    $VHDSizeInput = Read-Host "Enter VHD Size in GB (e.g., 128)"
}
$VHDSize = [int64]$VHDSizeInput * 1GB

$MemoryInput = Read-Host "Enter Memory in GB (e.g., 4)"
while (![int]::TryParse($MemoryInput, [ref]$null) -or [int]$MemoryInput -le 0) {
    Write-Host "Please enter a valid positive number for memory!" -ForegroundColor Red
    $MemoryInput = Read-Host "Enter Memory in GB (e.g., 4)"
}
$Memory = [int64]$MemoryInput * 1GB

$ProcessorCountInput = Read-Host "Enter Processor Count (e.g., 2)"
while (![int]::TryParse($ProcessorCountInput, [ref]$null) -or [int]$ProcessorCountInput -le 0) {
    Write-Host "Please enter a valid positive number for processor count!" -ForegroundColor Red
    $ProcessorCountInput = Read-Host "Enter Processor Count (e.g., 2)"
}
$ProcessorCount = [int]$ProcessorCountInput

# Get available virtual switches and let user choose
Write-Host ""
Write-Host "=== Virtual Switch Configuration ===" -ForegroundColor Yellow
$SwitchName = ""
$UseSwitch  = $false

# Silent internet connectivity check (1-second timeout, no output)
$hasInternet = Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet -ErrorAction SilentlyContinue

try {
    $AvailableSwitches = @(Get-VMSwitch | Select-Object Name, SwitchType)
    if ($AvailableSwitches.Count -eq 0) {
        Write-Host "No virtual switches found!" -ForegroundColor Yellow
        Write-Host "The VM will be created without network connectivity." -ForegroundColor Yellow
        Write-Host "You can add a network adapter later through Hyper-V Manager." -ForegroundColor Cyan
    } else {
        if (-not $hasInternet) {
            Write-Host "[!] No internet connection detected on this host." -ForegroundColor Yellow
            Write-Host "    Default Switch will not provide internet access to the VM." -ForegroundColor Yellow
            Write-Host ""
        }

        Write-Host "Available Virtual Switches:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $AvailableSwitches.Count; $i++) {
            $switchLabel = $AvailableSwitches[$i].Name
            if ($switchLabel -eq 'Default Switch') {
                if ($hasInternet) {
                    $switchLabel += ' (Note: Requires Internet Connection)'
                } else {
                    $switchLabel += ' [No Internet Detected - Limited Connectivity]'
                }
            }
            Write-Host "[$($i + 1)] $switchLabel ($($AvailableSwitches[$i].SwitchType))" -ForegroundColor Cyan
        }
        Write-Host "[0] Skip - Create VM without network adapter" -ForegroundColor Cyan

        Write-Host ""
        $SwitchChoice = Read-Host "Select a virtual switch by number (0-$($AvailableSwitches.Count))"
        while (![int]::TryParse($SwitchChoice, [ref]$null) -or [int]$SwitchChoice -lt 0 -or [int]$SwitchChoice -gt $AvailableSwitches.Count) {
            Write-Host "Please enter a valid number between 0 and $($AvailableSwitches.Count)!" -ForegroundColor Red
            $SwitchChoice = Read-Host "Select a virtual switch by number (0-$($AvailableSwitches.Count))"
        }

        if ([int]$SwitchChoice -eq 0) {
            Write-Host "Network configuration skipped. VM will be created without network connectivity." -ForegroundColor Yellow
        } else {
            $SwitchName = $AvailableSwitches[[int]$SwitchChoice - 1].Name
            $UseSwitch  = $true

            # Extra confirmation when selecting Default Switch with no internet
            if (-not $hasInternet -and $SwitchName -eq 'Default Switch') {
                Write-Host ""
                Write-Host "  Warning: You selected 'Default Switch' but no internet was detected on this host." -ForegroundColor Yellow
                Write-Host "  The VM will be connected to the switch but may have no internet access." -ForegroundColor Yellow
                $confirmSwitch = Read-Host "  Proceed with Default Switch anyway? (Y/N)"
                if ($confirmSwitch -ne 'Y' -and $confirmSwitch -ne 'y') {
                    $SwitchName = ""
                    $UseSwitch  = $false
                    Write-Host "Network configuration skipped. VM will be created without network connectivity." -ForegroundColor Yellow
                } else {
                    Write-Host "Selected switch: $SwitchName" -ForegroundColor Green
                }
            } else {
                Write-Host "Selected switch: $SwitchName" -ForegroundColor Green
            }
        }
    }
}
catch {
    Write-Host "Error retrieving virtual switches: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "The VM will be created without network connectivity." -ForegroundColor Yellow
    Write-Host "You can add a network adapter later through Hyper-V Manager." -ForegroundColor Cyan
}

# ISO Path - Allow empty/optional with browse option
Write-Host ""
Write-Host "=== ISO File Selection ===" -ForegroundColor Yellow
Write-Host "[1] Browse for ISO file" -ForegroundColor Cyan
Write-Host "[2] Enter ISO path manually" -ForegroundColor Cyan
Write-Host "[3] Skip ISO (no ISO will be mounted)" -ForegroundColor Cyan
Write-Host ""

$ISOChoice = Read-Host "Select an option (1-3)"
while (![int]::TryParse($ISOChoice, [ref]$null) -or [int]$ISOChoice -lt 1 -or [int]$ISOChoice -gt 3) {
    Write-Host "Please enter a valid number between 1 and 3!" -ForegroundColor Red
    $ISOChoice = Read-Host "Select an option (1-3)"
}

$ISOPath = ""
$UseISO = $false

switch ([int]$ISOChoice) {
    1 {
        # Browse for ISO file
        Write-Host "Opening file browser..." -ForegroundColor Cyan
        try {
            $ISOPath = Get-ISOPath
            if ($ISOPath -and (Test-Path $ISOPath)) {
                $UseISO = $true
                Write-Host "ISO selected: $ISOPath" -ForegroundColor Green
            } else {
                Write-Host "No ISO file selected. No ISO will be mounted." -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "Error opening file browser: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "No ISO will be mounted." -ForegroundColor Yellow
        }
    }
    2 {
        # Manual entry
        $ISOPath = Read-Host "Enter ISO Path (full path to ISO file)"
        if (![string]::IsNullOrWhiteSpace($ISOPath)) {
            while (!(Test-Path $ISOPath)) {
                Write-Host "ISO file not found at specified path!" -ForegroundColor Red
                $ISOPath = Read-Host "Enter ISO Path (full path to ISO file, or press Enter to skip)"
                if ([string]::IsNullOrWhiteSpace($ISOPath)) {
                    break
                }
            }

            if (![string]::IsNullOrWhiteSpace($ISOPath)) {
                $UseISO = $true
                Write-Host "ISO will be mounted: $ISOPath" -ForegroundColor Green
            } else {
                Write-Host "No ISO will be mounted." -ForegroundColor Yellow
            }
        } else {
            Write-Host "No ISO will be mounted." -ForegroundColor Yellow
        }
    }
    3 {
        # Skip ISO
        Write-Host "No ISO will be mounted." -ForegroundColor Yellow
    }
}

# Build VHD path from $VHDBasePath
$VHDPath = "$VHDBasePath\$VMName\Virtual Hard Disks\$VMName.vhdx"

Write-Host ""
Write-Host "=== VM Configuration Summary ===" -ForegroundColor Yellow
Write-Host "VM Name: $VMName"
Write-Host "VHD Size: $($VHDSize/1GB)GB"
Write-Host "Memory: $($Memory/1GB)GB"
Write-Host "Processor Count: $ProcessorCount"
if ($UseSwitch) {
    Write-Host "Virtual Switch: $SwitchName"
} else {
    Write-Host "Virtual Switch: None (No network connectivity)"
}
if ($UseISO) {
    Write-Host "ISO Path: $ISOPath"
} else {
    Write-Host "ISO Path: None (No ISO will be mounted)"
}
Write-Host "VM Config Path   : $VMPath"
Write-Host "VHD Path         : $VHDPath"
Write-Host ""

$confirm = Read-Host "Do you want to proceed with VM creation? (Y/N)"
if ($confirm -ne "Y" -and $confirm -ne "y") {
    Write-Host "VM creation cancelled." -ForegroundColor Yellow
    exit
}

Write-Host ""
Write-Host "=== Starting VM Creation ===" -ForegroundColor Green

try {
    # Create VHD parent directory before New-VHD
    Write-Host "Creating VHD directory..." -ForegroundColor Cyan
    $vhdDir = Split-Path $VHDPath -Parent
    if (!(Test-Path $vhdDir)) { New-Item -ItemType Directory -Path $vhdDir -Force | Out-Null }

    # Create the virtual hard disk
    Write-Host "Creating virtual hard disk..." -ForegroundColor Cyan
    New-VHD -Path $VHDPath -SizeBytes $VHDSize -Dynamic

    # Create the virtual machine
    Write-Host "Creating virtual machine..." -ForegroundColor Cyan
    New-VM -Name $VMName -Path $VMPath -VHDPath $VHDPath -Generation 2

    # Configure VM settings
    Write-Host "Configuring VM settings..." -ForegroundColor Cyan
    Set-VM -Name $VMName -ProcessorCount $ProcessorCount -MemoryStartupBytes $Memory -StaticMemory

    # Configure secure boot (Generation 2 VMs only)
    Write-Host "Configuring secure boot..." -ForegroundColor Cyan
    Set-VMFirmware -VMName $VMName -EnableSecureBoot On -SecureBootTemplate "MicrosoftWindows"

    # Add DVD drive and mount ISO only if ISO path is provided
    if ($UseISO) {
        Write-Host "Adding DVD drive and mounting ISO..." -ForegroundColor Cyan
        Add-VMDvdDrive -VMName $VMName -Path $ISOPath
    } else {
        Write-Host "Adding DVD drive (no ISO mounted)..." -ForegroundColor Cyan
        Add-VMDvdDrive -VMName $VMName
    }

    # Set DVD drive as first boot option
    Write-Host "Setting DVD drive as first boot option..." -ForegroundColor Cyan
    $DVDDrive       = Get-VMDvdDrive       -VMName $VMName
    $HardDrive      = Get-VMHardDiskDrive  -VMName $VMName
    $NetworkAdapter = Get-VMNetworkAdapter -VMName $VMName
    Set-VMFirmware -VMName $VMName -BootOrder $DVDDrive, $HardDrive, $NetworkAdapter

    # Configure Key Protector and enable TPM
    Write-Host "Configuring TPM..." -ForegroundColor Cyan
    Set-VMKeyProtector -VMName $VMName -NewLocalKeyProtector
    Enable-VMTPM -VMName $VMName

    # Connect the existing NIC (only if a switch was selected)
    if ($UseSwitch) {
        Write-Host "Connecting network adapter to switch '$SwitchName'..." -ForegroundColor Cyan
        Connect-VMNetworkAdapter -VMName $VMName -SwitchName $SwitchName
    } else {
        Write-Host "Skipping network adapter configuration (no switch selected)..." -ForegroundColor Yellow
    }

    # Configure additional VM settings
    Write-Host "Configuring additional settings..." -ForegroundColor Cyan
    Set-VM -Name $VMName -AutomaticCheckpointsEnabled $false -CheckpointType Standard

    # Enable all integration services
    Write-Host "Enabling integration services..." -ForegroundColor Cyan
    Enable-VMIntegrationService -VMName $VMName -Name 'Guest Service Interface', 'Heartbeat', 'Key-Value Pair Exchange', 'Shutdown', 'Time Synchronization', 'VSS'

    # Enable nested virtualization (if needed)
    Set-VMProcessor -VMName $VMName -ExposeVirtualizationExtensions $true

    # Enable Enhanced Session Mode
    Set-VM -VMName $VMName -EnhancedSessionTransportType VMBus

    Write-Host ""
    Write-Host "=== VM Creation Completed Successfully! ===" -ForegroundColor Green
    Write-Host ""
    Write-Host "Virtual Machine '$VMName' created successfully with the following configuration:" -ForegroundColor White
    Write-Host "- RAM: $($Memory/1GB)GB" -ForegroundColor White
    Write-Host "- Processors: $ProcessorCount" -ForegroundColor White
    Write-Host "- Hard Drive: $($VHDSize/1GB)GB (Dynamic VHDX)" -ForegroundColor White
    Write-Host "- Secure Boot: Enabled" -ForegroundColor White
    Write-Host "- TPM: Enabled" -ForegroundColor White
    if ($UseSwitch) {
        Write-Host "- Network Adapter: Connected to '$SwitchName'" -ForegroundColor White
    } else {
        Write-Host "- Network Adapter: Not Connected (can be configured later)" -ForegroundColor White
    }
    if ($UseISO) {
        Write-Host "- ISO Mounted: $ISOPath" -ForegroundColor White
    } else {
        Write-Host "- ISO Mounted: None" -ForegroundColor White
    }
    Write-Host "- Integration Services: Enabled (Guest Service Interface, Heartbeat, Key-Value Pair Exchange, Shutdown, Time Synchronization, VSS)" -ForegroundColor White
    Write-Host "- Expose Virtualization Extensions: Enabled" -ForegroundColor White
    Write-Host "- Enhanced Session Mode: Enabled" -ForegroundColor White
    Write-Host "- Boot Order: DVD Drive (First), Hard Drive, Network Adapter" -ForegroundColor White

    $startVM = Read-Host "Would you like to start the VM now? (Y/N)"
    if ($startVM -eq "Y" -or $startVM -eq "y") {
        Write-Host "Starting VM..." -ForegroundColor Cyan
        Start-VM -Name $VMName
        Write-Host "VM '$VMName' has been started!" -ForegroundColor Green
    }
}
catch {
    Write-Host ""
    Write-Host "=== Error Occurred ===" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please check the following:" -ForegroundColor Yellow
    Write-Host "- Hyper-V is enabled and you have administrative privileges"
    if ($UseSwitch) {
        Write-Host "- The specified virtual switch '$SwitchName' exists"
    }
    Write-Host "- There's enough disk space for the VHD"
    Write-Host "- The VM name doesn't already exist"
}
