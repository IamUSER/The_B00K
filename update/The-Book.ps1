# The-Book.ps1 - A Windows System Administration and Diagnostic Tool
# Version: 1.0.0
# Compatibility: Windows 7 through Windows 11
# Author: System Administrator

#Requires -Version 3.0
#Requires -RunAsAdministrator

# Import required modules
Import-Module -Name Microsoft.PowerShell.Utility
Import-Module -Name Microsoft.PowerShell.Management

# Import configuration
$script:Config = @{}
$script:ConfigPath = Join-Path $PSScriptRoot "The-Book.config.json"

# Initialize logging
$script:LogPath = Join-Path $PSScriptRoot "Logs"
$script:LogFile = Join-Path $script:LogPath "The-Book_$(Get-Date -Format 'yyyyMMdd').log"

# Ensure required directories exist
if (-not (Test-Path $script:LogPath)) {
    New-Item -ItemType Directory -Path $script:LogPath | Out-Null
}

# Load configuration
function Load-Configuration {
    try {
        if (Test-Path $script:ConfigPath) {
            $script:Config = Get-Content $script:ConfigPath | ConvertFrom-Json -AsHashtable
        } else {
            # Default configuration
            $script:Config = @{
                Logging = @{
                    MaxLogSize = 5MB
                    MaxLogFiles = 10
                    LogLevel = "Information"
                }
                SystemInfo = @{
                    CollectDetailedInfo = $true
                    IncludeNetworkInfo = $true
                    IncludeInstalledApps = $true
                }
                UI = @{
                    ShowBanner = $true
                    UseColors = $true
                }
            }
            # Save default configuration
            $script:Config | ConvertTo-Json | Set-Content $script:ConfigPath
        }
    }
    catch {
        Write-Log -Level Error -Message "Failed to load configuration: $_"
    }
}

# Logging function
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet("Error", "Warning", "Information", "Debug")]
        [string]$Level,
        
        [Parameter(Mandatory=$true)]
        [string]$Message
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Check log file size and rotate if necessary
    if (Test-Path $script:LogFile) {
        $logSize = (Get-Item $script:LogFile).Length
        if ($logSize -gt $script:Config.Logging.MaxLogSize) {
            $archivePath = Join-Path $script:LogPath "The-Book_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
            Move-Item $script:LogFile $archivePath
        }
    }
    
    # Write to log file
    Add-Content -Path $script:LogFile -Value $logEntry
    
    # Write to console based on log level
    switch ($Level) {
        "Error" { Write-Host $logEntry -ForegroundColor Red }
        "Warning" { Write-Host $logEntry -ForegroundColor Yellow }
        "Information" { Write-Host $logEntry -ForegroundColor White }
        "Debug" { Write-Host $logEntry -ForegroundColor Gray }
    }
}

# System Information Collection
function Get-SystemInformation {
    Write-Log -Level Information -Message "Collecting system information..."
    
    $systemInfo = @{
        ComputerName = $env:COMPUTERNAME
        UserName = $env:USERNAME
        OSVersion = [System.Environment]::OSVersion.VersionString
        Architecture = (Get-WmiObject Win32_OperatingSystem).OSArchitecture
        Processor = (Get-WmiObject Win32_Processor).Name
        Memory = [math]::Round((Get-WmiObject Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
        Drives = Get-WmiObject Win32_LogicalDisk | Select-Object DeviceID, Size, FreeSpace
        NetworkAdapters = Get-NetAdapter | Select-Object Name, InterfaceDescription, Status
    }
    
    if ($script:Config.SystemInfo.CollectDetailedInfo) {
        $systemInfo.DetailedInfo = @{
            BIOS = Get-WmiObject Win32_BIOS | Select-Object Manufacturer, SerialNumber, Version
            SystemManufacturer = (Get-WmiObject Win32_ComputerSystem).Manufacturer
            SystemModel = (Get-WmiObject Win32_ComputerSystem).Model
            SystemSerialNumber = (Get-WmiObject Win32_BIOS).SerialNumber
        }
    }
    
    if ($script:Config.SystemInfo.IncludeNetworkInfo) {
        $systemInfo.NetworkInfo = @{
            IPAddresses = Get-NetIPAddress | Where-Object { $_.AddressFamily -eq "IPv4" } | Select-Object IPAddress, InterfaceAlias
            DNS = Get-DnsClientServerAddress | Select-Object InterfaceAlias, ServerAddresses
            Routes = Get-NetRoute | Where-Object { $_.DestinationPrefix -ne "0.0.0.0/0" } | Select-Object DestinationPrefix, NextHop
        }
    }
    
    if ($script:Config.SystemInfo.IncludeInstalledApps) {
        $systemInfo.InstalledApps = Get-WmiObject Win32_Product | Select-Object Name, Version, Vendor
    }
    
    return $systemInfo
}

# Main menu function
function Show-MainMenu {
    Clear-Host
    if ($script:Config.UI.ShowBanner) {
        Write-Host @"
    ____  _           ____             _    
   | __ )| |_   _ ___| __ )  ___   ___| | __
   |  _ \| | | | / __|  _ \ / _ \ / __| |/ /
   | |_) | | |_| \__ \ |_) | (_) | (__|   < 
   |____/|_|\__,_|___/____/ \___/ \___|_|\_\
   
   Windows System Administration Tool
   Version 1.0.0
"@ -ForegroundColor Cyan
    }
    
    Write-Host "`nMain Menu:`n" -ForegroundColor White
    Write-Host "1. System Information" -ForegroundColor Green
    Write-Host "2. Network Tools" -ForegroundColor Green
    Write-Host "3. System Configuration" -ForegroundColor Green
    Write-Host "4. User Management" -ForegroundColor Green
    Write-Host "5. Diagnostics" -ForegroundColor Green
    Write-Host "6. Settings" -ForegroundColor Green
    Write-Host "7. Help" -ForegroundColor Green
    Write-Host "X. Exit" -ForegroundColor Red
    Write-Host "`n"
}

# Main script execution
try {
    # Load configuration
    Load-Configuration
    
    # Check for administrator privileges
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Log -Level Error -Message "This script requires administrator privileges. Please run as administrator."
        exit 1
    }
    
    # Main loop
    $running = $true
    while ($running) {
        Show-MainMenu
        $choice = Read-Host "Select an option"
        
        switch ($choice) {
            "1" {
                $systemInfo = Get-SystemInformation
                $systemInfo | ConvertTo-Json | Out-File (Join-Path $script:LogPath "SystemInfo_$(Get-Date -Format 'yyyyMMdd_HHmmss').json")
                Write-Host "System information collected and saved to logs directory." -ForegroundColor Green
                Pause
            }
            "2" {
                # Network tools menu
                Show-NetworkMenu
            }
            "3" {
                # System configuration menu
                Show-SystemConfigMenu
            }
            "4" {
                # User management menu
                Show-UserManagementMenu
            }
            "5" {
                # Diagnostics menu
                Show-DiagnosticsMenu
            }
            "6" {
                # Settings menu
                Show-SettingsMenu
            }
            "7" {
                # Help menu
                Show-HelpMenu
            }
            "X" {
                $running = $false
            }
            default {
                Write-Host "Invalid option. Please try again." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    }
}
catch {
    Write-Log -Level Error -Message "An error occurred: $_"
    Write-Log -Level Error -Message "Stack trace: $($_.ScriptStackTrace)"
    exit 1
}
finally {
    Write-Log -Level Information -Message "Script execution completed"
} 