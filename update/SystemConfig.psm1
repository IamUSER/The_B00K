# SystemConfig.psm1 - System configuration and management tools

function Write-SystemConfigLog {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet("Error", "Warning", "Information", "Debug")]
        [string]$Level,
        
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [string]$FeatureName,
        
        [Parameter(Mandatory=$false)]
        [string]$ServiceName,
        
        [Parameter(Mandatory=$false)]
        [string]$ProgramPath,
        
        [Parameter(Mandatory=$false)]
        [string]$AdditionalInfo
    )
    
    $logMessage = $Message
    if ($FeatureName) { $logMessage = $logMessage -replace "{FeatureName}", $FeatureName }
    if ($ServiceName) { $logMessage = $logMessage -replace "{ServiceName}", $ServiceName }
    if ($ProgramPath) { $logMessage = $logMessage -replace "{ProgramPath}", $ProgramPath }
    if ($AdditionalInfo) { $logMessage = $logMessage -replace "{AdditionalInfo}", $AdditionalInfo }
    
    Write-Log -Level $Level -Message $logMessage
}

function Show-SystemConfigMenu {
    Clear-Host
    Write-Host "System Configuration Menu:`n" -ForegroundColor White
    Write-Host "1. Windows Features" -ForegroundColor Green
    Write-Host "2. Services Management" -ForegroundColor Green
    Write-Host "3. Startup Programs" -ForegroundColor Green
    Write-Host "4. Power Settings" -ForegroundColor Green
    Write-Host "5. System Properties" -ForegroundColor Green
    Write-Host "B. Back to Main Menu" -ForegroundColor Yellow
    Write-Host "X. Exit" -ForegroundColor Red
    Write-Host "`n"
    
    $choice = Read-Host "Select an option"
    
    switch ($choice) {
        "1" { Show-WindowsFeatures }
        "2" { Show-ServicesMenu }
        "3" { Show-StartupPrograms }
        "4" { Show-PowerSettings }
        "5" { Show-SystemProperties }
        "B" { return }
        "X" { exit }
        default {
            Write-Host "Invalid option. Please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Show-SystemConfigMenu
        }
    }
}

function Show-WindowsFeatures {
    Write-SystemConfigLog -Level Information -Message "Displaying Windows features"
    
    if (Get-Command -Name Get-WindowsOptionalFeature -ErrorAction SilentlyContinue) {
        $features = Get-WindowsOptionalFeature -Online
        $enabledFeatures = $features | Where-Object { $_.State -eq "Enabled" }
        $disabledFeatures = $features | Where-Object { $_.State -eq "Disabled" }
        
        Write-Host "`nEnabled Features:" -ForegroundColor Green
        $enabledFeatures | Format-Table -Property FeatureName, State
        
        Write-Host "`nDisabled Features:" -ForegroundColor Yellow
        $disabledFeatures | Format-Table -Property FeatureName, State
        
        $choice = Read-Host "`nDo you want to enable/disable any features? (Y/N)"
        if ($choice -eq "Y") {
            Manage-WindowsFeatures
        }
    }
    else {
        Write-Host "Windows Features management is not available on this system." -ForegroundColor Red
    }
    
    Pause
    Show-SystemConfigMenu
}

function Manage-WindowsFeatures {
    $featureName = Read-Host "Enter feature name (or 'list' to see all features)"
    
    if ($featureName -eq "list") {
        Get-WindowsOptionalFeature -Online | Format-Table -Property FeatureName, State
        return
    }
    
    $action = Read-Host "Enable or Disable? (E/D)"
    try {
        if ($action -eq "E") {
            Enable-WindowsOptionalFeature -Online -FeatureName $featureName -NoRestart
            Write-Host "Feature enabled successfully." -ForegroundColor Green
            Write-SystemConfigLog -Level Information -Message "Enabled Windows feature: {FeatureName}" -FeatureName $featureName
        }
        elseif ($action -eq "D") {
            Disable-WindowsOptionalFeature -Online -FeatureName $featureName -NoRestart
            Write-Host "Feature disabled successfully." -ForegroundColor Green
            Write-SystemConfigLog -Level Information -Message "Disabled Windows feature: {FeatureName}" -FeatureName $featureName
        }
    }
    catch {
        Write-Host "Failed to modify feature: $_" -ForegroundColor Red
        Write-SystemConfigLog -Level Error -Message "Failed to modify Windows feature: {AdditionalInfo}" -AdditionalInfo $_
    }
}

function Show-ServicesMenu {
    Write-SystemConfigLog -Level Information -Message "Displaying services menu"
    
    Write-Host "`nServices Management:" -ForegroundColor Cyan
    Write-Host "1. List all services"
    Write-Host "2. List running services"
    Write-Host "3. List stopped services"
    Write-Host "4. Manage specific service"
    Write-Host "B. Back"
    
    $choice = Read-Host "Select an option"
    
    switch ($choice) {
        "1" { Get-Service | Format-Table -Property Name, Status, StartType }
        "2" { Get-Service | Where-Object { $_.Status -eq "Running" } | Format-Table -Property Name, Status, StartType }
        "3" { Get-Service | Where-Object { $_.Status -eq "Stopped" } | Format-Table -Property Name, Status, StartType }
        "4" { Manage-Service }
        "B" { return }
        default {
            Write-Host "Invalid option." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
    
    Pause
    Show-ServicesMenu
}

function Manage-Service {
    $serviceName = Read-Host "Enter service name"
    $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
    
    if ($service) {
        Write-Host "`nService: $($service.DisplayName)" -ForegroundColor Cyan
        Write-Host "Status: $($service.Status)"
        Write-Host "Start Type: $($service.StartType)"
        
        Write-Host "`nActions:" -ForegroundColor Yellow
        Write-Host "1. Start service"
        Write-Host "2. Stop service"
        Write-Host "3. Restart service"
        Write-Host "4. Change startup type"
        
        $action = Read-Host "Select action"
        
        try {
            switch ($action) {
                "1" {
                    Start-Service -Name $serviceName
                    Write-Host "Service started successfully." -ForegroundColor Green
                    Write-SystemConfigLog -Level Information -Message "Started service: {ServiceName}" -ServiceName $serviceName
                }
                "2" {
                    Stop-Service -Name $serviceName
                    Write-Host "Service stopped successfully." -ForegroundColor Green
                    Write-SystemConfigLog -Level Information -Message "Stopped service: {ServiceName}" -ServiceName $serviceName
                }
                "3" {
                    Restart-Service -Name $serviceName
                    Write-Host "Service restarted successfully." -ForegroundColor Green
                    Write-SystemConfigLog -Level Information -Message "Restarted service: {ServiceName}" -ServiceName $serviceName
                }
                "4" {
                    Write-Host "`nStartup Types:" -ForegroundColor Yellow
                    Write-Host "1. Automatic"
                    Write-Host "2. Automatic (Delayed Start)"
                    Write-Host "3. Manual"
                    Write-Host "4. Disabled"
                    
                    $startType = Read-Host "Select startup type"
                    $newStartType = switch ($startType) {
                        "1" { "Automatic" }
                        "2" { "AutomaticDelayedStart" }
                        "3" { "Manual" }
                        "4" { "Disabled" }
                        default { throw "Invalid startup type" }
                    }
                    
                    Set-Service -Name $serviceName -StartupType $newStartType
                    Write-Host "Startup type changed successfully." -ForegroundColor Green
                    Write-SystemConfigLog -Level Information -Message "Changed startup type for {ServiceName} to {AdditionalInfo}" -ServiceName $serviceName -AdditionalInfo $newStartType
                }
                default {
                    Write-Host "Invalid action." -ForegroundColor Red
                }
            }
        }
        catch {
            Write-Host "Failed to perform action: $_" -ForegroundColor Red
            Write-SystemConfigLog -Level Error -Message "Failed to manage service {ServiceName}: {AdditionalInfo}" -ServiceName $serviceName -AdditionalInfo $_
        }
    }
    else {
        Write-Host "Service not found." -ForegroundColor Red
    }
}

function Show-StartupPrograms {
    Write-SystemConfigLog -Level Information -Message "Displaying startup programs"
    
    $startupPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
    $commonStartupPath = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
    
    Write-Host "`nUser Startup Programs:" -ForegroundColor Cyan
    if (Test-Path $startupPath) {
        Get-ChildItem $startupPath | Format-Table -Property Name, LastWriteTime
    }
    
    Write-Host "`nCommon Startup Programs:" -ForegroundColor Cyan
    if (Test-Path $commonStartupPath) {
        Get-ChildItem $commonStartupPath | Format-Table -Property Name, LastWriteTime
    }
    
    $choice = Read-Host "`nDo you want to manage startup programs? (Y/N)"
    if ($choice -eq "Y") {
        Manage-StartupPrograms -StartupPath $startupPath -CommonStartupPath $commonStartupPath
    }
    
    Pause
    Show-SystemConfigMenu
}

function Manage-StartupPrograms {
    param(
        [Parameter(Mandatory=$true)]
        [string]$StartupPath,
        
        [Parameter(Mandatory=$true)]
        [string]$CommonStartupPath
    )
    
    $location = Read-Host "Manage programs in (1) User Startup or (2) Common Startup?"
    $path = if ($location -eq "1") { $StartupPath } else { $CommonStartupPath }
    
    Write-Host "`n1. Add program"
    Write-Host "2. Remove program"
    Write-Host "3. Back"
    
    $action = Read-Host "Select action"
    
    switch ($action) {
        "1" {
            $programPath = Read-Host "Enter full path to program"
            if (Test-Path $programPath) {
                $shortcutPath = Join-Path $path "$(Split-Path $programPath -Leaf).lnk"
                $WshShell = New-Object -ComObject WScript.Shell
                $Shortcut = $WshShell.CreateShortcut($shortcutPath)
                $Shortcut.TargetPath = $programPath
                $Shortcut.Save()
                Write-Host "Program added to startup." -ForegroundColor Green
                Write-SystemConfigLog -Level Information -Message "Added program to startup: {ProgramPath}" -ProgramPath $programPath
            }
            else {
                Write-Host "Program not found." -ForegroundColor Red
            }
        }
        "2" {
            $programs = Get-ChildItem $path
            for ($i = 0; $i -lt $programs.Count; $i++) {
                Write-Host "$($i + 1). $($programs[$i].Name)"
            }
            $choice = Read-Host "Select program to remove"
            if ($choice -match '^\d+$' -and $choice -le $programs.Count) {
                Remove-Item $programs[$choice - 1].FullName
                Write-Host "Program removed from startup." -ForegroundColor Green
                Write-SystemConfigLog -Level Information -Message "Removed program from startup: {ProgramPath}" -ProgramPath $programs[$choice - 1].Name
            }
        }
        "3" { return }
        default {
            Write-Host "Invalid action." -ForegroundColor Red
        }
    }
}

function Show-PowerSettings {
    Write-SystemConfigLog -Level Information -Message "Displaying power settings"
    
    Write-Host "`nCurrent Power Plan:" -ForegroundColor Cyan
    powercfg /getactivescheme
    
    Write-Host "`nAvailable Power Plans:" -ForegroundColor Cyan
    powercfg /list
    
    $choice = Read-Host "`nDo you want to change power settings? (Y/N)"
    if ($choice -eq "Y") {
        $scheme = Read-Host "Enter power scheme GUID (or 'balanced', 'high', 'power')"
        try {
            if ($scheme -eq "balanced") { $scheme = "381b4222-f694-41f0-9685-ff5bb260df2e" }
            elseif ($scheme -eq "high") { $scheme = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c" }
            elseif ($scheme -eq "power") { $scheme = "a1841308-3541-4fab-bc81-f71556f20b4a" }
            
            powercfg /setactive $scheme
            Write-Host "Power scheme changed successfully." -ForegroundColor Green
            Write-SystemConfigLog -Level Information -Message "Changed power scheme to {AdditionalInfo}" -AdditionalInfo $scheme
        }
        catch {
            Write-Host "Failed to change power scheme: $_" -ForegroundColor Red
            Write-SystemConfigLog -Level Error -Message "Failed to change power scheme: {AdditionalInfo}" -AdditionalInfo $_
        }
    }
    
    Pause
    Show-SystemConfigMenu
}

function Show-SystemProperties {
    Write-SystemConfigLog -Level Information -Message "Displaying system properties"
    
    Write-Host "`nSystem Properties:" -ForegroundColor Cyan
    $computerInfo = Get-ComputerInfo
    $computerInfo | Format-List -Property OsName, OsVersion, OsBuildNumber, OsArchitecture, ComputerName, Domain, TimeZone, LastBootUpTime
    
    Write-Host "`nEnvironment Variables:" -ForegroundColor Cyan
    Get-ChildItem env: | Format-Table -Property Name, Value
    
    Write-Host "`nSystem Path:" -ForegroundColor Cyan
    $env:Path -split ';' | Format-Table
    
    Pause
    Show-SystemConfigMenu
}

Export-ModuleMember -Function Show-SystemConfigMenu 