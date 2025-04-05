# Diagnostics.psm1 - System diagnostics and troubleshooting tools

function Show-DiagnosticsMenu {
    Clear-Host
    Write-Host "Diagnostics Menu:`n" -ForegroundColor White
    Write-Host "1. System File Check" -ForegroundColor Green
    Write-Host "2. Disk Health Check" -ForegroundColor Green
    Write-Host "3. Memory Diagnostics" -ForegroundColor Green
    Write-Host "4. Event Log Analysis" -ForegroundColor Green
    Write-Host "5. Performance Monitor" -ForegroundColor Green
    Write-Host "B. Back to Main Menu" -ForegroundColor Yellow
    Write-Host "X. Exit" -ForegroundColor Red
    Write-Host "`n"
    
    $choice = Read-Host "Select an option"
    
    switch ($choice) {
        "1" { Start-SystemFileCheck }
        "2" { Start-DiskHealthCheck }
        "3" { Start-MemoryDiagnostics }
        "4" { Show-EventLogAnalysis }
        "5" { Show-PerformanceMonitor }
        "B" { return }
        "X" { exit }
        default {
            Write-Host "Invalid option. Please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Show-DiagnosticsMenu
        }
    }
}

function Start-SystemFileCheck {
    Write-Log -Level Information -Message "Starting System File Check"
    
    Write-Host "`nRunning System File Checker (SFC)..." -ForegroundColor Cyan
    $result = sfc /scannow
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "No integrity violations found." -ForegroundColor Green
        Write-Log -Level Information -Message "SFC completed successfully with no violations"
    }
    else {
        Write-Host "Integrity violations found. Check the CBS.log for details." -ForegroundColor Yellow
        Write-Log -Level Warning -Message "SFC found integrity violations"
    }
    
    $choice = Read-Host "`nDo you want to run DISM health check? (Y/N)"
    if ($choice -eq "Y") {
        Write-Host "`nRunning DISM health check..." -ForegroundColor Cyan
        dism /online /cleanup-image /checkhealth
        dism /online /cleanup-image /scanhealth
    }
    
    Pause
    Show-DiagnosticsMenu
}

function Start-DiskHealthCheck {
    Write-Log -Level Information -Message "Starting Disk Health Check"
    
    Write-Host "`nAvailable Drives:" -ForegroundColor Cyan
    $drives = Get-PhysicalDisk | Where-Object { $_.OperationalStatus -eq "OK" }
    $drives | Format-Table -Property FriendlyName, HealthStatus, Size, MediaType
    
    $choice = Read-Host "`nDo you want to run CHKDSK? (Y/N)"
    if ($choice -eq "Y") {
        $driveLetter = Read-Host "Enter drive letter (e.g., C:)"
        Write-Host "`nRunning CHKDSK in read-only mode..." -ForegroundColor Cyan
        chkdsk $driveLetter
    }
    
    $choice = Read-Host "`nDo you want to check disk performance? (Y/N)"
    if ($choice -eq "Y") {
        Write-Host "`nRunning disk performance check..." -ForegroundColor Cyan
        Get-PhysicalDisk | Get-StorageReliabilityCounter | Format-Table -Property DeviceId, Temperature, Wear, PowerOnHours
    }
    
    Pause
    Show-DiagnosticsMenu
}

function Start-MemoryDiagnostics {
    Write-Log -Level Information -Message "Starting Memory Diagnostics"
    
    Write-Host "`nMemory Information:" -ForegroundColor Cyan
    $memory = Get-WmiObject Win32_PhysicalMemory
    $totalMemory = ($memory | Measure-Object -Property Capacity -Sum).Sum / 1GB
    Write-Host "Total Physical Memory: $totalMemory GB"
    
    $memory | Format-Table -Property DeviceLocator, Capacity, Speed, Manufacturer
    
    $choice = Read-Host "`nDo you want to run Windows Memory Diagnostic? (Y/N)"
    if ($choice -eq "Y") {
        Write-Host "`nScheduling Windows Memory Diagnostic for next restart..." -ForegroundColor Cyan
        mdsched.exe
    }
    
    Pause
    Show-DiagnosticsMenu
}

function Show-EventLogAnalysis {
    Write-Log -Level Information -Message "Starting Event Log Analysis"
    
    Write-Host "`nEvent Log Analysis:" -ForegroundColor Cyan
    Write-Host "1. System Events"
    Write-Host "2. Application Events"
    Write-Host "3. Security Events"
    Write-Host "4. Custom Filter"
    
    $choice = Read-Host "Select log type"
    
    $logName = switch ($choice) {
        "1" { "System" }
        "2" { "Application" }
        "3" { "Security" }
        "4" { Read-Host "Enter log name" }
        default { "System" }
    }
    
    $hours = Read-Host "Enter hours to look back (default: 24)"
    if (-not $hours) { $hours = 24 }
    
    $events = Get-WinEvent -LogName $logName -MaxEvents 50 -ErrorAction SilentlyContinue |
        Where-Object { $_.TimeCreated -gt (Get-Date).AddHours(-$hours) } |
        Select-Object TimeCreated, LevelDisplayName, ProviderName, Message
    
    if ($events) {
        $events | Format-Table -AutoSize
    }
    else {
        Write-Host "No events found in the specified time range." -ForegroundColor Yellow
    }
    
    Pause
    Show-DiagnosticsMenu
}

function Show-PerformanceMonitor {
    Write-Log -Level Information -Message "Starting Performance Monitor"
    
    Write-Host "`nPerformance Counters:" -ForegroundColor Cyan
    Write-Host "1. CPU Usage"
    Write-Host "2. Memory Usage"
    Write-Host "3. Disk I/O"
    Write-Host "4. Network Usage"
    
    $choice = Read-Host "Select counter to monitor"
    
    $counter = switch ($choice) {
        "1" { "\Processor(_Total)\% Processor Time" }
        "2" { "\Memory\% Committed Bytes In Use" }
        "3" { "\PhysicalDisk(_Total)\% Disk Time" }
        "4" { "\Network Interface(*)\Bytes Total/sec" }
        default { "\Processor(_Total)\% Processor Time" }
    }
    
    $duration = Read-Host "Enter duration in seconds (default: 10)"
    if (-not $duration) { $duration = 10 }
    
    Write-Host "`nMonitoring $counter for $duration seconds..." -ForegroundColor Cyan
    Get-Counter -Counter $counter -SampleInterval 1 -MaxSamples $duration | 
        ForEach-Object { 
            $timestamp = $_.Timestamp
            $value = $_.CounterSamples.CookedValue
            Write-Host "$timestamp : $value"
        }
    
    Pause
    Show-DiagnosticsMenu
}

Export-ModuleMember -Function Show-DiagnosticsMenu 