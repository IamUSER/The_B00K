# NetworkTools.psm1 - Network management and diagnostic tools

function Show-NetworkMenu {
    Clear-Host
    Write-Host "Network Tools Menu:`n" -ForegroundColor White
    Write-Host "1. Network Adapter Information" -ForegroundColor Green
    Write-Host "2. IP Configuration" -ForegroundColor Green
    Write-Host "3. DNS Configuration" -ForegroundColor Green
    Write-Host "4. Network Diagnostics" -ForegroundColor Green
    Write-Host "5. NTP Configuration" -ForegroundColor Green
    Write-Host "B. Back to Main Menu" -ForegroundColor Yellow
    Write-Host "X. Exit" -ForegroundColor Red
    Write-Host "`n"
    
    $choice = Read-Host "Select an option"
    
    switch ($choice) {
        "1" { Get-NetworkAdapterInfo }
        "2" { Get-IPConfiguration }
        "3" { Get-DNSConfiguration }
        "4" { Start-NetworkDiagnostics }
        "5" { Set-NTPConfiguration }
        "B" { return }
        "X" { exit }
        default {
            Write-Host "Invalid option. Please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Show-NetworkMenu
        }
    }
}

function Get-NetworkAdapterInfo {
    Write-Log -Level Information -Message "Retrieving network adapter information"
    
    $adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
    foreach ($adapter in $adapters) {
        Write-Host "`nAdapter: $($adapter.Name)" -ForegroundColor Cyan
        Write-Host "Status: $($adapter.Status)"
        Write-Host "Speed: $($adapter.Speed) Mbps"
        Write-Host "MAC Address: $($adapter.MacAddress)"
        
        $ipConfig = Get-NetIPAddress -InterfaceIndex $adapter.ifIndex | Where-Object { $_.AddressFamily -eq "IPv4" }
        if ($ipConfig) {
            Write-Host "IP Address: $($ipConfig.IPAddress)"
            Write-Host "Subnet Mask: $($ipConfig.PrefixLength)"
        }
    }
    
    Pause
    Show-NetworkMenu
}

function Get-IPConfiguration {
    Write-Log -Level Information -Message "Retrieving IP configuration"
    
    $ipConfig = Get-NetIPConfiguration
    foreach ($config in $ipConfig) {
        Write-Host "`nInterface: $($config.InterfaceAlias)" -ForegroundColor Cyan
        Write-Host "IPv4 Address: $($config.IPv4Address.IPAddress)"
        Write-Host "Subnet Mask: $($config.IPv4Address.PrefixLength)"
        Write-Host "Default Gateway: $($config.IPv4DefaultGateway.NextHop)"
        Write-Host "DNS Servers: $($config.DNSServer.ServerAddresses -join ', ')"
    }
    
    Pause
    Show-NetworkMenu
}

function Get-DNSConfiguration {
    Write-Log -Level Information -Message "Retrieving DNS configuration"
    
    $dnsConfig = Get-DnsClientServerAddress
    foreach ($config in $dnsConfig) {
        Write-Host "`nInterface: $($config.InterfaceAlias)" -ForegroundColor Cyan
        Write-Host "DNS Servers: $($config.ServerAddresses -join ', ')"
    }
    
    $choice = Read-Host "`nDo you want to change DNS servers? (Y/N)"
    if ($choice -eq "Y") {
        Set-DNSServers
    }
    
    Show-NetworkMenu
}

function Set-DNSServers {
    $adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
    $adapterNames = $adapters | Select-Object -ExpandProperty Name
    
    Write-Host "`nAvailable network adapters:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $adapterNames.Count; $i++) {
        Write-Host "$($i + 1). $($adapterNames[$i])"
    }
    
    $choice = Read-Host "`nSelect adapter number"
    if ($choice -match '^\d+$' -and $choice -le $adapterNames.Count) {
        $adapter = $adapters[$choice - 1]
        
        Write-Host "`nAvailable DNS server options:" -ForegroundColor Cyan
        Write-Host "1. Google DNS (8.8.8.8, 8.8.4.4)"
        Write-Host "2. Cloudflare DNS (1.1.1.1, 1.0.0.1)"
        Write-Host "3. Custom DNS servers"
        
        $dnsChoice = Read-Host "Select DNS option"
        switch ($dnsChoice) {
            "1" { $dnsServers = @("8.8.8.8", "8.8.4.4") }
            "2" { $dnsServers = @("1.1.1.1", "1.0.0.1") }
            "3" {
                $dns1 = Read-Host "Enter primary DNS server"
                $dns2 = Read-Host "Enter secondary DNS server"
                $dnsServers = @($dns1, $dns2)
            }
            default {
                Write-Host "Invalid option. No changes made." -ForegroundColor Red
                return
            }
        }
        
        try {
            Set-DnsClientServerAddress -InterfaceIndex $adapter.ifIndex -ServerAddresses $dnsServers
            Write-Host "DNS servers updated successfully." -ForegroundColor Green
            Write-Log -Level Information -Message "DNS servers updated for $($adapter.Name) to $($dnsServers -join ', ')"
        }
        catch {
            Write-Host "Failed to update DNS servers: $_" -ForegroundColor Red
            Write-Log -Level Error -Message "Failed to update DNS servers: $_"
        }
    }
    else {
        Write-Host "Invalid selection." -ForegroundColor Red
    }
}

function Start-NetworkDiagnostics {
    Write-Log -Level Information -Message "Starting network diagnostics"
    
    Write-Host "`nRunning network diagnostics..." -ForegroundColor Cyan
    
    # Test network connectivity
    Write-Host "`nTesting network connectivity..."
    Test-Connection -ComputerName "8.8.8.8" -Count 4
    
    # Test DNS resolution
    Write-Host "`nTesting DNS resolution..."
    Resolve-DnsName -Name "google.com" -Type A
    
    # Check network routes
    Write-Host "`nChecking network routes..."
    Get-NetRoute | Where-Object { $_.DestinationPrefix -ne "0.0.0.0/0" } | Format-Table
    
    Pause
    Show-NetworkMenu
}

function Set-NTPConfiguration {
    Write-Log -Level Information -Message "Configuring NTP settings"
    
    Write-Host "`nCurrent NTP configuration:" -ForegroundColor Cyan
    w32tm /query /configuration
    
    $choice = Read-Host "`nDo you want to change NTP server? (Y/N)"
    if ($choice -eq "Y") {
        Write-Host "`nAvailable NTP servers:" -ForegroundColor Cyan
        Write-Host "1. time.windows.com"
        Write-Host "2. time.nist.gov"
        Write-Host "3. pool.ntp.org"
        Write-Host "4. Custom NTP server"
        
        $ntpChoice = Read-Host "Select NTP server option"
        switch ($ntpChoice) {
            "1" { $ntpServer = "time.windows.com" }
            "2" { $ntpServer = "time.nist.gov" }
            "3" { $ntpServer = "pool.ntp.org" }
            "4" { $ntpServer = Read-Host "Enter NTP server address" }
            default {
                Write-Host "Invalid option. No changes made." -ForegroundColor Red
                return
            }
        }
        
        try {
            w32tm /config /manualpeerlist:$ntpServer /syncfromflags:manual /update
            w32tm /resync
            Write-Host "NTP configuration updated successfully." -ForegroundColor Green
            Write-Log -Level Information -Message "NTP server updated to $ntpServer"
        }
        catch {
            Write-Host "Failed to update NTP configuration: $_" -ForegroundColor Red
            Write-Log -Level Error -Message "Failed to update NTP configuration: $_"
        }
    }
    
    Pause
    Show-NetworkMenu
}

Export-ModuleMember -Function Show-NetworkMenu 