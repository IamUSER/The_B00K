# UserManagement.psm1 - User account management and security tools

function Show-UserManagementMenu {
    Clear-Host
    Write-Host "User Management Menu:`n" -ForegroundColor White
    Write-Host "1. User Accounts" -ForegroundColor Green
    Write-Host "2. Local Groups" -ForegroundColor Green
    Write-Host "3. User Permissions" -ForegroundColor Green
    Write-Host "4. Password Policies" -ForegroundColor Green
    Write-Host "5. Security Settings" -ForegroundColor Green
    Write-Host "B. Back to Main Menu" -ForegroundColor Yellow
    Write-Host "X. Exit" -ForegroundColor Red
    Write-Host "`n"
    
    $choice = Read-Host "Select an option"
    
    switch ($choice) {
        "1" { Show-UserAccounts }
        "2" { Show-LocalGroups }
        "3" { Show-UserPermissions }
        "4" { Show-PasswordPolicies }
        "5" { Show-SecuritySettings }
        "B" { return }
        "X" { exit }
        default {
            Write-Host "Invalid option. Please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Show-UserManagementMenu
        }
    }
}

function Show-UserAccounts {
    Write-UserManagementLog -Level Information -Message "Displaying user accounts"
    
    Write-Host "`nLocal User Accounts:" -ForegroundColor Cyan
    Get-LocalUser | Format-Table -Property Name, Enabled, LastLogon, PasswordLastSet
    
    Write-Host "`nActions:" -ForegroundColor Yellow
    Write-Host "1. Create new user"
    Write-Host "2. Modify user"
    Write-Host "3. Delete user"
    Write-Host "4. Back"
    
    $choice = Read-Host "Select action"
    
    switch ($choice) {
        "1" { New-LocalUserAccount }
        "2" { Modify-LocalUserAccount }
        "3" { Remove-LocalUserAccount }
        "4" { return }
        default {
            Write-Host "Invalid option." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
    
    Pause
    Show-UserManagementMenu
}

function New-LocalUserAccount {
    $username = Read-Host "Enter username"
    $fullName = Read-Host "Enter full name"
    $description = Read-Host "Enter description"
    $password = Read-Host "Enter password" -AsSecureString
    
    try {
        New-LocalUser -Name $username -FullName $fullName -Description $description -Password $password -PasswordNeverExpires $false -UserMayNotChangePassword $false
        Write-Host "User account created successfully." -ForegroundColor Green
        Write-UserManagementLog -Level Information -Message "Created new user account: {Username}" -Username $username
        
        $choice = Read-Host "Add user to any groups? (Y/N)"
        if ($choice -eq "Y") {
            Add-UserToGroups -Username $username
        }
    }
    catch {
        Write-Host "Failed to create user account: $_" -ForegroundColor Red
        Write-UserManagementLog -Level Error -Message "Failed to create user account: {AdditionalInfo}" -AdditionalInfo $_
    }
}

function Modify-LocalUserAccount {
    $username = Read-Host "Enter username to modify"
    $user = Get-LocalUser -Name $username -ErrorAction SilentlyContinue
    
    if ($user) {
        Write-Host "`nCurrent user information:" -ForegroundColor Cyan
        $user | Format-List
        
        Write-Host "`nModification options:" -ForegroundColor Yellow
        Write-Host "1. Change password"
        Write-Host "2. Enable/Disable account"
        Write-Host "3. Change description"
        Write-Host "4. Manage groups"
        Write-Host "5. Back"
        
        $choice = Read-Host "Select action"
        
        switch ($choice) {
            "1" {
                $password = Read-Host "Enter new password" -AsSecureString
                $user | Set-LocalUser -Password $password
                Write-Host "Password changed successfully." -ForegroundColor Green
                Write-UserManagementLog -Level Information -Message "Changed password for user {Username}" -Username $username
            }
            "2" {
                $enabled = !$user.Enabled
                $user | Set-LocalUser -Enabled $enabled
                Write-Host "Account $(if ($enabled) { 'enabled' } else { 'disabled' }) successfully." -ForegroundColor Green
                Write-UserManagementLog -Level Information -Message "$(if ($enabled) { 'Enabled' } else { 'Disabled' }) user account: {Username}" -Username $username
            }
            "3" {
                $description = Read-Host "Enter new description"
                $user | Set-LocalUser -Description $description
                Write-Host "Description updated successfully." -ForegroundColor Green
                Write-UserManagementLog -Level Information -Message "Updated description for user: {Username}" -Username $username
            }
            "4" {
                Add-UserToGroups -Username $username
            }
            "5" { return }
            default {
                Write-Host "Invalid option." -ForegroundColor Red
            }
        }
    }
    else {
        Write-Host "User not found." -ForegroundColor Red
    }
}

function Remove-LocalUserAccount {
    $username = Read-Host "Enter username to delete"
    $user = Get-LocalUser -Name $username -ErrorAction SilentlyContinue
    
    if ($user) {
        $confirm = Read-Host "Are you sure you want to delete user '$username'? (Y/N)"
        if ($confirm -eq "Y") {
            try {
                Remove-LocalUser -Name $username
                Write-Host "User account deleted successfully." -ForegroundColor Green
                Write-UserManagementLog -Level Information -Message "Deleted user account: {Username}" -Username $username
            }
            catch {
                Write-Host "Failed to delete user account: $_" -ForegroundColor Red
                Write-UserManagementLog -Level Error -Message "Failed to delete user account: {AdditionalInfo}" -AdditionalInfo $_
            }
        }
    }
    else {
        Write-Host "User not found." -ForegroundColor Red
    }
}

function Add-UserToGroups {
    param (
        [string]$Username
    )
    
    Write-Host "`nAvailable Groups:" -ForegroundColor Cyan
    Get-LocalGroup | Format-Table -Property Name, Description
    
    $groupName = Read-Host "Enter group name to add user to (or 'done' to finish)"
    while ($groupName -ne "done") {
        try {
            Add-LocalGroupMember -Group $groupName -Member $Username
            Write-Host "User added to group successfully." -ForegroundColor Green
            Write-UserManagementLog -Level Information -Message "Added user {Username} to group {GroupName}" -Username $Username -GroupName $groupName
        }
        catch {
            Write-Host "Failed to add user to group: $_" -ForegroundColor Red
            Write-UserManagementLog -Level Error -Message "Failed to add user {Username} to group {GroupName}: {AdditionalInfo}" -Username $Username -GroupName $groupName -AdditionalInfo $_
        }
        $groupName = Read-Host "Enter group name to add user to (or 'done' to finish)"
    }
}

function Show-LocalGroups {
    Write-UserManagementLog -Level Information -Message "Displaying local groups"
    
    Write-Host "`nLocal Groups:" -ForegroundColor Cyan
    Get-LocalGroup | Format-Table -Property Name, Description
    
    Write-Host "`nActions:" -ForegroundColor Yellow
    Write-Host "1. Create new group"
    Write-Host "2. Delete group"
    Write-Host "3. View group members"
    Write-Host "4. Back"
    
    $choice = Read-Host "Select action"
    
    switch ($choice) {
        "1" { New-LocalGroup }
        "2" { Remove-LocalGroup }
        "3" { View-GroupMembers }
        "4" { return }
        default {
            Write-Host "Invalid option." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
    
    Pause
    Show-UserManagementMenu
}

function New-LocalGroup {
    $groupName = Read-Host "Enter group name"
    $description = Read-Host "Enter description"
    
    try {
        New-LocalGroup -Name $groupName -Description $description
        Write-Host "Group created successfully." -ForegroundColor Green
        Write-UserManagementLog -Level Information -Message "Created new group: {GroupName}" -GroupName $groupName
    }
    catch {
        Write-Host "Failed to create group: $_" -ForegroundColor Red
        Write-UserManagementLog -Level Error -Message "Failed to create group: {AdditionalInfo}" -AdditionalInfo $_
    }
}

function Remove-LocalGroup {
    $groupName = Read-Host "Enter group name to delete"
    $group = Get-LocalGroup -Name $groupName -ErrorAction SilentlyContinue
    
    if ($group) {
        $confirm = Read-Host "Are you sure you want to delete group '$groupName'? (Y/N)"
        if ($confirm -eq "Y") {
            try {
                Remove-LocalGroup -Name $groupName
                Write-Host "Group deleted successfully." -ForegroundColor Green
                Write-UserManagementLog -Level Information -Message "Deleted group: {GroupName}" -GroupName $groupName
            }
            catch {
                Write-Host "Failed to delete group: $_" -ForegroundColor Red
                Write-UserManagementLog -Level Error -Message "Failed to delete group: {GroupName}" -GroupName $groupName
            }
        }
    }
    else {
        Write-Host "Group not found." -ForegroundColor Red
    }
}

function View-GroupMembers {
    $groupName = Read-Host "Enter group name"
    $group = Get-LocalGroup -Name $groupName -ErrorAction SilentlyContinue
    
    if ($group) {
        Write-Host ("`nMembers of {0}:" -f $groupName) -ForegroundColor Cyan
        Get-LocalGroupMember -Group $groupName | Format-Table -Property Name, PrincipalSource
    }
    else {
        Write-Host "Group not found." -ForegroundColor Red
    }
}

function Show-UserPermissions {
    Write-UserManagementLog -Level Information -Message "Displaying user permissions"
    
    Write-Host "`nUser Rights Assignment:" -ForegroundColor Cyan
    $rights = Get-WmiObject -Class Win32_UserRight
    $rights | Format-Table -Property Name, Description
    
    Pause
    Show-UserManagementMenu
}

function Show-PasswordPolicies {
    Write-UserManagementLog -Level Information -Message "Displaying password policies"
    
    Write-Host "`nPassword Policy Settings:" -ForegroundColor Cyan
    $policy = Get-LocalUser | Select-Object -First 1 | Get-LocalUserPasswordPolicy
    $policy | Format-List
    
    Pause
    Show-UserManagementMenu
}

function Show-SecuritySettings {
    Write-UserManagementLog -Level Information -Message "Displaying security settings"
    
    Write-Host "`nSecurity Settings:" -ForegroundColor Cyan
    Write-Host "1. Account Lockout Policy"
    Write-Host "2. Audit Policy"
    Write-Host "3. User Rights Assignment"
    Write-Host "4. Security Options"
    
    $choice = Read-Host "Select setting to view"
    
    switch ($choice) {
        "1" {
            Write-Host "`nAccount Lockout Policy:" -ForegroundColor Cyan
            $lockout = Get-WmiObject -Class Win32_AccountLockoutPolicy
            $lockout | Format-List
        }
        "2" {
            Write-Host "`nAudit Policy:" -ForegroundColor Cyan
            $audit = Get-WmiObject -Class Win32_AuditPolicy
            $audit | Format-Table -Property Category, Subcategory, Setting
        }
        "3" {
            Show-UserPermissions
        }
        "4" {
            Write-Host "`nSecurity Options:" -ForegroundColor Cyan
            $options = Get-WmiObject -Class Win32_SecuritySetting
            $options | Format-Table -Property Name, Value
        }
        default {
            Write-Host "Invalid option." -ForegroundColor Red
        }
    }
    
    Pause
    Show-UserManagementMenu
}

function Write-UserManagementLog {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet("Error", "Warning", "Information", "Debug")]
        [string]$Level,
        
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [string]$Username,
        
        [Parameter(Mandatory=$false)]
        [string]$GroupName,
        
        [Parameter(Mandatory=$false)]
        [string]$AdditionalInfo
    )
    
    $logMessage = $Message
    if ($Username) { $logMessage = $logMessage -replace "{Username}", $Username }
    if ($GroupName) { $logMessage = $logMessage -replace "{GroupName}", $GroupName }
    if ($AdditionalInfo) { $logMessage = $logMessage -replace "{AdditionalInfo}", $AdditionalInfo }
    
    Write-Log -Level $Level -Message $logMessage
}

Export-ModuleMember -Function Show-UserManagementMenu 