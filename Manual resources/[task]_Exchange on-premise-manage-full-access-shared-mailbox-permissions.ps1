# Fixed values
$AutoMapping = $false

try {
    <#----- Exchange On-Premises: Start -----#>
    # Connect to Exchange
    try {
        $adminSecurePassword = ConvertTo-SecureString -String "$ExchangeAdminPassword" -AsPlainText -Force
        $adminCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ExchangeAdminUsername, $adminSecurePassword
        $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
        $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -ErrorAction Stop 
        #-AllowRedirection
        $null = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber
        HID-Write-Status -Message "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" -Event Success
    }
    catch {
        HID-Write-Status -Message "Error connecting to Exchange using the URI [$exchangeConnectionUri]" -Event Error
        HID-Write-Status -Message "Error at line: $($_.InvocationInfo.ScriptLineNumber - 79): $($_.Exception.Message)" -Event Error
        if ($debug -eq $true) {
            HID-Write-Status -Message "$($_.Exception)" -Event Error
        }
        HID-Write-Summary -Message "Failed to connect to Exchange using the URI [$exchangeConnectionUri]" -Event Failed
        throw $_
    }

    Hid-Write-Status -Message "Checking if mailbox with identity '$($mailboxAlias)' exists" -Event Information
    $mailbox = Get-Mailbox -Identity $mailboxAlias -ErrorAction Stop
    if ($mailbox.Name.Count -eq 0) {
        throw "Could not find shared mailbox with identity '$($mailboxAlias)'"
    }
    # Add Full Access Permissions for Mail-enabled Security Group for users
    try { 
        # Add Full Access Permissions
        if ($usersToAddFullAccess -ne "[]") {
            HID-Write-Status -Message "Starting to add full access members to mailbox $($mailboxAlias)" -Event Information
            $usersToAddJson = $usersToAddFullAccess | ConvertFrom-Json
            foreach ($user in $usersToAddJson) {
                if ($AutoMapping) {
            
                    Add-MailboxPermission -Identity $mailboxAlias -AccessRights FullAccess -InheritanceType All -AutoMapping:$true -User $user.sAMAccountName -ErrorAction Stop
                }
        
                else {
                    Add-MailboxPermission -Identity $mailboxAlias -AccessRights FullAccess -InheritanceType All -AutoMapping:$false -User $user.sAMAccountName -ErrorAction Stop
                }
                Hid-Write-Status -Message "Assigned access rights [FullAccess] for mailbox [$($mailboxAlias)] to [$($user.sAMAccountName)] successfully" -Event Success
                HID-Write-Summary -Message "Assigned access rights [FullAccess] for mailbox [$($mailboxAlias)] to [$($user.sAMAccountName)] successfully" -Event Success
            }
        }
        
    }
    catch {
        HID-Write-Status -Message "Error assigning access rights [FullAccess] for mailbox [$($mailboxAlias)] to [$($user.sAMAccountName)]. Error: $($_.Exception.Message)" -Event Error
        HID-Write-Summary -Message "Error assigning access rights [FullAccess] for mailbox [$($mailboxAlias)] to [$($user.sAMAccountName)]" -Event Failed
        throw $_
    }
    
    # Remove Full Access Permissions for Mail-enabled Security Group for users
    try { 
        # Remove Full Access Permissions
        if ($usersToRemoveFullAccess -ne "[]") {
            HID-Write-Status -Message "Starting to remove full access members from mailbox $($mailboxAlias)" -Event Information
            $usersToRemoveJson = $usersToRemoveFullAccess | ConvertFrom-Json
            foreach ($user in $usersToRemoveJson) {
                
                Remove-MailboxPermission -Identity $mailboxAlias -AccessRights FullAccess -InheritanceType All -User $user.sAMAccountName -Confirm:$false -ErrorAction Stop
                Hid-Write-Status -Message "Removed access rights [FullAccess] for mailbox [$($mailboxAlias)] to [$($user.sAMAccountName)] successfully" -Event Success
                HID-Write-Summary -Message "Removed access rights [FullAccess] for mailbox [$($mailboxAlias)] to [$($user.sAMAccountName)] successfully" -Event Success   
            }
        }
        
    }
    catch {
        HID-Write-Status -Message "Error removing access rights [FullAccess] for mailbox [$($mailboxAlias)] to [$($user.sAMAccountName)]. Error: $($_.Exception.Message)" -Event Error
        HID-Write-Summary -Message "Error removing access rights [FullAccess] for mailbox [$($mailboxAlias)] to [$($user.sAMAccountName)]" -Event Failed
        throw $_
    }
}
catch {
    HID-Write-Status -Message "Error removing access rights for mailbox [$($mailboxAlias)] to the user [$($user.sAMAccountName)]. Error: $($_.Exception.Message)" -Event Error
    HID-Write-Summary -Message "Error removing access rights for mailbox [$($mailboxAlias)] to the user [$($user.sAMAccountName)]" -Event Failed
}
finally {
    # Disconnect from Exchange
    try {
        Remove-PsSession -Session $exchangeSession -Confirm:$false -ErrorAction Stop
        HID-Write-Status -Message "Successfully disconnected from Exchange" -Event Success
    }
    catch {
        HID-Write-Status -Message "Error disconnecting from Exchange" -Event Error
        HID-Write-Status -Message "Error at line: $($_.InvocationInfo.ScriptLineNumber - 79): $($_.Exception.Message)" -Event Error
        if ($debug -eq $true) {
            HID-Write-Status -Message "$($_.Exception)" -Event Error
        }
        HID-Write-Summary -Message "Failed to disconnect from Exchange" -Event Failed
        throw $_
    }
    <#----- Exchange On-Premises: End -----#>
}


