$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# variables configured in form
$mailboxAlias = $form.gridMailbox.Alias
$usersToAddFullAccess = $form.fullaccessList.leftToRight
$usersToRemoveFullAccess = $form.fullaccessList.rightToLeft
$AutoMapping = $form.blnautomapping

try {
    <#----- Exchange On-Premises: Start -----#>
    # Connect to Exchange
    try {
        $adminSecurePassword = ConvertTo-SecureString -String "$ExchangeAdminPassword" -AsPlainText -Force
        $adminCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ExchangeAdminUsername, $adminSecurePassword
        $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
        $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -ErrorAction Stop 
        #-AllowRedirection
        $session = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber
    
        Write-Information "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" 
    
        $Log = @{
            Action            = "UpdateResource" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" # required (free format text) 
            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log
    }
    catch {
        Write-Error "Error connecting to Exchange using the URI [$exchangeConnectionUri]. Error: $($_.Exception.Message)"
        $Log = @{
            Action            = "UpdateResource" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Failed to connect to Exchange using the URI [$exchangeConnectionUri]." # required (free format text) 
            IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log
    }

    Write-Information "Checking if mailbox with identity '$($mailboxAlias)' exists"
    $mailbox = Get-Mailbox -Identity $mailboxAlias -ErrorAction Stop | Select-Object Guid
    if ($mailbox.Guid.Count -eq 0) {
        throw "Could not find shared mailbox with identity '$($mailboxAlias)'"
    }
    # Add Full Access Permissions for Mail-enabled Security Group for users
    try { 
        # Add Full Access Permissions
        if ($usersToAddFullAccess.Count -gt 0) {
            Write-Information "Starting to add full access members to mailbox $($mailboxAlias)"
            
            foreach ($user in $usersToAddFullAccess) {
                try {            
                    if($AutoMapping -eq 'true') {
                        Add-MailboxPermission -Identity $mailboxAlias -AccessRights FullAccess -InheritanceType All -AutoMapping:$true -User $($user.sAMAccountName) -ErrorAction Stop
                    }
                    else
                    {
                        Add-MailboxPermission -Identity $mailboxAlias -AccessRights FullAccess -InheritanceType All -AutoMapping:$false -User $($user.sAMAccountName) -ErrorAction Stop
                    }
                    
                    Write-Information "Assigned access rights [FullAccess] on mailbox [$($mailboxAlias)] for [$($user.sAMAccountName)] successfully."
                    $Log = @{
                        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
                        System            = "Exchange On-Premise" # optional (free format text) 
                        Message           = "Assigned access rights [FullAccess] on mailbox [$($mailboxAlias)] for [$($user.sAMAccountName)] successfully." # required (free format text) 
                        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                        TargetDisplayName = $mailboxAlias # optional (free format text) 
                        TargetIdentifier  = $([string]$mailbox.Guid) # optional (free format text) 
                    }
                    #send result back  
                    Write-Information -Tags "Audit" -MessageData $log

                }
                catch {
                    Write-Error "Error assigning access rights [FullAccess] for $($user.sAMAccountName) on mailbox [$($mailboxAlias)]. Error: $($_.Exception.Message)" 
                    $Log = @{
                        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
                        System            = "Exchange On-Premise" # optional (free format text) 
                        Message           = "Error assigning access rights [FullAccess] to $($user.sAMAccountName) on mailbox [$($mailboxAlias)]" # required (free format text) 
                        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                        TargetDisplayName = $mailboxAlias # optional (free format text) 
                        TargetIdentifier  = $([string]$mailbox.Guid) # optional (free format text) 
                    }
                    #send result back  
                    Write-Information -Tags "Audit" -MessageData $log
                }
            }
        }
        
    }
    catch {
        Write-Error "Error assigning access rights [FullAccess] for $($usersToAddFullAccess.sAMAccountName) on mailbox [$($mailboxAlias)]. Error: $($_.Exception.Message)" 
        $Log = @{
            Action            = "UpdateResource" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Error assigning access rights [FullAccess] to $($usersToAddFullAccess.sAMAccountName) on mailbox [$($mailboxAlias)]" # required (free format text) 
            IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $mailboxAlias # optional (free format text) 
            TargetIdentifier  = $([string]$mailbox.Guid) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log
    }
    
    # Remove Full Access Permissions for Mail-enabled Security Group for users
    try { 
        # Remove Full Access Permissions
        if ($usersToRemoveFullAccess.Count -gt 0) {
            Write-Information "Starting to remove full access members from mailbox $($mailboxAlias)"
            
            foreach ($user in $usersToRemoveFullAccess) {
                try {
                    Remove-MailboxPermission -Identity $mailboxAlias -AccessRights FullAccess -InheritanceType All -User $($user.sAMAccountName) -Confirm:$false -ErrorAction Stop
                    Write-Information "Removing access rights [FullAccess] on mailbox [$($mailboxAlias)] for [$($user.sAMAccountName)] successfully"
                    $Log = @{
                        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
                        System            = "Exchange On-Premise" # optional (free format text) 
                        Message           = "Removing access rights [FullAccess] on mailbox [$($mailboxAlias)] for [$($user.sAMAccountName)] successfully." # required (free format text) 
                        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                        TargetDisplayName = $mailboxAlias # optional (free format text) 
                        TargetIdentifier  = $([string]$mailbox.Guid) # optional (free format text) 
                    }
                    #send result back  
                    Write-Information -Tags "Audit" -MessageData $log
                
                }
                catch {
                    Write-Error "Error removing access rights [FullAccess] on mailbox [$($mailboxAlias)] for [$($user.sAMAccountName)]. Error: $($_.Exception.Message)"
                    $Log = @{
                        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
                        System            = "Exchange On-Premise" # optional (free format text) 
                        Message           = "Error removing access rights [FullAccess] on mailbox [$($mailboxAlias)] for [$($user.sAMAccountName)]." # required (free format text) 
                        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                        TargetDisplayName = $mailboxAlias # optional (free format text) 
                        TargetIdentifier  = $([string]$mailbox.Guid) # optional (free format text) 
                    }
                    #send result back  
                    Write-Information -Tags "Audit" -MessageData $log
                }
            }
        }
        
    }
    catch {
        Write-Error "Error removing access rights [FullAccess] on mailbox [$($mailboxAlias)] for [$($usersToRemoveFullAccess.sAMAccountName)]."
        $Log = @{
            Action            = "UpdateResource" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Error removing access rights [FullAccess] on mailbox [$($mailboxAlias)] for [$($usersToRemoveFullAccess.sAMAccountName)]." # required (free format text) 
            IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $mailboxAlias # optional (free format text) 
            TargetIdentifier  = $([string]$mailbox.Guid) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log
    }
}
catch {
    Write-Error "Error setting access rights [FullAccess] on mailbox [$($mailboxAlias)]. Error: $($_.Exception.Message)"
    $Log = @{
        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Error setting access rights [FullAccess] on mailbox [$($mailboxAlias)]." # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
        TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log 
}
finally {
    # Disconnect from Exchange
    try {
        Remove-PsSession -Session $exchangeSession -Confirm:$false -ErrorAction Stop
        Write-Information "Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]"     
        $Log = @{
            Action            = "UpdateResource" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]" # required (free format text) 
            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log
    }
    catch {
        Write-Error "Error disconnecting from Exchange.  Error: $($_.Exception.Message)"
        $Log = @{
            Action            = "UpdateResource" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Failed to disconnect from Exchange using the URI [$exchangeConnectionUri]." # required (free format text) 
            IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log 
    }
}


