# variables configured in form
$mailbox = $form.gridMailbox
$usersToAdd = $form.fullaccessList.leftToRight
$usersToRemove = $form.fullaccessList.rightToLeft
$AutoMapping = $form.blnautomapping

# Global variables
# Outcommented as these are set from Global Variables
# $ExchangeConnectionUri = ""
# $ExchangeAdminUsername = ""
# $ExchangeAdminPassword = ""

# Fixed values
$commands = @(
    "Get-Mailbox",    
    "Add-MailboxPermission",
    "Remove-MailboxPermission"
)

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Set debug logging
$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

try {
     # Create credentials
    $actionMessage = "creating credentials object"
    
    $securePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $credential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername, $securePassword)
    
    Write-Verbose "Created credentials for user [$ExchangeAdminUsername]"

    # Connect to Exchange On-Premises
    # Docs: https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-servers-using-remote-powershell
    $actionMessage = "connecting to Exchange On-Premises"

    $sessionOptionParams = @{
        SkipCACheck         = $false
        SkipCNCheck         = $false
        SkipRevocationCheck = $false
    }

    $sessionOption = New-PSSessionOption @sessionOptionParams

    $sessionParams = @{
        Authentication    = 'Default'
        ConfigurationName = 'Microsoft.Exchange'
        Credential        = $credential
        ConnectionUri     = $ExchangeConnectionUri
        SessionOption     = $sessionOption
        ErrorAction       = "Stop"
    }

    $exchangeSession = New-PSSession @sessionParams
    $null = Import-PSSession -Session $exchangeSession -DisableNameChecking -AllowClobber -CommandName $commands -ErrorAction Stop

    # Send initial audit log
    $Log = @{
        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premises" # optional (free format text) 
        Message           = "Successfully connected to Exchange using URI [$ExchangeConnectionUri]" # required (free format text) 
        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $ExchangeConnectionUri # optional (free format text) 
        TargetIdentifier  = $([string]$exchangeSession.InstanceId) # optional (free format text) 
    }
    Write-Information -Tags "Audit" -MessageData $log

    # Add Full Access Permissions
    if ($usersToAdd.Count -gt 0) {
        Write-Information "Starting to grant permission [FullAccess] members to mailbox $($mailbox.DisplayName)"
        
        foreach ($user in $usersToAdd) {
            try {      
                $actionMessage = "granting permission [FullAccess] to mailbox [$($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress))] for user [$($user.userPrincipalName) ($($user.Guid))]"      
                
                $FullAccessPermissionSplatParams = @{
                    Identity      = $($mailbox.Guid)
                    User          = $($user.sAMAccountName)
                    AccessRights  = "FullAccess"
                    InheritanceType = "All"
                    AutoMapping   = [bool]$AutoMapping
                    ErrorAction   = "Stop"                 
                }
                $addFullAccessPermission = Add-MailboxPermission @FullAccessPermissionSplatParams
                
                Write-Information "Granting access rights [FullAccess] on mailbox [$($mailbox.DisplayName)] for [$($user.DisplayName)] successfully."
                $Log = @{
                    Action            = "GrantMembership" # optional. ENUM (undefined = default) 
                    System            = "Exchange On-Premises" # optional (free format text) 
                    Message           = "Granting access rights [FullAccess] on mailbox [$($mailbox.DisplayName)] for [$($user.DisplayName)] successfully." # required (free format text) 
                    IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                    TargetDisplayName = $($mailbox.DisplayName) # optional (free format text) 
                    TargetIdentifier  = $([string]$mailbox.Guid) # optional (free format text) 
                }
                #send result back  
                Write-Information -Tags "Audit" -MessageData $log

            }
            catch {
                $ex = $PSItem
                if (-not [string]::IsNullOrEmpty($ex.Exception.Message)) {
                    $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
                    $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
                }
                else {
                    $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception)"
                    $auditMessage = "Error $($actionMessage). Error: $($ex.Exception)"
                }

                Write-Error "Error granting access rights [FullAccess] for [$($user.DisplayName)] on mailbox [$($mailbox.DisplayName)]. Error: $($_.Exception.Message)" 
                $Log = @{
                    Action            = "GrantMembership" # optional. ENUM (undefined = default) 
                    System            = "Exchange On-Premises" # optional (free format text) 
                    Message           = "Error assigning access rights [FullAccess] to [$($user.DisplayName)] on mailbox [$($mailbox.DisplayName)]" # required (free format text) 
                    IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                    TargetDisplayName = $($mailbox.DisplayName) # optional (free format text) 
                    TargetIdentifier  = $([string]$mailbox.Guid) # optional (free format text) 
                }
                #send result back  
                Write-Information -Tags "Audit" -MessageData $log
                Write-Warning $warningMessage
                Write-Error $auditMessage            
            }
            break
        }
    }

    # Remove Full Access Permissions
    if ($usersToRemove.Count -gt 0) {
        Write-Information "Starting to revoke permission [FullAccess] on mailbox [$($mailbox.DisplayName)]"
        
        foreach ($user in $usersToRemove) {
            try {
                $actionMessage = "revoking permission [FullAccess] to mailbox [$($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress))] for user [$($user.userPrincipalName) ($($user.Guid))]"      
                $FullAccessPermissionSplatParams = @{
                    Identity      = $($mailbox.Guid)
                    User          = $($user.sAMAccountName)
                    AccessRights  = "FullAccess"
                    InheritanceType = "All"                    
                    ErrorAction   = "Stop"
                    Confirm      = $false
                }
                $removeFullAccessPermission = Remove-MailboxPermission @FullAccessPermissionSplatParams
                
                Write-Information "Revoking access rights [FullAccess] on mailbox [$($mailbox.DisplayName)] for [$($user.DisplayName)] successfully"
                $Log = @{
                    Action            = "RevokeMembership" # optional. ENUM (undefined = default) 
                    System            = "Exchange On-Premises" # optional (free format text) 
                    Message           = "Revoking access rights [FullAccess] on mailbox [$($mailbox.DisplayName)] for [$($user.DisplayName)] successfully." # required (free format text) 
                    IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                    TargetDisplayName = $($mailbox.DisplayName) # optional (free format text) 
                    TargetIdentifier  = $([string]$mailbox.Guid) # optional (free format text) 
                }                
            }
            catch {
                $ex = $PSItem
                if (-not [string]::IsNullOrEmpty($ex.Exception.Message)) {
                    $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
                    $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
                }
                else {
                    $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception)"
                    $auditMessage = "Error $($actionMessage). Error: $($ex.Exception)"
                }

                Write-Error "Error revoking access rights [FullAccess] on mailbox [$($mailbox.DisplayName)] for [$($user.DisplayName)]. Error: $($_.Exception.Message)"
                $Log = @{
                    Action            = "RevokeMembership" # optional. ENUM (undefined = default) 
                    System            = "Exchange On-Premises" # optional (free format text) 
                    Message           = "Error revoking access rights [FullAccess] on mailbox [$($mailbox.DisplayName)] for [$($user.DisplayName)]." # required (free format text) 
                    IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                    TargetDisplayName = $($mailbox.DisplayName) # optional (free format text) 
                    TargetIdentifier  = $([string]$mailbox.Guid) # optional (free format text) 
                }
                #send result back  
                Write-Information -Tags "Audit" -MessageData $log
                Write-Warning $warningMessage
                Write-Error $auditMessage
            }
            break
        }
    }
} catch {
    $ex = $PSItem
    if (-not [string]::IsNullOrEmpty($ex.Exception.Message)) {
        $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
    }
    else {
        $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception)"
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception)"
    }

    # Send error audit log to HelloID
    $Log = @{
        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premises" # optional (free format text) 
        Message           = $auditMessage # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $($mailbox.DisplayName) # optional (free format text) 
        TargetIdentifier  = $([string]$mailbox.Guid) # optional (free format text) 
    }
    
    Write-Information -Tags "Audit" -MessageData $log
    Write-Warning $warningMessage
    Write-Error $auditMessage
}
finally {
    # Disconnect from Exchange
    # Docs: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/remove-pssession
    if ($null -ne $exchangeSession) {
        try {
            $deleteExchangeSessionSplatParams = @{
                Session     = $exchangeSession
                Confirm     = $false
                ErrorAction = "Stop"
            }
            $null = Remove-PSSession @deleteExchangeSessionSplatParams

            # Send disconnect audit log
            $Log = @{
                Action            = "UpdateResource" # optional. ENUM (undefined = default) 
                System            = "Exchange On-Premises" # optional (free format text) 
                Message           = "Successfully disconnected from Exchange using URI [$ExchangeConnectionUri]" # required (free format text) 
                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $ExchangeConnectionUri # optional (free format text) 
                TargetIdentifier  = $([string]$exchangeSession.InstanceId) # optional (free format text) 
            }
            Write-Information -Tags "Audit" -MessageData $log
        }
        catch {
            Write-Warning "Failed to disconnect from Exchange using URI [$ExchangeConnectionUri]. Error: $($_.Exception.Message)"
        }
    }
}
