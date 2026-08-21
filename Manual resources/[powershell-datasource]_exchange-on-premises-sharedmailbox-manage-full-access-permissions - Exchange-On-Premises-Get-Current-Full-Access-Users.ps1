# variables configured in form
$mailbox = $datasource.selectedmailbox

# Global variables
# Outcommented as these are set from Global Variables
# $ExchangeConnectionUri = ""
# $ExchangeAdminUsername = ""
# $ExchangeAdminPassword = ""

# Fixed values
# Properties to select - Select only needed properties to limit memory usage and speed up processing
$commands = @(
    "Get-MailboxPermission"
    , "Get-Recipient"    
)

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Set debug logging
$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

#region functions
#endregion functions

# Read current mailbox
try {
    # Create credentials
    $actionMessage = "creating credentials object"
    
    $securePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $credential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername, $securePassword)

    # Connect to Exchange On-Premises
    # Docs: https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-servers-using-remote-powershell
    $actionMessage = "connecting to Exchange On-Premises using URI [$ExchangeConnectionUri]"

    $sessionOptionParams = @{
        SkipCACheck         = $false
        SkipCNCheck         = $false
        SkipRevocationCheck = $false
    }

    $sessionOption = New-PSSessionOption @sessionOptionParams

    $sessionParams = @{
        Authentication    = 'Default'
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri     = $ExchangeConnectionUri
        Credential        = $credential
        SessionOption     = $sessionOption
        ErrorAction       = "Stop"
    }

    $exchangeSession = New-PSSession @sessionParams
    $null = Import-PSSession -Session $exchangeSession -DisableNameChecking -AllowClobber -CommandName $commands -ErrorAction Stop

    # Get Mailboxes
    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/get-mailbox
    $actionMessage = "querying shared mailboxes that match Name [$($mailbox.displayName)]"

    $getMailboxesSplatParams = @{
        Identity    = $mailbox.Guid
        ResultSize  = "Unlimited"
        ErrorAction = 'Stop'
    }

    $permissions = Get-MailboxPermission @getMailboxesSplatParams | Where-Object { ($_.IsInherited -eq $false) -and -not ($_.User -like "*NT AUTHORITY*") -and ($_.AccessRights -like "*FullAccess*") } | Select-Object  @{Name = "Displayname"; Expression = { (Get-Recipient $_.user.ToString()).Displayname.ToString() } }, @{Name = "Samaccountname"; Expression = { (Get-Recipient $_.user.ToString()).sAMAccountName.ToString() } }
    Write-Information "Queried shared mailboxes that match Name [$($mailbox.displayName)]. Result fullaccess users count: $(($permissions | Measure-Object).Count)"

    $permissions = $permissions | Sort-Object -Property Displayname
    foreach ($permission in $permissions) {
        $displayValue = $permission.Displayname
        $returnObject = @{SamAccountName = $permission.Samaccountname; Name = $displayValue; }
        Write-Output $returnObject
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
    Write-Warning $warningMessage
    Write-Error $auditMessage
    # exit # use when using multiple try/catch and the script must stop
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
        }
        catch {
            Write-Warning "Failed to disconnect from Exchange using URI [$ExchangeConnectionUri]. Error: $($_.Exception.Message)"
        }
    }
}
