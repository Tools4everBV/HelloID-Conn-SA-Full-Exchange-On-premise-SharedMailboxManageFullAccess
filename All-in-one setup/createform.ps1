# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

#HelloID variables
#Note: when running this script inside HelloID; portalUrl and API credentials are provided automatically (generate and save API credentials first in your admin panel!)
$portalUrl = "https://CUSTOMER.helloid.com"
$apiKey = "API_KEY"
$apiSecret = "API_SECRET"
$delegatedFormAccessGroupNames = @("") #Only unique names are supported. Groups must exist!
$delegatedFormCategories = @("Exchange On-Premises") #Only unique names are supported. Categories will be created if not exists
$script:debugLogging = $false #Default value: $false. If $true, the HelloID resource GUIDs will be shown in the logging
$script:duplicateForm = $false #Default value: $false. If $true, the HelloID resource names will be changed to import a duplicate Form
$script:duplicateFormSuffix = "_tmp" #the suffix will be added to all HelloID resource names to generate a duplicate form with different resource names

#The following HelloID Global variables are used by this form. No existing HelloID global variables will be overriden only new ones are created.
#NOTE: You can also update the HelloID Global variable values afterwards in the HelloID Admin Portal: https://<CUSTOMER>.helloid.com/admin/variablelibrary
$globalHelloIDVariables = [System.Collections.Generic.List[object]]@();

#Global variable #1 >> ExchangeConnectionUri
$tmpName = @'
ExchangeConnectionUri
'@ 
$tmpValue = @'
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False" });

#Global variable #2 >> ExchangeAdminPassword
$tmpName = @'
ExchangeAdminPassword
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "True" });

#Global variable #3 >> ExchangeAdminUsername
$tmpName = @'
ExchangeAdminUsername
'@ 
$tmpValue = @'
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False" });


#make sure write-information logging is visual
$InformationPreference = "continue"

# Check for prefilled API Authorization header
if (-not [string]::IsNullOrEmpty($portalApiBasic)) {
    $script:headers = @{"authorization" = $portalApiBasic }
    Write-Information "Using prefilled API credentials"
}
else {
    # Create authorization headers with HelloID API key
    $pair = "$apiKey" + ":" + "$apiSecret"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $base64 = [System.Convert]::ToBase64String($bytes)
    $key = "Basic $base64"
    $script:headers = @{"authorization" = $Key }
    Write-Information "Using manual API credentials"
}

# Check for prefilled PortalBaseURL
if (-not [string]::IsNullOrEmpty($portalBaseUrl)) {
    $script:PortalBaseUrl = $portalBaseUrl
    Write-Information "Using prefilled PortalURL: $script:PortalBaseUrl"
}
else {
    $script:PortalBaseUrl = $portalUrl
    Write-Information "Using manual PortalURL: $script:PortalBaseUrl"
}

# Define specific endpoint URI
$script:PortalBaseUrl = $script:PortalBaseUrl.trim("/") + "/"  

# Make sure to reveive an empty array using PowerShell Core
function ConvertFrom-Json-WithEmptyArray([string]$jsonString) {
    # Running in PowerShell Core?
    if ($IsCoreCLR -eq $true) {
        $r = [Object[]]($jsonString | ConvertFrom-Json -NoEnumerate)
        return , $r  # Force return value to be an array using a comma
    }
    else {
        $r = [Object[]]($jsonString | ConvertFrom-Json)
        return , $r  # Force return value to be an array using a comma
    }
}

function Invoke-HelloIDGlobalVariable {
    param(
        [parameter(Mandatory)][String]$Name,
        [parameter(Mandatory)][String][AllowEmptyString()]$Value,
        [parameter(Mandatory)][String]$Secret
    )

    $Name = $Name + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        $uri = ($script:PortalBaseUrl + "api/v1/automation/variables/named/$Name")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false

        if ([string]::IsNullOrEmpty($response.automationVariableGuid)) {
            #Create Variable
            $body = @{
                name     = $Name;
                value    = $Value;
                secret   = $Secret;
                ItemType = 0;
            }    
            $body = ConvertTo-Json -InputObject $body -Depth 100

            $uri = ($script:PortalBaseUrl + "api/v1/automation/variable")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            $variableGuid = $response.automationVariableGuid

            Write-Information "Variable '$Name' created$(if ($script:debugLogging -eq $true) { ": " + $variableGuid })"
        }
        else {
            $variableGuid = $response.automationVariableGuid
            Write-Warning "Variable '$Name' already exists$(if ($script:debugLogging -eq $true) { ": " + $variableGuid })"
        }
    }
    catch {
        Write-Error "Variable '$Name', message: $_"
    }
}

function Invoke-HelloIDAutomationTask {
    param(
        [parameter(Mandatory)][String]$TaskName,
        [parameter(Mandatory)][String]$UseTemplate,
        [parameter(Mandatory)][String]$AutomationContainer,
        [parameter(Mandatory)][String][AllowEmptyString()]$Variables,
        [parameter(Mandatory)][String]$PowershellScript,
        [parameter()][String][AllowEmptyString()]$ObjectGuid,
        [parameter()][String][AllowEmptyString()]$ForceCreateTask,
        [parameter(Mandatory)][Ref]$returnObject
    )

    $TaskName = $TaskName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        $uri = ($script:PortalBaseUrl + "api/v1/automationtasks?search=$TaskName&container=$AutomationContainer")
        $responseRaw = (Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false) 
        $response = $responseRaw | Where-Object -filter { $_.name -eq $TaskName }

        if ([string]::IsNullOrEmpty($response.automationTaskGuid) -or $ForceCreateTask -eq $true) {
            #Create Task

            $body = @{
                name                = $TaskName;
                useTemplate         = $UseTemplate;
                powerShellScript    = $PowershellScript;
                automationContainer = $AutomationContainer;
                objectGuid          = $ObjectGuid;
                variables           = (ConvertFrom-Json-WithEmptyArray($Variables));
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100

            $uri = ($script:PortalBaseUrl + "api/v1/automationtasks/powershell")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            $taskGuid = $response.automationTaskGuid

            Write-Information "Powershell task '$TaskName' created$(if ($script:debugLogging -eq $true) { ": " + $taskGuid })"
        }
        else {
            #Get TaskGUID
            $taskGuid = $response.automationTaskGuid
            Write-Warning "Powershell task '$TaskName' already exists$(if ($script:debugLogging -eq $true) { ": " + $taskGuid })"
        }
    }
    catch {
        Write-Error "Powershell task '$TaskName', message: $_"
    }

    $returnObject.Value = $taskGuid
}

function Invoke-HelloIDDatasource {
    param(
        [parameter(Mandatory)][String]$DatasourceName,
        [parameter(Mandatory)][String]$DatasourceType,
        [parameter(Mandatory)][String][AllowEmptyString()]$DatasourceModel,
        [parameter()][String][AllowEmptyString()]$DatasourceStaticValue,
        [parameter()][String][AllowEmptyString()]$DatasourcePsScript,        
        [parameter()][String][AllowEmptyString()]$DatasourceInput,
        [parameter()][String][AllowEmptyString()]$AutomationTaskGuid,
        [parameter()][String][AllowEmptyString()]$DatasourceRunInCloud,
        [parameter(Mandatory)][Ref]$returnObject
    )

    $DatasourceName = $DatasourceName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    $datasourceTypeName = switch ($DatasourceType) { 
        "1" { "Native data source"; break } 
        "2" { "Static data source"; break } 
        "3" { "Task data source"; break } 
        "4" { "Powershell data source"; break }
    }

    try {
        $uri = ($script:PortalBaseUrl + "api/v1/datasource/named/$DatasourceName")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
    
        if ([string]::IsNullOrEmpty($response.dataSourceGUID)) {
            #Create DataSource
            $body = @{
                name               = $DatasourceName;
                type               = $DatasourceType;
                model              = (ConvertFrom-Json-WithEmptyArray($DatasourceModel));
                automationTaskGUID = $AutomationTaskGuid;
                value              = (ConvertFrom-Json-WithEmptyArray($DatasourceStaticValue));
                script             = $DatasourcePsScript;
                input              = (ConvertFrom-Json-WithEmptyArray($DatasourceInput));
                runInCloud         = $DatasourceRunInCloud;
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl + "api/v1/datasource")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            
            $datasourceGuid = $response.dataSourceGUID
            Write-Information "$datasourceTypeName '$DatasourceName' created$(if ($script:debugLogging -eq $true) { ": " + $datasourceGuid })"
        }
        else {
            #Get DatasourceGUID
            $datasourceGuid = $response.dataSourceGUID
            Write-Warning "$datasourceTypeName '$DatasourceName' already exists$(if ($script:debugLogging -eq $true) { ": " + $datasourceGuid })"
        }
    }
    catch {
        Write-Error "$datasourceTypeName '$DatasourceName', message: $_"
    }

    $returnObject.Value = $datasourceGuid
}

function Invoke-HelloIDDynamicForm {
    param(
        [parameter(Mandatory)][String]$FormName,
        [parameter(Mandatory)][String]$FormSchema,
        [parameter(Mandatory)][Ref]$returnObject
    )

    $FormName = $FormName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        try {
            $uri = ($script:PortalBaseUrl + "api/v1/forms/$FormName")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        }
        catch {
            $response = $null
        }

        if (([string]::IsNullOrEmpty($response.dynamicFormGUID)) -or ($response.isUpdated -eq $true)) {
            #Create Dynamic form
            $body = @{
                Name       = $FormName;
                FormSchema = (ConvertFrom-Json-WithEmptyArray($FormSchema));
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100

            $uri = ($script:PortalBaseUrl + "api/v1/forms")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body

            $formGuid = $response.dynamicFormGUID
            Write-Information "Dynamic form '$formName' created$(if ($script:debugLogging -eq $true) { ": " + $formGuid })"
        }
        else {
            $formGuid = $response.dynamicFormGUID
            Write-Warning "Dynamic form '$FormName' already exists$(if ($script:debugLogging -eq $true) { ": " + $formGuid })"
        }
    }
    catch {
        Write-Error "Dynamic form '$FormName', message: $_"
    }

    $returnObject.Value = $formGuid
}


function Invoke-HelloIDDelegatedForm {
    param(
        [parameter(Mandatory)][String]$DelegatedFormName,
        [parameter(Mandatory)][String]$DynamicFormGuid,
        [parameter()][Array][AllowEmptyString()]$AccessGroups,
        [parameter()][String][AllowEmptyString()]$Categories,
        [parameter(Mandatory)][String]$UseFaIcon,
        [parameter()][String][AllowEmptyString()]$FaIcon,
        [parameter()][String][AllowEmptyString()]$task,
        [parameter(Mandatory)][Ref]$returnObject
    )
    $delegatedFormCreated = $false
    $DelegatedFormName = $DelegatedFormName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        try {
            $uri = ($script:PortalBaseUrl + "api/v1/delegatedforms/$DelegatedFormName")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        }
        catch {
            $response = $null
        }

        if ([string]::IsNullOrEmpty($response.delegatedFormGUID)) {
            #Create DelegatedForm
            $body = @{
                name            = $DelegatedFormName;
                dynamicFormGUID = $DynamicFormGuid;
                isEnabled       = "True";
                useFaIcon       = $UseFaIcon;
                faIcon          = $FaIcon;
                task            = ConvertFrom-Json -inputObject $task;
            }
            if (-not[String]::IsNullOrEmpty($AccessGroups)) { 
                $body += @{
                    accessGroups = (ConvertFrom-Json-WithEmptyArray($AccessGroups));
                }
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100

            $uri = ($script:PortalBaseUrl + "api/v1/delegatedforms")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body

            $delegatedFormGuid = $response.delegatedFormGUID
            Write-Information "Delegated form '$DelegatedFormName' created$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormGuid })"
            $delegatedFormCreated = $true

            $bodyCategories = $Categories
            $uri = ($script:PortalBaseUrl + "api/v1/delegatedforms/$delegatedFormGuid/categories")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $bodyCategories
            Write-Information "Delegated form '$DelegatedFormName' updated with categories"
        }
        else {
            #Get delegatedFormGUID
            $delegatedFormGuid = $response.delegatedFormGUID
            Write-Warning "Delegated form '$DelegatedFormName' already exists$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormGuid })"
        }
    }
    catch {
        Write-Error "Delegated form '$DelegatedFormName', message: $_"
    }

    $returnObject.value.guid = $delegatedFormGuid
    $returnObject.value.created = $delegatedFormCreated
}

<# Begin: HelloID Global Variables #>
foreach ($item in $globalHelloIDVariables) {
    Invoke-HelloIDGlobalVariable -Name $item.name -Value $item.value -Secret $item.secret 
}
<# End: HelloID Global Variables #>


<# Begin: HelloID Data sources #>
<# Begin: DataSource "exchange-on-premises-sharedmailbox-manage-full-access-permissions | Exchange-On-Premises-Get-All-Users" #>
$tmpPsScript = @'
# Build filter
# Check for mailboxes matching the displayName, mailNickname (alias), primary email or proxy addresses
# This will check ALL users (enabled and disabled), including shared/room/equipment mailboxes
$filter = "RecipientType -eq 'UserMailbox'"

# Global variables
# Outcommented as these are set from Global Variables
# $ExchangeConnectionUri = ""
# $ExchangeAdminUsername = ""
# $ExchangeAdminPassword = ""

# Fixed values
# Properties to select - Select only needed properties to limit memory usage and speed up processing
$propertiesToSelect = @(
    "Guid"
    , "DisplayName"
    , "PrimarySmtpAddress"
    , "UserPrincipalName"    
    , "samAccountName"
    , "Name"
)

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Set debug logging
$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

#region functions
#endregion functions

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
    $null = Import-PSSession -Session $exchangeSession -DisableNameChecking -AllowClobber -CommandName "Get-Mailbox" -ErrorAction Stop

     # Get Mailboxes
    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/get-mailbox
    $actionMessage = "querying shared mailboxes that match filter [$($filter)]"
    $getMailboxesSplatParams = @{
        Filter      = $filter
        ResultSize  = "Unlimited"
        ErrorAction = 'Stop'
    }

    $mailboxes = Get-Mailbox @getMailboxesSplatParams | Select-Object -Property $propertiesToSelect
    Write-Information "Queried shared mailboxes that match filter [$($filter)]. Result count: $(($mailboxes | Measure-Object).Count)"

    # Sort and send results to HelloID
    $actionMessage = "sending results to HelloID"
    $mailboxes | Sort-Object -Property DisplayName | ForEach-Object {
        Write-Output $_
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
'@ 
$tmpModel = @'
[{"key":"Guid","type":0},{"key":"DisplayName","type":0},{"key":"PrimarySmtpAddress","type":0},{"key":"UserPrincipalName","type":0},{"key":"SamAccountName","type":0},{"key":"Name","type":0}]
'@ 
$tmpInput = @'
[]
'@ 
$dataSourceGuid_1 = [PSCustomObject]@{} 
$dataSourceGuid_1_Name = @'
exchange-on-premises-sharedmailbox-manage-full-access-permissions | Exchange-On-Premises-Get-All-Users
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_1_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -DataSourceRunInCloud "False" -returnObject ([Ref]$dataSourceGuid_1) 
<# End: DataSource "exchange-on-premises-sharedmailbox-manage-full-access-permissions | Exchange-On-Premises-Get-All-Users" #>

<# Begin: DataSource "exchange-on-premises-sharedmailbox-manage-full-access-permissions | Exchange-On-Premises-Get-Current-Full-Access-Users" #>
$tmpPsScript = @'
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
try{
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

    $permissions = Get-MailboxPermission @getMailboxesSplatParams | Where-Object {($_.IsInherited -eq $false) -and -not ($_.User -like "*NT AUTHORITY*") -and ($_.AccessRights -like "*FullAccess*")} | Select-Object  @{Name="Displayname"; Expression={(Get-Recipient $_.user.ToString()).Displayname.ToString()}}, @{Name="Samaccountname"; Expression={(Get-Recipient $_.user.ToString()).sAMAccountName.ToString()}}
    Write-Information "Queried shared mailboxes that match Name [$($mailbox.displayName)]. Result fullaccess users count: $(($permissions | Measure-Object).Count)"

    $permissions = $permissions | Sort-Object -Property Displayname
    foreach($permission in $permissions)
    {
        $displayValue = $permission.Displayname
        $returnObject = @{SamAccountName=$permission.Samaccountname;Name=$displayValue;}
        Write-Output $returnObject
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
'@ 
$tmpModel = @'
[{"key":"Name","type":0},{"key":"SamAccountName","type":0}]
'@ 
$tmpInput = @'
[{"description":null,"translateDescription":false,"inputFieldType":1,"key":"selectedMailbox","type":0,"options":1}]
'@ 
$dataSourceGuid_2 = [PSCustomObject]@{} 
$dataSourceGuid_2_Name = @'
exchange-on-premises-sharedmailbox-manage-full-access-permissions | Exchange-On-Premises-Get-Current-Full-Access-Users
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_2_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -DataSourceRunInCloud "False" -returnObject ([Ref]$dataSourceGuid_2) 
<# End: DataSource "exchange-on-premises-sharedmailbox-manage-full-access-permissions | Exchange-On-Premises-Get-Current-Full-Access-Users" #>

<# Begin: DataSource "exchange-on-premises-sharedmailbox-manage-full-access-permissions | Exchange-On-Premises-Get-Sharedmailbox-Wildcard-Name-Alias" #>
$tmpPsScript = @'
# Variables configured in form
$searchValue = $datasource.searchValue
if ($searchValue -eq "*") {
    $filter = "RecipientTypeDetails -eq 'SharedMailbox'"
}
else {
    $filter = "RecipientTypeDetails -eq 'SharedMailbox' -and (Name -like '*$searchValue*' -or SamAccountName -like '*$searchValue*' -or Alias -like '*$searchValue*' -or PrimarySmtpAddress -like '*$searchValue*')"
}

# Global variables
# Outcommented as these are set from Global Variables
# $ExchangeConnectionUri = ""
# $ExchangeAdminUsername = ""
# $ExchangeAdminPassword = ""

# Fixed values
# Properties to select - Select only needed properties to limit memory usage and speed up processing
$propertiesToSelect = @(
    "Guid"
    , "DisplayName"
    , "Name"
    , "Alias"
    , "PrimarySmtpAddress"
    , "EmailAddresses"
    , "UserPrincipalName"
    , "RecipientTypeDetails"
    , "HiddenFromAddressListsEnabled"
)

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Set debug logging
$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

#region functions
#endregion functions

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
    $null = Import-PSSession -Session $exchangeSession -DisableNameChecking -AllowClobber -CommandName "Get-Mailbox" -ErrorAction Stop

     # Get Mailboxes
    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/get-mailbox
    $actionMessage = "querying shared mailboxes that match filter [$($filter)]"

    $getMailboxesSplatParams = @{
        Filter      = $filter
        ResultSize  = "Unlimited"
        ErrorAction = 'Stop'
    }

    $mailboxes = Get-Mailbox @getMailboxesSplatParams | Select-Object -Property $propertiesToSelect
    Write-Information "Queried shared mailboxes that match filter [$($filter)]. Result count: $(($mailboxes | Measure-Object).Count)"

    # Sort and send results to HelloID
    $actionMessage = "sending results to HelloID"
    $mailboxes | Sort-Object -Property DisplayName | ForEach-Object {
        Write-Output $_
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
'@ 
$tmpModel = @'
[{"key":"Guid","type":0},{"key":"DisplayName","type":0},{"key":"Name","type":0},{"key":"Alias","type":0},{"key":"PrimarySmtpAddress","type":0},{"key":"EmailAddresses","type":0},{"key":"UserPrincipalName","type":0},{"key":"RecipientTypeDetails","type":0},{"key":"HiddenFromAddressListsEnabled","type":0}]
'@ 
$tmpInput = @'
[{"description":null,"translateDescription":false,"inputFieldType":1,"key":"searchValue","type":0,"options":1}]
'@ 
$dataSourceGuid_0 = [PSCustomObject]@{} 
$dataSourceGuid_0_Name = @'
exchange-on-premises-sharedmailbox-manage-full-access-permissions | Exchange-On-Premises-Get-Sharedmailbox-Wildcard-Name-Alias
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_0_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -DataSourceRunInCloud "False" -returnObject ([Ref]$dataSourceGuid_0) 
<# End: DataSource "exchange-on-premises-sharedmailbox-manage-full-access-permissions | Exchange-On-Premises-Get-Sharedmailbox-Wildcard-Name-Alias" #>
<# End: HelloID Data sources #>

<# Begin: Dynamic Form "Exchange On-Premises - Sharedmailbox - Manage full access permissions" #>
$tmpSchema = @"
[{"label":"Search Sharedmailbox","fields":[{"key":"searchMailbox","templateOptions":{"label":"Search Sharedmailbox","placeholder":""},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"gridMailbox","templateOptions":{"label":"Select Sharedmailbox","required":true,"grid":{"columns":[{"headerName":"Display Name","field":"DisplayName"},{"headerName":"Primary Smtp Address","field":"PrimarySmtpAddress"},{"headerName":"Email Addresses","field":"EmailAddresses"},{"headerName":"Alias","field":"Alias"},{"headerName":"Recipient Type Details","field":"RecipientTypeDetails"}],"height":300,"rowSelection":"single"},"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_0","input":{"propertyInputs":[{"propertyName":"searchValue","otherFieldValue":{"otherFieldKey":"searchMailbox"}}]}},"useDefault":false,"allowCsvDownload":true,"useFilter":true},"type":"grid","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":true}]},{"label":"Mailbox Permissions","fields":[{"key":"fullaccessList","templateOptions":{"label":"User(s) to manage full access permission","required":false,"filterable":true,"useDataSource":true,"dualList":{"options":[{"guid":"75ea2890-88f8-4851-b202-626123054e14","Name":"Apple"},{"guid":"0607270d-83e2-4574-9894-0b70011b663f","Name":"Pear"},{"guid":"1ef6fe01-3095-4614-a6db-7c8cd416ae3b","Name":"Orange"}],"optionKeyProperty":"SamAccountName","optionDisplayProperty":"Name"},"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_1","input":{"propertyInputs":[]}},"destinationDataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_2","input":{"propertyInputs":[{"propertyName":"selectedMailbox","otherFieldValue":{"otherFieldKey":"gridMailbox"}}]}}},"type":"duallist","summaryVisibility":"Show","sourceDataSourceIdentifierSuffix":"source-datasource","destinationDataSourceIdentifierSuffix":"destination-datasource","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"blnautomapping","templateOptions":{"label":"Automapping","useSwitch":true,"checkboxLabel":"Yes"},"type":"boolean","defaultValue":true,"summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false}]}]
"@ 

$dynamicFormGuid = [PSCustomObject]@{} 
$dynamicFormName = @'
Exchange On-Premises - Sharedmailbox - Manage full access permissions
'@ 
Invoke-HelloIDDynamicForm -FormName $dynamicFormName -FormSchema $tmpSchema  -returnObject ([Ref]$dynamicFormGuid) 
<# END: Dynamic Form #>

<# Begin: Delegated Form Access Groups and Categories #>
$delegatedFormAccessGroupGuids = @()
if (-not[String]::IsNullOrEmpty($delegatedFormAccessGroupNames)) {
    foreach ($group in $delegatedFormAccessGroupNames) {
        try {
            $uri = ($script:PortalBaseUrl + "api/v1/groups/$group")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
            $delegatedFormAccessGroupGuid = $response.groupGuid
            $delegatedFormAccessGroupGuids += $delegatedFormAccessGroupGuid
        
            Write-Information "HelloID (access)group '$group' successfully found$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormAccessGroupGuid })"
        }
        catch {
            Write-Error "HelloID (access)group '$group', message: $_"
        }
    }
    if ($null -ne $delegatedFormAccessGroupGuids) {
        $delegatedFormAccessGroupGuids = ($delegatedFormAccessGroupGuids | Select-Object -Unique | ConvertTo-Json -Depth 100 -Compress)
    }
}

$delegatedFormCategoryGuids = @()
foreach ($category in $delegatedFormCategories) {
    try {
        $uri = ($script:PortalBaseUrl + "api/v1/delegatedformcategories/$category")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        $response = $response | Where-Object { $_.name.en -eq $category }
    
        $tmpGuid = $response.delegatedFormCategoryGuid
        $delegatedFormCategoryGuids += $tmpGuid
    
        Write-Information "HelloID Delegated Form category '$category' successfully found$(if ($script:debugLogging -eq $true) { ": " + $tmpGuid })"
    }
    catch {
        Write-Warning "HelloID Delegated Form category '$category' not found"
        $body = @{
            name = @{"en" = $category };
        }
        $body = ConvertTo-Json -InputObject $body -Depth 100

        $uri = ($script:PortalBaseUrl + "api/v1/delegatedformcategories")
        $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
        $tmpGuid = $response.delegatedFormCategoryGuid
        $delegatedFormCategoryGuids += $tmpGuid

        Write-Information "HelloID Delegated Form category '$category' successfully created$(if ($script:debugLogging -eq $true) { ": " + $tmpGuid })"
    }
}
$delegatedFormCategoryGuids = (ConvertTo-Json -InputObject $delegatedFormCategoryGuids -Depth 100 -Compress)
<# End: Delegated Form Access Groups and Categories #>

<# Begin: Delegated Form #>
$delegatedFormRef = [PSCustomObject]@{guid = $null; created = $null } 
$delegatedFormName = @'
Exchange On-Premises - Sharedmailbox - Manage full access permissions
'@
$tmpTask = @'
{"name":"Exchange On-Premises - Sharedmailbox - Manage full access permissions","script":"# variables configured in form\r\n$mailbox = $form.gridMailbox\r\n$usersToAdd = $form.fullaccessList.leftToRight\r\n$usersToRemove = $form.fullaccessList.rightToLeft\r\n$AutoMapping = $form.blnautomapping\r\n\r\n# Global variables\r\n# Outcommented as these are set from Global Variables\r\n# $ExchangeConnectionUri = \"\"\r\n# $ExchangeAdminUsername = \"\"\r\n# $ExchangeAdminPassword = \"\"\r\n\r\n# Fixed values\r\n$commands = @(\r\n    \"Get-Mailbox\",    \r\n    \"Add-MailboxPermission\",\r\n    \"Remove-MailboxPermission\"\r\n)\r\n\r\n# Enable TLS1.2\r\n[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12\r\n\r\n# Set debug logging\r\n$VerbosePreference = \"SilentlyContinue\"\r\n$InformationPreference = \"Continue\"\r\n$WarningPreference = \"Continue\"\r\n\r\ntry {\r\n     # Create credentials\r\n    $actionMessage = \"creating credentials object\"\r\n    \r\n    $securePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force\r\n    $credential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername, $securePassword)\r\n    \r\n    Write-Verbose \"Created credentials for user [$ExchangeAdminUsername]\"\r\n\r\n    # Connect to Exchange On-Premises\r\n    # Docs: https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-servers-using-remote-powershell\r\n    $actionMessage = \"connecting to Exchange On-Premises\"\r\n\r\n    $sessionOptionParams = @{\r\n        SkipCACheck         = $false\r\n        SkipCNCheck         = $false\r\n        SkipRevocationCheck = $false\r\n    }\r\n\r\n    $sessionOption = New-PSSessionOption @sessionOptionParams\r\n\r\n    $sessionParams = @{\r\n        Authentication    = \u0027Default\u0027\r\n        ConfigurationName = \u0027Microsoft.Exchange\u0027\r\n        Credential        = $credential\r\n        ConnectionUri     = $ExchangeConnectionUri\r\n        SessionOption     = $sessionOption\r\n        ErrorAction       = \"Stop\"\r\n    }\r\n\r\n    $exchangeSession = New-PSSession @sessionParams\r\n    $null = Import-PSSession -Session $exchangeSession -DisableNameChecking -AllowClobber -CommandName $commands -ErrorAction Stop\r\n\r\n    # Send initial audit log\r\n    $Log = @{\r\n        Action            = \"UpdateResource\" # optional. ENUM (undefined = default) \r\n        System            = \"Exchange On-Premises\" # optional (free format text) \r\n        Message           = \"Successfully connected to Exchange using URI [$ExchangeConnectionUri]\" # required (free format text) \r\n        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n        TargetDisplayName = $ExchangeConnectionUri # optional (free format text) \r\n        TargetIdentifier  = $([string]$exchangeSession.InstanceId) # optional (free format text) \r\n    }\r\n    Write-Information -Tags \"Audit\" -MessageData $log\r\n\r\n    # Add Full Access Permissions\r\n    if ($usersToAdd.Count -gt 0) {\r\n        Write-Information \"Starting to grant permission [FullAccess] members to mailbox $($mailbox.DisplayName)\"\r\n        \r\n        foreach ($user in $usersToAdd) {\r\n            try {      \r\n                $actionMessage = \"granting permission [FullAccess] to mailbox [$($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress))] for user [$($user.userPrincipalName) ($($user.Guid))]\"      \r\n                \r\n                $FullAccessPermissionSplatParams = @{\r\n                    Identity      = $($mailbox.Guid)\r\n                    User          = $($user.sAMAccountName)\r\n                    AccessRights  = \"FullAccess\"\r\n                    InheritanceType = \"All\"\r\n                    AutoMapping   = [bool]$AutoMapping\r\n                    ErrorAction   = \"Stop\"                 \r\n                }\r\n                $addFullAccessPermission = Add-MailboxPermission @FullAccessPermissionSplatParams\r\n                \r\n                Write-Information \"Granting access rights [FullAccess] on mailbox [$($mailbox.DisplayName)] for [$($user.DisplayName)] successfully.\"\r\n                $Log = @{\r\n                    Action            = \"GrantMembership\" # optional. ENUM (undefined = default) \r\n                    System            = \"Exchange On-Premises\" # optional (free format text) \r\n                    Message           = \"Granting access rights [FullAccess] on mailbox [$($mailbox.DisplayName)] for [$($user.DisplayName)] successfully.\" # required (free format text) \r\n                    IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n                    TargetDisplayName = $($mailbox.DisplayName) # optional (free format text) \r\n                    TargetIdentifier  = $([string]$mailbox.Guid) # optional (free format text) \r\n                }\r\n                #send result back  \r\n                Write-Information -Tags \"Audit\" -MessageData $log\r\n\r\n            }\r\n            catch {\r\n                $ex = $PSItem\r\n                if (-not [string]::IsNullOrEmpty($ex.Exception.Message)) {\r\n                    $warningMessage = \"Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)\"\r\n                    $auditMessage = \"Error $($actionMessage). Error: $($ex.Exception.Message)\"\r\n                }\r\n                else {\r\n                    $warningMessage = \"Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception)\"\r\n                    $auditMessage = \"Error $($actionMessage). Error: $($ex.Exception)\"\r\n                }\r\n\r\n                Write-Error \"Error granting access rights [FullAccess] for [$($user.DisplayName)] on mailbox [$($mailbox.DisplayName)]. Error: $($_.Exception.Message)\" \r\n                $Log = @{\r\n                    Action            = \"GrantMembership\" # optional. ENUM (undefined = default) \r\n                    System            = \"Exchange On-Premises\" # optional (free format text) \r\n                    Message           = \"Error assigning access rights [FullAccess] to [$($user.DisplayName)] on mailbox [$($mailbox.DisplayName)]\" # required (free format text) \r\n                    IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n                    TargetDisplayName = $($mailbox.DisplayName) # optional (free format text) \r\n                    TargetIdentifier  = $([string]$mailbox.Guid) # optional (free format text) \r\n                }\r\n                #send result back  \r\n                Write-Information -Tags \"Audit\" -MessageData $log\r\n                Write-Warning $warningMessage\r\n                Write-Error $auditMessage            \r\n            }\r\n            break\r\n        }\r\n    }\r\n\r\n    # Remove Full Access Permissions\r\n    if ($usersToRemove.Count -gt 0) {\r\n        Write-Information \"Starting to revoke permission [FullAccess] on mailbox [$($mailbox.DisplayName)]\"\r\n        \r\n        foreach ($user in $usersToRemove) {\r\n            try {\r\n                $actionMessage = \"revoking permission [FullAccess] to mailbox [$($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress))] for user [$($user.userPrincipalName) ($($user.Guid))]\"      \r\n                $FullAccessPermissionSplatParams = @{\r\n                    Identity      = $($mailbox.Guid)\r\n                    User          = $($user.sAMAccountName)\r\n                    AccessRights  = \"FullAccess\"\r\n                    InheritanceType = \"All\"                    \r\n                    ErrorAction   = \"Stop\"\r\n                    Confirm      = $false\r\n                }\r\n                $removeFullAccessPermission = Remove-MailboxPermission @FullAccessPermissionSplatParams\r\n                \r\n                Write-Information \"Revoking access rights [FullAccess] on mailbox [$($mailbox.DisplayName)] for [$($user.DisplayName)] successfully\"\r\n                $Log = @{\r\n                    Action            = \"RevokeMembership\" # optional. ENUM (undefined = default) \r\n                    System            = \"Exchange On-Premises\" # optional (free format text) \r\n                    Message           = \"Revoking access rights [FullAccess] on mailbox [$($mailbox.DisplayName)] for [$($user.DisplayName)] successfully.\" # required (free format text) \r\n                    IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n                    TargetDisplayName = $($mailbox.DisplayName) # optional (free format text) \r\n                    TargetIdentifier  = $([string]$mailbox.Guid) # optional (free format text) \r\n                }                \r\n            }\r\n            catch {\r\n                $ex = $PSItem\r\n                if (-not [string]::IsNullOrEmpty($ex.Exception.Message)) {\r\n                    $warningMessage = \"Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)\"\r\n                    $auditMessage = \"Error $($actionMessage). Error: $($ex.Exception.Message)\"\r\n                }\r\n                else {\r\n                    $warningMessage = \"Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception)\"\r\n                    $auditMessage = \"Error $($actionMessage). Error: $($ex.Exception)\"\r\n                }\r\n\r\n                Write-Error \"Error revoking access rights [FullAccess] on mailbox [$($mailbox.DisplayName)] for [$($user.DisplayName)]. Error: $($_.Exception.Message)\"\r\n                $Log = @{\r\n                    Action            = \"RevokeMembership\" # optional. ENUM (undefined = default) \r\n                    System            = \"Exchange On-Premises\" # optional (free format text) \r\n                    Message           = \"Error revoking access rights [FullAccess] on mailbox [$($mailbox.DisplayName)] for [$($user.DisplayName)].\" # required (free format text) \r\n                    IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n                    TargetDisplayName = $($mailbox.DisplayName) # optional (free format text) \r\n                    TargetIdentifier  = $([string]$mailbox.Guid) # optional (free format text) \r\n                }\r\n                #send result back  \r\n                Write-Information -Tags \"Audit\" -MessageData $log\r\n                Write-Warning $warningMessage\r\n                Write-Error $auditMessage\r\n            }\r\n            break\r\n        }\r\n    }\r\n} catch {\r\n    $ex = $PSItem\r\n    if (-not [string]::IsNullOrEmpty($ex.Exception.Message)) {\r\n        $warningMessage = \"Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)\"\r\n        $auditMessage = \"Error $($actionMessage). Error: $($ex.Exception.Message)\"\r\n    }\r\n    else {\r\n        $warningMessage = \"Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception)\"\r\n        $auditMessage = \"Error $($actionMessage). Error: $($ex.Exception)\"\r\n    }\r\n\r\n    # Send error audit log to HelloID\r\n    $Log = @{\r\n        Action            = \"UpdateResource\" # optional. ENUM (undefined = default) \r\n        System            = \"Exchange On-Premises\" # optional (free format text) \r\n        Message           = $auditMessage # required (free format text) \r\n        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n        TargetDisplayName = $($mailbox.DisplayName) # optional (free format text) \r\n        TargetIdentifier  = $([string]$mailbox.Guid) # optional (free format text) \r\n    }\r\n    \r\n    Write-Information -Tags \"Audit\" -MessageData $log\r\n    Write-Warning $warningMessage\r\n    Write-Error $auditMessage\r\n}\r\nfinally {\r\n    # Disconnect from Exchange\r\n    # Docs: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/remove-pssession\r\n    if ($null -ne $exchangeSession) {\r\n        try {\r\n            $deleteExchangeSessionSplatParams = @{\r\n                Session     = $exchangeSession\r\n                Confirm     = $false\r\n                ErrorAction = \"Stop\"\r\n            }\r\n            $null = Remove-PSSession @deleteExchangeSessionSplatParams\r\n\r\n            # Send disconnect audit log\r\n            $Log = @{\r\n                Action            = \"UpdateResource\" # optional. ENUM (undefined = default) \r\n                System            = \"Exchange On-Premises\" # optional (free format text) \r\n                Message           = \"Successfully disconnected from Exchange using URI [$ExchangeConnectionUri]\" # required (free format text) \r\n                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n                TargetDisplayName = $ExchangeConnectionUri # optional (free format text) \r\n                TargetIdentifier  = $([string]$exchangeSession.InstanceId) # optional (free format text) \r\n            }\r\n            Write-Information -Tags \"Audit\" -MessageData $log\r\n        }\r\n        catch {\r\n            Write-Warning \"Failed to disconnect from Exchange using URI [$ExchangeConnectionUri]. Error: $($_.Exception.Message)\"\r\n        }\r\n    }\r\n}","runInCloud":false}
'@ 

Invoke-HelloIDDelegatedForm -DelegatedFormName $delegatedFormName -DynamicFormGuid $dynamicFormGuid -AccessGroups $delegatedFormAccessGroupGuids -Categories $delegatedFormCategoryGuids -UseFaIcon "True" -FaIcon "fa fa-pencil-square" -task $tmpTask -returnObject ([Ref]$delegatedFormRef) 
<# End: Delegated Form #>

