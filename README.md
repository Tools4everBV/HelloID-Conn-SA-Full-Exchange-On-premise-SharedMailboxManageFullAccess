# HelloID-Conn-SA-Full-Exchange-On-Premises-SharedMailbox-Manage-FullAccess-Permissions

| :warning: Important                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| :-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Best Practice:** Use **HelloID Products** for requesting and managing permissions (group memberships, mailbox access, application roles). Products provide governance, approval workflows, admin visibility, and full lifecycle management.<br>Use delegated forms for one-time operational actions (creating resources like shared mailboxes, password resets, attribute updates) only.<br><br>**[Read more: Products vs. Delegated Forms](https://docs.helloid.com/en/service-automation/products-vs--delegated-forms.html)** |

| :information_source: Information                                                                                                                                                                                                                                                                                                                                                          |
| :---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| This repository contains the connector and configuration code only. The implementer is responsible for acquiring the connection details such as username, password, certificate, etc. You might even need to sign a contract or agreement with the supplier before implementing this connector. Please contact the client's application manager to coordinate the connector requirements. |

## Description

HelloID-Conn-SA-Full-Exchange-On-Premises-SharedMailbox-Manage-FullAccess-Permissions is a template designed for use with HelloID Service Automation (SA) Delegated Forms. It can be imported into HelloID and customized according to your requirements.

This HelloID Service Automation Delegated Form provides Exchange On-Premises shared mailbox full access permission management functionality. The following options are available:

1.  Search for a shared mailbox to manage
2.  Select the shared mailbox
3.  Manage members who have full-access rights (Add/Remove)
4.  Confirm the changes
5.  Full access permissions are updated in Exchange On-Premises
6.  Changes are logged for audit purposes

## Getting started

### Requirements

- **Exchange On-Premises PowerShell Module**:<br>
  Access to Exchange On-Premises PowerShell remoting endpoint is required. Ensure the service account has appropriate permissions to manage mailbox permissions.
- **Network Connectivity**:<br>
  The HelloID agent or service must have network access to the Exchange On-Premises server PowerShell endpoint (typically http://servername/powershell).

- **Service Account Permissions**:<br>
  The service account must have sufficient permissions in Exchange to:
  - Query shared mailboxes
  - Query user mailboxes
  - Manage mailbox permissions (Add-MailboxPermission, Remove-MailboxPermission)

### Connection settings

The following user-defined variables are used by the connector.

| Setting               | Description                                 | Mandatory |
| --------------------- | ------------------------------------------- | --------- |
| ExchangeConnectionUri | The URI to the Exchange PowerShell endpoint | Yes       |
| ExchangeAdminUsername | The username of the service account         | Yes       |
| ExchangeAdminPassword | The password of the service account         | Yes       |

## Remarks

### Enhanced Error Handling and Logging

- The updated version includes comprehensive error handling and structured audit logging for all operations.
- All actions are logged with detailed information including action type, system, message, error status, and target identifiers for proper audit trails.

### Improved Connection Management

- Connection parameters have been enhanced with better session options and credential handling.
- TLS 1.2 is enforced for secure connections.
- Session configuration includes proper authentication settings and error handling.

### Optimized Performance

- Data sources now use property selection to limit memory usage and improve processing speed.
- Only necessary properties are retrieved from Exchange, reducing overhead.

### Enhanced Search Capabilities

- Wildcard search has been improved to search across multiple mailbox properties (Name, SamAccountName, Alias, PrimarySmtpAddress).
- Better filter handling for more accurate results.

### Security Improvements

- Password is now marked as secret in global variables.
- Secure credential handling throughout all scripts.

## Development resources

### PowerShell Cmdlets

The following Exchange PowerShell cmdlets are used by the connector:

| Cmdlet                   | Description                                   |
| ------------------------ | --------------------------------------------- |
| Get-Mailbox              | Retrieve mailbox information                  |
| Add-MailboxPermission    | Add full access permissions to a mailbox      |
| Remove-MailboxPermission | Remove full access permissions from a mailbox |

### API documentation

For more information about Exchange On-Premises PowerShell, please refer to:

- [Connect to Exchange servers using remote PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-servers-using-remote-powershell)
- [Exchange Server PowerShell (Exchange Management Shell)](https://learn.microsoft.com/en-us/powershell/exchange/exchange-management-shell)

## Getting help

> :bulb: **Tip:**  
> _For more information on Delegated Forms, please refer to our [documentation](https://docs.helloid.com/en/service-automation/delegated-forms.html) pages_.

## HelloID docs

The official HelloID documentation can be found at: https://docs.helloid.com/
