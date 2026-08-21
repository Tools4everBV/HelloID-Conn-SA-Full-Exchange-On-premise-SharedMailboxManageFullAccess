# Changelog

All notable changes to this project will be documented in this file. The format is based on [Keep a Changelog](https://keepachangelog.com/), and this project adheres to [Semantic Versioning](https://semver.org/).

## [2.0.0] - 2026-08-21

### Added

- Added best practice warning callout about using HelloID Products vs. Delegated Forms for permission management
- Added comprehensive audit logging with structured log objects for all operations
- Added TLS 1.2 enforcement for secure connections
- Added property selection in data sources to optimize performance and reduce memory usage
- Added enhanced search capabilities across multiple mailbox properties (Name, SamAccountName, Alias, PrimarySmtpAddress)
- Added improved error handling with detailed error messages and context
- Added verbose and information preference settings for better debugging
- Added documentation links to Microsoft Exchange PowerShell documentation

### Changed

- Updated file naming convention to use longer, more descriptive names with prefixes
- Updated category name from "Exchange On-Premise" to "Exchange On-Premises" for consistency
- Improved credential handling with more secure practices
- Enhanced connection management with better session configuration
- Updated data source names:
  - `Exchange-sharedmailbox-generate-table-wildcard-fullaccess` → `Exchange-On-Premises-Get-Sharedmailbox-Wildcard-Name-Alias`
  - `Exchange-sharedmailbox-generate-table-manage-permissions-fullaccess` → `Exchange-On-Premises-Get-Current-Full-Access-Users`
  - `Exchange-user-generate-table-sharedmailbox-manage-memberships` → `Exchange-On-Premises-Get-All-Users`
- Updated task name from "Exchange on-premise - Manage full access permissions shared mailbox" to "Exchange On-Premises - Sharedmailbox - Manage full access permissions"
- Refactored code structure for better maintainability and readability
- Updated filter logic for more accurate mailbox searches
- Improved session option configuration with better parameter handling

### Removed

- Removed ADsharedMailboxSearchOU global variable (simplified configuration)
- Removed outdated code comments and debugging artifacts

### Fixed

- Fixed inconsistent naming conventions across all files
- Fixed error handling in connection establishment
- Fixed search filter to properly handle wildcard searches

### Security

- Marked ExchangeAdminPassword as secret in global variables configuration
- Enhanced secure credential handling throughout all scripts

## [1.0.2] - 2022-08-24

### Added

- Added version number and updated code for SA-agent and auditlogging

## [1.0.1] - 2021-11-16

### Added

- Added version number and updated all-in-one script

## [1.0.0] - 2021-04-29

Initial release of HelloID-Conn-SA-Full-Exchange-On-Premises-SharedMailbox-Manage-FullAccess-Permissions.

### Added

- Initial release for managing Exchange On-Premises shared mailbox full access permissions
- Support for searching shared mailboxes
- Support for adding and removing full access permissions
- Auto-mapping configuration option
- All-in-one PowerShell setup script
