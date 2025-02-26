# Microsoft 365 Licence Management Tool

A comprehensive PowerShell tool for managing Microsoft 365 user mailboxes, licenses, and roles. This tool is designed to streamline the process of converting standard user mailboxes to shared mailboxes and properly managing associated licenses and security settings.

## Features

- **Generate Reports**: Identify inactive users and potential cost-saving opportunities
- **Mailbox Conversion**: Convert eligible user mailboxes to shared mailboxes
- **License Management**: Remove unnecessary licenses from converted mailboxes
- **Security Controls**:
  - Set receive limits on shared mailboxes
  - Block sign-in for converted shared mailboxes
  - Remove roles from converted users
- **Hybrid Environment Support**: Disable on-premises AD accounts for converted mailboxes
- **Comparison Reporting**: Generate before/after comparisons to track changes

## Prerequisites

- PowerShell 5.1 or newer
- Exchange Online PowerShell module
  ```powershell
  Install-Module -Name ExchangeOnlineManagement
  ```
- Microsoft Graph PowerShell SDK
  ```powershell
  Install-Module -Name Microsoft.Graph
  ```
- For hybrid environments (optional): Active Directory module
  ```powershell
  Install-WindowsFeature RSAT-AD-PowerShell
  ```

## Installation

1. Clone this repository
2. Ensure you have the required PowerShell modules installed
3. Run the script from PowerShell

```powershell
.\Microsoft365-LicenceManagement.ps1
```

## Usage

The tool provides an interactive menu with the following options:

1. Create report of eligible users
2. Convert eligible users to shared mailboxes
3. Set receive limit to 0KB for converted mailboxes
4. Block sign-in for converted shared mailboxes
5. Remove licenses from converted users
6. Remove roles from converted users
7. Disable on-premises AD accounts (hybrid environments)
8. Create before/after comparison report

Simply follow the on-screen prompts to perform the desired operations.

## Best Practices

1. Always run the report generation first (Option 1)
2. Review the report before performing any changes
3. Follow the menu options in sequential order for a complete mailbox lifecycle management
4. Back up your environment before making significant changes
5. Test in a non-production environment first

## License

This project is licensed under the MIT License - see the LICENSE file for details

## Acknowledgments

* Microsoft Graph and Exchange Online documentation
* PowerShell community resources
