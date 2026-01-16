# Barracuda XDR Microsoft 365 Monitoring Setup

A PowerShell GUI application that automates the setup of Barracuda XDR for Microsoft 365 monitoring. This tool simplifies the configuration process by automating application registration, API permissions, admin consent, client secret creation, and audit logging enablement.

## Features

- **Graphical User Interface**: Easy-to-use Windows Forms GUI for all operations
- **Automated Application Registration**: Creates or finds the required Entra ID application
- **API Permissions Management**: Automatically configures Microsoft Graph and Office 365 Management API permissions
- **Admin Consent**: Attempts to grant admin consent programmatically
- **Client Secret Generation**: Creates and displays client secrets with expiration
- **Audit Logging**: Enables unified audit logging for all mailboxes via Exchange Online
- **Individual Copy Buttons**: Quick copy buttons for each credential field
- **Ticket Notes Generator**: One-click copy of formatted ticket notes for documentation
- **Connection Testing**: Verify application registration and permissions

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- Microsoft 365 Global Administrator account
- Internet connection
- Required PowerShell modules (installed automatically):
  - `Microsoft.Graph` (for Entra ID operations)
  - `ExchangeOnlineManagement` (for audit logging)

## Quick Start

1. **Download the script**:
   ```powershell
   git clone https://github.com/monobrau/m365monitoring.git
   cd m365monitoring
   ```

2. **Run the script**:
   ```powershell
   .\Setup-BarracudaXDR.ps1
   ```

3. **Follow the GUI prompts**:
   - Click "Start Setup"
   - Authenticate with your Microsoft 365 Global Admin account
   - Review the generated credentials
   - Copy credentials to Barracuda XDR Dashboard

## What the Script Does

### 1. Module Installation
- Checks for and installs required PowerShell modules
- Installs `Microsoft.Graph` and `ExchangeOnlineManagement` if missing

### 2. Application Registration
- Searches for existing "SKOUTCYBERSECURITY" application
- Creates new application if not found
- Configures as single-tenant application

### 3. API Permissions Configuration
- **Microsoft Graph Permissions**:
  - `User.EnableDisableAccount.All`
  - `User.ReadWrite.All`
- **Office 365 Management API Permissions**:
  - `ActivityFeed.Read`
  - `ActivityFeed.ReadDlp`
  - `ServiceHealth.Read`

### 4. Admin Consent
- Attempts to grant admin consent programmatically
- Provides manual instructions if programmatic consent fails

### 5. Client Secret Creation
- Generates a new client secret
- Sets expiration to 24 months
- Displays secret value (save immediately - it won't be shown again)

### 6. Audit Logging
- Connects to Exchange Online
- Enables unified audit log ingestion for all mailboxes
- Verifies the configuration

## GUI Features

### Credential Fields
- **Application ID**: Displays the App (Client) ID
- **Tenant ID**: Displays the Directory (Tenant) ID
- **Client Secret**: Displays the secret value (masked by default)
- **Copy Buttons**: Individual copy buttons for each credential
- **Show/Hide**: Toggle visibility of the client secret

### Buttons
- **Start Setup**: Begins the automated setup process
- **Enable Audit Logging**: Manually enable audit logging (if skipped during setup)
- **Test Connection**: Verify application registration and permissions
- **Copy Ticket Notes**: Generate and copy formatted ticket notes
- **Save to File**: Save credentials to a JSON file
- **Copy All Credentials**: Copy all credentials at once

### Ticket Notes Format
The "Copy Ticket Notes" button generates formatted notes in this structure:
```
Task - 
    - Configure Barracuda XDR for Microsoft 365 monitoring

Step(s) performed - 
    - [List of all concrete steps performed]

Is the task resolved - 
    - Yes

Next step(s) - 
    - None - all tasks complete
```

## Manual Steps (if needed)

### Office 365 Management API Permissions
If the script cannot automatically add Office 365 Management API permissions:

1. Go to [Azure Portal](https://portal.azure.com) > App registrations > SKOUTCYBERSECURITY
2. Click "API permissions" > "Add a permission"
3. Select "APIs my organization uses"
4. Search for "Office 365 Management API"
5. Select "Application permissions"
6. Add: `ActivityFeed.Read`, `ActivityFeed.ReadDlp`, `ServiceHealth.Read`
7. Click "Add permissions" then "Grant admin consent"

### Admin Consent
If programmatic consent fails:

1. Go to Azure Portal > App registrations > SKOUTCYBERSECURITY
2. Click "API permissions"
3. Click "Grant admin consent for [Your Domain]"
4. Confirm the action

### Audit Logging
If audit logging needs to be enabled manually:

1. Connect to Exchange Online PowerShell:
   ```powershell
   Connect-ExchangeOnline
   ```
2. Enable audit logging:
   ```powershell
   Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true
   ```

## Barracuda XDR Dashboard Configuration

After running the script:

1. Log in to Barracuda XDR Dashboard
2. Navigate to **Administration > Integrations**
3. Click **Setup** on the Microsoft 365 card
4. Enter the credentials from the script:
   - **Application ID**: From the GUI
   - **Directory (Tenant) ID**: From the GUI
   - **Application Secret**: From the GUI
5. Click **Test** to verify connection
6. Click **Save**

## Troubleshooting

### Office 365 Management API Permissions Not Found
- These permissions may not be exposed as AppRoles in all tenants
- Use the manual steps above to add them via Azure Portal
- The script will detect them once manually added

### Exchange Online Connection Fails
- Ensure you're using a Global Administrator account
- Try running PowerShell as Administrator
- Check if Exchange Online is available in your tenant
- Use the "Enable Audit Logging" button to retry

### Admin Consent Fails
- Requires Global Administrator privileges
- Some tenants have policies that prevent programmatic consent
- Use the manual steps provided in the script output

## Files

- `Setup-BarracudaXDR.ps1` - Main PowerShell GUI script
- `Test-O365Permissions.ps1` - Test script for Office 365 Management API permissions
- `enable_audit_logging.ps1` - Standalone audit logging script
- `README.md` - This file
- `QUICKSTART.md` - Quick start guide
- `ADD_O365_PERMISSIONS.md` - Manual permission addition guide

## Version

**Version 1.0** - Initial release

## License

This project is provided as-is for use with Barracuda XDR and Microsoft 365.

## Support

For issues or questions:
- Check the troubleshooting section above
- Review the status messages in the GUI
- Check Azure Portal for application status
- Verify permissions in Barracuda XDR Dashboard

## Contributing

Contributions are welcome! Please ensure any changes maintain compatibility with the existing GUI and functionality.
