# Barracuda XDR Microsoft 365 Setup - PowerShell GUI

This PowerShell script provides a graphical user interface (GUI) for setting up Barracuda XDR monitoring of Microsoft 365.

## Features

- ✅ **Windows Forms GUI** - Easy-to-use graphical interface
- ✅ **Automatic application registration** - Creates "SKOUTCYBERSECURITY" app in Entra ID
- ✅ **API permissions configuration** - Automatically configures required permissions
- ✅ **Client secret generation** - Creates 24-month expiry secret
- ✅ **Audit logging enablement** - One-click audit logging setup
- ✅ **Connection testing** - Verify your setup
- ✅ **Credential management** - Copy to clipboard or save to file

## Prerequisites

- **Windows PowerShell 5.1+** or **PowerShell 7+**
- **Microsoft 365/Azure Global Administrator** or **Application Administrator** role
- **Internet connection**
- **Administrator rights** (recommended)

## Installation

1. **Open PowerShell as Administrator** (recommended)

2. **Set execution policy** (if needed):
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

3. **Run the script**:
```powershell
cd C:\Git\m365monitoring
.\Setup-BarracudaXDR.ps1
```

## Usage

### Main Window

The GUI provides the following sections:

1. **Setup Status** - Shows real-time progress and status messages
2. **Application Credentials** - Displays generated credentials
3. **Action Buttons**:
   - **Start Setup** - Begins the automated setup process
   - **Enable Audit Logging** - Enables audit logging via Exchange Online
   - **Test Connection** - Verifies the setup
   - **Close** - Exits the application

### Setup Process

1. Click **"Start Setup"**
2. A browser window will open for Microsoft 365 authentication
3. Select your admin account and authenticate
4. The script will automatically:
   - Install required PowerShell modules (if needed)
   - Connect to Microsoft Graph
   - Register or find the application
   - Configure API permissions
   - Generate client secret
5. Credentials will appear in the "Application Credentials" section

### Credential Management

- **Show/Hide Secret** - Toggle visibility of the client secret
- **Copy All Credentials** - Copies all credentials to clipboard
- **Save to File** - Saves credentials to a JSON file

### Enable Audit Logging

Click **"Enable Audit Logging"** to:
- Install Exchange Online Management module (if needed)
- Connect to Exchange Online
- Enable unified audit log ingestion
- Verify configuration

**Alternative**: Enable via Admin Center at https://portal.office.com > Compliance > Audit

## API Permissions Configured

### Office 365 Management API (Application Permissions)
- `ActivityFeed.Read` - Read activity data for your organization
- `ActivityFeed.ReadDlp` - Read DLP policy events including detected sensitive data
- `ServiceHealth.Read` - Read service health information for your organization

### Microsoft Graph API (Application Permissions)
- `User.EnableDisableAccount.All` - Enable and disable user accounts
- `User.ReadWrite.All` - Read and write all users' full profiles

## Manual Steps Required

### Grant Admin Consent

After running setup:

1. Go to https://portal.azure.com
2. Navigate to: **Azure Active Directory** > **App registrations**
3. Find **SKOUTCYBERSECURITY** application
4. Click **API permissions**
5. Click **Grant admin consent for [Your Domain]**
6. Verify all permissions show "Granted for [Your Domain]"

The script will display a direct link to the Azure Portal page.

### Configure Barracuda XDR Dashboard

1. Log in to Barracuda XDR Dashboard
2. Navigate to: **Administration** > **Integrations**
3. Click **Setup** on the Microsoft 365 card
4. Enter credentials from the GUI:
   - **Application ID**
   - **Directory (Tenant) ID**
   - **Client Secret**
5. Click **Test** to verify connection
6. Click **Save**

## Troubleshooting

### "Module not found" errors
- The script will attempt to install required modules automatically
- If installation fails, run manually:
  ```powershell
  Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
  Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
  ```

### Authentication issues
- Ensure you're using an account with Global Administrator or Application Administrator role
- Check that MFA is properly configured
- Verify internet connectivity

### Permission configuration fails
- Some permissions may need to be added manually in Azure Portal
- Ensure you grant admin consent after the script runs
- Verify Office 365 Management API is available in your tenant

### Audit logging script fails
- Ensure you have Exchange Administrator role
- Verify Exchange Online Management module is installed
- Try manual enablement via Admin Center

### Connection test fails
- Wait 24-48 hours after enabling audit logging
- Verify admin consent was granted for all permissions
- Check that credentials are correct
- Ensure audit logging has been active for at least a few hours

## Output Files

The script can save credentials to a JSON file with the following format:

```json
{
  "application_id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "directory_tenant_id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "client_secret": "your-secret-value",
  "created_at": "2026-01-16T15:30:00.0000000-08:00"
}
```

## Security Notes

- 🔒 **Keep credentials secure** - The client secret is shown only once
- 🔒 **Do not share credentials** - Store them securely
- 🔒 **Rotate secrets regularly** - Client secrets expire after 24 months
- 🔒 **Review permissions** - Ensure only necessary permissions are granted
- 🔒 **Monitor application usage** - Review in Azure Portal regularly

## Support

For issues or questions:
- Review the Barracuda XDR documentation
- Check Microsoft 365 audit log requirements
- Verify Entra ID app registration permissions
- Review PowerShell module documentation

## License

This script is provided as-is for automating Barracuda XDR setup.
