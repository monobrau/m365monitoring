# Quick Start Guide

## Overview

This PowerShell GUI tool helps you set up Barracuda XDR monitoring for Microsoft 365 by automating the Entra ID app registration and configuration process.

## Prerequisites Checklist

- [ ] Microsoft 365/Azure Global Administrator or Application Administrator role
- [ ] Windows PowerShell 5.1+ or PowerShell 7+ installed
- [ ] Administrator rights (recommended)
- [ ] Internet connection

## Step-by-Step Setup

### 1. Set Execution Policy (if needed)

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### 2. Run the PowerShell GUI Script

```powershell
cd c:\git\m365monitoring
.\Setup-BarracudaXDR.ps1
```

**What happens:**
- GUI window opens
- Click **"Start Setup"** button
- Browser opens for Entra ID authentication (select your admin account)
- Script automatically:
  - Installs required PowerShell modules (if needed)
  - Creates "SKOUTCYBERSECURITY" application in Entra ID
  - Configures API permissions
  - Generates client secret (24-month expiry)
- Credentials appear in the GUI

**Important:** Copy the client secret immediately - it's shown only once!

### 3. Enable Audit Logging

Click the **"Enable Audit Logging"** button in the GUI, or run manually:

```powershell
.\enable_audit_logging.ps1
```

**What happens:**
- Installs Exchange Online Management module (if needed)
- Connects to Exchange Online
- Enables unified audit log ingestion
- Verifies configuration

**Alternative:** Enable via Admin Center at https://portal.office.com > Compliance > Audit

### 4. Grant Admin Consent (Manual Step)

1. Go to https://portal.azure.com
2. Navigate to: **Azure Active Directory** > **App registrations**
3. Find **SKOUTCYBERSECURITY** application
4. Click **API permissions**
5. Click **Grant admin consent for [Your Domain]**
6. Verify all permissions show "Granted for [Your Domain]"

The GUI will display a direct link to the Azure Portal page.

### 5. Configure Barracuda XDR Dashboard

1. Log in to Barracuda XDR Dashboard
2. Go to: **Administration** > **Integrations**
3. Click **Setup** on the Microsoft 365 card
4. Enter credentials from the GUI (or use "Copy All Credentials" button):
   - **Application ID**
   - **Directory (Tenant) ID**
   - **Client Secret**
5. Click **Test** to verify connection
6. Click **Save**

## GUI Features

### Setup Status Panel
- Real-time progress updates
- Status messages with timestamps
- Color-coded success/error messages

### Application Credentials Panel
- **Application ID** - Displayed automatically after setup
- **Tenant ID** - Displayed automatically after setup
- **Client Secret** - Displayed automatically after setup
- **Show/Hide Secret** - Toggle secret visibility
- **Copy All Credentials** - Copy to clipboard
- **Save to File** - Save credentials to JSON file

### Action Buttons
- **Start Setup** - Runs the complete automated setup
- **Enable Audit Logging** - One-click audit logging enablement
- **Test Connection** - Verify application exists in Entra ID
- **Close** - Exit the application

## Expected Output

After clicking "Start Setup", you should see status messages like:

```
[15:30:00] ========================================
[15:30:00] Barracuda XDR Microsoft 365 Setup
[15:30:00] ========================================
[15:30:01] Checking required PowerShell modules...
[15:30:02] ✓ Microsoft.Graph is installed
[15:30:03] Connecting to Microsoft Graph...
[15:30:05] ✓ Connected as: admin@yourdomain.com
[15:30:05] ✓ Tenant ID: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
[15:30:06] Checking for existing 'SKOUTCYBERSECURITY' application...
[15:30:07] ✓ No existing application found
[15:30:08] Registering new application 'SKOUTCYBERSECURITY'...
[15:30:10] ✓ Application registered successfully
[15:30:11] Configuring API permissions...
[15:30:12]   Adding Office 365 Management API permissions...
[15:30:13]     ✓ Added ActivityFeed.Read
[15:30:13]     ✓ Added ActivityFeed.ReadDlp
[15:30:14]     ✓ Added ServiceHealth.Read
[15:30:15]   Adding Microsoft Graph API permissions...
[15:30:16]     ✓ Added User.EnableDisableAccount.All
[15:30:17]     ✓ Added User.ReadWrite.All
[15:30:18] Creating client secret...
[15:30:20] ✓ Client secret created successfully
[15:30:20] ⚠ IMPORTANT: Save this secret now - it won't be shown again!
[15:30:21] ========================================
[15:30:21] Setup completed successfully!
[15:30:21] ========================================
```

Credentials will appear in the "Application Credentials" panel.

## Troubleshooting

### "Module not found" errors
- The script will attempt to install required modules automatically
- If installation fails, run manually:
  ```powershell
  Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
  Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
  ```

### "Authentication failed"
- Ensure you're using an account with Global Administrator or Application Administrator role
- Check that MFA is properly configured
- Verify internet connectivity

### "Application registration failed"
- Verify you have Application Administrator or Global Administrator role
- Check that app registrations are allowed in your tenant

### "Permission configuration failed"
- Some permissions may need to be added manually in Azure Portal
- Ensure you grant admin consent after the script runs
- Verify Office 365 Management API is available in your tenant

### "Audit logging script fails"
- Ensure you have Exchange Administrator role
- Verify Exchange Online Management module installed: `Get-Module -ListAvailable ExchangeOnlineManagement`
- Try manual enablement via Admin Center

### "Barracuda XDR test fails"
- Wait 24-48 hours after enabling audit logging
- Verify admin consent was granted for all permissions
- Check that credentials are correct (no extra spaces)
- Ensure audit logging has been active for at least a few hours

## Next Steps

After successful setup:

1. **Wait 24-48 hours** for audit logs to populate
2. **Monitor Barracuda XDR Dashboard** for initial data ingestion
3. **Review security alerts** as they appear
4. **Configure alerting rules** in Barracuda XDR as needed

## Security Reminders

- 🔒 Keep credentials secure - Use the "Save to File" button to store securely
- 🔒 Do not commit credentials to version control
- 🔒 Rotate client secrets before expiry (24 months)
- 🔒 Review API permissions regularly
- 🔒 Monitor application usage in Azure Portal

## Support

- Review `README.md` for detailed documentation
- Review `README-PowerShell.md` for PowerShell-specific details
- Check Barracuda XDR documentation
- Verify Microsoft 365 audit log requirements
- Review Entra ID app registration best practices
