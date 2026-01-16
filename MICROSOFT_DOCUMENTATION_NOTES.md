# Microsoft Official Documentation Research Notes

## Office 365 Management API Permissions

### Key Findings:
- **ActivityFeed.Read**, **ActivityFeed.ReadDlp**, and **ServiceHealth.Read** are valid Office 365 Management API permissions
- These permissions are **Application permissions** (not delegated)
- They **always require admin consent** - cannot be granted by regular users
- These permissions may not always be exposed as AppRoles in the service principal
- **Must be added manually in Azure Portal** if not found programmatically

### Official Documentation:
- [Get started with Office 365 Management APIs](https://learn.microsoft.com/en-us/office/office-365-management-api/get-started-with-office-365-management-apis)
- [Office 365 Management APIs overview](https://learn.microsoft.com/en-us/office/office-365-management-api/office-365-management-apis-overview)

### Required Steps:
1. Register app in Microsoft Entra ID
2. Add Application permissions from Office 365 Management API
3. **Tenant admin must grant consent via browser-based flow**
4. Enable unified audit logging (required before APIs return data)

## Admin Consent Granting

### Key Findings:
- **Even Global Admins** may need to grant consent manually via Azure Portal
- Programmatic consent granting is **restricted** and may not work in all tenants
- Admin consent can be granted via:
  1. Azure Portal (most reliable method)
  2. Admin Consent URL (browser-based)
  3. Microsoft Graph API (may be restricted)

### Admin Consent URL Pattern:
```
https://login.microsoftonline.com/{tenant-id}/adminconsent
  ?client_id={appId}
  &scope={space-separated-scopes}
  &redirect_uri={redirectURI}
```

### Official Documentation:
- [Grant tenant-wide admin consent to an application](https://learn.microsoft.com/en-us/azure/active-directory/manage-apps/grant-admin-consent)
- [Microsoft Graph permissions overview](https://learn.microsoft.com/en-us/graph/permissions-overview)

### Required Scopes for Programmatic Consent:
- `Application.ReadWrite.All`
- `AppRoleAssignment.ReadWrite.All`
- `DelegatedPermissionGrant.ReadWrite.All`

**Note**: Even with these scopes, programmatic consent may be restricted by tenant policies.

## Exchange Online PowerShell 7 Connection

### Key Findings:
- **PowerShell 7 + Exchange Online Management** can have compatibility issues
- Null reference errors are common with WAM (Web Account Manager) authentication
- **Interactive browser authentication** is the recommended method
- Connection may succeed even if return value is null - verify by running commands

### Official Documentation:
- [Connect to Exchange Online PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell)
- Exchange Online Management module documentation

### Best Practices:
1. Use interactive browser authentication (`Connect-ExchangeOnline` without `-UseDeviceAuthentication`)
2. Don't rely on return value - verify connection by running commands
3. Wait 3-5 seconds after connection before running commands
4. Clear existing sessions before connecting

## Unified Audit Logging

### Key Findings:
- **Must be enabled** before Office 365 Management APIs return audit data
- Enabled via: `Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true`
- Can be enabled via:
  1. Microsoft Purview Compliance Portal (UI)
  2. Exchange Online PowerShell
  3. Security & Compliance Center

### Verification:
```powershell
Get-AdminAuditLogConfig | Format-List UnifiedAuditLogIngestionEnabled
```

### Official Documentation:
- [Enable or disable audit log search](https://learn.microsoft.com/en-us/purview/audit-log-enable-disable)

## Microsoft Graph Permissions

### Key Findings:
- Permissions are either **Delegated** (user context) or **Application** (app-only)
- Application permissions **always require admin consent**
- Permission names and IDs are documented in Microsoft Graph Permissions Reference

### Official Documentation:
- [Microsoft Graph permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [Microsoft Graph permissions overview](https://learn.microsoft.com/en-us/graph/permissions-overview)

## Recommendations for Script

### Office 365 Management API Permissions:
1. **Try to add programmatically** - but expect it may fail
2. **Provide clear manual instructions** if programmatic addition fails
3. **Use admin consent URL** as fallback for consent granting

### Admin Consent:
1. **Try programmatic consent** - but expect it may be restricted
2. **Always provide manual fallback** with direct Azure Portal link
3. **Consider using admin consent URL** for better user experience

### Exchange Online:
1. **Suppress null reference errors** during connection
2. **Verify connection by running commands** not checking return value
3. **Add longer wait times** for PowerShell 7 compatibility
4. **Clear sessions** before connecting

### Audit Logging:
1. **Always verify** if audit logging is enabled before trying to enable
2. **Provide clear error messages** if enablement fails
3. **Give manual instructions** as fallback
