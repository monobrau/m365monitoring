# Adding Office 365 Management API Permissions - Quick Guide

## Current Status
✅ Microsoft Graph permissions are configured:
- User.EnableDisableAccount.All
- User.ReadWrite.All

❌ Missing Office 365 Management API permissions:
- ActivityFeed.Read
- ActivityFeed.ReadDlp
- ServiceHealth.Read

## Steps to Add Missing Permissions

1. **Click "+ Add a permission"** (blue button at the top)

2. **Select "APIs my organization uses"** tab

3. **Search for "Office 365 Management API"**
   - Type in the search box: `Office 365 Management API`
   - Click on the result when it appears

4. **Select "Application permissions"** (NOT Delegated permissions)

5. **Check these 3 permissions:**
   - ✅ ActivityFeed.Read
   - ✅ ActivityFeed.ReadDlp  
   - ✅ ServiceHealth.Read

6. **Click "Add permissions"** button at the bottom

7. **Grant Admin Consent:**
   - Click "Grant admin consent for [Your Domain]"
   - Confirm the action
   - Wait for all permissions to show "Granted for [Your Domain]"

## Verification

After adding, you should see:
- **Microsoft Graph (2)** - already there ✅
- **Office 365 Management API (3)** - newly added ✅
  - ActivityFeed.Read
  - ActivityFeed.ReadDlp
  - ServiceHealth.Read

## Direct Link

If you need to return to this page:
```
https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/YOUR_APP_ID
```

## Notes

- These are **Application permissions** (not Delegated)
- **Admin consent is required** for all three permissions
- After granting consent, it may take a few minutes to propagate
- The script will detect these permissions once they're added and consent is granted
