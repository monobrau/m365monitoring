# Git Repository Setup

## Repository Initialized

The git repository has been initialized in `c:\git\m365monitoring`.

## To Push to Remote Repository

### Option 1: GitHub (Recommended)

1. Create a new repository on GitHub named `m365monitoring`
2. Then run:
   ```powershell
   git remote add origin https://github.com/YOUR_USERNAME/m365monitoring.git
   git branch -M main
   git push -u origin main
   ```

### Option 2: GitHub via SSH

```powershell
git remote add origin git@github.com:YOUR_USERNAME/m365monitoring.git
git branch -M main
git push -u origin main
```

### Option 3: Azure DevOps / GitLab / Other

Replace the URL with your repository URL:
```powershell
git remote add origin YOUR_REPO_URL
git branch -M main
git push -u origin main
```

## Current Status

- ✅ Git repository initialized
- ✅ All files committed
- ✅ Python files removed (none found)
- ⏳ Waiting for remote repository URL

## Files Included

- `Setup-BarracudaXDR.ps1` - Main PowerShell GUI setup script
- `Test-O365Permissions.ps1` - Test script for Office 365 Management API permissions
- `enable_audit_logging.ps1` - Standalone audit logging script
- `README.md` - Main documentation
- `README-PowerShell.md` - PowerShell-specific documentation
- `QUICKSTART.md` - Quick start guide
- `ADD_O365_PERMISSIONS.md` - Manual permission addition guide
- `MICROSOFT_DOCUMENTATION_NOTES.md` - Research notes
- `.gitignore` - Git ignore rules
