# Enable Audit Logging for Microsoft 365
# This script enables audit logging for all mailboxes via Exchange Online PowerShell

Write-Host "=" -NoNewline
Write-Host ("=" * 69) -ForegroundColor Cyan
Write-Host "Microsoft 365 Audit Logging Setup" -ForegroundColor Cyan
Write-Host "=" -NoNewline
Write-Host ("=" * 69) -ForegroundColor Cyan
Write-Host ""

# Check if ExchangeOnlineManagement module is installed
$moduleName = "ExchangeOnlineManagement"
$moduleInstalled = Get-Module -ListAvailable -Name $moduleName

if (-not $moduleInstalled) {
    Write-Host "Exchange Online Management module not found." -ForegroundColor Yellow
    Write-Host "Installing ExchangeOnlineManagement module..." -ForegroundColor Yellow
    
    try {
        Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber
        Write-Host "✓ Module installed successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Failed to install module: $_" -ForegroundColor Red
        Write-Host "Please install manually: Install-Module -Name ExchangeOnlineManagement" -ForegroundColor Yellow
        exit 1
    }
}

# Import the module
try {
    Import-Module $moduleName -ErrorAction Stop
    Write-Host "✓ Module imported successfully" -ForegroundColor Green
}
catch {
    Write-Host "✗ Failed to import module: $_" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
Write-Host "A browser window will open for authentication." -ForegroundColor Yellow
Write-Host ""

try {
    Connect-ExchangeOnline -ShowBanner:$false
    
    Write-Host "✓ Connected to Exchange Online" -ForegroundColor Green
    Write-Host ""
    
    # Check current audit log configuration
    Write-Host "Checking current audit log configuration..." -ForegroundColor Yellow
    $currentConfig = Get-AdminAuditLogConfig
    
    if ($currentConfig.UnifiedAuditLogIngestionEnabled) {
        Write-Host "✓ Audit logging is already enabled" -ForegroundColor Green
        Write-Host ""
        Write-Host "Current configuration:" -ForegroundColor Cyan
        Write-Host "  Unified Audit Log Ingestion: Enabled" -ForegroundColor Green
        Write-Host "  Admin Audit Log Enabled: $($currentConfig.AdminAuditLogEnabled)" -ForegroundColor Cyan
    }
    else {
        Write-Host "Audit logging is not enabled. Enabling now..." -ForegroundColor Yellow
        Write-Host ""
        
        try {
            Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true
            
            Write-Host "✓ Audit logging enabled successfully!" -ForegroundColor Green
            Write-Host ""
            Write-Host "Important Notes:" -ForegroundColor Yellow
            Write-Host "  - It may take up to 60 minutes for changes to take effect" -ForegroundColor Yellow
            Write-Host "  - Full audit log data may take several hours to become available" -ForegroundColor Yellow
            Write-Host "  - This integration may not work immediately after enabling" -ForegroundColor Yellow
        }
        catch {
            Write-Host "✗ Failed to enable audit logging: $_" -ForegroundColor Red
            Write-Host ""
            Write-Host "Please try enabling manually:" -ForegroundColor Yellow
            Write-Host "  1. Navigate to https://portal.office.com" -ForegroundColor Yellow
            Write-Host "  2. Go to Admin center > Compliance > Audit" -ForegroundColor Yellow
            Write-Host "  3. Click 'Turn on auditing'" -ForegroundColor Yellow
            exit 1
        }
    }
    
    Write-Host ""
    Write-Host "Verifying configuration..." -ForegroundColor Yellow
    $verifyConfig = Get-AdminAuditLogConfig
    
    if ($verifyConfig.UnifiedAuditLogIngestionEnabled) {
        Write-Host "✓ Verification successful - Audit logging is enabled" -ForegroundColor Green
    }
    else {
        Write-Host "⚠ Warning: Audit logging may not be fully enabled yet" -ForegroundColor Yellow
        Write-Host "  Please wait a few minutes and verify in the Admin Center" -ForegroundColor Yellow
    }
    
    Write-Host ""
    Write-Host "=" -NoNewline
    Write-Host ("=" * 69) -ForegroundColor Cyan
    Write-Host "Setup completed successfully!" -ForegroundColor Green
    Write-Host "=" -NoNewline
    Write-Host ("=" * 69) -ForegroundColor Cyan
    
    # Disconnect
    Write-Host ""
    Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "✓ Disconnected" -ForegroundColor Green
}
catch {
    Write-Host "✗ Error: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Troubleshooting:" -ForegroundColor Yellow
    Write-Host "  - Ensure you have Exchange Administrator or Global Administrator role" -ForegroundColor Yellow
    Write-Host "  - Verify your account has the Audit Logs role in Exchange Online" -ForegroundColor Yellow
    Write-Host "  - Check your internet connection" -ForegroundColor Yellow
    exit 1
}
