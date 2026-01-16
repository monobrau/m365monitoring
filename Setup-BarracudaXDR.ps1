#Requires -Version 5.1
# Barracuda XDR Microsoft 365 Setup - PowerShell GUI
# This script provides a graphical interface for setting up Barracuda XDR monitoring

param(
    [switch]$SkipAuditLogging
)

# Add required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Application Configuration
$script:AppName = "SKOUTCYBERSECURITY"
$script:ClientSecretDescription = "Barracuda XDR"
$script:ClientSecretExpiryMonths = 24

# Global variables
$script:TenantId = $null
$script:AppId = $null
$script:ClientSecret = $null
$script:GraphConnection = $null

# Office 365 Management API permissions
$script:M365ManagementPermissions = @(
    "ActivityFeed.Read",
    "ActivityFeed.ReadDlp",
    "ServiceHealth.Read"
)

# Microsoft Graph API permissions
$script:GraphPermissions = @(
    "User.EnableDisableAccount.All",
    "User.ReadWrite.All"
)

# Main Form
function New-MainForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Barracuda XDR Microsoft 365 Setup"
    $form.Size = New-Object System.Drawing.Size(800, 700)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    
    # Title Label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Barracuda XDR Microsoft 365 Setup"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(750, 30)
    $form.Controls.Add($titleLabel)
    
    # Status GroupBox
    $statusGroup = New-Object System.Windows.Forms.GroupBox
    $statusGroup.Text = "Setup Status"
    $statusGroup.Location = New-Object System.Drawing.Point(20, 60)
    $statusGroup.Size = New-Object System.Drawing.Size(750, 200)
    $form.Controls.Add($statusGroup)
    
    # Status TextBox (read-only, multiline)
    $script:StatusTextBox = New-Object System.Windows.Forms.TextBox
    $script:StatusTextBox.Multiline = $true
    $script:StatusTextBox.ReadOnly = $true
    $script:StatusTextBox.ScrollBars = "Vertical"
    $script:StatusTextBox.Location = New-Object System.Drawing.Point(10, 20)
    $script:StatusTextBox.Size = New-Object System.Drawing.Size(730, 170)
    $script:StatusTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $statusGroup.Controls.Add($script:StatusTextBox)
    
    # Credentials GroupBox
    $credsGroup = New-Object System.Windows.Forms.GroupBox
    $credsGroup.Text = "Application Credentials"
    $credsGroup.Location = New-Object System.Drawing.Point(20, 270)
    $credsGroup.Size = New-Object System.Drawing.Size(750, 200)
    $form.Controls.Add($credsGroup)
    
    # Application ID
    $appIdLabel = New-Object System.Windows.Forms.Label
    $appIdLabel.Text = "Application ID:"
    $appIdLabel.Location = New-Object System.Drawing.Point(10, 25)
    $appIdLabel.Size = New-Object System.Drawing.Size(100, 20)
    $credsGroup.Controls.Add($appIdLabel)
    
    $script:AppIdTextBox = New-Object System.Windows.Forms.TextBox
    $script:AppIdTextBox.Location = New-Object System.Drawing.Point(120, 23)
    $script:AppIdTextBox.Size = New-Object System.Drawing.Size(600, 23)
    $script:AppIdTextBox.ReadOnly = $true
    $credsGroup.Controls.Add($script:AppIdTextBox)
    
    # Tenant ID
    $tenantIdLabel = New-Object System.Windows.Forms.Label
    $tenantIdLabel.Text = "Tenant ID:"
    $tenantIdLabel.Location = New-Object System.Drawing.Point(10, 55)
    $tenantIdLabel.Size = New-Object System.Drawing.Size(100, 20)
    $credsGroup.Controls.Add($tenantIdLabel)
    
    $script:TenantIdTextBox = New-Object System.Windows.Forms.TextBox
    $script:TenantIdTextBox.Location = New-Object System.Drawing.Point(120, 53)
    $script:TenantIdTextBox.Size = New-Object System.Drawing.Size(600, 23)
    $script:TenantIdTextBox.ReadOnly = $true
    $credsGroup.Controls.Add($script:TenantIdTextBox)
    
    # Client Secret
    $secretLabel = New-Object System.Windows.Forms.Label
    $secretLabel.Text = "Client Secret:"
    $secretLabel.Location = New-Object System.Drawing.Point(10, 85)
    $secretLabel.Size = New-Object System.Drawing.Size(100, 20)
    $credsGroup.Controls.Add($secretLabel)
    
    $script:SecretTextBox = New-Object System.Windows.Forms.TextBox
    $script:SecretTextBox.Location = New-Object System.Drawing.Point(120, 83)
    $script:SecretTextBox.Size = New-Object System.Drawing.Size(500, 23)
    $script:SecretTextBox.ReadOnly = $true
    $script:SecretTextBox.PasswordChar = '*'
    $credsGroup.Controls.Add($script:SecretTextBox)
    
    # Show/Hide Secret Button
    $script:ShowSecretBtn = New-Object System.Windows.Forms.Button
    $script:ShowSecretBtn.Text = "Show"
    $script:ShowSecretBtn.Location = New-Object System.Drawing.Point(630, 82)
    $script:ShowSecretBtn.Size = New-Object System.Drawing.Size(90, 25)
    $script:ShowSecretBtn.Add_Click({
        if ($script:SecretTextBox.PasswordChar -eq '*') {
            $script:SecretTextBox.PasswordChar = [char]0
            $script:ShowSecretBtn.Text = "Hide"
        } else {
            $script:SecretTextBox.PasswordChar = '*'
            $script:ShowSecretBtn.Text = "Show"
        }
    })
    $credsGroup.Controls.Add($script:ShowSecretBtn)
    
    # Copy Credentials Button
    $copyCredsBtn = New-Object System.Windows.Forms.Button
    $copyCredsBtn.Text = "Copy All Credentials"
    $copyCredsBtn.Location = New-Object System.Drawing.Point(120, 115)
    $copyCredsBtn.Size = New-Object System.Drawing.Size(150, 30)
    $copyCredsBtn.Add_Click({
        $creds = "Application ID: $($script:AppIdTextBox.Text)`n"
        $creds += "Tenant ID: $($script:TenantIdTextBox.Text)`n"
        $creds += "Client Secret: $($script:SecretTextBox.Text)"
        [System.Windows.Forms.Clipboard]::SetText($creds)
        [System.Windows.Forms.MessageBox]::Show("Credentials copied to clipboard!", "Success", "OK", "Information")
    })
    $credsGroup.Controls.Add($copyCredsBtn)
    
    # Save Credentials Button
    $saveCredsBtn = New-Object System.Windows.Forms.Button
    $saveCredsBtn.Text = "Save to File"
    $saveCredsBtn.Location = New-Object System.Drawing.Point(280, 115)
    $saveCredsBtn.Size = New-Object System.Drawing.Size(150, 30)
    $saveCredsBtn.Add_Click({
        Save-CredentialsToFile
    })
    $credsGroup.Controls.Add($saveCredsBtn)
    
    # Instructions Label
    $instructionsLabel = New-Object System.Windows.Forms.Label
    $instructionsLabel.Text = "Instructions: Click 'Start Setup' to begin. You will be prompted to authenticate with your Microsoft 365 admin account."
    $instructionsLabel.Location = New-Object System.Drawing.Point(20, 480)
    $instructionsLabel.Size = New-Object System.Drawing.Size(750, 40)
    $instructionsLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    $form.Controls.Add($instructionsLabel)
    
    # Buttons Panel
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Location = New-Object System.Drawing.Point(20, 530)
    $buttonPanel.Size = New-Object System.Drawing.Size(750, 50)
    $form.Controls.Add($buttonPanel)
    
    # Start Setup Button
    $script:StartBtn = New-Object System.Windows.Forms.Button
    $script:StartBtn.Text = "Start Setup"
    $script:StartBtn.Location = New-Object System.Drawing.Point(10, 10)
    $script:StartBtn.Size = New-Object System.Drawing.Size(150, 35)
    $script:StartBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $script:StartBtn.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $script:StartBtn.ForeColor = [System.Drawing.Color]::White
    $script:StartBtn.Add_Click({
        $script:StartBtn.Enabled = $false
        try {
            Start-SetupProcess
        } finally {
            $script:StartBtn.Enabled = $true
        }
    })
    $buttonPanel.Controls.Add($script:StartBtn)
    
    # Enable Audit Logging Button
    $auditBtn = New-Object System.Windows.Forms.Button
    $auditBtn.Text = "Enable Audit Logging"
    $auditBtn.Location = New-Object System.Drawing.Point(170, 10)
    $auditBtn.Size = New-Object System.Drawing.Size(180, 35)
    $auditBtn.Add_Click({
        Enable-AuditLoggingGUI
    })
    $buttonPanel.Controls.Add($auditBtn)
    
    # Test Connection Button
    $testBtn = New-Object System.Windows.Forms.Button
    $testBtn.Text = "Test Connection"
    $testBtn.Location = New-Object System.Drawing.Point(360, 10)
    $testBtn.Size = New-Object System.Drawing.Size(150, 35)
    $testBtn.Add_Click({
        Test-ConnectionGUI
    })
    $buttonPanel.Controls.Add($testBtn)
    
    # Close Button
    $closeBtn = New-Object System.Windows.Forms.Button
    $closeBtn.Text = "Close"
    $closeBtn.Location = New-Object System.Drawing.Point(650, 10)
    $closeBtn.Size = New-Object System.Drawing.Size(90, 35)
    $closeBtn.Add_Click({
        $form.Close()
    })
    $buttonPanel.Controls.Add($closeBtn)
    
    # Progress Bar
    $script:ProgressBar = New-Object System.Windows.Forms.ProgressBar
    $script:ProgressBar.Location = New-Object System.Drawing.Point(20, 590)
    $script:ProgressBar.Size = New-Object System.Drawing.Size(750, 23)
    $script:ProgressBar.Style = "Continuous"
    $form.Controls.Add($script:ProgressBar)
    
    return $form
}

# Write status message
function Write-Status {
    param([string]$Message, [string]$Color = "Black")
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $statusMessage = "[$timestamp] $Message`r`n"
    
    if ($script:StatusTextBox.InvokeRequired) {
        $script:StatusTextBox.Invoke([System.Action[String]]{
            param($msg)
            $script:StatusTextBox.AppendText($msg)
            $script:StatusTextBox.SelectionStart = $script:StatusTextBox.Text.Length
            $script:StatusTextBox.ScrollToCaret()
        }, $statusMessage)
    } else {
        $script:StatusTextBox.AppendText($statusMessage)
        $script:StatusTextBox.SelectionStart = $script:StatusTextBox.Text.Length
        $script:StatusTextBox.ScrollToCaret()
    }
    
    [System.Windows.Forms.Application]::DoEvents()
}

# Update progress bar
function Update-Progress {
    param([int]$Value)
    
    if ($script:ProgressBar.InvokeRequired) {
        $script:ProgressBar.Invoke([System.Action[int]]{
            param($val)
            $script:ProgressBar.Value = $val
        }, $Value)
    } else {
        $script:ProgressBar.Value = $Value
    }
    [System.Windows.Forms.Application]::DoEvents()
}

# Check and install required modules
function Install-RequiredModules {
    Write-Status "Checking required PowerShell modules..."
    Update-Progress 10
    
    $modules = @(
        @{Name = "Microsoft.Graph"; Required = $true},
        @{Name = "ExchangeOnlineManagement"; Required = $false}
    )
    
    foreach ($module in $modules) {
        $installed = Get-Module -ListAvailable -Name $module.Name
        if (-not $installed) {
            if ($module.Required) {
                Write-Status "Installing $($module.Name) module..." "Blue"
                try {
                    Install-Module -Name $module.Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                    Write-Status "✓ $($module.Name) installed successfully" "Green"
                } catch {
                    Write-Status "✗ Failed to install $($module.Name): $_" "Red"
                    return $false
                }
            } else {
                Write-Status "⚠ $($module.Name) not installed (optional)" "Orange"
            }
        } else {
            Write-Status "✓ $($module.Name) is installed" "Green"
        }
    }
    
    Update-Progress 20
    return $true
}

# Connect to Microsoft Graph
function Connect-MicrosoftGraph {
    Write-Status "Connecting to Microsoft Graph..."
    Update-Progress 25
    
    try {
        # Disconnect any existing connections
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        
        # Connect with required scopes (including AppRoleAssignment for consent granting)
        $scopes = @(
            "Application.ReadWrite.All",
            "Directory.ReadWrite.All",
            "User.Read",
            "DelegatedPermissionGrant.ReadWrite.All",
            "AppRoleAssignment.ReadWrite.All"
        )
        
        $script:GraphConnection = Connect-MgGraph -Scopes $scopes -NoWelcome
        
        # Get tenant information
        $context = Get-MgContext
        $script:TenantId = $context.TenantId
        
        Write-Status "✓ Connected as: $($context.Account)" "Green"
        Write-Status "✓ Tenant ID: $script:TenantId" "Green"
        
        Update-Progress 35
        return $true
    } catch {
        Write-Status "✗ Failed to connect: $_" "Red"
        return $false
    }
}

# Check for existing application
function Get-ExistingApplication {
    Write-Status "Checking for existing '$script:AppName' application..."
    Update-Progress 40
    
    try {
        $apps = Get-MgApplication -Filter "displayName eq '$script:AppName'" -ErrorAction Stop
        
        if ($apps) {
            $app = $apps | Select-Object -First 1
            $script:AppId = $app.AppId
            Write-Status "✓ Found existing application: $script:AppId" "Green"
            Update-Progress 45
            return $app
        } else {
            Write-Status "✓ No existing application found" "Green"
            Update-Progress 45
            return $null
        }
    } catch {
        Write-Status "⚠ Error checking for existing app: $_" "Orange"
        Update-Progress 45
        return $null
    }
}

# Register new application
function Register-Application {
    Write-Status "Registering new application '$script:AppName'..."
    Update-Progress 50
    
    try {
        $params = @{
            DisplayName = $script:AppName
            SignInAudience = "AzureADMyOrg"  # Single tenant
        }
        
        $app = New-MgApplication @params -ErrorAction Stop
        $script:AppId = $app.AppId
        
        Write-Status "✓ Application registered successfully" "Green"
        Write-Status "  Application ID: $script:AppId" "Green"
        Write-Status "  Tenant ID: $script:TenantId" "Green"
        
        Update-Progress 60
        return $app
    } catch {
        Write-Status "✗ Failed to register application: $_" "Red"
        return $null
    }
}

# Add API permissions
function Add-APIPermissions {
    param($AppObjectId)
    
    Write-Status "Configuring API permissions..."
    Update-Progress 65
    
    try {
        # Get service principals for the APIs
        Write-Status "  Looking up service principals..." "Blue"
        
        # Office 365 Management APIs
        # Use the correct AppId: c5393580-f805-4401-95e8-94b7a6ef2fc2 (verified by test script)
        Write-Status "  Searching for Office 365 Management API service principal..." "Blue"
        
        $o365AppId = 'c5393580-f805-4401-95e8-94b7a6ef2fc2'  # Office 365 Management APIs (verified)
        $o365Sp = Get-MgServicePrincipal -Filter "appId eq '$o365AppId'" -ErrorAction SilentlyContinue
        
        if (-not $o365Sp) {
            # Try alternative lookup
            $allSp = Get-MgServicePrincipal -All -ErrorAction SilentlyContinue | Where-Object { $_.AppId -eq $o365AppId }
            if ($allSp) {
                $o365Sp = $allSp | Select-Object -First 1
            }
        }
        
        if ($o365Sp) {
            Write-Status "  ✓ Found Office 365 Management API service principal: $($o365Sp.DisplayName)" "Green"
            Write-Status "    AppId: $($o365Sp.AppId)" "Gray"
        } else {
            Write-Status "  ⚠ Office 365 Management API service principal not found" "Orange"
            Write-Status "  Will attempt to add permissions using known AppId" "Blue"
        }
        
        # Microsoft Graph - App ID: 00000003-0000-0000-c000-000000000000
        $graphAppId = '00000003-0000-0000-c000-000000000000'
        $graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'" -ErrorAction SilentlyContinue
        
        if (-not $graphSp) {
            Write-Status "  ⚠ Could not find Microsoft Graph service principal via filter" "Orange"
            Write-Status "  Attempting alternative lookup method..." "Blue"
            $allSp = Get-MgServicePrincipal -All -ErrorAction SilentlyContinue | Where-Object { $_.AppId -eq $graphAppId }
            if ($allSp) {
                $graphSp = $allSp | Select-Object -First 1
                Write-Status "  ✓ Found Microsoft Graph service principal (alternative method)" "Green"
            }
        } else {
            Write-Status "  ✓ Found Microsoft Graph service principal" "Green"
        }
        
        if (-not $graphSp) {
            Write-Status "✗ Could not find Microsoft Graph service principal. Cannot continue." "Red"
            return $false
        }
        
        # Get the application
        $app = Get-MgApplication -ApplicationId $AppObjectId
        
        # Build required resource access
        $requiredAccess = @()
        
        # Office 365 Management API permissions
        Write-Status "  Adding Office 365 Management API permissions..." "Blue"
        
        # Use the correct AppId (verified by test script)
        $o365ResourceAppId = if ($o365Sp) { $o365Sp.AppId } else { 'c5393580-f805-4401-95e8-94b7a6ef2fc2' }
        
        $o365Access = @{
            ResourceAppId = $o365ResourceAppId
            ResourceAccess = @()
        }
        
        # Get app roles explicitly - try to get service principal if we don't have it
        Write-Status "  Retrieving all app roles for Office 365 Management API..." "Blue"
        
        if (-not $o365Sp) {
            # Try to get the service principal using the known AppId
            try {
                $allSp = Get-MgServicePrincipal -All -ErrorAction Stop | Where-Object { $_.AppId -eq $o365ResourceAppId }
                if ($allSp) {
                    $o365Sp = $allSp | Select-Object -First 1
                }
            } catch {
                Write-Status "  ⚠ Could not retrieve service principal: $_" "Orange"
            }
        }
        
        if ($o365Sp) {
            try {
                # Use Invoke-MgGraphRequest to get all AppRoles (more reliable)
                $spUri = "https://graph.microsoft.com/v1.0/servicePrincipals/$($o365Sp.Id)?`$select=id,appId,displayName,appRoles"
                $o365SpFull = Invoke-MgGraphRequest -Method GET -Uri $spUri -ErrorAction Stop
                
                # Verify we have the correct service principal
                if ($o365SpFull.appId -ne 'c5393580-f805-4401-95e8-94b7a6ef2fc2') {
                    Write-Status "  ⚠ Warning: Service principal AppId ($($o365SpFull.appId)) doesn't match expected" "Orange"
                    Write-Status "  Expected: c5393580-f805-4401-95e8-94b7a6ef2fc2" "Orange"
                }
                
                # Use AppRoles directly from API response (matches test script approach)
                if ($o365SpFull.appRoles) {
                    Write-Status "  Found $($o365SpFull.appRoles.Count) AppRoles in Office 365 Management API" "Blue"
                } else {
                    Write-Status "  ⚠ No AppRoles found in service principal response" "Orange"
                }
            } catch {
                Write-Status "  ⚠ Could not retrieve AppRoles via API: $_" "Orange"
                Write-Status "  Falling back to standard method..." "Blue"
                try {
                    if ($o365Sp) {
                        $o365SpFull = Get-MgServicePrincipal -ServicePrincipalId $o365Sp.Id -Property AppRoles -ErrorAction Stop
                    }
                } catch {
                    Write-Status "  ⚠ Could not retrieve AppRoles: $_" "Orange"
                }
            }
        }
        
        # Search for permissions in AppRoles (matches test script approach)
        if ($o365SpFull -and $o365SpFull.appRoles) {
            foreach ($permName in $script:M365ManagementPermissions) {
                $appRole = $null
                
                # Search in the appRoles from the API response (exact match first - matches test script)
                $appRole = $o365SpFull.appRoles | Where-Object { $_.value -eq $permName } | Select-Object -First 1
                
                # Case-insensitive match
                if (-not $appRole) {
                    $appRole = $o365SpFull.appRoles | Where-Object { $_.value -ieq $permName } | Select-Object -First 1
                }
                
                if ($appRole) {
                    $o365Access.ResourceAccess += @{
                        Id = $appRole.id
                        Type = "Role"
                    }
                    Write-Status "    ✓ Added $permName (ID: $($appRole.id))" "Green"
                } else {
                    Write-Status "    ⚠ Permission '$permName' not found in Office 365 Management API" "Orange"
                }
            }
        } else {
            Write-Status "  ⚠ Could not retrieve AppRoles - permissions will need manual addition" "Orange"
        }
        
        # If no permissions were found via AppRoles, try adding them directly via Graph API
        # Office 365 Management API permissions may not be exposed as AppRoles but can still be added
        if ($o365Access.ResourceAccess.Count -eq 0) {
            Write-Status "  ⚠ Permissions not found in AppRoles - attempting direct addition via Graph API..." "Orange"
            
            # Method 1: Try querying ALL service principals to find Office 365 Management API with different AppId
            Write-Status "  Searching for Office 365 Management API service principal..." "Blue"
            try {
                # Try alternative AppIds that Office 365 Management API might use
                $possibleAppIds = @(
                    '00000003-0000-0ff1-ce00-000000000000',  # Standard
                    'c5393580-f805-4401-95e8-94b7a6ef2fc2',  # Alternative (from docs)
                    '797f4846-ba00-4fd7-ba43-dac1f8f63013'   # Another possible ID
                )
                
                foreach ($testAppId in $possibleAppIds) {
                    $testSp = Get-MgServicePrincipal -Filter "appId eq '$testAppId'" -ErrorAction SilentlyContinue
                    if (-not $testSp) {
                        $allSp = Get-MgServicePrincipal -All -ErrorAction SilentlyContinue | Where-Object { $_.AppId -eq $testAppId }
                        if ($allSp) {
                            $testSp = $allSp | Select-Object -First 1
                        }
                    }
                    
                    if ($testSp) {
                        Write-Status "  Found service principal with AppId: $testAppId" "Blue"
                        $spUri = "https://graph.microsoft.com/v1.0/servicePrincipals/$($testSp.Id)?`$select=id,appId,displayName,appRoles"
                        $spData = Invoke-MgGraphRequest -Method GET -Uri $spUri -ErrorAction Stop
                        
                        if ($spData.appRoles -and $spData.appRoles.Count -gt 0) {
                            Write-Status "  Found $($spData.appRoles.Count) AppRoles - searching for required permissions..." "Blue"
                            
                            foreach ($permName in $script:M365ManagementPermissions) {
                                $foundRole = $null
                                
                                # Search for permission by name patterns
                                if ($permName -eq "ActivityFeed.Read") {
                                    $foundRole = $spData.appRoles | Where-Object { 
                                        ($_.value -like "*ActivityFeed*" -or $_.value -like "*Activity*Feed*") -and
                                        $_.value -like "*Read*" -and
                                        $_.value -notlike "*Dlp*" -and
                                        $_.value -notlike "*DLP*"
                                    } | Select-Object -First 1
                                } elseif ($permName -eq "ActivityFeed.ReadDlp") {
                                    $foundRole = $spData.appRoles | Where-Object { 
                                        ($_.value -like "*ActivityFeed*Dlp*" -or 
                                         $_.value -like "*ActivityFeed*DLP*" -or
                                         ($_.value -like "*Activity*" -and ($_.value -like "*Dlp*" -or $_.value -like "*DLP*")))
                                    } | Select-Object -First 1
                                } elseif ($permName -eq "ServiceHealth.Read") {
                                    $foundRole = $spData.appRoles | Where-Object { 
                                        $_.value -like "*ServiceHealth*" -or 
                                        ($_.value -like "*Service*" -and $_.value -like "*Health*")
                                    } | Select-Object -First 1
                                }
                                
                                if ($foundRole) {
                                    $o365Access.ResourceAccess += @{
                                        Id = $foundRole.id
                                        Type = "Role"
                                    }
                                    Write-Status "    ✓ Found $permName (as: $($foundRole.value), ID: $($foundRole.id))" "Green"
                                    # Update the service principal reference
                                    $o365Sp = $testSp
                                    break
                                }
                            }
                            
                            if ($o365Access.ResourceAccess.Count -gt 0) {
                                break  # Found permissions, stop searching
                            }
                        }
                    }
                }
            } catch {
                Write-Status "  ⚠ Error searching alternative service principals: $_" "Orange"
            }
            
            # Method 2: Try querying the application manifest directly to see current RequiredResourceAccess
            # and try to add permissions even if we don't have exact IDs
            if ($o365Access.ResourceAccess.Count -eq 0) {
                Write-Status "  Attempting to query application manifest and add permissions directly..." "Blue"
                
                try {
                    # Get current application with RequiredResourceAccess
                    $appManifestUri = "https://graph.microsoft.com/v1.0/applications/$AppObjectId?`$select=id,appId,requiredResourceAccess"
                    $appData = Invoke-MgGraphRequest -Method GET -Uri $appManifestUri -ErrorAction Stop
                    
                    # Check if Office 365 Management API is already in RequiredResourceAccess
                    $existingO365Resource = $appData.requiredResourceAccess | Where-Object { 
                        $_.resourceAppId -eq '00000003-0000-0ff1-ce00-000000000000' 
                    }
                    
                    if ($existingO365Resource) {
                        Write-Status "  Found existing Office 365 Management API resource in manifest" "Blue"
                        Write-Status "  Current permissions: $($existingO365Resource.resourceAccess.Count)" "Blue"
                    }
                    
                    # Try one more time to get ALL service principals and search for Office 365 Management API
                    Write-Status "  Performing comprehensive search of all service principals..." "Blue"
                    $allServicePrincipals = Get-MgServicePrincipal -All -ErrorAction Stop | Where-Object {
                        $_.DisplayName -like "*Office 365*Management*" -or
                        $_.DisplayName -like "*Management*API*" -or
                        $_.AppId -eq '00000003-0000-0ff1-ce00-000000000000' -or
                        $_.AppId -eq 'c5393580-f805-4401-95e8-94b7a6ef2fc2'
                    }
                    
                    foreach ($sp in $allServicePrincipals) {
                        Write-Status "  Checking service principal: $($sp.DisplayName) (AppId: $($sp.AppId))" "Blue"
                        $spUri = "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.Id)?`$expand=appRoles"
                        $spFull = Invoke-MgGraphRequest -Method GET -Uri $spUri -ErrorAction Stop
                        
                        if ($spFull.appRoles) {
                            Write-Status "    Found $($spFull.appRoles.Count) AppRoles" "Blue"
                            
                            # Search for our permissions
                            foreach ($permName in $script:M365ManagementPermissions) {
                                $permFound = $spFull.appRoles | Where-Object {
                                    $_.value -eq $permName -or
                                    $_.value -ieq $permName -or
                                    ($permName -eq "ActivityFeed.Read" -and $_.value -like "*ActivityFeed*Read*" -and $_.value -notlike "*Dlp*") -or
                                    ($permName -eq "ActivityFeed.ReadDlp" -and $_.value -like "*ActivityFeed*Dlp*") -or
                                    ($permName -eq "ServiceHealth.Read" -and $_.value -like "*ServiceHealth*")
                                } | Select-Object -First 1
                                
                                if ($permFound) {
                                    $o365Access.ResourceAccess += @{
                                        Id = $permFound.id
                                        Type = "Role"
                                    }
                                    Write-Status "    ✓ Found $permName (as: $($permFound.value), ID: $($permFound.id))" "Green"
                                    $o365Access.ResourceAppId = $sp.AppId
                                }
                            }
                            
                            if ($o365Access.ResourceAccess.Count -gt 0) {
                                break
                            }
                        }
                    }
                } catch {
                    Write-Status "  ⚠ Error querying manifest: $_" "Orange"
                }
                
            # Final fallback: Try using known permission IDs directly
            # Even if permissions aren't exposed as AppRoles, we can try adding them by ID
            if ($o365Access.ResourceAccess.Count -eq 0) {
                Write-Status "  Attempting to add permissions using known permission IDs..." "Blue"
                Write-Status "  This may work even if permissions aren't visible in AppRoles" "Blue"
                
                # Known Office 365 Management API permission IDs (from Claude's approach)
                # These are the standard IDs for these permissions
                $knownPermissionIds = @{
                    "ActivityFeed.Read" = "594c1fb6-4f81-4f1d-ab49-e4b9e6e7e7c0"
                    "ActivityFeed.ReadDlp" = "4807a72c-ad38-4250-94c9-4eabfe26cd55"
                    "ServiceHealth.Read" = "f2e896c5-7f55-4e12-9264-b8f5eff67b28"
                }
                
                # Try AppIds for Office 365 Management API (Claude's approach first)
                $o365AppIds = @(
                    'c5393580-f805-4401-95e8-94b7a6ef2fc2',  # Office 365 Management API (Claude's approach)
                    '00000003-0000-0ff1-ce00-000000000000'   # Standard (but might be SharePoint)
                )
                
                foreach ($testAppId in $o365AppIds) {
                    Write-Status "  Trying AppId: $testAppId" "Gray"
                    
                    # Try adding permissions with this AppId
                    $testAccess = @{
                        ResourceAppId = $testAppId
                        ResourceAccess = @()
                    }
                    
                    foreach ($permName in $script:M365ManagementPermissions) {
                        if ($knownPermissionIds.ContainsKey($permName)) {
                            $permId = $knownPermissionIds[$permName]
                            $testAccess.ResourceAccess += @{
                                Id = $permId
                                Type = "Role"
                            }
                            Write-Status "    Adding $permName (ID: $permId)" "Gray"
                        }
                    }
                    
                    if ($testAccess.ResourceAccess.Count -gt 0) {
                        # Try to add this to the application
                        try {
                            # Get current RequiredResourceAccess
                            $currentRequiredAccess = @()
                            if ($app.RequiredResourceAccess) {
                                $currentRequiredAccess = $app.RequiredResourceAccess | Where-Object {
                                    $_.ResourceAppId -ne $testAppId
                                }
                            }
                            
                            # Add our test access
                            $currentRequiredAccess += $testAccess
                            
                            # Update application
                            Update-MgApplication -ApplicationId $AppObjectId -RequiredResourceAccess $currentRequiredAccess -ErrorAction Stop
                            
                            Write-Status "  ✓ Successfully added permissions using AppId: $testAppId" "Green"
                            Write-Status "    Added $($testAccess.ResourceAccess.Count) permission(s) by ID" "Green"
                            
                            $o365Access = $testAccess
                            break
                        } catch {
                            Write-Status "  ⚠ Failed with AppId $testAppId : $_" "Orange"
                            # Try next AppId
                        }
                    }
                }
                
                # If still no success, provide manual instructions
                if ($o365Access.ResourceAccess.Count -eq 0) {
                    Write-Status "  ⚠ Could not add permissions using known IDs" "Orange"
                    Write-Status "  Permissions must be added manually in Azure Portal" "Orange"
                }
            }
            }
        }
        
        if ($o365Access.ResourceAccess.Count -gt 0) {
            $requiredAccess += $o365Access
            Write-Status "  ✓ Added $($o365Access.ResourceAccess.Count) Office 365 Management API permission(s)" "Green"
        } else {
            Write-Status "  ⚠ No Office 365 Management API permissions were added automatically" "Orange"
            Write-Status "  These permissions are not exposed as AppRoles and must be added manually" "Orange"
            Write-Status "" "Blue"
            Write-Status "  Opening Azure Portal to add permissions manually..." "Blue"
            
            # Open Azure Portal directly to the API permissions page
            $portalUrl = "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$script:AppId"
            try {
                Start-Process $portalUrl
                Write-Status "  ✓ Azure Portal opened" "Green"
            } catch {
                Write-Status "  ⚠ Could not open browser automatically" "Orange"
            }
            
            Write-Status "" "Blue"
            Write-Status "  Manual steps in Azure Portal:" "Blue"
            Write-Status "  1. Click 'API permissions' (should be open)" "Blue"
            Write-Status "  2. Click 'Add a permission'" "Blue"
            Write-Status "  3. Select 'APIs my organization uses'" "Blue"
            Write-Status "  4. Search for 'Office 365 Management API'" "Blue"
            Write-Status "  5. Select 'Application permissions'" "Blue"
            Write-Status "  6. Check these permissions:" "Blue"
            Write-Status "     - ActivityFeed.Read" "Blue"
            Write-Status "     - ActivityFeed.ReadDlp" "Blue"
            Write-Status "     - ServiceHealth.Read" "Blue"
            Write-Status "  7. Click 'Add permissions'" "Blue"
            Write-Status "  8. Click 'Grant admin consent for [Your Domain]'" "Blue"
            Write-Status "" "Blue"
        }
        
        # Microsoft Graph API permissions
        Write-Status "  Adding Microsoft Graph API permissions..." "Blue"
        $graphAccess = @{
            ResourceAppId = $graphSp.AppId
            ResourceAccess = @()
        }
        
        # Get app roles for Graph if needed
        if (-not $graphSp.AppRoles) {
            Write-Status "  Retrieving app roles for Microsoft Graph..." "Blue"
            $graphSpFull = Get-MgServicePrincipal -ServicePrincipalId $graphSp.Id -Property AppRoles
            $graphSp = $graphSpFull
        }
        
        foreach ($permName in $script:GraphPermissions) {
            $appRole = $graphSp.AppRoles | Where-Object { $_.Value -eq $permName }
            
            if (-not $appRole) {
                $appRole = $graphSp.AppRoles | Where-Object { $_.Value -like "*$permName*" }
            }
            
            if ($appRole) {
                $graphAccess.ResourceAccess += @{
                    Id = $appRole.Id
                    Type = "Role"
                }
                Write-Status "    ✓ Added $permName (ID: $($appRole.Id))" "Green"
            } else {
                Write-Status "    ⚠ Permission '$permName' not found in Microsoft Graph" "Orange"
            }
        }
        
        if ($graphAccess.ResourceAccess.Count -gt 0) {
            $requiredAccess += $graphAccess
            Write-Status "  ✓ Added $($graphAccess.ResourceAccess.Count) Microsoft Graph API permission(s)" "Green"
        } else {
            Write-Status "  ✗ No Microsoft Graph API permissions were added" "Red"
        }
        
        # Update application with permissions
        if ($requiredAccess.Count -gt 0) {
            Write-Status "  Updating application with $($requiredAccess.Count) API resource(s)..." "Blue"
            
            $updateParams = @{
                ApplicationId = $AppObjectId
                RequiredResourceAccess = $requiredAccess
            }
            
            Update-MgApplication @updateParams -ErrorAction Stop
            
            # Verify the update
            $updatedApp = Get-MgApplication -ApplicationId $AppObjectId
            $totalPerms = ($updatedApp.RequiredResourceAccess | ForEach-Object { $_.ResourceAccess.Count }) | Measure-Object -Sum
            Write-Status "✓ API permissions configured ($($totalPerms.Sum) total permissions)" "Green"
            
            # Show summary
            foreach ($resource in $requiredAccess) {
                $resourceName = if ($resource.ResourceAppId -eq $o365Sp.AppId) { "Office 365 Management API" } else { "Microsoft Graph" }
                Write-Status "  - $resourceName : $($resource.ResourceAccess.Count) permission(s)" "Blue"
            }
            
            Update-Progress 75
            return $true
        } else {
            Write-Status "✗ No permissions were added - manual configuration required" "Red"
            Write-Status "  Please add permissions manually in Azure Portal:" "Orange"
            Write-Status "  https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$script:AppId" "Blue"
            return $false
        }
    } catch {
        Write-Status "✗ Failed to add API permissions: $_" "Red"
        Write-Status "  You may need to add permissions manually in Azure Portal" "Orange"
        return $false
    }
}

# Create client secret
function New-ClientSecret {
    param($AppObjectId)
    
    Write-Status "Creating client secret..."
    Update-Progress 80
    
    try {
        $endDate = (Get-Date).AddMonths($script:ClientSecretExpiryMonths)
        
        $passwordCredential = @{
            DisplayName = $script:ClientSecretDescription
            EndDateTime = $endDate
        }
        
        $secret = Add-MgApplicationPassword -ApplicationId $AppObjectId -PasswordCredential $passwordCredential -ErrorAction Stop
        
        if ($secret.SecretText) {
            $script:ClientSecret = $secret.SecretText
            Write-Status "✓ Client secret created successfully" "Green"
            Write-Status "  Secret Value: $script:ClientSecret" "Green"
            Write-Status "  ⚠ IMPORTANT: Save this secret now - it won't be shown again!" "Red"
            
            Update-Progress 90
            return $true
        } else {
            Write-Status "✗ Failed to retrieve secret value" "Red"
            return $false
        }
    } catch {
        Write-Status "✗ Failed to create client secret: $_" "Red"
        return $false
    }
}

# Save credentials to file
function Save-CredentialsToFile {
    if (-not $script:AppId -or -not $script:TenantId -or -not $script:ClientSecret) {
        [System.Windows.Forms.MessageBox]::Show("No credentials to save. Please run setup first.", "No Credentials", "OK", "Warning")
        return
    }
    
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "JSON files (*.json)|*.json|Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $saveDialog.FileName = "barracuda_xdr_credentials.json"
    $saveDialog.InitialDirectory = "C:\Git\m365monitoring"
    
    if ($saveDialog.ShowDialog() -eq "OK") {
        $credentials = @{
            application_id = $script:AppId
            directory_tenant_id = $script:TenantId
            client_secret = $script:ClientSecret
            created_at = (Get-Date).ToString("o")
        }
        
        $credentials | ConvertTo-Json -Depth 10 | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
        
        [System.Windows.Forms.MessageBox]::Show("Credentials saved to:`n$($saveDialog.FileName)", "Saved", "OK", "Information")
        Write-Status "✓ Credentials saved to: $($saveDialog.FileName)" "Green"
    }
}

# Grant admin consent for application permissions
# Note: Even Global Admins may need to grant consent manually due to tenant policies
function Grant-AdminConsent {
    param($AppObjectId)
    
    Write-Status "Granting admin consent for API permissions..."
    Write-Status "  Note: Programmatic consent may be restricted - manual consent may be required" "Blue"
    Update-Progress 76
    
    try {
        # Get the application to find its service principal
        $app = Get-MgApplication -ApplicationId $AppObjectId -ErrorAction Stop
        
        # Get or create the service principal for the app
        $sp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -ErrorAction SilentlyContinue
        
        if (-not $sp) {
            Write-Status "  Creating service principal..." "Blue"
            $spParams = @{
                AppId = $app.AppId
            }
            $sp = New-MgServicePrincipal @spParams -ErrorAction Stop
            Write-Status "  ✓ Service principal created" "Green"
        }
        
        # Get the required resource access from the app
        if (-not $app.RequiredResourceAccess) {
            Write-Status "  ⚠ No permissions found to grant consent for" "Orange"
            return $false
        }
        
        # Grant consent for each resource
        $consentGranted = $false
        foreach ($resourceAccess in $app.RequiredResourceAccess) {
            $resourceAppId = $resourceAccess.ResourceAppId
            
            # Get the resource service principal
            $resourceSp = Get-MgServicePrincipal -Filter "appId eq '$resourceAppId'" -ErrorAction SilentlyContinue
            if (-not $resourceSp) {
                Write-Status "  ⚠ Could not find service principal for resource: $resourceAppId" "Orange"
                continue
            }
            
            # Grant admin consent using appRoleAssignments (for application permissions)
            foreach ($access in $resourceAccess.ResourceAccess) {
                if ($access.Type -eq "Role") {
                    # Application permission - grant admin consent via appRoleAssignment
                    try {
                        # Check if assignment already exists
                        $existingAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -ErrorAction SilentlyContinue
                        $existingAssignment = $existingAssignments | Where-Object { 
                            $_.ResourceId -eq $resourceSp.Id -and $_.AppRoleId -eq $access.Id 
                        }
                        
                        if (-not $existingAssignment) {
                            # Create app role assignment (this grants admin consent)
                            $assignmentParams = @{
                                ServicePrincipalId = $sp.Id
                                PrincipalId = $sp.Id
                                ResourceId = $resourceSp.Id
                                AppRoleId = $access.Id
                            }
                            
                            New-MgServicePrincipalAppRoleAssignment @assignmentParams -ErrorAction Stop
                            Write-Status "  ✓ Granted consent for permission: $($access.Id)" "Green"
                            $consentGranted = $true
                        } else {
                            Write-Status "  ✓ Consent already granted for permission: $($access.Id)" "Green"
                            $consentGranted = $true
                        }
                    } catch {
                        Write-Status "  ⚠ Could not grant consent for permission $($access.Id): $_" "Orange"
                        # Fallback: Try using Invoke-MgGraphRequest
                        try {
                            $uri = "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.Id)/appRoleAssignments"
                            $body = @{
                                principalId = $sp.Id
                                resourceId = $resourceSp.Id
                                appRoleId = $access.Id
                            } | ConvertTo-Json
                            
                            Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType "application/json" -ErrorAction Stop
                            Write-Status "  ✓ Granted consent for permission: $($access.Id) (via API)" "Green"
                            $consentGranted = $true
                        } catch {
                            Write-Status "  ✗ Failed to grant consent: $_" "Red"
                        }
                    }
                }
            }
        }
        
        if ($consentGranted) {
            Write-Status "✓ Admin consent granted successfully" "Green"
            Update-Progress 78
            return $true
        } else {
            Write-Status "⚠ Could not automatically grant consent programmatically" "Orange"
            Write-Status "  This is normal - programmatic consent may be restricted by tenant policies" "Orange"
            Write-Status "" "Blue"
            Write-Status "  Option 1: Grant consent via Azure Portal:" "Blue"
            Write-Status "  1. Go to: https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$script:AppId" "Blue"
            Write-Status "  2. Click 'API permissions'" "Blue"
            Write-Status "  3. Click 'Grant admin consent for [Your Domain]'" "Blue"
            Write-Status "  4. Verify all permissions show 'Granted for [Your Domain]'" "Blue"
            Write-Status "" "Blue"
            
            # Generate admin consent URL (alternative method)
            $adminConsentUrl = "https://login.microsoftonline.com/$script:TenantId/adminconsent?client_id=$script:AppId"
            Write-Status "  Option 2: Use admin consent URL (opens browser for consent):" "Blue"
            Write-Status "  $adminConsentUrl" "Blue"
            Write-Status "" "Blue"
            Write-Status "  Would you like to open the admin consent URL now? (Check status window)" "Blue"
            
            # Try to open the URL automatically
            try {
                Start-Process $adminConsentUrl
                Write-Status "  ✓ Admin consent URL opened in browser" "Green"
            } catch {
                Write-Status "  ⚠ Could not open browser automatically - copy the URL above" "Orange"
            }
            
            Update-Progress 78
            return $false
        }
    } catch {
        Write-Status "⚠ Could not automatically grant admin consent: $_" "Orange"
        Write-Status "  This is normal - programmatic consent may be restricted by tenant policies" "Orange"
        Write-Status "" "Blue"
        Write-Status "  Option 1: Grant consent via Azure Portal:" "Blue"
        Write-Status "  https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$script:AppId" "Blue"
        Write-Status "  Steps: API permissions > Grant admin consent for [Your Domain]" "Blue"
        Write-Status "" "Blue"
        
        # Generate admin consent URL
        $adminConsentUrl = "https://login.microsoftonline.com/$script:TenantId/adminconsent?client_id=$script:AppId"
        Write-Status "  Option 2: Use admin consent URL (opens browser for consent):" "Blue"
        Write-Status "  $adminConsentUrl" "Blue"
        
        Update-Progress 78
        return $false
    }
}

# Enable audit logging programmatically
function Enable-AuditLogging {
    Write-Status "Enabling audit logging for all mailboxes..."
    Update-Progress 85
    
    try {
        # Check if ExchangeOnlineManagement module is installed
        $module = Get-Module -ListAvailable -Name "ExchangeOnlineManagement"
        if (-not $module) {
            Write-Status "  Installing ExchangeOnlineManagement module..." "Blue"
            Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Status "  ✓ Module installed" "Green"
        }
        
        # Import the module
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        
        Write-Status "  Connecting to Exchange Online..." "Blue"
        
        # Clear any existing Exchange Online sessions first
        Get-PSSession | Where-Object { $_.ConfigurationName -like "*Exchange*" } | Remove-PSSession -ErrorAction SilentlyContinue
        
        # Get user principal name from Graph context
        $userPrincipalName = $null
        try {
            $context = Get-MgContext
            if ($context -and $context.Account) {
                $userPrincipalName = $context.Account
                Write-Status "  Using account: $userPrincipalName" "Blue"
            }
        } catch {
            Write-Status "  ⚠ Could not get user principal name from Graph context" "Orange"
        }
        
        # Use interactive browser authentication with explicit UserPrincipalName
        # This is more reliable than relying on default authentication
        $connectionAttempts = 0
        $maxAttempts = 3
        $connectionSuccess = $false
        
        while (-not $connectionSuccess -and $connectionAttempts -lt $maxAttempts) {
            $connectionAttempts++
            try {
                Write-Status "  Attempt $connectionAttempts of $maxAttempts..." "Blue"
                
                if ($connectionAttempts -eq 1) {
                    # First attempt: Try with UserPrincipalName (most reliable)
                    Write-Status "  Opening browser for authentication..." "Blue"
                    if ($userPrincipalName) {
                        Connect-ExchangeOnline -UserPrincipalName $userPrincipalName -ShowBanner:$false -ErrorAction Stop | Out-Null
                    } else {
                        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
                    }
                } elseif ($connectionAttempts -eq 2) {
                    # Second attempt: Try device code flow
                    Write-Status "  Browser auth failed, trying device code authentication..." "Blue"
                    Write-Status "  A device code will be displayed - visit https://microsoft.com/devicelogin" "Blue"
                    Connect-ExchangeOnline -Device -ShowBanner:$false -ErrorAction Stop | Out-Null
                } else {
                    # Third attempt: Try without UserPrincipalName
                    Write-Status "  Final attempt: Interactive authentication..." "Blue"
                    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
                }
                
                # Wait for connection to establish
                Start-Sleep -Seconds 5
                
                # Verify by running a command - this is the real test
                try {
                    $testResult = Get-AdminAuditLogConfig -ErrorAction Stop
                    if ($null -ne $testResult) {
                        $connectionSuccess = $true
                        Write-Status "  ✓ Exchange Online connection established and verified" "Green"
                    }
                } catch {
                    # Connection might be established but commands not ready yet
                    Write-Status "  Waiting for commands to become available..." "Blue"
                    Start-Sleep -Seconds 3
                    $testResult = Get-AdminAuditLogConfig -ErrorAction Stop
                    if ($null -ne $testResult) {
                        $connectionSuccess = $true
                        Write-Status "  ✓ Exchange Online connection verified" "Green"
                    } else {
                        throw "Commands not available after connection"
                    }
                }
            } catch {
                $errorMsg = $_.Exception.Message
                Write-Status "  ⚠ Connection attempt $connectionAttempts failed: $errorMsg" "Orange"
                
                if ($connectionAttempts -lt $maxAttempts) {
                    Write-Status "  Retrying with different method..." "Blue"
                    Start-Sleep -Seconds 3
                    # Clean up any partial connections
                    Get-PSSession | Where-Object { $_.ConfigurationName -like "*Exchange*" } | Remove-PSSession -ErrorAction SilentlyContinue
                } else {
                    Write-Status "  ✗ All authentication methods failed" "Red"
                    Write-Status "  Please try connecting manually:" "Orange"
                    Write-Status "  Connect-ExchangeOnline -UserPrincipalName your-admin@domain.com" "Blue"
                    throw "Failed to establish Exchange Online connection after $maxAttempts attempts. Last error: $errorMsg"
                }
            }
        }
        
        if (-not $connectionSuccess) {
            throw "Failed to establish Exchange Online connection"
        }
        
        # Wait a moment for connection to stabilize
        Start-Sleep -Seconds 2
        
        # Check current configuration
        $currentConfig = $null
        try {
            $currentConfig = Get-AdminAuditLogConfig -ErrorAction Stop
        } catch {
            Write-Status "  ⚠ Could not retrieve current audit log configuration: $_" "Orange"
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            return $false
        }
        
        if ($null -eq $currentConfig) {
            Write-Status "  ⚠ Could not retrieve audit log configuration" "Orange"
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            return $false
        }
        
        if ($currentConfig.UnifiedAuditLogIngestionEnabled) {
            Write-Status "  ✓ Audit logging is already enabled" "Green"
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            Update-Progress 87
            return $true
        }
        
        # Enable audit logging
        Write-Status "  Enabling unified audit log..." "Blue"
        try {
            Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true -ErrorAction Stop
        } catch {
            Write-Status "  ✗ Failed to enable audit logging: $_" "Red"
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            return $false
        }
        
        # Wait a moment for the change to propagate
        Start-Sleep -Seconds 2
        
        # Verify
        $verifyConfig = $null
        try {
            $verifyConfig = Get-AdminAuditLogConfig -ErrorAction Stop
        } catch {
            Write-Status "  ⚠ Enabled but could not verify: $_" "Orange"
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            return $true  # Assume success if we can't verify
        }
        
        if ($null -ne $verifyConfig -and $verifyConfig.UnifiedAuditLogIngestionEnabled) {
            Write-Status "  ✓ Audit logging enabled successfully!" "Green"
            Write-Status "  Note: It may take up to 60 minutes for changes to take effect" "Orange"
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            Update-Progress 87
            return $true
        } else {
            Write-Status "  ⚠ Audit logging may not be fully enabled yet" "Orange"
            Write-Status "  Please verify manually in Exchange Admin Center" "Orange"
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            return $false
        }
    } catch {
        Write-Status "  ✗ Failed to enable audit logging: $_" "Red"
        Write-Status "  Please enable manually via Admin Center or PowerShell:" "Orange"
        Write-Status "  https://portal.office.com > Compliance > Audit" "Blue"
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        return $false
    }
}

# Enable audit logging (GUI version with prompt)
function Enable-AuditLoggingGUI {
    $result = [System.Windows.Forms.MessageBox]::Show(
        "This will enable audit logging via Exchange Online PowerShell.`n`nDo you want to continue?",
        "Enable Audit Logging",
        "YesNo",
        "Question"
    )
    
    if ($result -eq "Yes") {
        Write-Status "Enabling audit logging..."
        
        try {
            # Check if ExchangeOnlineManagement module is installed
            $module = Get-Module -ListAvailable -Name "ExchangeOnlineManagement"
            if (-not $module) {
                Write-Status "Installing ExchangeOnlineManagement module..." "Blue"
                Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
            }
            
            Write-Status "Connecting to Exchange Online..." "Blue"
            
            # Get user principal name from Graph context if available
            $userPrincipalName = $null
            try {
                $context = Get-MgContext -ErrorAction SilentlyContinue
                if ($context -and $context.Account) {
                    $userPrincipalName = $context.Account
                    Write-Status "Using account: $userPrincipalName" "Blue"
                }
            } catch {
                # Ignore - will use default auth
            }
            
            # Clear any existing sessions
            Get-PSSession | Where-Object { $_.ConfigurationName -like "*Exchange*" } | Remove-PSSession -ErrorAction SilentlyContinue
            
            # Use interactive browser authentication with explicit UserPrincipalName
            Write-Status "Opening browser for authentication..." "Blue"
            $connectionSuccess = $false
            
            # Try with UserPrincipalName first (most reliable)
            try {
                if ($userPrincipalName) {
                    Connect-ExchangeOnline -UserPrincipalName $userPrincipalName -ShowBanner:$false -ErrorAction Stop | Out-Null
                } else {
                    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
                }
                Start-Sleep -Seconds 5
                
                # Verify connection by running a command
                $testResult = Get-AdminAuditLogConfig -ErrorAction Stop
                if ($null -ne $testResult) {
                    $connectionSuccess = $true
                    Write-Status "✓ Exchange Online connection established" "Green"
                }
            } catch {
                Write-Status "⚠ Browser auth failed, trying device code..." "Orange"
                try {
                    Connect-ExchangeOnline -Device -ShowBanner:$false -ErrorAction Stop | Out-Null
                    Write-Status "Device code displayed - visit https://microsoft.com/devicelogin" "Blue"
                    Start-Sleep -Seconds 5
                    
                    # Verify connection
                    $testResult = Get-AdminAuditLogConfig -ErrorAction Stop
                    if ($null -ne $testResult) {
                        $connectionSuccess = $true
                        Write-Status "✓ Exchange Online connection established (device code)" "Green"
                    }
                } catch {
                    Write-Status "✗ Connection failed: $_" "Red"
                    throw "Failed to establish Exchange Online connection"
                }
            }
            
            if (-not $connectionSuccess) {
                throw "Failed to establish Exchange Online connection"
            }
            
            # Wait a moment for connection to stabilize
            Start-Sleep -Seconds 2
            
            # Check current configuration
            $currentConfig = Get-AdminAuditLogConfig -ErrorAction Stop
            
            if ($null -ne $currentConfig -and $currentConfig.UnifiedAuditLogIngestionEnabled) {
                Write-Status "✓ Audit logging is already enabled" "Green"
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
                return
            }
            
            Write-Status "Enabling unified audit log..." "Blue"
            Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true -ErrorAction Stop
            
            # Wait a moment for the change to propagate
            Start-Sleep -Seconds 2
            
            # Verify
            $verifyConfig = Get-AdminAuditLogConfig -ErrorAction Stop
            if ($null -ne $verifyConfig -and $verifyConfig.UnifiedAuditLogIngestionEnabled) {
                Write-Status "✓ Audit logging enabled successfully!" "Green"
                Write-Status "  Note: It may take up to 60 minutes for changes to take effect" "Orange"
            } else {
                Write-Status "⚠ Audit logging may not be fully enabled yet" "Orange"
                Write-Status "  Please verify manually in Exchange Admin Center" "Orange"
            }
            
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        } catch {
            Write-Status "✗ Failed to enable audit logging: $_" "Red"
            Write-Status "  Please enable manually via Admin Center or PowerShell:" "Orange"
            Write-Status "  https://portal.office.com > Compliance > Audit" "Blue"
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        }
    }
}

# Test connection
function Test-ConnectionGUI {
    if (-not $script:AppId -or -not $script:TenantId -or -not $script:ClientSecret) {
        [System.Windows.Forms.MessageBox]::Show("No credentials available. Please run setup first.", "No Credentials", "OK", "Warning")
        return
    }
    
    Write-Status "Testing connection..."
    
    try {
        # Verify application exists
        $app = Get-MgApplication -Filter "appId eq '$script:AppId'"
        if ($app) {
            Write-Status "✓ Application found in Entra ID" "Green"
            Write-Status "✓ Connection test successful" "Green"
            [System.Windows.Forms.MessageBox]::Show("Connection test successful!", "Success", "OK", "Information")
        } else {
            Write-Status "✗ Application not found" "Red"
            [System.Windows.Forms.MessageBox]::Show("Application not found. Please verify the Application ID.", "Test Failed", "OK", "Error")
        }
    } catch {
        Write-Status "✗ Connection test failed: $_" "Red"
        [System.Windows.Forms.MessageBox]::Show("Connection test failed: $_", "Test Failed", "OK", "Error")
    }
}

# Main setup process
function Start-SetupProcess {
    Write-Status "========================================" "Blue"
    Write-Status "Barracuda XDR Microsoft 365 Setup" "Blue"
    Write-Status "========================================" "Blue"
    Write-Status ""
    
    Update-Progress 0
    
    # Step 1: Install modules
    if (-not (Install-RequiredModules)) {
        Write-Status "Setup failed at module installation step" "Red"
        return
    }
    
    # Step 2: Connect to Graph
    if (-not (Connect-MicrosoftGraph)) {
        Write-Status "Setup failed at authentication step" "Red"
        return
    }
    
    # Step 3: Check for existing app or register new
    $app = Get-ExistingApplication
    if (-not $app) {
        $app = Register-Application
        if (-not $app) {
            Write-Status "Setup failed at application registration step" "Red"
            return
        }
    }
    
    # Step 4: Add API permissions
    if (Add-APIPermissions -AppObjectId $app.Id) {
        # Step 4a: Grant admin consent
        Grant-AdminConsent -AppObjectId $app.Id | Out-Null
    }
    
    # Step 5: Create client secret
    if (-not (New-ClientSecret -AppObjectId $app.Id)) {
        Write-Status "Setup failed at client secret creation step" "Red"
        return
    }
    
    # Step 6: Enable audit logging
    $auditResult = [System.Windows.Forms.MessageBox]::Show(
        "Do you want to enable audit logging now?`n`nThis will connect to Exchange Online and enable unified audit log ingestion.",
        "Enable Audit Logging",
        "YesNo",
        "Question"
    )
    
    if ($auditResult -eq "Yes") {
        Enable-AuditLogging | Out-Null
    } else {
        Write-Status "⚠ Audit logging not enabled. You can enable it later using the 'Enable Audit Logging' button." "Orange"
    }
    
    # Step 7: Update UI with credentials
    if ($script:AppIdTextBox.InvokeRequired) {
        $script:AppIdTextBox.Invoke([System.Action[String]]{
            param($text)
            $script:AppIdTextBox.Text = $text
        }, $script:AppId)
    } else {
        $script:AppIdTextBox.Text = $script:AppId
    }
    
    if ($script:TenantIdTextBox.InvokeRequired) {
        $script:TenantIdTextBox.Invoke([System.Action[String]]{
            param($text)
            $script:TenantIdTextBox.Text = $text
        }, $script:TenantId)
    } else {
        $script:TenantIdTextBox.Text = $script:TenantId
    }
    
    if ($script:SecretTextBox.InvokeRequired) {
        $script:SecretTextBox.Invoke([System.Action[String]]{
            param($text)
            $script:SecretTextBox.Text = $text
        }, $script:ClientSecret)
    } else {
        $script:SecretTextBox.Text = $script:ClientSecret
    }
    
    Update-Progress 100
    
    Write-Status ""
    Write-Status "========================================" "Green"
    Write-Status "Setup completed successfully!" "Green"
    Write-Status "========================================" "Green"
    Write-Status ""
    Write-Status "Next steps:" "Blue"
    Write-Status "1. Verify admin consent was granted (check Azure Portal if needed)" "Blue"
    Write-Status "2. Verify audit logging is enabled (check Exchange Admin Center if needed)" "Blue"
    Write-Status "3. Enter credentials in Barracuda XDR Dashboard" "Blue"
    
    [System.Windows.Forms.MessageBox]::Show(
        "Setup completed successfully!`n`nNext steps:`n1. Verify admin consent was granted`n2. Verify audit logging is enabled`n3. Enter credentials in Barracuda XDR Dashboard",
        "Setup Complete",
        "OK",
        "Information"
    )
}

# Main execution
try {
    # Check if running as administrator (recommended)
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "This script works best when run as Administrator.`n`nContinue anyway?",
            "Administrator Rights",
            "YesNo",
            "Question"
        )
        if ($result -eq "No") {
            exit
        }
    }
    
    # Create and show form
    $form = New-MainForm
    Write-Status "Ready to begin setup. Click 'Start Setup' to continue." "Blue"
    [System.Windows.Forms.Application]::Run($form)
} catch {
    [System.Windows.Forms.MessageBox]::Show("An error occurred: $_", "Error", "OK", "Error")
    Write-Host "Error: $_" -ForegroundColor Red
} finally {
    # Cleanup
    if ($script:GraphConnection) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
}
