# Test script for Office 365 Management API permissions
# This script tests various methods to find and add the required permissions

param(
    [string]$AppId = "",  # App (Client) ID - will search by name if not provided
    [string]$AppObjectId = ""  # Optional: Object ID if known
)

$ErrorActionPreference = "Continue"

# Required permissions
$requiredPermissions = @(
    "ActivityFeed.Read",
    "ActivityFeed.ReadDlp",
    "ServiceHealth.Read"
)

# Office 365 Management API App IDs to try
$o365AppIds = @(
    'c5393580-f805-4401-95e8-94b7a6ef2fc2',  # Office 365 Management API (from Claude)
    '00000003-0000-0ff1-ce00-000000000000',  # Standard (but might be SharePoint)
    '797f4846-ba00-4fd7-ba43-dac1f8f63013'   # Another possible
)

# Known permission IDs (from Claude's approach)
$knownPermissionIds = @{
    "ActivityFeed.Read" = "594c1fb6-4f81-4f1d-ab49-e4b9e6e7e7c0"
    "ActivityFeed.ReadDlp" = "4807a72c-ad38-4250-94c9-4eabfe26cd55"
    "ServiceHealth.Read" = "f2e896c5-7f55-4e12-9264-b8f5eff67b28"
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Office 365 Management API Permission Test" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Connect to Microsoft Graph
Write-Host "[1/6] Connecting to Microsoft Graph..." -ForegroundColor Yellow
try {
    $scopes = @(
        "Application.ReadWrite.All",
        "Directory.ReadWrite.All",
        "User.Read"
    )
    Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
    $context = Get-MgContext
    Write-Host "✓ Connected as: $($context.Account)" -ForegroundColor Green
    Write-Host "  Tenant: $($context.TenantId)" -ForegroundColor Gray
} catch {
    Write-Host "✗ Failed to connect: $_" -ForegroundColor Red
    exit 1
}

Write-Host ""

# Get the application
Write-Host "[2/6] Getting application details..." -ForegroundColor Yellow
try {
    # First, find the application by AppId (Client ID)
    Write-Host "  Searching for application with AppId: $AppId" -ForegroundColor Gray
    
    if ($AppObjectId) {
        # Use provided Object ID
        $app = Get-MgApplication -ApplicationId $AppObjectId -ErrorAction Stop
        Write-Host "  Using provided Object ID: $AppObjectId" -ForegroundColor Gray
    } else {
        # Search by AppId
        $allApps = Get-MgApplication -All -ErrorAction Stop | Where-Object { $_.AppId -eq $AppId }
        if ($allApps) {
            $app = $allApps | Select-Object -First 1
            Write-Host "  Found application by AppId" -ForegroundColor Gray
        } else {
            # Try searching by display name
            Write-Host "  AppId not found, searching by display name 'SKOUTCYBERSECURITY'..." -ForegroundColor Gray
            $allApps = Get-MgApplication -All -ErrorAction Stop | Where-Object { 
                $_.DisplayName -like "*SKOUT*" -or $_.DisplayName -like "*Barracuda*"
            }
            if ($allApps) {
                $app = $allApps | Select-Object -First 1
                Write-Host "  Found application by name: $($app.DisplayName)" -ForegroundColor Gray
            }
        }
    }
    
    if (-not $app) {
        throw "Application not found with AppId: $AppId"
    }
    
    Write-Host "✓ Found application: $($app.DisplayName)" -ForegroundColor Green
    Write-Host "  App (Client) ID: $($app.AppId)" -ForegroundColor Gray
    Write-Host "  Object ID: $($app.Id)" -ForegroundColor Gray
    
    # Store Object ID for later use
    $script:AppObjectId = $app.Id
    
    # Show current RequiredResourceAccess
    if ($app.RequiredResourceAccess) {
        Write-Host "  Current RequiredResourceAccess:" -ForegroundColor Gray
        foreach ($resource in $app.RequiredResourceAccess) {
            Write-Host "    - ResourceAppId: $($resource.ResourceAppId)" -ForegroundColor Gray
            Write-Host "      Permissions: $($resource.ResourceAccess.Count)" -ForegroundColor Gray
        }
    } else {
        Write-Host "  No RequiredResourceAccess configured" -ForegroundColor Gray
    }
} catch {
    Write-Host "✗ Failed to get application: $_" -ForegroundColor Red
    Write-Host "  Make sure the AppId is correct, or provide the Object ID with -AppObjectId parameter" -ForegroundColor Yellow
    exit 1
}

Write-Host ""

# Search for Office 365 Management API service principal
Write-Host "[3/6] Searching for Office 365 Management API service principal..." -ForegroundColor Yellow
$foundSp = $null

foreach ($appId in $o365AppIds) {
    Write-Host "  Trying AppId: $appId" -ForegroundColor Gray
    try {
        $sp = Get-MgServicePrincipal -Filter "appId eq '$appId'" -ErrorAction SilentlyContinue
        if (-not $sp) {
            $allSp = Get-MgServicePrincipal -All -ErrorAction SilentlyContinue | Where-Object { $_.AppId -eq $appId }
            if ($allSp) {
                $sp = $allSp | Select-Object -First 1
            }
        }
        
        if ($sp) {
            Write-Host "  ✓ Found service principal: $($sp.DisplayName)" -ForegroundColor Green
            Write-Host "    AppId: $($sp.AppId)" -ForegroundColor Gray
            Write-Host "    ID: $($sp.Id)" -ForegroundColor Gray
            
            # Get full details with AppRoles
            try {
                $spUri = "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.Id)?`$select=id,appId,displayName,appRoles"
                $spData = Invoke-MgGraphRequest -Method GET -Uri $spUri -ErrorAction Stop
                
                if ($spData.appRoles) {
                    Write-Host "    Found $($spData.appRoles.Count) AppRoles" -ForegroundColor Green
                    $foundSp = $spData
                    $foundSp.originalSp = $sp
                    break
                } else {
                    Write-Host "    ⚠ No AppRoles found" -ForegroundColor Yellow
                }
            } catch {
                Write-Host "    ⚠ Error getting AppRoles: $_" -ForegroundColor Yellow
            }
        }
    } catch {
        Write-Host "  ✗ Error: $_" -ForegroundColor Red
    }
}

if (-not $foundSp) {
    Write-Host "  Searching all service principals for Office 365 Management API..." -ForegroundColor Yellow
    Write-Host "  Note: AppId 00000003-0000-0ff1-ce00-000000000000 is SharePoint, not Management API" -ForegroundColor Yellow
    Write-Host ""
    
    try {
        # Search by various patterns
        $searchPatterns = @(
            "*Management*API*",
            "*Office 365*Management*",
            "*O365*Management*",
            "*Management Activity*",
            "*Activity API*"
        )
        
        $allSp = Get-MgServicePrincipal -All -ErrorAction Stop
        
        Write-Host "  Searching through $($allSp.Count) total service principals..." -ForegroundColor Gray
        
        $matchingSp = @()
        foreach ($sp in $allSp) {
            $match = $false
            foreach ($pattern in $searchPatterns) {
                if ($sp.DisplayName -like $pattern) {
                    $match = $true
                    break
                }
            }
            if ($match) {
                $matchingSp += $sp
            }
        }
        
        Write-Host "  Found $($matchingSp.Count) matching service principal(s)" -ForegroundColor Gray
        
        foreach ($sp in $matchingSp) {
            Write-Host "    Checking: $($sp.DisplayName) (AppId: $($sp.AppId))" -ForegroundColor Gray
            try {
                $spUri = "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.Id)?`$select=id,appId,displayName,appRoles"
                $spData = Invoke-MgGraphRequest -Method GET -Uri $spUri -ErrorAction Stop
                
                if ($spData.appRoles -and $spData.appRoles.Count -gt 0) {
                    Write-Host "      ✓ Found $($spData.appRoles.Count) AppRoles" -ForegroundColor Green
                    
                    # Check if this has the permissions we need
                    $hasActivityFeed = $spData.appRoles | Where-Object { $_.value -like "*ActivityFeed*" }
                    $hasServiceHealth = $spData.appRoles | Where-Object { $_.value -like "*ServiceHealth*" -or $_.value -like "*Health*" }
                    
                    if ($hasActivityFeed -or $hasServiceHealth) {
                        Write-Host "      ⚠ This might be the Management API!" -ForegroundColor Yellow
                        $foundSp = $spData
                        $foundSp.originalSp = $sp
                        break
                    }
                }
            } catch {
                Write-Host "      ⚠ Error: $_" -ForegroundColor Yellow
            }
        }
        
        # If still not found, list all service principals with "Management" in the name
        if (-not $foundSp) {
            Write-Host ""
            Write-Host "  Listing all service principals with 'Management' in name:" -ForegroundColor Yellow
            $managementSp = $allSp | Where-Object { $_.DisplayName -like "*Management*" } | Select-Object DisplayName, AppId -First 20
            foreach ($sp in $managementSp) {
                Write-Host "    - $($sp.DisplayName) (AppId: $($sp.AppId))" -ForegroundColor Gray
            }
        }
    } catch {
        Write-Host "  ✗ Error searching: $_" -ForegroundColor Red
    }
}

if (-not $foundSp) {
    Write-Host "✗ Could not find Office 365 Management API service principal" -ForegroundColor Red
    Write-Host ""
    Write-Host "This means the Office 365 Management API is not available in your tenant," -ForegroundColor Yellow
    Write-Host "or the permissions are not exposed as AppRoles." -ForegroundColor Yellow
    exit 1
}

Write-Host ""

# Search for required permissions
Write-Host "[4/6] Searching for required permissions..." -ForegroundColor Yellow
$foundPermissions = @()

if ($foundSp.appRoles) {
    Write-Host "  Available AppRoles ($($foundSp.appRoles.Count)):" -ForegroundColor Gray
    
    # Show first 10 AppRoles
    $shown = 0
    foreach ($role in $foundSp.appRoles) {
        if ($shown -lt 10) {
            Write-Host "    - $($role.value) (ID: $($role.id))" -ForegroundColor Gray
            $shown++
        }
    }
    if ($foundSp.appRoles.Count -gt 10) {
        Write-Host "    ... and $($foundSp.appRoles.Count - 10) more" -ForegroundColor Gray
    }
    
    Write-Host ""
    
    foreach ($permName in $requiredPermissions) {
        Write-Host "  Searching for: $permName" -ForegroundColor Yellow
        
        # Try exact match
        $found = $foundSp.appRoles | Where-Object { $_.value -eq $permName }
        
        # Try case-insensitive
        if (-not $found) {
            $found = $foundSp.appRoles | Where-Object { $_.value -ieq $permName }
        }
        
        # Try pattern matching
        if (-not $found) {
            if ($permName -eq "ActivityFeed.Read") {
                $found = $foundSp.appRoles | Where-Object {
                    ($_.value -like "*ActivityFeed*" -or $_.value -like "*Activity*Feed*") -and
                    $_.value -like "*Read*" -and
                    $_.value -notlike "*Dlp*" -and
                    $_.value -notlike "*DLP*"
                } | Select-Object -First 1
            } elseif ($permName -eq "ActivityFeed.ReadDlp") {
                $found = $foundSp.appRoles | Where-Object {
                    ($_.value -like "*ActivityFeed*Dlp*" -or
                     $_.value -like "*ActivityFeed*DLP*" -or
                     ($_.value -like "*Activity*" -and ($_.value -like "*Dlp*" -or $_.value -like "*DLP*")))
                } | Select-Object -First 1
            } elseif ($permName -eq "ServiceHealth.Read") {
                $found = $foundSp.appRoles | Where-Object {
                    $_.value -like "*ServiceHealth*" -or
                    ($_.value -like "*Service*" -and $_.value -like "*Health*")
                } | Select-Object -First 1
            }
        }
        
        if ($found) {
            Write-Host "    ✓ Found: $($found.value) (ID: $($found.id))" -ForegroundColor Green
            $foundPermissions += @{
                Name = $permName
                FoundAs = $found.value
                Id = $found.id
                Type = "Role"
            }
        } else {
            Write-Host "    ✗ Not found" -ForegroundColor Red
        }
    }
} else {
    Write-Host "✗ No AppRoles available in service principal" -ForegroundColor Red
}

Write-Host ""

# Try to add permissions
Write-Host "[5/6] Attempting to add permissions..." -ForegroundColor Yellow

if ($foundPermissions.Count -eq 0) {
    Write-Host "⚠ No permissions found via AppRoles search" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Attempting Claude's approach: Adding permissions using known IDs..." -ForegroundColor Yellow
    
    # Try Claude's approach - add permissions directly using known IDs
    $claudeAppId = 'c5393580-f805-4401-95e8-94b7a6ef2fc2'  # Office 365 Management API AppId
    
    Write-Host "  Using AppId: $claudeAppId" -ForegroundColor Gray
    Write-Host "  Adding permissions by known IDs:" -ForegroundColor Gray
    
    $resourceAccess = @{
        ResourceAppId = $claudeAppId
        ResourceAccess = @()
    }
    
    foreach ($permName in $requiredPermissions) {
        if ($knownPermissionIds.ContainsKey($permName)) {
            $permId = $knownPermissionIds[$permName]
            $resourceAccess.ResourceAccess += @{
                Id = $permId
                Type = "Role"
            }
            Write-Host "    - $permName (ID: $permId)" -ForegroundColor Gray
            $foundPermissions += @{
                Name = $permName
                FoundAs = "Known ID"
                Id = $permId
                Type = "Role"
            }
        }
    }
    
    if ($foundPermissions.Count -gt 0) {
        Write-Host "  ✓ Found $($foundPermissions.Count) permissions using known IDs" -ForegroundColor Green
        $o365Access = $resourceAccess
    } else {
        Write-Host "⚠ Could not add permissions using known IDs" -ForegroundColor Yellow
        Write-Host "The permissions are not exposed as AppRoles in your tenant." -ForegroundColor Yellow
        Write-Host "They must be added manually in Azure Portal." -ForegroundColor Yellow
    }
}

if ($foundPermissions.Count -gt 0) {
    Write-Host "  Found $($foundPermissions.Count) of $($requiredPermissions.Count) permissions" -ForegroundColor Green
    
    # Build RequiredResourceAccess
    $resourceAccess = @{
        ResourceAppId = $foundSp.appId
        ResourceAccess = @()
    }
    
    foreach ($perm in $foundPermissions) {
        $resourceAccess.ResourceAccess += @{
            Id = $perm.Id
            Type = $perm.Type
        }
        Write-Host "    ✓ Will add: $($perm.Name) -> $($perm.FoundAs)" -ForegroundColor Green
    }
    
    # Get current RequiredResourceAccess
    $currentRequiredAccess = $app.RequiredResourceAccess | Where-Object {
        $_.ResourceAppId -eq $foundSp.appId
    }
    
    if ($currentRequiredAccess) {
        Write-Host "  Found existing Office 365 Management API resource" -ForegroundColor Gray
        Write-Host "  Current permissions: $($currentRequiredAccess.ResourceAccess.Count)" -ForegroundColor Gray
        
        # Merge with existing
        $existingIds = $currentRequiredAccess.ResourceAccess | ForEach-Object { $_.Id }
        foreach ($perm in $foundPermissions) {
            if ($perm.Id -notin $existingIds) {
                $currentRequiredAccess.ResourceAccess += @{
                    Id = $perm.Id
                    Type = $perm.Type
                }
            }
        }
        $resourceAccess = $currentRequiredAccess
    }
    
    # Update application
    Write-Host ""
    Write-Host "  Updating application RequiredResourceAccess..." -ForegroundColor Yellow
    
    try {
        # Get all current RequiredResourceAccess
        $allRequiredAccess = @()
        foreach ($resource in $app.RequiredResourceAccess) {
            if ($resource.ResourceAppId -eq $foundSp.appId) {
                $allRequiredAccess += $resourceAccess
            } else {
                $allRequiredAccess += $resource
            }
        }
        
        # Add if not already present
        if (-not ($app.RequiredResourceAccess | Where-Object { $_.ResourceAppId -eq $foundSp.appId })) {
            $allRequiredAccess += $resourceAccess
        }
        
        # Update application (use Object ID, not App ID)
        $updateParams = @{
            ApplicationId = $script:AppObjectId
            RequiredResourceAccess = $allRequiredAccess
        }
        
        Update-MgApplication @updateParams -ErrorAction Stop
        Write-Host "  ✓ Application updated successfully" -ForegroundColor Green
    } catch {
        Write-Host "  ✗ Failed to update application: $_" -ForegroundColor Red
        Write-Host "    Error details: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""

# Verify
Write-Host "[6/6] Verifying permissions..." -ForegroundColor Yellow
try {
    $updatedApp = Get-MgApplication -ApplicationId $script:AppObjectId -ErrorAction Stop
    $o365Resource = $updatedApp.RequiredResourceAccess | Where-Object {
        $_.ResourceAppId -eq $foundSp.appId
    }
    
    if ($o365Resource) {
        Write-Host "✓ Office 365 Management API resource found in RequiredResourceAccess" -ForegroundColor Green
        Write-Host "  Permissions configured: $($o365Resource.ResourceAccess.Count)" -ForegroundColor Gray
        
        foreach ($access in $o365Resource.ResourceAccess) {
            $permInfo = $foundSp.appRoles | Where-Object { $_.id -eq $access.Id }
            if ($permInfo) {
                Write-Host "    - $($permInfo.value) (ID: $($access.Id))" -ForegroundColor Green
            } else {
                Write-Host "    - Unknown permission (ID: $($access.Id))" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host "⚠ Office 365 Management API resource not found in RequiredResourceAccess" -ForegroundColor Yellow
    }
} catch {
    Write-Host "✗ Failed to verify: $_" -ForegroundColor Red
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Test Complete" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "1. Check Azure Portal to see if permissions were added" -ForegroundColor Gray
Write-Host "2. Grant admin consent if permissions were added" -ForegroundColor Gray
Write-Host "3. If permissions weren't found, they must be added manually" -ForegroundColor Gray
Write-Host ""

# Disconnect
Disconnect-MgGraph -ErrorAction SilentlyContinue
