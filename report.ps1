# --- 1. Connect to Microsoft 365 Services ---
# Install modules if not already installed:
# Install-Module Microsoft.Graph -Scope CurrentUser
# Install-Module ExchangeOnlineManagement -Scope CurrentUser

# Connect to Microsoft Graph with required scopes for sign-in data
Connect-MgGraph -Scopes "AuditLog.Read.All", "User.Read.All", "Directory.Read.All"
# Connect to Exchange Online (for archive mailbox info)
Connect-ExchangeOnline

# --- 2. Define Inactivity Threshold (90 days) ---
$inactiveDays    = 90
$inactiveDate    = (Get-Date).AddDays(-$inactiveDays)

# --- 3. Retrieve all users with required properties from Graph ---
# Prepare an output list
$InactiveUsersReport = @()

# Get all user objects with relevant properties in one go (paged to avoid throttling)
$allUsers = Get-MgUser -All -PageSize 100 -Property `
    "Id", "DisplayName", "UserPrincipalName", "Mail", "Department", `
    "AccountEnabled", "SignInActivity", "AssignedLicenses", "CreatedDateTime"

# Filter to only enabled users (exclude already disabled accounts) for analysis
$activeUsers = $allUsers | Where-Object { $_.AccountEnabled -eq $true }

# --- 4. Identify users with last login older than threshold (or never logged in) ---
# Note: SignInActivity.LastSignInDateTime is null if the user never logged in 
# or last login was before Azure AD started tracking (Apr 2020)&#8203;:contentReference[oaicite:11]{index=11}.
$inactiveCandidates = foreach ($user in $activeUsers) {
    $lastSignIn = $user.SignInActivity.LastSignInDateTime
    # Determine if user is inactive:
    if (($lastSignIn -eq $null -and $user.CreatedDateTime -lt $inactiveDate) -or `
        ($lastSignIn -ne $null -and $lastSignIn -lt $inactiveDate)) {
        # Calculate inactive days (if never logged in, use days since creation)
        $daysInactive = if ($lastSignIn) { 
                           (New-TimeSpan -Start $lastSignIn -End (Get-Date)).Days 
                        } else {
                           (New-TimeSpan -Start $user.CreatedDateTime -End (Get-Date)).Days
                        }
        # Prepare a placeholder for license names (filled next)
        [PSCustomObject]@{
            UserId          = $user.Id
            UPN             = $user.UserPrincipalName
            DisplayName     = $user.DisplayName
            Email           = $user.Mail
            LastLoginDate   = if ($lastSignIn) { [DateTime]$lastSignIn } else { $null }
            InactiveDays    = $daysInactive
            LicenseSkuIds   = $user.AssignedLicenses.SkuId   # collection of GUIDs
            Department      = $user.Department
            ArchiveEnabled  = $false  # default, will update based on mailbox info
        }
    }
}

# --- 5. Map License GUIDs to License Type Names ---
# Retrieve all license SKUs for the tenant (to translate GUIDs to product names)
$skuMap = @{}
$skus = Get-MgSubscribedSku -All
foreach ($sku in $skus) {
    $skuMap[$sku.SkuId] = $sku.SkuPartNumber   # e.g., SkuPartNumber "ENTERPRISEPACK" for E3
}

# Add a friendly license name (or multiple names) to each inactive user entry
foreach ($entry in $inactiveCandidates) {
    if ($entry.LicenseSkuIds -and $entry.LicenseSkuIds.Count -gt 0) {
        $licenseNames = @()
        foreach ($skuId in $entry.LicenseSkuIds) {
            if ($skuMap.ContainsKey($skuId)) {
                $licenseNames += $skuMap[$skuId]
            } else {
                $licenseNames += $skuId  # fallback to GUID if not found in map
            }
        }
        $entry | Add-Member -NotePropertyName "LicenseType" `
                          -NotePropertyValue ($licenseNames -join "; ")
    } else {
        $entry | Add-Member -NotePropertyName "LicenseType" -NotePropertyValue "Unlicensed"
    }
}

# --- 6. Get archive mailbox status for each inactive user from Exchange Online ---
# Fetch all user mailboxes with archive info in one call
# (Using Get-ExoMailbox for efficiency in large tenants; falls back to Get-Mailbox if needed)
Try {
    $allMailboxes = Get-ExoMailbox -ResultSize Unlimited -Properties ArchiveStatus,ExternalDirectoryObjectId
} Catch {
    # If Get-ExoMailbox fails or not available, use classic Get-Mailbox
    $allMailboxes = Get-Mailbox -ResultSize Unlimited -Properties ArchiveStatus,ExternalDirectoryObjectId
}
# Create a lookup for archive status by user object ID (ExternalDirectoryObjectId matches Azure AD user Id)
$archiveStatusMap = @{}
foreach ($mbx in $allMailboxes) {
    $isArchiveEnabled = $false
    if ($mbx.ArchiveStatus -eq "Active") { $isArchiveEnabled = $true }
    # (Alternatively: if ($mbx.ArchiveGuid -and $mbx.ArchiveGuid -ne [Guid]::Empty) { $true })
    $archiveStatusMap[$mbx.ExternalDirectoryObjectId] = $isArchiveEnabled
}

# Update each inactive user entry with archive mailbox status
foreach ($entry in $inactiveCandidates) {
    if ($archiveStatusMap.ContainsKey($entry.UserId)) {
        $entry.ArchiveEnabled = $archiveStatusMap[$entry.UserId]
    } else {
        # No mailbox found (could be unlicensed for Exchange) -> archive not applicable
        $entry.ArchiveEnabled = $false
    }
}

# --- 7. Build final output objects (select and rename as needed) and export to CSV ---
$InactiveUsersReport = $inactiveCandidates | Select-Object `
    UPN, DisplayName, Email, 
    @{Name="LastLoginDate"; Expression={ if ($_.LastLoginDate) { [DateTime]$_.LastLoginDate } else {"Never"} } }, 
    @{Name="InactiveDays";  Expression={ $_.InactiveDays } },
    @{Name="LicenseType";   Expression={ $_.LicenseType } },
    Department,
    @{Name="ArchiveMailboxEnabled"; Expression={ if ($_.ArchiveEnabled) { "Yes" } else { "No" } } }

# Specify output path for the CSV report
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$OutputPath = "$env:USERPROFILE\Desktop\InactiveM365Users_$timestamp.csv"
$InactiveUsersReport | Export-Csv -Path $OutputPath -NoTypeInformation
Write-Host "Report generated: $OutputPath (Total inactive users: $($InactiveUsersReport.Count))"
