# Script to generate a report of Office 365 users with license information
# Includes a column indicating if users are eligible for license removal based on inactivity

# Check if the Microsoft Graph PowerShell module is installed, if not, install it
if (-not (Get-Module -Name Microsoft.Graph -ListAvailable)) {
    Write-Host "Microsoft Graph PowerShell module not found. Installing..."
    Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
}

# Check if the Exchange Online PowerShell module is installed, if not, install it
if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    Write-Host "Exchange Online PowerShell module not found. Installing..."
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
}

# Import required modules
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Identity.SignIns
Import-Module Microsoft.Graph.Reports

# Ensure Exchange Online Management module is imported correctly
Write-Host "Importing Exchange Online Management module..."
try {
    # Force import the module to ensure it's properly loaded
    Import-Module ExchangeOnlineManagement -Force -DisableNameChecking
    
    # List the cmdlets from the ExchangeOnlineManagement module to verify
    $exchangeCmdlets = Get-Command -Module ExchangeOnlineManagement
    Write-Host "Successfully imported Exchange Online Management module with $($exchangeCmdlets.Count) cmdlets" -ForegroundColor Green
    
    # Specifically check for Set-Mailbox cmdlet
    $setMailboxCmd = Get-Command Set-Mailbox -ErrorAction SilentlyContinue
    if ($setMailboxCmd) {
        Write-Host "Set-Mailbox cmdlet is available" -ForegroundColor Green
    } else {
        Write-Host "Set-Mailbox cmdlet is not available in the current session. Will attempt to connect to Exchange Online to access it." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Warning: Error importing Exchange Online Management module: $_" -ForegroundColor Yellow
    Write-Host "Will attempt to connect to Exchange Online anyway." -ForegroundColor Yellow
}

# This function will use the basic remote PowerShell session to connect to Exchange Online
function Connect-BasicExchangeOnline {
    Write-Host "Attempting to establish a basic Exchange Online Remote PowerShell session..." -ForegroundColor Cyan
    
    # Get credentials - use the current user's credentials
    $credential = Get-Credential -Message "Enter your Exchange Online credentials"
    
    # Create session options
    $sessionOptions = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
    
    # Try to establish a session to Exchange Online
    try {
        # Create a remote PowerShell session to Exchange Online
        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credential -Authentication Basic -AllowRedirection -SessionOption $sessionOptions -ErrorAction Stop
        
        # Import the session
        Import-PSSession $session -DisableNameChecking -AllowClobber -ErrorAction Stop | Out-Null
        
        # Test if Set-Mailbox is now available
        $setMailboxCmd = Get-Command Set-Mailbox -ErrorAction SilentlyContinue
        if ($setMailboxCmd) {
            Write-Host "Basic Exchange Online Remote PowerShell connection established successfully. Set-Mailbox cmdlet is now available." -ForegroundColor Green
            return $true
        } else {
            Write-Host "Basic Exchange Online Remote PowerShell connection established, but Set-Mailbox cmdlet is still not available." -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "Failed to establish basic Exchange Online Remote PowerShell session: $_" -ForegroundColor Red
        return $false
    }
}

# Helper function to ensure Exchange Online is connected
function Ensure-ExchangeOnlineConnection {
    try {
        # Try to run a simple Exchange Online command to test the connection
        $null = Get-Mailbox -ResultSize 1 -ErrorAction Stop
        Write-Host "Exchange Online connection is active." -ForegroundColor Green
    }
    catch {
        Write-Host "Exchange Online connection is not active. Reconnecting..." -ForegroundColor Yellow
        try {
            # Disconnect any existing sessions first
            try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
            
            # Make sure the module is imported
            Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop
            
            # Connect to Exchange Online
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            
            # Verify that Set-Mailbox is available
            $command = Get-Command Set-Mailbox -ErrorAction SilentlyContinue
            if ($command -eq $null) {
                Write-Host "Set-Mailbox cmdlet is not available after connecting with modern method." -ForegroundColor Yellow
                Write-Host "Trying basic Exchange Online Remote PowerShell connection..." -ForegroundColor Yellow
                
                # Try the basic connection method
                $basicConnectionSuccess = Connect-BasicExchangeOnline
                
                if (-not $basicConnectionSuccess) {
                    Write-Host "Failed to establish a connection that provides the Set-Mailbox cmdlet." -ForegroundColor Red
                    Write-Host "Try running the script in a new PowerShell window, or manually connect to Exchange Online first." -ForegroundColor Yellow
                    return $false
                }
            }
            else {
                Write-Host "Successfully reconnected to Exchange Online." -ForegroundColor Green
            }
        }
        catch {
            Write-Host "Failed to connect to Exchange Online with modern method: $_" -ForegroundColor Yellow
            Write-Host "Trying basic Exchange Online Remote PowerShell connection..." -ForegroundColor Yellow
            
            # Try the basic connection method
            $basicConnectionSuccess = Connect-BasicExchangeOnline
            
            if (-not $basicConnectionSuccess) {
                Write-Host "Failed to establish any Exchange Online connection. Error: $_" -ForegroundColor Red
                Write-Host "Please run the script again or manually connect to Exchange Online before proceeding." -ForegroundColor Red
                return $false
            }
        }
    }
    return $true
}

# Helper function to ensure Microsoft Graph is connected
function Ensure-MgGraphConnection {
    try {
        # Try to run a simple Graph command to test the connection
        $null = Get-MgUser -Top 1 -ErrorAction Stop
        Write-Host "Microsoft Graph connection is active." -ForegroundColor Green
    }
    catch {
        Write-Host "Microsoft Graph connection is not active. Reconnecting..." -ForegroundColor Yellow
        try {
            # Reconnect to Microsoft Graph
            Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All", "Directory.Read.All", "Reports.Read.All" -ErrorAction Stop
            Write-Host "Successfully reconnected to Microsoft Graph." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to connect to Microsoft Graph: $_" -ForegroundColor Red
            Write-Host "Please run the script again or manually connect to Microsoft Graph before proceeding." -ForegroundColor Red
            return $false
        }
    }
    return $true
}

# Connect to Microsoft Graph with required permissions
Write-Host "Connecting to Microsoft Graph..."
Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All", "Directory.Read.All", "Reports.Read.All"

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..."
Connect-ExchangeOnline -ShowBanner:$false

# Set the date thresholds
$inactiveDate = (Get-Date).AddDays(-90)
$newUserDate = (Get-Date).AddDays(-30)

# Declare global variables to store results
$global:results = $null
$global:totalLicensed = 0
$global:sharedMailboxCount = 0
$global:eligibleCount = 0
$global:reportPath = ""

# Function to generate the initial report
function Generate-InactivityReport {
    Write-Host "Generating report for all licensed Office 365 users..."
    Write-Host "Users who haven't authenticated to Office 365 or Outlook since $($inactiveDate.ToString('yyyy-MM-dd')) and were created before $($newUserDate.ToString('yyyy-MM-dd')) will be marked as eligible for license removal."

    # Get all users with necessary properties
    $users = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, Mail, CreatedDateTime, SignInActivity, AssignedLicenses, AccountEnabled

    # Get mailbox information from Exchange Online
    Write-Host "Retrieving mailbox information from Exchange Online..."
    $mailboxes = @{}
    try {
        $exchangeMailboxes = Get-Mailbox -ResultSize Unlimited
        foreach ($mailbox in $exchangeMailboxes) {
            $mailboxes[$mailbox.UserPrincipalName.ToLower()] = @{
                RecipientTypeDetails = $mailbox.RecipientTypeDetails
                IsShared = ($mailbox.RecipientTypeDetails -eq 'SharedMailbox')
                RecipientType = $mailbox.RecipientType
                ExchangeGuid = $mailbox.ExchangeGuid
            }
        }
        Write-Host "Retrieved information for $($mailboxes.Count) mailboxes."
    }
    catch {
        Write-Warning "Unable to retrieve complete mailbox information: $_"
        Write-Host "Continuing with available data..."
    }

    # Get detailed sign-in logs for more comprehensive authentication data
    # Note: This API might return a lot of data; adjust the date range if needed
    $startDate = (Get-Date).AddDays(-120) # Looking back further to ensure we catch all data
    $endDate = Get-Date
    Write-Host "Retrieving sign-in logs from $($startDate.ToString('yyyy-MM-dd')) to $($endDate.ToString('yyyy-MM-dd'))..."

    # Define Office 365 and Outlook related app names to track
    # These are common app display names for Office 365 and Outlook-related services
    $office365Apps = @(
        "Office 365", 
        "Office 365 Exchange Online", 
        "Microsoft Office", 
        "Microsoft Office 365 Portal",
        "Microsoft Exchange Online",
        "Outlook",
        "Outlook Web Access",
        "Outlook Mobile",
        "Microsoft Outlook",
        "Exchange Online PowerShell",
        "Microsoft 365 admin center",
        "Exchange Online",
        "Outlook.com",
        "Exchange ActiveSync"
    )

    # Create a hashtable to store last authentication details for each user
    $userAuthDetails = @{}

    # Create a hashtable to specifically track Office 365 and Outlook authentications
    $officeOutlookAuth = @{}

    # Get sign-in logs and process them
    try {
        # Filter for successful sign-ins only (status code 0)
        $signInLogs = Get-MgAuditLogSignIn -Filter "createdDateTime ge $($startDate.ToString('yyyy-MM-ddTHH:mm:ssZ')) and createdDateTime le $($endDate.ToString('yyyy-MM-ddTHH:mm:ssZ')) and status/errorCode eq 0" -All
        
        Write-Host "Processing sign-in logs for Office 365 and Outlook app activity..."
        foreach ($log in $signInLogs) {
            $upn = $log.UserPrincipalName.ToLower()
            
            # Track all authentication activity
            if (!$userAuthDetails.ContainsKey($upn) -or $userAuthDetails[$upn].LastAuthTime -lt $log.CreatedDateTime) {
                $userAuthDetails[$upn] = @{
                    LastAuthTime = $log.CreatedDateTime
                    ClientAppUsed = $log.ClientAppUsed
                    AppDisplayName = $log.AppDisplayName
                    IPAddress = $log.IpAddress
                    DeviceDetail = $log.DeviceDetail.DisplayName
                    Location = if ($log.Location.City) { "$($log.Location.City), $($log.Location.CountryOrRegion)" } else { "Unknown" }
                    IsInteractive = $log.IsInteractive
                    AuthMethod = $log.AuthenticationDetails.AuthenticationMethod -join ", "
                }
            }
            
            # Specifically track Office 365 and Outlook authentication
            $isOfficeOrOutlookApp = $office365Apps -contains $log.AppDisplayName
            
            if ($isOfficeOrOutlookApp) {
                if (!$officeOutlookAuth.ContainsKey($upn) -or $officeOutlookAuth[$upn].LastAuthTime -lt $log.CreatedDateTime) {
                    $officeOutlookAuth[$upn] = @{
                        LastAuthTime = $log.CreatedDateTime
                        AppDisplayName = $log.AppDisplayName
                        ClientAppUsed = $log.ClientAppUsed
                        DeviceDetail = $log.DeviceDetail.DisplayName
                    }
                }
            }
        }
        
        Write-Host "Retrieved authentication details for $($userAuthDetails.Count) users."
        Write-Host "Found $($officeOutlookAuth.Count) users with Office 365 or Outlook app activity."
    } 
    catch {
        Write-Warning "Unable to retrieve detailed sign-in logs: $_"
        Write-Host "Continuing with basic SignInActivity data only..."
    }

    # Filter for users with licenses only
    $licensedUsers = $users | Where-Object { $_.AssignedLicenses.Count -gt 0 }

    Write-Host "Found $($licensedUsers.Count) licensed users..."

    # Create a results array with the information we want
    $global:results = $licensedUsers | ForEach-Object {
        $user = $_
        $upn = $user.UserPrincipalName.ToLower()
        
        # Get mailbox information if available
        $mailboxInfo = if ($mailboxes.ContainsKey($upn)) { $mailboxes[$upn] } else { $null }
        $isSharedMailbox = if ($mailboxInfo -ne $null) { $mailboxInfo.IsShared } else { $false }
        $recipientType = if ($mailboxInfo -ne $null) { $mailboxInfo.RecipientTypeDetails } else { "Unknown" }
        
        # Get basic authentication data
        $lastBasicAuth = "Never"
        $lastBasicAuthType = "N/A"
        
        if ($user.SignInActivity -ne $null) {
            if ($user.SignInActivity.LastSignInDateTime -ne $null) {
                $lastBasicAuth = $user.SignInActivity.LastSignInDateTime
                $lastBasicAuthType = "Interactive"
            } elseif ($user.SignInActivity.LastNonInteractiveSignInDateTime -ne $null) {
                if ($lastBasicAuth -eq "Never" -or $user.SignInActivity.LastNonInteractiveSignInDateTime -gt $lastBasicAuth) {
                    $lastBasicAuth = $user.SignInActivity.LastNonInteractiveSignInDateTime
                    $lastBasicAuthType = "Non-Interactive"
                }
            }
        }
        
        # Get detailed authentication data if available
        $hasDetailedAuth = $userAuthDetails.ContainsKey($upn)
        $authDetail = if ($hasDetailedAuth) { $userAuthDetails[$upn] } else { $null }
        
        # Get Office 365/Outlook specific auth data
        $hasOfficeAuth = $officeOutlookAuth.ContainsKey($upn)
        $officeAuthDetail = if ($hasOfficeAuth) { $officeOutlookAuth[$upn] } else { $null }
        
        # Determine the most recent and accurate authentication data to report
        $reportAuthTime = $lastBasicAuth
        $reportAuthType = $lastBasicAuthType
        $clientApp = "Unknown"
        $appName = "Unknown"
        $ipAddress = "Unknown"
        $deviceName = "Unknown"
        $location = "Unknown"
        $authMethod = "Unknown"
        
        if ($hasDetailedAuth) {
            if ($lastBasicAuth -eq "Never" -or ($authDetail.LastAuthTime -gt $lastBasicAuth)) {
                $reportAuthTime = $authDetail.LastAuthTime
                $reportAuthType = if ($authDetail.IsInteractive) { "Interactive" } else { "Non-Interactive" }
                $clientApp = $authDetail.ClientAppUsed
                $appName = $authDetail.AppDisplayName
                $ipAddress = $authDetail.IPAddress
                $deviceName = if ($authDetail.DeviceDetail) { $authDetail.DeviceDetail } else { "Unknown" }
                $location = $authDetail.Location
                $authMethod = if ($authDetail.AuthMethod) { $authDetail.AuthMethod } else { "Unknown" }
            }
        }
        
        # Get Office 365/Outlook specific data
        $lastOfficeAuth = "Never"
        $lastOfficeApp = "N/A"
        $officeClientApp = "N/A"
        $officeDevice = "N/A"
        
        if ($hasOfficeAuth) {
            $lastOfficeAuth = $officeAuthDetail.LastAuthTime
            $lastOfficeApp = $officeAuthDetail.AppDisplayName
            $officeClientApp = $officeAuthDetail.ClientAppUsed
            $officeDevice = if ($officeAuthDetail.DeviceDetail) { $officeAuthDetail.DeviceDetail } else { "Unknown" }
        }
        
        # Calculate days since last authentication for all methods
        $daysSinceAuth = if ($reportAuthTime -eq "Never") { 999 } else { ((Get-Date) - $reportAuthTime).Days }
        $daysSinceOfficeAuth = if ($lastOfficeAuth -eq "Never") { 999 } else { ((Get-Date) - $lastOfficeAuth).Days }
        
        # Calculate days since account creation
        $daysSinceCreation = ((Get-Date) - $user.CreatedDateTime).Days
        
        # Determine if the user is eligible for license removal based on:
        # 1. No authentication to Office/Outlook apps in last 90 days
        # 2. Account is at least 30 days old (to avoid removing licenses from new users)
        # 3. Not already a shared mailbox
        $isEligibleForRemoval = ($daysSinceOfficeAuth -ge 90) -and ($daysSinceCreation -ge 30) -and (-not $isSharedMailbox)
        
        # Get licenses in a readable format
        $licenseInfo = $user.AssignedLicenses | ForEach-Object {
            $licenseId = $_.SkuId
            # You can map license SKU IDs to readable names here if needed
            $licenseId
        }
        
        [PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            DisplayName = $user.DisplayName
            Email = $user.Mail
            MailboxType = $recipientType
            IsSharedMailbox = $isSharedMailbox
            AccountEnabled = $user.AccountEnabled
            LastAuthentication = $reportAuthTime
            AuthenticationType = $reportAuthType
            DaysSinceLastAuthentication = $daysSinceAuth
            LastOfficeAuthentication = $lastOfficeAuth
            DaysSinceLastOfficeAuthentication = $daysSinceOfficeAuth
            LastOfficeApp = $lastOfficeApp
            ClientApp = $clientApp
            OfficeClientApp = $officeClientApp
            AppName = $appName
            IPAddress = $ipAddress
            DeviceName = $deviceName
            OfficeDevice = $officeDevice
            Location = $location
            AuthMethod = $authMethod
            CreatedDate = $user.CreatedDateTime
            DaysSinceCreation = $daysSinceCreation
            LicensePlans = ($licenseInfo -join ", ")
            EligibleForLicenseRemoval = $isEligibleForRemoval
        }
    }

    # Count eligible users
    $global:eligibleCount = ($global:results | Where-Object { $_.EligibleForLicenseRemoval -eq $true }).Count
    $global:totalLicensed = $global:results.Count
    $global:sharedMailboxCount = ($global:results | Where-Object { $_.IsSharedMailbox -eq $true }).Count

    # Export the results to CSV
    $dateStamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $global:reportPath = "$PSScriptRoot\LicensedUsers_$dateStamp.csv"
    $global:results | Export-Csv -Path $global:reportPath -NoTypeInformation

    Write-Host "Report exported to $global:reportPath"
    Write-Host "Total licensed users: $global:totalLicensed"
    Write-Host "Shared mailboxes with licenses: $global:sharedMailboxCount"
    Write-Host "Users eligible for license removal: $global:eligibleCount ($([math]::Round(($global:eligibleCount/$global:totalLicensed)*100))%)"
    
    return $true
}

# Interactive Menu Functions
function Show-MainMenu {
    Clear-Host
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "        Microsoft 365 License Management Tool    " -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "1: Create report of eligible users"
    Write-Host "2: Convert eligible users to shared mailboxes"
    Write-Host "3: Set receive limit to 0KB for converted mailboxes"
    Write-Host "4: Block sign-in for converted shared mailboxes"
    Write-Host "5: Remove licenses from converted users"
    Write-Host "6: Remove roles from converted users"
    Write-Host "7: Disable on-premises AD accounts (hybrid environments)"
    Write-Host "8: Create before/after comparison report"
    Write-Host "Q: Quit"
    Write-Host "============================================" -ForegroundColor Cyan
    
    $choice = Read-Host "Enter your choice"
    return $choice
}

function Convert-ToSharedMailbox {
    if (-not (Ensure-ExchangeOnlineConnection)) {
        Write-Host "Cannot proceed without Exchange Online connection." -ForegroundColor Red
        return
    }
    
    $eligibleUsers = $global:results | Where-Object { $_.EligibleForLicenseRemoval -eq $true }
    
    if ($eligibleUsers.Count -eq 0) {
        Write-Host "No eligible users found for conversion to shared mailboxes." -ForegroundColor Yellow
        return
    }
    
    Write-Host "Found $($eligibleUsers.Count) eligible users that can be converted to shared mailboxes." -ForegroundColor Cyan
    $confirmation = Read-Host "Do you want to convert these mailboxes to shared mailboxes? (Y/N)"
    
    if ($confirmation -ne "Y" -and $confirmation -ne "y") {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        return
    }
    
    $conversionResults = @()
    
    # Initialize progress bar
    $progressCounter = 0
    $totalUsers = $eligibleUsers.Count

    # Get all mailboxes to check their types
    Write-Host "Retrieving all mailboxes to check for special types..."
    $allMailboxes = @{}
    try {
        $exchangeMailboxes = Get-Mailbox -ResultSize Unlimited
        foreach ($mailbox in $exchangeMailboxes) {
            $allMailboxes[$mailbox.UserPrincipalName.ToLower()] = @{
                RecipientTypeDetails = $mailbox.RecipientTypeDetails
                ExchangeGuid = $mailbox.ExchangeGuid
                Identity = $mailbox.Identity
            }
        }
        Write-Host "Retrieved information for $($allMailboxes.Count) mailboxes." -ForegroundColor Green
    }
    catch {
        Write-Warning "Unable to retrieve complete mailbox information: $_"
        Write-Host "Will check mailbox types individually during conversion." -ForegroundColor Yellow
    }
    
    # Define special mailbox types that should be skipped
    $specialMailboxTypes = @(
        "SchedulingMailbox",
        "TeamMailbox", 
        "DiscoveryMailbox", 
        "RoomMailbox", 
        "EquipmentMailbox",
        "LinkedMailbox",
        "LinkedRoomMailbox",
        "AuditLogMailbox",
        "AuxAuditLogMailbox",
        "GroupMailbox",
        "SupervisoryReviewPolicyMailbox"
    )
    
    # Patterns in email addresses or display names that suggest special mailboxes
    $specialMailboxPatterns = @(
        "autoattendant",
        "auto attendant",
        "aa@",
        "phone system",
        "pabx",
        "queue",
        "resource",
        "room",
        "meeting"
    )
    
    foreach ($user in $eligibleUsers) {
        # Update progress bar
        $progressCounter++
        $percentComplete = ($progressCounter / $totalUsers) * 100
        Write-Progress -Activity "Converting user mailboxes to shared mailboxes" -Status "Processing $progressCounter of $totalUsers - $($user.DisplayName)" -PercentComplete $percentComplete
        
        Write-Host "Converting $($user.DisplayName) ($($user.UserPrincipalName)) to shared mailbox..." -NoNewline
        
        # Check if this is a special mailbox type that should be skipped
        $skipMailbox = $false
        $skipReason = ""
        
        # Check if it's already a special mailbox type
        $upn = $user.UserPrincipalName.ToLower()
        if ($allMailboxes.ContainsKey($upn)) {
            $mailboxInfo = $allMailboxes[$upn]
            if ($specialMailboxTypes -contains $mailboxInfo.RecipientTypeDetails) {
                $skipMailbox = $true
                $skipReason = "Special mailbox type: $($mailboxInfo.RecipientTypeDetails)"
            }
        }
        
        # Check for special patterns in the email or display name
        if (-not $skipMailbox) {
            foreach ($pattern in $specialMailboxPatterns) {
                if ($user.UserPrincipalName -match $pattern -or $user.DisplayName -match $pattern) {
                    $skipMailbox = $true
                    $skipReason = "Matches special mailbox pattern: $pattern"
                    break
                }
            }
        }
        
        # Skip this mailbox if needed
        if ($skipMailbox) {
            Write-Host "Skipped" -ForegroundColor Yellow
            Write-Host "  Reason: $skipReason" -ForegroundColor Yellow
            
            $conversionResults += [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                DisplayName = $user.DisplayName
                ConversionStatus = "Skipped"
                Error = $skipReason
            }
            continue
        }
        
        try {
            # First attempt - using standard cmdlet
            try {
                Set-Mailbox -Identity $user.UserPrincipalName -Type Shared -ErrorAction Stop
                
                Write-Host "Success" -ForegroundColor Green
                $conversionResults += [PSCustomObject]@{
                    UserPrincipalName = $user.UserPrincipalName
                    DisplayName = $user.DisplayName
                    ConversionStatus = "Success"
                    Error = ""
                }
            }
            catch {
                Write-Host "Trying alternative method..." -ForegroundColor Yellow
                
                # Alternative approach - for when the first method doesn't work
                # This creates a script block to execute the command
                $command = "Set-Mailbox -Identity '$($user.UserPrincipalName)' -Type Shared"
                $scriptBlock = [ScriptBlock]::Create($command)
                
                try {
                    & $scriptBlock
                    
                    Write-Host "Success" -ForegroundColor Green
                    $conversionResults += [PSCustomObject]@{
                        UserPrincipalName = $user.UserPrincipalName
                        DisplayName = $user.DisplayName
                        ConversionStatus = "Success (Alternative Method)"
                        Error = ""
                    }
                }
                catch {
                    Write-Host "Failed" -ForegroundColor Red
                    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
                    
                    $conversionResults += [PSCustomObject]@{
                        UserPrincipalName = $user.UserPrincipalName
                        DisplayName = $user.DisplayName
                        ConversionStatus = "Failed"
                        Error = $_.Exception.Message
                    }
                }
            }
        }
        catch {
            Write-Host "Failed" -ForegroundColor Red
            Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
            
            $conversionResults += [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                DisplayName = $user.DisplayName
                ConversionStatus = "Failed"
                Error = $_.Exception.Message
            }
        }
    }
    
    # Export conversion results
    $conversionLogPath = "$PSScriptRoot\ConversionToSharedMailbox_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $conversionResults | Export-Csv -Path $conversionLogPath -NoTypeInformation
    Write-Host "Conversion log saved to: $conversionLogPath" -ForegroundColor Cyan
    
    # Update the global results variable with new mailbox types
    foreach ($user in $global:results) {
        $convertedUser = $conversionResults | Where-Object { $_.UserPrincipalName -eq $user.UserPrincipalName -and $_.ConversionStatus -eq "Success" }
        if ($convertedUser) {
            $user.MailboxType = "SharedMailbox"
        }
    }
}

function Set-ZeroReceiveLimit {
    # Ensure Exchange Online is connected before proceeding
    if (-not (Ensure-ExchangeOnlineConnection)) {
        return
    }
    
    $convertedUsers = $global:results | Where-Object { $_.MailboxType -eq "SharedMailbox" -and $_.EligibleForLicenseRemoval -eq $true }
    
    if ($convertedUsers.Count -eq 0) {
        Write-Host "No converted shared mailboxes found to set receive limits." -ForegroundColor Yellow
        return
    }
    
    Write-Host "Found $($convertedUsers.Count) converted shared mailboxes to set receive limits." -ForegroundColor Cyan
    
    # Create a log file for the process
    $limitLogPath = "$PSScriptRoot\ReceiveLimitSet_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $limitResults = @()
    
    # Initialize progress bar
    $progressCounter = 0
    $totalUsers = $convertedUsers.Count
    
    foreach ($user in $convertedUsers) {
        # Update progress bar
        $progressCounter++
        $percentComplete = ($progressCounter / $totalUsers) * 100
        Write-Progress -Activity "Setting receive limit to 0KB" -Status "Processing $progressCounter of $totalUsers - $($user.DisplayName)" -PercentComplete $percentComplete
        
        Write-Host "Setting receive limit to 0KB for $($user.DisplayName) ($($user.UserPrincipalName))..." -NoNewline
        
        try {
            # First attempt - using standard cmdlet
            try {
                # Set the receive limit to 0KB
                Set-Mailbox -Identity $user.UserPrincipalName -MaxReceiveSize 0 -ErrorAction Stop
                Write-Host "Success" -ForegroundColor Green
                
                $limitResults += [PSCustomObject]@{
                    UserPrincipalName = $user.UserPrincipalName
                    DisplayName = $user.DisplayName
                    Status = "Success"
                    ErrorMessage = ""
                    ProcessDate = Get-Date
                }
            }
            catch {
                # If standard cmdlet fails, try alternative approach
                Write-Host "Trying alternative method..." -ForegroundColor Yellow
                
                # Create a script block to execute the command
                $scriptBlock = {
                    param($upn)
                    Set-Mailbox -Identity $upn -MaxReceiveSize 0
                }
                
                # Execute the script block in the current session
                $result = & $scriptBlock $user.UserPrincipalName 2>&1
                
                if ($result -is [System.Management.Automation.ErrorRecord]) {
                    throw $result
                }
                
                Write-Host "Success (alternative method)" -ForegroundColor Green
                
                $limitResults += [PSCustomObject]@{
                    UserPrincipalName = $user.UserPrincipalName
                    DisplayName = $user.DisplayName
                    Status = "Success"
                    ErrorMessage = ""
                    ProcessDate = Get-Date
                }
            }
        }
        catch {
            Write-Host "Failed" -ForegroundColor Red
            Write-Host "Error: $_" -ForegroundColor Red
            
            $limitResults += [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                DisplayName = $user.DisplayName
                Status = "Failed"
                ErrorMessage = $_.ToString()
                ProcessDate = Get-Date
            }
        }
    }
    
    # Complete the progress bar
    Write-Progress -Activity "Setting receive limit to 0KB" -Completed
    
    # Export results
    $limitResults | Export-Csv -Path $limitLogPath -NoTypeInformation
    Write-Host "Receive limit setting log saved to: $limitLogPath" -ForegroundColor Cyan
}

function Block-SharedMailboxSignIn {
    if (-not (Ensure-MgGraphConnection)) {
        Write-Host "Cannot proceed without Microsoft Graph connection." -ForegroundColor Red
        return
    }
    
    $convertedUsers = $global:results | Where-Object { $_.MailboxType -eq "SharedMailbox" -and $_.EligibleForLicenseRemoval -eq $true }
    
    if ($convertedUsers.Count -eq 0) {
        Write-Host "No converted shared mailboxes found to block sign-in." -ForegroundColor Yellow
        return
    }
    
    Write-Host "Found $($convertedUsers.Count) shared mailboxes that can have sign-in blocked." -ForegroundColor Cyan
    $confirmation = Read-Host "Do you want to block sign-in for these mailboxes? (Y/N)"
    
    if ($confirmation -ne "Y" -and $confirmation -ne "y") {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        return
    }
    
    $blockResults = @()
    
    # Initialize progress bar
    $progressCounter = 0
    $totalUsers = $convertedUsers.Count
    
    foreach ($user in $convertedUsers) {
        # Update progress bar
        $progressCounter++
        $percentComplete = ($progressCounter / $totalUsers) * 100
        Write-Progress -Activity "Blocking sign-in for shared mailboxes" -Status "Processing $progressCounter of $totalUsers - $($user.DisplayName)" -PercentComplete $percentComplete
        
        Write-Host "Blocking sign-in for $($user.DisplayName) ($($user.UserPrincipalName))..." -NoNewline
        
        try {
            # Get the user to update
            $mgUser = Get-MgUser -UserId $user.UserPrincipalName -ErrorAction Stop
            
            # Block sign-in by setting AccountEnabled to false
            Update-MgUser -UserId $mgUser.Id -AccountEnabled:$false -ErrorAction Stop
            
            Write-Host "Success" -ForegroundColor Green
            $blockResults += [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                DisplayName = $user.DisplayName
                Status = "Success"
                Error = ""
            }
            
            # Update the global results array to mark account as disabled
            try {
                $userObj = $global:results | Where-Object { $_.UserPrincipalName -eq $user.UserPrincipalName }
                if ($userObj) {
                    $userObj.AccountEnabled = $false
                }
            }
            catch {
                Write-Warning "Unable to update account status in results array: $_"
            }
        }
        catch {
            Write-Host "Failed" -ForegroundColor Red
            Write-Host "  Error: $_" -ForegroundColor Red
            
            $blockResults += [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                DisplayName = $user.DisplayName
                Status = "Failed"
                Error = $_.Exception.Message
            }
        }
    }
    
    # Complete the progress bar
    Write-Progress -Activity "Blocking sign-in for shared mailboxes" -Completed
    
    # Export results to CSV
    $blockLogPath = "$PSScriptRoot\SignInBlock_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $blockResults | Export-Csv -Path $blockLogPath -NoTypeInformation
    
    # Summary
    $successCount = ($blockResults | Where-Object { $_.Status -eq "Success" }).Count
    $failedCount = ($blockResults | Where-Object { $_.Status -eq "Failed" }).Count
    
    Write-Host "Sign-in blocking completed." -ForegroundColor Cyan
    Write-Host "Successfully blocked: $successCount" -ForegroundColor Green
    if ($failedCount -gt 0) {
        Write-Host "Failed to block: $failedCount" -ForegroundColor Red
    }
    Write-Host "Block log saved to: $blockLogPath" -ForegroundColor Cyan
}

function Remove-UserLicenses {
    # Ensure Microsoft Graph is connected before proceeding
    if (-not (Ensure-MgGraphConnection)) {
        return
    }
    
    $convertedUsers = $global:results | Where-Object { $_.MailboxType -eq "SharedMailbox" -and $_.EligibleForLicenseRemoval -eq $true }
    
    if ($convertedUsers.Count -eq 0) {
        Write-Host "No converted shared mailboxes found for license removal." -ForegroundColor Yellow
        return
    }
    
    Write-Host "Found $($convertedUsers.Count) converted shared mailboxes to remove licenses." -ForegroundColor Cyan
    
    # Create a log file for the process
    $licenseLogPath = "$PSScriptRoot\LicenseRemoval_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $licenseResults = @()
    
    # Initialize progress bar
    $progressCounter = 0
    $totalUsers = $convertedUsers.Count
    
    foreach ($user in $convertedUsers) {
        # Update progress bar
        $progressCounter++
        $percentComplete = ($progressCounter / $totalUsers) * 100
        Write-Progress -Activity "Removing licenses" -Status "Processing $progressCounter of $totalUsers - $($user.DisplayName)" -PercentComplete $percentComplete
        
        Write-Host "Removing licenses from $($user.DisplayName) ($($user.UserPrincipalName))..." -NoNewline
        
        try {
            # Get user's current license information
            $mgUser = Get-MgUser -UserId $user.UserPrincipalName -Property AssignedLicenses
            
            if ($mgUser.AssignedLicenses.Count -gt 0) {
                # Create a license removal object
                $licensesToRemove = @{
                    AddLicenses = @()
                    RemoveLicenses = $mgUser.AssignedLicenses.SkuId
                }
                
                # Remove all licenses
                Set-MgUserLicense -UserId $user.UserPrincipalName -BodyParameter $licensesToRemove -ErrorAction Stop
                Write-Host "Success" -ForegroundColor Green
                
                $licenseResults += [PSCustomObject]@{
                    UserPrincipalName = $user.UserPrincipalName
                    DisplayName = $user.DisplayName
                    PreviousLicensesCount = $mgUser.AssignedLicenses.Count
                    Status = "Success"
                    ErrorMessage = ""
                    ProcessDate = Get-Date
                }
            }
            else {
                Write-Host "No licenses found" -ForegroundColor Yellow
                
                $licenseResults += [PSCustomObject]@{
                    UserPrincipalName = $user.UserPrincipalName
                    DisplayName = $user.DisplayName
                    PreviousLicensesCount = 0
                    Status = "No Licenses"
                    ErrorMessage = ""
                    ProcessDate = Get-Date
                }
            }
        }
        catch {
            Write-Host "Failed" -ForegroundColor Red
            Write-Host "Error: $_" -ForegroundColor Red
            
            $licenseResults += [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                DisplayName = $user.DisplayName
                PreviousLicensesCount = $mgUser.AssignedLicenses.Count
                Status = "Failed"
                ErrorMessage = $_.ToString()
                ProcessDate = Get-Date
            }
        }
    }
    
    # Complete the progress bar
    Write-Progress -Activity "Removing licenses" -Completed
    
    # Export results
    $licenseResults | Export-Csv -Path $licenseLogPath -NoTypeInformation
    Write-Host "License removal log saved to: $licenseLogPath" -ForegroundColor Cyan
    
    # Update the global results variable with new license status
    foreach ($user in $global:results) {
        $updatedUser = $licenseResults | Where-Object { $_.UserPrincipalName -eq $user.UserPrincipalName -and $_.Status -eq "Success" }
        if ($updatedUser) {
            try {
                # Check if the property exists before trying to set it
                $userProps = $user | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name
                if ($userProps -contains "LicensePlans") {
                    $user.LicensePlans = "Removed"
                }
                else {
                    # Add the property if it doesn't exist
                    Add-Member -InputObject $user -MemberType NoteProperty -Name "LicensePlans" -Value "Removed" -Force
                }
            }
            catch {
                Write-Warning "Unable to update license status for $($user.UserPrincipalName): $_"
            }
        }
    }
}

function Remove-UserRoles {
    # Ensure Microsoft Graph is connected before proceeding
    if (-not (Ensure-MgGraphConnection)) {
        Write-Host "Cannot proceed without Microsoft Graph connection." -ForegroundColor Red
        return
    }
    
    $convertedUsers = $global:results | Where-Object { $_.MailboxType -eq "SharedMailbox" -and $_.EligibleForLicenseRemoval -eq $true }
    
    if ($convertedUsers.Count -eq 0) {
        Write-Host "No converted shared mailboxes found for role removal." -ForegroundColor Yellow
        return
    }
    
    Write-Host "Found $($convertedUsers.Count) shared mailboxes that may have roles to remove." -ForegroundColor Cyan
    $confirmation = Read-Host "Do you want to check and remove roles from these users? (Y/N)"
    
    if ($confirmation -ne "Y" -and $confirmation -ne "y") {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        return
    }
    
    $roleResults = @()
    
    # Initialize progress bar
    $progressCounter = 0
    $totalUsers = $convertedUsers.Count
    
    foreach ($user in $convertedUsers) {
        # Update progress bar
        $progressCounter++
        $percentComplete = ($progressCounter / $totalUsers) * 100
        Write-Progress -Activity "Removing roles from users" -Status "Processing $progressCounter of $totalUsers - $($user.DisplayName)" -PercentComplete $percentComplete
        
        Write-Host "Checking roles for $($user.DisplayName) ($($user.UserPrincipalName))..." -NoNewline
        
        try {
            # Get all directory role assignments for this user
            $userRoles = Get-MgUserMemberOf -UserId $user.UserPrincipalName -All
            $directoryRoles = $userRoles | Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.directoryRole' }
            
            if ($directoryRoles.Count -eq 0) {
                Write-Host "No roles found" -ForegroundColor Green
                $roleResults += [PSCustomObject]@{
                    UserPrincipalName = $user.UserPrincipalName
                    DisplayName = $user.DisplayName
                    Status = "No roles found"
                    RolesRemoved = ""
                    Error = ""
                }
                continue
            }
            
            Write-Host "Found $($directoryRoles.Count) roles" -ForegroundColor Yellow
            
            # List the roles
            $roleNames = @()
            foreach ($role in $directoryRoles) {
                $roleName = $role.AdditionalProperties.displayName
                $roleNames += $roleName
                Write-Host "  - $roleName" -ForegroundColor Yellow
            }
            
            # Remove the user from each role
            $rolesRemoved = @()
            $errors = @()
            
            foreach ($role in $directoryRoles) {
                $roleId = $role.Id
                $roleName = $role.AdditionalProperties.displayName
                
                try {
                    # Remove user from role
                    Write-Host "  Removing from role: $roleName..." -NoNewline
                    Remove-MgDirectoryRoleMemberByRef -DirectoryRoleId $roleId -DirectoryObjectId $user.Id -ErrorAction Stop
                    Write-Host "Success" -ForegroundColor Green
                    $rolesRemoved += $roleName
                }
                catch {
                    Write-Host "Failed" -ForegroundColor Red
                    $errors += "Failed to remove from ${roleName}: $($_.Exception.Message)"
                    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            
            $status = if ($errors.Count -eq 0) { "Success" } else { "Partial Success" }
            
            $roleResults += [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                DisplayName = $user.DisplayName
                Status = $status
                RolesRemoved = ($rolesRemoved -join "; ")
                Error = ($errors -join "; ")
            }
        }
        catch {
            Write-Host "Failed" -ForegroundColor Red
            Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
            
            $roleResults += [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                DisplayName = $user.DisplayName
                Status = "Failed"
                RolesRemoved = ""
                Error = $_.Exception.Message
            }
        }
    }
    
    # Complete the progress bar
    Write-Progress -Activity "Removing roles from users" -Completed
    
    # Export results to CSV
    $roleLogPath = "$PSScriptRoot\RoleRemoval_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $roleResults | Export-Csv -Path $roleLogPath -NoTypeInformation
    
    # Summary
    $successCount = ($roleResults | Where-Object { $_.Status -eq "Success" }).Count
    $partialSuccessCount = ($roleResults | Where-Object { $_.Status -eq "Partial Success" }).Count
    $failedCount = ($roleResults | Where-Object { $_.Status -eq "Failed" }).Count
    $noRolesCount = ($roleResults | Where-Object { $_.Status -eq "No roles found" }).Count
    
    Write-Host "Role removal completed." -ForegroundColor Cyan
    Write-Host "Users with no roles: $noRolesCount" -ForegroundColor Green
    Write-Host "Successful role removals: $successCount" -ForegroundColor Green
    if ($partialSuccessCount -gt 0) {
        Write-Host "Partial successful role removals: $partialSuccessCount" -ForegroundColor Yellow
    }
    if ($failedCount -gt 0) {
        Write-Host "Failed role removals: $failedCount" -ForegroundColor Red
    }
    Write-Host "Role removal log saved to: $roleLogPath" -ForegroundColor Cyan
}

function Disable-OnPremADAccounts {
    Write-Host "This function will help disable on-premises Active Directory accounts for converted shared mailboxes." -ForegroundColor Cyan
    Write-Host "This is intended for hybrid environments where users exist both in Azure AD and on-premises AD." -ForegroundColor Cyan
    Write-Host ""
    Write-Host "IMPORTANT: This requires:" -ForegroundColor Yellow
    Write-Host "1. The Active Directory PowerShell module" -ForegroundColor Yellow
    Write-Host "2. Appropriate permissions to modify AD accounts" -ForegroundColor Yellow
    Write-Host "3. Connection to your on-premises domain" -ForegroundColor Yellow
    Write-Host ""
    
    # Check if Active Directory module is available
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-Host "The Active Directory PowerShell module is not installed." -ForegroundColor Red
        Write-Host "Please install it by running: Install-WindowsFeature RSAT-AD-PowerShell" -ForegroundColor Yellow
        return
    }
    
    # Try to import the module
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-Host "Successfully loaded the Active Directory module." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to import the Active Directory module: $_" -ForegroundColor Red
        Write-Host "Please ensure you have the RSAT AD tools installed and try again." -ForegroundColor Yellow
        return
    }
    
    # Prompt for domain info
    $defaultDomain = $env:USERDNSDOMAIN
    if (-not $defaultDomain) {
        $defaultDomain = "yourdomain.com"
    }
    
    $domain = Read-Host "Enter your Active Directory domain (default: $defaultDomain)"
    if ([string]::IsNullOrWhiteSpace($domain)) {
        $domain = $defaultDomain
    }
    
    # Check connection to domain
    try {
        $null = Get-ADDomain -Identity $domain -ErrorAction Stop
        Write-Host "Successfully connected to domain: $domain" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to connect to Active Directory domain $domain`: $_" -ForegroundColor Red
        Write-Host "Please check your domain name and connection to the domain." -ForegroundColor Yellow
        return
    }
    
    # Get the list of converted mailboxes
    $convertedUsers = $global:results | Where-Object { $_.MailboxType -eq "SharedMailbox" -and $_.EligibleForLicenseRemoval -eq $true }
    
    if ($convertedUsers.Count -eq 0) {
        Write-Host "No converted shared mailboxes found to disable in on-premises AD." -ForegroundColor Yellow
        return
    }
    
    Write-Host "Found $($convertedUsers.Count) shared mailboxes that could be disabled in on-premises AD." -ForegroundColor Cyan
    $confirmation = Read-Host "Do you want to disable these accounts in on-premises AD? (Y/N)"
    
    if ($confirmation -ne "Y" -and $confirmation -ne "y") {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        return
    }
    
    $adResults = @()
    
    # Initialize progress bar
    $progressCounter = 0
    $totalUsers = $convertedUsers.Count
    
    foreach ($user in $convertedUsers) {
        # Update progress bar
        $progressCounter++
        $percentComplete = ($progressCounter / $totalUsers) * 100
        Write-Progress -Activity "Disabling on-premises AD accounts" -Status "Processing $progressCounter of $totalUsers - $($user.DisplayName)" -PercentComplete $percentComplete
        
        # Extract username from UPN (remove domain part)
        $upnParts = $user.UserPrincipalName -split '@'
        $samAccountName = $upnParts[0]
        
        Write-Host "Looking for AD account for $($user.DisplayName) (SamAccountName: $samAccountName)..." -NoNewline
        
        try {
            # Try to find the AD account
            $adUser = Get-ADUser -Filter "SamAccountName -eq '$samAccountName'" -ErrorAction Stop
            
            if ($adUser) {
                Write-Host "Found" -ForegroundColor Green
                
                # Check if already disabled
                if (-not $adUser.Enabled) {
                    Write-Host "  Account is already disabled." -ForegroundColor Yellow
                    
                    $adResults += [PSCustomObject]@{
                        UserPrincipalName = $user.UserPrincipalName
                        DisplayName = $user.DisplayName
                        SamAccountName = $samAccountName
                        Status = "Already Disabled"
                        Error = ""
                    }
                    continue
                }
                
                # Disable the account
                Write-Host "  Disabling account..." -NoNewline
                try {
                    Disable-ADAccount -Identity $adUser -ErrorAction Stop
                    Write-Host "Success" -ForegroundColor Green
                    
                    $adResults += [PSCustomObject]@{
                        UserPrincipalName = $user.UserPrincipalName
                        DisplayName = $user.DisplayName
                        SamAccountName = $samAccountName
                        Status = "Success"
                        Error = ""
                    }
                }
                catch {
                    Write-Host "Failed" -ForegroundColor Red
                    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
                    
                    $adResults += [PSCustomObject]@{
                        UserPrincipalName = $user.UserPrincipalName
                        DisplayName = $user.DisplayName
                        SamAccountName = $samAccountName
                        Status = "Failed"
                        Error = $_.Exception.Message
                    }
                }
            }
            else {
                Write-Host "Not found" -ForegroundColor Yellow
                
                $adResults += [PSCustomObject]@{
                    UserPrincipalName = $user.UserPrincipalName
                    DisplayName = $user.DisplayName
                    SamAccountName = $samAccountName
                    Status = "Not Found"
                    Error = "AD account not found"
                }
            }
        }
        catch {
            Write-Host "Error" -ForegroundColor Red
            Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
            
            $adResults += [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                DisplayName = $user.DisplayName
                SamAccountName = $samAccountName
                Status = "Error"
                Error = $_.Exception.Message
            }
        }
    }
    
    # Complete the progress bar
    Write-Progress -Activity "Disabling on-premises AD accounts" -Completed
    
    # Export results to CSV
    $adLogPath = "$PSScriptRoot\ADAccountDisable_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $adResults | Export-Csv -Path $adLogPath -NoTypeInformation
    
    # Summary
    $successCount = ($adResults | Where-Object { $_.Status -eq "Success" }).Count
    $alreadyDisabledCount = ($adResults | Where-Object { $_.Status -eq "Already Disabled" }).Count
    $notFoundCount = ($adResults | Where-Object { $_.Status -eq "Not Found" }).Count
    $failedCount = ($adResults | Where-Object { $_.Status -eq "Failed" -or $_.Status -eq "Error" }).Count
    
    Write-Host "AD account disabling completed." -ForegroundColor Cyan
    Write-Host "Successfully disabled: $successCount" -ForegroundColor Green
    Write-Host "Already disabled accounts: $alreadyDisabledCount" -ForegroundColor Green
    if ($notFoundCount -gt 0) {
        Write-Host "Accounts not found in AD: $notFoundCount" -ForegroundColor Yellow
    }
    if ($failedCount -gt 0) {
        Write-Host "Failed operations: $failedCount" -ForegroundColor Red
    }
    Write-Host "AD account disable log saved to: $adLogPath" -ForegroundColor Cyan
}

function Create-ComparisonReport {
    $beforeAfterReportPath = "$PSScriptRoot\BeforeAfterComparison_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    Write-Host "Creating before/after comparison report..." -ForegroundColor Cyan
    
    $comparisonResults = @()
    
    # Get eligible users for the report
    $eligibleUsers = $global:results | Where-Object { $_.EligibleForLicenseRemoval -eq $true }
    
    # Initialize progress bar
    $progressCounter = 0
    $totalUsers = $eligibleUsers.Count
    
    foreach ($user in $eligibleUsers) {
        # Update progress bar
        $progressCounter++
        $percentComplete = ($progressCounter / $totalUsers) * 100
        Write-Progress -Activity "Creating comparison report" -Status "Processing $progressCounter of $totalUsers - $($user.DisplayName)" -PercentComplete $percentComplete
        
        $mailboxType = $user.MailboxType
        
        # Check if LicensePlans property exists
        $licenseStatus = "Active"
        try {
            $userProps = $user | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name
            if ($userProps -contains "LicensePlans") {
                $licenseStatus = if ($user.LicensePlans -eq "Removed") { "Removed" } else { "Active" }
            }
        }
        catch {
            Write-Warning "Unable to check license status for $($user.UserPrincipalName): $_"
        }
        
        $comparisonResults += [PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            DisplayName = $user.DisplayName
            OriginalMailboxType = "UserMailbox"
            CurrentMailboxType = $mailboxType
            OriginalLicenseStatus = "Licensed"
            CurrentLicenseStatus = $licenseStatus
            LastAuth = $user.LastAuthentication
            DaysSinceLastAuth = $user.DaysSinceLastAuthentication
            IsEligible = $user.EligibleForLicenseRemoval
        }
    }
    
    # Complete the progress bar
    Write-Progress -Activity "Creating comparison report" -Completed
    
    # Export results
    $comparisonResults | Export-Csv -Path $beforeAfterReportPath -NoTypeInformation
    Write-Host "Before/After comparison report saved to: $beforeAfterReportPath" -ForegroundColor Cyan
}

# Run the interactive menu
$choice = ""
do {
    $choice = Show-MainMenu
    
    switch ($choice) {
        "1" {
            # Generate the report of eligible users
            if (Generate-InactivityReport) {
                Write-Host "Report generated successfully and saved to: $global:reportPath" -ForegroundColor Green
            } else {
                Write-Host "Failed to generate report." -ForegroundColor Red
            }
            Read-Host "Press Enter to continue"
        }
        "2" {
            # Check if report has been generated
            if ($null -eq $global:results) {
                Write-Host "Please generate the report first (Option 1)." -ForegroundColor Yellow
                Read-Host "Press Enter to continue"
                continue
            }
            
            Convert-ToSharedMailbox
            Read-Host "Press Enter to continue"
        }
        "3" {
            # Check if report has been generated
            if ($null -eq $global:results) {
                Write-Host "Please generate the report first (Option 1)." -ForegroundColor Yellow
                Read-Host "Press Enter to continue"
                continue
            }
            
            Set-ZeroReceiveLimit
            Read-Host "Press Enter to continue"
        }
        "4" {
            # Check if report has been generated
            if ($null -eq $global:results) {
                Write-Host "Please generate the report first (Option 1)." -ForegroundColor Yellow
                Read-Host "Press Enter to continue"
                continue
            }
            
            Block-SharedMailboxSignIn
            Read-Host "Press Enter to continue"
        }
        "5" {
            # Check if report has been generated
            if ($null -eq $global:results) {
                Write-Host "Please generate the report first (Option 1)." -ForegroundColor Yellow
                Read-Host "Press Enter to continue"
                continue
            }
            
            Remove-UserLicenses
            Read-Host "Press Enter to continue"
        }
        "6" {
            # Check if report has been generated
            if ($null -eq $global:results) {
                Write-Host "Please generate the report first (Option 1)." -ForegroundColor Yellow
                Read-Host "Press Enter to continue"
                continue
            }
            
            Remove-UserRoles
            Read-Host "Press Enter to continue"
        }
        "7" {
            # Check if report has been generated
            if ($null -eq $global:results) {
                Write-Host "Please generate the report first (Option 1)." -ForegroundColor Yellow
                Read-Host "Press Enter to continue"
                continue
            }
            
            Disable-OnPremADAccounts
            Read-Host "Press Enter to continue"
        }
        "8" {
            # Check if report has been generated
            if ($null -eq $global:results) {
                Write-Host "Please generate the report first (Option 1)." -ForegroundColor Yellow
                Read-Host "Press Enter to continue"
                continue
            }
            
            Create-ComparisonReport
            Read-Host "Press Enter to continue"
        }
        "Q" { Write-Host "Exiting..." -ForegroundColor Yellow }
        "q" { Write-Host "Exiting..." -ForegroundColor Yellow }
        default { 
            Write-Host "Invalid option, please try again." -ForegroundColor Red
            Read-Host "Press Enter to continue"
        }
    }
} while ($choice -ne "Q" -and $choice -ne "q")

# Disconnect from services if still connected
try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Write-Host "Disconnected from Microsoft Graph" -ForegroundColor Cyan
}
catch {
    # Already disconnected or error, ignore
}

try {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Disconnected from Exchange Online" -ForegroundColor Cyan
}
catch {
    # Already disconnected or error, ignore
}

Write-Host "Script execution completed." -ForegroundColor Green
