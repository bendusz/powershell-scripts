<#
.SYNOPSIS
    Adds a selected user as an owner or member to selected Microsoft Teams groups.
.DESCRIPTION
    This script checks for Microsoft Graph module, authenticates, verifies permissions,
    and allows selecting a user to add as an owner or member to all or selected Teams groups.
.NOTES
    File Name      : add-user-as-owner-to-all-teams.ps1
    Author         : Ben Vegh
    Prerequisite   : PowerShell 5.1 or higher, Microsoft Graph PowerShell SDK
.EXAMPLE
    .\add-user-as-owner-to-all-teams.ps1
#>

# Requires -Version 5.1

# Function to write logs
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Determine log file path based on OS
    if ($IsWindows -or $env:OS -like "*Windows*") {
        $logPath = Join-Path -Path $env:USERPROFILE -ChildPath "ScriptLogs\TeamsMembershipLogs"
    }
    else {
        $logPath = Join-Path -Path $HOME -ChildPath "ScriptLogs/TeamsMembershipLogs"
    }
    
    if (-not (Test-Path -Path $logPath)) {
        New-Item -Path $logPath -ItemType Directory -Force | Out-Null
    }
    
    $logFile = Join-Path -Path $logPath -ChildPath "TeamsMembership_$(Get-Date -Format 'yyyyMMdd').log"
    Add-Content -Path $logFile -Value $logMessage
    
    # Also output to console
    switch ($Level) {
        'Info' { Write-Host $logMessage -ForegroundColor Cyan }
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        'Error' { Write-Host $logMessage -ForegroundColor Red }
    }
}

# Function to check and install modules
function Ensure-Module {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModuleName
    )
    
    Write-Log -Message "Checking if $ModuleName module is installed..." -Level Info
    
    $module = Get-Module -Name $ModuleName -ListAvailable
    
    if (-not $module) {
        Write-Log -Message "$ModuleName module not found. Attempting to install..." -Level Warning
        
        # Check if PSGallery is trusted
        $PSGalleryTrusted = (Get-PSRepository -Name "PSGallery").InstallationPolicy -eq "Trusted"
        
        if (-not $PSGalleryTrusted) {
            Write-Log -Message "Setting PSGallery as trusted repository..." -Level Info
            Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
        }
        
        try {
            Install-Module -Name $ModuleName -Repository PSGallery -Force -AllowClobber -Scope CurrentUser
            Write-Log -Message "$ModuleName module installed successfully." -Level Info
        }
        catch {
            Write-Log -Message "Failed to install $ModuleName module: $_" -Level Error
            return $false
        }
    }
    else {
        Write-Log -Message "$ModuleName module is already installed." -Level Info
    }
    
    return $true
}

# Function to show a cross-platform selection menu
function Show-SelectionMenu {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Items,
        
        [Parameter(Mandatory = $true)]
        [string]$Title,
        
        [Parameter(Mandatory = $false)]
        [string]$DisplayProperty = $null,
        
        [Parameter(Mandatory = $false)]
        [switch]$AllowMultiple = $false,
        
        [Parameter(Mandatory = $false)]
        [int]$PageSize = 20
    )
    
    if ($Items.Count -eq 0) {
        Write-Log -Message "No items to display in menu." -Level Warning
        return $null
    }
    
    Write-Log -Message "Show-SelectionMenu called with $($Items.Count) items, DisplayProperty: '$DisplayProperty'" -Level Info
    
    # Check if Out-GridView is available (Windows)
    if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
        Write-Log -Message "Using Out-GridView for selection..." -Level Info
        if ($AllowMultiple) {
            return $Items | Out-GridView -Title $Title -OutputMode Multiple
        } else {
            return $Items | Out-GridView -Title $Title -OutputMode Single
        }
    }
    
    # Fallback to console menu for macOS/Linux
    Write-Log -Message "Using console menu for selection..." -Level Info
    
    $selectedItems = @()
    $currentPage = 0
    $totalPages = [math]::Ceiling($Items.Count / $PageSize)
    
    do {
        Clear-Host
        Write-Host "`n=== $Title ===" -ForegroundColor Green
        Write-Host "Page $($currentPage + 1) of $totalPages" -ForegroundColor Yellow
        
        $startIndex = $currentPage * $PageSize
        $endIndex = [math]::Min($startIndex + $PageSize - 1, $Items.Count - 1)
        $pageItems = $Items[$startIndex..$endIndex]
        
        # Display items
        for ($i = 0; $i -lt $pageItems.Count; $i++) {
            $item = $pageItems[$i]
            $displayText = if ($DisplayProperty -and $item.$DisplayProperty) { 
                $item.$DisplayProperty 
            } elseif ($DisplayProperty) {
                "<No $DisplayProperty>"
            } else { 
                $item.ToString() 
            }
            $itemNumber = $startIndex + $i + 1
            
            # Show selection indicator if already selected
            $indicator = if ($selectedItems -contains $item) { "[*]" } else { "[ ]" }
            Write-Host "$indicator $itemNumber. $displayText" -ForegroundColor Cyan
        }
        
        Write-Host "`nOptions:" -ForegroundColor Yellow
        if ($AllowMultiple) {
            Write-Host "Enter number(s) to toggle selection (e.g., 1,3,5 or 1-5): " -ForegroundColor White -NoNewline
        } else {
            Write-Host "Enter number to select: " -ForegroundColor White -NoNewline
        }
        
        if ($totalPages -gt 1) {
            Write-Host "`n[N]ext page, [P]revious page, " -ForegroundColor Gray -NoNewline
        }
        Write-Host "[Q]uit" -ForegroundColor Gray
        
        if ($AllowMultiple -and $selectedItems.Count -gt 0) {
            Write-Host "[F]inish selection ($($selectedItems.Count) selected)" -ForegroundColor Green
        }
        
        $userInput = Read-Host "`nChoice"
        
        switch ($userInput.ToUpper()) {
            'Q' { 
                Write-Log -Message "Selection cancelled by user." -Level Info
                return $null 
            }
            'N' { 
                if ($currentPage -lt $totalPages - 1) { $currentPage++ }
            }
            'P' { 
                if ($currentPage -gt 0) { $currentPage-- }
            }
            'F' {
                if ($AllowMultiple -and $selectedItems.Count -gt 0) {
                    Write-Log -Message "User finished selection with $($selectedItems.Count) items." -Level Info
                    return $selectedItems
                }
            }
            default {
                # Parse number input
                try {
                    if ($userInput -match '^\d+$') {
                        # Single number
                        $number = [int]$userInput
                        if ($number -ge 1 -and $number -le $Items.Count) {
                            $selectedItem = $Items[$number - 1]
                            if ($AllowMultiple) {
                                if ($selectedItems -contains $selectedItem) {
                                    $selectedItems = $selectedItems | Where-Object { $_ -ne $selectedItem }
                                    Write-Log -Message "Deselected item $number" -Level Info
                                } else {
                                    $selectedItems += $selectedItem
                                    Write-Log -Message "Selected item $number" -Level Info
                                }
                            } else {
                                Write-Log -Message "User selected item $number" -Level Info
                                return $selectedItem
                            }
                        } else {
                            Write-Host "Invalid selection. Please enter a number between 1 and $($Items.Count)." -ForegroundColor Red
                            Start-Sleep -Seconds 2
                        }
                    }
                    elseif ($userInput -match '^\d+(-\d+)?(,\d+(-\d+)?)*$' -and $AllowMultiple) {
                        # Range or comma-separated numbers
                        $numbers = @()
                        $parts = $userInput -split ','
                        foreach ($part in $parts) {
                            if ($part -match '^(\d+)-(\d+)$') {
                                $start = [int]$matches[1]
                                $end = [int]$matches[2]
                                $numbers += $start..$end
                            } else {
                                $numbers += [int]$part
                            }
                        }
                        
                        foreach ($number in $numbers) {
                            if ($number -ge 1 -and $number -le $Items.Count) {
                                $selectedItem = $Items[$number - 1]
                                if ($selectedItems -contains $selectedItem) {
                                    $selectedItems = $selectedItems | Where-Object { $_ -ne $selectedItem }
                                } else {
                                    $selectedItems += $selectedItem
                                }
                            }
                        }
                        Write-Log -Message "Processed selection: $userInput" -Level Info
                    }
                    else {
                        Write-Host "Invalid input format." -ForegroundColor Red
                        Start-Sleep -Seconds 2
                    }
                }
                catch {
                    Write-Host "Invalid input. Please try again." -ForegroundColor Red
                    Start-Sleep -Seconds 2
                }
            }
        }
    } while ($true)
}

# Function to check required permissions
function Check-RequiredPermissions {
    [CmdletBinding()]
    param()
    
    Write-Log -Message "Checking required permissions..." -Level Info
    
    $requiredPermissions = @(
        "User.Read.All",
        "Group.ReadWrite.All"
    )
    
    $scopes = (Get-MgContext).Scopes
    
    if (-not $scopes) {
        Write-Log -Message "No permissions found in current context." -Level Error
        return $false
    }
    
    $missingPermissions = $requiredPermissions | Where-Object { $_ -notin $scopes }
    
    if ($missingPermissions) {
        Write-Log -Message "Missing required permissions: $($missingPermissions -join ', ')" -Level Error
        return $false
    }
    
    Write-Log -Message "All required permissions are present." -Level Info
    return $true
}

# Function to connect to Microsoft Graph
function Connect-ToMicrosoftGraph {
    [CmdletBinding()]
    param()
    
    Write-Log -Message "Connecting to Microsoft Graph..." -Level Info
    
    try {
        # Check if already connected
        $context = Get-MgContext
        if ($context) {
            Write-Log -Message "Already connected to Microsoft Graph as $($context.Account)" -Level Info
            return $true
        }
        
        # Connect with required permissions
        Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All"
        Write-Log -Message "Successfully connected to Microsoft Graph as $($(Get-MgContext).Account)" -Level Info
        return $true
    }
    catch {
        Write-Log -Message "Failed to connect to Microsoft Graph: $_" -Level Error
        return $false
    }
}

# Function to get active users
function Get-ActiveUsers {
    [CmdletBinding()]
    param()
    
    Write-Log -Message "Retrieving active users..." -Level Info
    
    try {
        $users = Get-MgUser -Filter "accountEnabled eq true" -All | 
                 Select-Object DisplayName, UserPrincipalName, Id |
                 Sort-Object DisplayName
        
        Write-Log -Message "Retrieved $($users.Count) active users." -Level Info
        return $users
    }
    catch {
        Write-Log -Message "Failed to retrieve active users: $_" -Level Error
        return $null
    }
}

# Function to get Teams groups
function Get-TeamsGroups {
    [CmdletBinding()]
    param()
    
    Write-Log -Message "Retrieving Teams groups..." -Level Info
    
    try {
        $teamsGroups = Get-MgGroup -Filter "resourceProvisioningOptions/any(x:x eq 'Team')" -All | 
                       Select-Object DisplayName, Description, Id |
                       Sort-Object DisplayName
        
        Write-Log -Message "Retrieved $($teamsGroups.Count) Teams groups." -Level Info
        return $teamsGroups
    }
    catch {
        Write-Log -Message "Failed to retrieve Teams groups: $_" -Level Error
        return $null
    }
}

# Function to add user to Teams group
function Add-UserToTeamsGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupId,
        
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Owner', 'Member')]
        [string]$Role
    )
    
    try {
        # Get group info for logging
        $group = Get-MgGroup -GroupId $GroupId
        $user = Get-MgUser -UserId $UserId
        
        if ($Role -eq 'Owner') {
            #Check if user is already an owner
            $existingOwners = Get-MgGroupOwner -GroupId $GroupId -All | Where-Object { $_.Id -eq $UserId }
            if ($existingOwners) {
                Write-Log -Message "User $($user.DisplayName) is already an Owner of team: $($group.DisplayName)" -Level Warning
                return $true
            }
            # Add user as owner
            Write-Log -Message "Adding $($user.DisplayName) as Owner to team: $($group.DisplayName)" -Level Info
            New-MgGroupOwner -GroupId $GroupId -DirectoryObjectId $UserId
        }
        else {
            #Check if user is already a member
            $existingMembers = Get-MgGroupMember -GroupId $GroupId -All | Where-Object { $_.Id -eq $UserId }
            if ($existingMembers) {
                Write-Log -Message "User $($user.DisplayName) is already a Member of team: $($group.DisplayName)" -Level Warning
                return $true
            }
            # Add user as member
            Write-Log -Message "Adding $($user.DisplayName) as Member to team: $($group.DisplayName)" -Level Info
            New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $UserId
        }
        
        Write-Log -Message "Successfully added $($user.DisplayName) as $Role to $($group.DisplayName)" -Level Info
        return $true
    }
    catch {
        Write-Log -Message "Failed to add user $($UserId) as $Role to group $($GroupId): $_" -Level Error
        return $false
    }
}

# Main execution starts here
Write-Log -Message "Starting script execution..." -Level Info

# Check and install required modules
if (-not (Ensure-Module -ModuleName "Microsoft.Graph")) {
    Write-Log -Message "Required module Microsoft.Graph could not be installed. Exiting script." -Level Error
    exit 1
}

# Import the Microsoft Graph module
Import-Module Microsoft.Graph

# Connect to Microsoft Graph
if (-not (Connect-ToMicrosoftGraph)) {
    Write-Log -Message "Failed to connect to Microsoft Graph. Exiting script." -Level Error
    exit 1
}

# Check required permissions
if (-not (Check-RequiredPermissions)) {
    Write-Log -Message "Missing required permissions. Please reconnect with appropriate permissions. Exiting script." -Level Error
    Disconnect-MgGraph
    exit 1
}

# Get active users
$activeUsers = Get-ActiveUsers
if (-not $activeUsers -or $activeUsers.Count -eq 0) {
    Write-Log -Message "No active users found. Exiting script." -Level Error
    Disconnect-MgGraph
    exit 1
}

Write-Log -Message "Found $($activeUsers.Count) active users. First few users: $($activeUsers[0..2].DisplayName -join ', ')" -Level Info

# Let user select a user
Write-Log -Message "Prompting to select a user..." -Level Info
$selectedUser = Show-SelectionMenu -Items $activeUsers -Title "Select a user to add to Teams groups" -DisplayProperty "DisplayName"

if (-not $selectedUser) {
    Write-Log -Message "No user was selected. Exiting script." -Level Warning
    Disconnect-MgGraph
    exit 0
}

Write-Log -Message "Selected user: $($selectedUser.DisplayName) ($($selectedUser.UserPrincipalName))" -Level Info

# Let user select the role
Write-Log -Message "Prompting to select role (Owner or Member)..." -Level Info
$roleOptions = @(
    [PSCustomObject]@{Role = 'Owner'; Description = 'User will be added as an Owner to selected Teams groups'}
    [PSCustomObject]@{Role = 'Member'; Description = 'User will be added as a Member to selected Teams groups'}
)

$selectedRole = Show-SelectionMenu -Items $roleOptions -Title "Select role for the user" -DisplayProperty "Role"

if (-not $selectedRole) {
    Write-Log -Message "No role was selected. Exiting script." -Level Warning
    Disconnect-MgGraph
    exit 0
}

Write-Log -Message "Selected role: $($selectedRole.Role)" -Level Info

# Let user select scope (all Teams or specific Teams)
Write-Log -Message "Prompting to select scope (all Teams or specific Teams)..." -Level Info
$scopeOptions = @(
    [PSCustomObject]@{Scope = 'All Teams'; Description = 'Add user to all Teams groups'}
    [PSCustomObject]@{Scope = 'Specific Teams'; Description = 'Select specific Teams groups to add the user to'}
)

$selectedScope = Show-SelectionMenu -Items $scopeOptions -Title "Select scope for adding user" -DisplayProperty "Scope"

if (-not $selectedScope) {
    Write-Log -Message "No scope was selected. Exiting script." -Level Warning
    Disconnect-MgGraph
    exit 0
}

Write-Log -Message "Selected scope: $($selectedScope.Scope)" -Level Info

# Get Teams groups
$teamsGroups = Get-TeamsGroups
if (-not $teamsGroups -or $teamsGroups.Count -eq 0) {
    Write-Log -Message "No Teams groups found. Exiting script." -Level Error
    Disconnect-MgGraph
    exit 1
}

# Define which Teams to process based on selection
$teamsToProcess = @()

if ($selectedScope.Scope -eq 'All Teams') {
    Write-Log -Message "Will add user to all $($teamsGroups.Count) Teams groups." -Level Info
    $teamsToProcess = $teamsGroups
}
else {
    Write-Log -Message "Prompting to select specific Teams groups..." -Level Info
    $selectedTeams = Show-SelectionMenu -Items $teamsGroups -Title "Select Teams groups to add the user to" -DisplayProperty "DisplayName" -AllowMultiple
    
    if (-not $selectedTeams -or $selectedTeams.Count -eq 0) {
        Write-Log -Message "No Teams groups were selected. Exiting script." -Level Warning
        Disconnect-MgGraph
        exit 0
    }
    
    Write-Log -Message "Selected $($selectedTeams.Count) Teams groups." -Level Info
    $teamsToProcess = $selectedTeams
}

# Process Teams groups
$successCount = 0
$failCount = 0
$totalCount = $teamsToProcess.Count

Write-Log -Message "Starting to process $totalCount Teams groups..." -Level Info

foreach ($team in $teamsToProcess) {
    $result = Add-UserToTeamsGroup -GroupId $team.Id -UserId $selectedUser.Id -Role $selectedRole.Role
    
    if ($result) {
        $successCount++
    }
    else {
        $failCount++
    }
    
    # Show progress
    $percentComplete = [math]::Round(($successCount + $failCount) / $totalCount * 100, 0)
    Write-Progress -Activity "Adding $($selectedUser.DisplayName) as $($selectedRole.Role)" -Status "$percentComplete% Complete" -PercentComplete $percentComplete
}

Write-Progress -Activity "Adding $($selectedUser.DisplayName) as $($selectedRole.Role)" -Completed

# Summary
Write-Log -Message "Operation completed. Summary:" -Level Info
Write-Log -Message "- Total Teams groups processed: $totalCount" -Level Info
Write-Log -Message "- Successful operations: $successCount" -Level Info
Write-Log -Message "- Failed operations: $failCount" -Level Info

# Disconnect from Microsoft Graph
Write-Log -Message "Disconnecting from Microsoft Graph..." -Level Info
Disconnect-MgGraph

Write-Log -Message "Script execution completed." -Level Info
exit 0

