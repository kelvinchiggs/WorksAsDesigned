<#
.SYNOPSIS
    Comprehensive Microsoft Teams Security and Compliance Audit Script
    
.DESCRIPTION
    This script performs an enterprise-level audit of Microsoft Teams deployment covering:
    - Governance, Strategy & Controls
    - Identity & Access Management
    - Teams & Channels Architecture
    - Security & Compliance
    - Messaging, Meetings & Collaboration
    - Voice & Telephony (Teams Phone)
    - Network & Media Quality
    - Devices & Rooms
    - Apps, Integrations & Power Platform
    - SharePoint & OneDrive Integration
    
    The script produces comprehensive Excel reports with:
    - RAG (Red/Amber/Green) rated findings
    - Current vs Recommended vs Best Practice configurations
    - Risk impact and likelihood scoring
    - Remediation roadmap (30/90/180+ day timeline)
    
.PARAMETER TenantId
    The Microsoft 365 Tenant ID (optional - will auto-detect if not provided)
    
.PARAMETER IncludeVoice
    Include Teams Phone/Voice audit sections (may require additional permissions)
    
.PARAMETER IncludeDevices
    Include device and Teams Rooms audit sections
    
.PARAMETER OutputPath
    Custom output path for reports and logs. Defaults to My Documents\Invoke-TeamsComprehensiveAudit
    
.EXAMPLE
    .\Invoke-TeamsComprehensiveAudit.ps1
    Runs complete audit with default settings
    
.EXAMPLE
    .\Invoke-TeamsComprehensiveAudit.ps1 -IncludeVoice -IncludeDevices
    Runs complete audit including Voice and Device assessments
    
.NOTES
    Script Name : Invoke-TeamsComprehensiveAudit.ps1
    Author      : Kelvin Chigorimbo
    Version     : 1.0
    Requires    : PowerShell 7.0 or later
    Modules     : Microsoft.Graph, MicrosoftTeams, ExchangeOnlineManagement, ImportExcel
    Permissions : Reports.Read.All, Directory.Read.All, Policy.Read.All, Organization.Read.All
                  TeamSettings.Read.All, User.Read.All, SecurityEvents.Read.All
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeVoice,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeDevices,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

#Requires -Version 7.0

# Script configuration
$ErrorActionPreference = 'Continue'
$WarningPreference = 'Continue'
$VerbosePreference = 'Continue'

# Required modules configuration
$RequiredModules = @(
    @{ Name = 'Microsoft.Graph.Authentication'; MinVersion = '2.0.0' }
    @{ Name = 'Microsoft.Graph.Reports'; MinVersion = '2.0.0' }
    @{ Name = 'Microsoft.Graph.Identity.DirectoryManagement'; MinVersion = '2.0.0' }
    @{ Name = 'Microsoft.Graph.Identity.SignIns'; MinVersion = '2.0.0' }
    @{ Name = 'Microsoft.Graph.Users'; MinVersion = '2.0.0' }
    @{ Name = 'Microsoft.Graph.Groups'; MinVersion = '2.0.0' }
    @{ Name = 'MicrosoftTeams'; MinVersion = '5.0.0' }
    @{ Name = 'ExchangeOnlineManagement'; MinVersion = '3.0.0' }
    @{ Name = 'ImportExcel'; MinVersion = '7.8.0' }
)

# Required Graph permissions
$RequiredScopes = @(
    'Reports.Read.All'
    'Directory.Read.All'
    'Policy.Read.All'
    'Organization.Read.All'
    'TeamSettings.Read.All'
    'User.Read.All'
    'SecurityEvents.Read.All'
    'AuditLog.Read.All'
)

# Global variables for script execution
$Script:LogFile = $null
$Script:StartTime = Get-Date
$Script:AuditResults = @{}
$Script:RAGFindings = @()
$Script:RiskScores = @()
$Script:RemediationRoadmap = @()

#region Helper Functions

function Write-Log {
    <#
    .SYNOPSIS
        Writes messages to log file and console
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS', 'DEBUG')]
        [string]$Level = 'INFO'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to log file
    if ($Script:LogFile) {
        Add-Content -Path $Script:LogFile -Value $logMessage -ErrorAction SilentlyContinue
    }
    
    # Write to console with color coding
    switch ($Level) {
        'ERROR'   { Write-Host $logMessage -ForegroundColor Red }
        'WARNING' { Write-Host $logMessage -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $logMessage -ForegroundColor Green }
        'DEBUG'   { Write-Verbose $logMessage }
        default   { Write-Host $logMessage -ForegroundColor Cyan }
    }
}

function Test-RequiredModules {
    <#
    .SYNOPSIS
        Validates that all required PowerShell modules are installed
    #>
    [CmdletBinding()]
    param()
    
    Write-Log "Checking required PowerShell modules..." -Level INFO
    $missingModules = @()
    
    foreach ($module in $RequiredModules) {
        Write-Progress -Activity "Validating Modules" -Status "Checking $($module.Name)..." -PercentComplete 0
        
        $installedModule = Get-Module -ListAvailable -Name $module.Name | 
            Where-Object { $_.Version -ge $module.MinVersion } | 
            Select-Object -First 1
        
        if (-not $installedModule) {
            Write-Log "Module $($module.Name) version $($module.MinVersion) or later not found" -Level WARNING
            $missingModules += $module.Name
        } else {
            Write-Log "Module $($module.Name) version $($installedModule.Version) found" -Level SUCCESS
        }
    }
    
    Write-Progress -Activity "Validating Modules" -Completed
    
    if ($missingModules.Count -gt 0) {
        Write-Log "Missing required modules: $($missingModules -join ', ')" -Level ERROR
        Write-Log "Install missing modules using: Install-Module -Name <ModuleName> -Scope CurrentUser" -Level WARNING
        return $false
    }
    
    Write-Log "All required modules are installed" -Level SUCCESS
    return $true
}

function Initialize-OutputDirectory {
    <#
    .SYNOPSIS
        Creates output directory structure for reports and logs
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$CustomPath
    )
    
    if ([string]::IsNullOrEmpty($CustomPath)) {
        $documentsPath = [Environment]::GetFolderPath('MyDocuments')
        $outputDir = Join-Path -Path $documentsPath -ChildPath 'Invoke-TeamsComprehensiveAudit'
    } else {
        $outputDir = $CustomPath
    }
    
    # Create directory structure
    $directories = @(
        $outputDir
        (Join-Path -Path $outputDir -ChildPath 'Reports')
        (Join-Path -Path $outputDir -ChildPath 'Logs')
        (Join-Path -Path $outputDir -ChildPath 'Evidence')
    )
    
    foreach ($dir in $directories) {
        if (-not (Test-Path -Path $dir)) {
            New-Item -Path $dir -ItemType Directory -Force | Out-Null
            Write-Log "Created directory: $dir" -Level SUCCESS
        }
    }
    
    # Initialize log file
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $Script:LogFile = Join-Path -Path $outputDir -ChildPath "Logs\TeamsAudit_$timestamp.log"
    
    Write-Log "Output directory initialized: $outputDir" -Level SUCCESS
    Write-Log "Log file: $Script:LogFile" -Level INFO
    
    return $outputDir
}

function Connect-RequiredServices {
    <#
    .SYNOPSIS
        Connects to Microsoft Graph, Teams, and Exchange Online
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$TenantId
    )
    
    Write-Log "Connecting to Microsoft services..." -Level INFO
    
    try {
        # Connect to Microsoft Graph
        Write-Log "Connecting to Microsoft Graph..." -Level INFO
        $graphParams = @{
            Scopes = $RequiredScopes
            NoWelcome = $true
        }
        
        if (-not [string]::IsNullOrEmpty($TenantId)) {
            $graphParams['TenantId'] = $TenantId
        }
        
        Connect-MgGraph @graphParams -ErrorAction Stop
        Write-Log "Connected to Microsoft Graph successfully" -Level SUCCESS
        
        # Get tenant information
        $tenantInfo = Get-MgOrganization -ErrorAction Stop
        Write-Log "Tenant: $($tenantInfo.DisplayName) (ID: $($tenantInfo.Id))" -Level INFO
        $Script:AuditResults['TenantInfo'] = $tenantInfo
        
        # Connect to Microsoft Teams
        Write-Log "Connecting to Microsoft Teams..." -Level INFO
        Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
        Write-Log "Connected to Microsoft Teams successfully" -Level SUCCESS
        
        # Connect to Exchange Online
        Write-Log "Connecting to Exchange Online..." -Level INFO
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Log "Connected to Exchange Online successfully" -Level SUCCESS
        
        return $true
        
    } catch {
        Write-Log "Failed to connect to required services: $_" -Level ERROR
        return $false
    }
}

function Get-RAGRating {
    <#
    .SYNOPSIS
        Calculates RAG (Red/Amber/Green) rating based on compliance percentage
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$CompliancePercentage
    )
    
    if ($CompliancePercentage -ge 80) {
        return 'Green'
    } elseif ($CompliancePercentage -ge 50) {
        return 'Amber'
    } else {
        return 'Red'
    }
}

function Get-RiskScore {
    <#
    .SYNOPSIS
        Calculates risk score based on impact and likelihood
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Low', 'Medium', 'High', 'Critical')]
        [string]$Impact,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Low', 'Medium', 'High')]
        [string]$Likelihood
    )
    
    $impactScore = switch ($Impact) {
        'Low' { 1 }
        'Medium' { 2 }
        'High' { 3 }
        'Critical' { 4 }
    }
    
    $likelihoodScore = switch ($Likelihood) {
        'Low' { 1 }
        'Medium' { 2 }
        'High' { 3 }
    }
    
    $riskScore = $impactScore * $likelihoodScore
    
    $riskLevel = switch ($riskScore) {
        { $_ -le 2 } { 'Low' }
        { $_ -le 6 } { 'Medium' }
        { $_ -le 9 } { 'High' }
        default { 'Critical' }
    }
    
    return @{
        Score = $riskScore
        Level = $riskLevel
        Impact = $Impact
        Likelihood = $Likelihood
    }
}

function Add-Finding {
    <#
    .SYNOPSIS
        Adds an audit finding to the global findings collection
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Category,
        
        [Parameter(Mandatory = $true)]
        [string]$SubCategory,
        
        [Parameter(Mandatory = $true)]
        [string]$Finding,
        
        [Parameter(Mandatory = $true)]
        [string]$CurrentState,
        
        [Parameter(Mandatory = $true)]
        [string]$RecommendedState,
        
        [Parameter(Mandatory = $true)]
        [string]$BestPractice,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Red', 'Amber', 'Green')]
        [string]$RAGStatus,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Low', 'Medium', 'High', 'Critical')]
        [string]$Impact,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Low', 'Medium', 'High')]
        [string]$Likelihood,
        
        [Parameter(Mandatory = $true)]
        [string]$Remediation,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Quick Win (30 days)', 'Medium Term (90 days)', 'Strategic (6-12 months)')]
        [string]$Timeline
    )
    
    $riskScore = Get-RiskScore -Impact $Impact -Likelihood $Likelihood
    
    $finding = [PSCustomObject]@{
        Category = $Category
        SubCategory = $SubCategory
        Finding = $Finding
        CurrentState = $CurrentState
        RecommendedState = $RecommendedState
        BestPractice = $BestPractice
        RAGStatus = $RAGStatus
        Impact = $Impact
        Likelihood = $Likelihood
        RiskScore = $riskScore.Score
        RiskLevel = $riskScore.Level
        Remediation = $Remediation
        Timeline = $Timeline
        DateIdentified = Get-Date -Format 'yyyy-MM-dd'
    }
    
    $Script:RAGFindings += $finding
}

#endregion

#region Audit Functions - 1. Governance & Controls

function Get-GovernanceAudit {
    <#
    .SYNOPSIS
        Audits Teams governance, strategy, and control baseline
    #>
    [CmdletBinding()]
    param()
    
    Write-Log "Starting Governance & Controls audit..." -Level INFO
    $results = @{}
    
    try {
        # 1.1 Tenant Configuration Baseline
        Write-Progress -Activity "Governance Audit" -Status "Checking tenant configuration..." -PercentComplete 10
        
        # Get Teams tenant configuration
        $teamsConfig = Get-CsTeamsClientConfiguration -ErrorAction SilentlyContinue
        $tenantFederationConfig = Get-CsTenantFederationConfiguration -ErrorAction SilentlyContinue
        
        # Check coexistence mode
        $coexistenceMode = (Get-CsTeamsUpgradeConfiguration -ErrorAction SilentlyContinue).Mode
        $results['CoexistenceMode'] = $coexistenceMode
        
        if ($coexistenceMode -eq 'Islands') {
            Add-Finding -Category 'Governance' -SubCategory 'Tenant Configuration' `
                -Finding 'Legacy Islands mode detected' `
                -CurrentState "Islands mode enabled" `
                -RecommendedState "Teams Only mode" `
                -BestPractice "Microsoft recommends Teams Only mode for modern deployments" `
                -RAGStatus 'Amber' `
                -Impact 'Medium' -Likelihood 'High' `
                -Remediation 'Plan migration to Teams Only mode with user communication strategy' `
                -Timeline 'Medium Term (90 days)'
        }
        
        # 1.2 Admin Role Separation
        Write-Progress -Activity "Governance Audit" -Status "Analyzing admin roles..." -PercentComplete 30
        
        $adminRoles = Get-MgDirectoryRole -All -ErrorAction Stop
        $teamsAdminRole = $adminRoles | Where-Object { $_.DisplayName -eq 'Teams Administrator' }
        $globalAdminRole = $adminRoles | Where-Object { $_.DisplayName -eq 'Global Administrator' }
        
        if ($teamsAdminRole) {
            $teamsAdmins = Get-MgDirectoryRoleMember -DirectoryRoleId $teamsAdminRole.Id -All -ErrorAction Stop
            $results['TeamsAdminCount'] = $teamsAdmins.Count
            
            if ($teamsAdmins.Count -eq 0) {
                Add-Finding -Category 'Governance' -SubCategory 'Admin Roles' `
                    -Finding 'No dedicated Teams Administrators assigned' `
                    -CurrentState "0 Teams Administrators" `
                    -RecommendedState "At least 2-3 dedicated Teams Administrators" `
                    -BestPractice "Separate Teams administration from Global Admin for least privilege" `
                    -RAGStatus 'Red' `
                    -Impact 'High' -Likelihood 'High' `
                    -Remediation 'Assign dedicated Teams Administrator roles to appropriate personnel' `
                    -Timeline 'Quick Win (30 days)'
            }
        }
        
        # Check Global Admin count
        if ($globalAdminRole) {
            $globalAdmins = Get-MgDirectoryRoleMember -DirectoryRoleId $globalAdminRole.Id -All -ErrorAction Stop
            $results['GlobalAdminCount'] = $globalAdmins.Count
            
            if ($globalAdmins.Count -gt 5) {
                Add-Finding -Category 'Governance' -SubCategory 'Admin Roles' `
                    -Finding 'Excessive Global Administrators detected' `
                    -CurrentState "$($globalAdmins.Count) Global Administrators" `
                    -RecommendedState "Maximum 3-5 Global Administrators" `
                    -BestPractice "Minimize Global Admin assignments, use specific admin roles instead" `
                    -RAGStatus 'Amber' `
                    -Impact 'High' -Likelihood 'Medium' `
                    -Remediation 'Review and reduce Global Admin assignments, implement PIM' `
                    -Timeline 'Quick Win (30 days)'
            }
        }
        
        # 1.3 Privileged Access Model
        Write-Progress -Activity "Governance Audit" -Status "Checking privileged access controls..." -PercentComplete 50
        
        # Check for Conditional Access policies
        $caPolici = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction SilentlyContinue
        $results['ConditionalAccessPolicyCount'] = ($caPolicies | Measure-Object).Count
        
        if (($caPolicies | Measure-Object).Count -eq 0) {
            Add-Finding -Category 'Governance' -SubCategory 'Privileged Access' `
                -Finding 'No Conditional Access policies configured' `
                -CurrentState "0 Conditional Access policies" `
                -RecommendedState "Minimum 5-10 policies covering admin access, MFA, device compliance" `
                -BestPractice "Implement Conditional Access for zero-trust security model" `
                -RAGStatus 'Red' `
                -Impact 'Critical' -Likelihood 'High' `
                -Remediation 'Design and implement Conditional Access policy framework' `
                -Timeline 'Medium Term (90 days)'
        }
        
        # 1.4 Compliance Alignment - Microsoft Secure Score
        Write-Progress -Activity "Governance Audit" -Status "Retrieving Secure Score..." -PercentComplete 70
        
        try {
            $secureScore = Get-MgSecuritySecureScore -Top 1 -ErrorAction SilentlyContinue
            if ($secureScore) {
                $currentScore = $secureScore.CurrentScore
                $maxScore = $secureScore.MaxScore
                $scorePercentage = [math]::Round(($currentScore / $maxScore) * 100, 2)
                
                $results['SecureScore'] = @{
                    Current = $currentScore
                    Max = $maxScore
                    Percentage = $scorePercentage
                }
                
                $ragStatus = Get-RAGRating -CompliancePercentage $scorePercentage
                
                Add-Finding -Category 'Governance' -SubCategory 'Compliance' `
                    -Finding 'Microsoft Secure Score assessment' `
                    -CurrentState "$currentScore out of $maxScore ($scorePercentage%)" `
                    -RecommendedState "Target 80% or higher" `
                    -BestPractice "Maintain Secure Score above 80% for optimal security posture" `
                    -RAGStatus $ragStatus `
                    -Impact 'High' -Likelihood 'Medium' `
                    -Remediation 'Review and implement Secure Score recommendations systematically' `
                    -Timeline $(if ($scorePercentage -lt 50) { 'Medium Term (90 days)' } else { 'Quick Win (30 days)' })
            }
        } catch {
            Write-Log "Unable to retrieve Secure Score: $_" -Level WARNING
        }
        
        Write-Progress -Activity "Governance Audit" -Completed
        Write-Log "Governance & Controls audit completed" -Level SUCCESS
        
    } catch {
        Write-Log "Error in Governance audit: $_" -Level ERROR
    }
    
    return $results
}

#endregion

#region Audit Functions - 2. Identity & Access

function Get-IdentityAccessAudit {
    <#
    .SYNOPSIS
        Audits identity and access controls for Teams
    #>
    [CmdletBinding()]
    param()
    
    Write-Log "Starting Identity & Access audit..." -Level INFO
    $results = @{}
    
    try {
        # 2.1 Entra ID Posture - Conditional Access
        Write-Progress -Activity "Identity & Access Audit" -Status "Analyzing Conditional Access policies..." -PercentComplete 10
        
        $caPolicies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop
        $results['ConditionalAccessPolicies'] = $caPolicies
        
        # Check for Teams-specific CA policies
        $teamsPolicies = $caPolicies | Where-Object { 
            $_.Conditions.Applications.IncludeApplications -contains 'Microsoft Teams' -or
            $_.Conditions.Applications.IncludeApplications -contains 'cc15fd57-2c6c-4117-a88c-83b1d56b4bbe'
        }
        
        if ($teamsPolicies.Count -eq 0) {
            Add-Finding -Category 'Identity & Access' -SubCategory 'Conditional Access' `
                -Finding 'No Teams-specific Conditional Access policies' `
                -CurrentState "No CA policies targeting Teams application" `
                -RecommendedState "Minimum 2-3 CA policies for Teams (MFA, device compliance, location-based)" `
                -BestPractice "Implement dedicated CA policies for Teams to control access based on context" `
                -RAGStatus 'Red' `
                -Impact 'Critical' -Likelihood 'High' `
                -Remediation 'Create CA policies for Teams requiring MFA and compliant devices' `
                -Timeline 'Quick Win (30 days)'
        }
        
        # 2.2 MFA Coverage
        Write-Progress -Activity "Identity & Access Audit" -Status "Checking MFA coverage..." -PercentComplete 30
        
        try {
            $mfaUsers = Get-MgReportAuthenticationMethodUserRegistrationDetail -All -ErrorAction Stop
            $totalUsers = ($mfaUsers | Measure-Object).Count
            $mfaEnabledUsers = ($mfaUsers | Where-Object { $_.IsMfaRegistered -eq $true } | Measure-Object).Count
            
            if ($totalUsers -gt 0) {
                $mfaPercentage = [math]::Round(($mfaEnabledUsers / $totalUsers) * 100, 2)
                $results['MFACoverage'] = @{
                    TotalUsers = $totalUsers
                    MFAEnabled = $mfaEnabledUsers
                    Percentage = $mfaPercentage
                }
                
                $ragStatus = Get-RAGRating -CompliancePercentage $mfaPercentage
                
                Add-Finding -Category 'Identity & Access' -SubCategory 'MFA' `
                    -Finding 'MFA registration coverage assessment' `
                    -CurrentState "$mfaEnabledUsers of $totalUsers users registered ($mfaPercentage%)" `
                    -RecommendedState "100% MFA registration for all users" `
                    -BestPractice "Enforce MFA for all users accessing Teams and M365 services" `
                    -RAGStatus $ragStatus `
                    -Impact 'Critical' -Likelihood 'High' `
                    -Remediation 'Implement MFA registration campaign and CA policy enforcement' `
                    -Timeline $(if ($mfaPercentage -lt 80) { 'Quick Win (30 days)' } else { 'Quick Win (30 days)' })
            }
        } catch {
            Write-Log "Unable to retrieve MFA statistics: $_" -Level WARNING
        }
        
        # 2.3 Guest Access Configuration
        Write-Progress -Activity "Identity & Access Audit" -Status "Analyzing guest access settings..." -PercentComplete 50
        
        $guestConfig = Get-CsTeamsClientConfiguration -Identity Global -ErrorAction SilentlyContinue
        $tenantConfig = Get-CsTenant -ErrorAction SilentlyContinue
        
        if ($tenantConfig) {
            $allowGuestAccess = $tenantConfig.AllowGuestUser
            $results['GuestAccessEnabled'] = $allowGuestAccess
            
            if ($allowGuestAccess -eq $true) {
                # Get actual guest users
                $guestUsers = Get-MgUser -Filter "userType eq 'Guest'" -All -ErrorAction SilentlyContinue
                $guestCount = ($guestUsers | Measure-Object).Count
                $results['GuestUserCount'] = $guestCount
                
                Add-Finding -Category 'Identity & Access' -SubCategory 'Guest Access' `
                    -Finding 'Guest access enabled with active guests' `
                    -CurrentState "Guest access enabled, $guestCount guest users present" `
                    -RecommendedState "Guest access with time-limited invitations and access reviews" `
                    -BestPractice "Enable guest access with strict governance: expiration, reviews, and MFA requirements" `
                    -RAGStatus 'Amber' `
                    -Impact 'Medium' -Likelihood 'Medium' `
                    -Remediation 'Implement guest access reviews, expiration policies, and MFA requirements' `
                    -Timeline 'Medium Term (90 days)'
            }
        }
        
        # 2.4 External Access (Federation)
        Write-Progress -Activity "Identity & Access Audit" -Status "Checking external access settings..." -PercentComplete 70
        
        $federationConfig = Get-CsTenantFederationConfiguration -ErrorAction SilentlyContinue
        if ($federationConfig) {
            $allowFederation = $federationConfig.AllowFederatedUsers
            $allowPublicUsers = $federationConfig.AllowPublicUsers
            
            $results['ExternalAccess'] = @{
                FederationEnabled = $allowFederation
                PublicUsersAllowed = $allowPublicUsers
            }
            
            if ($allowPublicUsers -eq $true) {
                Add-Finding -Category 'Identity & Access' -SubCategory 'External Access' `
                    -Finding 'Public user communication enabled' `
                    -CurrentState "Communication with Skype for Business Online users allowed" `
                    -RecommendedState "Restrict to specific trusted domains only" `
                    -BestPractice "Limit external access to known business partners using allowed domain list" `
                    -RAGStatus 'Amber' `
                    -Impact 'Medium' -Likelihood 'Medium' `
                    -Remediation 'Review business requirements and implement domain whitelist for external access' `
                    -Timeline 'Medium Term (90 days)'
            }
        }
        
        Write-Progress -Activity "Identity & Access Audit" -Completed
        Write-Log "Identity & Access audit completed" -Level SUCCESS
        
    } catch {
        Write-Log "Error in Identity & Access audit: $_" -Level ERROR
    }
    
    return $results
}

#endregion

#region Audit Functions - 3. Teams & Channels Architecture

function Get-TeamsArchitectureAudit {
    <#
    .SYNOPSIS
        Audits Teams and channels architecture, lifecycle, and governance
    #>
    [CmdletBinding()]
    param()
    
    Write-Log "Starting Teams & Channels Architecture audit..." -Level INFO
    $results = @{}
    
    try {
        # 3.1 Team Lifecycle
        Write-Progress -Activity "Teams Architecture Audit" -Status "Analyzing team lifecycle policies..." -PercentComplete 10
        
        # Get all teams
        $teams = Get-Team -ErrorAction Stop
        $totalTeams = ($teams | Measure-Object).Count
        $results['TotalTeams'] = $totalTeams
        
        Write-Log "Found $totalTeams teams in tenant" -Level INFO
        
        # Check for expiration policy
        try {
            $expirationPolicy = Get-MgGroupLifecyclePolicy -All -ErrorAction SilentlyContinue
            $results['ExpirationPolicyExists'] = ($expirationPolicy | Measure-Object).Count -gt 0
            
            if (($expirationPolicy | Measure-Object).Count -eq 0) {
                Add-Finding -Category 'Teams Architecture' -SubCategory 'Team Lifecycle' `
                    -Finding 'No team expiration policy configured' `
                    -CurrentState "No automatic team expiration or renewal policy" `
                    -RecommendedState "365-day expiration policy with owner renewal" `
                    -BestPractice "Implement expiration policies to prevent team sprawl and inactive teams" `
                    -RAGStatus 'Amber' `
                    -Impact 'Medium' -Likelihood 'High' `
                    -Remediation 'Configure group expiration policy in Entra ID' `
                    -Timeline 'Quick Win (30 days)'
            }
        } catch {
            Write-Log "Unable to check expiration policy: $_" -Level WARNING
        }
        
        # Check for teams without owners
        Write-Progress -Activity "Teams Architecture Audit" -Status "Checking team ownership..." -PercentComplete 30
        
        $orphanedTeams = @()
        $teamsWithSingleOwner = @()
        
        $i = 0
        foreach ($team in $teams) {
            $i++
            $percentComplete = [math]::Round(($i / $totalTeams) * 40) + 30
            Write-Progress -Activity "Teams Architecture Audit" -Status "Checking team $($team.DisplayName)..." -PercentComplete $percentComplete
            
            try {
                $owners = Get-TeamUser -GroupId $team.GroupId -Role Owner -ErrorAction SilentlyContinue
                
                if (($owners | Measure-Object).Count -eq 0) {
                    $orphanedTeams += $team
                } elseif (($owners | Measure-Object).Count -eq 1) {
                    $teamsWithSingleOwner += $team
                }
            } catch {
                Write-Log "Unable to check owners for team $($team.DisplayName): $_" -Level WARNING
            }
        }
        
        $results['OrphanedTeams'] = $orphanedTeams.Count
        $results['TeamsWithSingleOwner'] = $teamsWithSingleOwner.Count
        
        if ($orphanedTeams.Count -gt 0) {
            Add-Finding -Category 'Teams Architecture' -SubCategory 'Team Lifecycle' `
                -Finding 'Teams without owners detected' `
                -CurrentState "$($orphanedTeams.Count) teams have no owners" `
                -RecommendedState "All teams must have at least 2 owners" `
                -BestPractice "Every team should have minimum 2 owners for business continuity" `
                -RAGStatus 'Red' `
                -Impact 'High' -Likelihood 'High' `
                -Remediation 'Identify and assign owners to orphaned teams or archive if no longer needed' `
                -Timeline 'Quick Win (30 days)'
        }
        
        if ($teamsWithSingleOwner.Count -gt 0) {
            Add-Finding -Category 'Teams Architecture' -SubCategory 'Team Lifecycle' `
                -Finding 'Teams with single owner detected' `
                -CurrentState "$($teamsWithSingleOwner.Count) teams have only one owner" `
                -RecommendedState "All teams should have minimum 2 owners" `
                -BestPractice "Multiple owners ensure business continuity if primary owner leaves" `
                -RAGStatus 'Amber' `
                -Impact 'Medium' -Likelihood 'Medium' `
                -Remediation 'Add secondary owners to teams with single ownership' `
                -Timeline 'Quick Win (30 days)'
        }
        
        # 3.2 Naming Convention
        Write-Progress -Activity "Teams Architecture Audit" -Status "Checking naming conventions..." -PercentComplete 75
        
        # Check if naming policy exists
        try {
            $namingPolicy = Get-MgGroupSetting -All -ErrorAction SilentlyContinue | 
                Where-Object { $_.DisplayName -eq 'Group.Unified' }
            
            $hasNamingPolicy = $false
            if ($namingPolicy) {
                $prefixSuffix = $namingPolicy.Values | Where-Object { $_.Name -eq 'PrefixSuffixNamingRequirement' }
                if ($prefixSuffix.Value) {
                    $hasNamingPolicy = $true
                }
            }
            
            $results['NamingPolicyExists'] = $hasNamingPolicy
            
            if (-not $hasNamingPolicy) {
                Add-Finding -Category 'Teams Architecture' -SubCategory 'Naming Convention' `
                    -Finding 'No team naming policy configured' `
                    -CurrentState "Teams can be created with any name" `
                    -RecommendedState "Enforced naming policy with prefix/suffix requirements" `
                    -BestPractice "Implement naming convention to improve discoverability and organization" `
                    -RAGStatus 'Amber' `
                    -Impact 'Low' -Likelihood 'High' `
                    -Remediation 'Define and implement team naming policy in Entra ID' `
                    -Timeline 'Medium Term (90 days)'
            }
        } catch {
            Write-Log "Unable to check naming policy: $_" -Level WARNING
        }
        
        # 3.3 Channel Analysis
        Write-Progress -Activity "Teams Architecture Audit" -Status "Analyzing channel configuration..." -PercentComplete 90
        
        $totalChannels = 0
        $privateChannels = 0
        $sharedChannels = 0
        
        foreach ($team in $teams[0..([Math]::Min(50, $teams.Count - 1))]) { # Sample first 50 teams for performance
            try {
                $channels = Get-TeamChannel -GroupId $team.GroupId -ErrorAction SilentlyContinue
                $totalChannels += ($channels | Measure-Object).Count
                $privateChannels += ($channels | Where-Object { $_.MembershipType -eq 'Private' } | Measure-Object).Count
                $sharedChannels += ($channels | Where-Object { $_.MembershipType -eq 'Shared' } | Measure-Object).Count
            } catch {
                Write-Log "Unable to get channels for team $($team.DisplayName): $_" -Level WARNING
            }
        }
        
        $results['ChannelStats'] = @{
            TotalChannels = $totalChannels
            PrivateChannels = $privateChannels
            SharedChannels = $sharedChannels
        }
        
        if ($sharedChannels -gt 0) {
            Add-Finding -Category 'Teams Architecture' -SubCategory 'Channel Strategy' `
                -Finding 'Shared channels in use' `
                -CurrentState "$sharedChannels shared channels detected" `
                -RecommendedState "Shared channels with documented cross-org collaboration policy" `
                -BestPractice "Shared channels require governance due to cross-tenant data sharing implications" `
                -RAGStatus 'Amber' `
                -Impact 'Medium' -Likelihood 'Medium' `
                -Remediation 'Document shared channel usage policy and implement approval workflow' `
                -Timeline 'Medium Term (90 days)'
        }
        
        Write-Progress -Activity "Teams Architecture Audit" -Completed
        Write-Log "Teams & Channels Architecture audit completed" -Level SUCCESS
        
    } catch {
        Write-Log "Error in Teams Architecture audit: $_" -Level ERROR
    }
    
    return $results
}

#endregion

#region Audit Functions - 4. Security & Compliance

function Get-SecurityComplianceAudit {
    <#
    .SYNOPSIS
        Audits security and compliance controls for Teams
    #>
    [CmdletBinding()]
    param()
    
    Write-Log "Starting Security & Compliance audit..." -Level INFO
    $results = @{}
    
    try {
        # 4.1 Threat Protection - Safe Links/Attachments
        Write-Progress -Activity "Security & Compliance Audit" -Status "Checking threat protection settings..." -PercentComplete 10
        
        try {
            $atpPolicies = Get-AtpPolicyForO365 -ErrorAction SilentlyContinue
            $results['DefenderForOffice365'] = $atpPolicies -ne $null
            
            if ($atpPolicies) {
                $enableTeamsProtection = $atpPolicies.EnableATPForTeams
                $results['TeamsATPEnabled'] = $enableTeamsProtection
                
                if (-not $enableTeamsProtection) {
                    Add-Finding -Category 'Security & Compliance' -SubCategory 'Threat Protection' `
                        -Finding 'Defender for Office 365 not protecting Teams' `
                        -CurrentState "ATP for Teams disabled" `
                        -RecommendedState "ATP for Teams enabled with Safe Links and Safe Attachments" `
                        -BestPractice "Enable Defender for Office 365 protection for Teams messages and files" `
                        -RAGStatus 'Red' `
                        -Impact 'Critical' -Likelihood 'High' `
                        -Remediation 'Enable ATP for Teams in Defender for Office 365 policy' `
                        -Timeline 'Quick Win (30 days)'
                }
            }
        } catch {
            Write-Log "Unable to check ATP policies: $_" -Level WARNING
            
            Add-Finding -Category 'Security & Compliance' -SubCategory 'Threat Protection' `
                -Finding 'Defender for Office 365 status unknown' `
                -CurrentState "Unable to verify Defender for Office 365 configuration" `
                -RecommendedState "Defender for Office 365 Plan 2 with Teams protection" `
                -BestPractice "Implement comprehensive threat protection for all workloads" `
                -RAGStatus 'Amber' `
                -Impact 'High' -Likelihood 'Medium' `
                -Remediation 'Verify Defender for Office 365 licensing and configuration' `
                -Timeline 'Quick Win (30 days)'
        }
        
        # 4.2 Data Protection - Sensitivity Labels
        Write-Progress -Activity "Security & Compliance Audit" -Status "Checking sensitivity labels..." -PercentComplete 30
        
        try {
            # Note: Sensitivity labels require Security & Compliance Center PowerShell
            # This is a simplified check
            $labelPolicies = Get-LabelPolicy -ErrorAction SilentlyContinue
            $results['SensitivityLabelPolicies'] = ($labelPolicies | Measure-Object).Count
            
            if (($labelPolicies | Measure-Object).Count -eq 0) {
                Add-Finding -Category 'Security & Compliance' -SubCategory 'Data Protection' `
                    -Finding 'No sensitivity label policies configured' `
                    -CurrentState "Sensitivity labels not deployed" `
                    -RecommendedState "Minimum 3-5 sensitivity labels with auto-labeling" `
                    -BestPractice "Deploy sensitivity labels for Teams, SharePoint, and Office apps" `
                    -RAGStatus 'Red' `
                    -Impact 'High' -Likelihood 'High' `
                    -Remediation 'Design and deploy sensitivity label framework' `
                    -Timeline 'Medium Term (90 days)'
            }
        } catch {
            Write-Log "Unable to check sensitivity labels - may require Security & Compliance PowerShell: $_" -Level WARNING
        }
        
        # 4.3 DLP Policies
        Write-Progress -Activity "Security & Compliance Audit" -Status "Checking DLP policies..." -PercentComplete 50
        
        try {
            $dlpPolicies = Get-DlpCompliancePolicy -ErrorAction SilentlyContinue
            $teamsDlpPolicies = $dlpPolicies | Where-Object { 
                $_.Workload -contains 'TeamChat' -or 
                $_.Workload -contains 'TeamsChannel'
            }
            
            $results['DLPPoliciesForTeams'] = ($teamsDlpPolicies | Measure-Object).Count
            
            if (($teamsDlpPolicies | Measure-Object).Count -eq 0) {
                Add-Finding -Category 'Security & Compliance' -SubCategory 'Data Protection' `
                    -Finding 'No DLP policies protecting Teams' `
                    -CurrentState "Teams chat and channels not covered by DLP" `
                    -RecommendedState "Minimum 2-3 DLP policies for sensitive data types (PII, financial, health)" `
                    -BestPractice "Implement DLP policies for Teams to prevent data leakage" `
                    -RAGStatus 'Red' `
                    -Impact 'Critical' -Likelihood 'High' `
                    -Remediation 'Create DLP policies for Teams covering key sensitive data types' `
                    -Timeline 'Quick Win (30 days)'
            } else {
                Add-Finding -Category 'Security & Compliance' -SubCategory 'Data Protection' `
                    -Finding 'DLP policies active for Teams' `
                    -CurrentState "$($teamsDlpPolicies.Count) DLP policies protecting Teams" `
                    -RecommendedState "Regular review and tuning of DLP policies" `
                    -BestPractice "Continuously monitor and refine DLP policies based on incidents" `
                    -RAGStatus 'Green' `
                    -Impact 'Medium' -Likelihood 'Low' `
                    -Remediation 'Schedule quarterly DLP policy review and optimization' `
                    -Timeline 'Quick Win (30 days)'
            }
        } catch {
            Write-Log "Unable to check DLP policies: $_" -Level WARNING
        }
        
        # 4.4 Retention Policies
        Write-Progress -Activity "Security & Compliance Audit" -Status "Checking retention policies..." -PercentComplete 70
        
        try {
            $retentionPolicies = Get-RetentionCompliancePolicy -ErrorAction SilentlyContinue
            $teamsRetentionPolicies = $retentionPolicies | Where-Object { 
                $_.TeamsChannelLocationException -ne $null -or 
                $_.TeamsChatLocationException -ne $null
            }
            
            $results['RetentionPoliciesForTeams'] = ($teamsRetentionPolicies | Measure-Object).Count
            
            if (($teamsRetentionPolicies | Measure-Object).Count -eq 0) {
                Add-Finding -Category 'Security & Compliance' -SubCategory 'Data Protection' `
                    -Finding 'No retention policies for Teams' `
                    -CurrentState "Teams messages and files not subject to retention" `
                    -RecommendedState "Retention policies based on business and legal requirements" `
                    -BestPractice "Implement retention policies for regulatory compliance and eDiscovery" `
                    -RAGStatus 'Amber' `
                    -Impact 'High' -Likelihood 'Medium' `
                    -Remediation 'Define retention requirements and implement appropriate policies' `
                    -Timeline 'Medium Term (90 days)'
            }
        } catch {
            Write-Log "Unable to check retention policies: $_" -Level WARNING
        }
        
        # 4.5 Audit Logging
        Write-Progress -Activity "Security & Compliance Audit" -Status "Checking audit logging..." -PercentComplete 90
        
        try {
            $auditConfig = Get-AdminAuditLogConfig -ErrorAction SilentlyContinue
            $results['UnifiedAuditLogEnabled'] = $auditConfig.UnifiedAuditLogIngestionEnabled
            
            if (-not $auditConfig.UnifiedAuditLogIngestionEnabled) {
                Add-Finding -Category 'Security & Compliance' -SubCategory 'Audit & Logging' `
                    -Finding 'Unified Audit Log not enabled' `
                    -CurrentState "Audit logging disabled" `
                    -RecommendedState "Unified Audit Log enabled with SIEM integration" `
                    -BestPractice "Enable audit logging for compliance and security monitoring" `
                    -RAGStatus 'Red' `
                    -Impact 'Critical' -Likelihood 'High' `
                    -Remediation 'Enable Unified Audit Log in Security & Compliance Center' `
                    -Timeline 'Quick Win (30 days)'
            }
        } catch {
            Write-Log "Unable to check audit configuration: $_" -Level WARNING
        }
        
        Write-Progress -Activity "Security & Compliance Audit" -Completed
        Write-Log "Security & Compliance audit completed" -Level SUCCESS
        
    } catch {
        Write-Log "Error in Security & Compliance audit: $_" -Level ERROR
    }
    
    return $results
}

#endregion

#region Audit Functions - 5. Messaging & Meetings

function Get-MessagingMeetingsAudit {
    <#
    .SYNOPSIS
        Audits messaging, meetings, and collaboration control policies
    #>
    [CmdletBinding()]
    param()
    
    Write-Log "Starting Messaging & Meetings audit..." -Level INFO
    $results = @{}
    
    try {
        # 5.1 Chat Policy
        Write-Progress -Activity "Messaging & Meetings Audit" -Status "Analyzing chat policies..." -PercentComplete 10
        
        $messagingPolicy = Get-CsTeamsMessagingPolicy -Identity Global -ErrorAction SilentlyContinue
        if ($messagingPolicy) {
            $results['GlobalMessagingPolicy'] = $messagingPolicy
            
            # Check Giphy settings
            if ($messagingPolicy.AllowGiphy -eq $true -and $messagingPolicy.GiphyRatingType -eq 'Strict') {
                Add-Finding -Category 'Messaging & Meetings' -SubCategory 'Chat Policy' `
                    -Finding 'Giphy enabled with content filtering' `
                    -CurrentState "Giphy allowed with Strict content rating" `
                    -RecommendedState "Giphy with Strict rating or disabled for regulated industries" `
                    -BestPractice "Balance user experience with content governance requirements" `
                    -RAGStatus 'Green' `
                    -Impact 'Low' -Likelihood 'Low' `
                    -Remediation 'No action required - current configuration appropriate' `
                    -Timeline 'Quick Win (30 days)'
            } elseif ($messagingPolicy.AllowGiphy -eq $true -and $messagingPolicy.GiphyRatingType -ne 'Strict') {
                Add-Finding -Category 'Messaging & Meetings' -SubCategory 'Chat Policy' `
                    -Finding 'Giphy enabled without strict content filtering' `
                    -CurrentState "Giphy allowed with $($messagingPolicy.GiphyRatingType) content rating" `
                    -RecommendedState "Giphy with Strict rating or disabled" `
                    -BestPractice "Use Strict content filtering to prevent inappropriate content" `
                    -RAGStatus 'Amber' `
                    -Impact 'Low' -Likelihood 'Medium' `
                    -Remediation 'Update Giphy content rating to Strict' `
                    -Timeline 'Quick Win (30 days)'
            }
            
            # Check chat deletion settings
            if ($messagingPolicy.AllowUserDeleteChat -eq $true -and $messagingPolicy.AllowUserDeleteMessages -eq $true) {
                Add-Finding -Category 'Messaging & Meetings' -SubCategory 'Chat Policy' `
                    -Finding 'Users can delete messages and chats' `
                    -CurrentState "Message and chat deletion enabled for users" `
                    -RecommendedState "Consider retention requirements before allowing deletion" `
                    -BestPractice "Align deletion capabilities with compliance and retention policies" `
                    -RAGStatus 'Amber' `
                    -Impact 'Medium' -Likelihood 'Medium' `
                    -Remediation 'Review deletion settings against compliance requirements' `
                    -Timeline 'Medium Term (90 days)'
            }
        }
        
        # 5.2 Meetings Policy
        Write-Progress -Activity "Messaging & Meetings Audit" -Status "Analyzing meeting policies..." -PercentComplete 40
        
        $meetingPolicy = Get-CsTeamsMeetingPolicy -Identity Global -ErrorAction SilentlyContinue
        if ($meetingPolicy) {
            $results['GlobalMeetingPolicy'] = $meetingPolicy
            
            # Check lobby settings
            if ($meetingPolicy.AutoAdmittedUsers -eq 'EveryoneInCompany') {
                Add-Finding -Category 'Messaging & Meetings' -SubCategory 'Meeting Policy' `
                    -Finding 'Automatic meeting admission for organization users' `
                    -CurrentState "Company users bypass lobby automatically" `
                    -RecommendedState "Consider lobby controls for sensitive meetings" `
                    -BestPractice "Balance user experience with security for different meeting types" `
                    -RAGStatus 'Amber' `
                    -Impact 'Low' -Likelihood 'Medium' `
                    -Remediation 'Create separate meeting policies for different security levels' `
                    -Timeline 'Medium Term (90 days)'
            } elseif ($meetingPolicy.AutoAdmittedUsers -eq 'Everyone') {
                Add-Finding -Category 'Messaging & Meetings' -SubCategory 'Meeting Policy' `
                    -Finding 'No lobby controls - all participants auto-admitted' `
                    -CurrentState "All participants bypass lobby (including external)" `
                    -RecommendedState "Organization users only, or organizer approval required" `
                    -BestPractice "Require lobby for external participants to prevent unauthorized access" `
                    -RAGStatus 'Red' `
                    -Impact 'Medium' -Likelihood 'High' `
                    -Remediation 'Update AutoAdmittedUsers to EveryoneInCompany or EveryoneInSameAndFederatedCompany' `
                    -Timeline 'Quick Win (30 days)'
            }
            
            # Check recording settings
            if ($meetingPolicy.AllowCloudRecording -eq $true) {
                $recordingLocation = if ($meetingPolicy.RecordingStorageMode -eq 'Stream') { 'OneDrive/SharePoint' } else { 'Unknown' }
                
                Add-Finding -Category 'Messaging & Meetings' -SubCategory 'Meeting Policy' `
                    -Finding 'Meeting recording enabled' `
                    -CurrentState "Cloud recording enabled, stored in $recordingLocation" `
                    -RecommendedState "Recording enabled with appropriate retention and DLP policies" `
                    -BestPractice "Ensure recording storage aligns with data governance requirements" `
                    -RAGStatus 'Amber' `
                    -Impact 'Medium' -Likelihood 'Low' `
                    -Remediation 'Verify retention policies cover meeting recordings' `
                    -Timeline 'Medium Term (90 days)'
            }
            
            # Check end-to-end encryption
            if ($meetingPolicy.AllowMeetEndToEndEncryption -eq $false) {
                Add-Finding -Category 'Messaging & Meetings' -SubCategory 'Meeting Policy' `
                    -Finding 'End-to-end encryption not available' `
                    -CurrentState "E2E encryption disabled for meetings" `
                    -RecommendedState "E2E encryption available for sensitive meetings" `
                    -BestPractice "Enable E2E encryption option for highly confidential discussions" `
                    -RAGStatus 'Amber' `
                    -Impact 'Medium' -Likelihood 'Low' `
                    -Remediation 'Enable AllowMeetEndToEndEncryption for sensitive meeting scenarios' `
                    -Timeline 'Quick Win (30 days)'
            }
            
            # Check watermarking
            if ($meetingPolicy.AllowWatermarkForCameraVideo -eq $false -and $meetingPolicy.AllowWatermarkForScreenSharing -eq $false) {
                Add-Finding -Category 'Messaging & Meetings' -SubCategory 'Meeting Policy' `
                    -Finding 'Meeting watermarking not enabled' `
                    -CurrentState "No watermarks on video or screen sharing" `
                    -RecommendedState "Watermarking available for sensitive meetings" `
                    -BestPractice "Enable watermarking to deter unauthorized recording and screenshots" `
                    -RAGStatus 'Amber' `
                    -Impact 'Low' -Likelihood 'Low' `
                    -Remediation 'Enable watermarking capabilities for high-sensitivity scenarios' `
                    -Timeline 'Medium Term (90 days)'
            }
        }
        
        # 5.3 Live Events Policy
        Write-Progress -Activity "Messaging & Meetings Audit" -Status "Checking live events configuration..." -PercentComplete 70
        
        $liveEventsPolicy = Get-CsTeamsMeetingBroadcastPolicy -Identity Global -ErrorAction SilentlyContinue
        if ($liveEventsPolicy) {
            $results['LiveEventsPolicy'] = $liveEventsPolicy
            
            if ($liveEventsPolicy.AllowBroadcastScheduling -eq $true) {
                Add-Finding -Category 'Messaging & Meetings' -SubCategory 'Live Events' `
                    -Finding 'Live events enabled' `
                    -CurrentState "Users can schedule live events" `
                    -RecommendedState "Live events with approval workflow for large broadcasts" `
                    -BestPractice "Implement governance for live events due to wide potential audience" `
                    -RAGStatus 'Amber' `
                    -Impact 'Medium' -Likelihood 'Low' `
                    -Remediation 'Consider approval process for live events, especially external broadcasts' `
                    -Timeline 'Medium Term (90 days)'
            }
        }
        
        Write-Progress -Activity "Messaging & Meetings Audit" -Completed
        Write-Log "Messaging & Meetings audit completed" -Level SUCCESS
        
    } catch {
        Write-Log "Error in Messaging & Meetings audit: $_" -Level ERROR
    }
    
    return $results
}

#endregion

#region Audit Functions - 6. Voice & Telephony

function Get-VoiceTelephonyAudit {
    <#
    .SYNOPSIS
        Audits Teams Phone and voice configuration (if enabled)
    #>
    [CmdletBinding()]
    param()
    
    Write-Log "Starting Voice & Telephony audit..." -Level INFO
    $results = @{}
    
    try {
        # 6.1 PSTN Connectivity
        Write-Progress -Activity "Voice & Telephony Audit" -Status "Checking PSTN configuration..." -PercentComplete 10
        
        # Check for voice routing policies
        $voiceRoutingPolicies = Get-CsOnlineVoiceRoutingPolicy -ErrorAction SilentlyContinue
        $results['VoiceRoutingPolicies'] = ($voiceRoutingPolicies | Measure-Object).Count
        
        if (($voiceRoutingPolicies | Measure-Object).Count -gt 0) {
            Write-Log "Voice routing detected - Teams Phone is in use" -Level INFO
            
            # Check PSTN usage
            $pstnUsages = Get-CsOnlinePstnUsage -ErrorAction SilentlyContinue
            $results['PSTNUsages'] = ($pstnUsages.Usage | Measure-Object).Count
            
            # Check voice routes
            $voiceRoutes = Get-CsOnlineVoiceRoute -ErrorAction SilentlyContinue
            $results['VoiceRoutes'] = ($voiceRoutes | Measure-Object).Count
            
            if (($voiceRoutes | Measure-Object).Count -eq 0 -and ($voiceRoutingPolicies | Measure-Object).Count -gt 0) {
                Add-Finding -Category 'Voice & Telephony' -SubCategory 'PSTN Configuration' `
                    -Finding 'Voice routing policies without voice routes' `
                    -CurrentState "Voice routing policies exist but no voice routes configured" `
                    -RecommendedState "Properly configured voice routes for call routing" `
                    -BestPractice "Voice routing policies require associated voice routes to function" `
                    -RAGStatus 'Red' `
                    -Impact 'High' -Likelihood 'High' `
                    -Remediation 'Configure voice routes to match routing policies' `
                    -Timeline 'Quick Win (30 days)'
            }
        }
        
        # 6.2 Emergency Calling (E911)
        Write-Progress -Activity "Voice & Telephony Audit" -Status "Checking emergency calling configuration..." -PercentComplete 40
        
        # Check emergency calling policies
        $emergencyCallingPolicies = Get-CsTeamsEmergencyCallingPolicy -ErrorAction SilentlyContinue
        $results['EmergencyCallingPolicies'] = ($emergencyCallingPolicies | Measure-Object).Count
        
        # Check emergency call routing policies
        $emergencyCallRoutingPolicies = Get-CsTeamsEmergencyCallRoutingPolicy -ErrorAction SilentlyContinue
        $results['EmergencyCallRoutingPolicies'] = ($emergencyCallRoutingPolicies | Measure-Object).Count
        
        if (($voiceRoutingPolicies | Measure-Object).Count -gt 0) {
            if (($emergencyCallingPolicies | Measure-Object).Count -eq 0) {
                Add-Finding -Category 'Voice & Telephony' -SubCategory 'Emergency Calling' `
                    -Finding 'No emergency calling policies configured' `
                    -CurrentState "Teams Phone in use without emergency calling policies" `
                    -RecommendedState "Emergency calling policies for all voice-enabled users" `
                    -BestPractice "Configure E911 for regulatory compliance (Kari's Law, RAY BAUM's Act)" `
                    -RAGStatus 'Red' `
                    -Impact 'Critical' -Likelihood 'High' `
                    -Remediation 'Implement emergency calling policies with location detection' `
                    -Timeline 'Quick Win (30 days)'
            }
            
            if (($emergencyCallRoutingPolicies | Measure-Object).Count -eq 0) {
                Add-Finding -Category 'Voice & Telephony' -SubCategory 'Emergency Calling' `
                    -Finding 'No emergency call routing policies configured' `
                    -CurrentState "Teams Phone in use without emergency call routing" `
                    -RecommendedState "Emergency call routing policies for proper E911 routing" `
                    -BestPractice "Emergency calls must route to appropriate PSAP based on location" `
                    -RAGStatus 'Red' `
                    -Impact 'Critical' -Likelihood 'High' `
                    -Remediation 'Configure emergency call routing policies for all locations' `
                    -Timeline 'Quick Win (30 days)'
            }
        }
        
        # 6.3 Network Sites and Locations
        Write-Progress -Activity "Voice & Telephony Audit" -Status "Checking network site configuration..." -PercentComplete 70
        
        if (($voiceRoutingPolicies | Measure-Object).Count -gt 0) {
            $networkSites = Get-CsTenantNetworkSite -ErrorAction SilentlyContinue
            $results['NetworkSites'] = ($networkSites | Measure-Object).Count
            
            if (($networkSites | Measure-Object).Count -eq 0) {
                Add-Finding -Category 'Voice & Telephony' -SubCategory 'Emergency Calling' `
                    -Finding 'No network sites defined' `
                    -CurrentState "Teams Phone without network topology configuration" `
                    -RecommendedState "Network sites configured for dynamic E911 location" `
                    -BestPractice "Define network sites for automatic location detection" `
                    -RAGStatus 'Red' `
                    -Impact 'Critical' -Likelihood 'High' `
                    -Remediation 'Configure network sites and subnets for location-based services' `
                    -Timeline 'Medium Term (90 days)'
            }
        }
        
        # 6.4 Call Queues and Auto Attendants
        Write-Progress -Activity "Voice & Telephony Audit" -Status "Checking call queues and auto attendants..." -PercentComplete 90
        
        $callQueues = Get-CsCallQueue -ErrorAction SilentlyContinue
        $autoAttendants = Get-CsAutoAttendant -ErrorAction SilentlyContinue
        
        $results['CallQueues'] = ($callQueues | Measure-Object).Count
        $results['AutoAttendants'] = ($autoAttendants | Measure-Object).Count
        
        Write-Log "Found $($results['CallQueues']) call queues and $($results['AutoAttendants']) auto attendants" -Level INFO
        
        Write-Progress -Activity "Voice & Telephony Audit" -Completed
        Write-Log "Voice & Telephony audit completed" -Level SUCCESS
        
    } catch {
        Write-Log "Error in Voice & Telephony audit: $_" -Level ERROR
    }
    
    return $results
}

#endregion

#region Audit Functions - 7. Devices & Rooms

function Get-DevicesRoomsAudit {
    <#
    .SYNOPSIS
        Audits Teams devices, phones, and rooms configuration
    #>
    [CmdletBinding()]
    param()
    
    Write-Log "Starting Devices & Rooms audit..." -Level INFO
    $results = @{}
    
    try {
        # 7.1 Teams Devices
        Write-Progress -Activity "Devices & Rooms Audit" -Status "Checking Teams devices..." -PercentComplete 20
        
        # Get Teams IP phones
        $teamsPhones = Get-CsOnlineUser -Filter {LineUri -ne $null} -ErrorAction SilentlyContinue | 
            Where-Object { $_.EnterpriseVoiceEnabled -eq $true }
        
        $results['TeamsPhones'] = ($teamsPhones | Measure-Object).Count
        Write-Log "Found $($results['TeamsPhones']) Teams Phone users" -Level INFO
        
        # 7.2 Teams Rooms
        Write-Progress -Activity "Devices & Rooms Audit" -Status "Checking Teams Rooms..." -PercentComplete 50
        
        # Get Teams Rooms (resource accounts with Room type)
        $teamsRooms = Get-CsOnlineUser -Filter {AccountType -eq 'ResourceAccount'} -ErrorAction SilentlyContinue
        $results['TeamsRooms'] = ($teamsRooms | Measure-Object).Count
        
        if (($teamsRooms | Measure-Object).Count -gt 0) {
            Write-Log "Found $($results['TeamsRooms']) Teams Rooms accounts" -Level INFO
            
            Add-Finding -Category 'Devices & Rooms' -SubCategory 'Teams Rooms' `
                -Finding 'Teams Rooms deployed' `
                -CurrentState "$($results['TeamsRooms']) Teams Rooms resource accounts" `
                -RecommendedState "Teams Rooms with update policies and monitoring" `
                -BestPractice "Implement update policies and monitoring for Teams Rooms devices" `
                -RAGStatus 'Amber' `
                -Impact 'Medium' -Likelihood 'Medium' `
                -Remediation 'Configure Teams Rooms update policies and health monitoring' `
                -Timeline 'Medium Term (90 days)'
        }
        
        # 7.3 IP Phone Policies
        Write-Progress -Activity "Devices & Rooms Audit" -Status "Checking IP phone policies..." -PercentComplete 80
        
        $ipPhonePolicies = Get-CsTeamsIPPhonePolicy -ErrorAction SilentlyContinue
        $results['IPPhonePolicies'] = ($ipPhonePolicies | Measure-Object).Count
        
        if (($teamsPhones | Measure-Object).Count -gt 0 -and ($ipPhonePolicies | Measure-Object).Count -eq 0) {
            Add-Finding -Category 'Devices & Rooms' -SubCategory 'IP Phone Policies' `
                -Finding 'Teams phones without dedicated policies' `
                -CurrentState "Teams phones in use without IP phone policies" `
                -RecommendedState "IP phone policies for device configuration management" `
                -BestPractice "Use IP phone policies to centrally manage device settings" `
                -RAGStatus 'Amber' `
                -Impact 'Low' -Likelihood 'Medium' `
                -Remediation 'Create and assign Teams IP phone policies' `
                -Timeline 'Medium Term (90 days)'
        }
        
        Write-Progress -Activity "Devices & Rooms Audit" -Completed
        Write-Log "Devices & Rooms audit completed" -Level SUCCESS
        
    } catch {
        Write-Log "Error in Devices & Rooms audit: $_" -Level ERROR
    }
    
    return $results
}

#endregion

#region Audit Functions - 8. Apps & Integration

function Get-AppsIntegrationAudit {
    <#
    .SYNOPSIS
        Audits Teams apps, integrations, and Power Platform governance
    #>
    [CmdletBinding()]
    param()
    
    Write-Log "Starting Apps & Integration audit..." -Level INFO
    $results = @{}
    
    try {
        # 8.1 App Permission Policies
        Write-Progress -Activity "Apps & Integration Audit" -Status "Checking app permission policies..." -PercentComplete 20
        
        $appPermissionPolicies = Get-CsTeamsAppPermissionPolicy -ErrorAction SilentlyContinue
        $results['AppPermissionPolicies'] = ($appPermissionPolicies | Measure-Object).Count
        
        # Check global policy
        $globalAppPolicy = $appPermissionPolicies | Where-Object { $_.Identity -eq 'Global' }
        
        if ($globalAppPolicy) {
            $defaultOrgWideApps = $globalAppPolicy.DefaultOrgWideAppsState
            
            if ($defaultOrgWideAppsState -eq 'AllApps') {
                Add-Finding -Category 'Apps & Integration' -SubCategory 'App Governance' `
                    -Finding 'All apps allowed by default' `
                    -CurrentState "Organization-wide apps setting: Allow all apps" `
                    -RecommendedState "Selective app approval with allowed list" `
                    -BestPractice "Use allow-listing approach for third-party apps to reduce risk" `
                    -RAGStatus 'Amber' `
                    -Impact 'Medium' -Likelihood 'Medium' `
                    -Remediation 'Review and implement app allow-list with approval workflow' `
                    -Timeline 'Medium Term (90 days)'
            }
        }
        
        # 8.2 App Setup Policies
        Write-Progress -Activity "Apps & Integration Audit" -Status "Checking app setup policies..." -PercentComplete 50
        
        $appSetupPolicies = Get-CsTeamsAppSetupPolicy -ErrorAction SilentlyContinue
        $results['AppSetupPolicies'] = ($appSetupPolicies | Measure-Object).Count
        
        # 8.3 Custom Apps
        Write-Progress -Activity "Apps & Integration Audit" -Status "Checking custom apps..." -PercentComplete 70
        
        try {
            $teamsApps = Get-TeamsApp -ErrorAction SilentlyContinue
            $customApps = $teamsApps | Where-Object { $_.DistributionMethod -eq 'Organization' }
            
            $results['CustomApps'] = ($customApps | Measure-Object).Count
            
            if (($customApps | Measure-Object).Count -gt 0) {
                Add-Finding -Category 'Apps & Integration' -SubCategory 'App Governance' `
                    -Finding 'Custom Teams apps deployed' `
                    -CurrentState "$($customApps.Count) custom apps in organization catalog" `
                    -RecommendedState "Custom apps with security review and approval process" `
                    -BestPractice "Implement security review for all custom apps before deployment" `
                    -RAGStatus 'Amber' `
                    -Impact 'Medium' -Likelihood 'Medium' `
                    -Remediation 'Establish app review process covering permissions and data access' `
                    -Timeline 'Medium Term (90 days)'
            }
        } catch {
            Write-Log "Unable to retrieve Teams apps: $_" -Level WARNING
        }
        
        Write-Progress -Activity "Apps & Integration Audit" -Completed
        Write-Log "Apps & Integration audit completed" -Level SUCCESS
        
    } catch {
        Write-Log "Error in Apps & Integration audit: $_" -Level ERROR
    }
    
    return $results
}

#endregion

#region Audit Functions - 9. SharePoint & OneDrive

function Get-SharePointOneDriveAudit {
    <#
    .SYNOPSIS
        Audits SharePoint and OneDrive settings that impact Teams
    #>
    [CmdletBinding()]
    param()
    
    Write-Log "Starting SharePoint & OneDrive audit..." -Level INFO
    $results = @{}
    
    try {
        # 9.1 External Sharing Settings
        Write-Progress -Activity "SharePoint & OneDrive Audit" -Status "Checking sharing settings..." -PercentComplete 20
        
        try {
            $spoTenant = Get-SPOTenant -ErrorAction SilentlyContinue
            
            if ($spoTenant) {
                $results['SharePointSharingCapability'] = $spoTenant.SharingCapability
                $results['OneDriveSharingCapability'] = $spoTenant.OneDriveSharingCapability
                
                if ($spoTenant.SharingCapability -eq 'ExternalUserAndGuestSharing') {
                    Add-Finding -Category 'SharePoint & OneDrive' -SubCategory 'External Sharing' `
                        -Finding 'Anyone links enabled for SharePoint' `
                        -CurrentState "SharePoint allows 'Anyone' sharing links" `
                        -RecommendedState "Limit to authenticated external users or specific people" `
                        -BestPractice "Disable 'Anyone' links to prevent accidental data exposure" `
                        -RAGStatus 'Amber' `
                        -Impact 'High' -Likelihood 'Medium' `
                        -Remediation 'Change sharing capability to ExternalUserSharingOnly or more restrictive' `
                        -Timeline 'Quick Win (30 days)'
                } elseif ($spoTenant.SharingCapability -eq 'Disabled') {
                    Add-Finding -Category 'SharePoint & OneDrive' -SubCategory 'External Sharing' `
                        -Finding 'External sharing completely disabled' `
                        -CurrentState "SharePoint external sharing is disabled" `
                        -RecommendedState "Consider enabling authenticated external user sharing if needed" `
                        -BestPractice "Balance collaboration needs with security requirements" `
                        -RAGStatus 'Green' `
                        -Impact 'Low' -Likelihood 'Low' `
                        -Remediation 'Current configuration is secure - review against business requirements' `
                        -Timeline 'Quick Win (30 days)'
                }
                
                # Check OneDrive sharing
                if ($spoTenant.OneDriveSharingCapability -eq 'ExternalUserAndGuestSharing') {
                    Add-Finding -Category 'SharePoint & OneDrive' -SubCategory 'External Sharing' `
                        -Finding 'Anyone links enabled for OneDrive' `
                        -CurrentState "OneDrive allows 'Anyone' sharing links" `
                        -RecommendedState "Limit to authenticated external users or more restrictive" `
                        -BestPractice "OneDrive 'Anyone' links can lead to uncontrolled data sharing" `
                        -RAGStatus 'Amber' `
                        -Impact 'High' -Likelihood 'Medium' `
                        -Remediation 'Change OneDrive sharing to ExternalUserSharingOnly or more restrictive' `
                        -Timeline 'Quick Win (30 days)'
                }
                
                # Check default link type
                if ($spoTenant.DefaultSharingLinkType -eq 'AnonymousAccess') {
                    Add-Finding -Category 'SharePoint & OneDrive' -SubCategory 'External Sharing' `
                        -Finding 'Default sharing link type is Anyone' `
                        -CurrentState "Default link type set to anonymous access" `
                        -RecommendedState "Default to 'Specific people' or 'Organization'" `
                        -BestPractice "Default link type should be most restrictive for security" `
                        -RAGStatus 'Amber' `
                        -Impact 'Medium' -Likelihood 'High' `
                        -Remediation 'Change DefaultSharingLinkType to Direct or Internal' `
                        -Timeline 'Quick Win (30 days)'
                }
            }
        } catch {
            Write-Log "Unable to retrieve SharePoint tenant settings: $_" -Level WARNING
        }
        
        # 9.2 Versioning Settings
        Write-Progress -Activity "SharePoint & OneDrive Audit" -Status "Checking versioning configuration..." -PercentComplete 60
        
        # This would require checking individual sites - providing general guidance
        Add-Finding -Category 'SharePoint & OneDrive' -SubCategory 'Data Protection' `
            -Finding 'Versioning configuration review needed' `
            -CurrentState "Site-level versioning settings vary" `
            -RecommendedState "Versioning enabled with appropriate limits (50-100 versions)" `
            -BestPractice "Enable versioning for all Teams-connected sites for data recovery" `
            -RAGStatus 'Amber' `
            -Impact 'Medium' -Likelihood 'Low' `
            -Remediation 'Audit and standardize versioning settings across Teams sites' `
            -Timeline 'Medium Term (90 days)'
        
        Write-Progress -Activity "SharePoint & OneDrive Audit" -Completed
        Write-Log "SharePoint & OneDrive audit completed" -Level SUCCESS
        
    } catch {
        Write-Log "Error in SharePoint & OneDrive audit: $_" -Level ERROR
    }
    
    return $results
}

#endregion

#region Reporting Functions

function Export-AuditReports {
    <#
    .SYNOPSIS
        Generates comprehensive Excel reports from audit findings
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )
    
    Write-Log "Generating audit reports..." -Level INFO
    
    try {
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $reportPath = Join-Path -Path $OutputPath -ChildPath "Reports\TeamsAudit_$timestamp.xlsx"
        
        # Prepare data for export
        $ragData = $Script:RAGFindings | Select-Object Category, SubCategory, Finding, CurrentState, 
            RecommendedState, BestPractice, RAGStatus, Impact, Likelihood, RiskLevel, RiskScore, 
            Remediation, Timeline, DateIdentified
        
        # Calculate summary statistics
        $totalFindings = ($ragData | Measure-Object).Count
        $redFindings = ($ragData | Where-Object { $_.RAGStatus -eq 'Red' } | Measure-Object).Count
        $amberFindings = ($ragData | Where-Object { $_.RAGStatus -eq 'Amber' } | Measure-Object).Count
        $greenFindings = ($ragData | Where-Object { $_.RAGStatus -eq 'Green' } | Measure-Object).Count
        
        $criticalRisks = ($ragData | Where-Object { $_.RiskLevel -eq 'Critical' } | Measure-Object).Count
        $highRisks = ($ragData | Where-Object { $_.RiskLevel -eq 'High' } | Measure-Object).Count
        $mediumRisks = ($ragData | Where-Object { $_.RiskLevel -eq 'Medium' } | Measure-Object).Count
        $lowRisks = ($ragData | Where-Object { $_.RiskLevel -eq 'Low' } | Measure-Object).Count
        
        # Create summary sheet data
        $summary = @(
            [PSCustomObject]@{
                'Metric' = 'Total Findings'
                'Value' = $totalFindings
                'Notes' = 'Total audit findings identified'
            }
            [PSCustomObject]@{
                'Metric' = 'Red (Critical Issues)'
                'Value' = $redFindings
                'Notes' = 'Immediate attention required'
            }
            [PSCustomObject]@{
                'Metric' = 'Amber (Warnings)'
                'Value' = $amberFindings
                'Notes' = 'Should be addressed'
            }
            [PSCustomObject]@{
                'Metric' = 'Green (Compliant)'
                'Value' = $greenFindings
                'Notes' = 'Meeting best practices'
            }
            [PSCustomObject]@{
                'Metric' = ''
                'Value' = ''
                'Notes' = ''
            }
            [PSCustomObject]@{
                'Metric' = 'Critical Risks'
                'Value' = $criticalRisks
                'Notes' = 'Highest priority remediation'
            }
            [PSCustomObject]@{
                'Metric' = 'High Risks'
                'Value' = $highRisks
                'Notes' = 'Significant risk exposure'
            }
            [PSCustomObject]@{
                'Metric' = 'Medium Risks'
                'Value' = $mediumRisks
                'Notes' = 'Moderate risk'
            }
            [PSCustomObject]@{
                'Metric' = 'Low Risks'
                'Value' = $lowRisks
                'Notes' = 'Minor concerns'
            }
            [PSCustomObject]@{
                'Metric' = ''
                'Value' = ''
                'Notes' = ''
            }
            [PSCustomObject]@{
                'Metric' = 'Audit Date'
                'Value' = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                'Notes' = 'Report generation timestamp'
            }
            [PSCustomObject]@{
                'Metric' = 'Tenant'
                'Value' = $Script:AuditResults['TenantInfo'].DisplayName
                'Notes' = $Script:AuditResults['TenantInfo'].Id
            }
        )
        
        # Export summary sheet
        Write-Progress -Activity "Generating Reports" -Status "Creating executive summary..." -PercentComplete 10
        $summary | Export-Excel -Path $reportPath -WorksheetName 'Executive Summary' -AutoSize -FreezeTopRow -BoldTopRow
        
        # Export RAG findings
        Write-Progress -Activity "Generating Reports" -Status "Exporting RAG findings..." -PercentComplete 30
        $ragData | Export-Excel -Path $reportPath -WorksheetName 'RAG Findings' -AutoSize -FreezeTopRow -BoldTopRow `
            -ConditionalText $(
                New-ConditionalText -Text 'Red' -BackgroundColor Red -ConditionalTextColor White
                New-ConditionalText -Text 'Amber' -BackgroundColor Orange -ConditionalTextColor Black
                New-ConditionalText -Text 'Green' -BackgroundColor Green -ConditionalTextColor White
                New-ConditionalText -Text 'Critical' -BackgroundColor DarkRed -ConditionalTextColor White
            )
        
        # Export findings by category
        Write-Progress -Activity "Generating Reports" -Status "Creating category breakdowns..." -PercentComplete 50
        
        $categories = $ragData | Group-Object -Property Category
        foreach ($category in $categories) {
            $worksheetName = $category.Name -replace '[\\\/\:\*\?\[\]]', '_'
            $worksheetName = $worksheetName.Substring(0, [Math]::Min($worksheetName.Length, 31))
            
            $category.Group | Export-Excel -Path $reportPath -WorksheetName $worksheetName -AutoSize -FreezeTopRow -BoldTopRow
        }
        
        # Export remediation roadmap
        Write-Progress -Activity "Generating Reports" -Status "Creating remediation roadmap..." -PercentComplete 70
        
        $quickWins = $ragData | Where-Object { $_.Timeline -eq 'Quick Win (30 days)' } | 
            Sort-Object -Property RiskScore -Descending |
            Select-Object Category, Finding, RAGStatus, RiskLevel, Remediation
        
        $mediumTerm = $ragData | Where-Object { $_.Timeline -eq 'Medium Term (90 days)' } | 
            Sort-Object -Property RiskScore -Descending |
            Select-Object Category, Finding, RAGStatus, RiskLevel, Remediation
        
        $strategic = $ragData | Where-Object { $_.Timeline -eq 'Strategic (6-12 months)' } | 
            Sort-Object -Property RiskScore -Descending |
            Select-Object Category, Finding, RAGStatus, RiskLevel, Remediation
        
        if ($quickWins) {
            $quickWins | Export-Excel -Path $reportPath -WorksheetName 'Quick Wins (30d)' -AutoSize -FreezeTopRow -BoldTopRow
        }
        
        if ($mediumTerm) {
            $mediumTerm | Export-Excel -Path $reportPath -WorksheetName 'Medium Term (90d)' -AutoSize -FreezeTopRow -BoldTopRow
        }
        
        if ($strategic) {
            $strategic | Export-Excel -Path $reportPath -WorksheetName 'Strategic (6-12m)' -AutoSize -FreezeTopRow -BoldTopRow
        }
        
        Write-Progress -Activity "Generating Reports" -Completed
        
        Write-Log "Audit report generated: $reportPath" -Level SUCCESS
        return $reportPath
        
    } catch {
        Write-Log "Error generating reports: $_" -Level ERROR
        return $null
    }
}

#endregion

#region Main Execution

function Invoke-MainAudit {
    <#
    .SYNOPSIS
        Main audit orchestration function
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$TenantId,
        
        [Parameter(Mandatory = $false)]
        [bool]$IncludeVoice,
        
        [Parameter(Mandatory = $false)]
        [bool]$IncludeDevices,
        
        [Parameter(Mandatory = $false)]
        [string]$OutputPath
    )
    
    Write-Host "`n==================================================================" -ForegroundColor Cyan
    Write-Host "  Microsoft Teams Comprehensive Security & Compliance Audit" -ForegroundColor Cyan
    Write-Host "  Author: Kelvin Chigorimbo" -ForegroundColor Cyan
    Write-Host "==================================================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Initialize output directory
    $outputDir = Initialize-OutputDirectory -CustomPath $OutputPath
    
    # Check required modules
    if (-not (Test-RequiredModules)) {
        Write-Log "Missing required modules - cannot continue" -Level ERROR
        return
    }
    
    # Connect to required services
    if (-not (Connect-RequiredServices -TenantId $TenantId)) {
        Write-Log "Failed to connect to required services - cannot continue" -Level ERROR
        return
    }
    
    Write-Host "`n------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "  Starting comprehensive audit - this may take several minutes..." -ForegroundColor Yellow
    Write-Host "------------------------------------------------------------------`n" -ForegroundColor Yellow
    
    # Execute audit sections with progress tracking
    $auditSections = @(
        @{ Name = '1. Governance & Controls'; Function = { Get-GovernanceAudit }; Weight = 15 }
        @{ Name = '2. Identity & Access'; Function = { Get-IdentityAccessAudit }; Weight = 15 }
        @{ Name = '3. Teams & Channels Architecture'; Function = { Get-TeamsArchitectureAudit }; Weight = 15 }
        @{ Name = '4. Security & Compliance'; Function = { Get-SecurityComplianceAudit }; Weight = 15 }
        @{ Name = '5. Messaging & Meetings'; Function = { Get-MessagingMeetingsAudit }; Weight = 10 }
        @{ Name = '8. Apps & Integration'; Function = { Get-AppsIntegrationAudit }; Weight = 10 }
        @{ Name = '9. SharePoint & OneDrive'; Function = { Get-SharePointOneDriveAudit }; Weight = 10 }
    )
    
    if ($IncludeVoice) {
        $auditSections += @{ Name = '6. Voice & Telephony'; Function = { Get-VoiceTelephonyAudit }; Weight = 10 }
    }
    
    if ($IncludeDevices) {
        $auditSections += @{ Name = '7. Devices & Rooms'; Function = { Get-DevicesRoomsAudit }; Weight = 10 }
    }
    
    $currentProgress = 0
    foreach ($section in $auditSections) {
        Write-Progress -Activity "Teams Comprehensive Audit" -Status "Executing: $($section.Name)" -PercentComplete $currentProgress
        
        try {
            $result = & $section.Function
            $Script:AuditResults[$section.Name] = $result
        } catch {
            Write-Log "Error in section $($section.Name): $_" -Level ERROR
        }
        
        $currentProgress += $section.Weight
    }
    
    Write-Progress -Activity "Teams Comprehensive Audit" -Completed
    
    # Generate reports
    Write-Host "`n------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "  Generating audit reports..." -ForegroundColor Yellow
    Write-Host "------------------------------------------------------------------`n" -ForegroundColor Yellow
    
    $reportPath = Export-AuditReports -OutputPath $outputDir
    
    # Display summary
    Write-Host "`n==================================================================" -ForegroundColor Green
    Write-Host "  Audit Complete!" -ForegroundColor Green
    Write-Host "==================================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Total Findings: $($Script:RAGFindings.Count)" -ForegroundColor White
    Write-Host "  Red (Critical):    $(($Script:RAGFindings | Where-Object { $_.RAGStatus -eq 'Red' } | Measure-Object).Count)" -ForegroundColor Red
    Write-Host "  Amber (Warning):   $(($Script:RAGFindings | Where-Object { $_.RAGStatus -eq 'Amber' } | Measure-Object).Count)" -ForegroundColor Yellow
    Write-Host "  Green (Compliant): $(($Script:RAGFindings | Where-Object { $_.RAGStatus -eq 'Green' } | Measure-Object).Count)" -ForegroundColor Green
    Write-Host ""
    Write-Host "Risk Distribution:" -ForegroundColor White
    Write-Host "  Critical: $(($Script:RAGFindings | Where-Object { $_.RiskLevel -eq 'Critical' } | Measure-Object).Count)" -ForegroundColor Red
    Write-Host "  High:     $(($Script:RAGFindings | Where-Object { $_.RiskLevel -eq 'High' } | Measure-Object).Count)" -ForegroundColor Red
    Write-Host "  Medium:   $(($Script:RAGFindings | Where-Object { $_.RiskLevel -eq 'Medium' } | Measure-Object).Count)" -ForegroundColor Yellow
    Write-Host "  Low:      $(($Script:RAGFindings | Where-Object { $_.RiskLevel -eq 'Low' } | Measure-Object).Count)" -ForegroundColor Green
    Write-Host ""
    Write-Host "Output Location:" -ForegroundColor White
    Write-Host "  Reports: $reportPath" -ForegroundColor Cyan
    Write-Host "  Logs:    $Script:LogFile" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Execution Time: $(((Get-Date) - $Script:StartTime).ToString('hh\:mm\:ss'))" -ForegroundColor White
    Write-Host "==================================================================" -ForegroundColor Green
    Write-Host ""
    
    # Disconnect sessions
    Write-Log "Disconnecting from Microsoft services..." -Level INFO
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        Write-Log "Disconnected successfully" -Level SUCCESS
    } catch {
        Write-Log "Error during disconnect: $_" -Level WARNING
    }
    
    Write-Log "Audit completed successfully" -Level SUCCESS
}

# Execute main audit
Invoke-MainAudit -TenantId $TenantId -IncludeVoice:$IncludeVoice -IncludeDevices:$IncludeDevices -OutputPath $OutputPath

#endregion