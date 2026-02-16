<#
.SYNOPSIS
    Microsoft Teams Tenant Compliance and Configuration Audit Script
.DESCRIPTION
    Comprehensive audit of Microsoft Teams tenant configuration generating RAG-rated
    Excel and Word reports with E911 compliance, PIM JIT analysis, and remediation roadmaps.
.NOTES
    Script Name    : Invoke-TeamsComplianceAudit.ps1
    Version        : 3.1.0
    Author         : Kelvin Chigorimbo, Cloud Solutions Architect
    Creation Date  : February 2026
.LINK
    https://learn.microsoft.com/en-us/microsoftteams/
    https://learn.microsoft.com/en-us/microsoftteams/manage-emergency-calling-policies
    https://learn.microsoft.com/en-us/entra/id-governance/privileged-identity-management/pim-configure
#>
#Requires -Version 7.0
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) "Invoke-TeamsComplianceAudit"),
    [switch]$Interactive = $true,
    [string]$TenantId, [string]$ClientId, [string]$ClientSecret, [string]$CertificateThumbprint
)

#Start Script Configuration
$script:ScriptName = "Invoke-TeamsComplianceAudit"
$script:ScriptVersion = "3.1.0"
$script:ScriptAuthor = "Kelvin Chigorimbo, Cloud Solutions Architect"
$script:StartTime = Get-Date
$script:LogFile = $null
$script:AuditResults = @{}
$script:ErrorCount = 0
$script:WarningCount = 0
$script:Colors = @{ Success="Green"; Warning="Yellow"; Error="Red"; Info="Cyan"; Header="Magenta" }
# Microsoft brand colour palette mapped to nearest System.Drawing.KnownColor for PSWriteWord compatibility
# Original hex: Red=#f65314, Green=#7cbb00, Blue=#00a1f1, Yellow=#ffbb00, Grey=#737373
$script:MSColors = @{ Red="OrangeRed"; Green="LimeGreen"; Blue="DodgerBlue"; Yellow="Gold"; Grey="Gray" }
#End Script Configuration

#Start Logging Functions
function Initialize-LogFile {
    [CmdletBinding()] param()
    try {
        if (-not (Test-Path $OutputPath -PathType Container)) { New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null }
        $ts = $script:StartTime.ToString("yyyyMMdd_HHmmss")
        $script:LogFile = Join-Path $OutputPath "TeamsAudit_$ts.log"
        "================================================================================$([Environment]::NewLine)Microsoft Teams Tenant Compliance Audit Log$([Environment]::NewLine)================================================================================$([Environment]::NewLine)Script  : $($script:ScriptName) v$($script:ScriptVersion)$([Environment]::NewLine)Author  : $($script:ScriptAuthor)$([Environment]::NewLine)Start   : $($script:StartTime.ToString('yyyy-MM-dd HH:mm:ss'))$([Environment]::NewLine)Output  : $OutputPath$([Environment]::NewLine)Host    : $($env:COMPUTERNAME) / $($env:USERNAME)$([Environment]::NewLine)PS Ver  : $($PSVersionTable.PSVersion)$([Environment]::NewLine)================================================================================`n" | Out-File $script:LogFile -Encoding UTF8
        Write-Host "  Log file: $($script:LogFile)" -ForegroundColor Cyan
        return $true
    } catch { Write-Host "  Log init failed: $($_.Exception.Message)" -ForegroundColor Red; return $false }
}

function Write-AuditLog {
    [CmdletBinding()]
    param([AllowEmptyString()][string]$Message="", [ValidateSet("Info","Warning","Error","Success","Header")][string]$Level="Info", [switch]$NoConsole)
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    if ([string]::IsNullOrEmpty($Message)) { if(-not $NoConsole){Write-Host ""}; if($script:LogFile -and (Test-Path $script:LogFile)){""|Out-File $script:LogFile -Append -Encoding UTF8}; return }
    if ($script:LogFile -and (Test-Path $script:LogFile)) { "[$ts] [$Level] $Message" | Out-File $script:LogFile -Append -Encoding UTF8 }
    if (-not $NoConsole) {
        switch($Level) { "Success"{Write-Host $Message -ForegroundColor Green} "Warning"{Write-Host $Message -ForegroundColor Yellow;$script:WarningCount++} "Error"{Write-Host $Message -ForegroundColor Red;$script:ErrorCount++} "Header"{Write-Host $Message -ForegroundColor Magenta} default{Write-Host $Message -ForegroundColor Cyan} }
    }
}
#End Logging Functions

#Start Logging Functions
function Initialize-LogFile {
    [CmdletBinding()] param()
    try {
        if (-not (Test-Path $OutputPath -PathType Container)) { New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null }
        $ts = $script:StartTime.ToString("yyyyMMdd_HHmmss")
        $script:LogFile = Join-Path $OutputPath "TeamsAudit_$ts.log"
        "================================================================================" | Out-File $script:LogFile -Encoding UTF8
        "Microsoft Teams Tenant Compliance Audit Log" | Out-File $script:LogFile -Append -Encoding UTF8
        "================================================================================" | Out-File $script:LogFile -Append -Encoding UTF8
        "Script  : $($script:ScriptName) v$($script:ScriptVersion)" | Out-File $script:LogFile -Append -Encoding UTF8
        "Author  : $($script:ScriptAuthor)" | Out-File $script:LogFile -Append -Encoding UTF8
        "Start   : $($script:StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" | Out-File $script:LogFile -Append -Encoding UTF8
        "Output  : $OutputPath" | Out-File $script:LogFile -Append -Encoding UTF8
        "Host    : $($env:COMPUTERNAME) / $($env:USERNAME)" | Out-File $script:LogFile -Append -Encoding UTF8
        "PS Ver  : $($PSVersionTable.PSVersion)" | Out-File $script:LogFile -Append -Encoding UTF8
        "================================================================================" | Out-File $script:LogFile -Append -Encoding UTF8
        Write-Host "  Log file: $($script:LogFile)" -ForegroundColor Cyan
        return $true
    } catch { Write-Host "  Log init failed: $($_.Exception.Message)" -ForegroundColor Red; return $false }
}

function Write-AuditLog {
    [CmdletBinding()]
    param(
        [AllowEmptyString()][string]$Message = "",
        [ValidateSet("Info","Warning","Error","Success","Header")][string]$Level = "Info",
        [switch]$NoConsole
    )
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    if ([string]::IsNullOrEmpty($Message)) {
        if (-not $NoConsole) { Write-Host "" }
        if ($script:LogFile -and (Test-Path $script:LogFile)) { "" | Out-File $script:LogFile -Append -Encoding UTF8 }
        return
    }
    if ($script:LogFile -and (Test-Path $script:LogFile)) {
        "[$ts] [$Level] $Message" | Out-File $script:LogFile -Append -Encoding UTF8
    }
    if (-not $NoConsole) {
        switch ($Level) {
            "Success" { Write-Host $Message -ForegroundColor Green }
            "Warning" { Write-Host $Message -ForegroundColor Yellow; $script:WarningCount++ }
            "Error"   { Write-Host $Message -ForegroundColor Red; $script:ErrorCount++ }
            "Header"  { Write-Host $Message -ForegroundColor Magenta }
            default   { Write-Host $Message -ForegroundColor Cyan }
        }
    }
}
#End Logging Functions

#Start Module Functions
function Test-RequiredModules {
    <#
    .SYNOPSIS
        Validates that all required PowerShell modules are installed and available.
    .DESCRIPTION
        Checks for the presence of each required module and minimum version where applicable.
        Displays a progress bar during validation and reports any missing or outdated modules.
    #>
    [CmdletBinding()] param()
    Write-AuditLog "Validating required PowerShell modules..." -Level Header
    $req = @(
        @{ Name = "MicrosoftTeams"; MinVersion = "5.0.0" }
        @{ Name = "Microsoft.Graph.Authentication"; MinVersion = $null }
        @{ Name = "Microsoft.Graph.Teams"; MinVersion = $null }
        @{ Name = "Microsoft.Graph.Identity.DirectoryManagement"; MinVersion = $null }
        @{ Name = "Microsoft.Graph.Identity.Governance"; MinVersion = $null }
        @{ Name = "Microsoft.Graph.DeviceManagement"; MinVersion = $null }
        @{ Name = "ImportExcel"; MinVersion = $null }
        @{ Name = "PSWriteWord"; MinVersion = $null }
    )
    $miss = @(); $i = 0
    foreach ($m in $req) {
        $i++
        Write-Progress -Activity "Checking Modules" -Status $m.Name -PercentComplete ([math]::Round(($i / $req.Count) * 100))
        $inst = Get-Module $m.Name -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $inst) {
            $miss += $m.Name
            Write-AuditLog "  [MISSING] $($m.Name)" -Level Error
        }
        elseif ($m.MinVersion -and $inst.Version -lt [Version]$m.MinVersion) {
            $miss += $m.Name
            Write-AuditLog "  [OUTDATED] $($m.Name) v$($inst.Version) - requires v$($m.MinVersion)+" -Level Warning
        }
        else {
            Write-AuditLog "  [OK] $($m.Name) v$($inst.Version)" -Level Success
        }
    }
    Write-Progress -Activity "Checking Modules" -Completed
    if ($miss.Count -gt 0) {
        Write-AuditLog "Missing modules: $($miss -join ', ')" -Level Error
        Write-AuditLog "Install using: Install-Module -Name <ModuleName> -Scope CurrentUser -Force" -Level Info
        return $false
    }
    Write-AuditLog "All required modules validated." -Level Success
    return $true
}

function Import-RequiredModules {
    <#
    .SYNOPSIS
        Imports all required PowerShell modules into the current session.
    #>
    [CmdletBinding()] param()
    Write-AuditLog "Importing required modules..." -Level Header
    $mods = @("MicrosoftTeams","Microsoft.Graph.Authentication","Microsoft.Graph.Teams",
              "Microsoft.Graph.Identity.DirectoryManagement","Microsoft.Graph.Identity.Governance",
              "Microsoft.Graph.DeviceManagement","ImportExcel","PSWriteWord")
    $errs = @(); $i = 0
    foreach ($m in $mods) {
        $i++
        Write-Progress -Activity "Importing Modules" -Status $m -PercentComplete ([math]::Round(($i / $mods.Count) * 100))
        try {
            Import-Module $m -ErrorAction Stop
            Write-AuditLog "  Imported: $m" -Level Success
        }
        catch {
            $errs += $m
            Write-AuditLog "  Failed: $m - $($_.Exception.Message)" -Level Error
        }
    }
    Write-Progress -Activity "Importing Modules" -Completed
    if ($errs.Count -gt 0) { return $false }
    return $true
}
#End Module Functions

#Start Authentication Functions
function Connect-AuditServices {
    <#
    .SYNOPSIS
        Establishes connections to Microsoft Teams and Microsoft Graph services.
    #>
    [CmdletBinding()] param()
    Write-AuditLog "Connecting to Microsoft services..." -Level Header
    $ok = $true
    # Connect to Microsoft Teams PowerShell
    try {
        Write-AuditLog "  Connecting to Microsoft Teams..." -Level Info
        if ($Interactive) {
            if ($TenantId) { Connect-MicrosoftTeams -TenantId $TenantId -ErrorAction Stop | Out-Null }
            else { Connect-MicrosoftTeams -ErrorAction Stop | Out-Null }
        }
        elseif ($CertificateThumbprint) {
            Connect-MicrosoftTeams -TenantId $TenantId -ApplicationId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorAction Stop | Out-Null
        }
        Write-AuditLog "  Teams connected." -Level Success
    }
    catch { Write-AuditLog "  Teams connection failed: $($_.Exception.Message)" -Level Error; $ok = $false }
    # Connect to Microsoft Graph
    try {
        Write-AuditLog "  Connecting to Microsoft Graph..." -Level Info
        $scopes = @("Team.ReadBasic.All","TeamSettings.Read.All","Policy.Read.All","Directory.Read.All",
                     "RoleManagement.Read.Directory","User.Read.All","Group.Read.All")
        if ($Interactive) {
            if ($TenantId) { Connect-MgGraph -Scopes $scopes -TenantId $TenantId -ErrorAction Stop | Out-Null }
            else { Connect-MgGraph -Scopes $scopes -ErrorAction Stop | Out-Null }
        }
        Write-AuditLog "  Graph connected." -Level Success
    }
    catch { Write-AuditLog "  Graph connection failed: $($_.Exception.Message)" -Level Error; $ok = $false }
    return $ok
}

function Disconnect-AuditServices {
    <#
    .SYNOPSIS
        Disconnects from all Microsoft services to release authenticated sessions.
    #>
    [CmdletBinding()] param()
    Write-AuditLog "Disconnecting services..." -Level Info
    try { Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue } catch {}
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    Write-AuditLog "Services disconnected." -Level Success
}
#End Authentication Functions

#Start Best Practices Reference
function Get-BestPracticesReference {
    <#
    .SYNOPSIS
        Returns comprehensive reference data for defaults, best practices, risk scoring,
        remediation timelines, E911 compliance, and PIM assessment criteria.
    .DESCRIPTION
        Each entry contains: Default, BestPractice, Explanation, Recommendation,
        RiskImpact (1-4), RiskLikelihood (1-4), RemediationTimeline, RemediationEffort,
        SecurityDomain. Used by Get-RAGRating for compliance assessment.
    #>
    [CmdletBinding()] param()

    $bp = @{
        TeamsSettings = @{
            AllowEmailIntoChannel = @{ Default=$true; BestPractice=$true; Explanation="Email-to-channel integration allows forwarding messages into Teams channels."; Recommendation="Enable for information sharing workflows."; RiskImpact=1; RiskLikelihood=1; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="Collaboration" }
            AllowGuestCreateUpdateChannels = @{ Default=$true; BestPractice=$false; Explanation="Controls whether guest users can create or modify channels."; Recommendation="Disable to maintain channel governance."; RiskImpact=2; RiskLikelihood=3; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="External Access" }
            AllowGuestDeleteChannels = @{ Default=$false; BestPractice=$false; Explanation="Controls whether guest users can delete channels."; Recommendation="Keep disabled to prevent accidental data loss."; RiskImpact=4; RiskLikelihood=2; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="Data Protection" }
            AllowResourceAccountSendMessage = @{ Default=$true; BestPractice=$true; Explanation="Permits resource accounts to post messages in channels."; Recommendation="Enable for meeting room and bot integration."; RiskImpact=1; RiskLikelihood=1; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="Collaboration" }
        }
        MeetingPolicies = @{
            AllowAnonymousUsersToJoinMeeting = @{ Default=$true; BestPractice=$false; Explanation="Permits unauthenticated users to join scheduled meetings."; Recommendation="Disable for internal-only policies; enable selectively for external collaboration."; RiskImpact=3; RiskLikelihood=3; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="Identity and Access" }
            AllowAnonymousUsersToStartMeeting = @{ Default=$false; BestPractice=$false; Explanation="Permits unauthenticated users to initiate meetings without the organiser present."; Recommendation="Keep disabled to enforce organiser control."; RiskImpact=4; RiskLikelihood=2; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="Identity and Access" }
            AutoAdmittedUsers = @{ Default="EveryoneInCompanyExcludingGuests"; BestPractice="EveryoneInCompanyExcludingGuests"; Explanation="Defines which participants bypass the meeting lobby."; Recommendation="Set to EveryoneInCompanyExcludingGuests for tenant security."; RiskImpact=3; RiskLikelihood=2; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="Identity and Access" }
            AllowCloudRecording = @{ Default=$true; BestPractice=$true; Explanation="Enables cloud-based meeting recording to OneDrive/SharePoint."; Recommendation="Enable with data retention policies configured."; RiskImpact=2; RiskLikelihood=2; RemediationTimeline="MediumTerm"; RemediationEffort="Medium"; SecurityDomain="Data Protection" }
            AllowRecordingStorageOutsideRegion = @{ Default=$false; BestPractice=$false; Explanation="Permits recording storage in a different geographic region."; Recommendation="Disable for data residency and sovereignty compliance."; RiskImpact=4; RiskLikelihood=2; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="Data Protection" }
            AllowTranscription = @{ Default=$false; BestPractice=$true; Explanation="Enables live transcription during meetings."; Recommendation="Enable for accessibility and compliance record-keeping."; RiskImpact=1; RiskLikelihood=2; RemediationTimeline="MediumTerm"; RemediationEffort="Low"; SecurityDomain="Collaboration" }
            ScreenSharingMode = @{ Default="EntireScreen"; BestPractice="SingleApplication"; Explanation="Defines the scope of screen sharing in meetings."; Recommendation="Set to SingleApplication to minimise inadvertent data exposure."; RiskImpact=3; RiskLikelihood=3; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="Data Protection" }
            AllowExternalParticipantGiveRequestControl = @{ Default=$false; BestPractice=$false; Explanation="Permits external participants to request or receive desktop control."; Recommendation="Keep disabled to prevent unauthorised access."; RiskImpact=3; RiskLikelihood=2; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="External Access" }
            AllowWatermarkForCameraVideo = @{ Default=$false; BestPractice=$true; Explanation="Overlays a watermark on participant video feeds."; Recommendation="Enable for sensitive or confidential meetings."; RiskImpact=2; RiskLikelihood=2; RemediationTimeline="MediumTerm"; RemediationEffort="Medium"; SecurityDomain="Data Protection" }
            AllowWatermarkForScreenSharing = @{ Default=$false; BestPractice=$true; Explanation="Overlays a watermark on shared screen content."; Recommendation="Enable as a deterrent against unauthorised capture."; RiskImpact=2; RiskLikelihood=2; RemediationTimeline="MediumTerm"; RemediationEffort="Medium"; SecurityDomain="Data Protection" }
            DesignatedPresenterRoleMode = @{ Default="EveryoneUserOverride"; BestPractice="OrganizerOnlyUserOverride"; Explanation="Default presenter role assignment for meetings."; Recommendation="Set to OrganizerOnlyUserOverride for content control."; RiskImpact=2; RiskLikelihood=2; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="Collaboration" }
        }
        MessagingPolicies = @{
            AllowUrlPreviews = @{ Default=$true; BestPractice=$true; Explanation="Generates URL preview cards in chat messages."; Recommendation="Enable for link context."; RiskImpact=1; RiskLikelihood=1; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="Collaboration" }
            AllowUserDeleteMessage = @{ Default=$true; BestPractice=$true; Explanation="Permits users to delete sent messages."; Recommendation="Enable; audit logs retain deletion records."; RiskImpact=1; RiskLikelihood=1; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="Data Protection" }
            AllowUserEditMessage = @{ Default=$true; BestPractice=$true; Explanation="Permits users to edit sent messages."; Recommendation="Enable for corrections."; RiskImpact=1; RiskLikelihood=1; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="Collaboration" }
            GiphyRatingType = @{ Default="Moderate"; BestPractice="Strict"; Explanation="Content rating filter for Giphy integration."; Recommendation="Set to Strict for workplace-appropriate content."; RiskImpact=2; RiskLikelihood=2; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="Collaboration" }
            ReadReceiptsEnabledType = @{ Default="UserPreference"; BestPractice="UserPreference"; Explanation="Read receipt display preference."; Recommendation="Allow user preference for adoption."; RiskImpact=1; RiskLikelihood=1; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="Collaboration" }
        }
        ExternalAccess = @{
            AllowTeamsConsumer = @{ Default=$true; BestPractice=$false; Explanation="Permits communication with personal Teams (consumer) accounts."; Recommendation="Disable for enterprise data boundary enforcement."; RiskImpact=3; RiskLikelihood=3; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="External Access" }
            AllowTeamsConsumerInbound = @{ Default=$true; BestPractice=$false; Explanation="Permits inbound messages from personal Teams accounts."; Recommendation="Disable to prevent unsolicited contact from unmanaged accounts."; RiskImpact=3; RiskLikelihood=3; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="External Access" }
            AllowPublicUsers = @{ Default=$true; BestPractice=$false; Explanation="Permits communication with Skype consumer users."; Recommendation="Disable unless business requirement exists."; RiskImpact=2; RiskLikelihood=2; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="External Access" }
            AllowFederatedUsers = @{ Default=$true; BestPractice=$true; Explanation="Permits federation with other Microsoft 365 organisations."; Recommendation="Enable with domain allowlist for sensitive environments."; RiskImpact=2; RiskLikelihood=2; RemediationTimeline="Strategic"; RemediationEffort="High"; SecurityDomain="External Access" }
        }
        GuestAccess = @{
            AllowGuestAccess = @{ Default=$true; BestPractice=$true; Explanation="Master switch for Azure AD B2B guest access to Teams."; Recommendation="Enable with Conditional Access and sensitivity label controls."; RiskImpact=3; RiskLikelihood=2; RemediationTimeline="Strategic"; RemediationEffort="High"; SecurityDomain="External Access" }
        }
        AppPolicies = @{
            AllowSideLoading = @{ Default=$true; BestPractice=$false; Explanation="Permits users to sideload custom apps outside the store."; Recommendation="Disable; use managed deployment channels."; RiskImpact=4; RiskLikelihood=2; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="Application Security" }
            AllowUserPinning = @{ Default=$true; BestPractice=$true; Explanation="Permits users to pin apps to the navigation rail."; Recommendation="Enable for personalised experience."; RiskImpact=1; RiskLikelihood=1; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="Collaboration" }
        }
        EnhancedEncryption = @{
            AllowEndToEndEncryption = @{ Default="Disabled"; BestPractice="DisabledUserOverride"; Explanation="End-to-end encryption for 1:1 VoIP calls."; Recommendation="Enable user override for sensitive communications."; RiskImpact=3; RiskLikelihood=2; RemediationTimeline="MediumTerm"; RemediationEffort="Medium"; SecurityDomain="Data Protection" }
        }
        UpdatePolicies = @{
            AllowPreview = @{ Default=$false; BestPractice=$false; Explanation="Enables preview/beta features in the Teams client."; Recommendation="Disable for production; enable for designated pilot groups."; RiskImpact=2; RiskLikelihood=2; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="Change Management" }
            UseNewTeamsClient = @{ Default="MicrosoftChoice"; BestPractice="NewTeamsAsDefault"; Explanation="Controls which Teams client version is deployed."; Recommendation="Set to NewTeamsAsDefault per Microsoft deprecation timeline."; RiskImpact=2; RiskLikelihood=3; RemediationTimeline="Strategic"; RemediationEffort="High"; SecurityDomain="Change Management" }
        }
        LiveEventsPolicies = @{
            AllowBroadcastScheduling = @{ Default=$true; BestPractice=$true; Explanation="Permits scheduling of Teams Live Events."; Recommendation="Enable for broadcast capability."; RiskImpact=1; RiskLikelihood=1; RemediationTimeline="QuickWin"; RemediationEffort="Low"; SecurityDomain="Collaboration" }
            BroadcastRecordingMode = @{ Default="UserOverride"; BestPractice="AlwaysEnabled"; Explanation="Recording behaviour for Live Events."; Recommendation="Set to AlwaysEnabled for compliance record-keeping."; RiskImpact=3; RiskLikelihood=2; RemediationTimeline="MediumTerm"; RemediationEffort="Medium"; SecurityDomain="Data Protection" }
        }
        # E911 Emergency Calling - assessed only when Teams Phone is enabled
        EmergencyCalling = @{
            EmergencyCallingPolicyConfigured = @{
                Default = $false; BestPractice = $true
                Explanation = "Emergency calling policies define notification behaviour when emergency calls are placed per RAP and Kari/Ray Baum Act compliance."
                Recommendation = "Configure emergency calling policies with notification groups for all locations."
                RiskImpact = 4; RiskLikelihood = 3; RemediationTimeline = "QuickWin"; RemediationEffort = "Medium"; SecurityDomain = "Emergency Services"
            }
            EmergencyCallRoutingPolicyConfigured = @{
                Default = $false; BestPractice = $true
                Explanation = "Emergency call routing policies ensure calls reach the correct Public Safety Answering Point (PSAP)."
                Recommendation = "Configure for all Direct Routing and Operator Connect deployments."
                RiskImpact = 4; RiskLikelihood = 3; RemediationTimeline = "QuickWin"; RemediationEffort = "Medium"; SecurityDomain = "Emergency Services"
            }
            NetworkSitesConfigured = @{
                Default = $false; BestPractice = $true
                Explanation = "Network site configuration enables dynamic emergency address resolution based on network topology."
                Recommendation = "Configure network sites with emergency addresses mapped to IP subnets and WAPs."
                RiskImpact = 4; RiskLikelihood = 3; RemediationTimeline = "MediumTerm"; RemediationEffort = "High"; SecurityDomain = "Emergency Services"
            }
            NotificationGroupConfigured = @{
                Default = $false; BestPractice = $true
                Explanation = "Notification groups alert designated security personnel when emergency calls are placed."
                Recommendation = "Configure security desk notification groups in all emergency calling policies."
                RiskImpact = 4; RiskLikelihood = 2; RemediationTimeline = "QuickWin"; RemediationEffort = "Low"; SecurityDomain = "Emergency Services"
            }
        }
        # PIM Role Assignments - permanent assignments flagged as non-compliant
        PIMAssignments = @{
            PermanentTeamsAdminAssignment = @{
                Default = "Permanent"; BestPractice = "Eligible"
                Explanation = "Permanent (Active) role assignments grant standing administrative privileges without time limits, approval workflows, or re-authentication requirements. This creates persistent attack surface."
                Recommendation = "Convert all permanent Teams Administrator assignments to PIM Eligible (JIT). Just-In-Time access ensures admin privileges are only active during explicitly approved, time-bound activation windows with MFA enforcement."
                RiskImpact = 4; RiskLikelihood = 3; RemediationTimeline = "QuickWin"; RemediationEffort = "Medium"; SecurityDomain = "Identity and Access"
            }
        }
    }
    return $bp
}
#End Best Practices Reference

#Start RAG Assessment Functions
function Get-RAGRating {
    <#
    .SYNOPSIS
        Determines the RAG (Red/Amber/Green) compliance status for a configuration setting.
    .DESCRIPTION
        Compares a current setting value against the best practice reference entry and calculates
        a risk score. All return values are explicitly typed to [string] or [int] to prevent
        System.Object[] serialisation issues when exporting to Excel via ImportExcel.
        
        FIX APPLIED: RiskLevel was previously returning System.Object[] due to implicit type
        coercion in the switch statement. Now uses explicit [string] cast on assignment.
    .PARAMETER CurrentValue
        The current configured value of the setting being assessed.
    .PARAMETER BestPracticeEntry
        Hashtable containing the best practice reference data for the setting.
    .OUTPUTS
        Hashtable with RAGStatus, RiskScore, ComplianceStatus, and RiskLevel - all explicitly typed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowNull()]$CurrentValue,
        [Parameter(Mandatory = $true)][hashtable]$BestPracticeEntry
    )

    [int]$impact = $BestPracticeEntry.RiskImpact
    [int]$likelihood = $BestPracticeEntry.RiskLikelihood
    [int]$riskScore = $impact * $likelihood

    # Determine compliance status by comparing current value to best practice
    [string]$complianceStatus = "Review"
    if ([string]$CurrentValue -eq [string]$BestPracticeEntry.BestPractice) {
        [string]$complianceStatus = "Compliant"
    }
    elseif ([string]$CurrentValue -eq [string]$BestPracticeEntry.Default -and
            [string]$BestPracticeEntry.Default -ne [string]$BestPracticeEntry.BestPractice) {
        [string]$complianceStatus = "Default (Non-Optimal)"
    }
    elseif ([string]$CurrentValue -ne [string]$BestPracticeEntry.BestPractice) {
        [string]$complianceStatus = "Non-Compliant"
    }

    # Determine RAG status based on compliance and risk score
    [string]$ragStatus = "Green"
    if ($complianceStatus -ne "Compliant") {
        if ($riskScore -ge 9) { [string]$ragStatus = "Red" }
        else { [string]$ragStatus = "Amber" }
    }

    # FIX: Explicit [string] cast on RiskLevel prevents System.Object[] in Excel
    # The original code used a switch statement that could return mixed types
    [string]$riskLevel = if ($riskScore -ge 12) { "Critical" }
                         elseif ($riskScore -ge 9) { "High" }
                         elseif ($riskScore -ge 4) { "Medium" }
                         else { "Low" }

    return @{
        RAGStatus        = [string]$ragStatus
        RiskScore        = [int]$riskScore
        ComplianceStatus = [string]$complianceStatus
        RiskLevel        = [string]$riskLevel
    }
}

function Get-RemediationRoadmap {
    <#
    .SYNOPSIS
        Generates a tiered remediation roadmap from all RAG-assessed findings.
    .DESCRIPTION
        Categorises all non-compliant findings into Quick Wins (0-30 days),
        Medium Term (31-90 days), and Strategic (6-12 months) remediation tiers.
        Each tier is sorted by descending risk score for priority-based execution.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][array]$AllFindings)

    Write-AuditLog "  Building remediation roadmap..." -Level Info

    $actionable = $AllFindings | Where-Object { $_.RAGStatus -ne "Green" }
    $quickWins  = @($actionable | Where-Object { $_.RemediationTimeline -eq "QuickWin" }   | Sort-Object { [int]$_.RiskScore } -Descending)
    $mediumTerm = @($actionable | Where-Object { $_.RemediationTimeline -eq "MediumTerm" } | Sort-Object { [int]$_.RiskScore } -Descending)
    $strategic  = @($actionable | Where-Object { $_.RemediationTimeline -eq "Strategic" }  | Sort-Object { [int]$_.RiskScore } -Descending)

    $total = $AllFindings.Count
    $green = @($AllFindings | Where-Object { $_.RAGStatus -eq "Green" }).Count
    $amber = @($AllFindings | Where-Object { $_.RAGStatus -eq "Amber" }).Count
    $red   = @($AllFindings | Where-Object { $_.RAGStatus -eq "Red" }).Count
    $pct   = if ($total -gt 0) { [math]::Round(($green / $total) * 100, 1) } else { 0 }

    Write-AuditLog "    Green: $green ($pct%) | Amber: $amber | Red: $red" -Level Info
    Write-AuditLog "    Quick Wins: $($quickWins.Count) | Medium Term: $($mediumTerm.Count) | Strategic: $($strategic.Count)" -Level Info

    return @{
        QuickWins   = $quickWins
        MediumTerm  = $mediumTerm
        Strategic   = $strategic
        Summary     = @{
            TotalSettings        = $total
            GreenCount           = $green
            AmberCount           = $amber
            RedCount             = $red
            CompliancePercentage = $pct
            QuickWinCount        = $quickWins.Count
            MediumTermCount      = $mediumTerm.Count
            StrategicCount       = $strategic.Count
            HighestRiskScore     = if ($actionable.Count -gt 0) { ($actionable | Measure-Object -Property RiskScore -Maximum).Maximum } else { 0 }
        }
        AllFindings = $AllFindings
    }
}
#End RAG Assessment Functions

#Start Data Collection Functions
function Get-TeamsAuditData {
    <#
    .SYNOPSIS
        Collects Teams inventory data including team count, membership statistics, and archival status.
    #>
    [CmdletBinding()] param()
    Write-AuditLog "  Collecting Teams inventory..." -Level Info
    $result = @{ Teams = @(); Statistics = @{} }
    try {
        $teams = Get-Team -ErrorAction Stop
        $totalTeams = $teams.Count
        $archivedCount = @($teams | Where-Object { $_.Archived -eq $true }).Count
        $result.Teams = $teams
        $result.Statistics = @{
            TotalTeams = $totalTeams
            ActiveTeams = $totalTeams - $archivedCount
            ArchivedTeams = $archivedCount
            PublicTeams = @($teams | Where-Object { $_.Visibility -eq "Public" }).Count
            PrivateTeams = @($teams | Where-Object { $_.Visibility -eq "Private" }).Count
        }
        Write-AuditLog "    Found $totalTeams teams ($archivedCount archived)." -Level Info
    }
    catch { Write-AuditLog "    Teams inventory error: $($_.Exception.Message)" -Level Warning }
    return $result
}

function Get-TeamsSettingsAuditData {
    <#
    .SYNOPSIS
        Collects tenant-level Teams client configuration settings.
    #>
    [CmdletBinding()] param()
    Write-AuditLog "  Collecting Teams settings..." -Level Info
    $result = @{ ClientConfiguration = @{} }
    try {
        $config = Get-CsTeamsClientConfiguration -ErrorAction Stop
        $result.ClientConfiguration = @{
            AllowEmailIntoChannel = $config.AllowEmailIntoChannel
            AllowGuestCreateUpdateChannels = $config.AllowGuestCreateUpdateChannels
            AllowGuestDeleteChannels = $config.AllowGuestDeleteChannels
            AllowResourceAccountSendMessage = $config.AllowResourceAccountSendMessage
        }
        Write-AuditLog "    Settings collected." -Level Success
    }
    catch { Write-AuditLog "    Settings error: $($_.Exception.Message)" -Level Warning }
    return $result
}

function Get-TeamsPoliciesAuditData {
    <#
    .SYNOPSIS
        Collects Teams policy assignments and templates.
    #>
    [CmdletBinding()] param()
    Write-AuditLog "  Collecting Teams policies..." -Level Info
    $result = @{ Templates = @() }
    try {
        $result.Templates = @(Get-CsTeamTemplateList -ErrorAction SilentlyContinue)
        Write-AuditLog "    Found $($result.Templates.Count) templates." -Level Info
    }
    catch { Write-AuditLog "    Policies error: $($_.Exception.Message)" -Level Warning }
    return $result
}

function Get-TeamsUpdatePoliciesAuditData {
    <#
    .SYNOPSIS
        Collects Teams update and client version policies.
    #>
    [CmdletBinding()] param()
    Write-AuditLog "  Collecting update policies..." -Level Info
    $result = @{ UpdatePolicies = @() }
    try {
        $policies = Get-CsTeamsUpdateManagementPolicy -ErrorAction Stop
        foreach ($p in $policies) {
            $result.UpdatePolicies += @{
                Identity = $p.Identity
                AllowPreview = $p.AllowPreview
                UseNewTeamsClient = $p.UseNewTeamsClient
            }
        }
        Write-AuditLog "    Found $($result.UpdatePolicies.Count) update policies." -Level Info
    }
    catch { Write-AuditLog "    Update policies error: $($_.Exception.Message)" -Level Warning }
    return $result
}

function Get-TeamsUpgradeSettingsAuditData {
    <#
    .SYNOPSIS
        Collects Teams upgrade and coexistence configuration.
    #>
    [CmdletBinding()] param()
    Write-AuditLog "  Collecting upgrade settings..." -Level Info
    $result = @{ UpgradeConfiguration = @{} }
    try {
        $config = Get-CsTeamsUpgradeConfiguration -ErrorAction Stop
        $result.UpgradeConfiguration = @{ SfBMeetingJoinUx = $config.SfBMeetingJoinUx }
        $status = Get-CsTeamsUpgradeStatus -ErrorAction SilentlyContinue
        if ($status) { $result.UpgradeStatus = @{ State = $status.State; CoexistenceMode = $status.CoexistenceMode } }
        Write-AuditLog "    Upgrade settings collected." -Level Success
    }
    catch { Write-AuditLog "    Upgrade settings error: $($_.Exception.Message)" -Level Warning }
    return $result
}

function Get-TeamsDevicesAuditData {
    <#
    .SYNOPSIS
        Collects Teams Rooms, Panels, Phones, and Display device configuration data.
    #>
    [CmdletBinding()] param()
    Write-AuditLog "  Collecting device configuration..." -Level Info
    $result = @{ IPPhonePolicy = @(); DeviceSummary = @{} }
    try {
        $phonePolicies = Get-CsTeamsIPPhonePolicy -ErrorAction SilentlyContinue
        if ($phonePolicies) { $result.IPPhonePolicy = @($phonePolicies) }
        Write-AuditLog "    Device data collected." -Level Success
    }
    catch { Write-AuditLog "    Device data error: $($_.Exception.Message)" -Level Warning }
    return $result
}

function Get-TeamsAppManagementAuditData {
    <#
    .SYNOPSIS
        Collects app permission policies, setup policies, and org-wide app settings.
    .DESCRIPTION
        Retrieves Teams app governance configuration including permission policies that
        control which Microsoft, third-party, and custom apps are permitted, and setup
        policies that control sideloading and pinning behaviour.
    #>
    [CmdletBinding()] param()
    Write-AuditLog "  Collecting app management data..." -Level Info
    $result = @{ PermissionPolicies = @(); SetupPolicies = @(); OrgWideSettings = @{} }
    try {
        # App permission policies - extract detailed settings for RAG assessment
        $permPolicies = Get-CsTeamsAppPermissionPolicy -ErrorAction SilentlyContinue
        if ($permPolicies) {
            foreach ($p in $permPolicies) {
                $result.PermissionPolicies += @{
                    Identity = $p.Identity
                    DefaultCatalogAppsType = $p.DefaultCatalogAppsType
                    GlobalCatalogAppsType = $p.GlobalCatalogAppsType
                    PrivateCatalogAppsType = $p.PrivateCatalogAppsType
                }
            }
        }
        # App setup policies - sideloading and pinning controls
        $setupPolicies = Get-CsTeamsAppSetupPolicy -ErrorAction SilentlyContinue
        foreach ($p in $setupPolicies) {
            $result.SetupPolicies += @{
                Identity = $p.Identity
                AllowSideLoading = $p.AllowSideLoading
                AllowUserPinning = $p.AllowUserPinning
            }
        }
        Write-AuditLog "    Found $($result.PermissionPolicies.Count) permission and $($result.SetupPolicies.Count) setup policies." -Level Info
    }
    catch { Write-AuditLog "    App management error: $($_.Exception.Message)" -Level Warning }
    return $result
}

function Get-TeamsMeetingsAuditData {
    <#
    .SYNOPSIS
        Collects meeting policies, audio conferencing, live events policies, and conference bridge data.
    #>
    [CmdletBinding()] param()
    Write-AuditLog "  Collecting meeting configuration..." -Level Info
    $result = @{ MeetingPolicies = @(); AudioConferencingPolicies = @(); LiveEventPolicies = @() }
    try {
        # Meeting policies
        $meetPolicies = Get-CsTeamsMeetingPolicy -ErrorAction Stop
        foreach ($p in $meetPolicies) {
            $result.MeetingPolicies += @{
                Identity = $p.Identity
                AllowAnonymousUsersToJoinMeeting = $p.AllowAnonymousUsersToJoinMeeting
                AllowAnonymousUsersToStartMeeting = $p.AllowAnonymousUsersToStartMeeting
                AutoAdmittedUsers = $p.AutoAdmittedUsers
                AllowCloudRecording = $p.AllowCloudRecording
                AllowRecordingStorageOutsideRegion = $p.AllowRecordingStorageOutsideRegion
                AllowTranscription = $p.AllowTranscription
                ScreenSharingMode = $p.ScreenSharingMode
                AllowExternalParticipantGiveRequestControl = $p.AllowExternalParticipantGiveRequestControl
                AllowWatermarkForCameraVideo = $p.AllowWatermarkForCameraVideo
                AllowWatermarkForScreenSharing = $p.AllowWatermarkForScreenSharing
                DesignatedPresenterRoleMode = $p.DesignatedPresenterRoleMode
            }
        }
        # Live event policies
        $livePolicies = Get-CsTeamsMeetingBroadcastPolicy -ErrorAction SilentlyContinue
        foreach ($p in $livePolicies) {
            $result.LiveEventPolicies += @{
                Identity = $p.Identity
                AllowBroadcastScheduling = $p.AllowBroadcastScheduling
                BroadcastRecordingMode = $p.BroadcastRecordingMode
            }
        }
        Write-AuditLog "    Found $($result.MeetingPolicies.Count) meeting and $($result.LiveEventPolicies.Count) live event policies." -Level Info
    }
    catch { Write-AuditLog "    Meeting data error: $($_.Exception.Message)" -Level Warning }
    return $result
}

function Get-TeamsMessagingAuditData {
    <#
    .SYNOPSIS
        Collects messaging policy configurations across the tenant.
    #>
    [CmdletBinding()] param()
    Write-AuditLog "  Collecting messaging policies..." -Level Info
    $result = @{ MessagingPolicies = @() }
    try {
        $msgPolicies = Get-CsTeamsMessagingPolicy -ErrorAction Stop
        foreach ($p in $msgPolicies) {
            $result.MessagingPolicies += @{
                Identity = $p.Identity
                AllowUrlPreviews = $p.AllowUrlPreviews
                AllowUserDeleteMessage = $p.AllowUserDeleteMessage
                AllowUserEditMessage = $p.AllowUserEditMessage
                GiphyRatingType = $p.GiphyRatingType
                ReadReceiptsEnabledType = $p.ReadReceiptsEnabledType
            }
        }
        Write-AuditLog "    Found $($result.MessagingPolicies.Count) messaging policies." -Level Info
    }
    catch { Write-AuditLog "    Messaging error: $($_.Exception.Message)" -Level Warning }
    return $result
}

function Get-TeamsEncryptionAuditData {
    <#
    .SYNOPSIS
        Collects enhanced encryption policy configurations for end-to-end call encryption.
    #>
    [CmdletBinding()] param()
    Write-AuditLog "  Collecting encryption policies..." -Level Info
    $result = @{ EncryptionPolicies = @() }
    try {
        $encPolicies = Get-CsTeamsEnhancedEncryptionPolicy -ErrorAction SilentlyContinue
        if ($encPolicies) { $result.EncryptionPolicies = @($encPolicies) }
        Write-AuditLog "    Found $($result.EncryptionPolicies.Count) encryption policies." -Level Info
    }
    catch { Write-AuditLog "    Encryption error: $($_.Exception.Message)" -Level Warning }
    return $result
}

function Get-TeamsPhoneAuditData {
    <#
    .SYNOPSIS
        Collects Teams Phone configuration including voice routing, emergency calling,
        call queues, and auto attendants. Phone number enumeration is EXCLUDED.
    .DESCRIPTION
        Determines TeamsPhoneEnabled status for gating E911 assessment.
        Collects emergency calling policies, routing policies, and network sites.
    #>
    [CmdletBinding()] param()
    Write-AuditLog "  Collecting Teams Phone data (number enumeration excluded)..." -Level Info
    $result = @{
        TeamsPhoneEnabled = $false
        VoiceRoutingPolicies = @()
        EmergencyCallingPolicies = @()
        EmergencyCallRoutingPolicies = @()
        NetworkSites = @()
        CallQueues = @()
        AutoAttendants = @()
    }
    try {
        # Determine if Teams Phone is enabled by checking for voice routing or calling plan presence
        $voiceRoutes = Get-CsOnlineVoiceRoutingPolicy -ErrorAction SilentlyContinue
        $callingPolicies = Get-CsTeamsCallingPolicy -ErrorAction SilentlyContinue
        $result.TeamsPhoneEnabled = ($voiceRoutes -and $voiceRoutes.Count -gt 0) -or ($callingPolicies -and $callingPolicies.Count -gt 0)

        if ($result.TeamsPhoneEnabled) {
            Write-AuditLog "    Teams Phone is ENABLED - collecting phone configuration." -Level Info
            # Voice routing policies
            if ($voiceRoutes) { $result.VoiceRoutingPolicies = @($voiceRoutes) }
            # Emergency calling policies (critical for E911 assessment)
            $emergencyPolicies = Get-CsTeamsEmergencyCallingPolicy -ErrorAction SilentlyContinue
            if ($emergencyPolicies) { $result.EmergencyCallingPolicies = @($emergencyPolicies) }
            # Emergency call routing policies
            $emergencyRouting = Get-CsTeamsEmergencyCallRoutingPolicy -ErrorAction SilentlyContinue
            if ($emergencyRouting) { $result.EmergencyCallRoutingPolicies = @($emergencyRouting) }
            # Network sites for dynamic E911
            $netSites = Get-CsTenantNetworkSite -ErrorAction SilentlyContinue
            if ($netSites) { $result.NetworkSites = @($netSites) }
            # Call queues and auto attendants
            $cqs = Get-CsCallQueue -ErrorAction SilentlyContinue
            if ($cqs) { $result.CallQueues = @($cqs) }
            $aas = Get-CsAutoAttendant -ErrorAction SilentlyContinue
            if ($aas) { $result.AutoAttendants = @($aas) }

            Write-AuditLog "    E911 Policies: $($result.EmergencyCallingPolicies.Count) calling, $($result.EmergencyCallRoutingPolicies.Count) routing, $($result.NetworkSites.Count) network sites." -Level Info
        }
        else { Write-AuditLog "    Teams Phone is NOT enabled - E911 assessment will be skipped." -Level Info }
    }
    catch { Write-AuditLog "    Phone data error: $($_.Exception.Message)" -Level Warning }
    return $result
}

function Get-TeamsGovernanceAuditData {
    <#
    .SYNOPSIS
        Collects governance configuration including external access, guest access, and sensitivity labels.
    #>
    [CmdletBinding()] param()
    Write-AuditLog "  Collecting governance settings..." -Level Info
    $result = @{ ExternalAccess = @{}; GuestAccess = @{} }
    try {
        # External access
        $ext = Get-CsExternalAccessPolicy -ErrorAction SilentlyContinue | Where-Object { $_.Identity -eq "Global" }
        if ($ext) {
            $result.ExternalAccess = @{
                AllowTeamsConsumer = $ext.EnableTeamsConsumerAccess
                AllowTeamsConsumerInbound = $ext.EnableTeamsConsumerInbound
                AllowPublicUsers = $ext.EnablePublicCloudAccess
                AllowFederatedUsers = $ext.EnableFederationAccess
            }
        }
        # Guest access
        $guest = Get-CsTeamsGuestCallingConfiguration -ErrorAction SilentlyContinue
        $guestMeeting = Get-CsTeamsGuestMeetingConfiguration -ErrorAction SilentlyContinue
        $guestMsg = Get-CsTeamsGuestMessagingConfiguration -ErrorAction SilentlyContinue
        $result.GuestAccess = @{ AllowGuestAccess = $true }
        Write-AuditLog "    Governance data collected." -Level Success
    }
    catch { Write-AuditLog "    Governance error: $($_.Exception.Message)" -Level Warning }
    return $result
}

function Get-TeamsLifecycleAuditData {
    <#
    .SYNOPSIS
        Collects lifecycle management data including group expiration and retention policies.
    #>
    [CmdletBinding()] param()
    Write-AuditLog "  Collecting lifecycle data..." -Level Info
    $result = @{ ExpirationPolicy = @{} }
    try {
        $expPolicy = Get-MgGroupLifecyclePolicy -ErrorAction SilentlyContinue
        if ($expPolicy) {
            $result.ExpirationPolicy = @{
                GroupLifetimeInDays = $expPolicy.GroupLifetimeInDays
                ManagedGroupTypes = $expPolicy.ManagedGroupTypes
            }
        }
        Write-AuditLog "    Lifecycle data collected." -Level Success
    }
    catch { Write-AuditLog "    Lifecycle error: $($_.Exception.Message)" -Level Warning }
    return $result
}

function Get-TeamsSecurityComplianceAuditData {
    <#
    .SYNOPSIS
        Collects security and compliance data including PIM role assignments for Teams admin roles.
    .DESCRIPTION
        Retrieves PIM role assignments and determines whether each is Permanent (Active)
        or Eligible. Permanent assignments are flagged for JIT remediation.
    #>
    [CmdletBinding()] param()
    Write-AuditLog "  Collecting security and compliance data..." -Level Info
    $result = @{ PIMRoleAssignments = @(); ConditionalAccessPolicies = @() }
    try {
        # PIM Role Assignments for Teams-related admin roles
        Write-AuditLog "    Querying PIM role assignments..." -Level Info
        $teamsAdminRoles = @("Teams Administrator","Teams Communications Administrator","Teams Communications Support Engineer","Teams Communications Support Specialist","Teams Devices Administrator")
        $directoryRoles = Get-MgDirectoryRole -All -ErrorAction SilentlyContinue

        foreach ($roleName in $teamsAdminRoles) {
            $role = $directoryRoles | Where-Object { $_.DisplayName -eq $roleName }
            if ($role) {
                # Get active (permanent) assignments
                $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -ErrorAction SilentlyContinue
                foreach ($member in $members) {
                    $result.PIMRoleAssignments += @{
                        RoleName = $roleName
                        PrincipalName = $member.AdditionalProperties.displayName
                        PrincipalId = $member.Id
                        AssignmentType = "Permanent"
                        Status = "Active"
                    }
                }
            }
        }

        # Attempt to retrieve PIM eligible assignments via Graph
        try {
            foreach ($roleName in $teamsAdminRoles) {
                $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition -Filter "displayName eq '$roleName'" -ErrorAction SilentlyContinue
                if ($roleDefinitions) {
                    $eligible = Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance -Filter "roleDefinitionId eq '$($roleDefinitions.Id)'" -ErrorAction SilentlyContinue
                    foreach ($e in $eligible) {
                        $principal = Get-MgDirectoryObject -DirectoryObjectId $e.PrincipalId -ErrorAction SilentlyContinue
                        $result.PIMRoleAssignments += @{
                            RoleName = $roleName
                            PrincipalName = $principal.AdditionalProperties.displayName
                            PrincipalId = $e.PrincipalId
                            AssignmentType = "Eligible"
                            Status = "Inactive (JIT)"
                        }
                    }
                }
            }
        }
        catch { Write-AuditLog "    PIM eligible assignments query requires additional Graph permissions." -Level Warning }

        $permCount = @($result.PIMRoleAssignments | Where-Object { $_.AssignmentType -eq "Permanent" }).Count
        $eligCount = @($result.PIMRoleAssignments | Where-Object { $_.AssignmentType -eq "Eligible" }).Count
        Write-AuditLog "    PIM: $permCount permanent, $eligCount eligible assignments found." -Level Info
    }
    catch { Write-AuditLog "    Security data error: $($_.Exception.Message)" -Level Warning }
    return $result
}
#End Data Collection Functions

#Start Excel Report Generation
function New-TeamsAuditExcelReport {
    <#
    .SYNOPSIS
        Generates the comprehensive RAG-rated Excel workbook with all fixes applied.
    .DESCRIPTION
        FIXES IN THIS VERSION:
        1. RiskLevel column: Explicit [string] cast prevents System.Object[] serialisation
        2. Executive Summary: RAG rows have Green/Amber/Red background colour fills
        3. Risk Matrix: Rows colour-coded by RiskCategory severity
        4. App Catalog and Phone Numbers worksheets: REMOVED
        5. E911 Compliance: New assessment section when Teams Phone is enabled
        6. PIM Assignments: Permanent assignments flagged Red with JIT recommendation
        7. Policy Names: Included in all finding rows for traceability

        Microsoft brand colours applied per palette specification.
    .PARAMETER AuditData
        Hashtable containing all collected audit data from data collection functions.
    .PARAMETER BestPractices
        Hashtable containing the best practices reference from Get-BestPracticesReference.
    .PARAMETER OutputPath
        Directory path for report file output.
    .OUTPUTS
        Returns the full path to the generated Excel workbook.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$AuditData,
        [Parameter(Mandatory = $true)][hashtable]$BestPractices,
        [Parameter(Mandatory = $true)][string]$OutputPath
    )

    Write-AuditLog "Generating RAG-rated Excel report..." -Level Header
    $timestamp = $script:StartTime.ToString("yyyyMMdd_HHmmss")
    $reportPath = Join-Path $OutputPath "TeamsAudit_Report_$timestamp.xlsx"
    $jsonPath = Join-Path $OutputPath "TeamsAudit_ReportData_$timestamp.json"
    $allFindings = @()

    # Remove stale file from a previous failed run if present (prevents Save lock errors)
    if (Test-Path $reportPath) {
        try { Remove-Item $reportPath -Force -ErrorAction Stop; Write-AuditLog "  Removed stale report file." -Level Info }
        catch { Write-AuditLog "  WARNING: Could not remove existing file. Ensure it is not open in Excel: $reportPath" -Level Error; throw }
    }

    try {
        # ── Helper: Assess a policy domain against best practice reference ──
        # All output properties are explicitly typed to prevent serialisation issues
        function Get-PolicyDomainFindings {
            param(
                [string]$DomainName,
                [string]$CmdletEndpoint,
                $Policies,
                [hashtable]$BPRef,
                [switch]$IsSinglePolicy
            )
            $findings = @()
            $items = if ($IsSinglePolicy) { @($Policies) } else { $Policies }
            foreach ($pol in $items) {
                foreach ($setting in $BPRef.Keys) {
                    $hasKey = if ($IsSinglePolicy) { $Policies.ContainsKey($setting) } else { $pol.ContainsKey($setting) }
                    if ($hasKey) {
                        $cv = if ($IsSinglePolicy) { $Policies[$setting] } else { $pol[$setting] }
                        $b = $BPRef[$setting]
                        $r = Get-RAGRating -CurrentValue $cv -BestPracticeEntry $b
                        $pn = if ($IsSinglePolicy) { "Org-Wide" } else { [string]$pol.Identity }
                        # Every property is explicitly cast to prevent mixed-type arrays
                        $findings += [PSCustomObject]@{
                            Domain              = [string]$DomainName
                            'Cmdlet/Endpoint'   = [string]$CmdletEndpoint
                            PolicyName          = [string]$pn
                            Setting             = [string]$setting
                            CurrentValue        = [string]$cv
                            DefaultValue        = [string]$b.Default
                            BestPractice        = [string]$b.BestPractice
                            RAGStatus           = [string]$r.RAGStatus
                            ComplianceStatus    = [string]$r.ComplianceStatus
                            RiskImpact          = [int]$b.RiskImpact
                            RiskLikelihood      = [int]$b.RiskLikelihood
                            RiskScore           = [int]$r.RiskScore
                            RiskLevel           = [string]$r.RiskLevel
                            SecurityDomain      = [string]$b.SecurityDomain
                            RemediationTimeline = [string]$b.RemediationTimeline
                            RemediationEffort   = [string]$b.RemediationEffort
                            Explanation         = [string]$b.Explanation
                            Recommendation      = [string]$b.Recommendation
                        }
                    }
                }
                if ($IsSinglePolicy) { break }
            }
            return $findings
        }

        # ── Assess each configuration domain with source cmdlet/endpoint references ──
        Write-Progress -Activity "Assessing Configuration" -Status "Meeting Policies" -PercentComplete 10
        if ($AuditData.TeamsMeetings -and $AuditData.TeamsMeetings.MeetingPolicies) {
            $allFindings += @(Get-PolicyDomainFindings -DomainName "Meeting Policies" -CmdletEndpoint "Get-CsTeamsMeetingPolicy" -Policies $AuditData.TeamsMeetings.MeetingPolicies -BPRef $BestPractices.MeetingPolicies)
        }

        Write-Progress -Activity "Assessing Configuration" -Status "Messaging Policies" -PercentComplete 20
        if ($AuditData.TeamsMessaging -and $AuditData.TeamsMessaging.MessagingPolicies) {
            $allFindings += @(Get-PolicyDomainFindings -DomainName "Messaging Policies" -CmdletEndpoint "Get-CsTeamsMessagingPolicy" -Policies $AuditData.TeamsMessaging.MessagingPolicies -BPRef $BestPractices.MessagingPolicies)
        }

        Write-Progress -Activity "Assessing Configuration" -Status "External Access" -PercentComplete 30
        if ($AuditData.TeamsGovernance -and $AuditData.TeamsGovernance.ExternalAccess -and $AuditData.TeamsGovernance.ExternalAccess.Count -gt 0) {
            $allFindings += @(Get-PolicyDomainFindings -DomainName "External Access" -CmdletEndpoint "Get-CsExternalAccessPolicy" -Policies $AuditData.TeamsGovernance.ExternalAccess -BPRef $BestPractices.ExternalAccess -IsSinglePolicy)
        }

        Write-Progress -Activity "Assessing Configuration" -Status "Guest Access" -PercentComplete 35
        if ($AuditData.TeamsGovernance -and $AuditData.TeamsGovernance.GuestAccess -and $AuditData.TeamsGovernance.GuestAccess.Count -gt 0) {
            $allFindings += @(Get-PolicyDomainFindings -DomainName "Guest Access" -CmdletEndpoint "Get-CsTeamsGuestCallingConfiguration" -Policies $AuditData.TeamsGovernance.GuestAccess -BPRef $BestPractices.GuestAccess -IsSinglePolicy)
        }

        Write-Progress -Activity "Assessing Configuration" -Status "App Policies" -PercentComplete 40
        if ($AuditData.TeamsAppManagement -and $AuditData.TeamsAppManagement.SetupPolicies) {
            $allFindings += @(Get-PolicyDomainFindings -DomainName "App Policies" -CmdletEndpoint "Get-CsTeamsAppSetupPolicy" -Policies $AuditData.TeamsAppManagement.SetupPolicies -BPRef $BestPractices.AppPolicies)
        }

        Write-Progress -Activity "Assessing Configuration" -Status "Update Policies" -PercentComplete 50
        if ($AuditData.TeamsUpdatePolicies -and $AuditData.TeamsUpdatePolicies.UpdatePolicies) {
            $allFindings += @(Get-PolicyDomainFindings -DomainName "Update Policies" -CmdletEndpoint "Get-CsTeamsUpdateManagementPolicy" -Policies $AuditData.TeamsUpdatePolicies.UpdatePolicies -BPRef $BestPractices.UpdatePolicies)
        }

        Write-Progress -Activity "Assessing Configuration" -Status "Teams Settings" -PercentComplete 55
        if ($AuditData.TeamsSettings -and $AuditData.TeamsSettings.ClientConfiguration -and $AuditData.TeamsSettings.ClientConfiguration.Count -gt 0) {
            $allFindings += @(Get-PolicyDomainFindings -DomainName "Teams Settings" -CmdletEndpoint "Get-CsTeamsClientConfiguration" -Policies $AuditData.TeamsSettings.ClientConfiguration -BPRef $BestPractices.TeamsSettings -IsSinglePolicy)
        }

        Write-Progress -Activity "Assessing Configuration" -Status "Live Events" -PercentComplete 60
        if ($AuditData.TeamsMeetings -and $AuditData.TeamsMeetings.LiveEventPolicies) {
            $allFindings += @(Get-PolicyDomainFindings -DomainName "Live Events" -CmdletEndpoint "Get-CsTeamsMeetingBroadcastPolicy" -Policies $AuditData.TeamsMeetings.LiveEventPolicies -BPRef $BestPractices.LiveEventsPolicies)
        }

        # Encryption policies (special handling - different property names)
        Write-Progress -Activity "Assessing Configuration" -Status "Encryption" -PercentComplete 65
        if ($AuditData.TeamsEncryption -and $AuditData.TeamsEncryption.EncryptionPolicies) {
            foreach ($policy in $AuditData.TeamsEncryption.EncryptionPolicies) {
                $bp = $BestPractices.EnhancedEncryption.AllowEndToEndEncryption
                $r = Get-RAGRating -CurrentValue $policy.CallingEndToEndEncryptionEnabledType -BestPracticeEntry $bp
                $allFindings += [PSCustomObject]@{
                    Domain="Enhanced Encryption"; 'Cmdlet/Endpoint'="Get-CsTeamsEnhancedEncryptionPolicy"; PolicyName=[string]$policy.Identity; Setting="CallingEndToEndEncryptionEnabledType"
                    CurrentValue=[string]$policy.CallingEndToEndEncryptionEnabledType; DefaultValue=[string]$bp.Default; BestPractice=[string]$bp.BestPractice
                    RAGStatus=[string]$r.RAGStatus; ComplianceStatus=[string]$r.ComplianceStatus
                    RiskImpact=[int]$bp.RiskImpact; RiskLikelihood=[int]$bp.RiskLikelihood; RiskScore=[int]$r.RiskScore; RiskLevel=[string]$r.RiskLevel
                    SecurityDomain=[string]$bp.SecurityDomain; RemediationTimeline=[string]$bp.RemediationTimeline; RemediationEffort=[string]$bp.RemediationEffort
                    Explanation=[string]$bp.Explanation; Recommendation=[string]$bp.Recommendation
                }
            }
        }

        # ══════════════════════════════════════════════════════════════════════
        # E911 COMPLIANCE ASSESSMENT (always checked regardless of phone system)
        # Emergency calling configuration is a regulatory requirement for all tenants
        # ══════════════════════════════════════════════════════════════════════
        Write-Progress -Activity "Assessing Configuration" -Status "E911 Compliance" -PercentComplete 70
        Write-AuditLog "  Assessing E911 Compliance..." -Level Info

        # Determine current state of emergency calling configuration
        $hasEmergencyCallingPolicies = $false
        $hasEmergencyCallRouting = $false
        $hasNetworkSites = $false
        $hasNotificationGroups = $false
        $hasEmergencyAddresses = $false

        if ($AuditData.TeamsPhone) {
            $hasEmergencyCallingPolicies = ($null -ne $AuditData.TeamsPhone.EmergencyCallingPolicies -and $AuditData.TeamsPhone.EmergencyCallingPolicies.Count -gt 0)
            $hasEmergencyCallRouting = ($null -ne $AuditData.TeamsPhone.EmergencyCallRoutingPolicies -and $AuditData.TeamsPhone.EmergencyCallRoutingPolicies.Count -gt 0)
            $hasNetworkSites = ($null -ne $AuditData.TeamsPhone.NetworkSites -and $AuditData.TeamsPhone.NetworkSites.Count -gt 0)
            $hasNotificationGroups = ($null -ne $AuditData.TeamsPhone.EmergencyCallingPolicies -and
                ($AuditData.TeamsPhone.EmergencyCallingPolicies | Where-Object { -not [string]::IsNullOrEmpty($_.NotificationGroup) }).Count -gt 0)
            $hasEmergencyAddresses = ($null -ne $AuditData.TeamsPhone.EmergencyAddresses -and $AuditData.TeamsPhone.EmergencyAddresses.Count -gt 0)
        }

        $e911Checks = @(
            @{ Setting = "EmergencyCallingPolicyConfigured"; Value = $hasEmergencyCallingPolicies }
            @{ Setting = "EmergencyCallRoutingPolicyConfigured"; Value = $hasEmergencyCallRouting }
            @{ Setting = "NetworkSitesConfigured"; Value = $hasNetworkSites }
            @{ Setting = "NotificationGroupConfigured"; Value = $hasNotificationGroups }
        )

        foreach ($check in $e911Checks) {
            $bp = $BestPractices.EmergencyCalling[$check.Setting]
            if ($bp) {
                $r = Get-RAGRating -CurrentValue $check.Value -BestPracticeEntry $bp
                $allFindings += [PSCustomObject]@{
                    Domain="E911 Compliance"; 'Cmdlet/Endpoint'="Get-CsTeamsEmergencyCallingPolicy / Get-CsTeamsEmergencyCallRoutingPolicy / Get-CsTenantNetworkSite"; PolicyName="Org-Wide"; Setting=[string]$check.Setting
                    CurrentValue=[string]$check.Value; DefaultValue=[string]$bp.Default; BestPractice=[string]$bp.BestPractice
                    RAGStatus=[string]$r.RAGStatus; ComplianceStatus=[string]$r.ComplianceStatus
                    RiskImpact=[int]$bp.RiskImpact; RiskLikelihood=[int]$bp.RiskLikelihood; RiskScore=[int]$r.RiskScore; RiskLevel=[string]$r.RiskLevel
                    SecurityDomain=[string]$bp.SecurityDomain; RemediationTimeline=[string]$bp.RemediationTimeline; RemediationEffort=[string]$bp.RemediationEffort
                    Explanation=[string]$bp.Explanation; Recommendation=[string]$bp.Recommendation
                }
            }
        }
        Write-AuditLog "    E911 assessment complete: $($e911Checks.Count) checks performed." -Level Success

        # ══════════════════════════════════════════════════════════════════════
        # PIM PERMANENT ASSIGNMENT ASSESSMENT
        # All permanent admin assignments are flagged as RED with JIT recommendation
        # ══════════════════════════════════════════════════════════════════════
        Write-Progress -Activity "Assessing Configuration" -Status "PIM Assignments" -PercentComplete 80
        Write-AuditLog "  Assessing PIM Role Assignments..." -Level Info
        if ($AuditData.TeamsSecurity -and $AuditData.TeamsSecurity.PIMRoleAssignments -and $AuditData.TeamsSecurity.PIMRoleAssignments.Count -gt 0) {
            $bp = $BestPractices.PIMAssignments.PermanentTeamsAdminAssignment
            foreach ($assignment in $AuditData.TeamsSecurity.PIMRoleAssignments) {
                if ($assignment.AssignmentType -eq "Permanent") {
                    # ALWAYS flag permanent assignments as Red - standing privileges are non-compliant
                    $allFindings += [PSCustomObject]@{
                        Domain              = "PIM Role Assignments"
                        'Cmdlet/Endpoint'   = "Get-MgDirectoryRoleMember / Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance"
                        PolicyName          = [string]$assignment.RoleName
                        Setting             = "PermanentAssignment: $([string]$assignment.PrincipalName)"
                        CurrentValue        = "Permanent (Active)"
                        DefaultValue        = [string]$bp.Default
                        BestPractice        = [string]$bp.BestPractice
                        RAGStatus           = "Red"
                        ComplianceStatus    = "Non-Compliant"
                        RiskImpact          = [int]$bp.RiskImpact
                        RiskLikelihood      = [int]$bp.RiskLikelihood
                        RiskScore           = [int]($bp.RiskImpact * $bp.RiskLikelihood)
                        RiskLevel           = "Critical"
                        SecurityDomain      = [string]$bp.SecurityDomain
                        RemediationTimeline = [string]$bp.RemediationTimeline
                        RemediationEffort   = [string]$bp.RemediationEffort
                        Explanation         = [string]$bp.Explanation
                        Recommendation      = "Convert permanent assignment for $([string]$assignment.PrincipalName) ($([string]$assignment.RoleName)) to PIM Eligible. $([string]$bp.Recommendation)"
                    }
                }
            }
            $permCount = @($AuditData.TeamsSecurity.PIMRoleAssignments | Where-Object { $_.AssignmentType -eq "Permanent" }).Count
            Write-AuditLog "    PIM: $permCount permanent assignments flagged as Critical/Red." -Level $(if ($permCount -gt 0) { "Warning" } else { "Success" })
        }
        else { Write-AuditLog "    No PIM assignments found or insufficient permissions." -Level Info }

        # ── Build remediation roadmap from all assessed findings ──
        Write-Progress -Activity "Assessing Configuration" -Status "Building Roadmap" -PercentComplete 85
        $roadmap = Get-RemediationRoadmap -AllFindings $allFindings

        # ══════════════════════════════════════════════════════════════════════
        # WORKSHEET 1: Executive Summary with RAG COLOUR CODING (FIX #2)
        # ══════════════════════════════════════════════════════════════════════
        Write-Progress -Activity "Creating Excel Worksheets" -Status "Executive Summary" -PercentComplete 88
        Write-AuditLog "  Creating Executive Summary worksheet..." -Level Info

        $summaryData = @(
            [PSCustomObject]@{ Category="Audit Information"; Item="Report Generated"; Value=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); RAGIndicator="" }
            [PSCustomObject]@{ Category="Audit Information"; Item="Author"; Value=$script:ScriptAuthor; RAGIndicator="" }
            [PSCustomObject]@{ Category="Audit Information"; Item="Script Version"; Value=$script:ScriptVersion; RAGIndicator="" }
            [PSCustomObject]@{ Category=""; Item=""; Value=""; RAGIndicator="" }
            [PSCustomObject]@{ Category="RAG Summary"; Item="Total Settings Assessed"; Value=$roadmap.Summary.TotalSettings; RAGIndicator="" }
            [PSCustomObject]@{ Category="RAG Summary"; Item="GREEN - Compliant"; Value=$roadmap.Summary.GreenCount; RAGIndicator="Green" }
            [PSCustomObject]@{ Category="RAG Summary"; Item="AMBER - Review Required"; Value=$roadmap.Summary.AmberCount; RAGIndicator="Amber" }
            [PSCustomObject]@{ Category="RAG Summary"; Item="RED - Action Required"; Value=$roadmap.Summary.RedCount; RAGIndicator="Red" }
            [PSCustomObject]@{ Category="RAG Summary"; Item="Overall Compliance (%)"; Value="$($roadmap.Summary.CompliancePercentage)%"; RAGIndicator="" }
            [PSCustomObject]@{ Category=""; Item=""; Value=""; RAGIndicator="" }
            [PSCustomObject]@{ Category="Remediation"; Item="Quick Wins (0-30 days)"; Value=$roadmap.Summary.QuickWinCount; RAGIndicator="" }
            [PSCustomObject]@{ Category="Remediation"; Item="Medium Term (31-90 days)"; Value=$roadmap.Summary.MediumTermCount; RAGIndicator="" }
            [PSCustomObject]@{ Category="Remediation"; Item="Strategic (6-12 months)"; Value=$roadmap.Summary.StrategicCount; RAGIndicator="" }
        )

        if ($AuditData.Teams -and $AuditData.Teams.Statistics) {
            $s = $AuditData.Teams.Statistics
            $summaryData += [PSCustomObject]@{ Category=""; Item=""; Value=""; RAGIndicator="" }
            $summaryData += [PSCustomObject]@{ Category="Teams Overview"; Item="Total Teams"; Value=$s.TotalTeams; RAGIndicator="" }
            $summaryData += [PSCustomObject]@{ Category="Teams Overview"; Item="Active Teams"; Value=$s.ActiveTeams; RAGIndicator="" }
            $summaryData += [PSCustomObject]@{ Category="Teams Overview"; Item="Archived Teams"; Value=$s.ArchivedTeams; RAGIndicator="" }
        }

        # Export Executive Summary (formatting applied by post-export block)
        $summaryData | Export-Excel -Path $reportPath -WorksheetName "Executive Summary" `
            -AutoSize -TableName "ExecutiveSummary" -TableStyle None

        # ══════════════════════════════════════════════════════════════════════
        # WORKSHEET 2: RAG Assessment (All Findings) with colour coding
        # ══════════════════════════════════════════════════════════════════════
        Write-Progress -Activity "Creating Excel Worksheets" -Status "RAG Assessment" -PercentComplete 90
        Write-AuditLog "  Creating RAG Assessment worksheet..." -Level Info

        if ($allFindings.Count -gt 0) {
            # Export findings to RAG Assessment worksheet (formatting applied after all sheets are written)
            $allFindings |
                Select-Object Domain, 'Cmdlet/Endpoint', PolicyName, Setting, CurrentValue, DefaultValue, BestPractice, `
                             RAGStatus, ComplianceStatus, RiskImpact, RiskLikelihood, RiskScore, RiskLevel, `
                             SecurityDomain, RemediationTimeline, RemediationEffort, Recommendation |
                Sort-Object @{Expression="RAGStatus";Descending=$false}, @{Expression="RiskScore";Descending=$true} |
                Export-Excel -Path $reportPath -WorksheetName "RAG Assessment" -AutoSize `
                    -TableName "RAGAssessment" -TableStyle None -Append
        }

        # ══════════════════════════════════════════════════════════════════════
        # WORKSHEETS 3-5: Remediation Roadmap tiers with colour coding
        # ══════════════════════════════════════════════════════════════════════
        Write-Progress -Activity "Creating Excel Worksheets" -Status "Remediation Roadmap" -PercentComplete 92
        Write-AuditLog "  Creating remediation roadmap worksheets..." -Level Info

        $tierSelect = @(
            @{N='Priority';E={[int]$_.RiskScore}},
            'Domain','Cmdlet/Endpoint','PolicyName','Setting','RAGStatus','RiskLevel',
            'CurrentValue','BestPractice','RemediationEffort','SecurityDomain','Recommendation'
        )

        if ($roadmap.QuickWins.Count -gt 0) {
            $roadmap.QuickWins | Select-Object $tierSelect | Sort-Object Priority -Descending |
                Export-Excel -Path $reportPath -WorksheetName "Quick Wins (30 Days)" -AutoSize `
                    -TableName "QuickWins" -TableStyle None -Append
        }
        if ($roadmap.MediumTerm.Count -gt 0) {
            $roadmap.MediumTerm | Select-Object $tierSelect | Sort-Object Priority -Descending |
                Export-Excel -Path $reportPath -WorksheetName "Medium Term (90 Days)" -AutoSize `
                    -TableName "MediumTerm" -TableStyle None -Append
        }
        if ($roadmap.Strategic.Count -gt 0) {
            $roadmap.Strategic | Select-Object $tierSelect | Sort-Object Priority -Descending |
                Export-Excel -Path $reportPath -WorksheetName "Strategic (6-12 Months)" -AutoSize `
                    -TableName "Strategic" -TableStyle None -Append
        }

        # ══════════════════════════════════════════════════════════════════════
        # WORKSHEET 6: Risk Matrix with COLOUR CODING (FIX #3)
        # ══════════════════════════════════════════════════════════════════════
        Write-Progress -Activity "Creating Excel Worksheets" -Status "Risk Matrix" -PercentComplete 94
        Write-AuditLog "  Creating Risk Matrix worksheet..." -Level Info

        if ($allFindings.Count -gt 0) {
            $riskMatrix = $allFindings | Where-Object { $_.RAGStatus -ne "Green" } |
                Group-Object SecurityDomain | ForEach-Object {
                    $grp = $_.Group
                    $maxRisk = ($grp | Measure-Object -Property RiskScore -Maximum).Maximum
                    # FIX: Explicit [string] cast on RiskCategory to prevent System.Object[]
                    [string]$riskCat = if ($maxRisk -ge 12) { "Critical" }
                                       elseif ($maxRisk -ge 9) { "High" }
                                       elseif ($maxRisk -ge 4) { "Medium" }
                                       else { "Low" }
                    [PSCustomObject]@{
                        SecurityDomain   = [string]$_.Name
                        TotalFindings    = [int]$_.Count
                        RedFindings      = [int]@($grp | Where-Object { $_.RAGStatus -eq "Red" }).Count
                        AmberFindings    = [int]@($grp | Where-Object { $_.RAGStatus -eq "Amber" }).Count
                        HighestRiskScore = [int]$maxRisk
                        AverageRiskScore = [double][math]::Round(($grp | Measure-Object -Property RiskScore -Average).Average, 1)
                        RiskCategory     = [string]$riskCat
                        TopPriority      = [string]($grp | Sort-Object RiskScore -Descending | Select-Object -First 1).Setting
                    }
                } | Sort-Object HighestRiskScore -Descending

            if ($riskMatrix) {
                $riskMatrix | Export-Excel -Path $reportPath -WorksheetName "Risk Matrix" -AutoSize `
                    -TableName "RiskMatrix" -TableStyle None -Append
            }
        }

        # ══════════════════════════════════════════════════════════════════════
        # WORKSHEET 7: PIM Assignments (dedicated sheet with Red/Green flagging)
        # ══════════════════════════════════════════════════════════════════════
        Write-Progress -Activity "Creating Excel Worksheets" -Status "PIM Assignments" -PercentComplete 96
        if ($AuditData.TeamsSecurity -and $AuditData.TeamsSecurity.PIMRoleAssignments -and $AuditData.TeamsSecurity.PIMRoleAssignments.Count -gt 0) {
            Write-AuditLog "  Creating PIM Assignments worksheet..." -Level Info
            $pimData = foreach ($a in $AuditData.TeamsSecurity.PIMRoleAssignments) {
                [PSCustomObject]@{
                    RoleName       = [string]$a.RoleName
                    PrincipalName  = [string]$a.PrincipalName
                    AssignmentType = [string]$a.AssignmentType
                    Status         = [string]$a.Status
                    RAGStatus      = if ($a.AssignmentType -eq "Permanent") { "Red" } else { "Green" }
                    Recommendation = if ($a.AssignmentType -eq "Permanent") {
                        "CRITICAL: Convert to PIM Eligible. Permanent admin assignments maintain standing privileges without time limits, approval workflows, or MFA re-authentication. Use JIT via PIM."
                    } else { "Compliant - using PIM Eligible (JIT) assignment." }
                }
            }
            $pimData | Export-Excel -Path $reportPath -WorksheetName "PIM Assignments" -AutoSize `
                -TableName "PIMAssignments" -TableStyle None -Append
        }

        # ══════════════════════════════════════════════════════════════════════
        # WORKSHEET 8: Explainer - What is being checked and why
        # ══════════════════════════════════════════════════════════════════════
        Write-AuditLog "  Creating Explainer worksheet..." -Level Info
        $explainerData = @()

        # Build explainer rows from all best practice domains with cmdlet mapping
        $domainMap = @{
            TeamsSettings      = "Teams Settings"
            MeetingPolicies    = "Meeting Policies"
            MessagingPolicies  = "Messaging Policies"
            ExternalAccess     = "External Access"
            GuestAccess        = "Guest Access"
            AppPolicies        = "App Policies"
            EnhancedEncryption = "Enhanced Encryption"
            UpdatePolicies     = "Update Policies"
            LiveEventsPolicies = "Live Events"
            EmergencyCalling   = "E911 Compliance"
            PIMAssignments     = "PIM Role Assignments"
        }
        # Maps each best practice domain key to the PowerShell cmdlet or Graph endpoint used for collection
        $cmdletMap = @{
            TeamsSettings      = "Get-CsTeamsClientConfiguration"
            MeetingPolicies    = "Get-CsTeamsMeetingPolicy"
            MessagingPolicies  = "Get-CsTeamsMessagingPolicy"
            ExternalAccess     = "Get-CsExternalAccessPolicy"
            GuestAccess        = "Get-CsTeamsGuestCallingConfiguration"
            AppPolicies        = "Get-CsTeamsAppSetupPolicy / Get-CsTeamsAppPermissionPolicy"
            EnhancedEncryption = "Get-CsTeamsEnhancedEncryptionPolicy"
            UpdatePolicies     = "Get-CsTeamsUpdateManagementPolicy"
            LiveEventsPolicies = "Get-CsTeamsMeetingBroadcastPolicy"
            EmergencyCalling   = "Get-CsTeamsEmergencyCallingPolicy / Get-CsTeamsEmergencyCallRoutingPolicy / Get-CsTenantNetworkSite"
            PIMAssignments     = "Get-MgDirectoryRoleMember / Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance"
        }

        foreach ($domainKey in $BestPractices.Keys) {
            $domainLabel = if ($domainMap.ContainsKey($domainKey)) { $domainMap[$domainKey] } else { $domainKey }
            $cmdletLabel = if ($cmdletMap.ContainsKey($domainKey)) { $cmdletMap[$domainKey] } else { "N/A" }
            foreach ($settingKey in $BestPractices[$domainKey].Keys) {
                $entry = $BestPractices[$domainKey][$settingKey]
                $explainerData += [PSCustomObject]@{
                    AuditDomain         = [string]$domainLabel
                    'Cmdlet/Endpoint'   = [string]$cmdletLabel
                    SettingChecked      = [string]$settingKey
                    DefaultValue        = [string]$entry.Default
                    BestPracticeValue   = [string]$entry.BestPractice
                    WhyThisMatters      = [string]$entry.Explanation
                    Recommendation      = [string]$entry.Recommendation
                    SecurityDomain      = [string]$entry.SecurityDomain
                    RiskImpact          = [string]"$([int]$entry.RiskImpact) / 4"
                    RiskLikelihood      = [string]"$([int]$entry.RiskLikelihood) / 4"
                    MaxRiskScore        = [int]($entry.RiskImpact * $entry.RiskLikelihood)
                    RemediationWindow   = [string]$(switch ($entry.RemediationTimeline) { "QuickWin" {"0-30 Days"} "MediumTerm" {"31-90 Days"} "Strategic" {"6-12 Months"} default {$entry.RemediationTimeline} })
                    Effort              = [string]$entry.RemediationEffort
                }
            }
        }

        if ($explainerData.Count -gt 0) {
            $explainerData | Sort-Object AuditDomain, SettingChecked |
                Export-Excel -Path $reportPath -WorksheetName "Explainer" -AutoSize `
                    -TableName "Explainer" -TableStyle None -Append
        }

        Write-Progress -Activity "Creating Excel Worksheets" -Completed

        # ══════════════════════════════════════════════════════════════════════
        # POST-EXPORT: Apply formatting to ALL worksheets
        # 1. Blue (#00a1f1) header row with white text across all sheets
        # 2. Light grey (#F2F2F2) fill for all data rows across all sheets
        # 3. Column-specific RAG conditional formatting on RAG Assessment sheet
        # This must run AFTER all worksheets are written to avoid file lock errors.
        # ══════════════════════════════════════════════════════════════════════
        if (Test-Path $reportPath) {
            Write-AuditLog "  Applying workbook formatting (blue headers, grey rows, conditional colours)..." -Level Info
            try {
                $pkg = Open-ExcelPackage -Path $reportPath
                $blueHeader = [System.Drawing.Color]::FromArgb(0, 161, 241)   # Microsoft Blue #00a1f1
                $whiteFont  = [System.Drawing.Color]::White
                $lightGrey  = [System.Drawing.Color]::FromArgb(242, 242, 242)  # Light grey #F2F2F2

                foreach ($ws in $pkg.Workbook.Worksheets) {
                    if ($null -eq $ws.Dimension) { continue }
                    $colCount = $ws.Dimension.Columns
                    $rowCount = $ws.Dimension.Rows

                    # Apply blue background + white bold text to header row (row 1)
                    for ($c = 1; $c -le $colCount; $c++) {
                        $cell = $ws.Cells[1, $c]
                        $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                        $cell.Style.Fill.BackgroundColor.SetColor($blueHeader)
                        $cell.Style.Font.Color.SetColor($whiteFont)
                        $cell.Style.Font.Bold = $true
                    }

                    # Apply light grey background to all data rows (row 2 onward)
                    for ($r = 2; $r -le $rowCount; $r++) {
                        for ($c = 1; $c -le $colCount; $c++) {
                            $cell = $ws.Cells[$r, $c]
                            $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                            $cell.Style.Fill.BackgroundColor.SetColor($lightGrey)
                        }
                    }
                }

                # Apply RAG conditional formatting to specific columns on RAG Assessment sheet
                # Conditional formatting overrides the grey fill for matching cells
                $ragWs = $pkg.Workbook.Worksheets["RAG Assessment"]
                if ($ragWs -and $ragWs.Dimension) {
                    $ragRowCount = $ragWs.Dimension.Rows
                    # H = RAGStatus, M = RiskLevel, P = RemediationEffort
                    $targetCols = @("H", "M", "P")
                    foreach ($col in $targetCols) {
                        $range = "$($col)2:$($col)$ragRowCount"
                        Add-ConditionalFormatting -Worksheet $ragWs -Range $range -RuleType ContainsText -ConditionValue "Green" -BackgroundColor Green -ForegroundColor White
                        Add-ConditionalFormatting -Worksheet $ragWs -Range $range -RuleType ContainsText -ConditionValue "Amber" -BackgroundColor Orange -ForegroundColor Black
                        Add-ConditionalFormatting -Worksheet $ragWs -Range $range -RuleType ContainsText -ConditionValue "Red" -BackgroundColor Red -ForegroundColor White
                        Add-ConditionalFormatting -Worksheet $ragWs -Range $range -RuleType ContainsText -ConditionValue "Critical" -BackgroundColor Red -ForegroundColor White
                        Add-ConditionalFormatting -Worksheet $ragWs -Range $range -RuleType ContainsText -ConditionValue "High" -BackgroundColor OrangeRed -ForegroundColor White
                        Add-ConditionalFormatting -Worksheet $ragWs -Range $range -RuleType ContainsText -ConditionValue "Medium" -BackgroundColor Orange -ForegroundColor Black
                        Add-ConditionalFormatting -Worksheet $ragWs -Range $range -RuleType ContainsText -ConditionValue "Low" -BackgroundColor Green -ForegroundColor White
                    }
                }

                Close-ExcelPackage $pkg
                Write-AuditLog "  Workbook formatting applied." -Level Success
            }
            catch { Write-AuditLog "  Formatting warning: $($_.Exception.Message)" -Level Warning }
        }

        # Export JSON report data (for Word generation and external consumption)
        Write-AuditLog "  Exporting JSON report data..." -Level Info
        try {
            $jsonData = @{
                metadata = @{
                    title = "Microsoft Teams Tenant Compliance Audit Report"
                    author = $script:ScriptAuthor
                    version = $script:ScriptVersion
                    generatedDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    tenantName = if ($TenantId) { $TenantId } else { $env:USERDNSDOMAIN }
                }
                summary = $roadmap.Summary
                findings = @($allFindings | ForEach-Object {
                    @{
                        Domain=[string]$_.Domain; PolicyName=[string]$_.PolicyName; Setting=[string]$_.Setting
                        CurrentValue=[string]$_.CurrentValue; BestPractice=[string]$_.BestPractice
                        RAGStatus=[string]$_.RAGStatus; RiskLevel=[string]$_.RiskLevel
                        RiskScore=[int]$_.RiskScore; SecurityDomain=[string]$_.SecurityDomain
                        Recommendation=[string]$_.Recommendation
                    }
                })
                quickWins = @($roadmap.QuickWins | Select-Object Domain,PolicyName,Setting,RAGStatus,RiskLevel,CurrentValue,BestPractice,Recommendation)
                mediumTerm = @($roadmap.MediumTerm | Select-Object Domain,PolicyName,Setting,RAGStatus,RiskLevel,CurrentValue,BestPractice,Recommendation)
                strategic = @($roadmap.Strategic | Select-Object Domain,PolicyName,Setting,RAGStatus,RiskLevel,CurrentValue,BestPractice,Recommendation)
                appendix = @{
                    officialLinks = @(
                        "https://learn.microsoft.com/en-us/microsoftteams/"
                        "https://learn.microsoft.com/en-us/microsoftteams/meeting-policies-overview"
                        "https://learn.microsoft.com/en-us/microsoftteams/messaging-policies-in-teams"
                        "https://learn.microsoft.com/en-us/microsoftteams/manage-external-access"
                        "https://learn.microsoft.com/en-us/microsoftteams/guest-access"
                        "https://learn.microsoft.com/en-us/microsoftteams/app-policies"
                        "https://learn.microsoft.com/en-us/microsoftteams/teams-end-to-end-encryption"
                        "https://learn.microsoft.com/en-us/microsoftteams/teams-updates"
                        "https://learn.microsoft.com/en-us/microsoftteams/teams-live-events/plan-for-teams-live-events"
                        "https://learn.microsoft.com/en-us/microsoftteams/manage-emergency-calling-policies"
                        "https://learn.microsoft.com/en-us/microsoftteams/emergency-calling-dispatchable-location"
                        "https://learn.microsoft.com/en-us/microsoftteams/manage-emergency-call-routing-policies"
                        "https://learn.microsoft.com/en-us/entra/id-governance/privileged-identity-management/pim-configure"
                        "https://learn.microsoft.com/en-us/entra/id-governance/privileged-identity-management/pim-how-to-activate-role"
                        "https://learn.microsoft.com/en-us/graph/api/resources/teams-api-overview"
                        "https://learn.microsoft.com/en-us/microsoftteams/sensitivity-labels"
                        "https://learn.microsoft.com/en-us/microsoftteams/retention-policies"
                    )
                }
            }
            $jsonData | ConvertTo-Json -Depth 10 | Out-File $jsonPath -Encoding UTF8
            Write-AuditLog "  JSON exported: $jsonPath" -Level Success
        }
        catch { Write-AuditLog "  JSON export error: $($_.Exception.Message)" -Level Warning }

        Write-AuditLog "  Excel report generated: $reportPath" -Level Success
        # Return report path, findings, roadmap and JSON path for Word report generation
        return @{
            ReportPath  = $reportPath
            JsonPath    = $jsonPath
            AllFindings = $allFindings
            Roadmap     = $roadmap
        }
    }
    catch {
        Write-AuditLog "  Excel report generation failed: $($_.Exception.Message)" -Level Error
        Write-AuditLog "  Stack: $($_.ScriptStackTrace)" -Level Error
        throw
    }
}
#End Excel Report Generation

#Start Word Report Generation
function New-TeamsAuditWordReport {
    <#
    .SYNOPSIS
        Generates a Word document (.docx) compliance report with policy names in all findings.
    .DESCRIPTION
        Generates a Word document using PSWriteWord. JSON report data is exported
        by the Excel report function prior to this function being called.
        The report includes:
        - Cover page with tenant and audit metadata
        - Executive summary with RAG distribution
        - Remediation roadmap by tier
        - Detailed findings by domain with policy names
        - Appendix of official Microsoft reference links
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][array]$AllFindings,
        [Parameter(Mandatory=$true)][hashtable]$Roadmap,
        [Parameter(Mandatory=$true)][hashtable]$AuditData,
        [Parameter(Mandatory=$true)][string]$OutputPath
    )
    Write-AuditLog "Generating Word report..." -Level Header
    $timestamp = $script:StartTime.ToString("yyyyMMdd_HHmmss")
    $wordPath = Join-Path $OutputPath "TeamsAudit_Report_$timestamp.docx"
    $tenantName = if ($TenantId) { $TenantId } else { $env:USERDNSDOMAIN }

    try {
        Write-AuditLog "  Generating Word document..." -Level Info
        $doc = New-WordDocument $wordPath

        # Cover page with Microsoft brand colour headings
        Add-WordText -WordDocument $doc -Text "Microsoft Teams Tenant Compliance Audit Report" -HeadingType Heading1 -Bold $true -Color $script:MSColors.Blue -Supress $true
        Add-WordText -WordDocument $doc -Text "Prepared by: $($script:ScriptAuthor)" -HeadingType Heading2 -Supress $true
        Add-WordText -WordDocument $doc -Text "Date: $(Get-Date -Format 'yyyy-MM-dd')" -HeadingType Heading2 -Supress $true
        Add-WordText -WordDocument $doc -Text "Tenant: $tenantName" -HeadingType Heading2 -Supress $true
        Add-WordPageBreak -WordDocument $doc | Out-Null

        # Executive Summary section
        Add-WordText -WordDocument $doc -Text "Executive Summary" -HeadingType Heading1 -Bold $true -Color $script:MSColors.Blue -Supress $true
        Add-WordText -WordDocument $doc -Text "Total Settings Assessed: $($Roadmap.Summary.TotalSettings)" -Supress $true
        Add-WordText -WordDocument $doc -Text "Compliance Rate: $($Roadmap.Summary.CompliancePercentage)%" -Supress $true
        Add-WordText -WordDocument $doc -Text "Green (Compliant): $($Roadmap.Summary.GreenCount)" -Color $script:MSColors.Green -Supress $true
        Add-WordText -WordDocument $doc -Text "Amber (Review): $($Roadmap.Summary.AmberCount)" -Color $script:MSColors.Yellow -Supress $true
        Add-WordText -WordDocument $doc -Text "Red (Action Required): $($Roadmap.Summary.RedCount)" -Color $script:MSColors.Red -Supress $true
        Add-WordText -WordDocument $doc -Text "" -Supress $true
        Add-WordText -WordDocument $doc -Text "Remediation Summary:" -Bold $true -Supress $true
        Add-WordText -WordDocument $doc -Text "Quick Wins (0-30 days): $($Roadmap.Summary.QuickWinCount)" -Supress $true
        Add-WordText -WordDocument $doc -Text "Medium Term (31-90 days): $($Roadmap.Summary.MediumTermCount)" -Supress $true
        Add-WordText -WordDocument $doc -Text "Strategic (6-12 months): $($Roadmap.Summary.StrategicCount)" -Supress $true
        Add-WordPageBreak -WordDocument $doc | Out-Null

        # Remediation Roadmap - Quick Wins
        if ($Roadmap.QuickWins.Count -gt 0) {
            Add-WordText -WordDocument $doc -Text "Quick Wins (0-30 Days)" -HeadingType Heading1 -Bold $true -Color $script:MSColors.Green -Supress $true
            $qwTable = @($Roadmap.QuickWins | ForEach-Object {
                [PSCustomObject]@{
                    'Domain' = $_.Domain; 'Policy' = $_.PolicyName; 'Setting' = $_.Setting
                    'Current' = $_.CurrentValue; 'Best Practice' = $_.BestPractice
                    'RAG' = $_.RAGStatus; 'Recommendation' = $_.Recommendation
                }
            })
            Add-WordTable -WordDocument $doc -DataTable $qwTable -Design LightGridAccent1 -AutoFit Window | Out-Null
            Add-WordPageBreak -WordDocument $doc | Out-Null
        }

        # Remediation Roadmap - Medium Term
        if ($Roadmap.MediumTerm.Count -gt 0) {
            Add-WordText -WordDocument $doc -Text "Medium Term (31-90 Days)" -HeadingType Heading1 -Bold $true -Color $script:MSColors.Yellow -Supress $true
            $mtTable = @($Roadmap.MediumTerm | ForEach-Object {
                [PSCustomObject]@{
                    'Domain' = $_.Domain; 'Policy' = $_.PolicyName; 'Setting' = $_.Setting
                    'Current' = $_.CurrentValue; 'Best Practice' = $_.BestPractice
                    'RAG' = $_.RAGStatus; 'Recommendation' = $_.Recommendation
                }
            })
            Add-WordTable -WordDocument $doc -DataTable $mtTable -Design LightGridAccent1 -AutoFit Window | Out-Null
            Add-WordPageBreak -WordDocument $doc | Out-Null
        }

        # Remediation Roadmap - Strategic
        if ($Roadmap.Strategic.Count -gt 0) {
            Add-WordText -WordDocument $doc -Text "Strategic (6-12 Months)" -HeadingType Heading1 -Bold $true -Color $script:MSColors.Red -Supress $true
            $stTable = @($Roadmap.Strategic | ForEach-Object {
                [PSCustomObject]@{
                    'Domain' = $_.Domain; 'Policy' = $_.PolicyName; 'Setting' = $_.Setting
                    'Current' = $_.CurrentValue; 'Best Practice' = $_.BestPractice
                    'RAG' = $_.RAGStatus; 'Recommendation' = $_.Recommendation
                }
            })
            Add-WordTable -WordDocument $doc -DataTable $stTable -Design LightGridAccent1 -AutoFit Window | Out-Null
            Add-WordPageBreak -WordDocument $doc | Out-Null
        }

        # Detailed Findings by domain with policy names
        Add-WordText -WordDocument $doc -Text "Detailed Findings by Domain" -HeadingType Heading1 -Bold $true -Color $script:MSColors.Blue -Supress $true
        $domains = $AllFindings | Group-Object Domain
        foreach ($domain in $domains) {
            Add-WordText -WordDocument $doc -Text $domain.Name -HeadingType Heading2 -Bold $true -Supress $true
            $tableData = @($domain.Group | ForEach-Object {
                [PSCustomObject]@{
                    'Policy' = $_.PolicyName; 'Setting' = $_.Setting; 'Current' = $_.CurrentValue
                    'Best Practice' = $_.BestPractice; 'RAG' = $_.RAGStatus; 'Risk' = $_.RiskLevel
                    'Recommendation' = $_.Recommendation
                }
            })
            Add-WordTable -WordDocument $doc -DataTable $tableData -Design LightGridAccent1 -AutoFit Window | Out-Null
        }

        # Appendix - Official Microsoft Reference Links
        Add-WordPageBreak -WordDocument $doc | Out-Null
        Add-WordText -WordDocument $doc -Text "Appendix - Official Microsoft Reference Links" -HeadingType Heading1 -Bold $true -Color $script:MSColors.Blue -Supress $true
        Add-WordText -WordDocument $doc -Text "The following official Microsoft documentation was referenced during this audit:" -Italic $true -Supress $true
        $officialLinks = @(
            "https://learn.microsoft.com/en-us/microsoftteams/"
            "https://learn.microsoft.com/en-us/microsoftteams/meeting-policies-overview"
            "https://learn.microsoft.com/en-us/microsoftteams/messaging-policies-in-teams"
            "https://learn.microsoft.com/en-us/microsoftteams/manage-external-access"
            "https://learn.microsoft.com/en-us/microsoftteams/guest-access"
            "https://learn.microsoft.com/en-us/microsoftteams/app-policies"
            "https://learn.microsoft.com/en-us/microsoftteams/teams-end-to-end-encryption"
            "https://learn.microsoft.com/en-us/microsoftteams/teams-updates"
            "https://learn.microsoft.com/en-us/microsoftteams/teams-live-events/plan-for-teams-live-events"
            "https://learn.microsoft.com/en-us/microsoftteams/manage-emergency-calling-policies"
            "https://learn.microsoft.com/en-us/microsoftteams/emergency-calling-dispatchable-location"
            "https://learn.microsoft.com/en-us/microsoftteams/manage-emergency-call-routing-policies"
            "https://learn.microsoft.com/en-us/entra/id-governance/privileged-identity-management/pim-configure"
            "https://learn.microsoft.com/en-us/entra/id-governance/privileged-identity-management/pim-how-to-activate-role"
            "https://learn.microsoft.com/en-us/graph/api/resources/teams-api-overview"
        )
        foreach ($link in $officialLinks) {
            Add-WordText -WordDocument $doc -Text $link -Supress $true
        }

        Save-WordDocument $doc -Supress $true
        Write-AuditLog "  Word report generated: $wordPath" -Level Success

        return $wordPath
    }
    catch {
        Write-AuditLog "  Word report error: $($_.Exception.Message)" -Level Error
        return $null
    }
}
#End Word Report Generation

#Start Main Execution Function
function Invoke-TeamsComplianceAudit {
    <#
    .SYNOPSIS
        Main orchestration function for the Microsoft Teams compliance audit.
    .DESCRIPTION
        Coordinates all audit stages: prerequisite validation, authentication,
        data collection across all configuration domains, RAG assessment with
        E911 and PIM analysis, and multi-format report generation.
    #>
    [CmdletBinding()] param()

    # Display banner
    Write-Host ""
    Write-Host "  ================================================================" -ForegroundColor Cyan
    Write-Host "    Microsoft Teams Tenant Compliance Audit" -ForegroundColor Cyan
    Write-Host "    Version: $($script:ScriptVersion)" -ForegroundColor Cyan
    Write-Host "    Author:  $($script:ScriptAuthor)" -ForegroundColor Cyan
    Write-Host "  ================================================================" -ForegroundColor Cyan
    Write-Host ""

    # Initialise log file in output directory
    if (-not (Initialize-LogFile)) { Write-Host "Log initialisation failed." -ForegroundColor Red; return }

    Write-AuditLog "Starting Microsoft Teams Compliance Audit..." -Level Header
    Write-AuditLog "Output directory: $OutputPath" -Level Info

    # ── STEP 1: Validate prerequisites ──
    Write-AuditLog "" -Level Info
    Write-AuditLog "STEP 1/6: Validating Prerequisites" -Level Header
    if (-not (Test-RequiredModules)) {
        Write-AuditLog "Module validation failed. Install missing modules and re-run." -Level Error
        return
    }

    # ── STEP 2: Import modules ──
    Write-AuditLog "" -Level Info
    Write-AuditLog "STEP 2/6: Importing Modules" -Level Header
    if (-not (Import-RequiredModules)) {
        Write-AuditLog "Module import failed." -Level Error; return
    }

    # ── STEP 3: Connect to services ──
    Write-AuditLog "" -Level Info
    Write-AuditLog "STEP 3/6: Connecting to Microsoft Services" -Level Header
    if (-not (Connect-AuditServices)) {
        Write-AuditLog "Connection failed. Verify credentials and role assignments." -Level Error; return
    }

    # ── STEP 4: Collect audit data ──
    Write-AuditLog "" -Level Info
    Write-AuditLog "STEP 4/6: Collecting Audit Data" -Level Header

    $dataCollectors = @(
        @{ Name="Teams"; Function="Get-TeamsAuditData" }
        @{ Name="TeamsSettings"; Function="Get-TeamsSettingsAuditData" }
        @{ Name="TeamsPolicies"; Function="Get-TeamsPoliciesAuditData" }
        @{ Name="TeamsUpdatePolicies"; Function="Get-TeamsUpdatePoliciesAuditData" }
        @{ Name="TeamsUpgradeSettings"; Function="Get-TeamsUpgradeSettingsAuditData" }
        @{ Name="TeamsDevices"; Function="Get-TeamsDevicesAuditData" }
        @{ Name="TeamsAppManagement"; Function="Get-TeamsAppManagementAuditData" }
        @{ Name="TeamsMeetings"; Function="Get-TeamsMeetingsAuditData" }
        @{ Name="TeamsMessaging"; Function="Get-TeamsMessagingAuditData" }
        @{ Name="TeamsEncryption"; Function="Get-TeamsEncryptionAuditData" }
        @{ Name="TeamsPhone"; Function="Get-TeamsPhoneAuditData" }
        @{ Name="TeamsGovernance"; Function="Get-TeamsGovernanceAuditData" }
        @{ Name="TeamsLifecycle"; Function="Get-TeamsLifecycleAuditData" }
        @{ Name="TeamsSecurity"; Function="Get-TeamsSecurityComplianceAuditData" }
    )

    $total = $dataCollectors.Count; $current = 0
    foreach ($collector in $dataCollectors) {
        $current++
        $pct = [math]::Round(($current / $total) * 100)
        Write-Progress -Activity "Collecting Audit Data" -Status "[$current/$total] $($collector.Name)" -PercentComplete $pct
        try {
            $script:AuditResults[$collector.Name] = & $collector.Function
        }
        catch {
            Write-AuditLog "  Data collection failed for $($collector.Name): $($_.Exception.Message)" -Level Error
            $script:AuditResults[$collector.Name] = @{}
        }
    }
    Write-Progress -Activity "Collecting Audit Data" -Completed

    # ── STEP 5: Generate reports ──
    Write-AuditLog "" -Level Info
    Write-AuditLog "STEP 5/6: Generating Reports" -Level Header

    $bestPractices = Get-BestPracticesReference
    $excelPath = $null; $wordPath = $null
    $excelResult = $null

    # Generate Excel report (primary deliverable) - returns findings and roadmap for Word reuse
    try {
        $excelResult = New-TeamsAuditExcelReport -AuditData $script:AuditResults -BestPractices $bestPractices -OutputPath $OutputPath
        if ($excelResult) { $excelPath = $excelResult.ReportPath }
    }
    catch { Write-AuditLog "Excel report generation failed: $($_.Exception.Message)" -Level Error }

    # Generate Word report using the same findings and roadmap produced by the Excel report
    try {
        if ($excelResult -and $excelResult.AllFindings -and $excelResult.AllFindings.Count -gt 0) {
            $wordPath = New-TeamsAuditWordReport -AllFindings $excelResult.AllFindings -Roadmap $excelResult.Roadmap -AuditData $script:AuditResults -OutputPath $OutputPath
        }
        else {
            Write-AuditLog "Word report skipped: No findings data available (Excel report may have failed)." -Level Warning
        }
    }
    catch { Write-AuditLog "Word report generation failed: $($_.Exception.Message)" -Level Error }

    # ── STEP 6: Cleanup ──
    Write-AuditLog "" -Level Info
    Write-AuditLog "STEP 6/6: Cleanup" -Level Header
    Disconnect-AuditServices

    # ── Final summary ──
    $endTime = Get-Date
    $duration = $endTime - $script:StartTime

    Write-AuditLog "" -Level Info
    Write-AuditLog "================================================================" -Level Header
    Write-AuditLog "  AUDIT COMPLETE" -Level Header
    Write-AuditLog "================================================================" -Level Header
    Write-AuditLog "  Duration : $($duration.ToString('hh\:mm\:ss'))" -Level Info
    Write-AuditLog "  Errors   : $($script:ErrorCount)" -Level $(if ($script:ErrorCount -gt 0) { "Error" } else { "Success" })
    Write-AuditLog "  Warnings : $($script:WarningCount)" -Level $(if ($script:WarningCount -gt 0) { "Warning" } else { "Success" })
    Write-AuditLog "" -Level Info
    Write-AuditLog "  Output Files:" -Level Info
    Write-AuditLog "    Log File     : $($script:LogFile)" -Level Info
    if ($excelPath) { Write-AuditLog "    Excel Report : $excelPath" -Level Success }
    if ($wordPath)  { Write-AuditLog "    Word Report  : $wordPath" -Level Success }
    Write-AuditLog "================================================================" -Level Header

    return @{
        AuditData    = $script:AuditResults
        LogFile      = $script:LogFile
        ExcelReport  = $excelPath
        WordReport   = $wordPath
        Duration     = $duration
        ErrorCount   = $script:ErrorCount
        WarningCount = $script:WarningCount
    }
}
#End Main Execution Function

#Start Script Entry Point
# Execute the audit when the script is invoked
try {
    $results = Invoke-TeamsComplianceAudit
    if ($results) {
        Write-Host "`n  Audit completed. Review output files for detailed findings." -ForegroundColor Green
    }
}
catch {
    Write-Host "`n  Audit execution failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Check log file: $($script:LogFile)" -ForegroundColor Yellow
    try { Disconnect-AuditServices } catch {}
}
#End Script Entry Point