<#
.SYNOPSIS
    Test Azure AD Application Permissions for Teams Audit
    
.DESCRIPTION
    This script tests if an Azure AD application has the correct permissions
    and admin consent for running the Teams Audit script.
    
.PARAMETER TenantId
    The Microsoft 365 Tenant ID
    
.PARAMETER ClientId
    The Application (Client) ID
    
.PARAMETER ClientSecret
    The Client Secret
    
.EXAMPLE
    .\Test-AuditAppPermissions.ps1 -TenantId "contoso.onmicrosoft.com" -ClientId "app-id" -ClientSecret "secret"
    
.NOTES
    Script Name : Test-AuditAppPermissions.ps1
    Author      : Kelvin Chigorimbo
    Version     : 1.0
    Requires    : PowerShell 7.0 or later
#>

#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $true)]
    [string]$ClientId,
    
    [Parameter(Mandatory = $true)]
    [string]$ClientSecret
)

$RequiredPermissions = @(
    'Reports.Read.All'
    'Directory.Read.All'
    'Policy.Read.All'
    'Organization.Read.All'
    'TeamSettings.Read.All'
    'User.Read.All'
    'SecurityEvents.Read.All'
    'AuditLog.Read.All'
)

Write-Host "`n================================================================" -ForegroundColor Cyan
Write-Host "  Testing Azure AD Application Permissions" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Tenant ID:  " -NoNewline -ForegroundColor Gray
Write-Host $TenantId -ForegroundColor White
Write-Host "Client ID:  " -NoNewline -ForegroundColor Gray
Write-Host $ClientId -ForegroundColor White
Write-Host ""

# Test connection
Write-Host "Step 1: Testing authentication..." -ForegroundColor Yellow

try {
    $secureSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
    $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $secureSecret
    
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential -NoWelcome -ErrorAction Stop
    
    Write-Host "  ✓ Authentication successful" -ForegroundColor Green
    
} catch {
    Write-Host "  ✗ Authentication failed" -ForegroundColor Red
    Write-Host "`nError: $($_.Exception.Message)" -ForegroundColor Red
    
    if ($_.Exception.Message -match "AADSTS7000215") {
        Write-Host "`nLikely Issue: Invalid client secret" -ForegroundColor Yellow
        Write-Host "  • Verify you copied the secret value (not the secret ID)" -ForegroundColor Gray
        Write-Host "  • Create a new secret if the original was lost" -ForegroundColor Gray
    } elseif ($_.Exception.Message -match "AADSTS700016") {
        Write-Host "`nLikely Issue: Application not found" -ForegroundColor Yellow
        Write-Host "  • Verify the Client ID is correct" -ForegroundColor Gray
        Write-Host "  • Check that the app is in the correct tenant" -ForegroundColor Gray
    } elseif ($_.Exception.Message -match "forbidden" -or $_.Exception.Message -match "403") {
        Write-Host "`nLikely Issue: Insufficient permissions or missing admin consent" -ForegroundColor Yellow
        Write-Host "  • Verify admin consent was granted" -ForegroundColor Gray
        Write-Host "  • Check all permissions are 'Application' type" -ForegroundColor Gray
    }
    
    Write-Host ""
    exit 1
}

# Test permissions
Write-Host "`nStep 2: Verifying permissions..." -ForegroundColor Yellow

$context = Get-MgContext

if ($context) {
    Write-Host "  Connected as: $($context.Account)" -ForegroundColor Gray
    Write-Host "  Auth type: $($context.AuthType)" -ForegroundColor Gray
    
    # Test organization read
    Write-Host "`nStep 3: Testing Organization.Read.All..." -ForegroundColor Yellow
    try {
        $org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
        Write-Host "  ✓ Organization read successful" -ForegroundColor Green
        Write-Host "    Tenant: $($org.DisplayName)" -ForegroundColor Gray
    } catch {
        Write-Host "  ✗ Organization read failed" -ForegroundColor Red
        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Gray
    }
    
    # Test directory read
    Write-Host "`nStep 4: Testing Directory.Read.All..." -ForegroundColor Yellow
    try {
        $users = Get-MgUser -Top 1 -ErrorAction Stop
        Write-Host "  ✓ Directory read successful" -ForegroundColor Green
    } catch {
        Write-Host "  ✗ Directory read failed" -ForegroundColor Red
        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Gray
    }
    
    # Test policy read
    Write-Host "`nStep 5: Testing Policy.Read.All..." -ForegroundColor Yellow
    try {
        $policies = Get-MgPolicyAuthorizationPolicy -ErrorAction Stop
        Write-Host "  ✓ Policy read successful" -ForegroundColor Green
    } catch {
        Write-Host "  ✗ Policy read failed" -ForegroundColor Red
        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Gray
    }
    
    # Test reports read
    Write-Host "`nStep 6: Testing Reports.Read.All..." -ForegroundColor Yellow
    try {
        $report = Get-MgReportAuthenticationMethodUserRegistrationDetail -Top 1 -ErrorAction Stop
        Write-Host "  ✓ Reports read successful" -ForegroundColor Green
    } catch {
        Write-Host "  ✗ Reports read failed" -ForegroundColor Red
        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Gray
    }
    
    # Test audit log read
    Write-Host "`nStep 7: Testing AuditLog.Read.All..." -ForegroundColor Yellow
    try {
        $audit = Get-MgAuditLogDirectoryAudit -Top 1 -ErrorAction Stop
        Write-Host "  ✓ Audit log read successful" -ForegroundColor Green
    } catch {
        Write-Host "  ✗ Audit log read failed" -ForegroundColor Red
        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Gray
    }
    
    # Test security events read
    Write-Host "`nStep 8: Testing SecurityEvents.Read.All..." -ForegroundColor Yellow
    try {
        $secureScore = Get-MgSecuritySecureScore -Top 1 -ErrorAction SilentlyContinue
        if ($secureScore) {
            Write-Host "  ✓ Security events read successful" -ForegroundColor Green
        } else {
            Write-Host "  ⚠ Security events read returned no data (may be normal)" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "  ✗ Security events read failed" -ForegroundColor Red
        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Gray
    }
    
} else {
    Write-Host "  ✗ No context available" -ForegroundColor Red
}

# Disconnect
Write-Host "`nDisconnecting..." -ForegroundColor Yellow
Disconnect-MgGraph -ErrorAction SilentlyContinue

Write-Host "`n================================================================" -ForegroundColor Cyan
Write-Host "  Permission Test Complete" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan

Write-Host "`nIf all tests passed, the application is ready for Teams Audit!" -ForegroundColor Green
Write-Host "Run: .\Invoke-TeamsComprehensiveAudit.ps1 -TenantId '$TenantId' -ClientId '$ClientId' -ClientSecret 'your-secret'" -ForegroundColor Cyan
Write-Host ""
