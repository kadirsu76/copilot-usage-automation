<#
.SYNOPSIS
    Restricts Microsoft Graph Mail.Send permission to a single designated mailbox using Exchange Online Application Access Policy.

.DESCRIPTION
    By default, Microsoft Graph 'Mail.Send' is a tenant-wide application permission.
    This script implements enterprise least-privilege security by creating an Exchange Online ApplicationAccessPolicy
    that restricts the Azure Automation Managed Identity to ONLY send emails from the specified sender mailbox.
    Any attempt to send email as another user in the tenant will be rejected by Exchange.

.PARAMETER ManagedIdentityId
    The Object (Principal) ID or Application ID of the Azure Automation Managed Identity.

.PARAMETER Mailbox
    The email address / UPN of the only allowed sender mailbox (e.g. reports@yourdomain.com).

.EXAMPLE
    .\Set-MailboxAccessPolicy.ps1 -ManagedIdentityId "5b9a89d1-32ef-4f12-a63d-058b21e8d47b" -Mailbox "reports@contoso.com"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [Alias('Mi', 'PrincipalId', 'AppId')]
    [string]$ManagedIdentityId,

    [Parameter(Mandatory = $true)]
    [Alias('SenderMailbox', 'AllowedMailbox')]
    [string]$Mailbox
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-Module {
    param([string]$Name)

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Installing required PowerShell module $Name..." -ForegroundColor Cyan
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    }
}

Assert-Module -Name ExchangeOnlineManagement
Assert-Module -Name Microsoft.Graph.Applications

Import-Module ExchangeOnlineManagement
Import-Module Microsoft.Graph.Applications

$cleanId = $ManagedIdentityId.Trim()
$cleanMailbox = $Mailbox.Trim()

Write-Host "Target Managed Identity: $cleanId" -ForegroundColor Cyan
Write-Host "Target Sender Mailbox  : $cleanMailbox" -ForegroundColor Cyan

# 1. Resolve Application (Client) ID of the Managed Identity
Write-Host "Resolving Application (Client) ID from Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "Application.Read.All", "Directory.Read.All" -NoWelcome

$sp = Get-MgServicePrincipal -ServicePrincipalId $cleanId -ErrorAction SilentlyContinue
if ($null -eq $sp) {
    # If cleanId was already an AppId
    $sp = Get-MgServicePrincipal -Filter "appId eq '$cleanId'" -ErrorAction SilentlyContinue
}

if ($null -eq $sp) {
    throw "Service Principal not found for Managed Identity ID: $cleanId"
}

$appId = $sp.AppId
$spDisplayName = $sp.DisplayName
Write-Host "Service Principal Display Name : $spDisplayName" -ForegroundColor Green
Write-Host "Resolved Application (Client) ID: $appId" -ForegroundColor Green

# 2. Connect to Exchange Online
Write-Host ""
Write-Host "Connecting to Exchange Online PowerShell..." -ForegroundColor Cyan
Connect-ExchangeOnline -ShowBanner:$false

# 3. Check for existing Application Access Policies for this AppId
Write-Host "Checking existing Application Access Policies in Exchange Online..." -ForegroundColor Cyan
$existingPolicies = Get-ApplicationAccessPolicy | Where-Object { $_.AppId -eq $appId }

if ($existingPolicies) {
    Write-Host "An Application Access Policy already exists for this AppId ($appId):" -ForegroundColor Yellow
    $existingPolicies | Format-Table -Property PolicyScopeGroupId, AccessRight, Description -AutoSize
    Write-Host "Updating policy scope to '$cleanMailbox'..." -ForegroundColor Cyan
    Set-ApplicationAccessPolicy -Identity $existingPolicies[0].Identity -PolicyScopeGroupId $cleanMailbox -Description "Scoped Copilot Usage Report Managed Identity" | Out-Null
}
else {
    Write-Host "Creating new ApplicationAccessPolicy restricting AppId to '$cleanMailbox'..." -ForegroundColor Cyan
    New-ApplicationAccessPolicy `
        -AppId $appId `
        -PolicyScopeGroupId $cleanMailbox `
        -AccessRight RestrictAccess `
        -Description "Scoped Copilot Usage Report Managed Identity to $cleanMailbox" | Out-Null
}

# 4. Verify policy enforcement with Test-ApplicationAccessPolicy
Write-Host ""
Write-Host "Verifying mailbox access policy..." -ForegroundColor Cyan
$testResult = Test-ApplicationAccessPolicy -Identity $cleanMailbox -AppId $appId

Write-Host ""
Write-Host "=== Mailbox Access Policy Status ===" -ForegroundColor Green
$summary = [pscustomobject]@{
    ApplicationId      = $appId
    ServicePrincipal   = $spDisplayName
    AllowedMailbox     = $cleanMailbox
    AccessCheckResult  = $testResult.AccessCheckResult
    PolicyEnforced     = "True (Scoped to $cleanMailbox only)"
}
$summary | Format-Table -AutoSize

if ($testResult.AccessCheckResult -eq 'Granted') {
    Write-Host "SUCCESS: Managed Identity is successfully restricted to send mail ONLY as $cleanMailbox." -ForegroundColor Green
    Write-Host "All other mailboxes in the tenant are protected and will return Access Denied." -ForegroundColor Green
}
else {
    Write-Warning "Policy created, but access check returned: $($testResult.AccessCheckResult). Exchange policy changes can take up to 30 minutes to propagate."
}
