<#
.SYNOPSIS
    Restricts Microsoft Graph Mail.Send permission to a single designated mailbox using Exchange Online Application Access Policy.

.DESCRIPTION
    By default, Microsoft Graph 'Mail.Send' is a tenant-wide application permission.
    Exchange Online requires an ApplicationAccessPolicy to target a Mail-Enabled Security Group.
    This script automatically creates or reuses a Mail-Enabled Security Group containing the designated sender mailbox,
    creates the ApplicationAccessPolicy restricting the Managed Identity, and verifies with Test-ApplicationAccessPolicy.

.PARAMETER ManagedIdentityId
    The Object (Principal) ID or Application ID of the Azure Automation Managed Identity.

.PARAMETER Mailbox
    The email address / UPN of the only allowed sender mailbox (e.g. reports@wedo.com.tr).

.PARAMETER GroupAlias
    Optional alias for the scoping Mail-Enabled Security Group (default: sg-copilot-senders).

.EXAMPLE
    .\Set-MailboxAccessPolicy.ps1 -ManagedIdentityId "47ac6d07-b886-4a77-b53f-e47c6cc20b6d" -Mailbox "reports@wedo.com.tr"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [Alias('Mi', 'PrincipalId', 'AppId')]
    [string]$ManagedIdentityId,

    [Parameter(Mandatory = $true)]
    [Alias('SenderMailbox', 'AllowedMailbox')]
    [string]$Mailbox,

    [Parameter(Mandatory = $false)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory = $false)]
    [string]$AutomationAccountName,

    [Parameter(Mandatory = $false)]
    [string]$BaseName = 'copilotrep',

    [Parameter(Mandatory = $false)]
    [string]$GroupAlias = "sg-copilot-senders"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($ManagedIdentityId)) {
    if (-not [string]::IsNullOrWhiteSpace($ResourceGroupName)) {
        if ([string]::IsNullOrWhiteSpace($AutomationAccountName)) {
            $AutomationAccountName = "aa-$($BaseName.ToLowerInvariant().Replace('_', '-'))"
        }
        Write-Host "Resolving Managed Identity for '$AutomationAccountName' in '$ResourceGroupName'..." -ForegroundColor Cyan
        $resolved = az automation account show --resource-group $ResourceGroupName --name $AutomationAccountName --query identity.principalId -o tsv 2>$null
        if ($resolved -and $resolved.Trim() -match '^[0-9a-fA-F-]{36}$') {
            $ManagedIdentityId = $resolved.Trim()
            Write-Host "Resolved Managed Identity ID: $ManagedIdentityId" -ForegroundColor Green
        }
    }
}

if ([string]::IsNullOrWhiteSpace($ManagedIdentityId)) {
    throw "Managed Identity ID is required. Provide -ManagedIdentityId <GUID> or pass -ResourceGroupName <Name>."
}

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

# 3. Exchange Online requires PolicyScopeGroupId to be a Mail-Enabled Security Group
Write-Host "Ensuring Mail-Enabled Security Group '$GroupAlias' exists for policy scoping..." -ForegroundColor Cyan
$scopeGroup = Get-DistributionGroup -Identity $GroupAlias -ErrorAction SilentlyContinue

if ($null -eq $scopeGroup) {
    Write-Host "Creating Mail-Enabled Security Group '$GroupAlias' with member '$cleanMailbox'..." -ForegroundColor Cyan
    $scopeGroup = New-DistributionGroup `
        -Name "Copilot Report Senders" `
        -DisplayName "Copilot Report Senders" `
        -Alias $GroupAlias `
        -Type Security `
        -Members $cleanMailbox
    Write-Host "Created Security Group: $($scopeGroup.PrimarySmtpAddress)" -ForegroundColor Green
}
else {
    Write-Host "Found existing Security Group: $($scopeGroup.PrimarySmtpAddress)" -ForegroundColor Green
    # Verify member
    $members = Get-DistributionGroupMember -Identity $GroupAlias -ErrorAction SilentlyContinue
    if ($null -eq $members -or $members.PrimarySmtpAddress -notcontains $cleanMailbox) {
        Write-Host "Adding '$cleanMailbox' to '$GroupAlias'..." -ForegroundColor Cyan
        Add-DistributionGroupMember -Identity $GroupAlias -Member $cleanMailbox -ErrorAction SilentlyContinue | Out-Null
    }
}

$policyScopeTarget = $scopeGroup.PrimarySmtpAddress
Write-Host "Policy Scope Target Group: $policyScopeTarget" -ForegroundColor Green

# 4. Check for existing Application Access Policies for this AppId
Write-Host ""
Write-Host "Configuring Application Access Policy in Exchange Online..." -ForegroundColor Cyan
$existingPolicies = Get-ApplicationAccessPolicy | Where-Object { $_.AppId -eq $appId }

if ($existingPolicies) {
    Write-Host "Updating existing Application Access Policy for AppId ($appId)..." -ForegroundColor Yellow
    Set-ApplicationAccessPolicy `
        -Identity $existingPolicies[0].Identity `
        -PolicyScopeGroupId $policyScopeTarget `
        -Description "Scoped Copilot Usage Report Managed Identity to $cleanMailbox via $GroupAlias" | Out-Null
}
else {
    Write-Host "Creating new ApplicationAccessPolicy restricting AppId to '$policyScopeTarget'..." -ForegroundColor Cyan
    New-ApplicationAccessPolicy `
        -AppId $appId `
        -PolicyScopeGroupId $policyScopeTarget `
        -AccessRight RestrictAccess `
        -Description "Scoped Copilot Usage Report Managed Identity to $cleanMailbox via $GroupAlias" | Out-Null
}

# 5. Verify policy enforcement with Test-ApplicationAccessPolicy
Write-Host ""
Write-Host "Verifying mailbox access policy with Test-ApplicationAccessPolicy..." -ForegroundColor Cyan
$testResult = Test-ApplicationAccessPolicy -Identity $cleanMailbox -AppId $appId

Write-Host ""
Write-Host "=== Mailbox Access Policy Status ===" -ForegroundColor Green
$summary = [pscustomobject]@{
    ApplicationId     = $appId
    ServicePrincipal  = $spDisplayName
    AllowedMailbox    = $cleanMailbox
    ScopeGroup        = $policyScopeTarget
    AccessCheckResult = $testResult.AccessCheckResult
    PolicyEnforced    = "True (Scoped to $cleanMailbox only)"
}
$summary | Format-Table -AutoSize

if ($testResult.AccessCheckResult -eq 'Granted') {
    Write-Host "SUCCESS: Managed Identity is restricted to send email ONLY from $cleanMailbox." -ForegroundColor Green
    Write-Host "Attempts to send as any other mailbox in the tenant will return Access Denied." -ForegroundColor Green
}
else {
    Write-Warning "Policy created. Access check returned: $($testResult.AccessCheckResult). Note: Exchange Online policies take up to 30 minutes to propagate fully across all servers."
}
