<#
.SYNOPSIS
    Assigns required Microsoft Graph Application Permissions to the Azure Automation Managed Identity.

.DESCRIPTION
    Grants 'Reports.Read.All' (to read Copilot telemetry) and 'Mail.Send' (to send report emails with Excel attachments)
    to the System-Assigned or User-Assigned Managed Identity of the Azure Automation Account.

.PARAMETER Mi
    The Object (Principal) ID GUID of the Managed Identity.

.PARAMETER GraphAppPermissions
    Array of Microsoft Graph application permissions to assign. Defaults to Reports.Read.All and Mail.Send.

.EXAMPLE
    .\Grant-GraphPermissions.ps1 -Mi "5b9a89d1-32ef-4f12-a63d-058b21e8d47b"
#>

param(
    [Parameter(Mandatory = $false)]
    [Alias('ManagedIdentityObjectId', 'PrincipalId')]
    [string]$Mi,

    [Parameter(Mandatory = $false)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory = $false)]
    [string]$AutomationAccountName,

    [Parameter(Mandatory = $false)]
    [string]$BaseName = 'copilotrep',

    [Parameter(Mandatory = $false)]
    [string[]]$GraphAppPermissions = @(
        'Reports.Read.All',
        'Mail.Send'
    )
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($Mi)) {
    if (-not [string]::IsNullOrWhiteSpace($ResourceGroupName)) {
        if ([string]::IsNullOrWhiteSpace($AutomationAccountName)) {
            $AutomationAccountName = "aa-$($BaseName.ToLowerInvariant().Replace('_', '-'))"
        }
        Write-Host "Resolving Managed Identity for '$AutomationAccountName' in '$ResourceGroupName'..." -ForegroundColor Cyan
        $resolved = az automation account show --resource-group $ResourceGroupName --name $AutomationAccountName --query identity.principalId -o tsv 2>$null
        if ($resolved -and $resolved.Trim() -match '^[0-9a-fA-F-]{36}$') {
            $Mi = $resolved.Trim()
            Write-Host "Resolved Managed Identity Principal ID: $Mi" -ForegroundColor Green
        }
    }
}

if ([string]::IsNullOrWhiteSpace($Mi)) {
    throw "Managed Identity ID is required. Provide -Mi <PrincipalIdGuid> or pass -ResourceGroupName <Name>."
}

$managedIdentityObjectId = $Mi.Trim()
if ($managedIdentityObjectId -notmatch '^[0-9a-fA-F-]{36}$') {
    throw "Invalid -Mi value ($managedIdentityObjectId). Provide a valid 36-character Managed Identity Principal (Object) ID GUID."
}

function Assert-Module {
    param([string]$Name)

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Installing required PowerShell module $Name..." -ForegroundColor Cyan
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    }
}

Assert-Module -Name Microsoft.Graph.Authentication
Assert-Module -Name Microsoft.Graph.Applications

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Applications

Write-Host "Target Managed Identity Principal ID: $managedIdentityObjectId" -ForegroundColor Cyan

$requiredScopes = @(
    'Application.Read.All',
    'AppRoleAssignment.ReadWrite.All',
    'Directory.Read.All'
)

Write-Host "Connecting to Microsoft Graph with required admin scopes..." -ForegroundColor Cyan
Connect-MgGraph -Scopes $requiredScopes -NoWelcome

# Well-known AppId for Microsoft Graph service principal
$graphServicePrincipal = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
if ($null -eq $graphServicePrincipal) {
    throw 'Microsoft Graph service principal not found in tenant.'
}

$miServicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $managedIdentityObjectId -ErrorAction SilentlyContinue
if ($null -eq $miServicePrincipal) {
    throw "Managed Identity service principal not found for ID: $managedIdentityObjectId"
}

Write-Host "Target Service Principal Name: $($miServicePrincipal.DisplayName)" -ForegroundColor Green

$existingAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $managedIdentityObjectId -All

$results = @()

foreach ($permission in $GraphAppPermissions) {
    $appRole = $graphServicePrincipal.AppRoles |
        Where-Object { $_.Value -eq $permission -and $_.AllowedMemberTypes -contains 'Application' -and $_.IsEnabled } |
        Select-Object -First 1

    if ($null -eq $appRole) {
        $results += [pscustomobject]@{
            Permission = $permission
            Status     = 'NotFound'
            AppRoleId  = ''
            Message    = 'App role not found on Microsoft Graph.'
        }
        continue
    }

    $alreadyAssigned = $existingAssignments |
        Where-Object {
            $_.ResourceId -eq $graphServicePrincipal.Id -and
            $_.AppRoleId -eq $appRole.Id
        } |
        Select-Object -First 1

    if ($null -ne $alreadyAssigned) {
        $results += [pscustomobject]@{
            Permission = $permission
            Status     = 'AlreadyAssigned'
            AppRoleId  = $appRole.Id
            Message    = 'Permission is already assigned.'
        }
        continue
    }

    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $managedIdentityObjectId -BodyParameter @{
        principalId = $managedIdentityObjectId
        resourceId  = $graphServicePrincipal.Id
        appRoleId   = $appRole.Id
    } | Out-Null

    $results += [pscustomobject]@{
        Permission = $permission
        Status     = 'Assigned'
        AppRoleId  = $appRole.Id
        Message    = 'Permission assigned successfully.'
    }
}

Write-Host ''
Write-Host 'Microsoft Graph Permission Assignment Summary:' -ForegroundColor Cyan
$results | Format-Table -AutoSize

Write-Host ''
Write-Host 'Note: Managed Identity tokens in Azure are cached. New permissions may take a few minutes to propagate.' -ForegroundColor Yellow
