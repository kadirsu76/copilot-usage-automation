<#
.SYNOPSIS
    Uploads the latest Python runbook code to Azure Automation and publishes it.

.PARAMETER ResourceGroupName
    Target Azure Resource Group.

.PARAMETER BaseName
    Base name used during deployment (default: copilotrep).

.PARAMETER SubscriptionId
    Azure Subscription ID (optional if already active in az cli).

.PARAMETER RunbookFile
    Path to the Python runbook file (default: ../runbook/Run-CopilotUsageReport.py).
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory = $false)]
    [string]$BaseName = 'copilotrep',

    [Parameter(Mandatory = $false)]
    [string]$SubscriptionId,

    [Parameter(Mandatory = $false)]
    [string]$RunbookFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true

function Assert-Command {
    param([Parameter(Mandatory = $true)][string]$Name)

    if (-not (Get-Command -Name $Name -ErrorAction SilentlyContinue)) {
        throw "Required command '$Name' is not installed or not in PATH."
    }
}

function Invoke-Az {
    param([Parameter(Mandatory = $true)][scriptblock]$Command)

    $output = & $Command
    if ($LASTEXITCODE -ne 0) {
        throw "Azure CLI command failed with exit code $LASTEXITCODE."
    }
    return $output
}

Assert-Command -Name 'az'

try {
    Invoke-Az { az account show --output none } | Out-Null
}
catch {
    Write-Host 'Authenticating to Azure...' -ForegroundColor Cyan
    Invoke-Az { az login --output none } | Out-Null
}

if (-not [string]::IsNullOrWhiteSpace($SubscriptionId)) {
    Invoke-Az { az account set --subscription $SubscriptionId } | Out-Null
}

$subscriptionIdValue = Invoke-Az { az account show --query id --output tsv }
if ([string]::IsNullOrWhiteSpace($subscriptionIdValue)) {
    throw 'Unable to resolve Azure subscription ID.'
}

$normalizedBaseName = $BaseName.ToLowerInvariant().Replace('_', '-')
$automationAccountName = "aa-$normalizedBaseName"
$runbookName = "rb-$normalizedBaseName-report"

if ([string]::IsNullOrWhiteSpace($RunbookFile)) {
    $RunbookFile = Join-Path $PSScriptRoot '..' 'runbook' 'Run-CopilotUsageReport.py'
}

$resolvedRunbookPath = [System.IO.Path]::GetFullPath($RunbookFile)
if (-not (Test-Path $resolvedRunbookPath)) {
    throw "Runbook file not found at: $resolvedRunbookPath"
}

Write-Host "Updating runbook '$runbookName' in '$automationAccountName'..." -ForegroundColor Cyan
Write-Host "Source file: $resolvedRunbookPath"

$draftUrl = "https://management.azure.com/subscriptions/${subscriptionIdValue}/resourceGroups/${ResourceGroupName}/providers/Microsoft.Automation/automationAccounts/${automationAccountName}/runbooks/${runbookName}/draft/content?api-version=2024-10-23"
$publishUrl = "https://management.azure.com/subscriptions/${subscriptionIdValue}/resourceGroups/${ResourceGroupName}/providers/Microsoft.Automation/automationAccounts/${automationAccountName}/runbooks/${runbookName}/publish?api-version=2024-10-23"
$runbookUrl = "https://management.azure.com/subscriptions/${subscriptionIdValue}/resourceGroups/${ResourceGroupName}/providers/Microsoft.Automation/automationAccounts/${automationAccountName}/runbooks/${runbookName}?api-version=2024-10-23"

# Upload draft content
Write-Host "Uploading draft content..." -ForegroundColor Cyan
Invoke-Az { az rest --method put --url $draftUrl --headers 'Content-Type=text/plain; charset=utf-8' --body "@$resolvedRunbookPath" } | Out-Null

# Publish runbook
Write-Host "Publishing runbook..." -ForegroundColor Cyan
Invoke-Az { az rest --method post --url $publishUrl } | Out-Null

# Poll until Published
Write-Host "Waiting for runbook to reach Published state..." -ForegroundColor Cyan
for ($i = 0; $i -lt 30; $i++) {
    Start-Sleep -Seconds 2
    $stateJson = Invoke-Az { az rest --method get --url $runbookUrl --output json }
    $state = $stateJson | ConvertFrom-Json
    if ($state.properties.state -eq 'Published') {
        Write-Host "Runbook '$runbookName' is now Published and ready." -ForegroundColor Green
        return
    }
}

Write-Host "Runbook publish request submitted. Verification timed out but publication is in progress." -ForegroundColor Yellow
