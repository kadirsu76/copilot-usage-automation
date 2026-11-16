<#
.SYNOPSIS
    Deploys the Microsoft 365 Copilot Usage Automation infrastructure to Azure.

.DESCRIPTION
    Provisions the Resource Group, Automation Account with System-Assigned Managed Identity,
    Python 3 Runbook, Automation Variables, and weekly/daily Execution Schedule.

.PARAMETER ResourceGroupName
    Target Azure Resource Group name.

.PARAMETER Location
    Azure Region (e.g. westeurope, northeurope, eastus).

.PARAMETER BaseName
    Base naming prefix (e.g. copilotrep => aa-copilotrep, rb-copilotrep-report).

.PARAMETER ReportPeriod
    Reporting window (D7, D28, D90, D180). Default: D28.

.PARAMETER MailSender
    Sender mailbox address (UPN) for dispatching reports.

.PARAMETER MailRecipients
    Comma-separated list of recipient email addresses.

.PARAMETER ScheduleFrequency
    Report frequency: Week (Mondays at 08:00 UTC) or Day.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory = $false)]
    [string]$Location = "westeurope",

    [Parameter(Mandatory = $false)]
    [string]$BaseName = "copilotrep",

    [Parameter(Mandatory = $false)]
    [ValidateSet("D7", "D28", "D90", "D180")]
    [string]$ReportPeriod = "D28",

    [Parameter(Mandatory = $false)]
    [string]$MailSender = "",

    [Parameter(Mandatory = $false)]
    [string]$MailRecipients = "",

    [Parameter(Mandatory = $false)]
    [ValidateSet("Week", "Day")]
    [string]$ScheduleFrequency = "Week"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$PSNativeCommandUseErrorActionPreference = $true

function Assert-Command {
    param([string]$Name)

    if (-not (Get-Command -Name $Name -ErrorAction SilentlyContinue)) {
        throw "Required command '$Name' is not installed or not in PATH."
    }
}

Assert-Command -Name "az"

$root = Split-Path -Path $PSScriptRoot -Parent
$templateFile = Join-Path $root "infra/main.bicep"

if (-not (Test-Path -Path $templateFile)) {
    throw "Bicep template not found: $templateFile"
}

Write-Host "Checking Azure authentication state..." -ForegroundColor Cyan
try {
    az account show --output none | Out-Null
}
catch {
    Write-Host "Authenticating to Azure..." -ForegroundColor Cyan
    az login --output none | Out-Null
}

Write-Host "Ensuring Resource Group '$ResourceGroupName' exists in '$Location'..." -ForegroundColor Cyan
az group create --name $ResourceGroupName --location $Location --output none | Out-Null

$parameters = @(
    "baseName=$BaseName"
    "location=$Location"
    "reportPeriod=$ReportPeriod"
    "mailSender=$MailSender"
    "mailRecipients=$MailRecipients"
    "scheduleFrequency=$ScheduleFrequency"
)

Write-Host "Deploying Bicep infrastructure template..." -ForegroundColor Cyan
$deploymentResult = az deployment group create `
    --resource-group $ResourceGroupName `
    --template-file $templateFile `
    --parameters $parameters `
    --output json | ConvertFrom-Json

$outputs = $deploymentResult.properties.outputs
$miObjectId = $outputs.automationPrincipalId.value
$aaName = $outputs.automationAccountName.value
$rbName = $outputs.runbookName.value
$schName = $outputs.scheduleName.value

Write-Host ""
Write-Host "=== Deployment Succeeded! ===" -ForegroundColor Green
Write-Host "Automation Account: $aaName"
Write-Host "Managed Identity Principal ID: $miObjectId"
Write-Host "Published Runbook: $rbName"
Write-Host "Configured Schedule: $schName"
Write-Host ""
Write-Host "=== NEXT STEP: Grant Microsoft Graph Permissions ===" -ForegroundColor Yellow
Write-Host "Run the following command to grant Reports.Read.All and Mail.Send to the Managed Identity:"
Write-Host "pwsh ./scripts/Grant-GraphPermissions.ps1 -Mi $miObjectId" -ForegroundColor White
