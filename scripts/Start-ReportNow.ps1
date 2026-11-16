<#
.SYNOPSIS
    Triggers the Copilot Usage Report Runbook on-demand in Azure Automation.

.PARAMETER ResourceGroupName
    Target Azure Resource Group.

.PARAMETER BaseName
    Base name used during deployment (default: copilotrep).

.PARAMETER Wait
    Wait for the job to complete and stream progress.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory = $false)]
    [string]$BaseName = 'copilotrep',

    [Parameter(Mandatory = $false)]
    [string]$SubscriptionId,

    [Parameter(Mandatory = $false)]
    [switch]$Wait,

    [Parameter(Mandatory = $false)]
    [switch]$SyncRunbook,

    [Parameter(Mandatory = $false)]
    [ValidateRange(5, 300)]
    [int]$PollSeconds = 10,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 120)]
    [int]$TimeoutMinutes = 30
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

if ($SyncRunbook.IsPresent) {
    $updateScript = Join-Path $PSScriptRoot 'Update-Runbook.ps1'
    if (Test-Path $updateScript) {
        Write-Host "Syncing runbook code before triggering job..." -ForegroundColor Cyan
        & $updateScript -ResourceGroupName $ResourceGroupName -BaseName $BaseName -SubscriptionId $subscriptionIdValue
    }
}

$jobName = [guid]::NewGuid().ToString()
$jobUrl = "https://management.azure.com/subscriptions/${subscriptionIdValue}/resourceGroups/${ResourceGroupName}/providers/Microsoft.Automation/automationAccounts/${automationAccountName}/jobs/${jobName}?api-version=2024-10-23"
$jobBody = '{"properties":{"runbook":{"name":"' + $runbookName + '"}}}'

Write-Host "Triggering runbook '$runbookName' in '$automationAccountName'..." -ForegroundColor Cyan
$raw = Invoke-Az { az rest --method put --url $jobUrl --headers 'Content-Type=application/json' --body $jobBody --output json }
$null = $raw | ConvertFrom-Json

Write-Host "Report Job Started Successfully." -ForegroundColor Green
Write-Host "Subscription      : $subscriptionIdValue"
Write-Host "Resource Group    : $ResourceGroupName"
Write-Host "Automation Account: $automationAccountName"
Write-Host "Runbook           : $runbookName"
Write-Host "Job ID            : $jobName"

if (-not $Wait.IsPresent) {
    Write-Host "Job submitted in background. Pass -Wait to stream status." -ForegroundColor Yellow
    return
}

$terminalStatuses = @('Completed', 'Failed', 'Stopped', 'Suspended', 'Blocked')
$lastStatus = ''
$deadline = (Get-Date).ToUniversalTime().AddMinutes($TimeoutMinutes)

while ($true) {
    if ((Get-Date).ToUniversalTime() -ge $deadline) {
        throw "Timed out waiting for job after $TimeoutMinutes minutes."
    }

    Start-Sleep -Seconds $PollSeconds
    $raw = Invoke-Az { az rest --method get --url $jobUrl --output json }
    $job = $raw | ConvertFrom-Json
    $status = [string]$job.properties.status

    if ($status -ne $lastStatus) {
        Write-Host "Current job status: $status" -ForegroundColor Cyan
        $lastStatus = $status
    }

    if ($terminalStatuses -contains $status) {
        if ($status -ne 'Completed') {
            $exceptionText = [string]$job.properties.exception
            throw "Job finished with status '$status'. Exception: $exceptionText"
        }
        Write-Host "Report generation and email dispatch completed successfully." -ForegroundColor Green
        break
    }
}
