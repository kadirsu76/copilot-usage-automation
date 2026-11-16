$ErrorActionPreference = 'Stop'

Import-Module Az.Accounts -ErrorAction Stop

Connect-AzAccount -Identity | Out-Null
$subscriptionId = (Get-AzContext).Subscription.Id

$resourceGroupName = $env:RG_NAME
$automationAccountName = $env:AA_NAME
$runbookName = $env:RUNBOOK_NAME
$scheduleName = $env:SCHEDULE_NAME
$scheduleFrequency = if ($env:SCHEDULE_FREQUENCY) { $env:SCHEDULE_FREQUENCY } else { 'Week' }
$apiVersion = if ($env:AUTOMATION_API_VERSION) { $env:AUTOMATION_API_VERSION } else { '2024-10-23' }
$runbookContentBase64 = $env:RUNBOOK_CONTENT_B64
$runbookContent = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($runbookContentBase64))

function Assert-RequiredValue {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $false)][string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw "Required environment variable is empty: $Name"
    }

    return $Value.Trim()
}

$subscriptionId = Assert-RequiredValue -Name 'subscriptionId' -Value $subscriptionId
$resourceGroupName = Assert-RequiredValue -Name 'RG_NAME' -Value $resourceGroupName
$automationAccountName = Assert-RequiredValue -Name 'AA_NAME' -Value $automationAccountName
$runbookName = Assert-RequiredValue -Name 'RUNBOOK_NAME' -Value $runbookName
$scheduleName = Assert-RequiredValue -Name 'SCHEDULE_NAME' -Value $scheduleName

if ([string]::IsNullOrWhiteSpace($runbookContent)) {
    throw 'RUNBOOK_CONTENT_B64 is empty. Cannot continue.'
}

$baseUrl = [System.UriBuilder]::new('https://management.azure.com').Uri.AbsoluteUri.TrimEnd('/') +
    "/subscriptions/$([Uri]::EscapeDataString($subscriptionId))/resourceGroups/$([Uri]::EscapeDataString($resourceGroupName))/providers/Microsoft.Automation/automationAccounts/$([Uri]::EscapeDataString($automationAccountName))/runbooks/$([Uri]::EscapeDataString($runbookName))"
$automationBaseUrl = $baseUrl.Substring(0, $baseUrl.LastIndexOf('/runbooks/'))
$scheduleUrl = "$automationBaseUrl/schedules/$([Uri]::EscapeDataString($scheduleName))?api-version=$apiVersion"
$jobSchedulesUrl = "$automationBaseUrl/jobSchedules?api-version=$apiVersion"

function Get-PlainTextToken {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ResourceUrl
    )

    $tokenResponse = Get-AzAccessToken -ResourceUrl $ResourceUrl -ErrorAction Stop
    $tokenValue = $tokenResponse.Token

    if ($tokenValue -is [System.Security.SecureString]) {
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenValue)
        try {
            return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
        }
        finally {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }

    return [string]$tokenValue
}

$token = Get-PlainTextToken -ResourceUrl 'https://management.azure.com/'

$headers = @{
    Authorization = "Bearer $token"
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock,
        [int]$MaxAttempts = 12,
        [int]$DelaySeconds = 10
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return & $ScriptBlock
        }
        catch {
            if ($attempt -eq $MaxAttempts) {
                throw
            }
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

Write-Host "Updating draft content for runbook: $runbookName"
Invoke-WithRetry -ScriptBlock {
    Invoke-RestMethod -Method Put -Uri "${baseUrl}/draft/content?api-version=${apiVersion}" -Headers @{
        Authorization = "Bearer $token"
        'Content-Type' = 'text/plain; charset=utf-8'
    } -Body $runbookContent | Out-Null
} | Out-Null

Write-Host "Publishing runbook..."
Invoke-WithRetry -ScriptBlock {
    Invoke-RestMethod -Method Post -Uri "${baseUrl}/publish?api-version=${apiVersion}" -Headers $headers | Out-Null
} | Out-Null

$published = $false
for ($i = 0; $i -lt 60; $i++) {
    Start-Sleep -Seconds 5
    $state = Invoke-RestMethod -Method Get -Uri "${baseUrl}?api-version=${apiVersion}" -Headers $headers
    if ($state.properties.state -eq 'Published') {
        $published = $true
        Write-Host "Runbook is now in Published state."
        break
    }
}

if (-not $published) {
    throw 'Runbook publish did not reach Published state in allotted time.'
}

# Ensure Schedule exists
Write-Host "Configuring schedule: $scheduleName"
try {
    Invoke-RestMethod -Method Get -Uri $scheduleUrl -Headers $headers | Out-Null
    Write-Host "Schedule already exists."
}
catch {
    $startTime = (Get-Date).ToUniversalTime().Date.AddDays(1).AddHours(8).ToString('yyyy-MM-ddTHH:mm:ssZ')
    if ($scheduleFrequency -eq 'Week') {
        $scheduleBody = @{
            properties = @{
                description = 'Weekly Copilot Usage Report schedule on Monday at 08:00 UTC'
                startTime   = $startTime
                frequency   = 'Week'
                interval    = 1
                advancedSchedule = @{
                    weekDays = @('Monday')
                }
                timeZone    = 'Etc/UTC'
            }
        } | ConvertTo-Json -Depth 5 -Compress
    }
    else {
        $scheduleBody = @{
            properties = @{
                description = 'Daily Copilot Usage Report schedule at 08:00 UTC'
                startTime   = $startTime
                frequency   = 'Day'
                interval    = 1
                timeZone    = 'Etc/UTC'
            }
        } | ConvertTo-Json -Depth 5 -Compress
    }

    Invoke-RestMethod -Method Put -Uri $scheduleUrl -Headers @{
        Authorization  = "Bearer $token"
        'Content-Type' = 'application/json'
    } -Body $scheduleBody | Out-Null
    Write-Host "Schedule created successfully."
}

# Ensure Runbook is linked to Schedule via Job Schedule
Write-Host "Linking runbook to schedule..."
$jobSchedules = Invoke-RestMethod -Method Get -Uri $jobSchedulesUrl -Headers $headers
$existingLink = @($jobSchedules.value | Where-Object {
    $_.properties.runbook.name -eq $runbookName -and $_.properties.schedule.name -eq $scheduleName
} | Select-Object -First 1)

if ($existingLink.Count -eq 0) {
    $jobScheduleGuid = [guid]::NewGuid().ToString()
    $jobScheduleUrl = "$automationBaseUrl/jobSchedules/${jobScheduleGuid}?api-version=$apiVersion"
    $jobScheduleBody = @{
        properties = @{
            runbook  = @{ name = $runbookName }
            schedule = @{ name = $scheduleName }
        }
    } | ConvertTo-Json -Depth 5 -Compress

    Invoke-RestMethod -Method Put -Uri $jobScheduleUrl -Headers @{
        Authorization  = "Bearer $token"
        'Content-Type' = 'application/json'
    } -Body $jobScheduleBody | Out-Null
    Write-Host "Job schedule created: $jobScheduleGuid"
}
else {
    Write-Host "Runbook is already linked to schedule."
}

$DeploymentScriptOutputs = @{
    runbookState = 'Published'
    scheduleName = $scheduleName
}
