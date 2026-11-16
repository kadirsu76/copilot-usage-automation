targetScope = 'resourceGroup'

@description('Resource location')
param location string = resourceGroup().location

@description('Base name used to derive all resource names (e.g. copilotrep => aa-copilotrep, rb-copilotrep-report)')
@minLength(3)
@maxLength(20)
param baseName string = 'copilotrep'

@description('Copilot usage telemetry reporting window: D7, D28, D90, D180')
@allowed([
  'D7'
  'D28'
  'D90'
  'D180'
])
param reportPeriod string = 'D28'

@description('Sender mailbox address for report email delivery (leave blank for dry-run/preview)')
param mailSender string = ''

@description('Comma-separated list of recipient email addresses')
param mailRecipients string = ''

@description('Subject line of the dispatched email report')
param mailSubject string = 'Microsoft 365 Copilot Adoption & Usage Report'

@description('Execution frequency: Week (every Monday at 08:00 UTC) or Day')
@allowed([
  'Week'
  'Day'
])
param scheduleFrequency string = 'Week'

var automationContributorRoleId = subscriptionResourceId('Microsoft.Authorization/roleDefinitions', 'f353d9bd-d4a6-484e-a77a-8050b599b867')
var runbookContentBase64 = base64(loadTextContent('../runbook/Run-CopilotUsageReport.py'))
var publishRunbookScriptContent = loadTextContent('publish-runbook.ps1')
var publishRunbookScriptContentHash = base64(publishRunbookScriptContent)

var normalizedBaseName = toLower(replace(baseName, '_', '-'))
var automationAccountName = 'aa-${normalizedBaseName}'
var runbookName = 'rb-${normalizedBaseName}-report'
var scheduleName = 'sch-${normalizedBaseName}-${toLower(scheduleFrequency)}ly'
var deploymentIdentityName = 'mi-${normalizedBaseName}-ds'
var publishRunbookScriptName = 'ds-${normalizedBaseName}-publish-runbook'

// 1. Temporary Deployment Identity for publishing runbook and linking schedule
resource deploymentIdentity 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = {
  name: deploymentIdentityName
  location: location
}

// 2. Azure Automation Account with System-Assigned Managed Identity
resource automationAccount 'Microsoft.Automation/automationAccounts@2023-11-01' = {
  name: automationAccountName
  location: location
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    sku: {
      name: 'Basic'
    }
    publicNetworkAccess: true
  }
}

// 3. Automation Python 3 Runbook
resource runbook 'Microsoft.Automation/automationAccounts/runbooks@2024-10-23' = {
  name: runbookName
  parent: automationAccount
  location: location
  properties: {
    description: 'Generates Microsoft 365 Copilot adoption report with vector SVG trend and Excel attachment, dispatched via Graph API.'
    logActivityTrace: 0
    logProgress: true
    logVerbose: true
    runbookType: 'Python3'
    draft: {}
  }
}

// 4. Automation Variables
resource periodVariable 'Microsoft.Automation/automationAccounts/variables@2024-10-23' = {
  name: 'CopilotReport-Period'
  parent: automationAccount
  properties: {
    description: 'Copilot reporting period window (D7, D28, D90, D180).'
    isEncrypted: false
    value: '"${reportPeriod}"'
  }
}

resource mailSenderVariable 'Microsoft.Automation/automationAccounts/variables@2024-10-23' = {
  name: 'CopilotReport-MailSender'
  parent: automationAccount
  properties: {
    description: 'Sender mailbox address for report email delivery.'
    isEncrypted: false
    value: '"${mailSender}"'
  }
}

resource mailRecipientsVariable 'Microsoft.Automation/automationAccounts/variables@2024-10-23' = {
  name: 'CopilotReport-MailRecipients'
  parent: automationAccount
  properties: {
    description: 'Comma-separated recipient email addresses.'
    isEncrypted: false
    value: '"${mailRecipients}"'
  }
}

resource mailSubjectVariable 'Microsoft.Automation/automationAccounts/variables@2024-10-23' = {
  name: 'CopilotReport-MailSubject'
  parent: automationAccount
  properties: {
    description: 'Subject line of the email report.'
    isEncrypted: false
    value: '"${mailSubject}"'
  }
}

// 5. Grant Automation Contributor to Deployment Identity scoped to the Automation Account
resource automationRoleForDeploymentScript 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(automationAccount.id, deploymentIdentity.name, 'automationContributor')
  scope: automationAccount
  properties: {
    roleDefinitionId: automationContributorRoleId
    principalId: deploymentIdentity.properties.principalId
    principalType: 'ServicePrincipal'
  }
}

// 6. Deployment Script to publish Runbook and create JobSchedule
resource publishRunbookScript 'Microsoft.Resources/deploymentScripts@2023-08-01' = {
  name: publishRunbookScriptName
  location: location
  kind: 'AzurePowerShell'
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${deploymentIdentity.id}': {}
    }
  }
  properties: {
    azPowerShellVersion: '15.0'
    cleanupPreference: 'OnSuccess'
    retentionInterval: 'P1D'
    timeout: 'PT30M'
    forceUpdateTag: uniqueString(runbookContentBase64, publishRunbookScriptContentHash)
    environmentVariables: [
      {
        name: 'RG_NAME'
        value: resourceGroup().name
      }
      {
        name: 'AA_NAME'
        value: automationAccount.name
      }
      {
        name: 'RUNBOOK_NAME'
        value: runbook.name
      }
      {
        name: 'SCHEDULE_NAME'
        value: scheduleName
      }
      {
        name: 'SCHEDULE_FREQUENCY'
        value: scheduleFrequency
      }
      {
        name: 'AUTOMATION_API_VERSION'
        value: '2024-10-23'
      }
      {
        name: 'RUNBOOK_CONTENT_B64'
        value: runbookContentBase64
      }
    ]
    scriptContent: publishRunbookScriptContent
  }
  dependsOn: [
    automationRoleForDeploymentScript
  ]
}

output automationAccountResourceId string = automationAccount.id
output automationAccountName string = automationAccount.name
output automationPrincipalId string = automationAccount.identity.principalId
output runbookName string = runbook.name
output scheduleName string = scheduleName
output scheduleResourceId string = resourceId('Microsoft.Automation/automationAccounts/schedules', automationAccount.name, scheduleName)
