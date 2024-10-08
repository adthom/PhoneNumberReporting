@description('The name of the automation account')
param automationAccountName string = '${resourceGroup().name}automation'

@description('The SharePoint domain to use for the SharePoint Lists')
param SharePointDomain string

@description('The name of the SharePoint site to use for the SharePoint Lists')
param Site string = 'PhoneNumberManagement'

@description('The location of the runbooks')
param RunbookFileLocationBaseUri string = 'https://raw.githubusercontent.com/adthom/PhoneNumberReporting/main/Runbooks/'

@description('The location of the automation account, defaults to the resource group location')
param location string = resourceGroup().location

@description('The name of the DID <-> Department Map SharePoint List')
param DIDDepartmentMap string = 'DID-Department Map'

@description('The name of the RBAC SharePoint List')
param RBACList string = 'Report RBAC'

@description('The name of the Report SharePoint List')
param Report string = 'Phone Number Report'

@description('The environment of the SharePoint tenant')
@allowed(['Production', 'PPE', 'China', 'Germany', 'USGovernment', 'USGovernmentHigh', 'USGovernmentDoD'])
param Environment string = 'Production'

@description('The time to start the schedule in ISO 8601 format, defaults to 1 hour from now, must be more than 5 minutes from now')
param scheduleStart string = dateTimeAdd(utcNow(), 'PT1H')

var verboseEnabled = true
var progressEnabled = false
var traceLevel = 0
var powerShellVersion = 'PowerShell'

#disable-next-line no-hardcoded-env-urls // PSGallery Uri not available in all clouds
var PSGalleryUri = 'https://devopsgallerystorage.blob.core.windows.net/packages/'

var neededModules = [
    {
        name: 'MicrosoftTeams'
        version: '6.1.0'
    }
    {
        name: 'PnP.PowerShell'
        version: '1.12.0'
    }
]
var neededVariables = [
    {
        name: 'DIDDepartmentMapName'
        description: 'The name of the DID <-> Department Map SharePoint List'
        value: DIDDepartmentMap
    }
    {
        name: 'RBACListName'
        description: 'The name of the RBAC SharePoint List'
        value: RBACList
    }
    {
        name: 'ReportName'
        description: 'The name of the Report SharePoint List'
        value: Report
    }
    {
        name: 'SharePointDomain'
        description: 'The SharePoint domain to use for the SharePoint Lists'
        value: SharePointDomain
    }
    {
        name: 'Site'
        description: 'The name of the SharePoint site to use for the SharePoint Lists'
        value: Site
    }
    {
        name: 'RBACLastRun'
        description: 'The last time the RBAC runbook was run'
        value: dateTimeFromEpoch(0)
    }
    {
        name: 'Environment'
        description: 'The environment of the SharePoint tenant'
        value: Environment
    }
]
var neededAdhocRunbooks = [
    'CreateLists'
]
var neededScheduledRunbooks = [
    {
        name: 'UpdateRBAC'
        schedulename: 'RBAC Update Schedule'
        schedule: {
            startTime: scheduleStart
            interval: 4
            frequency: 'Hour'
        }
    }
    {
        name: 'UpdateDIDReportAssignments'
        schedulename: 'Daily Update Assignments Schedule'
        schedule: {
            startTime: scheduleStart
            interval: 1
            frequency: 'Day'
        }
    }
]

resource automationAccount 'Microsoft.Automation/automationAccounts@2022-08-08' = {
    name: automationAccountName
    location: location
    identity: {
        type: 'SystemAssigned'
    }
    properties: {
        publicNetworkAccess: true
        disableLocalAuth: false
        sku: {
            name: 'Basic'
        }
        encryption: {
            keySource: 'Microsoft.Automation'
        }
    }

    resource modules 'modules' = [for module in neededModules: {
        name: module.name
        properties: {
            contentLink: {
                uri: uri(PSGalleryUri, '${toLower(module.name)}.${module.version}.nupkg')
            }
        }
    }]

    resource variables 'variables' = [for variable in neededVariables:  {
        name: variable.name
        properties: {
            description: variable.description
            isEncrypted: false
            value: '"${variable.value}"'
        }
    }]

    resource schedules 'schedules' = [for runbook in neededScheduledRunbooks: {
        name: runbook.schedulename
        properties: runbook.schedule
    }]

    resource adhocRunbooks 'runbooks' = [for runbook in neededAdhocRunbooks: {
        name: runbook
        location: location
        properties: {
            runbookType: powerShellVersion
            logVerbose: verboseEnabled
            logProgress: progressEnabled
            logActivityTrace: traceLevel
            publishContentLink: {
                uri: uri(RunbookFileLocationBaseUri, '${runbook}.ps1')
            }
        }
        dependsOn: [ 
            variables
            modules
        ]
    }]

    resource scheduledRunbooks 'runbooks' = [for runbook in neededScheduledRunbooks: {
        name: runbook.name
        location: location
        properties: {
            runbookType: powerShellVersion
            logVerbose: verboseEnabled
            logProgress: progressEnabled
            logActivityTrace: traceLevel
            publishContentLink: {
                uri: uri(RunbookFileLocationBaseUri, '${runbook.name}.ps1')
            }
        }
        dependsOn: [ 
            schedules
            variables
            modules
        ]
    }]

    resource jobSchedules 'jobSchedules' = [for runbook in neededScheduledRunbooks: {
        name: guid(runbook.name)
        properties: {
            schedule: {
                name: runbook.schedulename
            }
            runbook: {
                name: runbook.name
            }
        }
        dependsOn: [ 
            scheduledRunbooks
            schedules
        ]
    }]
}

output automationAccountIdentity string = automationAccount.identity.principalId
