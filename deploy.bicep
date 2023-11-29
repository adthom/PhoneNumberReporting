@description('The name of the automation account')
param automationAccountName string

@description('The location of the automation account, defaults to the resource group location')
param location string = resourceGroup().location

@description('The location of the runbooks')
param RunbookFileLocationBaseUri string = 'https://raw.githubusercontent.com/adthom/PhoneNumberReporting/main/Runbooks'

@description('The SharePoint domain to use for the SharePoint Lists')
param SharePointDomain string

@description('The name of the DID <-> Department Map SharePoint List')
param DIDDepartmentMap string = 'DID-Department Map'

@description('The name of the RBAC SharePoint List')
param RBACList string = 'Report RBAC'

@description('The name of the Report SharePoint List')
param Report string = 'Phone Number Report'

@description('The name of the SharePoint site to use for the SharePoint Lists')
param Site string = 'PhoneNumberManagement'

var verboseEnabled = true
var progressEnabled = false
var traceLevel = 0
var powerShellVersion = 'PowerShell7'

var PSGalleryUri = 'https://www.powershellgallery.com/api/v2/package'
var neededModules = [
    {
        name: 'MicrosoftTeams'
        version: '5.8.0'
    }
    {
        name: 'PnP.PowerShell'
        version: '2.2.0'
    }
    {
        name: 'Az'
        version: '11.0.0'
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
]
var neededAdhocRunbooks = [
    'CreateLists'
]
var neededScheduledRunbooks = [
    {
        name: 'UpdateRBAC'
        schedulename: 'RBAC Update Schedule'
        schedule: {
            startTime: dateTimeFromEpoch(0)
            interval: 4
            frequency: 'Hour'
        }
    }
    {
        name: 'UpdateDIDReportAssignments'
        schedulename: 'Daily Update Assignments Schedule'
        schedule: {
            startTime: dateTimeFromEpoch(0)
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
                uri: uri(PSGalleryUri, '${module.name}/${module.version}')
            }
        }
    }]

    resource variables 'variables' = [for variable in neededVariables:  {
        name: variable.name
        properties: {
            description: variable.description
            isEncrypted: false
            value: variable.value
        }
    }]

    resource schedules 'schedules' = [for runbook in neededScheduledRunbooks: {
        name: runbook.schedulename
        properties: runbook.schedule
    }]

    resource adhocRunbooks 'runbooks' = [for runbook in neededAdhocRunbooks: {
        name: runbook
        properties: {
            runbookType: powerShellVersion
            logVerbose: verboseEnabled
            logProgress: progressEnabled
            logActivityTrace: traceLevel
            draft: {
                inEdit: false
                draftContentLink: {
                    uri: uri(RunbookFileLocationBaseUri, '${runbook}.ps1')
                }
            }
        }

        dependsOn: [ 
            variables
            modules
        ]
    }]

    resource scheduledRunbooks 'runbooks' = [for runbook in neededScheduledRunbooks: {
        name: runbook.name
        properties: {
            runbookType: powerShellVersion
            logVerbose: verboseEnabled
            logProgress: progressEnabled
            logActivityTrace: traceLevel
            draft: {
                inEdit: false
                draftContentLink: {
                    uri: uri(RunbookFileLocationBaseUri, '${runbook.name}.ps1')
                }
            }
        }
        dependsOn: [ 
            schedules
            variables
            modules
        ]
    }]

    resource jobSchedules 'jobSchedules' = [for runbook in neededScheduledRunbooks: {
        name: runbook.name
        properties: {
            schedule: {
                name: runbook.schedulename
            }
            runbook: {
                name: runbook.name
            }
        }
    }]
}

output automationAccountIdentity string = automationAccount.identity.principalId
