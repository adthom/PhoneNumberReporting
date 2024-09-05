#Requires -Version 7.2
#Requires -Modules @{ModuleName='PnP.PowerShell';RequiredVersion='2.2.0'}, @{ModuleName='Microsoft.Graph.Identity.Governance';RequiredVersion='2.10.0'}
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]
    $SharePointDomain,

    [Parameter(Mandatory)]
    [string]
    $Site,

    [Parameter(Mandatory)]
    [string]
    $AutomationAccountObjectId,

    [Parameter()]
    [string]
    $TeamsAdminRoleName = 'Teams Communications Administrator',

    [ValidateSet('Production', 'PPE', 'China', 'Germany', 'USGovernment', 'USGovernmentHigh', 'USGovernmentDoD')]
    [string]
    $Environment = 'Production',

    [Parameter()]
    [string]
    $Tenant
)

$ConnectMgGraphParams = @{
    Scopes      = @('RoleManagement.ReadWrite.Directory')
    NoWelcome   = $true
    ErrorAction = 'Stop'
}
$GraphEnvironmentName = switch ($Environment) {
    'China' { 'China'; break }
    'Germany' { 'Germany'; break }
    'USGovernmentHigh' { 'USGov'; break }
    'USGovernmentDoD' { 'USGovDoD'; break }
    default { $null }
}
if ($GraphEnvironmentName) {
    $ConnectMgGraphParams['Environment'] = $GraphEnvironmentName
}

Connect-MgGraph @ConnectMgGraphParams

$SharepointTLD = switch ($Environment) {
    'China' { 'sharepoint.cn'; break }
    'USGovernmentHigh' { 'sharepoint.us'; break }
    'USGovernmentDoD' { 'sharepoint-mil.us'; break }
    default { 'sharepoint.com' }
}

$SiteUrl = "https://${SharePointDomain}.${SharepointTLD}/sites/${Site}"

$ConnectPnPOnlineParams = @{
    Url              = $SiteUrl
    Interactive      = $true
    AzureEnvironment = $Environment
    ErrorAction      = 'Stop'
}
if ($Tenant) {
    $ConnectPnPOnlineParams['Tenant'] = $Tenant
}

$AutomationIdentity = @{
    ObjectId = $AutomationAccountObjectId
}

Connect-PnPOnline @ConnectPnPOnlineParams

$Role = @{
    AppRole     = 'Sites.Selected'
    BuiltInType = 'SharePointOnline'
}

$ServicePrincipal = Get-PnPAzureADServicePrincipal @AutomationIdentity -ErrorAction Stop
$RoleInfo = Get-PnPAzureADServicePrincipalAssignedAppRole -Principal $AutomationIdentity['ObjectId'] | Where-Object { $_.AppRoleName -eq $Role['AppRole'] }
if (!$RoleInfo -or $RoleInfo.Count -eq 0) {
    $RoleInfo = $ServicePrincipal | Add-PnPAzureADServicePrincipalAppRole @Role -ErrorAction Stop
}
$Permission = Grant-PnPAzureADAppSitePermission -Site $SiteUrl -Permissions Write -AppId $ServicePrincipal.AppId -DisplayName $ServicePrincipal.DisplayName -ErrorAction Stop
Set-PnPAzureADAppSitePermission -Site $SiteUrl -Permissions FullControl -PermissionId $Permission.Id -ErrorAction Stop

$TeamsRole = Get-MgRoleManagementDirectoryRoleDefinition -Filter "DisplayName eq '$TeamsAdminRoleName'" -ErrorAction Stop
$Existing = Get-MgRoleManagementDirectoryRoleAssignment -Filter "principalId eq '$AutomationAccountObjectId' and roleDefinitionId eq '$($TeamsRole.Id)'" -ErrorAction SilentlyContinue
if (!$Existing) {
    New-MgRoleManagementDirectoryRoleAssignment -PrincipalId $AutomationAccountObjectId -RoleDefinitionId $TeamsRole.Id -DirectoryScopeId '/' -ErrorAction Stop
}