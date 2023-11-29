#Requires -Modules PnP.PowerShell, Microsoft.Graph.Identity.RoleManagement
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
    $TeamsAdminRoleName = 'Teams Communications Administrator'
)

$SiteUrl = "https://${SharePointDomain}.sharepoint.com/sites/${Site}"
$AutomationIdentity = @{
    ObjectId = $AutomationAccountObjectId
}

Connect-PnPOnline -Url $SiteUrl -Interactive -ErrorAction Stop
Connect-MgGraph -Scopes 'RoleManagement.ReadWrite.Directory' -NoWelcome -ErrorAction Stop

$Role = @{
    AppRole = 'Sites.Selected'
    BuiltInType = 'SharePointOnline'
}

$ServicePrincipal = Get-PnPAzureADServicePrincipal @AutomationIdentity -ErrorAction Stop
$RoleInfo = $ServicePrincipal | Add-PnPAzureADServicePrincipalAppRole @Role -ErrorAction Stop
$Permission = Grant-PnPAzureADAppSitePermission -Site $SiteUrl -Permissions Write -AppId $ServicePrincipal.AppId -DisplayName $ServicePrincipal.DisplayName -ErrorAction Stop
Set-PnPAzureADAppSitePermission -Site $SiteUrl -Permissions FullControl -PermissionId $Permission.Id -ErrorAction Stop

$TeamsRole = Get-MgRoleManagementDirectoryRoleDefinition -Filter "DisplayName eq '$TeamsAdminRoleName'" -ErrorAction Stop
$Existing = Get-MgRoleManagementDirectoryRoleAssignment -Filter "principalId eq '$AutomationAccountObjectId' and roleDefinitionId eq '$($TeamsRole.Id)'" -ErrorAction SilentlyContinue
if (!$Existing) {
    New-MgRoleManagementDirectoryRoleAssignment -PrincipalId $AutomationAccountObjectId -RoleDefinitionId $TeamsRole.Id -DirectoryScopeId '/' -ErrorAction Stop
}