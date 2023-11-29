param(
    [Parameter(Mandatory)]
    [string]
    $SharePointDomain,

    [Parameter(Mandatory)]
    [string]
    $Site,

    [Parameter(Mandatory)]
    [string]
    $AutomationAccountObjectId
)

$Site = @{
    Site = "https://${SharePointDomain}.sharepoint.com/sites/${Site}"
}
$AutomationIdentity = @{
    ObjectId = $AutomationAccountObjectId
}

Connect-PnPOnline @Site -Interactive

$Role = @{
    AppRole = 'Sites.Selected'
    BuiltInType = 'SharePointOnline'
}

$ServicePrincipal = Get-PnPAzureADServicePrincipal @AutomationIdentity
$RoleInfo = $ServicePrincipal | Add-PnPAzureADServicePrincipalAppRole @Role
$Permission = Grant-PnPAzureADAppSitePermission @Site -Permissions Write -AppId $ServicePrincipal.AppId -DisplayName $ServicePrincipal.DisplayName
Set-PnPAzureADAppSitePermission @Site -Permissions FullControl -PermissionId $Permission.Id
