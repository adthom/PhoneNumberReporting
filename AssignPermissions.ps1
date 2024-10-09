#Requires -Version 7.2
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Applications, Microsoft.Graph.Identity.Governance, Microsoft.Graph.Sites
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
    $TenantId
)

$GraphEnvironmentName = switch ($Environment) {
    'China' { 'China'; break }
    'Germany' { 'Germany'; break }
    'USGovernmentHigh' { 'USGov'; break }
    'USGovernmentDoD' { 'USGovDoD'; break }
    default { 'Global' }
}

$ConnectMgGraphParams = @{
    Scopes      = @('RoleManagement.ReadWrite.Directory','Sites.FullControl.All','Application.Read.All')
    NoWelcome   = $true
    ErrorAction = 'Stop'
    Environment = $GraphEnvironmentName
}
if ($TenantId) {
    $ConnectMgGraphParams['TenantId'] = $TenantId
}

Connect-MgGraph @ConnectMgGraphParams

$RootSite = Get-MgSite -SiteId root -ErrorAction Stop

$SiteUri = [Uri]::new([Uri]$RootSite.WebUrl,"sites/$Site")

$SharePointRoles = @(
    'Sites.Selected'
)

$SharePointSPN = Get-MgServicePrincipalByAppId -AppId 00000003-0000-0ff1-ce00-000000000000 -ErrorAction Stop
$SharePointAppRoles = @($SharePointSPN | Select-Object -ExpandProperty AppRoles | Where-Object {$_.Value -in $SharePointRoles})
$ServicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $AutomationAccountObjectId -ExpandProperty AppRoleAssignments -ErrorAction Stop

$ExistingAssignments = @($ServicePrincipal.AppRoleAssignments | Where-Object { $_.AppRoleId -in $SharePointAppRoles.Id -and $_.ResourceId -eq $SharePointSPN.Id })
if ($ExistingAssignments.Count -lt $SharePointAppRoles.Count) {
    $SharePointAppRoles | Where-Object {
      $_.Id -notin $ExistingAssignments.AppRoleId
    } | ForEach-Object {
      $NewAppRoleAssignmentParams = @{
        ServicePrincipalId = $AutomationAccountObjectId
        BodyParameter = @{
          principalId = $AutomationAccountObjectId
          resourceId = $SharePointSPN.Id
          appRoleId = $_.Id
        }
        ErrorAction = 'Stop'
      }
      New-MgServicePrincipalAppRoleAssignment @NewAppRoleAssignmentParams
    }
}

$SharepointSite = Get-MgSite -SiteId ('{0}:{1}' -f $SiteUri.Host,$SiteUri.LocalPath) -ErrorAction Stop
$InvokeMgGraphParams = @{
    Method = 'POST'
    Uri = "v1.0/sites/$($SharepointSite.Id)/permissions"
    Body = @{
        roles = @( 'FullControl' )
        grantedToIdentities = @(
            @{
                application = @{
                    id = $ServicePrincipal.AppId
                    displayName = $ServicePrincipal.DisplayName
                }
            }
        )
        grantedToIdentitiesV2 = @(
            @{
                application = @{
                    id = $ServicePrincipal.AppId
                    displayName = $ServicePrincipal.DisplayName
                }
            }
        )
    } | ConvertTo-Json -Compress -Depth 3
    ContentType = 'application/json'
    ErrorAction = 'Stop'
    OutputType = 'PSObject'
}

$null = Invoke-MgGraphRequest @InvokeMgGraphParams

$TeamsRoleParams = @{
    Filter = "displayName eq '$TeamsAdminRoleName'"
    ErrorAction = 'Stop'
}
$TeamsRole = Get-MgRoleManagementDirectoryRoleDefinition @TeamsRoleParams
$ExistingRoleParams = @{
    Filter = "principalId eq '$AutomationAccountObjectId' and roleDefinitionId eq '$($TeamsRole.Id)'"
    ErrorAction = 'SilentlyContinue'
}
$Existing = Get-MgRoleManagementDirectoryRoleAssignment @ExistingRoleParams
if (!$Existing) {
    $NewRoleParams = @{
        PrincipalId = $AutomationAccountObjectId
        RoleDefinitionId = $TeamsRole.Id
        DirectoryScopeId = '/'
        ErrorAction = 'Stop'
    }
    New-MgRoleManagementDirectoryRoleAssignment @NewRoleParams
}
