#Requires -Modules Microsoft.Graph.Applications
[CmdletBinding()]
param(
    [Parameter()]
    [string]
    $DisplayName = 'PnP PowerShell for PhoneNumberReporting',

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
    Environment = $GraphEnvironmentName
    Scopes      = @('Application.ReadWrite.All')
    NoWelcome   = $true
    ErrorAction = 'Stop'
}
if ($TenantId) {
    $ConnectMgGraphParams['TenantId'] = $TenantId
}

Connect-MgGraph @ConnectMgGraphParams

$GraphContext = Get-MgContext
$TenantId = $GraphContext.TenantId
$GraphEnvironment = Get-MgEnvironment $GraphContext.Environment
$loginEndpoint = $GraphEnvironment.AzureADEndpoint

$NewAppParams = @{
    BodyParameter = @{
        isFallbackPublicClient = $true
        displayName            = $DisplayName
        signInAudience         = 'AzureADMyOrg'
        publicClient           = @{
            redirectUris = @(
                "${loginEndpoint}/common/oauth2/nativeclient"
                'http://localhost'
            )
        }
        requiredResourceAccess = @(
            @{
                resourceAppId  = '00000003-0000-0ff1-ce00-000000000000'
                resourceAccess = @(
                    @{
                        id   = '56680e0d-d2a3-4ae1-80d8-3c4f2100e3d0'
                        type = 'Scope'
                    }
                )
            }
        )
    }
    ErrorAction = 'Stop'
}
$app = New-MgApplication @NewAppParams

$app.PublicClient.RedirectUris += "ms-appx-web://microsoft.aad.brokerplugin/$($app.Id)"
$UpdateAppParams = @{
    ApplicationId = $app.Id
    BodyParameter = @{
        publicClient = $app.PublicClient
    }
    ErrorAction = 'Stop'
}
Update-MgApplication @UpdateAppParams

$consentUrl = "${loginEndpoint}/${TenantId}/v2.0/adminconsent?client_id=$($app.AppId)&scope=00000003-0000-0ff1-ce00-000000000000/.default&redirect_uri=http://localhost"
Write-Host "Please open the following URL in a browser to grant consent to the application:"
Write-Host $consentUrl

"App created. You can now connect to your tenant using Connect-PnPOnline -Url <yourtenanturl> -Interactive -ClientId $($app.AppId)"
[PSCustomObject]@{
    ClientId = $app.AppId
}