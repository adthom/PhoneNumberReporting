#Requires -Version 7.2
#Requires -Modules @{ModuleName='PnP.PowerShell';RequiredVersion='2.2.0'}
[CmdletBinding()]
param(
    [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
    [string]
    $Department,

    [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
    [ValidateRange(10000000000, 19999999999)]
    [double]
    $DID,

    [Parameter(Mandatory, HelpMessage='The subdomain of the SharePoint site that contains the report (e.g. use "contoso" for "contoso.sharepoint.com")')]
    [string]
    $SharePointDomain,

    [Parameter(Mandatory, HelpMessage='The id of the SharePoint site that contains the report')]
    [string]
    $Site,

    [Parameter()]
    [string]
    $DIDDepartmentMapName = 'DID-Department Map',

    [Parameter()]
    [string]
    $RBACListName = 'Report RBAC',

    [Parameter()]
    [string]
    $ReportName = 'Phone Number Report',

    [Parameter()]
    [switch]
    $ClearExisting,

    [ValidateSet('Production', 'PPE', 'China', 'Germany', 'USGovernment', 'USGovernmentHigh', 'USGovernmentDoD')]
    [string]
    $Environment = 'Production',

    [Parameter(Mandatory)]
    [string]
    $PnPClientId,

    [Parameter()]
    [string]
    $PnPTenant
)

begin {
    $SharepointTLD = switch ($Environment) {
        'China' { 'sharepoint.cn'; break }
        'USGovernmentHigh' { 'sharepoint.us'; break }
        'USGovernmentDoD' { 'sharepoint-mil.us'; break }
        default { 'sharepoint.com' }
    }
    
    $ConnectPnPOnlineParams = @{
        Url              = "https://${SharePointDomain}.${SharepointTLD}/sites/${Site}"
        Interactive      = $true
        AzureEnvironment = $Environment
        ErrorAction      = 'Stop'
        ClientId         = $PnPClientId
    }
    if ($PnPTenant) {
        $ConnectPnPOnlineParams['Tenant'] = $PnPTenant
    }

    Connect-PnPOnline @ConnectPnPOnlineParams

    $DIDList = Get-PnPList -Identity $DIDDepartmentMapName
    $RBACList = Get-PnPList -Identity $RBACListName
    $ReportList = Get-PnPList -Identity $ReportName

    $DIDs = [Collections.Generic.List[PSCustomObject]]@()
}
process {
    $DIDs.Add([PSCustomObject]@{
            Department = $Department
            DID        = $DID
        })
}
end {
    if ($ClearExisting) {
        Write-Information "Clearing existing data"        
        $batch = New-PnPBatch
        $RBACList | Get-PnPListItem | Remove-PnPListItem -List $RBACList -Batch $batch
        $DIDList | Get-PnPListItem | Remove-PnPListItem -List $DIDList -Batch $batch
        $ReportList | Get-PnPListItem | Remove-PnPListItem -List $ReportList -Batch $batch
        Invoke-PnPBatch -Batch $batch -ErrorAction Stop
        Write-Information "Existing data cleared"
    }

    Write-Information "Adding Departments to RBAC List"
    $batch = New-PnPBatch
    $DIDs.Department | Sort-Object -Unique | ForEach-Object { Add-PnPListItem -List $RBACList -Values @{Department = $_ } -Batch $batch }
    Invoke-PnPBatch -Batch $batch
    if ($batch.RequestCount -gt 0) {
        Write-Warning "Batch Failed"
        return
    }
    Write-Information "Departments added to RBAC List"

    Write-Information "Adding items to DID <-> Department Map"
    $batch = New-PnPBatch
    $DIDs | ForEach-Object { Add-PnPListItem -List $DIDList -Values @{Department = $_.Department; DID = $_.DID } -Batch $batch }
    Invoke-PnPBatch -Batch $batch
    if ($batch.RequestCount -gt 0) {
        Write-Warning "Batch Failed"
        return
    }
    Write-Information "All items added to DID <-> Department Map"
    
    $DIDLookupList = @{}
    Get-PnPListItem -List $DIDList -Fields Id, DID -PageSize 1000 | ForEach-Object { $DIDLookupList[$_.FieldValues['DID']] = $_.Id }
    
    Write-Information "Adding items to Report List"
    $batch = New-PnPBatch
    $DIDs | ForEach-Object { Add-PnPListItem -List $ReportList -Values @{ DID = $DIDLookupList[[double]$_.DID] } -Batch $batch }
    Invoke-PnPBatch -Batch $batch
    if ($batch.RequestCount -gt 0) {
        Write-Warning "Batch Failed"
        return
    }
    Write-Information "All items added to Report List"
}