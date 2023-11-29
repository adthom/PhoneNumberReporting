param(
    [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
    [string]
    $Department,

    [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
    [ValidateRange(10000000000,19999999999)]
    [double]
    $DID,

    [Parameter(Mandatory)]
    [string]
    $SharePointDomain,

    [Parameter(Mandatory)]
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
    $ClearExisting
)

begin {
    Connect-PnPOnline -Url "https://${SharePointDomain}.sharepoint.com/sites/${Site}" -Interactive -ErrorAction Stop
    
    $DIDList = Get-PnPList -Identity $DIDDepartmentMapName
    $RBACList = Get-PnPList -Identity $RBACListName
    $ReportList = Get-PnPList -Identity $ReportName

    $DIDs = [Collections.Generic.List[PSCustomObject]]@()
}
process {
    $DIDs.Add([PSCustomObject]@{
        Department = $Department
        DID = $DID
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
    $DIDs.Department | Sort-Object -Unique | ForEach-Object { Add-PnPListItem -List $RBACList -Values @{Department=$_} -Batch $batch }
    Invoke-PnPBatch -Batch $batch -ErrorAction Stop
    Write-Information "Departments added to RBAC List"

    Write-Information "Adding items to DID <-> Department Map"
    $batch = New-PnPBatch
    $DIDs | ForEach-Object { Add-PnPListItem -List $DIDList -Values @{Department=$_.Department;DID=$_.DID} -Batch $batch }
    Invoke-PnPBatch -Batch $batch -ErrorAction Stop
    Write-Information "All items added to DID <-> Department Map"
    
    $DIDLookupList = @{}
    Get-PnPListItem -List $DIDList -Fields Id,DID -PageSize 1000 | ForEach-Object { $DIDLookupList[$_.FieldValues['DID']] = $_.Id }
    
    Write-Information "Adding items to Report List"
    $batch = New-PnPBatch
    $DIDs | ForEach-Object { Add-PnPListItem -List $ReportList -Values @{ DID = $DIDLookupList[[double]$_.DID] } -Batch $batch }
    Invoke-PnPBatch -Batch $batch -ErrorAction Stop
    Write-Information "All items added to Report List"
}