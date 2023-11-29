"Connecting to Azure" | Write-Output
$null = Connect-AzAccount -Identity

"Getting Variables" | Write-Output
$Site = Get-AutomationVariable -Name Site -ErrorAction Stop
$SharePointDomain = Get-AutomationVariable -Name SharePointDomain -ErrorAction Stop
$RBACListName = Get-AutomationVariable -Name RBACListName -ErrorAction Stop
$ReportName = Get-AutomationVariable -Name ReportName -ErrorAction Stop

"Connecting to https://${SharePointDomain}.sharepoint.com/sites/${Site}" | Write-Output
try {
    $null = Connect-PnPOnline -Url "https://${SharePointDomain}.sharepoint.com/sites/${Site}" -ManagedIdentity -ErrorAction Stop
}
catch {
    Write-Error $_
    throw
}

"Getting Current Site Owners" | Write-Output
$Owners = Get-PnPUser "$Site Owners" -ErrorAction Stop
"Getting Current RBAC List" | Write-Output
$RBACList = Get-PnPList $RBACListName -ErrorAction Stop
"Getting Current Report List" | Write-Output
$ReportList = Get-PnPList $ReportName -ErrorAction Stop

# Get last run date from variable
$LastRun = Get-AutomationVariable -Name RBACLastRun -ErrorAction SilentlyContinue
if ($null -eq $LastRun -or $LastRun.AddYears(1) -lt [DateTime]::Today) {
    # if last run is null or more than a year ago, set to a year ago
    $LastRun = [DateTime]::Today.AddYears(-1)
}
"Getting All RBAC Changes Since $($LastRun.ToString('o'))" | Write-Output
$ModifiedQuery = '<View><Query><Where><Geq><FieldRef Name="Modified"/><Value Type="DateTime" IncludeTimeValue="TRUE">{0:o}</Value></Geq></Where></Query>{{0}}</View>' -f $LastRun.ToUniversalTime()
$LastRun = [DateTime]::Now
$RBACFields = '<ViewFields><FieldRef Name="Identity"/><FieldRef Name="Modified"/><FieldRef Name="Department"/></ViewFields>'
$ReportListFields = '<ViewFields><FieldRef Name="Id"/></ViewFields>'
$DepartmentQuery = '<View><Query><Where><Eq><FieldRef Name="Department"/><Value Type="Text">{0}</Value></Eq></Where></Query></View>'
$CurrentErrorCount = $Error.Count
Get-PnPListItem -List $RBACList -Query ($ModifiedQuery -f $RBACFields) -PageSize 1000 -ScriptBlock { param($items) $items.Context.ExecuteQuery() } |
    ForEach-Object -Begin {$j=0;$d=0} -Process {
        $updated = $_
        $Department = $updated.FieldValues['Department']
        $d++
        $RBAC = $updated.FieldValues['Identity']
        $DepQuery = $DepartmentQuery -f $Department, $ReportListFields
        Get-PnPListItem -List $ReportList -Query $DepQuery -PageSize 1000 -ScriptBlock { param($items) $items.Context.ExecuteQuery() } |
            ForEach-Object -Begin {$i=0} -Process {
                $ID = $_.Id
                Set-PnPListItemPermission -List $ReportList -Identity $Id -AddRole 'Full Control' -Group $Owners.Title -ClearExisting
                if ($RBAC.Count -gt 0) {
                    $RBAC | ForEach-Object { Set-PnPListItemPermission -List $ReportList -Identity $Id -AddRole Read -User $_.Email }
                }
                $i++
                if((++$j % 25) -eq 0) { "$Department $i rows processed ($j rows processed)" | Write-Output }
            } -End {
                "$Department $i rows processed ($j rows processed)" | Write-Output
            }
    } -End {
        "Processed $d RBAC changes ($j rows affected)" | Write-Output
    }

if ($CurrentErrorCount -lt $Error.Count) {
    Write-Error "There were errors during the RBAC update, please review the error stream"
    throw
}
# save last run to variable for next runs
$LastRun = Set-AutomationVariable RBACLastRun -Value $LastRun
