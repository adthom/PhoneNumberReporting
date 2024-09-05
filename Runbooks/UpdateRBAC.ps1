"Getting Variables" | Write-Output
$SiteDisplayName = Get-AutomationVariable -Name Site -ErrorAction Stop
$SharePointDomain = Get-AutomationVariable -Name SharePointDomain -ErrorAction Stop
$DIDDepartmentMapName = Get-AutomationVariable DIDDepartmentMapName -ErrorAction Stop
$RBACListName = Get-AutomationVariable -Name RBACListName -ErrorAction Stop
$ReportName = Get-AutomationVariable -Name ReportName -ErrorAction Stop
$Environment = Get-AutomationVariable -Name Environment -ErrorAction Stop
$Tenant = Get-AutomationVariable -Name Tenant -ErrorAction Stop

$Site = $SiteDisplayName -replace '\s', ''
$SharepointTLD = switch ($Environment) {
    'China' { 'sharepoint.cn'; break }
    'USGovernmentHigh' { 'sharepoint.us'; break }
    'USGovernmentDoD' { 'sharepoint-mil.us'; break }
    default { 'sharepoint.com' }
}
$ConnectPnPOnlineParams = @{
    Url              = "https://${SharePointDomain}.${SharepointTLD}/sites/${Site}"
    ManagedIdentity  = $true
    ErrorAction      = 'Stop'
    AzureEnvironment = $Environment
    Tenant           = $Tenant
}

"Connecting to $($ConnectPnPOnlineParams['Url'])" | Write-Output
try {
    $env:PNPPOWERSHELL_UPDATECHECK = 'Off'
    $null = Connect-PnPOnline @ConnectPnPOnlineParams
}
catch {
    Write-Error $_
    throw
}

"Getting Current Site Owners" | Write-Output
$Owners = Get-PnPGroup "$SiteDisplayName Owners" -ErrorAction Stop
"Getting Current RBAC List" | Write-Output
$RBACList = Get-PnPList $RBACListName -ErrorAction Stop
"Getting Current DID <-> Department List" | Write-Output
$DIDDepartmentMap = Get-PnPList $DIDDepartmentMapName -ErrorAction Stop
"Getting Current Report List" | Write-Output
$ReportList = Get-PnPList $ReportName -ErrorAction Stop

# Get last run date from variable
$LastRun = Get-AutomationVariable -Name RBACLastRun -ErrorAction SilentlyContinue
if ($null -eq $LastRun -or $LastRun.AddYears(1) -lt [DateTime]::Today) {
    # if last run is null or more than a year ago, set to a year ago
    $LastRun = [DateTime]::Today.AddYears(-1)
}
"Getting All RBAC Changes Since $($LastRun.ToString('o'))" | Write-Output
$ModifiedQuery = '<View><Query><Where><Geq><FieldRef Name="Modified"/><Value Type="DateTime" IncludeTimeValue="TRUE">{0:o}</Value></Geq></Where></Query>{{0}}</View>' -f $LastRun.AddHours(-7).ToUniversalTime()
$LastRun = [DateTime]::Now
$RBACFields = '<ViewFields><FieldRef Name="Identity"/><FieldRef Name="Modified"/><FieldRef Name="Department"/></ViewFields>'

try {
    "Getting All Report Items" | Write-Output
    $ReportItems = Get-PnPListItem -List $ReportList -PageSize 1000 -ErrorAction Stop
}
catch {
    Write-Error $_
    throw
}

try {
    "Getting All DID <-> Department Items" | Write-Output
    $DIDDepartmentItems = Get-PnPListItem -List $DIDDepartmentMap -PageSize 1000 -ErrorAction Stop
}
catch {
    Write-Error $_
    throw
}

$CurrentErrorCount = $Error.Count
Get-PnPListItem -List $RBACList -Query ($ModifiedQuery -f $RBACFields) -PageSize 1000 -ScriptBlock { param($items) $items.Context.ExecuteQuery() } |
    ForEach-Object -Begin {$j=0;$k=0;$d=0} -Process {
        $updated = $_
        $Department = $updated.FieldValues['Department']
        $d++
        $RBAC = $updated.FieldValues['Identity']
        $ReportItems | Where-Object { $_.FieldValues['Department'].LookupValue -eq $Department } |
            ForEach-Object -Begin {$i=0} -Process {
                $ID = $_.Id
                Set-PnPListItemPermission -List $ReportList -Identity $Id -AddRole 'Full Control' -Group $Owners.Title -ClearExisting
                if ($RBAC.Count -gt 0) {
                    $RBAC | ForEach-Object { Set-PnPListItemPermission -List $ReportList -Identity $Id -AddRole Read -User $_.Email }
                }
                $i++
                if((++$j % 25) -eq 0) { "$Department $i report rows processed ($j report rows and $k source rows rows processed)" | Write-Output }
            } -End {
                "$Department $i report rows processed ($j report rows and $k source rows processed)" | Write-Output
            }
        $DIDDepartmentItems | Where-Object { $_.FieldValues['Department'] -eq $Department } |
            ForEach-Object -Begin {$i=0} -Process {
                    $ID = $_.Id
                    Set-PnPListItemPermission -List $DIDDepartmentMap -Identity $Id -AddRole 'Full Control' -Group $Owners.Title -ClearExisting
                    if ($RBAC.Count -gt 0) {
                        $RBAC | ForEach-Object { Set-PnPListItemPermission -List $DIDDepartmentMap -Identity $Id -AddRole Read -User $_.Email }
                    }
                    $i++
                    if((++$k % 25) -eq 0) { "$Department $i source rows processed ($j report rows and $k source rows processed)" | Write-Output }
                } -End {
                    "$Department $i source rows processed ($j report rows and $k source rows processed)" | Write-Output
                }
    } -End {
        "Processed $d RBAC changes ($j report rows and $k source rows affected)" | Write-Output
    }

if ($CurrentErrorCount -lt $Error.Count) {
    Write-Error "There were errors during the RBAC update, please review the error stream"
    throw
}
# save last run to variable for next runs
$LastRun = Set-AutomationVariable RBACLastRun -Value $LastRun
