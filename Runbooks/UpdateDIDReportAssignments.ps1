"Getting Variables" | Write-Output
$Site = Get-AutomationVariable -Name Site -ErrorAction Stop
$SharePointDomain = Get-AutomationVariable -Name SharePointDomain -ErrorAction Stop
$ReportName = Get-AutomationVariable -Name ReportName -ErrorAction Stop

"Connecting to https://${SharePointDomain}.sharepoint.com/sites/${Site}" | Write-Output
try {
    $env:PNPPOWERSHELL_UPDATECHECK='Off'
    $null = Connect-PnPOnline -Url "https://${SharePointDomain}.sharepoint.com/sites/${Site}" -ManagedIdentity -ErrorAction Stop
}
catch {
    Write-Error $_
    throw
}

"Connecting to MicrosoftTeams" | Write-Output
try {
    $null = Connect-MicrosoftTeams -Identity -ErrorAction Stop
}
catch {
    Write-Error $_
    throw
}

"Getting Current Phone Number Report And Items" | Write-Output
$ReportList = Get-PnPList -Identity $ReportName -ErrorAction Stop
$PhoneReportQuery = '<View><ViewFields><FieldRef Name="DID"/><FieldRef Name="Department"/><FieldRef Name="AssignedIdentity"/></ViewFields></View>'
$ListResults = Get-PnPListItem -List $ReportList -Query $PhoneReportQuery -PageSize 1000

$GetPhoneNumberAssignmentParams = @{
    NumberType  = 'DirectRouting'
    Top         = 1000
    ErrorAction = 'Stop'
}
$NumberLookup = @{}
do {
    "Getting next $($GetPhoneNumberAssignmentParams['Top']) numbers" | Write-Output
    $Results = @(try { Get-CsPhoneNumberAssignment @GetPhoneNumberAssignmentParams | Select-Object TelephoneNumber, AssignedPstnTargetId }catch { Write-Error $_; throw })
    $GetPhoneNumberAssignmentParams['Skip'] = $GetPhoneNumberAssignmentParams['Top'] + $GetPhoneNumberAssignmentParams['Skip']
    $Results | ForEach-Object {
        $Number = $_.TelephoneNumber
        $NumberLookup[$Number.TrimStart('+')] = [PSCustomObject]@{
            Number               = $Number
            AssignedPstnTargetId = $_.AssignedPstnTargetId
        }
    }
} while ($Results.Count -eq $GetPhoneNumberAssignmentParams['Top'])

$batch = New-PnPBatch
$CurrentErrorCount = $Error.Count
"0 rows processed" | Write-Output
$changes = 0
$ListResults | ForEach-Object -Begin { $i = 0 } -Process {
    $key = $_.FieldValues['DID'].LookupValue.Split('.', 2)[0]
    $AssignmentInfo = $NumberLookup[$key]
    if ($NumberLookup.ContainsKey($key)) { $NumberLookup.Remove($key) }
    $AssignedIdentity = $null
    $AssignedIdentity = if ($AssignmentInfo.AssignedPstnTargetId) { (Get-CsOnlineUser -Identity $AssignmentInfo.AssignedPstnTargetId).UserPrincipalName } else { $null }
    if ($AssignedIdentity -ne $_.FieldValues['AssignedIdentity'].Email) {
        $changes++
        Set-PnPListItem -List $ReportList -Identity $_.Id -Values @{ AssignedIdentity = $AssignedIdentity } -Batch $batch
    }
    if ((++$i % 250) -eq 0) { "$i rows of $($ListResults.Count) processed" | Write-Output }
} -End { "$i rows of $($ListResults.Count) processed" | Write-Output }

"Found $changes Changes" | Write-Output
if ($changes.Count -gt 0) {
    "Invoking Batch" | Write-Output
    Invoke-PnPBatch -Batch $batch -Force
}

if ($NumberLookup.Count -gt 0) {
    Write-Warning "$($NumberLookup.Count) numbers not found in report"
    # No mechanism yet for how to notify an admin that there are numbers not in the report
    foreach ($number in $NumberLookup.Keys) {
        "Report is missing: $number" | Write-Output
    }
}

if ($CurrentErrorCount -lt $Error.Count) {
    Write-Error "There were errors during the report update, please review the error stream"
    throw
}
