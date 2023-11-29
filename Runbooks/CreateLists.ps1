"Connecting to Azure" | Write-Output
$null = Connect-AzAccount -Identity

"Getting Variables" | Write-Output
$Site = Get-AutomationVariable -Name Site -ErrorAction Stop
$DIDDepartmentMapName = Get-AutomationVariable DIDDepartmentMapName -ErrorAction Stop
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

"Getting or Creating the DID <-> Department Map List" | Write-Output
$DIDList = & {
    $DIDList = Get-PnPList -Identity $DIDDepartmentMapName -ErrorAction SilentlyContinue
    if ($null -eq $DIDList) {
        $Result = New-PnPList -Title $DIDDepartmentMapName -Template GenericList
        $DIDList = Get-PnPList -Identity $DIDDepartmentMapName
    
        $DIDField = Get-PnPField -Identity DID -List $DIDList -ErrorAction SilentlyContinue
        if ($null -eq $DIDField) {
            $Formatter = '{"$schema":"https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json","elmType":"span","style":{"overflow":"hidden","text-overflow":"ellipsis","padding":"0 3px"},"txtContent":"=''+1 (''+substring(toString(@currentField),1,4)+'') ''+substring(toString(@currentField),4,7)+''-''+substring(toString(@currentField),7,11)","attributes":{"class":""}}'
            $xml = '<Field CommaSeparator="FALSE" CustomUnitOnRight="TRUE" Decimals="0" DisplayName="DID" EnforceUniqueValues="TRUE" Format="Dropdown" Indexed="TRUE" IsModern="TRUE" Max="19999999999" Min="10000000000" Name="DID" Percentage="FALSE" Required="TRUE" Title="DID" Type="Number" Unit="None" StaticName="DID" RowOrdinal="0" CustomFormatter="{0}" ><Validation Message="DID must be 11 digits starting with a 1.">=LEN(TRIM(DID))=11</Validation></Field>' -f $Formatter.Replace('"','&quot;')
            $Result = Add-PnPFieldFromXml -List $DIDList -FieldXml $xml
            $DIDField = Get-PnPField -Identity DID -List $DIDList -ErrorAction Stop
        }
    
        $DepartmentField = Get-PnpField -List $DIDList -Identity Department -ErrorAction SilentlyContinue
        if ($null -eq $DepartmentField) {
            $xml = '<Field Name="Department" StaticName="Department" Description="The Department which owns the DID" DisplayName="Department" ReadOnly="FALSE" Type="Text" FromBaseType="TRUE" />'
            $Result = Add-PnPFieldFromXml -List $DIDList -FieldXml $xml
            $DepartmentField = Get-PnPField -Identity Department -List $DIDList -ErrorAction Stop
        }
    
        $TitleField = Get-PnpField -List $DIDList -Identity Title
        $TitleField.Hidden = $true
        $TitleField.Required = $false
        $TitleField.UpdateAndPushChanges($true)
    
        $TitleField = Get-PnpField -List $DIDList -Identity LinkTitle
        $TitleField.Hidden = $true
        $TitleField.Required = $false
        $TitleField.UpdateAndPushChanges($true)
    
        $DIDList = Get-PnPList -Identity $DIDDepartmentMapName -ErrorAction Stop
    
        $Result = $DIDList | Get-PnpView -Identity 'All Items' | Set-PnpView -Fields DID,Department

        $DIDList.BreakRoleInheritance($false,$true)
        Set-PnPListPermission -Identity $DIDList -AddRole 'Full Control' -Group $Owners.Title
    }
    return $DIDList
}

"Getting or Creating the RBAC List" | Write-Output
$RBACList = & {
    $RBACList = Get-PnPList -Identity $RBACListName -ErrorAction SilentlyContinue
    if ($null -eq $RBACList) {
        $Result = New-PnPList -Title $RBACListName -Template GenericList
        $RBACList = Get-PnPList -Identity $RBACListName

        $RBACDepartmentField = Get-PnpField -List $RBACList -Identity Department -ErrorAction SilentlyContinue
        if ($null -eq $RBACDepartmentField) {
            $xml = '<Field Type="Text" Name="Department" DisplayName="Department" Required="TRUE" StaticName="Department" ColName="nvarchar7" RowOrdinal="0" Description="The Agency/Commission/Board with which to associate the user(s) or group(s)" EnforceUniqueValues="TRUE" Hidden="FALSE" Indexed="TRUE" ShowInFiltersPane="Pinned" />'
            $Result = Add-PnPFieldFromXml -List $RBACList -FieldXml $xml
            $RBACDepartmentField = Get-PnPField -Identity Department -List $RBACList -ErrorAction Stop
        }

        $IdentityField = Get-PnPField -Identity Identity -List $RBACList -ErrorAction SilentlyContinue
        if ($null -eq $IdentityField) {
            $Result = Add-PnPFieldFromXml -List $RBACList -FieldXml '<Field Description="The User(s) Or Group(s) needing read-only access to the report" DisplayName="Identity" Format="Dropdown" List="UserInfo" Mult="TRUE" Name="Identity" Title="Identity" Type="UserMulti" UserDisplayOptions="NamePhoto" UserSelectionMode="1" UserSelectionScope="0" StaticName="Identity" ColName="int1" RowOrdinal="0" />'
            $IdentityField = Get-PnPField -Identity Identity -List $RBACList -ErrorAction Stop
        }

        $TitleField = Get-PnpField -List $RBACList -Identity Title
        $TitleField.Hidden = $true
        $TitleField.Required = $false
        $TitleField.UpdateAndPushChanges($true)

        $TitleField = Get-PnpField -List $RBACList -Identity LinkTitle
        $TitleField.Hidden = $true
        $TitleField.Required = $false
        $TitleField.UpdateAndPushChanges($true)

        $RBACList = Get-PnPList -Identity $RBACListName -ErrorAction Stop
        $Result = $RBACList | Get-PnpView -Identity 'All Items' | Set-PnpView -Fields Department,Identity

        $RBACList.BreakRoleInheritance($false,$true)
        Set-PnPListPermission -Identity $RBACList -AddRole 'Full Control' -Group $Owners.Title
    }
    return $RBACList
}

"Getting or Creating the Report List" | Write-Output
$ReportList = & {
    $ReportList = Get-PnPList -Identity $ReportName -ErrorAction SilentlyContinue
    if ($null -eq $ReportList) {
        $Result = New-PnPList -Title $ReportName -Template GenericList -OnQuickLaunch
        $ReportList = Get-PnPList -Identity $ReportName

        $DIDLookupField = Get-PnPField -Identity DID -List $ReportList -ErrorAction SilentlyContinue
        if ($null -eq $DIDLookupField) {

            $Formatter = '{"$schema":"https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json","elmType":"span","style":{"overflow":"hidden","text-overflow":"ellipsis","padding":"0 3px"},"txtContent":"=''+1 (''+substring(toString(@currentField.lookupValue),1,4)+'') ''+substring(toString(@currentField.lookupValue),4,7)+''-''+substring(toString(@currentField.lookupValue),7,11)","attributes":{"class":""}}'
            $xml = '<Field DisplayName="DID" Format="Dropdown" EnforceUniqueValues="TRUE" Indexed="TRUE" IsModern="TRUE" List="{0}" Name="DID" Required="TRUE" ShowField="DID" Title="DID" Type="Lookup" StaticName="DID" ColName="int1" RowOrdinal="0" CustomFormatter="{1}" />' -f $DIDList.Id.Guid, $Formatter.Replace('"','&quot;')
            $Result = Add-PnPFieldFromXml -List $ReportList -FieldXml $xml
            $DIDLookupField = Get-PnPField -Identity DID -List $ReportList -ErrorAction Stop
        }

        $DepartmentLookupField = Get-PnPField -Identity Department -List $ReportList -ErrorAction SilentlyContinue
        if ($null -eq $DepartmentLookupField) {
            $xml = '<Field DisplayName="Department" FieldRef="{1}" Format="Dropdown" IsModern="TRUE" List="{0}" Name="Department" Required="TRUE" ShowField="Department" Title="Department" Type="Lookup" StaticName="Department" ReadOnly="TRUE" />' -f $DIDList.Id.Guid, $DIDLookupField.Id.Guid
            $Result = Add-PnPFieldFromXml -List $ReportList -FieldXml $xml
            $DepartmentLookupField = Get-PnPField -Identity Department -List $ReportList -ErrorAction Stop
        }

        $AssignedIdentityField = Get-PnPField -Identity AssignedIdentity -List $ReportList -ErrorAction SilentlyContinue
        if ($null -eq $AssignedIdentityField) {
            $Result = Add-PnPFieldFromXml -List $ReportList -FieldXml '<Field DisplayName="Assigned Identity" Format="Dropdown" EnforceUniqueValues="TRUE" Indexed="TRUE" IsModern="TRUE" List="UserInfo" Name="AssignedIdentity" Title="Assigned Identity" Type="User" UserDisplayOptions="NamePhoto" UserSelectionMode="0" UserSelectionScope="0" StaticName="AssignedIdentity" ColName="int2" RowOrdinal="0" />'
            $AssignedIdentityField = Get-PnPField -Identity AssignedIdentity -List $ReportList -ErrorAction Stop
        }

        $TitleField = Get-PnpField -List $ReportList -Identity Title
        $TitleField.Hidden = $true
        $TitleField.Required = $false
        $TitleField.UpdateAndPushChanges($true)

        $TitleField = Get-PnpField -List $ReportList -Identity LinkTitle
        $TitleField.Hidden = $true
        $TitleField.Required = $false
        $TitleField.UpdateAndPushChanges($true)

        $ReportList = Get-PnPList -Identity $ReportName -ErrorAction Stop

        $Result = $ReportList | Get-PnpView -Identity 'All Items' | Set-PnpView -Fields Department,DID,AssignedIdentity

        $ReportList.BreakRoleInheritance($false,$true)
        Set-PnPListPermission -Identity $ReportList -AddRole 'Full Control' -Group $Owners.Title
    }
    return $ReportList
}
