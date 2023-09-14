$SPhost = "https://pscuat.sp.gov.ab.ca/PSCFOIP"
$path = ".\Groups.csv"

$GroupName = ""
$Owner = ""
$Description = ""

if (Get-Module -ListAvailable -Name SharePointPnPPowerShell2016) {
    Write-Output "Module SharePointPnPPowerShell2016 exists"
    Import-Module -name SharePointPnPPowerShell2016 -DisableNameChecking 
} 
else {
    Write-Output "Module SharePointPnPPowerShell2016 does not exist. Installing it"
    Install-Module -name SharePointPnPPowerShell2016 -Scope CurrentUser -AllowClobber
    Import-Module -name SharePointPnPPowerShell2016 -DisableNameChecking 
}

$global:SPconnection = Connect-PnPOnline -Url $SPhost -Credentials:My-Dot-Z -ReturnConnection

<#Import-Csv $path | Foreach-Object { 

    foreach ($property in $_.PSObject.Properties)
    {
        # doSomething $property.Name, $property.Value
        # New-PnPGroup -Title "New SP Group" -Description "New SP Group Description" -AllowMembersEditMembership -DisallowMembersViewMembership -AllowRequestToJoinLeave -AutoAcceptRequestToJoinLeave -Owner "Site Admins Group"
        
    } 

}#>

$csv = Import-Csv $path
foreach($line in $csv)
{ 
    $properties = $line | Get-Member -MemberType Properties
    for($i=0; $i -lt $properties.Count;$i++)
    {
        $column = $properties[$i]
        $columnvalue = $line | Select-Object -ExpandProperty $column.Name
        # Write-Output ("Set "+$column.Name+" to "+$columnvalue)
        
        # doSomething $column.Name $columnvalue 
        # doSomething $i $columnvalue
        Set-Variable -Name $column.Name -Value $columnvalue -Scope global
    }
    $GroupName
    $Owner
    $Description
    Write-Output "*** eol ***"
    New-PnPGroup -Title $GroupName -Description $Description -Owner $Owner
} 
