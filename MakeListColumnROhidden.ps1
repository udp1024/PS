$SPhost = "https://aepuat.sp.gov.ab.ca/Parks/KCP"

$List = "KCPxClients"
$targetColumnName = "KCPxClientNumber"

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

# Get Context
$clientContext = Get-PnPContext
 
# -List: The list object or name of the list
# -Identity: The field object or name
$targetField = Get-PnPField -List $List -Identity $targetColumnName
 
# Make list column Read-Only
$targetField.ReadOnlyField = 0   

#Make List column hidden in forms
$targetField.SetShowInEditForm($true)


# write the new properties to the field definition, and ...
$targetField.Update()

# execute.
$clientContext.ExecuteQuery()
 
Disconnect-PnPOnline
