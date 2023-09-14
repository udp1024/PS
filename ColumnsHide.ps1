if (Get-Module -ListAvailable -Name SharePointPnPPowerShell2016) {
    Write-Output "Module SharePointPnPPowerShell2016 exists"
    Import-Module -name SharePointPnPPowerShell2016 -DisableNameChecking 
} 
else {
    Write-Output "Module SharePointPnPPowerShell2016 does not exist. Installing it"
    Install-Module -name SharePointPnPPowerShell2016 -Scope CurrentUser -AllowClobber
    Import-Module -name SharePointPnPPowerShell2016 -DisableNameChecking 
}

#$SPhost = "https://myapsuat.sp.alberta.ca"
$SPhost = "https://li.sp.gov.ab.ca/OHSIMP2021"

$List = "Notification Subscribers - SIT Reports"
#$targetColumnNames = @("Title", "ApplicationCreator", "ApplicationCreatedOn", "ApplicationsSubmittedBy")
$targetColumnNames = @("Title")

$global:SPconnection = Connect-PnPOnline -Url $SPhost -Credentials:My-Dot-Z -ReturnConnection

# Get Context
$clientContext = Get-PnPContext

foreach ($column in $targetColumnNames){ 
# -List: The list object or name of the list
# -Identity: The field object or name
$targetField = Get-PnPField -List $List -Identity $column
 
# Make list column Read-Only
$targetField.ReadOnlyField = 1

#Make List column hidden in forms
$targetField.SetShowInEditForm($false)


# write the new properties to the field definition, and ...
$targetField.Update()
}

# execute.
$clientContext.ExecuteQuery()
 
Disconnect-PnPOnline