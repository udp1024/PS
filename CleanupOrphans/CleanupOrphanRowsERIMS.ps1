# ensure PnP is installed and loaded
$pnp = Get-Command Connect-PnpOnline -ErrorAction SilentlyContinue
if (!$pnp) {Install-Module SharePointPnPPowerShell2016 -Force}
Import-Module SharePointPnPPowerShell2016

# set constants
$OrphansList = "2021-22 Operational Plan (All Staff)"
$OrphanListfields = "PItemID"
$SourceList = "2021-22 Operational Plan"
$SourceListfields = "ID"

<#$Stages = @(
    "https://chrshare.alberta.ca/StrategicPlanning"
    "https://chrshareuat.sp.alberta.ca/StrategicPlanning"
    "https://chrsharedev.sp.alberta.ca/StrategicPlanning"
)#>
$Stages = @(
    "https://chrsharedev.sp.alberta.ca/StrategicPlanning"
)


foreach ($stage in $stages.GetEnumerator()) {
    $URL = $stage

    #Connect
    $connection = Connect-PnPOnline -URL $URL -Credentials:SP-workflows -ReturnConnection

    #Get Source List
    $SourceItems = Get-PnPListItem -List $SourceList -Connection $connection -Fields $SourceListfields
    $arrSourceItemIDs = @()
    foreach ($item in $SourceItems) {
        $arrSourceItemIDs = $arrSourceItemIDs + $item["ID"]
    }

    $AllStaffItems = Get-PnPListItem -List $OrphansList -Connection $connection -Fields $OrphanListfields 
    
    foreach ($item in $AllStaffItems) {
        if ($item["PItemID"] -notin $arrSourceItemIDs){
            Write-Host "Item ID: " $item["ID"] "P Item ID: " $item["PItemID"]
            Remove-PnPListItem -List $OrphansList -Identity $item["ID"] -Recycle -Force
        }
    }

    Disconnect-PnPOnline -Connection $connection
} #end of foreach ($stage in $stages...
