[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
Import-Module SharePointPnPPowerShell2016 -DisableNameChecking

$SiteUrl = "https://chrsharedev.sp.alberta.ca/WHS"
$ListName="Certification of Facilitating Working Minds"
$ImportFile ="F:\Powershell\Import CSV to SP List\Certification of Facilitating Working Minds.csv"

#$connection = Connect-PnPOnline -Url $SiteUrl -Credentials (Get-Credential) -ReturnConnection
#$connection = Connect-PnPOnline -Url $SiteUrl -CurrentCredentials -ReturnConnection
Connect-PnPOnline -Url $SiteUrl -Credentials:My-Dot-Z

#Get the List
$List = Get-PnpListItem -List $ListName
 
#Get the Data from CSV and Add to SharePoint List
$data = Import-Csv $ImportFile

Foreach ($row in $data) {
    #add item to List

    Add-PnPListItem -List $ListName -Values @{
    "Employee_x0020_Name" = $row.'Employee_Name';
    "Department" = $row.'Department';
    "Course1_x0020_Date" = $row.'Course1_Date';
    "Course2_x0020_Date" = $row.'Course2_Date';
    "Status" = $row.'Status';
    "Expiry" = $row.'Expiry';
    }    
}
Write-host "CSV data Imported to SharePoint List Successfully!"

$strQuery = "<View><Query><OrderBy><FieldRef Name='ID' /></OrderBy></Query></View>"
$items=Get-PnPListItem -List $ListName -Query $strQuery

Foreach ($item in $items) {
	
}