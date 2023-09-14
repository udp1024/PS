Import-Module -Name SharePointPnPPowerShell2016
Connect-PnPOnline -Url https://share.tbfsp.gov.ab.ca/

Import-Module AzureAD
$Credential = Get-Credential
Connect-AzureAD -credential $Credential

$aduser = Get-AzureADUser -ObjectId "christine.Sewell@gov.ab.ca"
# Get-PnPUser | ? Email -eq "user@tenant.onmicrosoft.com"
