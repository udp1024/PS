if (!(Get-PSRepository -Name local)) {
    $parameters = @{
        Name = "local"
        SourceLocation = "F:/Powershell/PS-Module/"
        PublishLocation = "F:/Powershell/PS-Module/"
        InstallationPolicy = "Trusted"
      }
      Register-PSRepository @parameters

}

<# if (Get-Module -ListAvailable -Name PnP.PowerShell) {
    #Write-Output "Module PnP.PowerShell exists"
    Import-Module -name PnP.PowerShell -DisableNameChecking 
} 
else {
    Write-Output "Module PnP.PowerShell does not exist. Installing it"
    Install-Module -name PnP.PowerShell -Scope CurrentUser -AllowClobber -Repository local
    Import-Module -name PnP.PowerShell -DisableNameChecking 
}#>
if (Get-Module -ListAvailable -Name SharePointPnPPowerShell2016) {
    #Write-Output "Module SharePointPnPPowerShell2016 exists"
    Import-Module -name SharePointPnPPowerShell2016 -DisableNameChecking 
} 
else {
    Write-Output "Module SharePointPnPPowerShell2016 does not exist. Installing it"
    Install-Module -name SharePointPnPPowerShell2016 -Scope CurrentUser -AllowClobber -Repository local
    Import-Module -name SharePointPnPPowerShell2016 -DisableNameChecking 
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$configFile = "LibFunctionsUAT.conf"
Foreach ($i in $(Get-Content $configFile)){
    $varName = $i.split("=")[0].trim()
    $varValue = $i.split("=",2)[1].trim()
    #write-host "varName = " $varName
    #write-host "varValue = " $varValue
    Set-Variable -Name $varName -Value $varValue -Scope global
}

#$global:SPconnection = Connect-PnPOnline -Url $SPhost -Credentials:SP-PnP-Connect -ReturnConnection
#$global:SPconnection = Connect-PnPOnline -Url $SPhost -Credentials:PPM-API-Sync.S -ReturnConnection
$global:SPconnection = Connect-PnPOnline -Url $SPhost -Credentials:My-Dot-Z -ReturnConnection
$global:startDateTime = (Get-Date).ToUniversalTime()

#$user = Get-MsolUser -UserPrincipalName Elizabeth.Wightman@gov.ab.ca
# $user = Get-ADUser Elizabeth.Wightman -properties "SID"

# $spUser = Get-PnPUser -WithRightsAssignedDetailed
$spUser = Get-PnPUser | Where-Object Email -eq "Elizabeth.Wightman@gov.ab.ca"
If ($spUser.UserId.NameId -ne $user.SID.Value) {
    Write-Host "SID does not match"
}

Write-Debug $spUser.UserId.NameId
Write-Debug $user.SID.Value

$user










Disconnect-PnPOnline -Connection $global:SPconnection
