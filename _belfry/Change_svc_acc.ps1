$userNameWithDomain = "GOA\SP2013-UAT-SVC.S"
$password = ConvertTo-SecureString “Y98GZy5dDXCE” -AsPlainText -Force
$sharePointCredential = New-Object System.Management.Automation.PSCredential $userNameWithDomain, $password

# Get an array of all services
$services = Get-SPServiceApplication

# Iterate each item in the array to replace that service's account
foreach ($row in $services) {
 
[SPServiceApplication] $svc = Get-SPServiceApplication -Identity $row.Id
#
# To set the account associated with a particular Service Instance using Windows PowerShell 
# we simply get the ProcessIdentity property of the Service Instance and set its Username property. 
# Once set we call Update() to update the Configuration Database and then Deploy() to push 
# the change out to all Service Instances.
#
<#    $pi = $svc.Service.($row.Id)
    $pi = $svc.Service.ProcessIdentity 
    if ($pi.Username-ne $username) { 
       $pi.Username= $username 
       $pi.Update() 
       $pi.Deploy() 
     } 
 #>

    $pi = $svc.Service.ProcessIdentity 

 }