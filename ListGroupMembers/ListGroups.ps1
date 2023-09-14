$cred = Get-Credential  
Connect-PnPOnline -Url https://myaps.alberta.ca -Credential $cred  
$groups = Get-PnPGroup | Select-Object Title,Users  
$groups | format-table @{Expression = {$_.Title};Label='Group'},@{Expression = {$_.Users.Title};Label='Users'},@{Expression = {$_.Users.Count};Label='UsersCount'}
