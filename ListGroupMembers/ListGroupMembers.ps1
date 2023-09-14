$SiteUrl = 'https://myaps.alberta.ca'
$GroupName = 'Style Resource Readers'

$cred = Get-Credential
Connect-PnPOnline -Url $SiteUrl -Credential $cred 
$group = Get-PnPGroup -Identity "$GroupName" -Includes Users
$group.Users | select-Object -property Title | Export-Csv -Path .\GroupMembers-$GroupName.csv -NoTypeInformation
