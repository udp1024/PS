.\_pssetup.ps1
$webapp=Get-SPWebApplication https://excshareuat.alberta.ca;
$webapp.Sites | %{Write-Host $_.Url ” -Count: ” $_.AllWebs.Count};
