.\_pssetup.ps1

$webApp = Get-SPWebApplication "https://excshareuat.alberta.ca"
[Microsoft.SharePoint.Publishing.PublishingCache]::FlushBlobCache($webApp)
Write-Host "Flushed the BLOB cache for:" $webApp
