$farm = Get-SPFarm
$cacheService = $farm.Services | where {$_.Name -eq "AppFabricCachingService"} 
$accnt = Get-SPManagedAccount -Identity GOA\SP2013-UAT-SVC.S 
        $cacheService.ProcessIdentity.CurrentIdentityType = "SpecificUser" 
        $cacheService.ProcessIdentity.ManagedAccount = $accnt 
        $cacheService.ProcessIdentity.Update() 
        $cacheService.ProcessIdentity.Deploy()

