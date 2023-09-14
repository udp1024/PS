Import-Module -name SharePointPnPPowerShell2016 -DisableNameChecking
cd (Split-Path -Parent $PSCommandPath)
Connect-PnPOnline -Url https://extshare.sp.alberta.ca/EAP -Credentials:SP-EOT-WF.A

$users = Get-PnPGroupMembers -Identity "ERC Council Members" | Select-Object Title,LoginName
$usersTable = @()

if ($users.Count -gt 0){
    foreach($user in $users) {
        if ($user.LoginName.StartsWith("i:0#.w|extern\")){ # is this an EXTERN domain account?
            #Write-Host("External Account "+ $user.Title)
            $adUser = Get-ADUser -Server "EXTERN.gov.ab.ca" -Identity $user.Title -properties DisplayName,passwordlastset,LockedOut,otherMailbox
            if ($adUser.LockedOut) {$adUserStatus = "Locked"}else{$adUserStatus = ""}
            #$tableUser = New-Object psobject -Property ([Ordered]@{Title = $adUser.DisplayName;  PassWordExpires = $adUser.PasswordLastSet.AddDays(60).ToShortDatestring(); DaysToExpiry = (([DateTime]::Today) - $adUser.PasswordLastSet.AddDays(60)).Days;Status = $adUserStatus})
            $tableUser = New-Object psobject -Property ([Ordered]@{Title = $adUser.DisplayName;  PassWordExpires = $adUser.PasswordLastSet.AddYears(1).ToShortDatestring(); DaysToExpiry = (([DateTime]::Today) - $adUser.PasswordLastSet.AddYears(1)).Days;Status = $adUserStatus; Notify = [system.String]::Join(";", $adUser.otherMailbox.ValueList)})
            $usersTable += $tableUser
        }
    } # foreach ($user ...
} # if ($users.Count ...

$arr = ($PWD.Drive.CurrentLocation).Split("\")
switch ($arr[-1]) {
    "UAT" { $configFile = "ERCnotifyUAT.conf" }
    "Dev" { $configFile = "ERCnotifyDEV.conf" }
    "PPMsync" { $configFile = "ERCnotify.conf" }
    default { $configFile = "ERCnotifyDEV.conf" }
}

$hash=ConvertFrom-StringData -StringData (Get-Content $configFile|out-string)
$config = New-Object PsObject -Property $hash
$config.mailBody = (Get-Content $config.mailBodyFile|out-string)
#$h = $usersTable | Sort-Object -Property DaysToExpiry,Title |Format-Table -Property Title, PassWordExpires, DaysToExpiry |Out-String
$h = $usersTable | Sort-Object -Property DaysToExpiry,Title| ConvertTo-Html -Fragment|Out-String
$config.mailBody += $h|Out-String
Send-MailMessage -To $config.mailTo -From $config.mailFrom -Subject $config.mailSubject -Body $config.mailBody -BodyAsHtml -SmtpServer $config.mailServer
