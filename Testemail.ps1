Import-Module -name SharePointPnPPowerShell2016 -DisableNameChecking
Connect-PnPOnline -Url https://extshare.sp.alberta.ca/EAP -Credentials:SP-EOT-WF.A

$users = Get-PnPGroupMembers -Identity "ERC Council Members" | Select-Object Title,LoginName
$usersTable = @()

if ($users.Count -gt 0){
    foreach($user in $users) {
        if ($user.LoginName.StartsWith("i:0#.w|extern\")){ # is this an EXTERN domain account?
            Write-Host("External Account "+ $user.Title)
            $adUser = Get-ADUser -Server "EXTERN.gov.ab.ca" -Identity $user.Title -properties DisplayName,passwordlastset
            $tableUser = New-Object psobject -Property @{Title = $adUser.DisplayName;  PassWordExpires = $adUser.PasswordLastSet.AddDays(60).Tostring(); DaysToExpiry = (([DateTime]::Today) - $adUser.PasswordLastSet.AddDays(60)).Days;}
            $usersTable += $tableUser
            <#
            $daysToX = ($m.PasswordLastSet.AddDays(60)) - ([DateTime]::Today)
            if (($daysToX.Days -eq 30) -or ($daysToX.Days -eq 15)) {
                #Add to notification text
                $userTable += @([PSCustomObject]@{Title = $user.Title;  PWexpiry = $m;})
            }#>
        }
    } # foreach ($user ...
} # if ($users.Count ...
$userTableHtml = $usersTable |Format-Table -Property Title, PassWordExpires, DaysToExpiry|ConvertTo-Html -Fragment



<#
$arr = ($PWD.Drive.CurrentLocation).Split("\")
switch ($arr[-1]) {
    "UAT" { $configFile = "ERCnotifyUAT.conf" }
    "Dev" { $configFile = "ERCnotifyDEV.conf" }
    "PPMsync" { $configFile = "ERCnotify.conf" }
    default { $configFile = "ERCnotifyDEV.conf" }
}

$hash=ConvertFrom-StringData -StringData (Get-Content $configFile|out-string)
$config = New-Object PsObject -Property $hash
$config.mailBody = Get-Content $config.mailBodyFile|out-string

#Send-MailMessage -To “salman.siddiqui@gov.ab.ca” -From “salman.siddiqui@gov.ab.ca”  -Subject “test message subject” -Body “<p><h1>test</h1></p><p>important plain text!</p>” -BodyAsHtml -SmtpServer “xmail.gov.ab.ca”
Send-MailMessage -To $config.mailTo -From $config.mailFrom -Subject $config.mailSubject -Body $config.mailBody -BodyAsHtml -SmtpServer $config.mailServer
#>