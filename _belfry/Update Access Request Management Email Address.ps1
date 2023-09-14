<#
    Author: Salman Siddiqui
    Date: 2017/03/22
    Department: Public Service Commission

    This script iterates through all sites in $targetWeb site collection,
    checks if User Access Management Email address is specified, then compares
    the specified address against $currentEmail1 and $currentEmail2. Updates 
    the email address if diferent than the standard.
#>

# ************* Constants **********************
$file = "C:\Windows\Temp\PSC-scheduledJob-Output.csv"
$targetWeb = "http://edm-goa-apm-816/" #"https://www.chrshare.alberta.ca"
$currentEmail1 = "psc.sharepointadministrators@gov.ab.ca";
$currentEmail2 = "contacthrcommunity@gov.ab.ca";
$newEmail = "psc.sharepointadministrators@gov.ab.ca";

# ************* end Constants **********************

if ( (Get-PSSnapin -Name "Microsoft.SharePoint.PowerShell") -eq $null )
{
    Add-PsSnapin "Microsoft.SharePoint.PowerShell"
}

$webapp = Get-SPWebApplication $targetWeb
#$fileContent = Import-csv $file -header "URL", "inherited", "email", "standard", "status" 
$psObject | Export-Csv $fileContent -NoTypeInformation

$url = $inherited = $emailAddress = $standard = $status = ""

#SMTP Server: xmail.gov.ab.ca
 
foreach($site in $webapp.Sites)
{
   foreach($web in $site.AllWebs)
   {
     $url = $inherited = $emailAddress = $standard = $status = ""

     $url = $web.url
     #Write-host $url

     if (!$web.HasUniquePerm)
     {
            #Write-Host "inherted from parent"
            $inherited = "inherted from parent"
     }
       elseif($web.RequestAccessEnabled)
       {
            #Write-Host "Not inherited."
            $inherited =  "Not inherited."
            #write-host $web.RequestAccessEmail
            $emailAddress =  $web.RequestAccessEmail

            if (($web.RequestAccessEmail -ine $currentEmail1) -or ($web.RequestAccessEmail -ine $currentEmail2))
            {
                #Write-Host "non-standard"
                $standard = "non-standard"
                $web.RequestAccessEmail = $newEmail
                $web.Update()
                #Write-Host "updated to standard\n"
                $status = "updated to standard"
            }
 
       }
       else
      {
            #Write-Host "inherited"
            $inherited =  "inherited"
      }
      $hashTable = @{"URL" = $url; "inherited" = $inherited; "email" = $emailAddress; "standard" = $standard; "status" = $status;}
      $newRow = New-Object PsObject -Property $hashTable
      Export-Csv $fileContent -inputobject $newrow -append -Force
   }
}

Export-Csv $fileContent 