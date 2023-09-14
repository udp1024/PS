#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Import-Module CredentialManager

#Function to set a Field to Read Only in SharePoint Online List
Function Set-SPOFieldReadOnly($SiteURL, $ListName, $FieldInternalName, [Bool]$IsReadOnly)
{
    Try {
        #Get Credentials to connect
        # $Cred= Get-Credential

        $Cred = Get-StoredCredential -Target "My-Dot-Z"
  
        #Setup the context
        $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
        #$Ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.Username, $Cred.Password)
        $Ctx.Credentials = $Cred

        #Get the List
        $List=$Ctx.Web.Lists.GetByTitle($ListName)
 
        #Get the Field
        $Field = $List.Fields.GetByInternalNameOrTitle($FieldInternalName)
        $Field.ReadOnlyField = $IsReadOnly
        $Field.SetShowInEditForm($true)
        $Field.Update()
        $Ctx.ExecuteQuery()
  
        Write-host -f Green "Read Only Settings Update for the Field Successfully!"
    }
    Catch {
        write-host -f Red "Error:" $_.Exception.Message
    }
}
#Set parameter values
$SiteURL = "https://aepuat.sp.gov.ab.ca/Parks/KCP"
$ListName = "KCPxClients"
$FieldInternalName = "Title" #Internal Name
$IsReadOnly = $false
 
Set-SPOFieldReadOnly -SiteURL $SiteURL -ListName $ListName -FieldInternalName $FieldInternalName -IsReadOnly $IsReadOnly


#Read more: https://www.sharepointdiary.com/2018/04/sharepoint-online-make-list-field-read-only-using-powershell.html#ixzz7Ddjz7AYU