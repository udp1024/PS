Invoke-Command -ScriptBlock {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
    Import-Module .\LibFunctions.psm1
    Set-Variable -Name "LogFile" -Value "" -Scope script

    SetupModules
    setStaticVars

    $AllFiles =@()
    $queryStr = "<View><ViewFields><FieldRef Name='Title'/><FieldRef Name='Modified'/><FieldRef Name='GUID'/></ViewFields><Query><Where><Geq><FieldRef Name='Modified'/><Value Type='DateTime'><Today/></Value></Geq></Where></Query></View>"
    $ListItems = Get-PnPListItem -List $DocLibraryName -FolderServerRelativeUrl $Folder1URL -Query $queryStr -Connection $SPconnection

    ForEach($Item in $ListItems)
    {
        #Add file details to Result array
        $AllFiles += New-Object PSObject -property $([ordered]@{ 
            FileName  = $Item.FieldValues["FileLeafRef"]            
            FileID = $Item.FieldValues["UniqueId"]
            FileType = $Item.FieldValues["File_x0020_Type"]
            RelativeURL = $Item.FieldValues["FileRef"]
            CreatedByEmail = $Item.FieldValues["Author"].Email
            CreatedTime   = $Item.FieldValues["Created"]
            LastModifiedTime   = $Item.FieldValues["Modified"]
            ModifiedByEmail  = $Item.FieldValues["Editor"].Email
            FileSize_KB = [Math]::Round(($Item.FieldValues["File_x0020_Size"]/1024), 2) #File size in KB
        })
    }

$AllFiles[0].FileName
$SPconnection.Url

}
