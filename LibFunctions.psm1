#===========================================================================================================
function SetupModules {
    if (!(Get-PSRepository -Name local)) {
        $parameters = @{
            Name = "local"
            SourceLocation = "F:/Powershell/PS-Module/"
            PublishLocation = "F:/Powershell/PS-Module/"
            InstallationPolicy = "Trusted"
          }
          Register-PSRepository @parameters

    }

    <# if (Get-Module -ListAvailable -Name PnP.PowerShell) {
        #Write-Output "Module PnP.PowerShell exists"
        Import-Module -name PnP.PowerShell -DisableNameChecking 
    } 
    else {
        Write-Output "Module PnP.PowerShell does not exist. Installing it"
        Install-Module -name PnP.PowerShell -Scope CurrentUser -AllowClobber -Repository local
        Import-Module -name PnP.PowerShell -DisableNameChecking 
    }#>
    if (Get-Module -ListAvailable -Name SharePointPnPPowerShell2016) {
        #Write-Output "Module SharePointPnPPowerShell2016 exists"
        Import-Module -name SharePointPnPPowerShell2016 -DisableNameChecking 
    } 
    else {
        Write-Output "Module SharePointPnPPowerShell2016 does not exist. Installing it"
        Install-Module -name SharePointPnPPowerShell2016 -Scope CurrentUser -AllowClobber -Repository local
        Import-Module -name SharePointPnPPowerShell2016 -DisableNameChecking 
    }
}
function setStaticVars {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $arr = ($PWD.Drive.CurrentLocation).Split("\") # Get current folder name
    switch ($arr[-1]) {
        "UAT" { $configFile = "LibFunctionsUAT.conf" }
        "Dev" { $configFile = "LibFunctionsDev.conf" }
        "Prod" { $configFile = "LibFunctions.conf" }
        default { $configFile = "LibFunctionsDev.conf" }
    }
    
    Foreach ($i in $(Get-Content $configFile)){
        $varName = $i.split("=")[0].trim()
        $varValue = $i.split("=",2)[1].trim()
        #write-host "varName = " $varName
        #write-host "varValue = " $varValue
        Set-Variable -Name $varName -Value $varValue -Scope global
    }
    
    #$global:SPconnection = Connect-PnPOnline -Url $SPhost -Credentials:SP-PnP-Connect -ReturnConnection
    $global:SPconnection = Connect-PnPOnline -Url $SPhost -Credentials:PPM-API-Sync.S -ReturnConnection
    $global:startDateTime = (Get-Date).ToUniversalTime()
    
    #$global:arrProjectFields = @("D_x002d_Ministry","D_x002d_Sector_x0020_Name","D_x002d_Short_x0020_Description","D_x002d_Assigned_x0020_To","D_x002d_Intake_x0020_Number","D_x002d_Intake_x0020_OIIE","D_x002d_Project_x0020_Informatio","D_x002d_Intake_x0020_Funded","D_x002d_Project_x0020_Code","D_x002d_Project_x0020_Informatio0","D_x002d_Funding_x0020_Type","D_x002d_Funding_x0020_Source","D_x002d_Resourcing","D_x002d_Project_x0020_Status","D_x002d_Short_x0020_Status","D_x002d_Scheduled_x0020_Finish","D_x002d_Project_x0020_Site","D_x002d_2nd_x0020_Project_x0020_","D_x002d_Updated_x0020_Budget_x00","D_x002d_Request_x0020_Type","D_x002d_Project_x0020_Meetings","D_x002d_Project_x0020_Structure_","D_x002d_Location","D_x002d_Scope","D_x002d_Schedule","D_x002d_Budget","D_x002d_Start_x0020_Date","D_x002d_Completion_x0020__x0025_","D_x002d_Estimated_x0020_Finish_x","D_x002d_Scheduled_x0020_Finish_x","D_x002d_Status_x0020_Summary","D_x002d_2nd_x0020_Project_x0020_0","D_x002d_Project_x0020_Health_x00","D_x002d_SOW_x0020__x0023_","D_x002d_Corrective_x0020_Actions","D_x002d_Status_x0020_Date","D_x002d_Total_x0020_Projected_x0","D_x002d_SDE_x0020_Intake","D_x002d_SDE_x0020_Intake_x003a_P","D_x002d_Purposed_x0020_Action_x0","D_x002d_Final_x0020_Action_x0020","D_x002d_Primary_x0020_Contact","D_x002d_Project_x0020_Status_x00","D_x002d_Short_x0020_Status0","D_x002d_Short_x0020_Status1","D_x002d_Status_x0020_History","Project_x0020_Server_x0020_Statu","Entry_x0020_updated_x0020_date","D_x002d_Actual_x0020_Costs","Project_x0020_Completed_x0020_or","D_x002d_Cost_x002d_Variance_x002","Managed_x0020_By_x0020_Org_x002f","D_x002d_Intake_x0020_Reference_x","D_x002d_Intake_x0020_Funded_x002","D_x002d_Intake_x0020_Project_x00","Intake_x0020_Delivery_x0020_Alig","D_x002d_Project_x0020_Number","D_x002d_Demand_x0020_Number","Intake_x0020_Delivery_x0020_Alig0","I_x002d_Intake_x0023_","D_x002d_Intake_x0020__x0023__x00","I_x002d_Intake_x0023__x003a_I_x0","I_x002d_Intake_x0023__x003a_I_x00","I_x002d_Intake_x0023__x003a_I_x01","I_x002d_Intake_x0023__x003a_I_x02","I_x002d_Intake_x0023__x003a_I_x03","D_x002d_Overall_x0020_Project_x0","ID","Modified","Created","Author","Editor","GUID","Last_x0020_Modified","Created_x0020_Date","EncodedAbsUrl")
}
#===========================================================================================================
function DisconnectSP {
    Disconnect-PnPOnline -Connection $global:SPconnection
}
#===========================================================================================================
Function RotateLog([String]$Source, [String]$DestinationPath,[String]$RententionInDays){
    $Format = '{0}_[{1}]{2}'
    $RententionInDays = $RententionInDays -as [int]

    If(Test-Path -LiteralPath $Source){
     $S = Get-Item -LiteralPath $Source
     $Destination = Join-Path -Path $DestinationPath -ChildPath ($Format -F $S.BaseName, ((Get-Date).ToString($DateTimeFormat)), $S.Extension)
    
     If(!(Test-Path -LiteralPath $DestinationPath)){$Null = New-Item -Path $DestinationPath -Type Directory -Force}
     Copy-Item -Path $Source -Destination $Destination -Force
     Clear-Content -LiteralPath $Source -Force
   
     Get-ChildItem -LiteralPath $DestinationPath -File -Filter ($Format -F $S.BaseName, '*',$S.Extension) | Where-Object LastWriteTime -le ((Get-Date).AddDays(-$RententionInDays)) | Remove-Item -ErrorAction SilentlyContinue
    }
}
#===========================================================================================================
function WriteLog ([string]$LogFile, [string]$LogString) {
    #Param ([string]$LogString)
    if (!(Test-Path -Path ".\log")) { New-Item -Path ".\" -Name "log" -ItemType "directory"}
    #$LogFile = ".\log\$(Get-Content env:computername).log"
    $DateTime = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
    $LogMessage = "$Datetime $LogString"
    Add-content -Path $LogFile -value $LogMessage
}
#===========================================================================================================
function currencyConvertToNumber ($currency) {
    $currency | 
    ForEach-Object{

        $_ -replace'[^0-9.]'
    }
    return $currency
}
#===========================================================================================================
function Write-DebugInfo {
    Write-Output " ************* Start Debug Info **************"
    Write-Output "SPhost " $SPhost
    Write-Output "SPconnection " ($SPconnection | ConvertTo-Json)
    Write-Output "SPListProjects " $SPlistProjects
    Write-Output "startDateTime " $startDateTime
}
#===========================================================================================================
function IgnoreSSLcert {
# BEGIN: SSL issues ignored
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
if (-not ([System.Management.Automation.PSTypeName]'ServerCertificateValidationCallback').Type)
{
$certCallback = @"
    using System;
    using System.Net;
    using System.Net.Security;
    using System.Security.Cryptography.X509Certificates;
    public class ServerCertificateValidationCallback
    {
        public static void Ignore()
        {
            if(ServicePointManager.ServerCertificateValidationCallback ==null)
            {
                ServicePointManager.ServerCertificateValidationCallback += 
                    delegate
                    (
                        Object obj, 
                        X509Certificate certificate, 
                        X509Chain chain, 
                        SslPolicyErrors errors
                    )
                    {
                        return true;
                    };
            }
        }
    }
"@
    Add-Type $certCallback
    }
[ServerCertificateValidationCallback]::Ignore()
# END: SSL issues
}
#===========================================================================================================
function SPgetLastUpdateRunDate {
    #$strQuery = "<ViewFields><FieldRef Name='ID' /><FieldRef Name='LastUpdateDateTime' /></ViewFields><OrderBy><FieldRef Name='ID' /></OrderBy>"
    $strQuery = "<ViewFields><FieldRef Name='LastUpdateDateTime' /></ViewFields><OrderBy><FieldRef Name='ID' /></OrderBy>"
    
    $items = Get-PnPListItem -Connection $SPconnection -List $SPlistLastRunDate -Query $strQuery | Select-Object -Last 1
    <#foreach ($item in $items) {
       # Write-Output "function getLastUpdateRunDate found ID: " $item["ID"] ". Title: " $item["Title"] "- Date: " $item["LastUpdateDateTime"]
        $DateTime = [datetime]$item["LastUpdateDateTime"]
        $DateTime.ToUniversalTime()
        #MA return [string] $DateTime.ToUniversalTime()
        return $DateTime.ToUniversalTime()
    } #>
    $DateTime = [datetime]$items["LastUpdateDateTime"][0]
    return $DateTime.ToUniversalTime() 
}
#===========================================================================================================
function SPgetLastUpdateRunInDateLocal {
    #$strQuery = "<ViewFields><FieldRef Name='ID' /><FieldRef Name='LastUpdateDateTime' /></ViewFields><OrderBy><FieldRef Name='ID' /></OrderBy>"
    $strQuery = "<ViewFields><FieldRef Name='LastUpdateDateTime' /></ViewFields><OrderBy><FieldRef Name='ID' /></OrderBy>"
    
    $items = Get-PnPListItem -Connection $SPconnection -List $SPlistLastRunDate -Query $strQuery | Select-Object -Last 1

    return  $items
    #$DateTime = [datetime]$items["LastUpdateDateTime"][0]
    # return $DateTime.ToUniversalTime()
}
#===========================================================================================================
function SPgetLastUpdateRunOutDate {
    #$strQuery = "<ViewFields><FieldRef Name='ID' /><FieldRef Name='LastUpdateDateTime' /></ViewFields><OrderBy><FieldRef Name='ID' /></OrderBy>"
    $strQuery = "<ViewFields><FieldRef Name='LastUpdateDateTimeOut' /></ViewFields><OrderBy><FieldRef Name='ID' /></OrderBy>"
    
    $items = Get-PnPListItem -Connection $SPconnection -List $SPlistLastRunDate -Query $strQuery | Select-Object -Last 1

    return  $items
}
#===========================================================================================================
function SPgetLastUpdateRunDateID {
    $strQuery = "<ViewFields><FieldRef Name='ID' /><FieldRef Name='LastUpdateDateTime' /></ViewFields><OrderBy><FieldRef Name='ID' /></OrderBy>"
    
    $items = Get-PnPListItem -Connection $SPconnection -List $SPlistLastRunDate -Query $strQuery | Select-Object -Last 1

    return $items
    <#foreach ($item in $items) {
        Write-Output "function getLastUpdateRunDate found ID: " $item["ID"] ". Title: " $item["Title"] "- Date: " $item["LastUpdateDateTime"]
        #$DateTime = [datetime]$item["LastUpdateDateTime"]
        #$DateTime.ToUniversalTime()
        #return [string]$DateTime.ToUniversalTime()
        return $item["ID"].ToString()
    }#>
}
#===========================================================================================================
function SPsetLastUpdateRunDateBy ($strTitle, $dateLastUpdateDateTime, $InOut) {
    If ([string]::IsNullOrEmpty($strTitle) -or [string]::IsNullOrEmpty($dateLastUpdateDateTime)) {
        Write-Output "function SPsetLastUpdateRunDateByID expects parameter strTitle and dateLastUpdateDateTime. Script execution will be terminated."
        return $false
    }

    $strID = SPgetLastUpdateRunDateID
    
    $result = Set-PnPListItem -Connection $SPconnection -List $SPlistLastRunDate -Identity $strID -values @{"Title" = $strTitle; $InOut = $dateLastUpdateDateTime}

    return !([string]::IsNullOrEmpty($result))
}
#===========================================================================================================
function SPgetProjectByProjectNumber($strProjectNumber) {
    If ([string]::IsNullOrWhiteSpace($strProjectNumber)) {
        Write-Output "function SPgetProjectByProjectNumber expects parameter strProjectNumber. Script execution will be terminated."
        
    }
    $strQuery = "<View><Query><Where><Eq><FieldRef Name='D_x002d_Project_x0020_Number' /><Value Type='Text'>"+$strProjectNumber+"</Value></Eq></Where></Query></View>"
    #$strQuery = "<View><Query><Where><Eq><FieldRef Name='D_x002d_Project_x0020_Number'/><Value Type='Text'>PRJ0010002</Value></Eq></Where></Query></View>"


   # $items = Get-PnPListItem -Connection $SPconnection -List $SPlistProjects -Query $strQuery
   $items = Get-PnPListItem -Connection $SPconnection -List $SPlistProjects -Query $strQuery

    #Write-Output GetType($items)
    return $items

}
#===========================================================================================================
function SPgetFilteredProjects {

    $SPLastRunOutDate = SPgetLastUpdateRunOutDate
 

    $strQuery ="<View><Query><Where><And><Gt><FieldRef Name='Modified' /><Value IncludeTimeValue='TRUE' Type='DateTime' StorageTZ='TRUE'>"+ $SPLastRunOutDate.FieldValues.LastUpdateDateTimeOut.ToString("yyyy/MM/dd hh:mm:ss tt")+"</Value></Gt><IsNotNull><FieldRef Name='D_x002d_Project_x0020_Number' /></IsNotNull></And></Where></Query></View>"
    #$strQuery ="<View><Query><Where><And><Gt><FieldRef Name='Modified' /><Value IncludeTimeValue='TRUE' Type='DateTime' StorageTZ='TRUE'>2021-06-10 4:1</Value></Gt><IsNotNull><FieldRef Name='D_x002d_Project_x0020_Number' /></IsNotNull></And></Where></Query></View>"

    $items = Get-PnPListItem -Connection $SPconnection -List $SPlistProjects -Query $strQuery


    #$message = "SharePoint Items Collection Count:" + $items.Count
    #WriteLog $message

    return $items
}
#===========================================================================================================
function SPsetProjectByProjectNumberLocal($project, $strColumn, $projObj) {

  
    $strQuery = "<View><Query><Where><Eq><FieldRef Name='D_x002d_Project_x0020_Number' /><Value Type='Text'>"+$project.projectNumber+"</Value></Eq></Where></Query></View>"        
    
    $items = Get-PnPListItem -Connection $SPconnection -List $SPlistProjects -Query $strQuery

    if ($strColumn -eq "projectName"){
            try{
                    Set-PnPListItem -Connection $SPconnection -List $SPlistProjects -Identity $items -Values @{"D_x002d_Intake_x0020_Project_x00" = $project.projectName;  "IN" = "Success"; "inExcept" = ""; "sys_itemSyncDateIn" = (Get-Date).ToUniversalTime()}
                    $message ="Success, Project Number:" + $project.projectNumber + "Project Name: from" + $projObj.FieldValues.D_x002d_Intake_x0020_Project_x00 + " to " + $project.projectName
           
            }
            catch {    

                    Get-PnPException
                    $message = "Fail, Project Number:" + $project.projectNumber +  "StatusCode:" + $_.Exception.Response.StatusCode.value__ + "StatusDescription:" + $_.Exception.Response.StatusDescription + "Exception Message: " + $_.Exception.Message
                    Set-PnPListItem -Connection $SPconnection -List $SPlistProjects -Identity $items -Values @{"IN" = "Fail"; "inExcept" = $message}             
                     }

                     WriteLog $message
            
    }
    if ($strColumn -eq "primaryContact") {

            try{
                    Set-PnPListItem -Connection $SPconnection -List $SPlistProjects -Identity $items -Values @{"D_x002d_Primary_x0020_Contact" = $project.primaryContact; "IN" = "Success"; "inExcept" = ""; "sys_itemSyncDateIn" = (Get-Date).ToUniversalTime()}             
                    $message ="Success, Project Number:" + $project.projectNumber + "Primary Contact: from "+ $projObj.FieldValues.D_x002d_Primary_x0020_Contact.LookupValue + " to " + $project.primaryContact 
                                        
            }
            catch {

                    Get-PnPException
                    $message = "Fail, Project Number:" + $project.projectNumber +  "StatusCode:" + $_.Exception.Response.StatusCode.value__ + "StatusDescription:" + $_.Exception.Response.StatusDescription + "Exception Message: " + $_.Exception.Message
                    Set-PnPListItem -Connection $SPconnection -List $SPlistProjects -Identity $items -Values @{"IN" = "Fail"; "inExcept" = $message}             

                    }
                    WriteLog $message  
    }     
    if ($strColumn -eq "Demand") {

            try{
                    Set-PnPListItem -Connection $SPconnection -List $SPlistProjects -Identity $items -Values @{"D_x002d_Demand_x0020_Number" = $project.intakeNumber; "IN" = "Success"; "inExcept" = ""; "sys_itemSyncDateIn" = (Get-Date).ToUniversalTime()}             
                    $message ="Success, Project Number:" + $project.projectNumber + "Demand Number: from "+ $projObj.FieldValues.D_x002d_Demand_x0020_Number + " to " + $project.intakeNumber 
            }
            catch{
                    Get-PnPException
                    $message = "Fail, Project Number:" + $project.projectNumber +  "StatusCode:" + $_.Exception.Response.StatusCode.value__ + "StatusDescription:" + $_.Exception.Response.StatusDescription + "Exception Message: " + $_.Exception.Message
                    Set-PnPListItem -Connection $SPconnection -List $SPlistProjects -Identity $items -Values @{"IN" = "Fail"; "inExcept" = $message}             
                    
                    }
                    WriteLog $message              


    }   
    if ($strColumn -eq "DemandName") {

            try{
                    Set-PnPListItem -Connection $SPconnection -List $SPlistProjects -Identity $items -Values @{"D_x002d_Demand_x0020_Name" = $project.intakeName; "IN" = "Success"; "inExcept" = ""; "sys_itemSyncDateIn" = (Get-Date).ToUniversalTime()}             
                    $message ="Success, Project Number:" + $project.projectNumber + "Demand Name: from "+ $projObj.FieldValues.D_x002d_Demand_x0020_Name + " to " + $project.intakeName 
            }
            catch{
                    Get-PnPException
                    $message = "Fail, Project Number:" + $project.projectNumber +  "StatusCode:" + $_.Exception.Response.StatusCode.value__ + "StatusDescription:" + $_.Exception.Response.StatusDescription + "Exception Message: " + $_.Exception.Message
                    Set-PnPListItem -Connection $SPconnection -List $SPlistProjects -Identity $items -Values @{"IN" = "Fail"; "inExcept" = $message}
                    }
                    WriteLog $message    

    }   
    
    if ($strColumn -eq "DemandMinistry") {

            try{
                    Set-PnPListItem -Connection $SPconnection -List $SPlistProjects -Identity $items -Values @{"D_x002d_Demand_x0020_Ministry" = $project.intakeMinistry; "IN" = "Success"; "inExcept" = ""; "sys_itemSyncDateIn" = (Get-Date).ToUniversalTime()}             
                    $message ="Success, Project Number:" + $project.projectNumber + "Demand Ministry: from "+ $projObj.FieldValues.D_x002d_Demand_x0020_Ministry + " to " + $project.intakeMinistry 
            }
            catch{
                    Get-PnPException
                    $message = "Fail, Project Number:" + $project.projectNumber +  "StatusCode:" + $_.Exception.Response.StatusCode.value__ + "StatusDescription:" + $_.Exception.Response.StatusDescription + "Exception Message: " + $_.Exception.Message

                    Set-PnPListItem -Connection $SPconnection -List $SPlistProjects -Identity $items -Values @{"IN" = "Fail"; "inExcept" = $message}

                    }
                    WriteLog $message    

    }

    if ($strColumn -eq "LeadSector") {

            try{
                    Set-PnPListItem -Connection $SPconnection -List $SPlistProjects -Identity $items -Values @{"D_x002d_Lead_x0020_Sector" = $project.leadSector; "IN" = "Success"; "inExcept" = ""; "sys_itemSyncDateIn" = (Get-Date).ToUniversalTime()}             
                    $message ="Success, Project Number:" + $project.projectNumber + "Lead Sector: from "+ $projObj.FieldValues.D_x002d_Lead_x0020_Sector + " to " + $project.leadSector 
            
            }
            catch{

                    Get-PnPException
                    $message = "Fail, Project Number:" + $project.projectNumber +  "StatusCode:" + $_.Exception.Response.StatusCode.value__ + "StatusDescription:" + $_.Exception.Response.StatusDescription + "Exception Message: " + $_.Exception.Message

                    Set-PnPListItem -Connection $SPconnection -List $SPlistProjects -Identity $items -Values @{"IN" = "Fail"; "inExcept" = $message}             

                    }
                    WriteLog $message 

    }

    if ($strColumn -eq "TotalBudget") {

        try{
                Set-PnPListItem -Connection $SPconnection -List $SPlistProjects -Identity $items -Values @{"D_x002d_Updated_x0020_Budget_x00" = $project.totalApprovedBudget; "IN" = "Success"; "inExcept" = ""; "sys_itemSyncDateIn" = (Get-Date).ToUniversalTime()}             
                $message ="Success, Project Number:" + $project.projectNumber + "Total Approved Budget: from "+ $projObj.FieldValues.D_x002d_Updated_x0020_Budget_x00 + " to " + $project.totalApprovedBudget 
        
        }
        catch{

                Get-PnPException
                $message = "Fail, Project Number:" + $project.projectNumber +  "StatusCode:" + $_.Exception.Response.StatusCode.value__ + "StatusDescription:" + $_.Exception.Response.StatusDescription + "Exception Message: " + $_.Exception.Message

                Set-PnPListItem -Connection $SPconnection -List $SPlistProjects -Identity $items -Values @{"IN" = "Fail"; "inExcept" = $message}             

                }
                WriteLog $message 

}
                        
}
#===========================================================================================================
function SNgetAllProjects($lastRunDateTime) {
    # the URL is case sensitive
    
        If ([string]::IsNullOrWhiteSpace($lastRunDateTime)) {
            Write-Output "function SNgetAllProjects expects parameter lastRunDateTime. Function will return null response."
            return $jsonResponse
        }
        else {
            #$message = "function getAllProjects received parameter lastRunDateTime "+$lastRunDateTime
            #WriteLog $message
            
            #$RESThostURI = "https://servicenowconnector-pgs-dev.os99.gov.ab.ca/servicenowconnector/v1/projects?lastRunDateTime=" + [string]$lastRunDateTime #DEV
            #$RESThostURI = "https://servicenowconnector-pgs-uat.os99.gov.ab.ca/servicenowconnector/v1/projects?lastRunDateTime=" + [string]$lastRunDateTime #UAT

            $encodedLastRunDate = [System.Web.HttpUtility]::UrlEncode($lastRunDateTime)

           # $RESThostURI = $RESThost + "/servicenowconnector/v1/projects?lastRunDateTime=" + [string]$lastRunDateTime
           $RESThostURI = $RESThost + "/servicenowconnector/v1/projects?lastRunDateTime=" + $encodedLastRunDate

            $RESTuri = [uri]$RESThostURI
            # Write-Output "Host: " $RESTuri.host
    
            #if ($RESTuri.host -eq "servicenowconnector-pgs-dev.os99.gov.ab.ca") {IgnoreSSLcert}#DEV
            if ($RESTuri.host -eq "servicenowconnector-pgs-uat.os99.gov.ab.ca") {IgnoreSSLcert}#UAT

            $Params = @{
                "URI"     = $RESThostURI
                "Method"  = 'GET'
                "Headers" = @{
                    "Content-Type"  = 'application/json'
                    "X-API-Authentication" = $strJWTToken
                    "Authorization" = $strJWTToken
                }
            }
            try {
                $jsonResponse = Invoke-RestMethod @Params
            } catch {


                $message = "Fail, Project Number:" + $strProjectNumber +  "StatusCode:" + $_.Exception.Response.StatusCode.value__ + "StatusDescription:" + $_.Exception.Response.StatusDescription + "Exception Message: " + $_.Exception.Message
                WriteLog $message
                $message = $Params | ConvertTo-Json 
                WriteLog $message
                # SPsetsys_OutitemSyncInforbyProjectNumber $strProjectNumber "Fail" "Update Project Status" $message     
            }
            return $jsonResponse
        }
    }
#===========================================================================================================
function SPsetProjectByProjectNumber($projectNumber) {
    $strQuery = "<View><Query><Where><Eq><FieldRef Name='D_x002d_Project_x0020_Number' /><Value Type='Text'>" + $project.projectNumber + "</Value></Eq></Where></Query></View>"        
        
    $items = Get-PnPListItem -Connection $SPconnection -List $SPlistProjects -Query $strQuery

    if ($project[1] -eq "projectName") {
        try {
            Set-PnPListItem -Connection $SPconnection -List $SPlistProjects -Identity $items -Values @{"D_x002d_Intake_x0020_Project_x00" = $project.projectName }
        }
        catch {    
            Write-Host $_.Exception.Message`n
        }
            
    }
    if ($project[1] -eq "primaryContact") {

        try {
            Set-PnPListItem -Connection $SPconnection -List $SPlistProjects -Identity $items -Values @{"D_x002d_Primary_x0020_Contact" = $project.primaryContact }             
        }
        catch {
            Write-Host $_.Exception.Message`n
        }
    }              
}
#===========================================================================================================
function SPsetsys_OutitemSyncInforbyProjectNumber($projectNumber,$success,$title,$errorMessage) {
    $strQuery = "<View><Query><Where><Eq><FieldRef Name='D_x002d_Project_x0020_Number' /><Value Type='Text'>" + $projectNumber + "</Value></Eq></Where></Query></View>"        
    $items = Get-PnPListItem -Connection $SPconnection -List $SPlistProjects -Query $strQuery

if ($success -eq "Success" -AND $global:OldStatus -ne "Fail")
{
            Set-PnPListItem -Connection $SPconnection -List $SPlistProjects -Identity $items -Values @{"sys_itemSyncDateOut" = (Get-Date).ToUniversalTime(); "OUT" = "Success"; "outExcept" = ""}           
            $message = "Success, " + $title + " Project Number:" + $projectNumber + " sys_itemSyncDateOut:" + $startDateTime
            $global:OldStatus = "Success"
            WriteLog $message
        
}elseif ($success -eq "Fail") {


            $message = $title + $errorMessage
            Set-PnPListItem -Connection $SPconnection -List $SPlistProjects -Identity $items -Values @{"OUT" = "Fail"; "outExcept" = $message + "`n" + $global:Oldmessage}          
            $global:Oldmessage = $message # save the error message, will be combine with Project status message
            $global:OldStatus = "Fail"
            WriteLog $message

}           
}
#===========================================================================================================

