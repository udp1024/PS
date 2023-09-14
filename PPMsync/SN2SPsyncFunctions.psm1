#===========================================================================================================
function setStaticVars {
#===================================DEV Connection========================================================================
    #$global:strJWTToken = "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIxMjM0NTY3ODkwIiwibmFtZSI6IkpvaG4gRG9lIiwiaWF0IjoxNjE3MjM5MDIyLCJzdWJzY3JpcHRpb25JZCI6InN1YjEifQ.gfZXp51gkGNL2iHeL6-efOAGJjd-2qaUxQode4NPMbs"#DEV
    #$global:Token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIxMjM0NTY3ODkwIiwibmFtZSI6IkpvaG4gRG9lIiwiaWF0IjoxNjE3MjM5MDIyLCJzdWJzY3JpcHRpb25JZCI6InN1YjEifQ.gfZXp51gkGNL2iHeL6-efOAGJjd-2qaUxQode4NPMbs"#DEV
    #$global:RESThost = "https://servicenowconnector-pgs-dev.os99.gov.ab.ca" #DEV
    #$global:SPhost = "https://cdmsdev.sp.gov.ab.ca/Delivery" #DEV
#===========================================================================================================

#====================================UAT Connection=======================================================================
    #$global:strJWTToken = "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIxMjM0NTY3ODkwIiwibmFtZSI6IkpvaG4gRG9lIiwiaWF0IjoxNjE3MjM5MDIyLCJzdWJzY3JpcHRpb25JZCI6InN1YjEifQ.gfZXp51gkGNL2iHeL6-efOAGJjd-2qaUxQode4NPMbs"#UAT
    #$global:Token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIxMjM0NTY3ODkwIiwibmFtZSI6IkpvaG4gRG9lIiwiaWF0IjoxNjE3MjM5MDIyLCJzdWJzY3JpcHRpb25JZCI6InN1YjEifQ.gfZXp51gkGNL2iHeL6-efOAGJjd-2qaUxQode4NPMbs"#UAT
    #$global:RESThost = "https://servicenowconnector-pgs-uat.os99.gov.ab.ca/" #UAT
    #$global:SPhost = "https://cdmsuat.sp.gov.ab.ca/Delivery" #UAT
#===========================================================================================================

#$global:SPlistProjects = "Delivery"
#$global:SPlistIntake = "Intake"
#$global:SPlistLastRunDate = "SN2SPlastRunDate"

    if (Get-Module -ListAvailable -Name SharePointPnPPowerShell2016) {
        #Write-Output "Module SharePointPnPPowerShell2016 exists"
        Import-Module -name SharePointPnPPowerShell2016 -DisableNameChecking 
    } 
    else {
        Write-Output "Module SharePointPnPPowerShell2016 does not exist. Installing it"
        Install-Module -name SharePointPnPPowerShell2016 -Scope CurrentUser -AllowClobber
        Import-Module -name SharePointPnPPowerShell2016 -DisableNameChecking 
    }
<#
    if (Get-Module -ListAvailable -Name PnP.PowerShell) {
        #Write-Output "Module PnP.PowerShell exists"
        Import-Module -name PnP.PowerShell -DisableNameChecking 
    } 
    else {
        Write-Output "Module PnP.PowerShell does not exist. Installing it"
        Install-Module -name PnP.PowerShell -Scope CurrentUser -AllowClobber
        Import-Module -name PnP.PowerShell -DisableNameChecking 
    }
#>
    $arr = ($PWD.Drive.CurrentLocation).Split("\")
    switch ($arr[-1]) {
        "UAT" { $configFile = "SN2SPsyncFunctionsUAT.conf" }
        "Dev" { $configFile = "SN2SPsyncFunctionsDEV.conf" }
        "PPMsync" { $configFile = "SN2SPsyncFunctions.conf" }
        default { $configFile = "SN2SPsyncFunctionsDEV.conf" }
    }

    Foreach ($i in $(Get-Content $configFile)){
        $varName = $i.split("=")[0].trim()
        $varValue = $i.split("=",2)[1].trim()
        #write-host "varName = " $varName
        #write-host "varValue = " $varValue
        Set-Variable -Name $varName -Value $varValue -Scope global
    }

    #$global:SPconnection = Connect-PnPOnline -Url $SPhost -Credentials:SP-PnP-Connect -ReturnConnection
    $global:SPconnection = Connect-PnPOnline -Url $SPhost -Credentials:SP-PnP-Connect -ReturnConnection #-TransformationOnPrem 
    $global:startDateTime = (Get-Date).ToUniversalTime()
    $global:arrProjectFields = @("D_x002d_Ministry","D_x002d_Sector_x0020_Name","D_x002d_Short_x0020_Description","D_x002d_Assigned_x0020_To","D_x002d_Intake_x0020_Number","D_x002d_Intake_x0020_OIIE","D_x002d_Project_x0020_Informatio","D_x002d_Intake_x0020_Funded","D_x002d_Project_x0020_Code","D_x002d_Project_x0020_Informatio0","D_x002d_Funding_x0020_Type","D_x002d_Funding_x0020_Source","D_x002d_Resourcing","D_x002d_Project_x0020_Status","D_x002d_Short_x0020_Status","D_x002d_Scheduled_x0020_Finish","D_x002d_Project_x0020_Site","D_x002d_2nd_x0020_Project_x0020_","D_x002d_Updated_x0020_Budget_x00","D_x002d_Request_x0020_Type","D_x002d_Project_x0020_Meetings","D_x002d_Project_x0020_Structure_","D_x002d_Location","D_x002d_Scope","D_x002d_Schedule","D_x002d_Budget","D_x002d_Start_x0020_Date","D_x002d_Completion_x0020__x0025_","D_x002d_Estimated_x0020_Finish_x","D_x002d_Scheduled_x0020_Finish_x","D_x002d_Status_x0020_Summary","D_x002d_2nd_x0020_Project_x0020_0","D_x002d_Project_x0020_Health_x00","D_x002d_SOW_x0020__x0023_","D_x002d_Corrective_x0020_Actions","D_x002d_Status_x0020_Date","D_x002d_Total_x0020_Projected_x0","D_x002d_SDE_x0020_Intake","D_x002d_SDE_x0020_Intake_x003a_P","D_x002d_Purposed_x0020_Action_x0","D_x002d_Final_x0020_Action_x0020","D_x002d_Primary_x0020_Contact","D_x002d_Project_x0020_Status_x00","D_x002d_Short_x0020_Status0","D_x002d_Short_x0020_Status1","D_x002d_Status_x0020_History","Project_x0020_Server_x0020_Statu","Entry_x0020_updated_x0020_date","D_x002d_Actual_x0020_Costs","Project_x0020_Completed_x0020_or","D_x002d_Cost_x002d_Variance_x002","Managed_x0020_By_x0020_Org_x002f","D_x002d_Intake_x0020_Reference_x","D_x002d_Intake_x0020_Funded_x002","D_x002d_Intake_x0020_Project_x00","Intake_x0020_Delivery_x0020_Alig","D_x002d_Project_x0020_Number","D_x002d_Demand_x0020_Number","Intake_x0020_Delivery_x0020_Alig0","I_x002d_Intake_x0023_","D_x002d_Intake_x0020__x0023__x00","I_x002d_Intake_x0023__x003a_I_x0","I_x002d_Intake_x0023__x003a_I_x00","I_x002d_Intake_x0023__x003a_I_x01","I_x002d_Intake_x0023__x003a_I_x02","I_x002d_Intake_x0023__x003a_I_x03","D_x002d_Overall_x0020_Project_x0","ID","Modified","Created","Author","Editor","GUID","Last_x0020_Modified","Created_x0020_Date","EncodedAbsUrl")

    $global:Oldmessage = "" #to track if their is message from UpdateProjectDetails function, so message can combine with insert/updateProjectstatus function
    $global:OldStatus = "" #to track if their is status from last updateProjectDetails function

}
#===========================================================================================================
function initializeSN2SPsync {

    if (Get-Module -ListAvailable -Name JWTDetails) {
        #Write-Output "Module JWTDetails exists"
        Import-Module -name JWTDetails
    } 
    else {
        Write-Output "Module JWTDetails does not exist. Installing it"
        Install-Module -name JWTDetails -Scope CurrentUser
    }
    #Get-JWTDetails($strJWTToken)

    $userName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    <#If ($userName -eq "GOA\PPM-API-Sync.S" ){
            Write-Output "user context GOA\PPM-API-Sync.S"
        }
    else
        {#>
            Write-Output "user context $UserName"
        #}
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
   
     Get-ChildItem -LiteralPath $DestinationPath -File -Filter ($Format -F $S.BaseName, '*',$S.Extension) | ? LastWriteTime -le ((Get-Date).AddDays(-$RententionInDays)) | Remove-Item -ErrorAction SilentlyContinue
    }
}
#===========================================================================================================
function WriteLog
{
    Param ([string]$LogString)
    if (!(Test-Path -Path ".\log")) { New-Item -Path ".\" -Name "log" -ItemType "directory"}
    $LogFile = ".\log\$(gc env:computername).log"
    $DateTime = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
    $LogMessage = "$Datetime $LogString"
    Add-content $LogFile -value $LogMessage
}
#===========================================================================================================

function currencyConvertToNumber ($currency)
{
    $currency | 
    ForEach-Object{

        $_ -replace'[^0-9.]'
    }
    return $currency
}

#===========================================================================================================

function Write-DebugInfo {
    Write-Output " ************* Start Debug Info **************"
    Write-Output "JWTToken " $strJWTToken
    Write-Output "Token " $Token
    Write-Output "RESThost " $RESThost
    Write-Output "SPhost " $SPhost
    Write-Output "SPconnection " ($SPconnection | ConvertTo-Json)
    Write-Output "SPListProjects " $SPlistProjects
    Write-Output "SPListIntake " $SPlistIntake
    Write-Output "SPlistLastRunDate " $SPlistLastRunDate
    Write-Output "startDateTime " $startDateTime
}
#===========================================================================================================
function NewProjectObject {
    $dicProject = @{"Title"  =  "";
    "D_x002d_Ministry"  =  "";
    "D_x002d_Sector_x0020_Name"  =  "";
    "D_x002d_Short_x0020_Description"  =  "";
    "D_x002d_Assigned_x0020_To"  =  "";
    "D_x002d_Intake_x0020_Number"  =  "";
    "D_x002d_Intake_x0020_OIIE"  =  "";
    "D_x002d_Project_x0020_Informatio"  =  "";
    "D_x002d_Intake_x0020_Funded"  =  "";
    "D_x002d_Project_x0020_Code"  =  "";
    "D_x002d_Project_x0020_Informatio0"  =  "";
    "D_x002d_Funding_x0020_Type"  =  "";
    "D_x002d_Funding_x0020_Source"  =  "";
    "D_x002d_Resourcing"  =  "";
    "D_x002d_Project_x0020_Status"  =  "";
    "D_x002d_Short_x0020_Status"  =  "";
    "D_x002d_Scheduled_x0020_Finish"  =  "";
    "D_x002d_Project_x0020_Site"  =  "";
    "D_x002d_2nd_x0020_Project_x0020_"  =  "";
    "D_x002d_Updated_x0020_Budget_x00"  =  "";
    "D_x002d_Request_x0020_Type"  =  "";
    "D_x002d_Project_x0020_Meetings"  =  "";
    "D_x002d_Project_x0020_Structure_"  =  "";
    "D_x002d_Location"  =  "";
    "D_x002d_Scope"  =  "";
    "D_x002d_Schedule"  =  "";
    "D_x002d_Budget"  =  "";
    "D_x002d_Start_x0020_Date"  =  "";
    "D_x002d_Completion_x0020__x0025_"  =  "";
    "D_x002d_Estimated_x0020_Finish_x"  =  "";
    "D_x002d_Scheduled_x0020_Finish_x"  =  "";
    "D_x002d_Status_x0020_Summary"  =  "";
    "D_x002d_2nd_x0020_Project_x0020_0"  =  "";
    "D_x002d_Project_x0020_Health_x00"  =  "";
    "D_x002d_SOW_x0020__x0023_"  =  "";
    "D_x002d_Corrective_x0020_Actions"  =  "";
    "D_x002d_Status_x0020_Date"  =  "";
    "D_x002d_Total_x0020_Projected_x0"  =  "";
    "D_x002d_SDE_x0020_Intake"  =  "";
    "D_x002d_SDE_x0020_Intake_x003a_P"  =  "";
    "D_x002d_Purposed_x0020_Action_x0"  =  "";
    "D_x002d_Final_x0020_Action_x0020"  =  "";
    "D_x002d_Primary_x0020_Contact"  =  "";
    "D_x002d_Project_x0020_Status_x00"  =  "";
    "D_x002d_Short_x0020_Status0"  =  "";
    "D_x002d_Short_x0020_Status1"  =  "";
    "D_x002d_Status_x0020_History"  =  "";
    "Project_x0020_Server_x0020_Statu"  =  "";
    "Entry_x0020_updated_x0020_date"  =  "";
    "D_x002d_Actual_x0020_Costs"  =  "";
    "Project_x0020_Completed_x0020_or"  =  "";
    "D_x002d_Cost_x002d_Variance_x002"  =  "";
    "Managed_x0020_By_x0020_Org_x002f"  =  "";
    "D_x002d_Intake_x0020_Reference_x"  =  "";
    "D_x002d_Intake_x0020_Funded_x002"  =  "";
    "D_x002d_Intake_x0020_Project_x00"  =  "";
    "Intake_x0020_Delivery_x0020_Alig"  =  "";
    "D_x002d_Project_x0020_Number"  =  "";
    "D_x002d_Demand_x0020_Number"  =  "";
    "Intake_x0020_Delivery_x0020_Alig0"  =  "";
    "I_x002d_Intake_x0023_"  =  "";
    "D_x002d_Intake_x0020__x0023__x00"  =  "";
    "I_x002d_Intake_x0023__x003a_I_x0"  =  "";
    "I_x002d_Intake_x0023__x003a_I_x00"  =  "";
    "I_x002d_Intake_x0023__x003a_I_x01"  =  "";
    "I_x002d_Intake_x0023__x003a_I_x02"  =  "";
    "I_x002d_Intake_x0023__x003a_I_x03"  =  "";
    "D_x002d_Overall_x0020_Project_x0"  =  "";
    "ID"  =  "";
    }

    return $dicProject
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
function SNupdateProjectStatus ($strProjectNumber,$Body) {
    #$RESThostURI = 'https://servicenowconnector-pgs-dev.os99.gov.ab.ca/servicenowconnector/v1/projects/'+$strProjectNumber+'/projectstatus' #DEV
    #$RESThostURI = 'https://servicenowconnector-pgs-uat.os99.gov.ab.ca/servicenowconnector/v1/projects/'+$strProjectNumber+'/projectstatus' #UAT
    $RESThostURI = $RESThost + '/servicenowconnector/v1/projects/'+ $strProjectNumber + '/projectstatus'
    
    $RESTuri = [uri]$RESThostURI

    # Params needs to be a Hashtable
    $Params = @{
        'URI'     = $RESThostURI
        'Method'  = 'PATCH'
        'ContentType'  = 'application/json'
    }

    # Headers needs tobe iDictionary
    $Headers = @{
        Authorization = $strJWTToken
    }

    try {
        $jsonResponse = Invoke-RestMethod @Params -Headers $Headers -Body $Body #-SkipCertificateCheck
        SPsetsys_OutitemSyncInforbyProjectNumber $strProjectNumber "Success" "Update Project Status" ""

    } catch {

        #Use Stream reader to read response body from restful service
        $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
        $ErrResp = $streamReader.ReadToEnd() | ConvertFrom-Json
        $streamReader.Close()

        foreach ($err in $ErrResp.errors) {
            $totalerr = $err.message + " | " + $totalerr
        }

        # Dig into the exception to get the Response details.
        # Note that value__ is not a typo.
        $message = "Fail, Project Number:" + $strProjectNumber+"Error message:"+$totalerr
        SPsetsys_OutitemSyncInforbyProjectNumber $strProjectNumber "Fail" "Update Project Status" $message
    }
    return $jsonResponse
}
#===========================================================================================================
function SNgetProjectByProjectNumber($strProjectNumber) {
    # the URL is case sensitive
    
        If ([string]::IsNullOrWhiteSpace($strProjectNumber)) {
            Write-Output "function SNgetProjectByProjectNumber expects parameter strProjectNumber. Function will return null response."
            return $jsonResponse
        }
        else {
            Write-Output "function SNgetProjectByProjectNumber received parameter lastRunDateTime " $strProjectNumber
            #$RESThostURI = "https://servicenowconnector-pgs-dev.os99.gov.ab.ca/servicenowconnector/v1/projects/" + [string]$strProjectNumber #DEV
            #$RESThostURI = "https://servicenowconnector-pgs-uat.os99.gov.ab.ca/servicenowconnector/v1/projects/" + [string]$strProjectNumber #UAT
            $RESThostURI = $RESThost + "/servicenowconnector/v1/projects/" + [string]$strProjectNumber

            $RESTuri = [uri]$RESThostURI
            # Write-Output "Host: " $RESTuri.host
    
            #if ($RESTuri.host -eq "servicenowconnector-pgs-dev.os99.gov.ab.ca") {IgnoreSSLcert} #DEV
            if ($RESTuri.host -eq "servicenowconnector-pgs-uat.os99.gov.ab.ca") {IgnoreSSLcert} #UAT

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
            }
            catch {
                $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
                $ErrResp = $streamReader.ReadToEnd() | ConvertFrom-Json
                $streamReader.Close()

                foreach ($err in $ErrResp.errors) {
                    $totalerr = $err.message + " | " + $totalerr
                }

                $message = "function SNgetProjectByProjectNumber:" + $strProjectNumber+"Error message:"+$totalerr
                WriteLog $message
                
            }
           
            return $jsonResponse
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
function SNinsertProjectStatus ($strProjectNumber,$Body) {
    #$RESThostURI = 'https://servicenowconnector-pgs-dev.os99.gov.ab.ca/servicenowconnector/v1/projects/'+$strProjectNumber+'/projectstatus' #DEV
    #$RESThostURI = 'https://servicenowconnector-pgs-uat.os99.gov.ab.ca/servicenowconnector/v1/projects/'+$strProjectNumber+'/projectstatus' #UAT
    $RESThostURI = $RESThost + '/servicenowconnector/v1/projects/'+$strProjectNumber+'/projectstatus'
    
    $RESTuri = [uri]$RESThostURI

    #if ($RESTuri.host -eq "servicenowconnector-pgs-dev.os99.gov.ab.ca") {IgnoreSSLcert}

    # Params needs to be a Hashtable
    $Params = @{
        'URI'     = $RESThostURI
        'Method'  = 'POST'
        'ContentType'  = 'application/json'
    }

    # Headers needs tobe iDictionary
    $Headers = @{
        Authorization = $strJWTToken
    }

    # body needs to be an Object
    #Write-Output $Body 


    #$body
    <#Write-Output "Params" $Params.GetType()
    Write-Output "Headers" $Headers.GetType()
    Write-Output "body" $body.GetType()
    Write-Host "Body " $body
    Write-Host "JWTtoken " $strJWTToken#>

    try {
        $jsonResponse = Invoke-RestMethod @Params -Headers $Headers -Body $Body #-SkipCertificateCheck
        SPsetsys_OutitemSyncInforbyProjectNumber $strProjectNumber "Success" "Insert Project Status" ""

        <# Debug code block #>
        <#Write-Output ($jsonResponse)
        #Write-Output ($jsonResponse)
        Write-Output ($jsonResponse.GetType())#>
        <# End debug code block #>
    } catch {

        #Use Stream reader to read response body from restful service
        $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
        $ErrResp = $streamReader.ReadToEnd() | ConvertFrom-Json
        $streamReader.Close()

        foreach ($err in $ErrResp.errors) {
            $totalerr = $err.message + " | " + $totalerr
        }

        $message = "Fail, Project Number:" + $strProjectNumber +  "Error Message:"+$totalerr
        SPsetsys_OutitemSyncInforbyProjectNumber $strProjectNumber "Fail" "Insert Project Status" $message
    }
    return $jsonResponse
}

#===========================================================================================================
#*********************************DATA Unifying to Service Now format*********************************
function DataUnifying($SPproject){
    if ([string]::IsNullOrEmpty($SPproject.FieldValues.D_x002d_Project_x0020_State)){$SPproject.FieldValues.D_x002d_Project_x0020_State = ""}
    if ([string]::IsNullOrEmpty($SPproject.FieldValues.D_x002d_Project_x0020_Phase)){$SPproject.FieldValues.D_x002d_Project_x0020_Phase = ""}

    if (![string]::IsNullOrEmpty($SPproject.FieldValues.D_x002d_Start_x0020_Date)){$SPproject.FieldValues.D_x002d_Start_x0020_Date = $SPproject.FieldValues.D_x002d_Start_x0020_Date.ToString("yyyy-MM-dd'T'HH:mm:ss+00:00")}else{$SPproject.FieldValues.D_x002d_Start_x0020_Date = $null}
    if (![string]::IsNullOrEmpty($SPproject.FieldValues.D_x002d_Estimated_x0020_Finish_x)){$SPproject.FieldValues.D_x002d_Estimated_x0020_Finish_x = $SPproject.FieldValues.D_x002d_Estimated_x0020_Finish_x.ToString("yyyy-MM-dd'T'HH:mm:ss+00:00")}else{$SPproject.FieldValues.D_x002d_Estimated_x0020_Finish_x = $null}
    if (![string]::IsNullOrEmpty($SPproject.FieldValues.D_x002d_Scheduled_x0020_Finish_x)){$SPproject.FieldValues.D_x002d_Scheduled_x0020_Finish_x = $SPproject.FieldValues.D_x002d_Scheduled_x0020_Finish_x.ToString("yyyy-MM-dd'T'HH:mm:ss+00:00")}else{$SPproject.FieldValues.D_x002d_Scheduled_x0020_Finish_x =$null}
    if (![string]::IsNullOrEmpty($SPproject.FieldValues.Project_x0020_Completed_x0020_or)){$SPproject.FieldValues.Project_x0020_Completed_x0020_or = $SPproject.FieldValues.Project_x0020_Completed_x0020_or.ToString("yyyy-MM-dd'T'HH:mm:ss+00:00")}else{$SPproject.FieldValues.Project_x0020_Completed_x0020_or = $null}
    if ([string]::IsNullOrEmpty($SPproject.FieldValues.D_x002d_Completion_x0020__x0025_)){$SPproject.FieldValues.D_x002d_Completion_x0020__x0025_ = ""}
    
    if ([string]::IsNullOrEmpty($SPproject.FieldValues.D_x002d_Actual_x0020_Costs)){$SPproject.FieldValues.D_x002d_Actual_x0020_Costs = "0"}
    
    if ([string]::IsNullOrEmpty($SPproject.FieldValues.D_x002d_SOW_x0020__x0023_)){$SPproject.FieldValues.D_x002d_SOW_x0020__x0023_ = ""}
    if ([string]::IsNullOrEmpty($SPproject.FieldValues.D_x002d_Project_x0020_Site.Url)){$SPproject.FieldValues.D_x002d_Project_x0020_Site = ""}
    
    $CostVarianceType = $SPproject.FieldValues.D_x002d_Cost_x002d_Variance_x002.GetType()
    if ($CostVarianceType.Name -eq "FieldCalculatedErrorValue") {$SPproject.FieldValues.D_x002d_Cost_x002d_Variance_x002 = "0"} 
    else { $SPproject.FieldValues.D_x002d_Cost_x002d_Variance_x002 = [double] $SPproject.FieldValues.D_x002d_Cost_x002d_Variance_x002.toString().Split('#')[1] }
    #if ([string]::IsNullOrEmpty($SPproject.FieldValues.D_x002d_Cost_x002d_Variance_x002)){$SPproject.FieldValues.D_x002d_Cost_x002d_Variance_x002 = ""}
    
    return $SPproject
}
#===========================================================================================================
function SNpatchProjectdetails ($strProjectNumber,$Body) {
    #$RESThostURI =  "https://servicenowconnector-pgs-dev.os99.gov.ab.ca/servicenowconnector/v1/projects/"+$strProjectNumber#DEV
    #$RESThostURI =  "https://servicenowconnector-pgs-uat.os99.gov.ab.ca/servicenowconnector/v1/projects/"+$strProjectNumber #UAT
    $RESThostURI =  $RESThost + "/servicenowconnector/v1/projects/"+$strProjectNumber #UAT

    $RESTuri = [uri]$RESThostURI

    #if ($RESTuri.host -eq "servicenowconnector-pgs-dev.os99.gov.ab.ca") {IgnoreSSLcert}

    # Params needs to be a Hashtable

    $Params = @{
        "URI"     = $RESThostURI
        "Method"  = 'PATCH'
        "Body" = $Body # TODO: Needs to be revisited to pass "green fields" and their values
        "Headers" = @{
            "Content-Type"  = 'application/json'
            "X-API-Authentication" = $strJWTToken
            "Authorization" = $strJWTToken
            "accept" = "application/json;odata=verbose"
        }
    }

    # Headers needs tobe iDictionary
    <#$Headers = @{
        Authorization = $strJWTToken
    }#>

    # body needs to be an Object
    #Write-Output $Body 


    #$body
    <#Write-Output "Params" $Params.GetType()
    Write-Output "Headers" $Headers.GetType()
    Write-Output "body" $body.GetType()
    Write-Host "Body " $body
    Write-Host "JWTtoken " $strJWTToken#>

    try {
        #if ($RESTuri.host -eq "servicenowconnector-pgs-dev.os99.gov.ab.ca") {IgnoreSSLcert} #DEV
        if ($RESTuri.host -eq "servicenowconnector-pgs-uat.os99.gov.ab.ca") {IgnoreSSLcert} #UAT
        $jsonResponse = Invoke-RestMethod @Params #-ErrorVariable $err #-SkipCertificateCheck

        #record timestamp on sys_itemSyncDateOut and set OUT status when success
        SPsetsys_OutitemSyncInforbyProjectNumber $strProjectNumber "Success" "Update Project Details" ""

    } catch {

        #Use Stream reader to read response body from restful service
        $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
        $ErrResp = $streamReader.ReadToEnd() | ConvertFrom-Json
        $streamReader.Close()

        foreach ($err in $ErrResp.errors) {
            $totalerr = $err.message + " | " + $totalerr
        }

        # Dig into the exception to get the Response details.
        $message = "Fail, Project Number:"+$strProjectNumber+"Error Message:"+$totalerr
        SPsetsys_OutitemSyncInforbyProjectNumber $strProjectNumber "Fail" "Update Project Details" $message

    }

    #$ErrResp

    return $jsonResponse
}
#===========================================================================================================
