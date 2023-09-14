Import-module -name .\SN2SPsyncFunctions.psm1
    
setStaticVars -ErrorAction SilentlyContinue
initializeSN2SPsync

#===========================================================================================================

function OutErrorHandling(){


    
}

#---------- Control Routine  
$spFilteredProjects = SPgetFilteredProjects

$message = "***************************[POST]Loop In Start at:" + (Get-Date).ToUniversalTime() +"***************************"
WriteLog $message

#DOC: Loop through the filtered items in SP Delivery list 
#DOC: For each row:
foreach ($spFilteredProject in $spFilteredProjects) {    

$global:Oldmessage = "" #refresh the global variable
$global:OldStatus = "" #refresh the global variable
    #DOC: If [SP Last Modified Date] > [Item Sync Date OUT]
    # $projectLastUpdatedDate = Get-Date($spFilteredProject.FieldValues.Modified)
    # $projectLastSyncDate = Get-Date($spFilteredProject.FieldValues.sys_itemSyncDateOut)
    # $itemLastSyncDate = $spFilteredProject.FieldValues.sys_itemSyncDateOut # we don't need this as it is a property of the item in the iteration
    <#if ($spFilteredProject.FieldValues.sys_itemSyncDateOut::IsNullOrWhiteSpace) {
        $projectLastSyncDate = Get-Date("0/0/0 00:00")
    }else {
        $projectLastSyncDate = Get-Date($spFilteredProject.FieldValues.sys_itemSyncDateOut)
    }#>
    
    $projectNumber = $spFilteredProject.FieldValues.D_x002d_Project_x0020_Number
   # $strProjectNumber = [System.Convert]::ToBase64String($projectNumber.ToString())    
       
    #$SNProject = SNgetProjectByProjectNumber("PRJ0010004")
    $SNProject = SNgetProjectByProjectNumber($projectNumber)

    
    if(![string]::IsNullOrEmpty($SNProject[2]))
    # TO check if Service Now return any project, $SNProject[2] will return value if the call is success and return null if the call is fail
           
    {

            if ([string]::IsNullOrEmpty($SNProject.lastProjectStatusUpdatedDate)) {
                #check last project status update date. If empty, assign a very low date value
                $SNlastProjectStatusUpdatedDate  =(Get-Date).AddYears(-2000).ToString("yyyy-MM-dd HH:mm:ss")
            }else{
                $SNlastProjectStatusUpdatedDate = $SNProject.lastProjectStatusUpdatedDate
            }


            if ([string]::IsNullOrEmpty($spFilteredProject.FieldValues.sys_itemSyncDateOut)) {
                # Give an earliest date and time when frist time run this program
                $spFilteredProject.FieldValues.sys_itemSyncDateOut =(Get-Date).AddYears(-2000)
            }

            if ($spFilteredProject.FieldValues.Modified.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") -gt $SNlastProjectStatusUpdatedDate -AND $spFilteredProject.FieldValues.Modified.ToString("yyyy-MM-dd HH:mm:ss") -gt $spFilteredProject.FieldValues.sys_itemSyncDateOut.AddSeconds(5).ToString("yyyy-MM-dd HH:mm:ss")){
                #Add 5 more second to the sys_itemSyncDateOut, beacuse Last modify will always be greater than sys_itemSyncDateOut when there is no change this item

            
                #DOC: Project # exists and data returned
                #DOC:  Is [SP Last Modified date] > [Service now last update date for project] 
                #DOC: ‘Sharepoint record is newer than the service now record
                #DOC: *** Update the Project table in SN ***
                #DOC: ### call the [update parent record in Project] <only project related green fields>
                #DOC: PATCH /servicenowconnector/v1/projects/{projectNumber})           
                #DOC: ‘This code is smart enough to do an update or insert or ignore
                #DOC: **** Update or Insert Project Status table in SN ****
                
                $SPproject = DataUnifying($spFilteredProject)

                $patchBody = [pscustomobject]@{
                    projectState                    =  $SPproject.FieldValues.D_x002d_Project_x0020_State
                    projectPhase                    =  $SPproject.FieldValues.D_x002d_Project_x0020_Phase
                    completionPercentage            =  $SPproject.FieldValues.D_x002d_Completion_x0020__x0025_
                    actualStartDate                 =  $SPproject.FieldValues.D_x002d_Start_x0020_Date
                    estimateFinishDate              =  $SPproject.FieldValues.D_x002d_Estimated_x0020_Finish_x
                    scheduleFinishDate              =  $SPproject.FieldValues.D_x002d_Scheduled_x0020_Finish_x
                    projectCompletedOrCancelledDate =  $SPproject.FieldValues.Project_x0020_Completed_x0020_or
                    actualCost                      =  $SPproject.FieldValues.D_x002d_Actual_x0020_Costs 
                    sowAndRASNumber                 =  $SPproject.FieldValues.D_x002d_SOW_x0020__x0023_ 
                    projectSite                     =  $SPproject.FieldValues.D_x002d_Project_x0020_Site.Url
                    costVariance                    =  $SPproject.FieldValues.D_x002d_Cost_x002d_Variance_x002
                } | ConvertTo-Json

                # save PathchBody to Log
                $message = "Patch "+$spFilteredProject.FieldValues.D_x002d_Project_x0020_Number+"`n"+$patchBody
                WriteLog $message

                $TempResponse = SNpatchProjectdetails $spFilteredProject.FieldValues.D_x002d_Project_x0020_Number $patchBody

                #Error handling
           
                #$tempProj = SNgetProjectByProjectNumber($spFilteredProject.FieldValues.D_x002d_Project_x0020_Number)
                
                #TODO: Check response handle errors and recover

                #DOC: If [D-Status Date?] is not null 
                #DOC: 		    {
                #DOC: 			If [D-Status Date?]  > the newest s_on date field 
                #DOC: or if the newest s_on date is null  (s_on date in the record retrieved from SN) 
                





                # Three possibilities; $SNlastStatusUpdatedDate = null~add/old~nothing/equal~add/new~add
                #if (![string]::IsNullOrEmpty(($spFilteredProject.FieldValues.Entry_x0020_updated_x0020_date)) -OR (($spFilteredProject.FieldValues.Entry_x0020_updated_x0020_date -gt $SNlastProjectStatusUpdatedDate) -OR [string]::IsNullOrEmpty($SNlastProjectStatusUpdatedDate))) {
                    #DOC: ###Insert new record in Project Status table [passing green fields)    


                    $statusDate = $spFilteredProject.FieldValues.Entry_x0020_updated_x0020_date
                    if (![string]::IsNullOrEmpty($statusDate)){$statusDate = $spFilteredProject.FieldValues.Entry_x0020_updated_x0020_date.ToString("yyyy-MM-dd HH:mm:ss")}


                    #$spFilteredProject.FieldValues.Entry_x0020_updated_x0020_date.ToLocalTime() #convert to locatime
                    #[datetime]::parseexact((($SNlastProjectStatusUpdatedDate.Split(" "))[0]), 'yyyy-MM-dd', $null) #convert to date time format, make sure it is first hour of date
                    #[datetime]::parseexact((($SNlastProjectStatusUpdatedDate.Split(" "))[0]), 'yyyy-MM-dd', $null) #convert to date time format, make sure it is first hour of date


                    if (![string]::IsNullOrEmpty($statusDate) -AND  $spFilteredProject.FieldValues.Entry_x0020_updated_x0020_date.ToLocalTime() -gt [datetime]::parseexact((($SNlastProjectStatusUpdatedDate.Split(" "))[0]), 'yyyy-MM-dd', $null)) {

                    
                        #TODO: add values from properties of hte SharePoint Project
                    $insertBody = [pscustomobject]@{ #Green fields from excel file 
                        shortStatus          = $spFilteredProject.FieldValues.D_x002d_Short_x0020_Status1
                        scope                = $spFilteredProject.FieldValues.D_x002d_Scope
                        schedule             = $spFilteredProject.FieldValues.D_x002d_Schedule
                        finance              = $spFilteredProject.FieldValues.D_x002d_Budget
                        overallProjectHealth = $spFilteredProject.FieldValues.D_x002d_Overall_x0020_Project_x0
                        projectHealthReason  = $spFilteredProject.FieldValues.D_x002d_Project_x0020_Health_x00
                        correctiveActions    = $spFilteredProject.FieldValues.D_x002d_Corrective_x0020_Actions
                        statusDate           = $statusDate

                    } | ConvertTo-Json

                    #Save insertBody to log
                    $message = "Insert Status for "+$spFilteredProject.FieldValues.D_x002d_Project_x0020_Number+"`n"+$insertBody
                    WriteLog $message

                    SNinsertProjectStatus $spFilteredProject.FieldValues.D_x002d_Project_x0020_Number $insertBody

                    #DOC: 			Else
                }
                elseif (![string]::IsNullOrEmpty($statusDate)) {	

                    $Body = [pscustomobject]@{ #Green fields from excel file 
                        shortStatus          = $spFilteredProject.FieldValues.D_x002d_Short_x0020_Status1
                        scope                = $spFilteredProject.FieldValues.D_x002d_Scope
                        schedule             = $spFilteredProject.FieldValues.D_x002d_Schedule
                        finance              = $spFilteredProject.FieldValues.D_x002d_Budget
                        overallProjectHealth = $spFilteredProject.FieldValues.D_x002d_Overall_x0020_Project_x0
                        projectHealthReason  = $spFilteredProject.FieldValues.D_x002d_Project_x0020_Health_x00
                        correctiveActions    = $spFilteredProject.FieldValues.D_x002d_Corrective_x0020_Actions
                        statusDate           = $statusDate
                    } | ConvertTo-Json

                    #Save Update Body to log
                    $message = "Insert Status for "+$spFilteredProject.FieldValues.D_x002d_Project_x0020_Number+"`n"+$Body
                    WriteLog $message

                    SNupdateProjectStatus $spFilteredProject.FieldValues.D_x002d_Project_x0020_Number $Body
                }
            }  
        
        else {
            #DOC:DO NOTHING
       
        }

    }

}

SPsetLastUpdateRunDateBy "SN2SPsync" $startDateTime.ToString() "LastUpdateDateTimeOut"
$message = "***************************[POST]Loop In end at" + (Get-Date).ToUniversalTime() + "***************************"
WriteLog $message
