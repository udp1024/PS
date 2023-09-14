Import-module -name .\SN2SPsyncFunctions.psm1 -Force

setStaticVars -ErrorAction SilentlyContinue
initializeSN2SPsync

$SPLastUpdate = SPgetLastUpdateRunInDateLocal

if ([string]::IsNullOrEmpty($SPLastUpdate.FieldValues.LastUpdateDateTimeIn)){$SPLastUpdate.FieldValues.LastUpdateDateTimeIn = (Get-Date).AddYears(-2000)} # for fist time run, give earlist value

<#$SPLastUpdate = SPgetLastUpdateRunDate#>
#$P = SNgetAllProjects($SPLastUpdate.FieldValues.LastUpdateDateTime)
$P = SNgetAllProjects($SPLastUpdate.FieldValues.LastUpdateDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss+00:00"))
#$P = SNgetAllProjects($SPLastUpdate.FieldValues.LastUpdateDateTime.ToString("yyyy-MM-dd HH:mm:ss"))
#$P = SNgetAllProjects($SPLastUpdate.FieldValues.LastUpdateDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss+00:00"))


$message = "***************************[GET]Loop In Start at:" + (Get-Date).ToUniversalTime()+"***************************"
WriteLog $message

foreach ($project in $P) {

      <#  Write-Output "projectNumber " $project[projectNumber]#>
        if (![string]::IsNullOrEmpty($project.projectNumber)){
                
                $projObj = SPgetProjectByProjectNumber($project.projectNumber)   


                try {

                        if (![string]::IsNullOrEmpty($projObj.FieldValues.D_x002d_Project_x0020_Number)) # project exist in SharePoint
                        {

                                #unify data format to "", data from SN and SP are different so can't compare null with ""
                                
                                $projectName = $projObj.FieldValues.D_x002d_Intake_x0020_Project_x00
                                $primaryContact = $projObj.FieldValues.D_x002d_Primary_x0020_Contact.LookupValue
                                $demandNumber = $projObj.FieldValues.D_x002d_Demand_x0020_Number
                                $demandName = $projObj.FieldValues.D_x002d_Demand_x0020_Name
                                $demandMinistry = $projObj.FieldValues.D_x002d_Demand_x0020_Ministry
                                $leadSector = $projObj.FieldValues.D_x002d_Lead_x0020_Sector
                                $totalApprovedBudget = $projObj.FieldValues.D_x002d_Updated_x0020_Budget_x00
                                $SNtotalApprovedBudget = currencyConvertToNumber($project.totalApprovedBudget) #need to convert currecy format to number for comparing
                                

                                


                                if ([string]::IsNullOrEmpty($projectName)){$projectName = ""}
                                if ([string]::IsNullOrEmpty($project.projectName)){$project.projectName =""}


                                if ([string]::IsNullOrEmpty($PrimaryContact)){$PrimaryContact = ""}
                                if ([string]::IsNullOrEmpty($project.primaryContact)){$project.primaryContact =""}

                                if ([string]::IsNullOrEmpty($demandNumber)){$demandNumber = ""}
                                if ([string]::IsNullOrEmpty($project.intakeNumber)){$project.intakeNumber = ""}

                                if ([string]::IsNullOrEmpty($demandName)){$demandName = ""}
                                if ([string]::IsNullOrEmpty($project.intakeName)){$project.intakeName =""}

                                if ([string]::IsNullOrEmpty($demandMinistry)){$demandMinistry = ""}
                                if ([string]::IsNullOrEmpty($project.intakeMinistry)){$project.intakeMinistry =""}

                                if ([string]::IsNullOrEmpty($leadSector)){$leadSector = ""}
                                if ([string]::IsNullOrEmpty($project.leadSector)){$project.leadSector =""}

                                #log project information get from serviceNow

                                $message = "Get From Service Now, Project Number:"+$project.projectNumber+"`n Project Name:"+$project.projectName+"`n Primary Contact"+$project.primaryContact+"`n Demand Number:"+$project.intakeNumber+"`n Demand Name:"+$project.intakeName+"`n Ministry:"+$project.intakeMinistry+"`n Lead Sector:"+$project.leadSector+"`n Total Budget:"+$SNtotalApprovedBudget
                                WriteLog $message

                                #if ([string]::IsNullOrEmpty($totalApprovedBudget)){$totalApprovedBudget = ""}
                                #if ([string]::IsNullOrEmpty($project.totalApprovedBudget)){$project.totalApprovedBudget =""}

                                if ($projectName -ne $project.projectName) 
                                {
                                        SPsetProjectByProjectNumberLocal $project "projectName" $projObj
                                } 
                                if ($PrimaryContact -ne $project.primaryContact) 
                                {
                                        SPsetProjectByProjectNumberLocal $project "primaryContact" $projObj
                                } 
                                if ($demandNumber -ne $project.intakeNumber) 
                                {
                                        SPsetProjectByProjectNumberLocal $project "Demand" $projObj
                                }   
                                
                                if ($demandName -ne $project.intakeName) 
                                {
                                        SPsetProjectByProjectNumberLocal $project "DemandName" $projObj
                                } 
                                if ($demandMinistry -ne $project.intakeMinistry) 
                                {
                                        SPsetProjectByProjectNumberLocal $project "DemandMinistry" $projObj
                                } 

                                if ($leadSector -ne $project.leadSector) 
                                {     
                                        SPsetProjectByProjectNumberLocal $project "LeadSector" $projObj
                                }      
                                
                                if ($totalApprovedBudget -ne $SNtotalApprovedBudget[0]) 
                                {
                                        SPsetProjectByProjectNumberLocal $project "TotalBudget" $projObj
                                } 

                                #$strQuery = "<Query><Where><Eq><FieldRef Name='D_x002d_Project_x0020_Number' /><Value Type='Text'>"+$project.projectNumber+"</Value></Eq></Where></Query>"

                                #Write-Output "strQuery " $strQuery
                                #Write-Output GetType($SPconnection)

                                #$item = Get-PnPListItem -Connection $SPconnection -List $SPlistProjects -Query $strQuery
                                
                                # check each value from SN against values from SP. If value is different - Log the field-name and values

                                # Set-PnPListItem -List $SPlistProjects -Identity $item -Values @{"D_x002d_Primary_x0020_Contact" = $project.primaryContact; "D_x002d_Intake_x0020_Project_x00"= $project.projectName}
                        }
                }
                catch {
                        Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__ 
                        Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription
                        Write-Host "Exception Message: " $_.Exception.Message
                        
                }

        }


}

SPsetLastUpdateRunDateBy "SN2SPsync" $startDateTime.ToString() "LastUpdateDateTime"
$message = "***************************[GET]Loop In end at" + (Get-Date).ToUniversalTime() + "***************************"
WriteLog $message
.\SPpostSNProjects.ps1
