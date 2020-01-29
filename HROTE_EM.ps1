Import-Module Microsoft.Online.Sharepoint.PowerShell

#Get list items to Array with passed field titles
Function GetItemArray
{
    Param($list,$fieldtitles)
    $itemArray=@()
    $items=Get-PnPListItem -List $list -Fields $fieldTitles
    foreach($item in $items)
    {
        $object  = New-Object -TypeName PSObject
        foreach($fieldTitle in $fieldTitles)
        {
            if($item[$fieldTitle] -ne $null)
            {
                $val =[string]$item[$fieldTitle]
                
                $LookUpVal = $val.replace(' ',' ').split(' ')
                $UserVal = $val.replace(' ',' ').split(' ')

                if($val -eq [Microsoft.SharePoint.Client.FieldLookupValue])
                {           
                $val=$item[$fieldTitle].LookupId
                }
                
                if($LookUpVal[0] -eq [Microsoft.SharePoint.Client.FieldLookupValue])
                {
                    $val=$item[$fieldTitle].LookupId 
                    $val=($val | Select-Object -Unique) -Join ","
                }
                
                if($val -eq [Microsoft.SharePoint.Client.FieldUserValue])
                {           
                $val=$item[$fieldTitle].LookupId
                }

                if($UserVal[0] -eq [Microsoft.SharePoint.Client.FieldUserValue])
                {
                    $val=$item[$fieldTitle].LookupId 
                    $val=($val | Select-Object -Unique) -Join ","
                }
          
                $object | Add-Member -MemberType Noteproperty -Name $fieldTitle -Value $val
            }
            else
            {
                $object | Add-Member -MemberType Noteproperty -Name $fieldTitle -Value ''
            }
        }
        #$object
        $itemArray+=$object
    }
    Return ($itemArray)
}

$User = "rtamatannaga@vmware.com"
$File = "D:\Ramakrishna\SharePointO365ProdPassword5.txt"
 
$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, (Get-Content $File | ConvertTo-SecureString)

#Site Collection 
#$siteURL = "https://onevmw.sharepoint.com/sites/sales-emea/Managers%20Portal/"
$siteURL ="https://onevmw.sharepoint.com/sites/sales-emea/stage/"

#Connects and Creates Context
Connect-PnPOnline -Url $siteUrl -Credentials $cred -Verbose
#Client Context
$ctx=Get-PnPContext -ErrorAction Continue
$ctx.RequestTimeout=[System.Threading.Timeout]::Infinite
Execute-PnPQuery
$web=Get-PnPWeb
$directorypath = 'D:\Ramakrishna\EMEA_NewScripts\Data'
$lists=Get-PnPList
$users=Get-PnPUser
$HROTE_EMList=$lists | where {$_.Title -eq 'HROTE_EM'}
$HROTE_EM_DL=$lists | where {$_.Title -eq 'HROTE_EM_DL'}
$unprocessedFiles=Get-PnPListItem -List $HROTE_EM_DL | where {[System.Convert]::ToBoolean($_['Processed']) -eq $false}
$HROTE_EMArray=@()
$fieldTitles='ID','EMC_ID'

Function AddItemHROTEEMList()
{
        foreach($unprocessedfile in $unprocessedFiles)
        {
        $HROTE_EMArray=GetItemArray $HROTE_EMList $fieldTitles
        $fileName=$unprocessedfile.FieldValues.FileLeafRef
        $fileUrl=$unprocessedfile.FieldValues.FileDirRef+'/'+$unprocessedfile.FieldValues.FileLeafRef
        Write-Host $fileUrl
        $csvVariable=Get-PnPFile -Url $fileUrl -Path $directorypath -AsFile
        $filePath=$directorypath+"\"+$fileName
        $csvVariable=Import-Csv $filePath
            foreach ($row in $csvVariable)
            {                
            $StartDate=$row.StartDate
            if(![string]::IsNullOrWhiteSpace($StartDate))
            {
            [DateTime]$StartDateVal=[DateTime]$row.StartDate
            }
            #if($StartDate -eq ""){$StartDateVal=[Nullable[DateTime]]$null} 
            #else{[DateTime]$StartDateVal=[DateTime]$row.StartDate}             
            $EndDate=$row.EndDate
            if(![string]::IsNullOrWhiteSpace($EndDate))
            {
            [DateTime]$EndDateVal=[DateTime]$row.EndDate
            }
            #if($EndDate -eq ""){$EndDateVal=[Nullable[DateTime]]$null} 
            #else{[DateTime]$EndDateVal=[DateTime]$row.EndDate}             
            $NRDEndDate=$row.NRD_x0020_End_x0020_Date
            if(![string]::IsNullOrWhiteSpace($NRDEndDate))
            {
            [DateTime]$NRDEndDateVal=[DateTime]$row.NRD_x0020_End_x0020_Date
            }
            #if($NRDEndDate -eq ""){$NRDEndDateVal=[Nullable[DateTime]]$null} 
            #else{[DateTime]$NRDEndDateVal=[DateTime]$row.NRD_x0020_End_x0020_Date}                       
            $OpsManager=$row.Ops_x0020_Manager            
            $OpsManagerVal=$users | where {$_.ID -eq $OpsManager}
            if(![string]::IsNullOrWhiteSpace($OpsManagerVal))
            {
            $OpsManagerEmail=$OpsManagerVal.Email
            $OpsManagerEmailID=$OpsManagerVal.ID
            }
            $MBOProcessOwner=$row.MBO_x0020_Process_x0020_Owner            
            $MBOProcessOwnerVal=$users | where {$_.ID -eq $MBOProcessOwner}
            if(![string]::IsNullOrWhiteSpace($MBOProcessOwnerVal))
            {
            $MBOProcessOwnerEmail=$MBOProcessOwnerVal.Email
            $MBOProcessOwnerEmailID=$MBOProcessOwnerVal.ID
            }
            #$HROTE_EMArray=GetItemArray $HROTE_EMList $fieldTitles
            $emcID=$row.EMC_x0020_ID
            $emcID=$emcID.PadLeft(6,"0")

                $HROTE_EMItem=$HROTE_EMArray | where {$_.EMC_ID -eq $emcID }
                $itemID=$null
                if(![string]::IsNullOrWhiteSpace($HROTE_EMItem) -and ($HROTE_EMItem -isnot [System.Array]) -and ($HROTE_EMArray -isnot [System.Array]))
                {
                $itemID=$HROTE_EMArray.ID
                }
                if($HROTE_EMArray -is [System.Array] -and (![string]::IsNullOrWhiteSpace($HROTE_EMItem)))
                {
                $itemID=$HROTE_EMItem.ID 
                }
                if($HROTE_EMItem -is [System.Array])
                {
                $itemID=$HROTE_EMItem[0].ID
                Remove-PnPListItem -List "HROTE_EM" -Identity HROTE_EMItem[1].ID
                Write-Host "Removed-"$emcID -foregroundcolor red 
                }             
                if((![string]::IsNullOrWhiteSpace($row.EMC_x0020_ID)) -and ($row.Plan_x0020_No_x002e_ -eq 1) -and ([string]::IsNullOrWhiteSpace($itemID)))
                {
                $creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($cred.UserName, $cred.Password)
                $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)
                $ctx.credentials = $creds
                $lists = $ctx.web.Lists  
                try{
                        $list = $lists.GetByTitle("HROTE_EM") 
                        $listItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation  
                        $newItem = $list.AddItem($listItemInfo)
                      #Single Line of Text
                        if($row.Title -ne ""){$newItem["Title"]=$row.Title.PadLeft(6,"0")} 
                        if($row.Ops_x0020_Group -ne ""){$newItem["Ops_x0020_Group"]=$row.Ops_x0020_Group}                        
                        if($row.EMC_x0020_ID -ne ""){$newItem["EMC_ID"]=$row.EMC_x0020_ID.PadLeft(6,"0")} 
                        if($row.EMC_x0020_ID -ne ""){$newItem["EMCID"]=$row.EMC_x0020_ID.PadLeft(6,"0")}                         
                        if($row.Employee_x0020_Name -ne ""){$newItem["EmployeeName"]=$row.Employee_x0020_Name}  
                        if($row.Manager -ne ""){$newItem["Manager"]=$row.Manager}
                        if($row.Plan_x0020_Type -ne ""){$newItem["Plan_x0020_Type1"]=$row.Plan_x0020_Type}
                        if($row.Territory_x0020_Type -ne ""){$newItem["Territory_x0020_Type1"]=$row.Territory_x0020_Type}
                        if($row.NRD_x0020_Status -ne ""){$newItem["NRDStatus"]=$row.NRD_x0020_Status}
                        if($row.MBO_x0020_Status -ne ""){$newItem["MBO_x0020_Status"]=$row.MBO_x0020_Status}                                                                     
                    #Lookup Columns                        
                        if($row.Region0 -ne ""){$newItem["Regionlookup"]=$row.Region0 }
                        if($row.Sub_x002d_region0 -ne ""){$newItem["Sub_x002d_region0"]=$row.Region0 }
                        if($row.Manager_x0020__x0028_HR_x0029_ -ne ""){$newItem["ManagerHR"]=$row.Manager_x0020__x0028_HR_x0029_ }
                        if($row.Position0 -ne ""){$newItem["Position"]=$row.Position0 }
                        if($row.Half_x0020_PlanID -ne ""){$newItem["Half_x0020_PlanID1"]=$row.Half_x0020_PlanID }
                        if($row.Territory -ne ""){$newItem["Territory1"]=$row.Territory }                                                                
                    #Date Time
                        if(![string]::IsNullOrWhiteSpace($row.NRD_x0020_End_x0020_Date)){$newItem["NRD_x0020_End_x0020_Date"]=$NRDEndDateVal }
                        if(![string]::IsNullOrWhiteSpace($row.StartDate)){$newItem["StartDate"]=$StartDateVal}
                        if(![string]::IsNullOrWhiteSpace($row.EndDate)){$newItem["EndDate"]=$EndDateVal }            
                    #Number Columns
                        if($row.Plan_x0020_No_x002e_ -ne ""){$newItem["Plan_x0020_No_x002e_1"]=[int]$row.Plan_x0020_No_x002e_ -replace "[^0-9.]",''}
                    #Choice Columns   
                        if($row.Status -ne ""){$newItem["Status"]=$row.Status }
                        if($row.GermanWorkContract -ne ""){$newItem["German_x0020_Work_x0020_Contract"]=$row.GermanWorkContract}
                    #Person Or Group     
                        if($row.Ops_x0020_Manager -ne ""){$newItem["OpsManager"]=$OpsManagerEmailID }                
                        if($row.MBO_x0020_Process_x0020_Owner -ne ""){$newItem["MBO_x0020_Process_x0020_Owner"]=$MBOProcessOwnerEmailID }                                                       
                    #Yes/No Columns
                        $newItem.Update()
                        $ctx.load($newItem)      
                        $ctx.executeQuery()  
                        Write-Host "Item Added with ID-"$newItem.Id "EMC_ID-"$emcID                        
                        }
                        catch
                        {  
                            write-host "$($_.Exception.Message)" -foregroundcolor red  
                        }
                }
                if(![string]::IsNullOrWhiteSpace($itemID) -and ($itemID -isnot [System.Array]) -and (![string]::IsNullOrWhiteSpace($row.EMC_x0020_ID)) -and ($row.Plan_x0020_No_x002e_ -eq 2))
                {
                  $creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($cred.UserName, $cred.Password)
                  $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)
                  $ctx.credentials = $creds
                  try{
                  $lists = $ctx.web.Lists  
                  $list = $lists.GetByTitle("HROTE_EM") 
                  $listItem = $list.GetItemById($itemID)
             #Single Line of Text
                    if($row.Plan_x0020_Type -ne ""){$listItem["Plan_x0020_Type2"]=$row.Plan_x0020_Type}
                    if($row.Territory_x0020_Type -ne ""){$listItem["Territory_x0020_Type2"]=$row.Territory_x0020_Type}
                #Lookup Columns
                    if($row.Half_x0020_PlanID -ne ""){$listItem["Half_x0020_PlanID2"]=$row.Half_x0020_PlanID }                                     
                    if($row.Territory -ne ""){$listItem["Territory2"]=$row.Territory }                                            
                #Number Columns
                    if($row.Plan_x0020_No_x002e_ -ne ""){$listItem["Plan_x0020_No_x002e_2"]=[int]$row.Plan_x0020_No_x002e_ -replace "[^0-9.]",''}                                                
                  $listItem.Update()  
                  $ctx.load($listItem)      
                  $ctx.executeQuery()  
                  Write-Host "Item Updated with ID-"$listItem.Id "EMC_ID-"$emcID -foregroundcolor Green                  
                  }
                  catch
                  {  
                    write-host "$($_.Exception.Message)" -foregroundcolor red  
                  }
                }
                if(![string]::IsNullOrWhiteSpace($itemID) -and ($itemID -isnot [System.Array]) -and (![string]::IsNullOrWhiteSpace($row.EMC_x0020_ID)) -and ($row.Plan_x0020_No_x002e_ -eq 3) )
                {
                  $creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($cred.UserName, $cred.Password)
                  $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)
                  $ctx.credentials = $creds
                  try{
                  $lists = $ctx.web.Lists  
                  $list = $lists.GetByTitle("HROTE_EM") 
                  $listItem = $list.GetItemById($itemID)
                 #Single Line of Text
                    if($row.Plan_x0020_Type -ne ""){$listItem["Plan_x0020_Type3"]=$row.Plan_x0020_Type}
                    if($row.Territory_x0020_Type -ne ""){$listItem["Territory_x0020_Type3"]=$row.Territory_x0020_Type}
                #Lookup Columns
                    if($row.Half_x0020_PlanID -ne ""){$listItem["Half_x0020_PlanID3"]=$row.Half_x0020_PlanID }                                     
                    if($row.Territory -ne ""){$listItem["Territory3"]=$row.Territory }                                            
                #Number Columns
                    if($row.Plan_x0020_No_x002e_ -ne ""){$listItem["Plan_x0020_No_x002e_3"]=[int]$row.Plan_x0020_No_x002e_ -replace "[^0-9.]",''}                                                
                  $listItem.Update()  
                  $ctx.load($listItem)      
                  $ctx.executeQuery()  
                  Write-Host "Item Updated with ID-"$listItem.Id "EMC_ID-"$emcID -foregroundcolor Yellow
                  }
                  catch
                  {  
                    write-host "$($_.Exception.Message)" -foregroundcolor red  
                  }
                }
          }
            $processedFileID=$unprocessedfile.FieldValues.ID
            Set-PnPListItem -List $HROTE_EM_DL -Identity  $processedFileID -Values @{"Processed"=$true}
        }
 }


 AddItemHROTEEMList
 #Disconnect-PnPOnline