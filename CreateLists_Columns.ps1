Import-Module Microsoft.Online.Sharepoint.PowerShell

$User = "ramakrishna@weaverit2.onmicrosoft.com"
$File = "D:\WeaverIT2\WeaverIT2pwd.txt" 
$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, (Get-Content $File | ConvertTo-SecureString)
#Site Collection 
$siteURL ="https://weaverit2.sharepoint.com/sites/dev/"
#Connects and Creates Context
Connect-PnPOnline -Url $siteUrl -Credentials $cred -Verbose
#Client Context
$ctx=Get-PnPContext -ErrorAction Continue
Execute-PnPQuery
$web=Get-PnPWeb
$lists=Get-PnPList
$designationList=$lists |where {$_.Title -eq 'Designation'}
$employeeList=$lists |where {$_.Title -eq 'Employee'}

function CreateList()
{
    #Create desinationList List
    if($designationList -eq $null)
    {
     New-PnPList -Title "Designation" -Template GenericList -Url "Designation"
     Write-Host "Designation - List Created Successfully!" -foregroundcolor Green 
    }
    else
    {
     Write-Host "Designation - List Already exist!" -ForegroundColor Red
    }

    #Create Employee List
    if($employeeList -eq $null)
    {
     New-PnPList -Title "Employee" -Template GenericList -Url "Employee"
     Write-Host "Employee - List Created Successfully!" -foregroundcolor Green 

     $EmployeeNameField=$fields |where {$_.Title -eq 'EmployeeName'}
            if($EmployeeNameField -eq $null)
            {
            Add-PnPField -DisplayName "EmployeeName" -InternalName "EmployeeName" -Type Text  -List "Employee" -AddToDefaultView
                                             
            write-host "Created Field EmployeeName in the List" -foregroundcolor Green 
            }
            else
            {
            write-host "Field EmployeeName Already Exists in the List" -foregroundcolor Red
            }  
            
     $EmployeeDeptField=$fields |where {$_.Title -eq 'EmployeeDept'}
            if($EmployeeDeptField -eq $null)
            {
            Add-PnPField -DisplayName "EmployeeDept" -InternalName "EmployeeDept" -Type Text  -List "Employee" -AddToDefaultView
                                             
            write-host "Created Field EmployeeDept in the List" -foregroundcolor Green 
            }
            else
            {
            write-host "Field EmployeeDept Already Exists in the List" -foregroundcolor Red
            }  
            
     $DeskNoField=$fields |where {$_.Title -eq 'DeskNo'}
            if($DeskNoField -eq $null)
            {
            Add-PnPField -DisplayName "DeskNo" -InternalName "DeskNo" -Type Number  -List "Employee" -AddToDefaultView
                                             
            write-host "Created Field DeskNo in the List" -foregroundcolor Green 
            }
            else
            {
            write-host "Field DeskNo Already Exists in the List" -foregroundcolor Red
            }             
     
            
      $DateOfBirthField=$fields |where {$_.Title -eq 'DateOfBirth'}
            if($DateOfBirthField -eq $null)
            {
            Add-PnPField -DisplayName "DateOfBirth" -InternalName "DateOfBirth" -Type DateTime  -List "Employee" -AddToDefaultView
                                             
            write-host "Created Field DateOfBirth in the List" -foregroundcolor Green 
            }
            else
            {
            write-host "Field DateOfBirth Already Exists in the List" -foregroundcolor Red
            }  

            $genderField=$fields |where {$_.Title -eq 'Gender'}
            if($genderField -eq $null)
            {
            Add-PnPField -DisplayName "Gender" -InternalName "Gender" -Type Choice -Choices @("Male","Female") -List "Employee" -AddToDefaultView
            write-host "Created Field Gender in the List" -foregroundcolor Green 
            }
            else
            {
            write-host "Field Gender Already Exists in the List" -foregroundcolor Red
            }

            $SystemAllotedField=$fields |where {$_.Title -eq 'SystemAlloted'}
            if($SystemAllotedField -eq $null)
            {
            #Set parameter values
            $ListName="Employee"
            $Name="SystemAlloted"
            $DisplayName="SystemAlloted"
            $Description="SystemAlloted"
            $DefaultValue="0" #0 for No / 1 for Yes
 
            #Call the function to add column to list
            Add-YesNoColumnToList -SiteURL $SiteURL -ListName $ListName -Name $Name -DisplayName $DisplayName -Description $Description -DefaultValue $DefaultValue
            write-host "Created Field SystemAlloted in the List" -foregroundcolor Green 
            }
            else
            {
            write-host "Field SystemAlloted Already Exists in the List" -foregroundcolor Red
            }  
            
      $UrlField=$fields |where {$_.Title -eq 'Url'}
            if($UrlField -eq $null)
            {
            $listName="Employee"
            $Name="Url"
            $DisplayName="Url"
            $Format="Hyperlink" #or "Image"
            #Call the function to add column to list
            Add-HyperLinkPictureColumnToList -SiteURL $siteURL -ListName $listName -Name $Name -DisplayName $DisplayName -Description $Description -Format $Format
            }
            else
            {
            write-host "Field Url Already Exists in the List" -foregroundcolor Red
            }
            
       
      $DesignationField=$fields |where {$_.Title -eq 'Designation'}
            if($DesignationField -eq $null)
            {   
            #Set parameter values
            $listName="Employee"
            $Name="Designation"
            $DisplayName="Designation"
            $Description="Designation"
            $listLookup = "Designation" #Parent List to Lookup
            $LookupField="Title"

            #Call the function to add column to list
            AddLookupColumnToList -SiteURL $siteURL -ListName $listName -Name $Name -DisplayName $DisplayName -Description $Description -LookupListName $listLookup -LookupField $LookupField
            }
            else
            {
            write-host "Field Designation Already Exists in the List" -foregroundcolor Red
            }  

      $ManagerField=$fields |where {$_.Title -eq 'Manager'}
            if($ManagerField -eq $null)
            {
            Add-PnPField -DisplayName "Manager" -InternalName "Manager" -Type User  -List "Employee" -AddToDefaultView
                                            
            write-host "Created Field Manager in the List" -foregroundcolor Green 
            }
            else
            {
            write-host "Field Manager Already Exists in the List" -foregroundcolor Red
            }         
    }
    else
    {
     Write-Host "Employee - List Already exist!" -ForegroundColor Red
    }
}

#Custom function to add column to list
Function Add-HyperLinkPictureColumnToList()
{
    param
    (
        [Parameter(Mandatory=$true)] [string] $SiteURL,
        [Parameter(Mandatory=$true)] [string] $ListName,
        [Parameter(Mandatory=$true)] [string] $Name,
        [Parameter(Mandatory=$true)] [string] $DisplayName,
        [Parameter(Mandatory=$false)] [string] $Description="",
        [Parameter(Mandatory=$false)] [string] $IsRequired = "FALSE",
        [Parameter(Mandatory=$false)] [string] $Format ="Hyperlink"
    )
 
    #Generate new GUID for Field ID
    $FieldID = New-Guid
 
    Try {
        #$Cred= Get-Credential
        $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.Username, $Cred.Password)
 
        #Setup the context
        $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
        $Ctx.Credentials = $Credentials
         
        #Get the List
        $List = $Ctx.Web.Lists.GetByTitle($ListName)
        $Ctx.Load($List)
        $Ctx.ExecuteQuery()
 
        #Check if the column exists in list already
        $Fields = $List.Fields
        $Ctx.Load($Fields)
        $Ctx.executeQuery()
        $NewField = $Fields | where { ($_.Internalname -eq $Name) -or ($_.Title -eq $DisplayName) }
        if($NewField -ne $NULL) 
        {
            Write-host "Column $Name already exists in the List!" -f Yellow
        }
        else
        {
            #Define XML for Field Schema
            $FieldSchema = "<Field Type='URL' ID='{$FieldID}' DisplayName='$DisplayName' Name='$Name' Description='$Description' Required='$IsRequired' Format='$Format' />"
            $NewField = $List.Fields.AddFieldAsXml($FieldSchema,$True,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint)
            $Ctx.ExecuteQuery()   
 
            Write-host "New Column Added to the List Successfully!" -ForegroundColor Green 
        }
    }
    Catch {
        write-host -f Red "Error Adding Column to List!" $_.Exception.Message
    }
}


#Custom function to add column to list
Function AddLookupColumnToList()
{
    param
    (
        [Parameter(Mandatory=$true)] [string] $SiteURL,
        [Parameter(Mandatory=$true)] [string] $ListName,
        [Parameter(Mandatory=$true)] [string] $Name,
        [Parameter(Mandatory=$true)] [string] $DisplayName,
        [Parameter(Mandatory=$false)] [string] $Description="",
        [Parameter(Mandatory=$false)] [string] $IsRequired = "FALSE",
        [Parameter(Mandatory=$false)] [string] $EnforceUniqueValues = "FALSE",
        [Parameter(Mandatory=$true)] [string] $LookupListName,
        [Parameter(Mandatory=$true)] [string] $LookupField
    )
 
    #Generate new GUID for Field ID
    $FieldID = New-Guid
 
    Try {
        $creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($cred.UserName, $cred.Password)
        #Setup the context
        $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
        $ctx.credentials = $creds
         
        #Get the web, List and Lookup list
        $Web = $Ctx.web
        $List = $Web.Lists.GetByTitle($ListName)
        $LookupList = $Web.Lists.GetByTitle($LookupListName)
        $Ctx.Load($Web)
        $Ctx.Load($List)
        $Ctx.Load($LookupList)
        $Ctx.ExecuteQuery()
 
        #Check if the column exists in list already
        $Fields = $List.Fields
        $Ctx.Load($Fields)
        $Ctx.executeQuery()
        $NewField = $Fields | where { ($_.Internalname -eq $Name) -or ($_.Title -eq $DisplayName) }
        if($NewField -ne $NULL) 
        {
            Write-host "Column $Name already exists in the List!" -f Yellow
        }
        else
        {
            #Get IDs of Lookup List and Web
            $LookupListID= $LookupList.id
            $LookupWebID=$web.Id
 
            $FieldSchema = "<Field Type='Lookup' ID='{$FieldID}' DisplayName='$DisplayName' Name='$Name' Description='$Description' Required='$IsRequired' EnforceUniqueValues='$EnforceUniqueValues' List='$LookupListID' WebId='$LookupWebID' ShowField='$LookupField' />"
            $NewField = $List.Fields.AddFieldAsXml($FieldSchema,$True,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint)
            $Ctx.ExecuteQuery()   
 
            Write-host "New Column Added to the List Successfully!" -ForegroundColor Green 
        }
    }
    Catch {
        write-host -f Red "Error Adding Column to List!" $_.Exception.Message
    }
}

#Custom function to add column to list
Function Add-YesNoColumnToList()
{
    param
    (
        [Parameter(Mandatory=$true)] [string] $SiteURL,
        [Parameter(Mandatory=$true)] [string] $ListName,
        [Parameter(Mandatory=$true)] [string] $Name,
        [Parameter(Mandatory=$true)] [string] $DisplayName,
        [Parameter(Mandatory=$false)] [string] $Description="",
        [Parameter(Mandatory=$false)] [string] $DefaultValue = "0"
    )
 
    #Generate new GUID for Field ID
    $FieldID = New-Guid
 
    Try {
        $creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($cred.UserName, $cred.Password)
        #Setup the context
        $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
        $ctx.credentials = $creds
         
        #Get the List
        $List = $Ctx.Web.Lists.GetByTitle($ListName)
        $Ctx.Load($List)
        $Ctx.ExecuteQuery()
 
        #Check if the column exists in list already
        $Fields = $List.Fields
        $Ctx.Load($Fields)
        $Ctx.executeQuery()
        $NewField = $Fields | where { ($_.Internalname -eq $Name) -or ($_.Title -eq $DisplayName) }
        if($NewField -ne $NULL) 
        {
            Write-host "Column $Name already exists in the List!" -f Yellow
        }
        else
        {
            #Define XML for Field Schema
            $FieldSchema = "<Field Type='Boolean' ID='{$FieldID}' DisplayName='$DisplayName' Name='$Name' Description='$Description'><Default>$DefaultValue</Default></Field>"
            $NewField = $List.Fields.AddFieldAsXml($FieldSchema,$True,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint)
            $Ctx.ExecuteQuery()   
 
            Write-host "New Column Added to the List Successfully!" -ForegroundColor Green 
        }
    }
    Catch {
        write-host -f Red "Error Adding Column to List!" $_.Exception.Message
    }
}



CreateList
#Remove-PnPList -Identity Designation -Force
#Remove-PnPList -Identity Employee -Force
#http://www.sharepointdiary.com/2016/10/sharepoint-online-add-yes-no-column-to-list-using-powershell.html
