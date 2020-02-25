#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.UserProfiles.dll"
 
#Import Azure AD Module
Import-Module MSOnline
 
Function Export-AllUserProfiles()
{
    param
    (
        [Parameter(Mandatory=$true)] [string] $TenantURL,
        [Parameter(Mandatory=$true)] [string] $CSVPath
    )   
    Try {
        #Setup Credentials to connect
        $Cred= Get-Credential
        $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.Username, $Cred.Password)
 
        #Setup the context
        $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($TenantURL)
        $Ctx.Credentials = $Credentials
         
        #Delete the CSV report file if exists
        if (Test-Path $CSVPath) { Remove-Item $CSVPath }
 
        #Get all Users
        Connect-MsolService -Credential $Cred
        $Users = Get-MsolUser -All |  Select-Object -ExpandProperty UserPrincipalName
         
        Write-host "Total Number of Profiles Found:"$Users.count -f Yellow
        #Get User Profile Manager
        $PeopleManager = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($Ctx)
        #Array to hold result
        $UserProfileData = @()
 
        Foreach ($User in $Users)
        {
            Write-host "Processing User Name:"$User
            #Get the User Profile
            $UserLoginName = "i:0#.f|membership|" + $User  #format to claims
            $UserProfile = $PeopleManager.GetPropertiesFor($UserLoginName)
            $Ctx.Load($UserProfile)
            $Ctx.ExecuteQuery()
            if($UserProfile.Email -ne $Null)
            {
            #Send Data to object array
            $UserProfileData += New-Object PSObject -Property @{
            'User Account' = $UserProfile.UserProfileProperties["UserName"]
            'Full Name' = $UserProfile.UserProfileProperties["PreferredName"]
            'E-mail' =  $UserProfile.UserProfileProperties["WorkEmail"]
            'Department' = $UserProfile.UserProfileProperties["Department"]
            'Location' = $UserProfile.UserProfileProperties["Office"]
            'Phone' = $UserProfile.UserProfileProperties["WorkPhone"]
            'Job Title' = $UserProfile.UserProfileProperties["Title"]
            }
            }
        }
        #Export the data to CSV
        $UserProfileData | Export-Csv $CSVPath -Append -NoTypeInformation
 
        write-host -f Green "User Profiles Data Exported Successfully to:" $CSVPath
  }
    Catch {
        write-host -f Red "Error Exporting User Profile Properties!" $_.Exception.Message
    }
}
 
#Call the function
$TenantURL="https://psrsolutions1.sharepoint.com"
$CSVPath="D:\SP\PnP\CSV\UserProfiles.csv"
 
Export-AllUserProfiles -TenantURL $TenantURL -CSVPath $CSVPath


