#Set parameter values
$AdminSiteURL="https://psrsolutions1-admin.sharepoint.com"
$LoginName="proxy@psrsolutions1.onmicrosoft.com"
 
#Get Credentials to connect
$Credentials = Get-Credential
Connect-SPOService -url $AdminSiteURL -credential $Credential

$csvPath="D:\SP\PnP\CSV\RemoveUsers.csv"

$csvData=Import-Csv -Path $csvPath

foreach($item in $csvData)
{ 
    $SiteURL=$item.SiteURL
        #Get all groups 
        $Groups= Get-SPOSiteGroup -Site $SiteURL
 
        #Remove user from each group
        Foreach($Group in $Groups)
        {
           Write-Host "Stop Here"
           if($Group.Users -ccontains $LoginName)
           {
           Remove-SPOUser -Site $SiteURL -LoginName $LoginName -Group $Group.Name
           }
        }
}

