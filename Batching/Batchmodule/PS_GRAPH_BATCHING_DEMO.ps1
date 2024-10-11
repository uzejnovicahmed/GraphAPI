#region ms learn
#.. https://learn.microsoft.com/en-us/graph/json-batching
#endregion

#region load modules

if(!(Get-Module Microsoft.Graph.*))
{
try
{
    Write-Output "Try to Import Microsoft.Graph Module"
    Import-Module Microsoft.Graph -ErrorAction Stop
    Write-Output "Import Microsoft.Graph Module was successfull"
}
catch
{
    Write-Output "Installing Microsoft.Graph Modules please Wait"

    Install-Module Microsoft.Graph.Authentication -Scope AllUsers -Force -Confirm:$false
    Install-Module Microsoft.Graph.Users.Actions -Scope AllUsers -Force -Confirm:$false
    Install-Module Microsoft.Graph.People -Scope AllUsers -Force -Confirm:$false
    Install-Module Microsoft.Graph.Users -Scope AllUsers -Force -Confirm:$false
    Install-Module Microsoft.Graph.PersonalContacts -Scope AllUsers -Force -Confirm:$false
    Install-Module Microsoft.Graph.Groups -Scope AllUsers -Force -AllowClobber -Confirm:$false

    Write-Output "Installation of Microsoft.Graph Modules finished"
}
}
else
{
    Write-Output "Microsoft.Graph Module was found"
}


Import-Module "C:\dev_github\DEMO_GRAPH_BATCHING\DEMO_GRAPH_BATCHING\DEMO_MODULES_BATCHNG.psm1" -DisableNameChecking -Force    

#endregion

#region authentication
#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#

$accesstoken = Request-AccessToken -TenantName "uzeah.onmicrosoft.com" -ClientId "82a377a7-ab4b-43a2-94f1-14b795d6897f" -ClientSecret ""    
Connect-MgGraph -AccessToken $accesstoken

#endregion authentication
#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#

#region glob variables
$GALSYNCFOLDERNAME = "CONTACTS_BATCHING_DEMO_DEV"
$UserID = 'alexw@uzeah.onmicrosoft.com'
#endregion

#$UserID = (Get-MgUser -All | ? {$_.mail -ne $null} | select mail).Mail
#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#

#region get contacts to sync into User Contacts Folder
$GAL_Users = $null
$GAL_Users = Get-GalSyncUsers -DebugPreference 'continue'
#endregion 

#region run Sync
Measure-Command {Update-GalSycnUserContactsParallelJobs -Contacts $GAL_Users -UserMails $UserID -GALSYNCFOLDERNAME $GALSYNCFOLDERNAME -accesstoken $($accesstoken) -DebugPreference 'Continue'}
#endregion

#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#
#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#

#region demonstration total time wasted without batching


$Usercontactfolder = Get-MgUserContactFolder -UserID $UserID | where {$_.DisplayName -eq "$($GALSYNCFOLDERNAME)"}
            
if($Usercontactfolder){
        Remove-MgUserContactFolder -ContactFolderId "$($Usercontactfolder.ID)" -UserID $UserID
    }
    
        Write-Debug "Usercontactfolder existiert noch nicht und muss angelegt werden"
        $ContactFolder = $null
        $ContactFolder = New-MgUserContactFolder -UserId $UserID -DisplayName "$($GALSYNCFOLDERNAME)"
        Write-Debug "Usercontactfolder $($GALSYNCFOLDERNAME) wurde angelegt"


$measureresults = @{}

#create contacts with batches Graph API HTTP Request
$BATCHJOBRESULT = Measure-Command {Push-GalSycnUserContacts -Contacts $GAL_Users -UserMail $UserID -ContactFolderID $($ContactFolder.Id) -DebugPreference 'Continue' -accesstoken $($accesstoken) -BatchJob $true}
$measureresults.'Batch' = "$($BATCHJOBRESULT.Seconds) Seconds"


#create contacts with SDK Graph Module
$SDKGRAPHMODULERESULT = Measure-Command {Push-GalSycnUserContacts -Contacts $GAL_Users -UserMail $UserID -ContactFolderID $($ContactFolder.Id) -DebugPreference 'Continue' -accesstoken $($accesstoken) -BatchJob $false}
$measureresults.'Graph Module' = "$($SDKGRAPHMODULERESULT.Seconds) Seconds"

#create contacts with GRaph API HTTP Request
$HTTPGRAPHRESULT = Measure-Command {Push-GalSycnUserContacts -Contacts $GAL_Users -UserMail $UserID -ContactFolderID $($ContactFolder.Id) -DebugPreference 'Continue' -accesstoken $($accesstoken) -BatchJob $false -httprequest}
$measureresults.'Graph HTTP Request' = "$($HTTPGRAPHRESULT.Seconds) Seconds"

$measureresults | fl


#endregion


