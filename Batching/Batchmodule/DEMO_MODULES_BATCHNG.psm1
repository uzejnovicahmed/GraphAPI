function Request-AccessToken{

    [CmdletBinding()]
    param (
        # The Microsoft Azure AD Tenant Name
        [Parameter(ParameterSetName = 'ClientAuth', Mandatory = $true)]  [string]$TenantName,
        # The Microsoft Azure AD TenantId (GUID or domain)
        [Parameter(ParameterSetName = 'ClientAuth', Mandatory = $true)]  [string]$ClientId,
         # An authentication secret of the Microsoft Azure AD Application Registration
         [Parameter(ParameterSetName = 'ClientAuth', Mandatory = $true)] [string]$ClientSecret
    )
    

    
    $resource = "https://graph.microsoft.com/"  
    
    $tokenBody = @{  
        Grant_Type    = 'client_credentials'  
        Scope         = 'https://graph.microsoft.com/.default'  
        Client_Id     = $ClientId  
        Client_Secret = $clientSecret  
    }  
    
    try{$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $tokenBody -ErrorAction Stop}catch{return "Error generating access token $($_)"}
    Write-Debug "Successfully generated authentication token"
    return $($tokenResponse.access_token)
    }



function Get-GalSyncUsers {

    [CmdletBinding()]
    param (
    # A Azure AD Group to get Only Users if Members in this Group
    [Parameter()][string]$GroupId = $null,
    [Parameter()]$DebugPreference = 'SilentlyContinue'
    )
    #$GroupId = 'ff9a0215-9f50-4f92-abef-a950cf6d55aa'

    $properties = @(

        ,'AccountEnabled'         
        ,'BusinessPhones'          
        ,'City'  
        ,'Country'       
        ,'CompanyName'         
        ,'Department'               
        ,'DisplayName'          
        ,'EmployeeId'         
        ,'FaxNumber'         
        ,'GivenName'            
        ,'Id'         
        ,'Addresses'
        ,'JobTitle'
        ,'title'         
        ,'Mail'         
        ,'MailNickname'           
        ,'Manager'         
        ,'MobilePhone'               
        ,'OfficeLocation'                    
        ,'OnPremisesExtensionAttributes'         
        ,'OnPremisesSamAccountName'          
        ,'OnPremisesSecurityIdentifier'         
        ,'OnPremisesUserPrincipalName'         
        ,'OtherMails'                      
        ,'PostalCode'                  
        ,'ProxyAddresses'             
        ,'Surname'         
        ,'UserPrincipalName'              
        ,'AdditionalProperties'
    )        



    if(([string]::IsNullOrEmpty($GroupId)) -and ([string]::IsNullOrWhiteSpace($GroupId))){ #if Group Parameter is empty get ALL Users

        Write-Debug "Group is empty"
        Write-Debug "Get All Users from Graph Endpoint where EmployeeId is not null"
        $UserObjects = Get-MgUser -All -Property $properties | select $properties
        #$UserObjects = $UserObjects | Where-Object {$_.EmployeeId -ne $null}
        
        Write-Debug "Found $($UserObjects.count) Users with EmployeeId"
    }
    else{# if Group parameter is set get users only member of this group

        Write-Debug "Group is present"
        $Groupmembers = $null
        $Groupmembers = Get-MgGroupMember -GroupId $GroupId        
        $UserObjects = Get-MgUser -Property * | Where-Object {$_.id -in $Groupmembers.ID} | select $properties
        Write-Debug "Users of Group selected and saved as Objects"
    }




$GalUserObjects = @()

foreach($line in $UserObjects)
{

$GalUserObjects += [pscustomobject]@{

   # birthday =  "$($line.Birthday)"
    fileAs =  ""
    displayName = "$($line.displayName)"
    givenName = "$($line.GivenName)"
    initials = ""
    middleName = ""
    nickName = "$($line.MailNickname)"
    surname = "$($line.surname)"
    title = "$($line.personalTitle)"
    yomiGivenName = ""
    yomiSurname = ""
    yomiCompanyName = ""
    generation = ""
    imAddresses = @($line.imAddresses)
    jobTitle = "$($line.JobTitle)"
    companyName = "$($line.CompanyName)"
    department = "$($line.Department)"
    officeLocation = "$($line.OfficeLocation)"
    profession = ""
    businessHomePage = ""
    assistantName = ""
    manager = ""
    homePhones = @()
    mobilePhone = "$($line.MobilePhone)"
    businessPhones = @("$($line.businessPhones)")
    spouseName = ""
    personalNotes = ""
    emailAddresses = @(
@{

            address= "$($line.Mail)"
            name= "$($line.DisplayName)"

}
    )
<#
    homeAddress =@{
      street = "$($line.StreetAddress)"
      city = "$($line.City)"
      state = "$($line.State)"
      countryOrRegion = "$($line.Country)"
      postalCode = "$($line.PostalCode)"
        }
  #>  
    businessAddress =@{
        street = "$($line.StreetAddress)"
        city = "$($line.City)"
        state = "$($line.State)"
        countryOrRegion = "$($line.Country)"
        postalCode = "$($line.PostalCode)"
        }
    <#  
    otherAddress =@{
        street = "$($line.StreetAddress)"
        city = "$($line.City)"
        state = "$($line.State)"
        countryOrRegion = "$($line.Country)"
        postalCode = "$($line.PostalCode)"
        }
    #>
    }
}
#####################################
#####################################
#####################################



return $GalUserObjects

}


Function Push-GalSycnUserContacts{

    [CmdletBinding()]
    param (
    [Parameter()]$Contacts,
    [Parameter()][string]$UserMail,
    [Parameter()][string]$ContactFolderID,
    [Parameter()][string]$accesstoken,
    [Parameter()][bool]$BatchJob = $true,
    [Parameter()][switch]$httprequest,
    [Parameter()]$DebugPreference = 'SilentlyContinue'
    )


    $headers = @{
        "Authorization" = "Bearer $($accesstoken)"
        "Content-type"  = "application/json;charset=utf-8"
        "Accept"        = "application/json"
    }


if($BatchJob -eq $true){

        Write-Debug "Start Create GalSyncUserContacts with Batching"
    
        
    
        $batchurl = "https://graph.microsoft.com/v1.0/`$batch"
        [int]$requestID = 0
        $myBatchRequests = @()

        Write-Debug "Working with $($Contacts.count) Contacts"

        foreach ($Contact in $Contacts) {
    
            $requestID ++

            $myRequest = [pscustomobject][ordered]@{ 
                id      = $($requestID)
                method  = "POST"
                url     = "/users/$($usermail)/contactFolders/$($ContactFolderID)/contacts"
                headers = $($headers)
                body    = $($contact)
            } 
            $myBatchRequests += $myRequest 

        }

        #region batchgroups

        <#
        Create Batchgroups because of  batching limitation of 20 requests at one time (1 Batch Request = 20 Contacts into 1 User)
        #>

        $counter = [pscustomobject] @{ Value = 0 }
        $BatchgroupSize = 20
        $BatchGroups = $($myBatchRequests) | Group-Object -Property { [math]::Floor($counter.Value++ / $BatchgroupSize) }

        #endregion batchgroups

        Write-Debug "$($BatchGroups.count) Batches to Process"

        foreach($BatchGroup in $BatchGroups)
        {

            Write-Debug "Working with Request Group $($BatchGroup.Name)"

            $BatchRequests = [pscustomobject][ordered]@{ 
                requests = $($BatchGroup.Group)
            }

            $batchBody = $BatchRequests | ConvertTo-Json -Depth 6
            $batchBodybytes = [System.Text.Encoding]::UTF8.GetBytes($batchBody)
            try{$BatchRequestResult = Invoke-RestMethod -Method Post -Uri $batchUrl -Body $batchBodybytes -Headers $headers -ErrorAction Stop}catch{Write-Error "Error Invoke Batch Request $($BatchGroup.Name) | Errormessage : $_";$BatchRequestResult.responses}
            Write-Debug "Successfully created Batch $($BatchGroup.Name)"
            $BatchRequestResult.responses | select id, status
        }
        
    

}

if($BatchJob -eq $false){



    if($http.IsPresent)
    {
        Write-Debug "Standard HTTP Query will be executed"

            foreach ($contact in $Contacts) {
                $contactbody = $null
                $contactbody = $contact | ConvertTo-Json -Depth 6
                $contactBodybytes = [System.Text.Encoding]::UTF8.GetBytes($contactbody )
                $URL = "https://graph.microsoft.com/v1.0/users/$($usermail)/contactFolders/$($ContactFolderID)/contacts" 
                try {
                    $HTTPRequestResult = Invoke-RestMethod -Method Post -Uri $URL -Body $contactBodybytes -Headers $headers -ErrorAction Stop
                }
                catch {
                    Write-Error "Error Invoke HTTP Request for Contact $($Contact.DisplayName) | Errormessage : $($_)"
                }

        
            } 
       


    }
    else
    {

            foreach($contact in $Contacts)
            {
                $contactbody = $null
                $contactbody = $contact | ConvertTo-Json -Depth 6
                try{New-MgUserContact -UserId $UserID -BodyParameter $contactbody -ErrorAction Stop}catch{Write-Debug "$_"}
                try{New-MgUserContactFolderContact -ContactFolderId $ContactFolderID -UserId $UserMail -BodyParameter $contactbody -ErrorAction Stop}catch{Write-Debug "$_"}
            }
    }

}




}




Function Update-GalSycnUserContacts{

    [CmdletBinding()]
    param (
    [Parameter()]$Contacts,
    [Parameter()][array]$UserMails,
    [Parameter()][string]$GALSYNCFOLDERNAME = "GAL_SYNC_DEV",
    [Parameter()][string]$accesstoken,
    [Parameter()]$DebugPreference = 'SilentlyContinue'

    )


    foreach($User in $UserMails){


    $Usercontactfolder = Get-MgUserContactFolder -UserID $User | where {$_.DisplayName -eq "$($GALSYNCFOLDERNAME)"}

    if($Usercontactfolder){
            Remove-MgUserContactFolder -ContactFolderId "$($Usercontactfolder.ID)" -UserID $User
        }
        
            Write-Debug "Usercontactfolder existiert noch nicht und muss angelegt werden"
            $ContactFolder = $null
            $ContactFolder = New-MgUserContactFolder -UserId $User -DisplayName "$($GALSYNCFOLDERNAME)"
            Write-Debug "Usercontactfolder $($GALSYNCFOLDERNAME) wurde angelegt"
        
        
        if($ContactFolder){
            Write-Debug "Start to Create GAL SYNC USER CONTACT"
            Push-GalSycnUserContacts -Contacts $Contacts -UserMail $User -ContactFolderID $($ContactFolder.Id) -DebugPreference 'Continue' -accesstoken $($accesstoken) -BatchJob $true
        
        }

    }
    
    }


Function Update-GalSycnUserContactsParallel{ #powershell 7 only

        [CmdletBinding()]
        param (
        [Parameter()]$Contacts,
        [Parameter()][array]$UserMails,
        [Parameter()][string]$GALSYNCFOLDERNAME = "GAL_SYNC_DEV",
        [Parameter()][string]$accesstoken,
        [Parameter()]$DebugPreference = 'SilentlyContinue'
    
        )
    
            Write-Debug "Parrallel Process will be started"
    
            $UserMails | ForEach-Object -Parallel {
    
            #Import-Module "GAL_SYNC.psm1"    
            Import-Module "C:\Users\uzeah\OneDrive - base-IT Gmbh\SMA_DEV\Module_GAL_SYNC\SMA - Codesamples\Exchange Online\GAL_SYNC.psm1" -force -DisableNameChecking      
            Write-Debug "Run Connect-MgGraph"
            Connect-MgGraph -AccessToken $($using:accesstoken)
    
    
            $UserID = $($_)
            Write-Debug "Working with $($UserID)"
            
            $Usercontactfolder = Get-MgUserContactFolder -UserID $($UserID) | where {$_.DisplayName -eq "$($using:GALSYNCFOLDERNAME)"}
            
            if($Usercontactfolder){
                    Remove-MgUserContactFolder -ContactFolderId "$($Usercontactfolder.ID)" -UserID $UserID
                }
                
                    Write-Debug "Usercontactfolder existiert noch nicht und muss angelegt werden"
                    $ContactFolder = $null
                    $ContactFolder = New-MgUserContactFolder -UserId $UserID -DisplayName "$($using:GALSYNCFOLDERNAME)"
                    Write-Debug "Usercontactfolder $($using:GALSYNCFOLDERNAME) wurde angelegt"
                
                
                if($ContactFolder){
                    Write-Debug "Start to Create GAL SYNC USER CONTACT"
                    Push-GalSycnUserContacts -Contacts $($using:Contacts) -UserMail $UserID -ContactFolderID $($ContactFolder.Id) -accesstoken $($using:accesstoken) -BatchJob $true -DebugPreference 'Continue'
                }
        
            } -ThrottleLimit $throttlelimit
    
    
        
        }
    





Function Update-GalSycnUserContactsParallelJobs{ # powershell Jobs

    [CmdletBinding()]
    param (
    [Parameter()]$Contacts,
    [Parameter()][array]$UserMails,
    [Parameter()][string]$GALSYNCFOLDERNAME = "GAL_SYNC_DEV",
    [Parameter()][string]$accesstoken,
    [Parameter()][string]$joblitmit = 15,
    [Parameter()]$DebugPreference = 'Continue'

    )

#region new 

$counter = [pscustomobject] @{ Value = 0 }
$BatchgroupSize = [math]::ceiling(($UserMails.count / $joblitmit))
$BatchGroups = $($UserMails) | Group-Object -Property { [math]::Floor($counter.value ++ / $BatchgroupSize) }



foreach($bgroup in $BatchGroups){


$UserMails = ($bgroup.Group)

Write-Debug "Start Jobs "

Start-Job -Name "JOB $($bgroup.Name)" -ScriptBlock {

    $UserMails = $args[0]
    $Contacts = $args[1]
    $GALSYNCFOLDERNAME = $args[2]
    $accesstoken = $args[3]

    Import-Module "C:\dev_github\DEMO_GRAPH_BATCHING\DEMO_GRAPH_BATCHING\DEMO_MODULES_BATCHNG.psm1" -force -DisableNameChecking      
    $ExecutionContext.SessionState.LanguageMode = "FullLanguage"

    Write-Debug "Run Connect-MgGraph"
    Connect-MgGraph -AccessToken $($accesstoken)


        foreach($user in $UserMails)

        {

                $Usercontactfolder = Get-MgUserContactFolder -UserID $User | where {$_.DisplayName -eq "$($GALSYNCFOLDERNAME)"}
            
                if($Usercontactfolder){
                        Remove-MgUserContactFolder -ContactFolderId "$($Usercontactfolder.ID)" -UserID $User
                    }
                    
                        Write-Debug "Usercontactfolder existiert noch nicht und muss angelegt werden"
                        $ContactFolder = $null
                        $ContactFolder = New-MgUserContactFolder -UserId $User -DisplayName "$($GALSYNCFOLDERNAME)"
                        Write-Debug "Usercontactfolder $($GALSYNCFOLDERNAME) wurde angelegt"


                    if($ContactFolder){
                        Write-Debug "Start to Create GAL SYNC USER CONTACT"
                        Push-GalSycnUserContacts -Contacts $Contacts -UserMail $User -ContactFolderID $($ContactFolder.Id) -DebugPreference 'Continue' -accesstoken $($accesstoken) -BatchJob $true

                    }

                

        }
    
    } -ArgumentList $($usermails),$($Contacts),$($GALSYNCFOLDERNAME),$($accesstoken) | Receive-Job



Write-Debug "Jobs from JOB $($bgroup.Name) Group started"


}

Get-Job | wait-Job | Receive-Job

}






#endregion new


    


Export-ModuleMember Push-GalSycnUserContacts
Export-ModuleMember Request-AccessToken
Export-ModuleMember Get-GalSyncUsers
Export-ModuleMember Update-GalSycnUserContacts
Export-ModuleMember Update-GalSycnUserContactsParallelJobs