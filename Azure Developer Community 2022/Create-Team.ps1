#MS Graph Application permissions:
#permissions: Team.Create
#permissions: User.Read.All
#permissions: Sites.Manage.All
#permissions: Sites.FullControl.All

#========================================================================
# Global Variables
#========================================================================
#region global variables
$SPOSourceSiteName = "Azure Developer Community Days 2022"
$SPOSourceListName = "Teams Request List"
$DefaultOwner = "Spiderman"
$appID = "ee25ff8e-64e2-4335-8992-0874ecb6336a"
$tenantID = "ac531135-db87-4577-9701-f3351b528e74"
$clientSecret = "Als8Q~e-7C3ZERXWvITNmUGlvR4pNuYMLYxK3aFK"
#endregion

#========================================================================
# Functions
#========================================================================
#region functions

function Get-GraphAuthorizationToken {
    param
    (
        [string]$ResourceURL = 'https://graph.microsoft.com',
        [string][parameter(Mandatory)]
        $TenantID,
        [string][Parameter(Mandatory)]
        $ClientKey,
        [string][Parameter(Mandatory)]
        $AppID
    )
	
    #$Authority = "https://login.windows.net/$TenantID/oauth2/token"
	$Authority = "https://login.microsoftonline.com/$TenantID/oauth2/token"
	
    [Reflection.Assembly]::LoadWithPartialName("System.Web") | Out-Null
    $EncodedKey = [System.Web.HttpUtility]::UrlEncode($ClientKey)
	
    $body = "grant_type=client_credentials&client_id=$AppID&client_secret=$EncodedKey&resource=$ResourceUrl"
	
    # Request a Token from the graph api
    $result = Invoke-RestMethod -Method Post -Uri $Authority -ContentType 'application/x-www-form-urlencoded' -Body $body
	
    $script:APIHeader = @{ 'Authorization' = "Bearer $($result.access_token)" }
}

#========================================================================
function Normalize-String {
    param(
        [Parameter(Mandatory = $true)][string]$str
    )
	
    $str = $str.ToLower()
    $str = $str.Replace(" ", "")
    $str = $str.Replace("ä", "ae")
    $str = $str.Replace("ö", "oe")
    $str = $str.Replace("ü", "ue")
    $str = $str.Replace("ß", "ss")
	
    Write-Output $str
}

#========================================================================
#endregion

#========================================================================
# Scriptstart
#========================================================================
clear

#region get Graph Token for further processing
Get-GraphAuthorizationToken -TenantID $tenantID -AppID $appID -ClientKey $clientSecret
if($script:APIHeader){
    Write-Output "Token from Microsoft Graph acquired!"
    #Write-Output ($script:APIHeader).values
}
else{
    Write-Output "Error acquiring token, quitting job!"
    exit
}
#endregion

#identify teh standard owner in AAD
# more info about Graph Queries: https://learn.microsoft.com/en-us/graph/query-parameters 
$uri = "https://graph.microsoft.com/v1.0/users?`$filter=startswith(displayName,'$($DefaultOwner)')"
$DefaultUser = Invoke-RestMethod -Uri $uri -Method get -Headers $script:APIHeader
$DefaultUserID = $DefaultUser.value.id


#region detect the source list
$uri = "https://graph.microsoft.com/v1.0/sites"
#site
$SPOSite = Invoke-RestMethod -uri $uri -Headers $script:APIHeader -Method Get | Select-Object -ExpandProperty value | Where-Object {$_.DisplayName -ilike "*$($SPOSourceSiteName)*"}
$SPOSiteID = $SPOSite.id.Split(",")[1]
if(!$SPOSiteID){
    $SPOSiteID = $SPOSite.value.ID.Split(",")[1]
}
#get Queue list
$uri = "https://graph.microsoft.com/v1.0/sites/$SPOSiteID/lists/" 
$SPOList = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:APIHeader | Select-Object -ExpandProperty value | Where-Object {$_.displayName -ilike "*$($SPOSourceListName)*"}
$SPOListID = $SPOList.id
if(!$SPOListID){
    $SPOListID = $SPOListID.value.id
}
if(!$SPOListID){
    Write-Output "Queue could not be identified! Exiting job"
    exit
}
#endregion

#region process the requests
$uri = "https://graph.microsoft.com/v1.0/sites/$SPOSiteID/lists/$SPOListID/items?expand=fields"
$Requests = Invoke-RestMethod -Uri $uri -Method Get -Headers $script:APIHeader
if((!$Requests.value) -or ($Requests.value.Count -eq 0)){
    Write-Output "Queue is empty, nothing to do!"
    exit
}

foreach($QueueItem in $Requests.value){
    Write-Output "Processing Share Request '$($QueueItem.fields.Title)' by user '$($QueueItem.createdBy.user.displayName) ($($QueueItem.createdBy.user.email))' "
    Write-Output "============================================================================================================="

    $TeamName = Normalize-String -str $QueueItem.fields.Title

    #specify the attributes of the Team
    $TeamsAttributes = @{
        "Template@odata.bind" = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
        DisplayName = $TeamName
        Description = $QueueItem.fields.'Team Description'
        Members = @(
		@{
			"@odata.type" = "#microsoft.graph.aadUserConversationMember"
			Roles = @(
				"owner"
			)
			"User@odata.bind" = "https://graph.microsoft.com/v1.0/users('$($DefaultUserID)')"
		    }
	    )
    }   #extend with more attributes as needed, e.g. funSettings: https://learn.microsoft.com/en-us/graph/api/resources/team?view=graph-rest-1.0

    #now create the Team
    $uri = "https://graph.microsoft.com/v1.0/teams"
    $newTeam = Invoke-WebRequest -Uri $uri -Method Post -Headers $script:APIHeader -Body ($TeamsAttributes | ConvertTo-Json -Depth 5) -ContentType "application/json; charset=utf-8" -UseBasicParsing
    if(!$newTeam){
        #TODO:
        #add error handling here
        continue
    }

    Write-Output "Team $TeamName successfully created!"

    #delete item in requests list:
    $uri = "https://graph.microsoft.com/v1.0/sites/$SPOSiteID/lists/$SPOListID/items/$($QueueItem.id)"
    Invoke-WebRequest -Uri $uri -Method Delete -Headers $script:APIHeader -UseBasicParsing 
}

#endregion