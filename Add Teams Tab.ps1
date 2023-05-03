$Tenantname="caje77sharepoint"
$TenantID = ''
$ClientID = ''
$ClientSecret = ''
$redirectURI="https://localhost"
$resource = "https://graph.microsoft.com/"
$siteName="zTestCajeStd1641"
Connect-PnPOnline -Url "https://$TenantName.sharepoint.com/sites/$siteName" -ClientId "ca70390d-9a0e-4290-88b4-9656d22a1340"  -ClientSecret "7KG8Q~~Qg2JYqduYOmUvj08k6NjPAP.NKW3vhaZ4" 

$grid=Get-PnPPropertyBag -Key GroupId

$tenantId="cajesharepoint.onmicrosoft.com"

$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

# Construct Body
$body = @{
    client_id     = "ca70390d-9a0e-4290-88b4-9656d22a1340"
    scope         = "https://graph.microsoft.com/.default"
    client_secret = "7KG8Q~~Qg2JYqduYOmUvj08k6NjPAP.NKW3vhaZ4"
    grant_type    = "client_credentials"
}

# Get OAuth 2.0 Token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

# Access Token
$accessToken = ($tokenRequest.Content | ConvertFrom-Json).access_token


$apiUrl = "https://graph.microsoft.com/v1.0/teams/$grid/channels?$filter=displayName eq 'General'"
##Call the graph
$channelDetails = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken" } -Uri $apiUrl -ContentType 'application/json' -Method Get



$apiUrl = "https://graph.microsoft.com/v1.0/groups/$grId/planner/plans"
##Call the graph - Special Access (It seems you can only access the plans by using user credentials because you can only access the plans to which you have permissions) 
$planDetails = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken"} -Uri $apiUrl -ContentType 'application/json' -Method Get

$apiUrl="https://graph.microsoft.com/v1.0/planner/plans"
$body = @{
     owner= $grid
     title= "title-value" 
    
    }
      
$createplan=Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken"} -Uri $apiUrl -ContentType 'application/json'  -Body ($body | ConvertTo-Json) -Method Post


$web=Get-PnPWeb

#$tab=Get-PnPTeamsTab -Team $team -Channel "General"  -Identity "Planner"
$chan=Get-PnPTeamsChannel -Team $te -Identity "General"
$plan=Get-PnPPlannerPlan -Group $grid
$tenant="cajesharepoint.onmicrosoft.com"
$tab=Add-PnPTeamsTab -Team $team -Channel "General" -DisplayName "Planner" -Type Planner -ContentURL "https://tasks.office.com/$tenant/Home/PlannerFrame?page=7&planId=$($plan.id)"
<#$tab.TeamsAppId="com.microsoft.teamspace.tab.planner"
$tab.Configuration.EntityId="tt.c_"+$chan.Id+"_p_"+$plan.Id
$tab.Configuration.ContentUrl ="https://tasks.office.com/cajesharepoint.onmicrosoft.com/Home/PlannerFrame?page=7&planId=wgZi-BFHYU2Yuvg4ItsjA8gACUw_"
$tab.Configuration.RemoveUrl ="https://tasks.office.com/cajesharepoint.onmicrosoft.com/Home/PlannerFrame?page=7&planId=wgZi-BFHYU2Yuvg4ItsjA8gACUw_"
$tab.Configuration.WebsiteUrl ="https://tasks.office.com/cajesharepoint.onmicrosoft.com/Home/PlannerFrame?page=7&planId=wgZi-BFHYU2Yuvg4ItsjA8gACUw_"#>

<#https://graph.microsoft.com/beta/teams/ccdae018-6f7e-4f7a-ab6d-8f5211cdf903/channels/19:B-psLznM5ZUvrGWR4gZ8EVxAwgT3dHmZPRyLd3_OG_Q1@thread.tacv2/tabs

{
    "name": "Planner1",
    "displayName": "Planner Backlog",
    "teamsAppId": "com.microsoft.teamspace.tab.planner",
    "configuration": {
        "entityId": "9ugIfZ8hIkKLSgnjcOXKMsgAFxWU",ner-
        "contentUrl": "https://tasks.office.com/cajesharepoint.onmicrosoft.com/Home/PlannerFrame?page=7&planId=9ugIfZ8hIkKLSgnjcOXKMsgAFxWU",
        "removeUrl": "https://tasks.office.com/cajesharepoint.onmicrosoft.com/Home/PlannerFrame?page=7&planId=9ugIfZ8hIkKLSgnjcOXKMsgAFxWU",
        "websiteUrl": "https://tasks.office.com/cajesharepoint.onmicrosoft.com/Home/PlannerFrame?page=7&planId=9ugIfZ8hIkKLSgnjcOXKMsgAFxWU"
    }
}#>




