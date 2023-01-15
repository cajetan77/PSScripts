Connect-PnpOnline -Url "https://ilcnz.sharepoint.com/sites/cGreatSouth" -Interactive
<#Content fropm template site#>
$sourceweb=Get-PnpWeb
 $searchv4paginations=Get-PnPPageComponent -Page $sourceweb.Title |  Where-Object { $_.Title -eq 'Search Pagination' -and $_.Section.Order -eq 1 -and $_.Column.Order -eq 1}
foreach($searchv4pagination in $searchv4paginations )
    {
     if(($searchv4pagination.Section.Order -eq 1)  -and ($searchv4pagination.Column.Order -eq 1) )
     {
      Remove-PnPClientSideComponent -Page $sourceweb.Title -InstanceId $searchv4pagination.InstanceId -Force
     }
    }
$searchv4Results=Get-PnPPageComponent -Page $sourceweb.Title |  Where-Object { $_.PropertiesJson -like'*Case*' -and $_.Title -eq 'PnP - Search Results'}
$searchv4Filter=Get-PnPPageComponent -Page $sourceweb.Title |  Where-Object { $_.Title -like'PnP - Search Filters'}
Disconnect-PnPOnline
<#Content from client site#>
Connect-PnpOnline -Url "https://ilcnz.sharepoint.com/sites/cSeafoodNZ" -Interactive
$web=Get-PnpWeb
 $SitePage = Get-PnPListItem -List "Site Pages" -Id 1
    ForEach ($Page in $SitePages) {
    }
$searchv3Results=Get-PnPPageComponent -Page $SitePage.FieldValues.FileLeafRef |  Where-Object { $_.PropertiesJson -like'*Case*' -and $_.Title -eq 'Search Results'}
$searchv3Filter=Get-PnPPageComponent -Page $SitePage.FieldValues.FileLeafRef |  Where-Object { $_.Title -like'Search Refiners'}

<#get location of search v3#>
$sectionResults= $searchv3Results.Section.Order
$columnResults=$searchv3Results.Column.Order


$sectionFilter= $searchv3Filter.Section.Order
$columnFilter=$searchv3Filter.Column.Order

 $searchpaginations=Get-PnPPageComponent -Page $SitePage.FieldValues.FileLeafRef |  Where-Object { $_.Title -eq 'Search Pagination' -and $_.Section.Order -eq 1 -and $_.Column.Order -eq 1}
foreach($searchpagination in $searchpaginations )
    {
     if(($searchpagination.Section.Order -eq 1)  -and ($searchpagination.Column.Order -eq 1) )
     {
      Remove-PnPClientSideComponent -Page $SitePage.FieldValues.FileLeafRef -InstanceId $searchpagination.InstanceId -Force
     }
    }



if($searchv3Results)
{

<#Remove searchv3 results#>
Remove-PnPClientSideComponent -Page $SitePage.FieldValues.FileLeafRef -InstanceId $searchv3Results.InstanceId -Force
}

if($searchv3Filter)
{
<#Remove searchv3 filter#>
Remove-PnPClientSideComponent -Page $SitePage.FieldValues.FileLeafRef -InstanceId $searchv3Filter.InstanceId -Force
}
<#Add v4 search web parts#>
$newWebPartResults= Add-PnPPageWebPart -Page $SitePage.FieldValues.FileLeafRef -Component "PnP - Search Results" -Section $sectionResults -Column $columnResults
$newWebPartFilter= Add-PnPPageWebPart -Page $SitePage.FieldValues.FileLeafRef -Component "PnP - Search Filters" -Section $sectionFilter -Column $columnFilter
<#get datasource reference to connect search results and filter web part#>
$oldfilterds=$searchv4Filter.WebPartId+"."+$searchv4Filter.InstanceId.Guid
$filterds=$newWebPartFilter.WebPartId+"."+$newWebPartFilter.InstanceId.Guid
$searchV4resultsJson=$searchv4Results.PropertiesJson.Replace($oldfilterds,$filterds)


$oldresultds=$searchv4Results.WebPartId+"."+$searchv4Results.InstanceId.Guid
$resultds=$newWebPartResults.WebPartId+"."+$newWebPartResults.InstanceId.Guid
$searchV4FilterJson=$searchv4Filter.PropertiesJson.Replace($oldresultds,$resultds)

#$webpartv4 = Get-PnPClientSideComponent -Page $web.Title -InstanceId 7abe2ee5-9d45-4225-8ed5-6ccf67a84bcc 
$customerName=Get-PnPPropertyBag -Key "ILDS.Case"


if($newWebPartResults)
{
Set-PnPPageWebPart -Page $SitePage.FieldValues.FileLeafRef  -Identity $newWebPartResults.InstanceId -PropertiesJson $searchv4Results.PropertiesJson
Set-PnPPageWebPart -Page $SitePage.FieldValues.FileLeafRef -Identity $newWebPartResults.InstanceId -PropertiesJson $searchV4resultsJson
Set-PnPClientSideWebPart -Page $SitePage.FieldValues.FileLeafRef -Identity $newWebPartResults.InstanceId -PropertiesJson $searchV4resultsJson.Replace("Great South",$customerName)
}



if($newWebPartFilter)
{
Set-PnPPageWebPart -Page $SitePage.FieldValues.FileLeafRef  -Identity $newWebPartFilter.InstanceId -PropertiesJson $searchv4Filter.PropertiesJson
Set-PnPPageWebPart -Page $SitePage.FieldValues.FileLeafRef  -Identity $newWebPartFilter.InstanceId -PropertiesJson $searchV4filterJson
}
