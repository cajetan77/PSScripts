Connect-PnpOnline -Url "https://caje77sharepoint.sharepoint.com/sites/ReadAccess" -Interactive
<#Content fropm template site#>
$searchv4Results=Get-PnPPageComponent -Page "Home1" |  Where-Object { $_.PropertiesJson -like'*Case*'}
$searchv4Filter=Get-PnPPageComponent -Page "Home1" |  Where-Object { $_.Title -like'PnP - Search Filters'}
<#Content from client site#>
$searchv3Results=Get-PnPPageComponent -Page "Home" |  Where-Object { $_.PropertiesJson -like'*Case*'}
$searchv3Filter=Get-PnPPageComponent -Page "Home" |  Where-Object { $_.Title -like'Search Filters'}

<#get location of search v3#>
$sectionResults= $searchv3Results.Section.Order
$columnResults=$searchv3Results.Column.Order


$sectionFilter= $searchv3Filter.Section.Order
$columnFilter=$searchv3Filter.Column.Order

if($searchv3Results)
{

<#Remove searchv3 results#>
Remove-PnPClientSideComponent -Page "Home" -InstanceId $searchv3Results.InstanceId -Force
}

if($searchv3Filter)
{
<#Remove searchv3 filter#>
Remove-PnPClientSideComponent -Page "Home" -InstanceId $searchv3Filter.InstanceId -Force
}
<#Add v4 search web parts#>
$newWebPartResults= Add-PnPPageWebPart -Page "Home" -Component "PnP - Search Results" -Section $sectionResults -Column $columnResults
$newWebPartFilter= Add-PnPPageWebPart -Page "Home" -Component "PnP - Search Filters" -Section $sectionFilter -Column $columnFilter
<#get datasource reference to connect search results and filter web part#>
$oldfilterds=$searchv4Filter.WebPartId+"."+$searchv4Filter.InstanceId.Guid
$filterds=$newWebPartFilter.WebPartId+"."+$newWebPartFilter.InstanceId.Guid
$searchV4resultsJson=$searchv4Results.PropertiesJson.Replace($oldfilterds,$filterds)


$oldresultds=$searchv4Results.WebPartId+"."+$searchv4Results.InstanceId.Guid
$resultds=$newWebPartResults.WebPartId+"."+$newWebPartResults.InstanceId.Guid
$searchV4FilterJson=$searchv4Filter.PropertiesJson.Replace($oldresultds,$resultds)

if($newWebPartResults)
{
Set-PnPPageWebPart -Page Home -Identity $newWebPartResults.InstanceId -PropertiesJson $searchv4Results.PropertiesJson
Set-PnPPageWebPart -Page Home -Identity $newWebPartResults.InstanceId -PropertiesJson $searchV4resultsJson
}



if($newWebPartFilter)
{
Set-PnPPageWebPart -Page Home -Identity $newWebPartFilter.InstanceId -PropertiesJson $searchv4Filter.PropertiesJson
Set-PnPPageWebPart -Page Home -Identity $newWebPartFilter.InstanceId -PropertiesJson $searchV4filterJson
}



