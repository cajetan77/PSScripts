Connect-PnpOnline -Url "https://ilcnz.sharepoint.com/sites/cTasman" -Interactive
<#Content fropm template site#>
$searchv4Results=Get-PnPPageComponent -Page "Tasman District Council" |  Where-Object { $_.PropertiesJson -like'*Case*' -and $_.Title -eq 'PnP - Search Results'}
$searchv4Filter=Get-PnPPageComponent -Page "Tasman District Council" |  Where-Object { $_.Title -like'PnP - Search Filters'}
<#Content from client site#>
Connect-PnpOnline -Url "https://ilcnz.sharepoint.com/sites/cToyotaNZ" -Interactive
$web=Get-PnpWeb
$searchv3Results=Get-PnPPageComponent -Page $web.Title |  Where-Object { $_.PropertiesJson -like'*Case*' -and $_.Title -eq 'Search Results'}
$searchv3Filter=Get-PnPPageComponent -Page $web.Title |  Where-Object { $_.Title -like'Search Refiners'}

<#get location of search v3#>
$sectionResults= $searchv3Results.Section.Order
$columnResults=$searchv3Results.Column.Order


$sectionFilter= $searchv3Filter.Section.Order
$columnFilter=$searchv3Filter.Column.Order

if($searchv3Results)
{

<#Remove searchv3 results#>
Remove-PnPClientSideComponent -Page $web.Title -InstanceId $searchv3Results.InstanceId -Force
}

if($searchv3Filter)
{
<#Remove searchv3 filter#>
Remove-PnPClientSideComponent -Page $web.Title -InstanceId $searchv3Filter.InstanceId -Force
}
<#Add v4 search web parts#>
$newWebPartResults= Add-PnPPageWebPart -Page $web.Title -Component "PnP - Search Results" -Section $sectionResults -Column $columnResults
$newWebPartFilter= Add-PnPPageWebPart -Page $web.Title-Component "PnP - Search Filters" -Section $sectionFilter -Column $columnFilter
<#get datasource reference to connect search results and filter web part#>
$oldfilterds=$searchv4Filter.WebPartId+"."+$searchv4Filter.InstanceId.Guid
$filterds=$newWebPartFilter.WebPartId+"."+$newWebPartFilter.InstanceId.Guid
$searchV4resultsJson=$searchv4Results.PropertiesJson.Replace($oldfilterds,$filterds)


$oldresultds=$searchv4Results.WebPartId+"."+$searchv4Results.InstanceId.Guid
$resultds=$newWebPartResults.WebPartId+"."+$newWebPartResults.InstanceId.Guid
$searchV4FilterJson=$searchv4Filter.PropertiesJson.Replace($oldresultds,$resultds)

if($newWebPartResults)
{
Set-PnPPageWebPart -Page $web.Title  -Identity $newWebPartResults.InstanceId -PropertiesJson $searchv4Results.PropertiesJson
Set-PnPPageWebPart -Page $web.Title -Identity $newWebPartResults.InstanceId -PropertiesJson $searchV4resultsJson
}



if($newWebPartFilter)
{
Set-PnPPageWebPart -Page "PrimePort Timaru"  -Identity $newWebPartFilter.InstanceId -PropertiesJson $searchv4Filter.PropertiesJson
Set-PnPPageWebPart -Page "PrimePort Timaru"  -Identity $newWebPartFilter.InstanceId -PropertiesJson $searchV4filterJson
}

