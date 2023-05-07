Connect-AzureAD

$appTitle = @("iWorkplace Library Builder 365",
"iWorkplace Awhina",
"iWorkplace Smart Records",
"iWorkplace Columns",
"iWorkplace Smart Case Files",
"iWorkplace Smart Provisioning",
"iWorkplace Smart Metadata Enterprise",
"iWorkplace Template Central 365")

$clientIds = @("8051d25d-fb2a-46a3-8b97-4ded8deb1d19",
"392cffc8-168a-4ca7-8758-228b86050e94",
"636b6508-b5b4-4332-9a38-81443add3e2c",
"154eca3c-6617-4b82-adb5-69b713fb19f2",
"802b6a3d-584f-44d5-a1b5-8abcf39df50e",
"a8c9df43-4431-4172-9ad6-f2f030e252ff",
"6abf605b-312d-4ffa-8f5a-7234eac8cb31",
"58aded1f-7524-47ba-a0b0-4a28c2bdcd4e");

$clientSecret = ("",
"",
"",
"",
"",
"",
""
);

$startDate = Get-Date
$endDate = Get-Date -Year 2026 -Month 04 -Day 30 -Hour 11 -Minute 59 -Second 59
$appDomain = @("lb365.iworkplace.net",
"awhina.iworkplace.net",
"smartrecords.iworkplace.net",
"columns.iworkplace.net",
"scf365.iworkplace.net",
"sia365.iworkplace.net",
"sme365.iworkplace.net",
"tc365.iworkplace.net");
$redirectUri = ("https://lb365.iworkplace.net/pages/default.aspx",
"https://awhina.iworkplace.net/pages/default.aspx",
"https://smartrecords.iworkplace.net/pages/default.aspx",
"https://columns.iworkplace.net/pages/default.aspx",
"https://scf365.iworkplace.net/pages/default.aspx",
"https://sia365.iworkplace.net/pages/default.aspx",
"https://sme365.iworkplace.net/pages/default.aspx",
"https://tc365.iworkplace.net/Pages/TCConfig.aspx");
$i=-1
foreach($clientId in $clientIds)
{
$i=$i+1
$app = Get-AzureADServicePrincipal -filter ("AppId eq '$($clientId)'")

if($app) 
{
$oldKeyCreds = $app.KeyCredentials
$oldPWCreds = $app.PasswordCredentials
$servicePrincipalName=$clientId + "/"+$appDomain[$i]
    try {
       
New-AzureADServicePrincipalKeyCredential -ObjectId $app.ObjectId -StartDate $startDate -EndDate $endDate -Value $clientSecret[$i] -Usage Verify -Type Symmetric
New-AzureADServicePrincipalKeyCredential -ObjectId $app.ObjectId -StartDate $startDate -EndDate $endDate -Value $clientSecret[$i] -Usage Sign -Type Symmetric
New-AzureADServicePrincipalPasswordCredential -ObjectId $app.ObjectId -EndDate $endDate -Value $clientSecret[$i]

    }
    Catch {
        Write-Host "Failed to register $($appDetails.Title)" -ForegroundColor Red;
        Write-Host "Reason: $($_.Exception.Message)" -ForegroundColor Red;
    }
    foreach($key in $oldKeyCreds){
    Remove-AzureADServicePrincipalKeyCredential -KeyId $key.KeyId -ObjectId $app.ObjectId
}

foreach($key in $oldPWCreds){
    Remove-AzureADServicePrincipalPasswordCredential -KeyId $key.KeyId -ObjectId $app.ObjectId
}
Write-Host("Updated app " +$appTitle[$i])
}
}