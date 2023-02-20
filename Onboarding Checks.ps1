
# PowerShell function to get read details from configuration file
# Author: Cajetan Trindade
$logFile = "Onboarding_Checks" + $TenantName + $(Get-Date).Year + ".log"

function Get-TimeStamp {
    
    return "{0:MM/dd/yy} {0:HH:mm:ss}" -f (Get-Date)
    
}
Function WriteLog([string]$logText, [switch]$ForceDisplay) {
    $datelogtext = Get-Date -Format "MM/dd/yyyy HH:mm:ss" #Format the Date

    $datelogText += " " + $logText
    if ($forceDisplay) {
        Write-Host    $datelogText 
    }
    $datelogText | out-file $logFile -Append
    }





Function Get-Config {
    $datappids = Get-Content -Path config.json | ConvertFrom-Json 

    $dataccounts = Get-Content -Path config.json | ConvertFrom-Json
    $dataccounts | ForEach $_ { $_.Accounts
        $TenantName = $_.Accounts.TenantName
        $ServiceAccount = $_.Accounts.ServiceAccount
        $FlowAccount = $_.Accounts.FlowAccount
        $B2B = $_.Accounts.B2BAccount
        $BAAccount = $_.Accounts.ECMAccount
    } 

    $siteURL = "https://$TenantName-admin.sharepoint.com/"
 #Get-GroupingNameRequirementandPnPManagementShell $TenantName

 Get-UserDetails $TenantName $ServiceAccount $FlowAccount $BAAccount
 Check-SPOTenantDetails $TenantName
 Restrict-GroupCreation $ServiceAccount
    

    $apphigh = $TenantName + 'High'
    $credhigh=Add-PnPStoredCredential -Name $apphigh -Username $appCredential.UserName -Password $appCredential.Password 


    $appCredential = Get-Credential -Message "Add IWorkplace Medium Details"
    $appmedium = $TenantName + 'Medium'
    Add-PnPStoredCredential -Name $appmedium -Username $appCredential.UserName -Password $appCredential.Password

    $appCredential = Get-Credential -Message "Add IWorkplace Low Details"
    $applow = $TenantName + 'Low'
    Add-PnPStoredCredential -Name $applow -Username $appCredential.UserName -Password $appCredential.Password#>


    $datappids = Get-Content config.json | ConvertFrom-Json
    $datappids.AppId.'IWorkplaceHigh'.storedcred = $apphigh
    $datappids | ConvertTo-Json | set-content config.json

    $datappids = Get-Content config.json | ConvertFrom-Json
    $datappids.AppId.'IWorkplaceMedium'.storedcred = $appmedium
    $datappids | ConvertTo-Json | set-content config.json


    $datappids = Get-Content config.json | ConvertFrom-Json
    $datappids.AppId.'IWorkplaceLow'.storedcred = $applow
    $datappids | ConvertTo-Json | set-content config.json

    <#$storedAppCredentials = Get-PnPStoredCredential -Name $apphigh#>
    $clientIDHigh = $storedAppCredentials.UserName


    <#$storedAppCredentials = Get-PnPStoredCredential -Name $appmedium#>
    $clientIDMedium = $storedAppCredentials.UserName

<#    $storedAppCredentials = Get-PnPStoredCredential -Name $applow#>
    $clientIDLow = $storedAppCredentials.UserName



    #$data1=Get-Content -Path config.json | ConvertFrom-Json
    WriteLog ("This will run a check through to see if all checks are done for IWorkplace Apps for" + $TenantName) -ForceDisplay

 
 #Get-UserDetails $TenantName $ServiceAccount
 <#   TestAppSecret1 -Site $TenantName -Creds $apphigh
    TestAppSecret1 -Site $TenantName -Creds $appmedium
    TestAppSecret1 -Site $TenantName -Creds $applow#>

   <# PermissionAppIds -PermissionType Application -AppId $clientIdHigh -Type High
    PermissionAppIds -PermissionType Application -AppId  $clientIdMedium -Type Medium    
    PermissionAppIds -PermissionType Application -AppId $clientIdLow -Type Low   #>
    Start-Sleep -s 5
    
    





    $datappids | format-List
    
    
     
    #Get-B2BDetails $TenantName  $B2B 
   
   
    Get-ServicePrincipalDetails
   
   
   
}

<#Function PermissionAppIds1 {
    
    param(

        [ValidateSet('Any', 'Delegated', 'Application')]
        [string] $PermissionType,
        [string] $AppId,
        [string] $Type

    )
    Connect-MgGraph
    $servicePrincipal = Get-MgServicePrincipal -top 999 | ? { $_.AppId -like $AppId }
    if ($Type -eq "High") {
        $requiredPermissions = 'Calendars.ReadWrite' , 'Directory.ReadWrite.All', 'Domain.Read.All', 'Files.ReadWrite.All', 'Group.Create', 'Group.Read.All', 'Group.ReadWrite.All', 'GroupMember.ReadWrite.All', 'Notes.ReadWrite.All', 'Schedule.ReadWrite.All', 'Sites.FullControl.All', 'TeamsAppInstallation.ReadWriteForTeam.All', 'User.ReadWrite.All' | Find-MgGraphPermission -PermissionType Application -Exact | Select-Object Id
        $requiredPermissionsDel = 'Group.ReadWrite.All' | Find-MgGraphPermission -PermissionType  Delegated -Exact | Select-Object Id
        $requiredPermissionsSP = @("741f803b-c850-494e-b5df-cde7c675a1ca", "c8e3537c-ec53-43b9-bed3-b2bd3617ae97", "678536fe-1083-478a-9c59-b99265e6b0d3")
        $requiredPermissionsOn = @("10af711e-5051-4838-8ae7-9767c43d8a9c")
       
        $currentpermissionsSP = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.id | ? { $_.resourceDisplayName -like "Office 365 SharePoint Online" } | select -expand AppRoleId
        $currentpermissionsGraph = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.id | ? { $_.resourceDisplayName -like "Microsoft Graph" } | select -expand AppRoleId
        $currentpermissionsON = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.id | ? { $_.resourceDisplayName -like "OneNote" } | select -expand AppRoleId
             
        $csv = Import-Csv .\GraphPermsions.csv | group -AsHashTable -Property Key

        WriteLog("Checking High App Permissions")-ForceDisplay
        foreach ($currentpermissionON in $currentpermissionsOn) {
            if ($requiredPermissionsOn -cmatch $currentpermissionON) {
                WriteLog ("One Note Permission exist" + $currentpermissionON )
            }
            else {
                WriteLog ("One Note Permission does not exist" + $currentpermissionON) -ForceDisplay
            }
        }

        foreach ($currentpermissionSP in $currentpermissionsSP) {
            if ($requiredPermissionsSP -cmatch $currentpermissionSP) {
                WriteLog ("Sharepoint Permission exist" + $currentpermissionSP) 
            }
            else {
                WriteLog ("Some Sharepoint Permission does not exist" + $currentpermissionSP)-ForceDisplay
            }
        }

              
        foreach ($currentpermissionGraph in $currentpermissionsGraph) {
            if ($requiredPermissions -cmatch $currentpermissionGraph) {
                WriteLog ("Graph Permission exist" + $currentpermissionGraph) 
            }
            else {
                WriteLog ("Some Graph Permission does not exist" + $currentpermissionGraph) -ForceDisplay
            }
        }
          
    }
          
    if ($Type -eq "Medium") {
        $requiredPermissionsSP = @("741f803b-c850-494e-b5df-cde7c675a1ca", "c8e3537c-ec53-43b9-bed3-b2bd3617ae97", "678536fe-1083-478a-9c59-b99265e6b0d3")
        $requiredPermissions = 'Files.ReadWrite.All', 'Group.Read.All', 'GroupMember.Read.All', 'Sites.FullControl.All', 'User.Read.All', 'Sites.FullControl.All', 'TermStore.ReadWrite.All', 'User.ReadWrite.All' | Find-MgGraphPermission -PermissionType  Application -Exact | Select-Object Id
        $currentpermissionsSP = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.id | ? { $_.resourceDisplayName -like "Office 365 SharePoint Online" } | select -expand AppRoleId
        $currentpermissionsGraph = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.id | ? { $_.resourceDisplayName -like "Microsoft Graph" } | select -expand AppRoleId
        WriteLog("Checking Medium App Permissions")-ForceDisplay
        foreach ($currentpermissionGraph in $currentpermissionsGraph) {
            if ($requiredPermissions -cmatch $currentpermissionGraph) {
                WriteLog ("Graph Permission exist" + $currentpermissionGraph) 
            }
            else {
                WriteLog ("Some Graph Permission does not exist" + $currentpermissionGraph) -ForceDisplay
            }
        }

        foreach ($currentpermissionSP in $currentpermissionsSP) {
            if ($requiredPermissionsSP -cmatch $currentpermissionSP) {
                WriteLog ("Sharepoint Permission exist" + $currentpermissionSP) 
            }
            else {
                WriteLog ("Some Sharepoint Permission does not exist" + $currentpermissionSP)-ForceDisplay
            }
        }
          
    }
    if ($Type -eq "Low") {
        $requiredPermissionsSP = @("9bff6588-13f2-4c48-bbf2-ddab62256b36")

        $currentpermissionsSP = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.id | ? { $_.resourceDisplayName -like "Office 365 SharePoint Online" } | select -expand AppRoleId
           
        WriteLog("Checking Low App Permissions")-ForceDisplay

        foreach ($currentpermissionSP in $currentpermissionsSP) {
            if ($requiredPermissionsSP -cmatch $currentpermissionSP) {
                WriteLog ("Sharepoint Permission exist" + $currentpermissionSP) 
            }
            else {
                WriteLog ("Some Sharepoint Permission does not exist" + $currentpermissionSP)-ForceDisplay
            }
        }
          
    }
  

           


}#>

Function Check-SPOTenantDetails {
    param ($TenantName)

    $adminUrl = "https://$TenantName-admin.sharepoint.com/"
    $rootsite = "https://$TenantName.sharepoint.com/"

    Connect-PnPOnline $AdminURL -Interactive
 
    #Get the Tenant Site Object
    $Site = Get-PnPTenantSite -Url $rootsite
  
    if ($Site.DenyAddAndCustomizePages -eq "Disabled") {
        WriteLog("Root Site Scripting is enabled")-ForceDisplay
    }
    else {
        WriteLog("Error: Root Site Scripting is not enabled")-ForceDisplay
    }

    Disconnect-PnPOnline




    Connect-SPOService -Url $adminUrl; 
    try {
        $spo = Get-SPOTenant 

        if (($spo.LegacyAuthProtocolsEnabled) -eq $true) {
            WriteLog('All good with respect to Legacy Authentication') -ForceDisplay
        }
        else {
            WriteLog("Need to enable legacy authentication by running the following Set-SPOTenant -LegacyAuthProtocolsEnabled $true;") -ForceDisplay;
        }

        if (($spo.DisableCustomAppAuthentication) -eq $false) {
            WriteLog('All good with respect to DisableCustomAppAuthentication')
        }
        else {
            WriteLog("Need to disable custom app authentication by running the following Set-SPOTenant -DisableCustomAppAuthentication $false;;")-ForceDisplay;
        }
    }
    catch {
        WriteLog ("Error:") -ForceDisplay -ForegroundColor red;
        WriteLog ($_) -ForceDisplay -ForegroundColor red;
    }
    Disconnect-SPOService;
    Start-Sleep -s 10
}


Function Get-GroupingNameRequirementandPnPManagementShell {
    param ($TenantName)
 Connect-AzureAD 
    WriteLog("Getting group naming requirement") -ForceDisplay
    Start-Sleep -s 5
    $settings = Get-AzureADDirectorySetting  | where-object { $_.displayname -eq "Group.Unified"
    }
    WriteLog("Group Naming Requirement: " + $settings[
        "PrefixSuffixNamingRequirement"
        ]) -ForceDisplay
        $ServicePrincipalList = Get-AzureADServicePrincipal -Filter "AppId eq '31359c7f-bd7e-475c-86db-fdb8c937548e'"
        if($ServicePrincipalList)
  {
  WriteLog("PNp Management Shell Registered") -ForceDisplay
  }
  else
  {
  WriteLog("PNp Management Shell Needs to be  Registered") -ForceDisplay
  }
}

Function Get-UserDetails {
     param ($TenantName, $ServiceAccount, $FlowAccount, $BAAccount)
     #param ($TenantName, $ServiceAccount)
 
    Connect-AzureAD #-AccountId $ServiceAccount
    WriteLog("Getting user account details");

    $accounts = @($ServiceAccount, $FlowAccount, $BAAccount)
      #$accounts = @($ServiceAccount)

    foreach ($account in $accounts) {
        $licensePlanList = Get-AzureADSubscribedSku
        # $user=Get-AzureADUser -UserPrincipalName $_.Accounts 
        try {

            $user = Get-AzureADUser -ObjectId $account -ErrorAction Continue

  
            if ($user) {
                WriteLog($user.UserPrincipalName + " exist")
                $roleNames = @("SharePoint Administrator", "Teams Administrator")
                WriteLog("Getting details for" + $user.UserPrincipalName)
                Start-Sleep 5;
                foreach ($roleName in $roleNames) {
                    $role = Get-AzureADDirectoryRole | Where { $_.displayName -eq $roleName }
                    $spdamins = Get-AzureADDirectoryRole | Where { $_.DisplayName -eq $roleName } | Get-AzureADDirectoryRoleMember 
                    foreach ($spdamin in $spdamins) {
                        if ($user.ObjectId -eq $spdamin.ObjectId) {
                            WriteLog($roleName + "exist for" + $user.UserPrincipalName) -ForceDisplay
                            #break;
                        }
                      
                       
                    }
                    
                }

                $pwdexpire = Get-AzureADUser -ObjectId $user.UserPrincipalName | Select-Object UserprincipalName, @{
                    N = "PasswordNeverExpires"; E = { $_.PasswordPolicies -contains "DisablePasswordExpiration"
                    }
                }
                if ($pwdexpire.PasswordNeverExpires -eq $false) {
                    WriteLog('Warning:We recommend to disable password expiration for the accounts' + $user.UserPrincipalName + 'To set the password of one user to never expire, run the following cmdlet by using the UPN or the user ID of the user:') -ForceDisplay
                }
                else {
                    WriteLog('Password expiration policy all good for' + $user.UserPrincipalName) -ForceDisplay
                }

                $userList = Get-AzureADUser -ObjectID $user.UserPrincipalName | Select -ExpandProperty AssignedLicenses | Select SkuID 

                if ($userList) {
                    $userList | ForEach { $sku = $_.SkuId ; $licensePlanList | ForEach { If ( $sku -eq $_.ObjectId.substring($_.ObjectId.length - 36,
                                    36) ) {
                                WriteLog ("License Exist for user" + $user.UserPrincipalName + "License Name" + $_.SkuPartNumber)-ForceDisplay
                            }
                        }
                    }
                    # WriteLog("License Exist" + $userList) -ForceDisplay
                }
                else {
                    WriteLog("Error: License Does Not Exist for" + $user.UserPrincipalName) -ForceDisplay
                }
            }
            else {
                WriteLog("Error: " + $user.UserPrincipalName + " does not exist") -ForceDisplay
            }
        }
        catch {
            WriteLog ("Error:") -ForceDisplay -ForegroundColor red;
            WriteLog ($_) -ForceDisplay -ForegroundColor red;
        }
    }
}

Function Get-B2BDetails {
    param ($TenantName, $B2B)
    $gr1 = (Get-AzureADGroup  | Where { $_.DisplayName -eq $B2B } -ErrorAction SilentlyContinue ) 

    if ($gr1) {
        WriteLog("Group Exist" + $gr1.DisplayName) -ForceDisplay
    }
    else {
        WriteLog("Error group: " + $gr1.DisplayName + "does not exist") -ForceDisplay
    }
}




function TestAppSecret1 { 
    Param(
        [Parameter(Mandatory = $true, Position = 1)] [string]$site,
        [Parameter(Mandatory = $true, Position = 2)] [string]$Creds
    ) 
 
    $rootsite = "https://$TenantName.sharepoint.com/"

    $storedAppCredentials = Get-PnPStoredCredential -Name $Creds
    $clientID = $storedAppCredentials.UserName
    $clientSecret = [System.Net.NetworkCredential]::new("", $storedAppCredentials.Password).Password

    Connect-PnPOnline -Url $rootsite  -ClientId $clientID -ClientSecret $clientSecret
  
    $ctx = Get-PnPContext;
    try {
        $web = Get-PnPWeb;
        WriteLog ("App ID:}" -f $clientID) -ForegroundColor White -ForceDisplay
        WriteLog ("Context OK: {0}" -f $ctx.Url) -ForceDisplay
        WriteLog ("Web access OK: `n" -f $web.Url) -ForceDisplay
        Start-Sleep 2;
    }  
    catch {
        WriteLog ("App ID: {0}" -f $clientID) -ForceDisplay
        writeLog ("Error getting context for: {0}`n" -f $Site) -ForceDisplay
        Start-Sleep 2;
    }
    finally {
        Disconnect-PnPOnline;
    }
}


Function Get-ServicePrincipalDetails {
 
    Connect-AzureAD
    $Applications = Get-AzureADApplication -SearchString "IWorkplace"

    foreach ($app in $Applications) {
        $AppName = $app.DisplayName
       
        $ObjId = $app.objectid
        $ApplID = $app.AppId
        $AppCreds = Get-AzureADApplication -ObjectId $ObjId | select PasswordCredentials, KeyCredentials
        $secret = $AppCreds.PasswordCredentials
        $cert = $AppCreds.KeyCredentials

        foreach ($s in $secret) {
            $StartDate = $s.StartDate
            
            $EndDate = $s.EndDate
            WriteLog("Application Name:" + $AppName + "Ends at:" + $EndDate) -ForceDisplay
   
        }
    }
}


Function Restrict-GroupCreation {
    param ($ServiceAccount)
    Connect-AzureAD

    $settingsObjectID = (Get-AzureADDirectorySetting | Where-object -Property Displayname -Value "Group.Unified" -EQ).id

    if ($settingsObjectID) {
        $grp = (Get-AzureADDirectorySetting -Id $settingsObjectID).Values
        foreach ($props in $grp) {
            if ($props.Name -eq "GroupCreationAllowedGroupId" -and $props.Value -ne $null) {
                WriteLog( $props.Name + " " + $props.Value) -ForceDisplay

                $group = (Get-AzureADGroup  | Where { $_.ObjectId -eq $props.Value } -ErrorAction SilentlyContinue ) 

                $grpmembers = Get-AzureADGroupMember -ObjectId $props.Value | Where-Object { $_.UserPrincipalName -eq $ServiceAccount } 

                if ($grpmembers) {
                    WriteLog($ServiceAccount + " exist in the Group " + $group.DisplayName + "  which restricts Group Creation") -ForceDisplay
                }
                else {
                    WriteLog("Warning" + $ServiceAccount + " does not exist in the Group" + $group.DisplayName + " which restricts Group Creation") -ForceDisplay
                }

      


            }

        }
    }
}


Function PermissionAppIds {
    
    param(

        [ValidateSet('Any', 'Delegated', 'Application')]
        [string] $PermissionType,
        [string] $AppId,
        [string] $Type

    )

    $data = Import-Csv .\GraphPermsions.csv
    $table = $data | Group-Object -AsHashTable -AsString -Property Key
    Connect-MgGraph
    $servicePrincipal = Get-MgServicePrincipal -top 999 | ? { $_.AppId -like $AppId }
    if ($Type -eq "High") {
        $requiredPermissions = 'Calendars.ReadWrite' , 'Directory.ReadWrite.All', 'Domain.Read.All', 'Files.ReadWrite.All', 'Group.Create', 'Group.Read.All', 'Group.ReadWrite.All', 'GroupMember.ReadWrite.All', 'Notes.ReadWrite.All', 'Schedule.ReadWrite.All', 'Sites.FullControl.All', 'TeamsAppInstallation.ReadWriteForTeam.All', 'User.ReadWrite.All' | Find-MgGraphPermission -PermissionType Application -Exact | Select-Object Id
        $requiredPermissionsDel = 'Group.ReadWrite.All' | Find-MgGraphPermission -PermissionType  Delegated -Exact | Select-Object Id
        $requiredPermissionsSP = @("741f803b-c850-494e-b5df-cde7c675a1ca", "c8e3537c-ec53-43b9-bed3-b2bd3617ae97", "678536fe-1083-478a-9c59-b99265e6b0d3")
        $requiredPermissionsOn = @("10af711e-5051-4838-8ae7-9767c43d8a9c")
       
        $currentpermissionsSP = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.id | ? { $_.resourceDisplayName -like "Office 365 SharePoint Online" } | select -expand AppRoleId
        $currentpermissionsGraph = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.id | ? { $_.resourceDisplayName -like "Microsoft Graph" } | select -expand AppRoleId
        $currentpermissionsON = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.id | ? { $_.resourceDisplayName -like "OneNote" } | select -expand AppRoleId
             
      
        WriteLog("Checking High App Permissions")-ForceDisplay
        foreach ($permissionON in $requiredPermissionsOn) {
            if ( $currentpermissionsON -ccontains ($permissionON)) {
                WriteLog("One Note Permission" + ($table.$permissionON)."value" + "is present")
            }
            else {
                WriteLog("Error: One Note Permission" + ($table.$permissionON)."value" + "is not present") -ForceDisplay
            }
        }

        foreach ($permissionSP in $requiredPermissionsSP) {
            if ( $currentpermissionsSP -ccontains ($permissionSP)) {
                WriteLog("SP Permission" + ($table.$permissionSP)."value" + "is present")
            }
            else {
                WriteLog("SP Permission" + ($table.$permissionSP)."value" + "is not present") -ForceDisplay
            }
        }

        foreach ($permissionGraph in $requiredPermissions) {
             $permgid=($permissionGraph).Id
            if ( $currentpermissionsGraph -ccontains $permgid) {
                WriteLog("Graph Permission" + ($table.$permgid)."value" + "is present")
            }
            else {
                WriteLog("Graph Permission" + ($table.$permgid)."value" + "is not present") -ForceDisplay
            }
        }
          
    }
          
    if ($Type -eq "Medium") {
        $requiredPermissionsSP = @("741f803b-c850-494e-b5df-cde7c675a1ca", "c8e3537c-ec53-43b9-bed3-b2bd3617ae97", "678536fe-1083-478a-9c59-b99265e6b0d3")
        $requiredPermissions = 'Files.ReadWrite.All', 'Group.Read.All', 'GroupMember.Read.All', 'Sites.FullControl.All', 'User.Read.All' | Find-MgGraphPermission -PermissionType  Application -Exact | Select-Object Id
        $currentpermissionsSP = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.id | ? { $_.resourceDisplayName -like "Office 365 SharePoint Online" } | select -expand AppRoleId
        $currentpermissionsGraph = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.id | ? { $_.resourceDisplayName -like "Microsoft Graph" } | select -expand AppRoleId
        WriteLog("Checking Medium App Permissions")-ForceDisplay
        foreach ($permissionSP in $requiredPermissionsSP) {
            if ( $currentpermissionsSP -ccontains ($permissionSP)) {
                WriteLog("SP Permission" + ($table.$permissionSP)."value" + "is present")
            }
            else {
                WriteLog("SP Permission" + ($table.$permissionSP)."value" + "is not present") -ForceDisplay
            }
        }

        foreach ($permissionGraph in $requiredPermissions) {
             $permgid=($permissionGraph).Id
            if ( $currentpermissionsGraph -ccontains $permgid) {
                WriteLog("Graph Permission" + ($table.$permgid)."value" + "is present") -ForceDisplay
            }
            else {
                WriteLog("Graph Permission" + ($table.$permgid)."value" + "is not present") -ForceDisplay
            }
        }
          
    }
          

if ($Type -eq "Low") {
    $requiredPermissionsSP = @("9bff6588-13f2-4c48-bbf2-ddab62256b36")

    $currentpermissionsSP = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.id | ? { $_.resourceDisplayName -like "Office 365 SharePoint Online" } | select -expand AppRoleId
           
    WriteLog("Checking Low App Permissions")-ForceDisplay

    foreach ($permissionSP in $requiredPermissionsSP) {
        if ( $currentpermissionsSP -ccontains ($permissionSP)) {
            WriteLog("SP Permission" + ($table.$permissionSP)."value" + "is present")
        }
        else {
            WriteLog("SP Permission" + ($table.$permissionSP)."value" + "is not present") -ForceDisplay
        }
    }
          
}
  
}




Function CheckPermissionAppIds {
    
    param(

       
        [string] $AppId,
        [string] $Type

    )

    $data = Import-Csv .\GraphPermsions.csv
    $table = $data | Group-Object -AsHashTable -AsString -Property Key
    Connect-MgGraph
    $servicePrincipal = Get-MgServicePrincipal -top 999 | ? { $_.AppId -like $AppId }
    if ($Type -eq "High") {
        $requiredPermissionsApp =  'Calendars.ReadWrite' , 'Directory.ReadWrite.All', 'Domain.Read.All', 'Sites.FullControl.All','Files.ReadWrite.All','Group.Create','Group.Read.All','GroupMember.ReadWrite.All','Group.ReadWrite.All','Notes.ReadWrite.All','Schedule.ReadWrite.All','TeamsAppInstallation.ReadWriteForTeam.All','User.ReadWrite.All' |Microsoft.Graph.Authentication\Find-MgGraphPermission -PermissionType Application -Exact  | Select-Object Id,Name
        $requiredPermissionsDel =  'Group.ReadWrite.All'|Microsoft.Graph.Authentication\Find-MgGraphPermission -PermissionType Delegated -Exact  | Select-Object Id,Name
        $requiredPermissionsGraph = $requiredPermissionsApp+ $requiredPermissionsDel

        $requiredPermissionsSP=@("c8e3537c-ec53-43b9-bed3-b2bd3617ae97","741f803b-c850-494e-b5df-cde7c675a1ca", "678536fe-1083-478a-9c59-b99265e6b0d3")
        $requiredPermissionsON=@("10af711e-5051-4838-8ae7-9767c43d8a9c")

       
        $currentpermissionsSP = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.id | ? { $_.resourceDisplayName -like "Office 365 SharePoint Online" } | select -expand AppRoleId
        $currentpermissionsGraph = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.id | ? { $_.resourceDisplayName -like "Microsoft Graph" } | select -expand AppRoleId
        $currentpermissionsON = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.id | ? { $_.resourceDisplayName -like "OneNote" } | select -expand AppRoleId
             
      
        WriteLog("Checking High App Permissions")-ForceDisplay
        foreach ($permissionON in $requiredPermissionsOn) {
            if ( $currentpermissionsON -ccontains ($permissionON)) {
                WriteLog("One Note Permission" + ($table.$permissionON)."value" + "is present")
            }
            else {
                WriteLog("Error: One Note Permission" + ($table.$permissionON)."value" + "is not present") -ForceDisplay
            }
        }

        foreach ($permissionSP in $requiredPermissionsSP) {
            if ( $currentpermissionsSP -ccontains ($permissionSP)) {
                WriteLog("SP Permission" + ($table.$permissionSP)."value" + "is present")
            }
            else {
                WriteLog("SP Permission" + ($table.$permissionSP)."value" + "is not present") -ForceDisplay
            }
        }

        foreach ($permissionGraph in $requiredPermissions) {
             $permgid=($permissionGraph).Id
            if ( $currentpermissionsGraph -ccontains $permgid) {
                WriteLog("Graph Permission" + ($table.$permgid)."value" + "is present")
            }
            else {
                WriteLog("Graph Permission" + ($table.$permgid)."value" + "is not present") -ForceDisplay
            }
        }
          
    }
          
    if ($Type -eq "Medium") {
        $requiredPermissionsSP = @("741f803b-c850-494e-b5df-cde7c675a1ca", "c8e3537c-ec53-43b9-bed3-b2bd3617ae97", "678536fe-1083-478a-9c59-b99265e6b0d3")
        $requiredPermissions = 'Files.ReadWrite.All', 'Group.Read.All', 'GroupMember.Read.All', 'Sites.FullControl.All', 'User.Read.All' | Find-MgGraphPermission -PermissionType  Application -Exact | Select-Object Id
        $currentpermissionsSP = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.id | ? { $_.resourceDisplayName -like "Office 365 SharePoint Online" } | select -expand AppRoleId
        $currentpermissionsGraph = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.id | ? { $_.resourceDisplayName -like "Microsoft Graph" } | select -expand AppRoleId
        WriteLog("Checking Medium App Permissions")-ForceDisplay
        foreach ($permissionSP in $requiredPermissionsSP) {
            if ( $currentpermissionsSP -ccontains ($permissionSP)) {
                WriteLog("SP Permission" + ($table.$permissionSP)."value" + "is present")
            }
            else {
                WriteLog("SP Permission" + ($table.$permissionSP)."value" + "is not present") -ForceDisplay
            }
        }

        foreach ($permissionGraph in $requiredPermissions) {
             $permgid=($permissionGraph).Id
            if ( $currentpermissionsGraph -ccontains $permgid) {
                WriteLog("Graph Permission" + ($table.$permgid)."value" + "is present") -ForceDisplay
            }
            else {
                WriteLog("Graph Permission" + ($table.$permgid)."value" + "is not present") -ForceDisplay
            }
        }
          
    }
          

if ($Type -eq "Low") {
    $requiredPermissionsSP = @("9bff6588-13f2-4c48-bbf2-ddab62256b36")

    $currentpermissionsSP = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.id | ? { $_.resourceDisplayName -like "Office 365 SharePoint Online" } | select -expand AppRoleId
           
    WriteLog("Checking Low App Permissions")-ForceDisplay

    foreach ($permissionSP in $requiredPermissionsSP) {
        if ( $currentpermissionsSP -ccontains ($permissionSP)) {
            WriteLog("SP Permission" + ($table.$permissionSP)."value" + "is present")
        }
        else {
            WriteLog("SP Permission" + ($table.$permissionSP)."value" + "is not present") -ForceDisplay
        }
    }
          
}
  
}
           







Get-Config