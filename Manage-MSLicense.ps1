<#
.SYNOPSIS
Manage-UserLicense.ps1 - This script assigns and removes Microsoft 365 licenses.

.DESCRIPTION 
This script assigns and removes Microsoft 365 licenses to users and groups in your tenant. 
More precisely, it prompts you interactively a list of the options to allow easily to choose the license and the relative service plans
to add / remove, using the friendly name of the product (e.g., you find "Microsoft 365 E5" for "SPE_E5").
The script needs the .csv file "LicenseMappingTable.csv" that must be saved in the same folder of the script.

.INPUTS
You must have installed the Microsoft.Graph
You must have valid credentials for connecting to Microsoft Graph with the permissions of Users Administrator or Groups Administrator.
The Microsoft Graph PowerShell Enterprise Application must have the following delegated API permissions:
    - User.ReadWrite.All
    - Organization.Read.All
    - Group.ReadWrite.All
If not granted, the script prompts you to give them (in this case you must be Global Administrator)

.OUTPUTS
In the shell, you can choose the license that you prefer to assign / remove.
You can choose to generate a report which summarizes the actions (licenses assigned, remaining licenses, etc.) made by the script.

.PARAMETER UserName
Insert the UPN of the single user for which you want manage license

.PARAMETER UserList
Insert the name of the relative path of a .csv file containing the list of users' UPN for which you want to manage the license.
The .csv must have a single column with UserPrincipalName as header.

.PARAMETER GroupName
Insert the Display Name of the single group for which you want manage license

.PARAMETER GroupList
Insert the name of the relative path of a .csv file containing the list of groups' Display Name for which you want to manage the license.
The .csv must have a single column with DisplayName as header.

.EXAMPLE
.\Manage-UserMSLicense.ps1 -UserName john.smith@contoso.com --> Manage licenses to the user john.smith@contoso.com;
.\Manage-UserMSLicense.ps1 -UserList .\Users.csv --> Manage licenses for the users listed in the .csv file Users;
.\Manage-UserMSLicense.ps1 -GroupName Group --> Manage licenses for the group whose Display Name is Group;
.\Manage-UserMSLicense.ps1 -GroupList .\Groups.csv --> Manage licenses for the groups listed in the .csv file Groups;

.NOTES
Written by: Stefano Viti - stefano.viti1995@gmail.com
Follow me at https://www.linkedin.com/in/stefano-viti/

#>

param(
    [Parameter(ParameterSetName='Category')]
    [string]$UserName,

	[Parameter(ParameterSetName='Category')]
    [string]$UserList,

    [Parameter(ParameterSetName='Category')]
    [string]$GroupName,

    [Parameter(ParameterSetName='Category')]
    [string]$GroupList
)


Clear-Host

#Environmental Variables
$AllSKUs = Import-csv -path .\LicensesMappingTable.csv -Encoding UTF8

#Connect to MSGraph
$MgContext = Get-MgContext
if($Null -eq $MgContext){
    try{
        Connect-MgGraph -Scopes User.ReadWrite.All, Organization.Read.All, Group.ReadWrite.All -NoWelcome
        Write-Host "Connected to Microsoft Graph!" -ForegroundColor Green
    }
    catch{
        Write-Host "Connection Failed! Check the synopsis of the scipt to see if you have all the required modules" -ForegroundColor Red
        exit
    }
}
else{
    try{
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Connect-MgGraph -Scopes User.ReadWrite.All, Organization.Read.All -NoWelcome
        Write-Host "Connected to Microsoft Graph!" -ForegroundColor Green
    }
    catch{
        Write-Host "Connection Failed! Check the synopsis of the scipt to see if you have all the required modules" -ForegroundColor Red
        exit
    }
}

#Check which SKUs you have in your tenant and choose which you want to assign
$AvailableSKUs = Get-MgSubscribedSku -All
$AssignableSKUArray = @()
$i = 0
Foreach($AvailableSKU in $AvailableSKUs){
    if($AvailableSKU.AppliesTo -eq "User"){
        $i ++
        $PrepaidLicense = (Get-MgSubscribedSku | Where-Object {$_.SkuId -eq $AvailableSKU.SkuId} | Select-Object -ExpandProperty PrepaidUnits).Enabled
        $ConsumedUnits = (Get-MgSubscribedSku | Where-Object {$_.SkuId -eq $AvailableSKU.SkuId} | Select-Object ConsumedUnits).ConsumedUnits
        $AvailabledLicense = $PrepaidLicense - $ConsumedUnits
        $AssignableSKU = $AllSKUs | Where-Object {$_.GUID -eq $AvailableSKU.SkuId} | Select-Object 'Product Name'
        $hashSKU = [ordered]@{
            ID = $i
            SKU = $AssignableSKU.'Product Name'
            AvailableLicense = $AvailabledLicense
        }
        $SKUItem = New-Object psobject -Property $hashSKU
        $AssignableSKUArray = $AssignableSKUArray + $SKUItem
    }
}

#Choose the license to assign and remove:

Write-Host "-------------------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host "                                                                               " -ForegroundColor Yellow
Write-Host "-------------------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host "In your tenant, the following license are available:" -ForegroundColor Cyan
Foreach($SKU in $AssignableSKUArray){
    Write-Host "[$($SKU.ID)] - $($SKU.SKU)" -ForegroundColor Cyan
}

$AssignIDList = Read-Host "Choose the licenses to assign, separated by a semi-column if more than one (e.g, 1;3;5). Proceed if you do not want to assign any licenses"
$RemoveIDList = Read-Host "Choose the licenses to remove, separated by a semi-column if more than one (e.g, 1;3;5). Proceed if you do not want to remove any licenses"

if($AssignIDList -eq "" -and $RemoveIDList -eq ""){
    Write-Host "You have not selected any license to assign or remove!" -ForegroundColor Red
    exit
}

$AssignIDArray = $AssignIDList -split ";"
$RemoveIDArray = $RemoveIDList -split ";"

#Prepare the table containing the licenses to remove:

$LicenseToRemove = @()
Foreach($RemoveID in $RemoveIDArray){
    $SelectedSKU = $AssignableSKUArray | Where-Object {$_.ID -eq $RemoveID}
    $SKUToRemove = $AllSKUs | Where-Object {$_.'Product name' -eq $SelectedSKU.SKU}
    $RemoveHash = [ordered]@{
        SKUName = $SelectedSKU.SKU
        SKUGUID = $SKUToRemove.GUID
    }
    $RemoveItem = New-Object psobject -Property $RemoveHash
    $LicenseToRemove = $LicenseToRemove + $RemoveItem  
}

#Prepare the Table summing up the licenses and the plans to assign:

Clear-Host

$LicenseToAssign = @()
Foreach($AssignID in $AssignIDArray){
    $SelectedSKU = $AssignableSKUArray | Where-Object {$_.ID -eq $AssignID}
    $SKUToAssign = $AllSKUs | Where-Object {$_.'Product name' -eq $SelectedSKU.SKU}
    $ServicePlans = $SKUToAssign.'Service plans included (friendly names)'.split("|")
    Write-Host "The SKU $($SelectedSKU.SKU) has the following Service Plans:" -ForegroundColor Red -BackgroundColor White
    $j = 0
    $ServicePlansArray = @()
    Foreach($Plan in $ServicePlans){
        $j ++
        $hashPlan = [ordered]@{
            IDPlan = $j
            ServicePlan = $Plan
        }
        $PlanItem = New-Object psobject -Property $hashPlan
        $ServicePlansArray = $ServicePlansArray + $PlanItem     
        Write-Host "[$($j)] - $($Plan)" -ForegroundColor Cyan
    }
    $AssignPlanIDList = Read-Host "Choose the service plans to DISABLE, separated by a semi-column if more than one (e.g, 1;3;5). Press Enter if no plans must be disabled"
    $AssignPlanIDArray = $AssignPlanIDList -split ";" 
    $SelectedPlanList = ""
    $SelectedPlanGUIDList = ""
    Foreach($PlanID in $AssignPlanIDArray){
       $SelectedPlan =  ($ServicePlansArray | Where-Object {$_.IDPlan -eq $PlanID}).ServicePlan
       $SelectedPlanGUID = ($SelectedPlan.Substring($SelectedPlan.Length-38).replace("(","")).replace(")","")
       $SelectedPlanList = $SelectedPlanList + $SelectedPlan + "|"
       $SelectedPlanGUIDList = $SelectedPlanGUIDList + $SelectedPlanGUID + "|"
    }
    $AssignHash = [ordered]@{
        SKUName = $SelectedSKU.SKU
        AvailableLicense = $SelectedSKU.AvailableLicense
        SKUGUID = $SKUToAssign.GUID
        ServicePlansToDisable = $SelectedPlanList.Substring(0,$SelectedPlanList.Length-1)
        ServicePlansGUID = $SelectedPlanGUIDList.Substring(0,$SelectedPlanGUIDList.Length-1)
    }
    $AssignItem = New-Object psobject -Property $AssignHash
    $LicenseToAssign = $LicenseToAssign + $AssignItem
}

#The variable $RemoveToAssign contains the summary of the license to remove
#The variable $LicenseToAssign contains the summary of the license to assign

#Now you can proceed to manage users and/or groups license

Clear-Host

if($Null -ne $UserList){
    $Users = Import-csv -path $UserList -Encoding utf8
    Foreach($User in $Users){
        $UPN = $User.UserPrincipalName
        #Remove the selected licenses
        Foreach($SKU in $LicenseToRemove){
            Write-Host "You are going to remove $($Users.count) $($SKU.SKUName) licenses" -ForegroundColor Cyan
            $RemoveLicense = @(
                @{
                    SkuId = $SKU.SKUGUID
                }
            )
            Set-MgUserLicense -UserID $UPN -AddLicenses @{} -RemoveLicenses $RemoveLicense
        }
        #Add the selected licenses
        Foreach($SKU in $LicenseToAssign){
            if($Users.count -gt $SKU.AvailableLicense){
                Write-Host "You are going to assign more $($SKU.SKUName) license than the available ones!" -ForegroundColor Red
                $Check = Read-Host: "Are you sure you want to proceed? [Y/N] (Default is N)"
                if($Check -ne "Y"){
                    exit
                }
            }
            Write-Host "You are going to assign $($Users.count) $($SKU.SKUName) licenses" -ForegroundColor Cyan
            $AddLicense = @(
                @{
                    SkuId = $SKU.SKUGUID
                    DisabledPlans = $SKU.ServicePlansGUID.split("|")
                }
            ) 
            Set-MgUserLicense -UserID $UPN -AddLicenses $AddLicense -RemoveLicenses @{}
        }
    }
}

if($Null -ne $GroupList){
    $Groups = Import-csv -path $GroupList -Encoding utf8
    Foreach($Group in $Groups){
        $GroupID = (Get-MgGroup -Filter "DisplayName eq '$($Group.DisplayName)'").GroupID
        $GroupMembers = Get-MgGroupMember -All -GroupID $GroupID
        #Remove the selected licenses
        Foreach($SKU in $LicenseToRemove){
            Write-Host "You are going to remove $($GroupMembers.count) $($SKU.SKUName) licenses" -ForegroundColor Cyan
            $RemoveLicense = @(
                @{
                    SkuId = $SKU.SKUGUID
                }
            )
            Set-MgGroupLicense -UserID $GroupID -AddLicenses @{} -RemoveLicenses $RemoveLicense
        }
        #Add the selected licenses
        Foreach($SKU in $LicenseToAssign){
            if($GroupMembers.count -gt $SKU.AvailableLicense){
                Write-Host "You are going to assign more $($SKU.SKUName) license than the available ones!" -ForegroundColor Red
                $Check = Read-Host: "Are you sure you want to proceed? [Y/N] (Default is N)"
                if($Check -ne "Y"){
                    exit
                }
            }
            Write-Host "You are going to assign $($GroupMembers.count) $($SKU.SKUName) licenses" -ForegroundColor Cyan
            $AddLicense = @(
                @{
                    SkuId = $SKU.SKUGUID
                    DisabledPlans = $SKU.ServicePlansGUID.split("|")
                }
            ) 
            Set-MgGroupLicense -UserID $GroupID -AddLicenses $AddLicense -RemoveLicenses @{}
        }
    }
}

if($Null -ne $UserName){
    $UPN = $UserName
    Foreach($SKU in $LicenseToRemove){
        $RemoveLicense = @(
            @{
                SkuId = $SKU.SKUGUID
            }
        )
        Set-MgUserLicense -UserID $UPN -AddLicenses @{} -RemoveLicenses $RemoveLicense
    }
    #Add the selected licenses
    Foreach($SKU in $LicenseToAssign){
        $AddLicense = @(
            @{
                SkuId = $SKU.SKUGUID
                DisabledPlans = $SKU.ServicePlansGUID.split("|")
            }
        ) 
        Set-MgUserLicense -UserID $UPN -AddLicenses $AddLicense -RemoveLicenses @{}
    }
}

if($Null -ne $GroupName){
    $GroupID = (Get-MgGroup -Filter "DisplayName eq '$($GroupName)'").GroupID
    $GroupMembers = Get-MgGroupMember -All -GroupID $GroupID
    #Remove the selected licenses
    Foreach($SKU in $LicenseToRemove){
        Write-Host "You are going to remove $($GroupMembers.count) $($SKU.SKUName) licenses" -ForegroundColor Cyan
        $RemoveLicense = @(
            @{
                SkuId = $SKU.SKUGUID
            }
        )
        Set-MgGroupLicense -UserID $GroupID -AddLicenses @{} -RemoveLicenses $RemoveLicense
    }
    #Add the selected licenses
    Foreach($SKU in $LicenseToAssign){
        Write-Host "You are going to assign $($GroupMembers.count) $($SKU.SKUName) licenses" -ForegroundColor Cyan
        $AddLicense = @(
            @{
                SkuId = $SKU.SKUGUID
                DisabledPlans = $SKU.ServicePlansGUID.split("|")
            }
        ) 
        Set-MgGroupLicense -UserID $GroupID -AddLicenses $AddLicense -RemoveLicenses @{}
    }
}