<# contains scripts to add/view/and remove licenses from users using either msonline module or ms entra module **untested as of now** #>

<#
# MS Online version:
Install-Module MSOnline
# connect to the service
Connect-MsolService
# sign in with a global admin credential for your Office 365 tenant

# for more details on a specific license, we can use:
#$licenses | Where-Object { $_.AccountSkuId -eq "yourTenant:SKUName" } | Format-List

# assign a license to an existing user
# Replace 'yourTenant:SKUName' with the correct AccountSkuId
$userPrincipalName = "jdoe@paradigmcos.com" 
$licenseToAdd = "yourTenant:SKUName"
# to check subscriptions/licenses and get their official names, run:
$licenses = Get-MsolAccountSku
$licenses | Format-List

#  the user already has some licenses and you only intend to add the specified license, use -AddLicenses
Set-MsolUserLicense -UserPrincipalName $userPrincipalName -AddLicenses $licenseToAdd

# verify if the license assignment by retrieving user's licenses
Get-MsolUser -UserPrincipalName $userPrincipalName | Select-Object DisplayName, Licenses
#>

## MS Entra version:
#https://learn.microsoft.com/en-us/powershell/entra-powershell/how-to-manage-user-licenses?view=entra-powershell
Install-Module -Name Microsoft.Entra -Repository PSGallery -Scope AllUsers -Force -AllowClobber
#install submodule for user management
Install-Module -Name Microsoft.Entra.Users -Repository PSGallery -Force -AllowClobber
#verify the installation version
Get-InstalledModule -Name Microsoft.Entra* | Where-Object { $_.Name -notmatch "Beta" } |
Format-Table Name, Version, InstalledLocation -AutoSize
#enable user.readwrite.all and organization.read.all permissions
Connect-Entra -Scopes 'User.ReadWrite.All','Organization.Read.All'
#To set a user's location, run:
Set-EntraUser -UserId 'jdoe@paradigmcos.com' -UsageLocation 'US'

## review available entra licenses
Get-EntraSubscribedSku | Select-Object -Property Sku*, ConsumedUnits -ExpandProperty PrepaidUnits

<#
## find and audit users with specific licenses
# get the SKU ID for EMSPREMIUM license plan
$skuId = (Get-EntraSubscribedSku | Where-Object { $_.SkuPartNumber -eq 'EMSPREMIUM' }).SkuId

# find users who have this license
$usersWithLicense = Get-EntraUser -All | Where-Object {
    $_.AssignedLicenses -and 
    ($_.AssignedLicenses.SkuId -contains $skuId)
}
# display the results
$usersWithLicense | Select-Object DisplayName, UserPrincipalName, AccountEnabled |
    Format-Table -AutoSize

## view licenses assigned to a specific user
$userLicenses = Get-EntraUserLicenseDetail -UserId 'jdoe@paradigmcos.com'
$userLicenses
#>


## assign license to a user
# get user details
$user = Get-EntraUser -UserId 'jdoe@paradigmcos.com'

# defines the license plan to assign to the user
$license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
$license.SkuId = (Get-EntraSubscribedSku | Where-Object { $_.SkuPartNumber -eq 'AAD_PREMIUM_P2' }).SkuId

$licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
$licenses.AddLicenses = $license

# assigns the license to the user
Set-EntraUserLicense -UserId $user.Id -AssignedLicenses $licenses

#verify the license assignment
Get-EntraUserLicenseDetail -UserId 'jdoe@paradigmcos.com'


<#
## assign multiple licenses to a user
# retrieves the SkuId for the desired license plans
$skuId1 = (Get-EntraSubscribedSku | Where-Object { $_.SkuPartNumber -eq 'AAD_PREMIUM_P2' }).SkuId
$skuId2 = (Get-EntraSubscribedSku | Where-Object { $_.SkuPartNumber -eq 'EMS' }).SkuId

# gets the user to assign the licenses to
$user = Get-EntraUser -UserId 'jdoe@paradigmcos.com'

# creates license assignment objects
$license1 = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
$license1.SkuId = $skuId1

$license2 = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
$license2.SkuId = $skuId2

$licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
$licenses.AddLicenses = $license1, $license2

# assign the licenses to the user
Set-EntraUserLicense -UserId $user.Id -AssignedLicenses $licenses
#>

<#
## assign license to user by copying license from another user
# defines the source and target users
$licensedUser = Get-EntraUser -UserId 'jdoe@paradigmcos.com'
$targetUser = Get-EntraUser -UserId 'jsmith@paradigmcos.com' 

# retrieves the source user and their licenses
$sourceUserLicenses = $licensedUser.AssignedLicenses

# creates license assignment objects for each license and assign them to the target user
$licensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses 
foreach ($license in $sourceUserLicenses) {
    $assignedLicense = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
    $assignedLicense.SkuId = $license.SkuId
    $licensesToAssign.AddLicenses= $assignedLicense
    Set-EntraUserLicense -UserId $targetUser.Id -AssignedLicenses $licensesToAssign
}
#>



<#
## remove user's licenses
# gets user details
$user = Get-EntraUser -UserId 'jsmith@paradigmcos.com'

# gets the license assigned to the user
$skuId = (Get-EntraUserLicenseDetail -UserId $user.Id).SkuId

# define the license object
$licensesToRemove = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
$licensesToRemove.RemoveLicenses = $skuId
# removes the assigned license
Set-EntraUserLicense -UserId $user.Id -AssignedLicenses $licensesToRemove
#>
