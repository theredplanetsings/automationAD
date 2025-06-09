Import-Module ActiveDirectory
#constants
$domainname = "paradigmcos.com"
$defaultpassword = ConvertTo-SecureString "Password123@" -AsPlainText -Force

# prompts for user details
$FirstName = Read-Host "Enter First Name"
$LastName = Read-Host "Enter Last Name"
$Office = Read-Host "Enter Office"
$Telephone = Read-Host "Enter Telephone Number"
$Email = Read-Host "Enter Email Address"
$MailNickname = Read-Host "Enter Mail Nickname"
$ProxyAddressesIn = Read-Host "Enter Proxy Addresses (comma separated, e.g. smtp:addr1,smtp:addr2)"
$JobTitle = Read-Host "Enter Job Title"
$Department = Read-Host "Enter Department"
$Company = Read-Host "Enter Company"
$GroupsInput = Read-Host "Enter comma-separated Security Groups (if any, leave blank if none)"

# produce derived values
$username = "$($FirstName.Substring(0,1))$LastName".ToLower()
$userPrincipalName = "$username@$domainname"

# creates the new AD user
New-ADUser `
    -Name "$FirstName $LastName" `
    -GivenName $FirstName `
    -Surname $LastName `
    -SamAccountName $username `
    -UserPrincipalName $userPrincipalName `
    -Path "OU=Users,DC=domain,DC=com" ` #adjust accordingly
    -AccountPassword $defaultpassword `
    -Enabled $true `
    -PassThru | Out-Null

# will assign security groups (if provided)
if ($GroupsInput) {
    $groupList = $GroupsInput -split ","
    foreach ($grp in $groupList) {
        Add-ADGroupMember -Identity $grp.Trim() -Members $username
    }
}

# sets the additional user properties for new user
Set-ADUser $username `
    -Office $Office `
    -OfficePhone $Telephone `
    -EmailAddress $Email `
    -Replace @{
        mailNickname   = $MailNickname
        proxyAddresses = ($ProxyAddressesIn -split ",")
        title          = $JobTitle
        department     = $Department
        company        = $Company
    }
