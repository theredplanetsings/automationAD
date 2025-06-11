Import-Module ActiveDirectory

# constants
$domainname = "paradigmcos.local"
$defaultpassword = ConvertTo-SecureString "Password123@" -AsPlainText -Force

# prompts for user details
$FirstName = Read-Host "Enter First Name"
$LastName = Read-Host "Enter Last Name"
$Office = Read-Host "Enter Physical Delivery Office Name"
$Address = Read-Host "Enter Street Address"
$PostalCode = Read-Host "Enter Postal Code"
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
$fullName = "$FirstName $LastName"

# creates the new AD user in the default OU (paradigmcos.local\Users)
New-ADUser `
    -Name $fullName `
    -GivenName $FirstName `
    -Surname $LastName `
    -SamAccountName $username `
    -UserPrincipalName $userPrincipalName `
    -Path "OU=Users,DC=paradigmcos,DC=local" ` # adjust accordingly if needed
    -AccountPassword $defaultpassword `
    -Enabled $true `
    -PassThru

# will assign security groups (if provided)
if ($GroupsInput) {
    $groupList = $GroupsInput -split ","
    foreach ($grp in $groupList) {
        Add-ADGroupMember -Identity $grp.Trim() -Members $username
    }
}

# sets the additional user properties for new user
Set-ADUser $username `
    -Replace @{
        physicalDeliveryOfficeName = $Office
        streetAddress              = $Address
        postalCode                 = $PostalCode
        telephoneNumber            = $Telephone
        mail                       = $Email
        mailNickname               = $MailNickname
        proxyAddresses             = ($ProxyAddressesIn -split ",")
        title                      = $JobTitle
        department                 = $Department
        company                    = $Company
        name                       = $fullName
    }
