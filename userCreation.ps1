<# Basic script to test user creation/addition to security groups/user property customisation#>
Import-Module ActiveDirectory

# global constants/hardcoded values for testing purposes
$domainname = "paradigmcos.com"
$company = "Paradigm Companies"
$defaultpassword = ConvertTo-SecureString "Password123@" -AsPlainText -Force
$newUsername = "jdoe"
$firstname = "John"
$lastname = "Doe"
$displayname = "$firstname $lastname"
$physicalDeliveryOfficeName = "Corporate"
$streetAddress = "123 Main St"
$postalCode = "12345"
$telephoneNumber = "123-456-7890"
$mail = "$newUsername@$domainname"
$mailNickname = $newUsername
$proxyAddresses = "smtp:$newUsername@$domainname"
$title = "IT Specialist"
$department = "IT"
$l = "CityName" # City
$st = "VA" # state (abbreviated)

# creating the new user with starter properties
New-ADUser  `
    -Name $displayname `
    -AccountPassword $defaultpassword `
    -GivenName $firstname `
    -Surname $lastname `
    -DisplayName $displayname `
    -SamAccountName $newUsername `
    -UserPrincipalName "$newUsername@$domainname" `
    -ChangePasswordAtLogon $true `
    -Path "OU=Internal,OU=Users,OU=PDC-SERVICES,DC=paradigmcos,DC=local" `
    -Enabled $true

# assigning the user to security groups
Add-ADGroupMember -Identity "VPN Users" -Members $newUsername
# (To remove a user from a group, you can use
# Remove-ADGroupMember -Identity "IT-Security" -Members $newUsername -Confirm:$false)

# additional user attributes to be set after initial creation
Set-ADUser `
    -Identity $newUsername `
    -Office $physicalDeliveryOfficeName `
    -OfficePhone $telephoneNumber `
    -EmailAddress $mail `
    -Replace @{
        streetAddress = $streetAddress
        postalCode = $postalCode
        telephoneNumber = $telephoneNumber
        mail = $mail
        mailNickname = $mailNickname
        proxyAddresses = $proxyAddresses
        title = $title
        department = $department
        company = $company
        l = $l
        st = $st
    }

# modify variables & attributes as needed for further testing