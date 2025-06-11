Import-Module ActiveDirectory
#global constants
$domainname = "paradigmcos.com"
$company = "Paradigm Companies"
$defaultpassword = ConvertTo-SecureString "Password123@" -AsPlainText -Force
$newUsername = "jdoe"

# step 1: creating a new user in the AD [placeholders used for properties until actual values are known]
New-ADUser  `
    -Name "John Doe" `
    -AccountPassword $defaultpassword `
    -GivenName "John" `
    -Surname "Doe"  `
    -SamAccountName $newUsername `
    -UserPrincipalName "$newUsername@$domainname" `
    -Path "OU=Internal,OU=Users,OU=PDC-SERVICES,DC=paradigmcos,DC=local" `
    -Enabled $true
    #-PassThru

# step 2: assigning specified security groups to the user [currently using placeholders until actual property names are known]
Add-ADGroupMember -Identity "VPN Users" -Members "$newUsername"
# if we want to unassign a security group from a user:
# Remove-ADGroupMember -Identity "IT-Security" -Members "jdoe" -Confirm:$false

# step 3: assign additional user properties [currently using placeholders until actual property names are known]
Set-ADUser `
    -Identity "$newUsername" `
    -Office "Main Office" `
    -OfficePhone "123-456-7890" `
    -EmailAddress "$newUsername@$domainname" `
    -Replace @{
        mailNickname = "jdoe"
        proxyAddresses = "smtp:$newUsername@$domainname"
        title = "IT Specialist"
        department = "IT"
        company = "$company"
    }

# modify variables & attributes as needed for further testing
