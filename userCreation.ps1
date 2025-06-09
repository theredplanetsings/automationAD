Import-Module ActiveDirectory
#global constants
$domainname = "paradigmcos.com"
$company = "Paradigm Companies"
$defaultpassword = ConvertTo-SecureString "Password123@" -AsPlainText -Force

# step 1: creating a new user in the AD [placeholders used for properties until actual values are known]
New-ADUser `
    -Name "John Doe" `
    -GivenName "John" `
    -Surname "Doe" `
    -SamAccountName "jdoe" `
    -UserPrincipalName "jdoe@$domainname" `
    -Path "OU=Users,DC=domain,DC=com" `
    -AccountPassword $defaultpassword `
    -Enabled $true `
    -PassThru | Out-Null

# step 2: assigning specified security groups to the user [currently using placeholders until actual property names are known]
Add-ADGroupMember -Identity "IT-Security" -Members "jdoe"
# if we want to unassign a security group from a user:
# Remove-ADGroupMember -Identity "IT-Security" -Members "jdoe" -Confirm:$false

# step 3: assign additional user properties [currently using placeholders until actual property names are known]
Set-ADUser "jdoe" `
    -Office "Main Office" `
    -OfficePhone "123-456-7890" `
    -EmailAddress "jdoe@$domainname" `
    -Replace @{
        mailNickname = "jdoe"
        proxyAddresses = "smtp:jdoe@$domainname"
        title = "IT Specialist"
        department = "IT"
        company = "$company"
    }

# modify variables & attributes as needed for further testing
