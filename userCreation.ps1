Import-Module ActiveDirectory
#global constants
$domainname = "paradigmcos.local"
$company = "Paradigm Services" #edit as needed
$defaultpassword = ConvertTo-SecureString "Password123@" -AsPlainText -Force
$mailaddy = "jdoe@paradigmcos.com"
# step 1: creating a new user in the AD [placeholders used for properties until actual values are known]
New-ADUser `
    -Name "John Doe" `
    -GivenName "John" `
    -Surname "Doe" `
    -sAMAccountName "jdoe" `
    -mailNickname "jdoe" `
    -mail $mailaddy `
    -UserPrincipalName "jdoe@paradigmcos.com" `
    -Path "OU=Users,DC=paradigmcos,DC=local" ` # adjust accordingly if needed
    -AccountPassword $defaultpassword `
    -Enabled $true `
    -PassThru

$proxyaddy = "SMTP:$mailaddy"
# step 2: assigning specified security groups to the user [currently using placeholders until actual property names are known]
Add-ADGroupMember -Identity "IT Staff" -Members "jdoe"
# if we want to unassign a security group from a user:
# Remove-ADGroupMember -Identity "IT Staff" -Members "jdoe" -Confirm:$false

# step 3: assign additional user properties [currently using placeholders until actual property names are known]
Set-ADUser $username `
    -Replace @{
        physicalDeliveryOfficeName = "Corporate"
        streetAddress              = "123 Jane Doe Lane"
        postalCode                 = "12345"
        telephoneNumber            = "123-456-7891"
        mailNickname               = "jdoe"
        proxyAddresses             = $proxyaddy
        title                      = "IT Staff"
        department                 = "Admin"
        company                    = $company
        st                         = "VA"
    }
# modify variables & attributes as needed for further testing
