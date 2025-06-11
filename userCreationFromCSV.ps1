Import-Module ActiveDirectory
# global constants
$domainname = "paradigmcos.local"
# default password for new users
$defaultpassword = ConvertTo-SecureString "Password123@" -AsPlainText -Force
# import CSV; ensure your CSV file has the correct headers such as:
# FirstName, LastName, Office, Address, PostalCode, Telephone, Email, MailNickname, ProxyAddresses, JobTitle, Department, Company, Groups
$users = Import-Csv -Path "C:\Users\$env:username\Downloads\it-tools\INSERT_FILE_NAME.csv"

foreach ($user in $users) {
    # produces username: first letter of first name + last name, all in lowercase
    $username = "$($user.FirstName.Substring(0,1))$($user.LastName)".ToLower()
    $userPrincipalName = "$username@$domainname"
    $fullName = "$($user.FirstName) $($user.LastName)"

    # creating new user in the domain's Users OU
    New-ADUser `
        -Name $fullName `
        -GivenName $user.FirstName `
        -Surname $user.LastName `
        -SamAccountName $username `
        -UserPrincipalName $userPrincipalName `
        -Path "OU=Users,DC=paradigmcos,DC=local" `
        -AccountPassword $defaultpassword `
        -Enabled $true `
        -PassThru

    # assigning specified security groups to the user (CSV Groups column should be comma separated)
    if ($user.Groups) {
        $groupList = $user.Groups -split ","
        foreach ($grp in $groupList) {
            Add-ADGroupMember -Identity $grp.Trim() -Members $username
        }
    }

    # setting additional user properties for new user
    Set-ADUser $username `
        -Replace @{
            physicalDeliveryOfficeName = $user.Office
            streetAddress              = $user.Address
            postalCode                 = $user.PostalCode
            telephoneNumber            = $user.Telephone
            mail                       = $user.Email
            mailNickname               = $user.MailNickname
            proxyAddresses             = ($user.ProxyAddresses -split ",")
            title                      = $user.JobTitle
            department                 = $user.Department
            company                    = $user.Company
            name                       = $fullName
        }
}
