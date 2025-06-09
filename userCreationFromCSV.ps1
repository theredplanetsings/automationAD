Import-Module ActiveDirectory
# global constants
$domainname = "paradigmcos.com"
# default password for new users
$defaultpassword = ConvertTo-SecureString "Password123@" -AsPlainText -Force
# import CSV; ensure your CSV file has the correct headers
$users = Import-Csv -Path "C:\Users\$env:username\Downloads\it-tools\INSERT_FILE_NAME.csv"

foreach ($user in $users) {
    # produces username: first letter of first name + last name, all in lowercase
    $username = "$($user.FirstName.Substring(0,1))$($user.LastName)".ToLower()
    $userPrincipalName = "$username@$domainname"

    # creating new user in the AD
    New-ADUser `
        -Name "$($user.FirstName) $($user.LastName)" `
        -GivenName $user.FirstName `
        -Surname $user.LastName `
        -SamAccountName $username `
        -UserPrincipalName $userPrincipalName `
        -Path "OU=Users,DC=domain,DC=com" `  # Update this path as needed.
        -AccountPassword $defaultpassword `
        -Enabled $true `
        -PassThru

    # assigning specified security groups to the user
    # expecting comma separated values in "groups" column of CSV
    if ($user.Groups) {
        $groupList = $user.Groups -split ","
        foreach ($grp in $groupList) {
            Add-ADGroupMember -Identity $grp.Trim() -Members $username
        }
    }

    # setting additional user properties
    Set-ADUser $username `
        -Office $user.Office `
        -OfficePhone $user.Telephone `
        -EmailAddress $user.Email `
        -Replace @{
            mailNickname = $user.MailNickname
            proxyAddresses = ($user.ProxyAddresses -split ",") # supports multiple SMTP addresses if needed
            title = $user.JobTitle
            department = $user.Department
            company = $user.Company
        }
}
