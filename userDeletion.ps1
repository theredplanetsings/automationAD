Import-Module ActiveDirectory
# user to be deleted
$username = Read-Host "Enter the username to delete"

# gets specified user within AD
$user = Get-ADUser -Identity $username -ErrorAction SilentlyContinue

#error handling if user isn't located
if (!$user) {
    Write-Host "User with username '$username' not found."
    exit
}
#confirms if user found
Write-Host "Found user: $($user.Name)"
$confirm = Read-Host "Are you sure you want to delete this user? (Y/N)"
if ($confirm -ne "Y") {
    Write-Host "Operation cancelled."
    exit
}

# removes user from all group memberships
$groups = Get-ADPrincipalGroupMembership $user | Select-Object -ExpandProperty Name
if ($groups) {
    foreach ($grp in $groups) {
        Remove-ADGroupMember -Identity $grp -Members $user -Confirm:$false
        Write-Host "Removed $username from group $grp"
    }
}

# removes the user from Active Directory
Remove-ADUser -Identity $user -Confirm
#debug statement confirming deletion
Write-Host "User '$username' has been deleted from Active Directory."
