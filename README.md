# Active Directory Automation Tools

These PowerShell scripts automate common Active Directory tasks, including:

- **Creating AD Users interactively:** Prompt for user details and create a new AD user with additional properties and optional security group assignments.
- **Creating AD Users from CSV:** Bulk-create AD users based on a CSV file containing the necessary user properties.
- **Deleting AD Users:** Locate an AD user by username, remove them from all group memberships, and delete the user from AD.

> **Note:** These scripts require the Active Directory module. Make sure that RSAT (Remote Server Administration Tools) is installed and that you run these scripts with elevated rights appropriate for AD administration.

## Directory Structure

```
automationAD
├── userCreationFromInput.ps1   # Interactive user creation script
├── userCreationFromCSV.ps1     # Bulk user creation script reading from a CSV file
├── userDeletion.ps1            # Script for deleting an AD user (includes group removal)
└── README.md                   # This file
```

## Prerequisites

- **Operating System:** Windows 10/11 or Windows Server.
- **PowerShell:** Version 5.1 or later.
- **Modules:** ActiveDirectory (provided by RSAT)
- **Permissions:** Administrator privileges or an account with sufficient AD rights for creating, modifying, and deleting user objects.

## Scripts Overview

### 1. Interactive User Creation (`userCreationFromInput.ps1`)

This script prompts for the following AD user details:

- First & Last Name  
- Office  
- Telephone Number  
- Email Address  
- Mail Nickname  
- Proxy Addresses (comma separated, e.g., `smtp:addr1,smtp:addr2`)  
- Job Title  
- Department  
- Company  
- Security Groups (comma separated, if any)

It then:
- Generates a username (first initial + last name in lowercase)
- Creates a new AD user with a default password
- Assigns the new user to provided security groups (if any)
- Sets additional properties such as office, phone number, email, and more

**Usage:**

1. Open PowerShell as an administrator.
2. Navigate to the project folder:
   ```powershell
   cd "C:\Users\crutherford\Downloads\it-tools\automationAD"
   ```
3. Run the script:
   ```powershell
   .\userCreationFromInput.ps1
   ```
4. Follow the prompts to input the required details.

### 2. Bulk User Creation from CSV (`userCreationFromCSV.ps1`)

This script reads user details from a CSV file. The CSV should have headers such as:

- FirstName
- LastName
- Office
- Telephone
- Email
- MailNickname
- ProxyAddresses (comma separated if more than one address)
- JobTitle
- Department
- Company
- Groups (comma separated security groups)

It then:
- Generates a username from the first letter of the first name and the last name (all lowercase)
- Creates a new AD user for each record in the CSV
- Assigns the user to any specified security groups
- Sets additional user properties

**Usage:**

1. Prepare your CSV file with the proper headers.
2. Update the CSV file path in the script (`$users = Import-Csv -Path "file_path_here"`)
3. Open PowerShell as an administrator.
4. Navigate to the project folder:
   ```powershell
   cd "file_path_here"
   ```
5. Run the script:
   ```powershell
   .\userCreationFromCSV.ps1
   ```

### 3. User Deletion (`userDeletion.ps1`)

This script allows you to delete an AD user by:
- Prompting for the username
- Validating that the user exists
- Confirming deletion (with an interactive prompt)
- Removing the user from all group memberships
- Deleting the user from Active Directory

**Usage:**

1. Open PowerShell as an administrator.
2. Navigate to the project folder:
   ```powershell
   cd "file_path_here"
   ```
3. Run the script:
   ```powershell
   .\userDeletion.ps1
   ```
4. Enter the username when prompted and follow the on-screen instructions to confirm deletion.

## Customization & Considerations

- **OU Paths:**  
  The scripts use the OU path `"OU=Users,DC=domain,DC=com"`. Update this parameter to match your AD configuration.

- **Default Passwords:**  
  Passwords are set using a hardcoded default. Modify these values to suit your organization's policies before running in production.

- **Group Membership Removal:**  
  In `userDeletion.ps1`, the script automatically removes the user from all group memberships. If needed, you can adjust this behavior to prompt for confirmation.

- **Testing:**  
  Always test these scripts in a lab environment before deployment in a production environment to ensure they meet your requirements and work with your AD configurations.

## Disclaimer

These scripts modify live Active Directory data. Use caution and perform thorough testing in a controlled environment before using in production. The scripts are provided "AS IS" without warranty of any kind.

## License

This project is provided as-is without any warranty. Use or modify it at your own risk.
