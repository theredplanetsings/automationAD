# Active Directory Automation Tools

These PowerShell scripts automate common Active Directory tasks, including:
- **Creating AD Users via GUI (with robust, CSV-driven logic)**
- **Automated group assignment based on OU and manager role**
- **User deletion with group cleanup**

> **Note:** These scripts require the Active Directory module (RSAT) and appropriate AD permissions. The GUI (`adduser.ps1`) uses CSV-driven logic for user creation with additive group assignment for manager roles.

## Directory Structure & Script Purposes

```
automationAD
├── userCreation.ps1           # Standalone script for creating a test AD user (for reference/testing)
├── adduser.ps1                # GUI for creating AD users with CSV-driven logic and group assignment
├── automationad.ps1           # Same as adduser.ps1 but with detailed comments for extending manager roles
├── addLicenses.ps1            # General scripts for assigning licenses to users (reference/utility)
├── userDeletion.ps1           # Script for deleting an AD user (includes group removal)
├── correctlyformatted.csv     # Example CSV with all required columns for user creation
├── incorrectlyformatted.csv   # Example CSV missing required columns (for validation testing)
├── secgroups.csv              # Security group mapping CSV template (OU path columns, group rows)
└── README.md                  # self explanatory
```

> **Note:** `adduser.ps1` and `automationad.ps1` are functionally identical. Use `automationad.ps1` if you need to understand or modify the manager role group assignment logic, as it contains detailed comments explaining how to add additional manager/assistant manager role patterns.

## Quick Start Guide for adduser.ps1 / automationad.ps1

### Prerequisites
- **ActiveDirectory PowerShell module** (installed via RSAT)
- **Administrative privileges** for AD user creation
- **Two CSV files** prepared according to specifications below

> **Script Choice:** Use either `adduser.ps1` or `automationad.ps1` - they are functionally identical. Choose `automationad.ps1` if you plan to modify manager role logic, as it contains detailed comments for extending the group assignment patterns.

### Basic Usage Steps
1. **Run the script:** `.\adduser.ps1` or `.\automationad.ps1` in PowerShell
2. **Page 1:** Enter basic user information and select your User Properties CSV
3. **Page 2:** Fill in detailed user attributes (populated from CSV)
4. **Page 3:** Select OU/Security Group CSV, choose organizational unit, and review summary
5. **Create User:** Click "Create User" after reviewing all information

## CSV File Requirements

### User Properties CSV (Required for Page 1-2)

The User Properties CSV populates dropdown menus for user attributes. This file **must** contain the following columns in any order:

#### Required Columns:
- **Office** - Physical office locations (e.g., "Corporate", "Branch A")
- **Company** - Company divisions or subsidiaries (e.g., "Paradigm Development", "Paradigm Services")
- **State** - State abbreviations or full names (e.g., "VA", "Maryland")
- **City** - City names (e.g., "Arlington", "Washington D.C.")
- **PostalCode** - ZIP/postal codes (e.g., "12345", "23456")
- **StreetAddress** - Street addresses (e.g., "123 Main St", "456 Secondary St")
- **Department** - Department names (e.g., "IT", "HR", "Finance")
- **Title** - Job titles (e.g., "IT Specialist", "Manager", "Analyst")

#### CSV Format Example:
```csv
Office,Company,State,City,PostalCode,StreetAddress,Department,Title
Corporate,Paradigm Development,VA,Arlington,12345,123 Main St,IT,IT Specialist
Branch A,Paradigm Services,MD,Washington D.C.,23456,456 Secondary St,HR,Manager
Branch B,Paradigm Management,DC,Alexandria,34567,789 Third St,Finance,Analyst
```

#### Important Notes:
- **Column names must match exactly** (case-sensitive)
- **All columns must be present** - missing columns will cause validation errors
- **Values can repeat** across rows to provide multiple options
- **Unique values** from each column populate the respective dropdown menus
- **Empty cells are allowed** but will not appear in dropdowns

### OU/Security Groups CSV (Required for Page 3)

The OU/Security Groups CSV maps organizational units to security groups for automatic assignment.

#### Column Header Format/Examples:
Each column header **must** represent an OU path starting with your domain:
```
yourdomain.local\OUName
yourdomain.local\ParentOU\ChildOU
yourdomain.local\ParentOU\ChildOU\SubOU
```

#### Security Group Assignment Rows:
Each row represents a "layer" of security groups:
- **Row 1:** Base groups assigned to ALL users in the OU
- **Row 2:** Additional groups for specific OUs
- **Row 3:** Assistant Manager groups (containing `_AsstMgr_` or `_Asstmgr_`)
- **Row 4:** Manager groups (containing `_Mgr_`)
- **Additional rows:** Custom group assignments as needed

#### CSV Format Example:
```csv
yourdomain.local\OUName,yourdomain.local\ParentOU\ChildOU,yourdomain.local\ParentOU\ChildOU\SubOU
group1,group1,group1
,group_2,group_2
,,Property_AsstMgr_Group
,,Property_Mgr_Group
```

#### Group Assignment Logic:
1. **Base Groups:** All users receive groups from rows that don't contain manager-specific patterns
2. **Assistant Manager:** Additionally receives groups containing `_AsstMgr_` or `_Asstmgr_`
3. **Manager:** Additionally receives groups containing `_Mgr_`
4. **Additive Assignment:** Manager roles add groups, never replace base groups

#### Important Notes:
- **OU paths must start with your domain** (e.g., `yourdomain.local\`)
- **Empty cells are allowed** and will be skipped
- **Group names should match existing AD security groups**
- **Manager-specific groups are identified by naming patterns** in the group names
- **The script converts OU paths to LDAP format** automatically

## Script Details

### `adduser.ps1` / `automationad.ps1` - Detailed Usage Guide

#### Overview
Interactive Windows Forms GUI for creating Active Directory users with comprehensive validation and CSV-driven automation. Both scripts are functionally identical.

> **Development Guide:** If you need to modify or extend the manager role group assignment logic, use `automationad.ps1` as it contains detailed comments explaining the pattern matching system and how to add new role types (e.g., Property Manager, Assistant Property Manager roles).

#### Features
- **3-Page Workflow:** Guided user creation process
- **Smart Username Generation:** Auto-generates username from first initial + last name
- **Conflict Detection:** Validates username and UPN uniqueness
- **CSV Integration:** Populates form fields from CSV data
- **Flexible OU Selection:** Supports nested OU structures with sub-directory options
- **Automated Group Assignment:** Maps users to security groups based on OU and role
- **Manager Role Support:** Additive group assignment for management positions
- **Input Validation:** Comprehensive validation at each step
- **Summary Review:** Complete user information review before creation

#### Page-by-Page Workflow

##### Page 1: Basic Information
**Required Fields:**
- **First Name** - User's first name
- **Last Name** - User's surname  
- **Username** - Auto-generated (editable) as first initial + last name
- **Company** - Selected from CSV dropdown

**Required Action:**
- **Select CSV** - Choose your User Properties CSV file

**Validation:**
- All fields must be completed
- CSV must contain all required columns
- Username auto-updates as you type names

##### Page 2: Detailed Attributes
**All fields required and populated from CSV:**
- **Office** - Physical office location
- **Department** - User's department
- **Job Title** - User's position/role
- **Street Address** - Physical address
- **City** - City location
- **State** - State/province
- **Postal Code** - ZIP/postal code
- **Telephone Number** - Phone number (manual entry)

**Validation:**
- All dropdown selections must be made
- Telephone number must be entered manually

##### Page 3: OU Selection and Review
**Required Actions:**
- **Select OU CSV** - Choose your OU/Security Groups CSV file
- **Select Organizational Unit** - Choose from available OUs
- **Select Sub-Directory** - If applicable (shown automatically)
- **Manager Role** - Choose manager level (if management OU selected)

**Review Summary:**
- Complete user information display
- OU path in both human-readable and LDAP format
- Security groups that will be assigned
- Final validation before user creation

#### Automatic Features

##### Username Generation
- **Format:** First initial + full last name (lowercase)
- **Example:** John Smith → `jsmith`
- **Editable:** You can modify the generated username

##### Email Address Creation
- **Format:** `username@yourdomain.com`
- **Automatic:** Generated from username and domain in script

##### OU Path Conversion
- **Input:** Human-readable OU path from CSV
- **Output:** LDAP Distinguished Name format
- **Example:** `domainname.com\PDC-HQ\Users` → `OU=Users,OU=PDC-HQ,DC=domainname,DC=com`

##### Security Group Assignment
1. **Base Groups:** All non-manager groups for the selected OU
2. **Manager Groups:** Additional groups based on manager role selection
3. **Additive Logic:** Manager roles receive base groups PLUS manager-specific groups

#### Error Handling and Validation

##### Username Conflicts
- **Check:** Validates username doesn't already exist in AD
- **Action:** Displays error message and prevents creation
- **Resolution:** User must choose a different username

##### UPN Conflicts  
- **Check:** Validates User Principal Name uniqueness
- **Action:** Displays error message and prevents creation
- **Resolution:** User must choose a different username

##### CSV Validation
- **Check:** Verifies all required columns are present
- **Action:** Displays specific missing columns
- **Resolution:** User must select a correctly formatted CSV

##### Required Field Validation
- **Check:** Ensures all required fields are completed at each page
- **Action:** Displays specific missing fields
- **Resolution:** User must complete all required information

##### OU Selection Validation
- **Check:** Validates proper OU and sub-OU selection
- **Action:** Prevents user creation with incomplete OU selection
- **Resolution:** User must select valid OU path

#### User Account Creation Details

##### Default Settings
- **Password:** `Password123@` (change before production use)
- **Change Password at Logon:** True
- **Account Enabled:** True
- **Domain:** Set in script (update for your environment)

##### Attributes Set
- **Basic Identity:** Name, username, UPN, display name
- **Contact Information:** Email, phone, office
- **Location Information:** Street address, city, state, postal code
- **Organizational:** Department, title, company
- **Exchange Attributes:** Mail nickname, proxy addresses

#### Troubleshooting Common Issues

##### "CSV missing columns" Error
- **Cause:** User Properties CSV doesn't have all required columns
- **Solution:** Ensure CSV has: Office, Company, State, City, PostalCode, StreetAddress, Department, Title

##### "Please select a valid Organizational Unit" Error
- **Cause:** OU selection incomplete or invalid
- **Solution:** Select both main OU and sub-directory if applicable

##### "Username already exists" Error
- **Cause:** Chosen username conflicts with existing AD user
- **Solution:** Modify the username field to something unique

##### "Could not convert OU path to LDAP format" Error
- **Cause:** Selected OU path doesn't match OU CSV format
- **Solution:** Verify OU CSV column headers start with domain name

##### Dropdown Menus Empty
- **Cause:** CSV file not properly loaded or formatted
- **Solution:** Re-select CSV file and ensure proper formatting

### `userCreation.ps1`
- **Purpose:** Standalone script for creating a single AD user with hardcoded/test values
- **Usage:** Run directly in PowerShell. Useful for testing or as a template for scripting

### `addLicenses.ps1`
- **Purpose:** Reference/utility script containing general-purpose code snippets for assigning and managing licenses in Microsoft Entra (Azure AD) or MSOnline
- **Usage:** Copy/paste or adapt code blocks as needed in your own automation scripts

### `userDeletion.ps1`
- **Purpose:** Command-line script to delete an AD user, remove them from all groups, and confirm actions interactively
- **Usage:** Run directly in PowerShell. Follows prompts for username and confirmation

## Configuration and Customization

### Domain Configuration
Update the following variables in `adduser.ps1` or `automationad.ps1` for your environment:
```powershell
$domainname = "yourdomain.com"          # Line ~22: Email domain
# Search for "paradigmcos.local" and replace with your AD domain
```

### Default Password
**CRITICAL:** Change the default password before production use:
```powershell
$defaultpassword = ConvertTo-SecureString "YourSecurePassword123!" -AsPlainText -Force
```

### Manager Role Group Patterns
Customize the patterns used to identify manager-specific groups. For detailed comments on extending these patterns, see `automationad.ps1`:

```powershell
# Assistant Manager groups (modify line ~124)
$asstMgrGroups = $allGroups | Where-Object { $_ -match '_(AsstMgr|Asstmgr|AsstPropertyMgr)_' }

# Manager groups (modify line ~127) 
$mgrGroups = $allGroups | Where-Object { $_ -match '_(Mgr|PropertyMgr)_' -and $_ -notmatch '_(AsstMgr|Asstmgr)_' }
```

> **Developer Note:** `automationad.ps1` contains extensive comments explaining how to add additional role patterns (e.g., `_PropertyMgr_`, `_AsstPropertyMgr_`) and gracefully handle cases where specific role groups may not exist in your environment.

### Custom Validation Rules
Modify required columns for User Properties CSV (line ~158):
```powershell
$requiredColumns = @("Office","Company","State","City","PostalCode","StreetAddress","Department","Title")
```

> **Extending Manager Roles:** For detailed instructions on adding new manager role patterns (e.g., Regional Manager, Department Manager), see the commented code sections in `automationad.ps1` around lines 119-174. The script uses PowerShell's `-match` and `-notmatch` operators with regex patterns to identify and assign role-specific security groups.

## Best Practices and Security Considerations

### Before Production Use
1. **Test in Lab Environment:** Always test with non-production AD first
2. **Change Default Passwords:** Never use default passwords in production
3. **Validate CSV Data:** Ensure CSV files contain accurate, up-to-date information
4. **Review Group Assignments:** Verify security group assignments match your policies
5. **Backup Strategy:** Ensure AD backups are current before bulk user creation

### Security Recommendations
- **Least Privilege:** Run script with minimum required AD permissions
- **Secure Password Policy:** Implement strong default passwords and force change at logon
- **Group Policy:** Ensure proper GPOs apply to created users
- **Audit Trail:** Monitor AD logs for user creation activities
- **Regular Review:** Periodically audit created accounts and group memberships

### Maintenance Tips
- **Keep CSVs Updated:** Regularly update CSV files as organizational structure changes
- **Version Control:** Track changes to CSV files and script modifications
- **Documentation:** Document any customizations made to the script
- **Regular Testing:** Periodically test user creation process with test accounts

## Example Workflow

### Preparing CSV Files

1. **Create User Properties CSV:**
   ```csv
   Office,Company,State,City,PostalCode,StreetAddress,Department,Title
   Headquarters,Acme Corp,CA,San Francisco,94102,123 Main St,IT,Systems Admin
   Branch Office,Acme Corp,NY,New York,10001,456 Broadway,Sales,Sales Rep
   ```

2. **Create OU/Security Groups CSV:**
   ```csv
   acme.local\Users,acme.local\IT\Users,acme.local\Sales\Users
   All_Employees,All_Employees,All_Employees
   ,IT_Department,Sales_Department
   ,IT_AsstMgr_Access,Sales_AsstMgr_Access
   ,IT_Mgr_Access,Sales_Mgr_Access
   ```

### Creating a User

1. **Launch Script:** `.\adduser.ps1` or `.\automationad.ps1`
2. **Page 1:**
   - Enter: John, Doe, jdoe, Acme Corp
   - Select User Properties CSV
   - Click Next
3. **Page 2:**
   - Select: Headquarters, IT, Systems Admin, etc.
   - Enter phone number
   - Click Next  
4. **Page 3:**
   - Select OU/Security Groups CSV
   - Choose: acme.local\IT\Users
   - Select manager role if applicable
   - Review summary
   - Click Create User

### Result
- User created in `OU=Users,OU=IT,DC=acme,DC=local`
- Assigned to: All_Employees, IT_Department groups
- Additional manager groups if role selected
- Email: jdoe@acme.com
- Must change password at next logon

## Disclaimer

These scripts modify live Active Directory data. Use caution and perform thorough testing in a controlled environment before using in production. The scripts are provided "AS IS" without warranty of any kind.

## Licence

Creative Commons Zero v1.0 Universal (CC0) - Public domain dedication for maximum freedom