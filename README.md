
# Active Directory Automation Tools

These PowerShell scripts automate common Active Directory tasks, including:
- **Creating AD Users via GUI (with robust, CSV-driven logic)**
- **Searching, viewing, and managing users**
- **Assigning and managing licenses**
- **Automated group assignment based on OU and manager role**

> **Note:** These scripts require the Active Directory module. Make sure that RSAT (Remote Server Administration Tools) is installed and that you run these scripts with elevated rights appropriate for AD administration.

## Directory Structure & Script Purposes

```
automationAD
├── userCreation.ps1           # Standalone script for creating a test AD user (for reference/testing)
├── interface.ps1              # All-in-one GUI for user creation, user search/summary, license management, and CSV import
├── adduser.ps1                # GUI for creating AD users only (streamlined version of interface.ps1)
├── addLicenses.ps1            # General scripts for assigning licenses to users (reference/utility)
├── userDeletion.ps1           # Script for deleting an AD user (includes group removal)
├── correctlyformatted.csv     # Example CSV with all required columns for user creation
├── incorrectlyformatted.csv   # Example CSV missing required columns (for validation testing)
├── secgroups.csv              # Security group mapping CSV template (OU path columns, group rows)
└── README.md                  # This file
```

> **Note:** Both GUIs (`interface.ps1` and `adduser.ps1`) use the same robust, CSV-driven logic for user creation and group assignment. All group assignment is additive for manager roles, and OU path conversion is handled automatically.

## Script Details

### `interface.ps1`
- **Purpose:** All-in-one Windows Forms GUI for:
  - Creating users interactively with a multi-page workflow
  - CSV import for dropdown population (Office, Company, Department, etc.)
  - Flexible Organizational Unit (OU) selection with two-level dropdown system
  - Searching for and viewing existing users and their properties/groups
  - Managing user licenses (search, view, assign available licenses)
- **Usage:** Run directly for a self-contained GUI experience
- **Features:**
  - **3-Page User Creation Workflow:** Basic info → Detailed attributes → OU selection & summary
  - **Smart OU Selection:** Automatically shows sub-directory options when needed
  - **Input Validation:** Username/UPN conflict detection, required field validation
  - **CSV Integration:** Populates dropdowns from CSV data with validation
  - **CSV-Driven Security Group Assignment:** Uses `secgroups.csv` for OU-to-group mapping
  - **Additive Group Assignment:** Manager/Assistant Manager roles receive additional groups
  - **OU Path Conversion:** Converts user-friendly OU paths to LDAP DN format automatically
  - **User Search & Summary:** Search for users, view properties and group memberships
  - **License Management:** View and assign licenses interactively

### `adduser.ps1`
- **Purpose:** Streamlined Windows Forms GUI for creating AD users only (no search or license management features)
- **Usage:** Run directly in PowerShell for a focused user creation experience
- **Features:**
  - Multi-page user creation workflow (basic info, detailed attributes, OU selection, summary)
  - CSV-driven dropdowns and OU/group mapping (identical to `interface.ps1`)
  - Additive group assignment logic for manager roles
  - Robust OU path conversion utility

### `userCreation.ps1`
- **Purpose:** Standalone script for creating a single AD user with hardcoded/test values
- **Usage:** Run directly in PowerShell. Useful for testing or as a template for scripting

### `addLicenses.ps1`
- **Purpose:** Reference/utility script containing general-purpose code snippets for assigning and managing licenses in Microsoft Entra (Azure AD) or MSOnline
- **Usage:** Copy/paste or adapt code blocks as needed in your own automation scripts

### `userDeletion.ps1`
- **Purpose:** Command-line script to delete an AD user, remove them from all groups, and confirm actions interactively
- **Usage:** Run directly in PowerShell. Follows prompts for username and confirmation

### CSV Files
- **`correctlyformatted.csv`:** Example of a properly formatted CSV file with all required columns for user creation (Office, Company, State, City, PostalCode, StreetAddress, Department, Title)
- **`incorrectlyformatted.csv`:** Example of a CSV missing required columns. Used for testing validation logic
- **`secgroups.csv`:** Security group mapping template. Each column header represents an OU path; each row contains group names to assign for that OU

## Prerequisites

- **Operating System:** Windows 10/11 or Windows Server
- **PowerShell:** Version 5.1 or later
- **Modules:** ActiveDirectory (provided by RSAT), Microsoft.Entra (for license management)
- **Permissions:** Administrator privileges or an account with sufficient AD and Entra rights for creating, modifying, deleting user objects, and managing licenses

## Usage Overview

### Interactive User Creation
Use `interface.ps1` or `adduser.ps1` for GUI-based user creation with a 3-page workflow:
1. **Page 1:** Basic information (Name, Username, Company) + CSV selection
2. **Page 2:** Detailed attributes (Office, Department, Title, Address, Phone)
3. **Page 3:** Organizational Unit selection + Summary review

### OU Selection
The scripts support flexible OU selection with automatic sub-directory options based on your AD structure.

### Security Group Assignment from CSV
For automated group assignment, configure the `secgroups.csv` file:
- Each column header is an OU path (e.g., `domain.local\Users`, `domain.local\Department\Users`)
- Each row contains group names to assign for that OU
- Example format:
  ```csv
  domain.local\Users,domain.local\Management\Users
  All_Staff_Group,All_Management_Group
  Basic_Access_Group,Advanced_Access_Group
  ```

### Group Assignment Logic (Manager Roles)
- Users always receive all non-manager groups for their OU
- If "Assistant Manager" is selected, they also receive groups containing `_AsstMgr_` or `_Asstmgr_` (modify as needed)
- If "Manager" is selected, they also receive groups containing `_Mgr_` (modify as needed)
- Manager role groups are assigned in addition to base OU groups, never as a replacement

### Required CSV Columns for User Data
Office, Company, State, City, PostalCode, StreetAddress, Department, Title

## Customization & Considerations

- **Domain Configuration:** Update domain references in the scripts to match your environment
- **Default Passwords:** Change hardcoded passwords before production use
- **CSV Structure:** Modify required columns in the script if your organization uses different attribute names
- **Security Group Mapping:** Customize `secgroups.csv` to match your AD structure and group policies
- **Testing:** Always test in a lab environment before production
- **UI Controls:** The GUI prevents accidental edits in selection lists and license displays
- **Validation:** Comprehensive validation for username conflicts, UPN conflicts, and required field completion

## Disclaimer

These scripts modify live Active Directory data. Use caution and perform thorough testing in a controlled environment before using in production. The scripts are provided "AS IS" without warranty of any kind.

## License

This project is provided as-is without any warranty. Use or modify it at your own risk.