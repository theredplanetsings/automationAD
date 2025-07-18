# Active Directory Automation Tools

These PowerShell scripts automate common Active Directory tasks, including:
- **Creating AD Users via GUI**

> **Note:** These scripts require the Active Directory module. Make sure that RSAT (Remote Server Administration Tools) is installed and that you run these scripts with elevated rights appropriate for AD administration.

## Directory Structure & Script Purposes

```
automationAD
├── userCreation.ps1           # Standalone script for creating a test AD user (for reference/testing)
├── interface.ps1              # All-in-one GUI for user creation, user search/summary, license management, and CSV import
│                                 # Supports additive group assignment for manager roles (see below)
├── addLicenses.ps1            # General scripts for assigning licenses to users (reference/utility)
├── userDeletion.ps1           # Script for deleting an AD user (includes group removal)
├── correctlyformatted.csv     # Example CSV with all required columns for user creation
├── incorrectlyformatted.csv   # Example CSV missing required columns (for validation testing)
└── README.md                  # This file
```

> **Note:** The GUI (`interface.ps1`) and group assignment logic are fully CSV-driven. When creating a user, all non-manager groups for the selected OU are assigned by default. If a manager role is selected, the user also receives the appropriate manager groups (additive, not replacement) as defined in the OU's group mapping. See the "Group Assignment Logic (Manager Roles)" section below for details.

### Script Details & Differences

#### `userCreation.ps1`
- **Purpose:** Standalone script for creating a single AD user with hardcoded/test values. Useful for testing or as a template for scripting.
- **Usage:** Run directly in PowerShell. Not interactive or GUI-based.

#### `interface.ps1`
- **Purpose:** All-in-one, non-modular Windows Forms GUI for:
  - Creating users interactively with multi-page workflow
  - CSV import for dropdown population (Office, Company, Department, etc.)
  - Flexible Organizational Unit (OU) selection with two-level dropdown system
  - Searching for and viewing existing users and their properties/groups
  - Managing user licenses (search, view, assign available licenses)
- **Usage:** Run directly for a self-contained GUI experience. Good for quick testing or single-file deployment.
- **Features:**
  - **3-Page User Creation Workflow:** Basic info → Detailed attributes → OU selection & summary
  - **Smart OU Selection:** Automatically shows sub-directory options when needed (External/Internal)
  - **Input Validation:** Username/UPN conflict detection, required field validation
  - **CSV Integration:** Populates dropdowns from CSV data with validation
- **Note:** Not designed for code reuse or modularity.

#### `addLicenses.ps1`
- **Purpose:** Reference/utility script containing general-purpose code snippets for assigning and managing licenses in Microsoft Entra (Azure AD) or MSOnline. Not a standalone program, but a collection of reusable script blocks for integration into larger automation workflows.
- **Usage:** Copy/paste or adapt code blocks as needed in your own automation scripts or GUIs.

#### `userDeletion.ps1`
- **Purpose:** Command-line script to delete an AD user, remove them from all groups, and confirm actions interactively.
- **Usage:** Run directly in PowerShell. Follows prompts for username and confirmation.

#### CSV Files
- **`correctlyformatted.csv`:** Example of a properly formatted CSV file with all required columns for user creation (Office, Company, State, City, PostalCode, StreetAddress, Department, Title). Use as a template for bulk user creation.
- **`incorrectlyformatted.csv`:** Example of a CSV missing required columns. Used for testing validation logic in the GUIs/scripts.

## Prerequisites

- **Operating System:** Windows 10/11 or Windows Server.
- **PowerShell:** Version 5.1 or later.
- **Modules:** ActiveDirectory (provided by RSAT), Microsoft.Entra (for license management)
- **Permissions:** Administrator privileges or an account with sufficient AD and Entra rights for creating, modifying, deleting user objects, and managing licenses.

## Usage Overview

### Interactive/GUI User Creation, Search, and License Management
- Use `interface.ps1` for a self-contained GUI that supports:
  - **Adding new users** with a 3-page workflow:
    - Page 1: Basic information (Name, Username, Company) + CSV selection
    - Page 2: Detailed attributes (Office, Department, Title, Address, Phone)
    - Page 3: Organizational Unit selection + Summary review
  - **OU Selection:** Choose from multiple organizational units with automatic sub-directory selection:
    - Users (direct placement)
    - Service Accounts (direct placement)  
    - PDC-CONSTRUCTION\USERS (requires External/Internal selection)
    - PDC-HQ\USERS (requires External/Internal selection)
    - PDC-SERVICES\USERS (requires External/Internal selection)
  - **Searching** for existing users and viewing their properties/groups
  - **Viewing and assigning** licenses to users

### Bulk User Creation and Security Group Assignment from CSV
- Prepare a CSV matching the format of `correctlyformatted.csv` for user attributes.

For security group assignment automation, use the provided `sgroups_template.csv` as your mapping file:
  - Each column header is an OU path (e.g., `paradigmcos.local\Users`, `paradigmcos.local\PDC-MANAGEMENT\360hstreet\Users`, etc.).
  - For each OU column, enter the security groups (comma-separated) that should be assigned to users created in that OU.
  - Example row:
    ```csv
    paradigmcos.local\\Users,paradigmcos.local\\PDC-MANAGEMENT\\360hstreet\\Users
    All_Paradigm_Staff@paradigmcos.com,All_Management_Staff@paradigmcos.com
    ```

**Group Assignment Logic (Manager Roles):**
- Users always receive all non-manager groups for their OU.
- If "Assistant Manager" is selected, they also receive all groups in their OU containing `_AsstMgr_` or `_Asstmgr_` (additive).
- If "Manager" is selected, they also receive all groups in their OU containing `_Mgr_` (additive).
- This ensures that manager role groups are assigned in addition to the base OU groups, never as a replacement.

- The GUI script imports this mapping and automatically assigns the specified groups when creating a user in the matching OU, applying the above logic for manager roles.
- **Required User CSV Columns:** Office, Company, State, City, PostalCode, StreetAddress, Department, Title
- CSV data automatically populates dropdown menus for consistent data entry.

### License Assignment
- Use the license management page in `interface.ps1` for interactive assignment.
- Use code snippets from `addLicenses.ps1` in your own scripts or automation tools as needed.

### User Deletion
- Run `userDeletion.ps1` and follow the prompts.

## Customization & Considerations

- **OU Paths:** The interface includes built-in OU selection and now supports external mapping of OU paths to security groups via CSV. Use `sgroups_template.csv` as your starting point for OU-to-group mapping.
- **Security Group Mapping CSV:** To automate group assignment, fill in the security groups for each OU column in `sgroups_template.csv`. Each cell should contain a comma-separated list of group names to assign for that OU.
- **Domain Configuration:** Update domain references (`paradigmcos.com`, `paradigmcos.local`) to match your environment.
- **Default Passwords:** Change hardcoded passwords (`Password123@`) before production use.
- **CSV Structure:** Modify required columns in the script if your organization uses different attribute names.
- **Testing:** Always test in a lab environment before production.
- **UI Controls:** The GUI disables typing in selection lists and license displays to prevent accidental edits; only selection is allowed where appropriate.
- **Validation:** The interface includes comprehensive validation for username conflicts, UPN conflicts, and required field completion.

## Disclaimer

These scripts modify live Active Directory data. Use caution and perform thorough testing in a controlled environment before using in production. The scripts are provided "AS IS" without warranty of any kind.

## License

This project is provided as-is without any warranty. Use or modify it at your own risk.