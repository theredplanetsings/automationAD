# Active Directory Automation Tools

These PowerShell scripts automate common Active Directory tasks, including:

- **Creating AD Users interactively or via GUI**
- **Bulk-creating AD Users from a CSV file**
- **Deleting AD Users**
- **Assigning and managing licenses for users**

> **Note:** These scripts require the Active Directory module. Make sure that RSAT (Remote Server Administration Tools) is installed and that you run these scripts with elevated rights appropriate for AD administration.

## Directory Structure & Script Purposes

```
automationAD
├── userCreation.ps1           # Standalone script for creating a test AD user (for reference/testing)
├── interface.ps1              # All-in-one GUI for user creation, user search/summary, license management, and CSV import (non-modular)
├── interfacemodular.ps1       # Modular GUI, imports ad-user.ps1 for user creation logic
│   └── ad-user.ps1            # Contains the Create-ADUserFromForm function, only used by interfacemodular.ps1
├── addLicenses.ps1            # General scripts for assigning licenses to users (reference/utility)
├── userDeletion.ps1           # Script for deleting an AD user (includes group removal)
├── correctlyformatted.csv     # Example CSV with all required columns for user creation
├── incorrectlyformatted.csv   # Example CSV missing required columns (for validation testing)
└── README.md                  # This file
```

### Script Details & Differences

#### `userCreation.ps1`
- **Purpose:** Standalone script for creating a single AD user with hardcoded/test values. Useful for testing or as a template for scripting.
- **Usage:** Run directly in PowerShell. Not interactive or GUI-based.

#### `interface.ps1`
- **Purpose:** All-in-one, non-modular Windows Forms GUI for:
  - Creating users interactively or via CSV import
  - Searching for and viewing existing users and their properties/groups
  - Managing user licenses (search, view, assign available licenses)
- **Usage:** Run directly for a self-contained GUI experience. Good for quick testing or single-file deployment.
- **Note:** Not designed for code reuse or modularity.

#### `interfacemodular.ps1` & `ad-user.ps1`
- **Purpose:**
  - `interfacemodular.ps1` provides the GUI and user interaction logic.
  - `ad-user.ps1` contains the actual user creation function (`Create-ADUserFromForm`).
- **How they work together:** `interfacemodular.ps1` imports `ad-user.ps1` and calls its function to perform user creation. This modular approach separates the interface from the business logic, making it easier to maintain and extend.
- **Usage:** Run `interfacemodular.ps1` to launch the modular GUI. Do not run `ad-user.ps1` directly.
- **Note:** `ad-user.ps1` is not a standalone script; it is a module used exclusively by `interfacemodular.ps1`.

#### `addLicenses.ps1`
- **Purpose:** Reference/utility script containing general-purpose code snippets for assigning and managing licenses in Microsoft Entra (Azure AD) or MSOnline. Not a standalone program, but a collection of reusable script blocks for integration into larger automation workflows (similar to how `ad-user.ps1` is used for user creation logic).
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
  - Adding new users
  - Searching for existing users and viewing their properties/groups
  - Viewing and assigning licenses to users
- Use `interfacemodular.ps1` (with `ad-user.ps1`) for a modular, maintainable GUI focused on user creation.

### Bulk User Creation from CSV
- Prepare a CSV matching the format of `correctlyformatted.csv`.
- Use the GUI scripts to import and validate the CSV, then create users.

### License Assignment
- Use the license management page in `interface.ps1` for interactive assignment.
- Use code snippets from `addLicenses.ps1` in your own scripts or automation tools as needed.

### User Deletion
- Run `userDeletion.ps1` and follow the prompts.

## Customization & Considerations

- **OU Paths:** Update the OU path in scripts to match your AD configuration.
- **Default Passwords:** Change hardcoded passwords before production use.
- **Testing:** Always test in a lab environment before production.
- **UI Controls:** The GUI disables typing in selection lists and license displays to prevent accidental edits; only selection is allowed where appropriate.

## Disclaimer

These scripts modify live Active Directory data. Use caution and perform thorough testing in a controlled environment before using in production. The scripts are provided "AS IS" without warranty of any kind.

## License

This project is provided as-is without any warranty. Use or modify it at your own risk.