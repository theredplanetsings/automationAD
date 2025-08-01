<#============================================
  Automation AD - Main Program
============================================#>

# =========================
# Imports and Initial Setup
# =========================
Get-Module -ListAvailable ActiveDirectory
Import-Module ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =========================
# User Creation Functions
# =========================
function Create-ADUserFromForm {
    # gather inputs
    $firstName = $txtFirstName.Text.Trim()
    $lastName = $txtLastName.Text.Trim()
    if (-not $firstName -or -not $lastName) {
        [System.Windows.Forms.MessageBox]::Show("Please enter both first and last name.", "Input Error")
        return
    }
    $fullName = "$firstName $lastName"
    $username = $txtUsername.Text.Trim()
    $domainname = "paradigmcos.com"
    $userPrincipalName = "$username@$domainname"
    $defaultpassword = ConvertTo-SecureString "Password123@" -AsPlainText -Force
    
    # Checks if username already exists
    try {
        $existingUser = Get-ADUser -Filter "SamAccountName -eq '$username'" -ErrorAction SilentlyContinue
        if ($existingUser) {
            [System.Windows.Forms.MessageBox]::Show("Username '$username' already exists. Please choose a different username.", "Username Conflict", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    } catch {
        # Continues if user doesn't exist
    }
    # Checks if UPN already exists
    try {
        $existingUPN = Get-ADUser -Filter "UserPrincipalName -eq '$userPrincipalName'" -ErrorAction SilentlyContinue
        if ($existingUPN) {
            [System.Windows.Forms.MessageBox]::Show("User Principal Name '$userPrincipalName' already exists. Please choose a different username.", "UPN Conflict", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    } catch {
        # Continues if UPN doesn't exist
    }
    # retrieves values from dropdowns and textboxes - convert to strings
    $physicalDeliveryOfficeName = if ($cmbOffice.SelectedItem) { $cmbOffice.SelectedItem.ToString() } else { "" }
    $company = if ($cmbCompany.SelectedItem) { $cmbCompany.SelectedItem.ToString() } else { "" }
    $st = if ($cmbState.SelectedItem) { $cmbState.SelectedItem.ToString() } else { "" }
    $l = if ($cmbCity.SelectedItem) { $cmbCity.SelectedItem.ToString() } else { "" }
    $postalCode = if ($cmbPostalCode.SelectedItem) { $cmbPostalCode.SelectedItem.ToString() } else { "" }
    $streetAddress = if ($cmbStreetAddress.SelectedItem) { $cmbStreetAddress.SelectedItem.ToString() } else { "" }
    $department = if ($cmbDepartment.SelectedItem) { $cmbDepartment.SelectedItem.ToString() } else { "" }
    $title = if ($cmbTitle.SelectedItem) { $cmbTitle.SelectedItem.ToString() } else { "" }
    $telephoneNum = $txtTelephone.Text.Trim()
    $mail = "$username@$domainname"
    $mailNickname = $username
    $proxyAddresses = "smtp:$mail"
    # console output, not fully necessary but useful for debugging
    Write-Host "Debug Values:"
    Write-Host "Office: '$physicalDeliveryOfficeName'"
    Write-Host "Company: '$company'"
    Write-Host "Department: '$department'"
    Write-Host "Title: '$title'"
    Write-Host "Street: '$streetAddress'"
    Write-Host "City: '$l'"
    Write-Host "State: '$st'"
    Write-Host "Postal: '$postalCode'"
    Write-Host "Phone: '$telephoneNum'"
    
    # Validates our OU selection
    $ouPath = Get-FullOUPath
    if (-not $ouPath) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid Organizational Unit. If you selected a main OU with sub-directories, you must also select a sub-directory.", "OU Selection Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    # Converts to LDAP path for AD
    $ldapOUPath = Convert-OUPathToLDAP $ouPath
    if (-not $ldapOUPath) {
        [System.Windows.Forms.MessageBox]::Show("Could not convert OU path to LDAP format. Please check your OU selection.", "OU Path Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    try {
        # creating the new user with starter properties
        New-ADUser  `
            -Name $fullName `
            -AccountPassword $defaultpassword `
            -GivenName $firstName `
            -Surname $lastName `
            -DisplayName $fullName `
            -SamAccountName $username `
            -UserPrincipalName $userPrincipalName `
            -ChangePasswordAtLogon $true `
            -Path $ldapOUPath `
            -Enabled $true
        # Only set attributes that have values
        $attributesToSet = @{}
        if ($streetAddress) { $attributesToSet.streetAddress = $streetAddress }
        if ($postalCode) { $attributesToSet.postalCode = $postalCode }
        if ($mailNickname) { $attributesToSet.mailNickname = $mailNickname }
        if ($proxyAddresses) { $attributesToSet.proxyAddresses = $proxyAddresses }
        if ($title) { $attributesToSet.title = $title }
        if ($department) { $attributesToSet.department = $department }
        if ($company) { $attributesToSet.company = $company }
        if ($l) { $attributesToSet.l = $l }
        if ($st) { $attributesToSet.st = $st }
        # sets additional attributes
        Set-ADUser `
            -Identity $username `
            -Office $physicalDeliveryOfficeName `
            -OfficePhone $telephoneNum `
            -EmailAddress $mail `
            -Replace $attributesToSet
        # Assigns security groups based on OU selection and manager role (additive)
        $groupsToAdd = @()
        if ($script:ouGroups.ContainsKey($ouPath)) {
            $allGroups = $script:ouGroups[$ouPath]
            $mgrRole = if ($cmbMgrRole.Visible) { $cmbMgrRole.SelectedItem } else { "Not a manager" }
            # Assign all non-manager groups
            $baseGroups = $allGroups | Where-Object { ($_ -notmatch '_(AsstMgr|Asstmgr)_' ) -and ($_ -notmatch '_Mgr_') }
            $groupsToAdd = @($baseGroups)
            if ($mgrRole -eq "Assistant Manager") {
                $asstMgrGroups = $allGroups | Where-Object { $_ -match '_(AsstMgr|Asstmgr)_' }
                $groupsToAdd += $asstMgrGroups
            } elseif ($mgrRole -eq "Manager") {
                $mgrGroups = $allGroups | Where-Object { $_ -match '_Mgr_' }
                $groupsToAdd += $mgrGroups
            }
            # Removes duplicates
            $groupsToAdd = $groupsToAdd | Sort-Object -Unique
        }
        # Adds user to each group
        foreach ($group in $groupsToAdd) {
            try {
                Add-ADGroupMember -Identity $group -Members $username -ErrorAction Stop
            } catch {
                Write-Host "Error adding $username to ${group}: $($_.Exception.Message)"
            }
        }
        [System.Windows.Forms.MessageBox]::Show("User created successfully!","Success")
        $ShowPage1.Invoke()
        $form.Show()
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Error")
    }
}

# =========================
# Form Initialisation
# =========================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Automation AD - Dashboard"
$form.Size = New-Object System.Drawing.Size(800,600)
$form.MinimumSize = New-Object System.Drawing.Size(800,600)
$form.MaximumSize = New-Object System.Drawing.Size(800,600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# =========================
# CSV Import/Validation
# =========================
# Required columns for user property CSV
$requiredColumns = @("Office","Company","State","City","PostalCode","StreetAddress","Department","Title")
$script:csvData = $null
$script:csvError = $false
# Security group CSV (OU/group mapping)
$script:secGroupCsv = $null
$script:secGroupCsvError = $false
$script:ouPaths = @()
$script:ouGroups = @{}

function Populate-DropdownsFromCsv {
    param($data)
    $cmbOffice.Items.Clear()
    $cmbCompany.Items.Clear()
    $cmbState.Items.Clear()
    $cmbCity.Items.Clear()
    $cmbPostalCode.Items.Clear()
    $cmbStreetAddress.Items.Clear()
    $cmbDepartment.Items.Clear()
    $cmbTitle.Items.Clear()

    $cmbOffice.Items.AddRange(($data | Select-Object -ExpandProperty Office | Sort-Object -Unique))
    $cmbCompany.Items.AddRange(($data | Select-Object -ExpandProperty Company | Sort-Object -Unique))
    $cmbState.Items.AddRange(($data | Select-Object -ExpandProperty State | Sort-Object -Unique))
    $cmbCity.Items.AddRange(($data | Select-Object -ExpandProperty City | Sort-Object -Unique))
    $cmbPostalCode.Items.AddRange(($data | Select-Object -ExpandProperty PostalCode | Sort-Object -Unique))
    $cmbStreetAddress.Items.AddRange(($data | Select-Object -ExpandProperty StreetAddress | Sort-Object -Unique))
    $cmbDepartment.Items.AddRange(($data | Select-Object -ExpandProperty Department | Sort-Object -Unique))
    $cmbTitle.Items.AddRange(($data | Select-Object -ExpandProperty Title | Sort-Object -Unique))
}
# Security group CSV loader
function Load-SecGroupCsv {
    param($csvPath)
    $script:secGroupCsvError = $false
    $script:ouPaths = @()
    $script:ouGroups = @{}
    try {
        $csv = Import-Csv -Path $csvPath
        $headers = (Get-Content $csvPath -First 1).Split(',')
        foreach ($header in $headers) {
            if (-not $header.Trim().StartsWith('paradigmcos.local\')) {
                [System.Windows.Forms.MessageBox]::Show("OU column '$header' must start with 'paradigmcos.local\'", "CSV Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $script:secGroupCsvError = $true
                return
            }
        }
        $script:ouPaths = $headers
        foreach ($header in $headers) {
            $groups = @()
            foreach ($row in $csv) {
                $val = $row.$header
                if ($val -and $val.Trim() -ne '') { $groups += $val.Trim() }
            }
            $script:ouGroups[$header] = $groups
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error reading secgroup CSV: $($_.Exception.Message)", "CSV Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $script:secGroupCsvError = $true
    }
}

# =========================
# Page 1 - Add User Controls
# =========================
# CSV select button
$btnSelectCsv = New-Object System.Windows.Forms.Button
$btnSelectCsv.Text = "Select CSV"
$btnSelectCsv.Location = New-Object System.Drawing.Point(350,75)
$btnSelectCsv.Width = 100
$btnSelectCsv.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $openFileDialog.Title = "Select a CSV file for dropdown data"
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        try {
            $script:csvData = Import-Csv -Path $openFileDialog.FileName
            $csvColumns = $script:csvData | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
            $missing = $requiredColumns | Where-Object { $_ -notin $csvColumns }
            if ($missing.Count -gt 0) {
                [System.Windows.Forms.MessageBox]::Show("CSV missing columns: $($missing -join ', ')", "CSV Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $script:csvError = $true
            } else {
                $script:csvError = $false
                Populate-DropdownsFromCsv $script:csvData
                $cmbCompany.Items.Clear()
                $cmbCompany.Items.AddRange(($script:csvData | Select-Object -ExpandProperty Company -Unique))
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error reading CSV: $($_.Exception.Message)", "CSV Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $script:csvError = $true
        }
    }
})

$form.Controls.Add($btnSelectCsv)
# first name
$lblFirstName = New-Object System.Windows.Forms.Label
$lblFirstName.Text = "First Name:"
$lblFirstName.Location = New-Object System.Drawing.Point(175,150)
$lblFirstName.AutoSize = $true
$form.Controls.Add($lblFirstName)

$txtFirstName = New-Object System.Windows.Forms.TextBox
$txtFirstName.Location = New-Object System.Drawing.Point(300,148)
$txtFirstName.Width = 200
$txtFirstName.ShortcutsEnabled = $true 
$form.Controls.Add($txtFirstName)
$txtFirstName.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $sender.SelectAll()
        $e.SuppressKeyPress = $true
    }
})

# last name
$lblLastName = New-Object System.Windows.Forms.Label
$lblLastName.Text = "Last Name:"
$lblLastName.Location = New-Object System.Drawing.Point(175,200)
$lblLastName.AutoSize = $true
$form.Controls.Add($lblLastName)

$txtLastName = New-Object System.Windows.Forms.TextBox
$txtLastName.Location = New-Object System.Drawing.Point(300,198)
$txtLastName.Width = 200
$txtLastName.ShortcutsEnabled = $true  
$form.Controls.Add($txtLastName)
$txtLastName.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $sender.SelectAll()
        $e.SuppressKeyPress = $true
    }
})

# username, auto-populated from first and last name, editable
$lblUsername = New-Object System.Windows.Forms.Label
$lblUsername.Text = "Username:"
$lblUsername.Location = New-Object System.Drawing.Point(175,250)
$lblUsername.AutoSize = $true
$form.Controls.Add($lblUsername)

$txtUsername = New-Object System.Windows.Forms.TextBox
$txtUsername.Location = New-Object System.Drawing.Point(300,248)
$txtUsername.Width = 200
$txtUsername.ShortcutsEnabled = $true
$form.Controls.Add($txtUsername)

# company dropdown (populated from CSV)
$lblCompany = New-Object System.Windows.Forms.Label
$lblCompany.Text = "Company:"
$lblCompany.Location = New-Object System.Drawing.Point(175,300)
$lblCompany.AutoSize = $true
$form.Controls.Add($lblCompany)

$cmbCompany = New-Object System.Windows.Forms.ComboBox
$cmbCompany.Location = New-Object System.Drawing.Point(300,298)
$cmbCompany.Width = 200
$cmbCompany.DropDownStyle = 'DropDownList'
$form.Controls.Add($cmbCompany)

$btnBackToDashboardFromAdd = New-Object System.Windows.Forms.Button
$btnBackToDashboardFromAdd.Text = "Back to Dashboard"
$btnBackToDashboardFromAdd.Location = New-Object System.Drawing.Point(333,432)
$btnBackToDashboardFromAdd.Width = 130
$btnBackToDashboardFromAdd.Visible = $false
$btnBackToDashboardFromAdd.Add_Click({ $ShowDashboard.Invoke() })
$form.Controls.Add($btnBackToDashboardFromAdd)

$btnNext = New-Object System.Windows.Forms.Button
$btnNext.Text = "Next"
$btnNext.Location = New-Object System.Drawing.Point(350,371)
$btnNext.Width = 100
$form.Controls.Add($btnNext)
$btnNext.Add_Click({
    $firstName = $txtFirstName.Text.Trim()
    $lastName = $txtLastName.Text.Trim()
    $username = $txtUsername.Text.Trim()
    $company = $cmbCompany.SelectedItem
    if (-not $firstName -or -not $lastName -or -not $username -or -not $company) {
        [System.Windows.Forms.MessageBox]::Show("Please fill out all fields (First Name, Last Name, Username, Company) before continuing.", "Input Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ($script:csvError -or -not $script:csvData) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid CSV file before continuing.", "CSV Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    $ShowPage2.Invoke()
})

# =========================
# Username Auto-Population
# =========================
function Update-Username {
    $firstName = $txtFirstName.Text.Trim()
    $lastName = $txtLastName.Text.Trim()
    if ($firstName.Length -ge 1 -and $lastName.Length -ge 1) {
        $txtUsername.Text = ("{0}{1}" -f $firstName.Substring(0,1), $lastName).ToLower()
    } else {
        $txtUsername.Text = ""
    }
}
$txtFirstName.Add_TextChanged({ Update-Username })
$txtLastName.Add_TextChanged({ Update-Username })

# =========================
# Page 2 - Add User Controls
# =========================
# office
$lblOffice = New-Object System.Windows.Forms.Label
$lblOffice.Text = "Office:"
$lblOffice.Location = New-Object System.Drawing.Point(175,60)
$lblOffice.AutoSize = $true
$lblOffice.Visible = $false
$form.Controls.Add($lblOffice)

$cmbOffice = New-Object System.Windows.Forms.ComboBox
$cmbOffice.Location = New-Object System.Drawing.Point(300,58)
$cmbOffice.Width = 200
$cmbOffice.DropDownStyle = 'DropDownList'
$cmbOffice.Visible = $false
$form.Controls.Add($cmbOffice)

# dept
$lblDepartment = New-Object System.Windows.Forms.Label
$lblDepartment.Text = "Department:"
$lblDepartment.Location = New-Object System.Drawing.Point(175,100)
$lblDepartment.AutoSize = $true
$lblDepartment.Visible = $false
$form.Controls.Add($lblDepartment)

$cmbDepartment = New-Object System.Windows.Forms.ComboBox
$cmbDepartment.Location = New-Object System.Drawing.Point(300,98)
$cmbDepartment.Width = 200
$cmbDepartment.DropDownStyle = 'DropDownList'
$cmbDepartment.Visible = $false
$form.Controls.Add($cmbDepartment)

# job title
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Job Title:"
$lblTitle.Location = New-Object System.Drawing.Point(175,140)
$lblTitle.AutoSize = $true
$lblTitle.Visible = $false
$form.Controls.Add($lblTitle)

$cmbTitle = New-Object System.Windows.Forms.ComboBox
$cmbTitle.Location = New-Object System.Drawing.Point(300,138)
$cmbTitle.Width = 200
$cmbTitle.DropDownStyle = 'DropDownList'
$cmbTitle.Visible = $false
$form.Controls.Add($cmbTitle)

# street address
$lblStreet = New-Object System.Windows.Forms.Label
$lblStreet.Text = "Street Address:"
$lblStreet.Location = New-Object System.Drawing.Point(175,180)
$lblStreet.AutoSize = $true
$lblStreet.Visible = $false
$form.Controls.Add($lblStreet)

$cmbStreetAddress = New-Object System.Windows.Forms.ComboBox
$cmbStreetAddress.Location = New-Object System.Drawing.Point(300,178)
$cmbStreetAddress.Width = 200
$cmbStreetAddress.DropDownStyle = 'DropDownList'
$cmbStreetAddress.Visible = $false
$form.Controls.Add($cmbStreetAddress)

# city
$lblCity = New-Object System.Windows.Forms.Label
$lblCity.Text = "City:"
$lblCity.Location = New-Object System.Drawing.Point(175,220)
$lblCity.AutoSize = $true
$lblCity.Visible = $false
$form.Controls.Add($lblCity)

$cmbCity = New-Object System.Windows.Forms.ComboBox
$cmbCity.Location = New-Object System.Drawing.Point(300,218)
$cmbCity.Width = 200
$cmbCity.DropDownStyle = 'DropDownList'
$cmbCity.Visible = $false
$form.Controls.Add($cmbCity)

# state
$lblState = New-Object System.Windows.Forms.Label
$lblState.Text = "State:"
$lblState.Location = New-Object System.Drawing.Point(175,260)
$lblState.AutoSize = $true
$lblState.Visible = $false
$form.Controls.Add($lblState)

$cmbState = New-Object System.Windows.Forms.ComboBox
$cmbState.Location = New-Object System.Drawing.Point(300,258)
$cmbState.Width = 200
$cmbState.DropDownStyle = 'DropDownList'
$cmbState.Visible = $false
$form.Controls.Add($cmbState)

# post code
$lblPostalCode = New-Object System.Windows.Forms.Label
$lblPostalCode.Text = "Postal Code:"
$lblPostalCode.Location = New-Object System.Drawing.Point(175,300)
$lblPostalCode.AutoSize = $true
$lblPostalCode.Visible = $false
$form.Controls.Add($lblPostalCode)

$cmbPostalCode = New-Object System.Windows.Forms.ComboBox
$cmbPostalCode.Location = New-Object System.Drawing.Point(300,298)
$cmbPostalCode.Width = 200
$cmbPostalCode.DropDownStyle = 'DropDownList'
$cmbPostalCode.Visible = $false
$form.Controls.Add($cmbPostalCode)

# telephone number, manually entered
$lblTelephone = New-Object System.Windows.Forms.Label
$lblTelephone.Text = "Telephone Number:"
$lblTelephone.Location = New-Object System.Drawing.Point(175,340)
$lblTelephone.AutoSize = $true
$lblTelephone.Visible = $false
$form.Controls.Add($lblTelephone)

$txtTelephone = New-Object System.Windows.Forms.TextBox
$txtTelephone.Location = New-Object System.Drawing.Point(300,338)
$txtTelephone.Width = 200
$txtTelephone.ShortcutsEnabled = $true 
$txtTelephone.Visible = $false
$form.Controls.Add($txtTelephone)
$txtTelephone.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $sender.SelectAll()
        $e.SuppressKeyPress = $true
    }
})

$btnNext2 = New-Object System.Windows.Forms.Button
$btnNext2.Text = "Next"
$btnNext2.Location = New-Object System.Drawing.Point(450,450)
$btnNext2.Width = 100
$btnNext2.Visible = $false
$btnNext2.Add_Click({
    $office = $cmbOffice.SelectedItem
    $department = $cmbDepartment.SelectedItem
    $title = $cmbTitle.SelectedItem
    $street = $cmbStreetAddress.SelectedItem
    $city = $cmbCity.SelectedItem
    $state = $cmbState.SelectedItem
    $postalCode = $cmbPostalCode.SelectedItem
    $telephoneNum = $txtTelephone.Text.Trim()

    if (-not $office -or -not $department -or -not $title -or -not $street -or -not $city -or -not $state -or -not $postalCode -or -not $telephoneNum) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please fill out all fields (Office, Department, Job Title, Street Address, City, State, Postal Code, Telephone Number) before continuing.",
            "Input Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }
    $ShowPage3.Invoke()
})
$form.Controls.Add($btnNext2)

$btnBack = New-Object System.Windows.Forms.Button
$btnBack.Text = "Back"
$btnBack.Location = New-Object System.Drawing.Point(200,450)
$btnBack.Width = 100
$btnBack.Visible = $false
$btnBack.Add_Click({
    $ShowPage1.Invoke()
})
$form.Controls.Add($btnBack)

# =========================
# Page 3 - Add User Controls
# =========================
# OU Selection (CSV-driven)
$lblOU = New-Object System.Windows.Forms.Label
$lblOU.Text = "Organizational Unit:"
$lblOU.Location = New-Object System.Drawing.Point(175,60)
$lblOU.AutoSize = $true
$lblOU.Visible = $false
$form.Controls.Add($lblOU)

$cmbOU = New-Object System.Windows.Forms.ComboBox
$cmbOU.Location = New-Object System.Drawing.Point(300,58)
$cmbOU.Width = 300
$cmbOU.DropDownStyle = 'DropDownList'
$cmbOU.Visible = $false
$form.Controls.Add($cmbOU)
# Sub-OU Selection (CSV-driven)
$lblSubOU = New-Object System.Windows.Forms.Label
$lblSubOU.Text = "Sub-Directory:"
$lblSubOU.Location = New-Object System.Drawing.Point(175,100)
$lblSubOU.AutoSize = $true
$lblSubOU.Visible = $false
$form.Controls.Add($lblSubOU)

$cmbSubOU = New-Object System.Windows.Forms.ComboBox
$cmbSubOU.Location = New-Object System.Drawing.Point(300,98)
$cmbSubOU.Width = 300
$cmbSubOU.DropDownStyle = 'DropDownList'
$cmbSubOU.Visible = $false
$form.Controls.Add($cmbSubOU)
# Manager Role Dropdown
$lblMgrRole = New-Object System.Windows.Forms.Label
$lblMgrRole.Text = "Manager Role:"
$lblMgrRole.Location = New-Object System.Drawing.Point(175,140)
$lblMgrRole.AutoSize = $true
$lblMgrRole.Visible = $false
$form.Controls.Add($lblMgrRole)

$cmbMgrRole = New-Object System.Windows.Forms.ComboBox
$cmbMgrRole.Location = New-Object System.Drawing.Point(300,138)
$cmbMgrRole.Width = 200
$cmbMgrRole.DropDownStyle = 'DropDownList'
$cmbMgrRole.Items.AddRange(@("Not a manager", "Assistant Manager", "Manager"))
$cmbMgrRole.SelectedIndex = 0
$cmbMgrRole.Visible = $false
$form.Controls.Add($cmbMgrRole)

$txtSummary = New-Object System.Windows.Forms.TextBox
$txtSummary.Multiline = $true
$txtSummary.ReadOnly = $true
$txtSummary.Location = New-Object System.Drawing.Point(175,180)
$txtSummary.Size = New-Object System.Drawing.Size(425, 270)
$txtSummary.ScrollBars = 'Vertical'
$txtSummary.Font = New-Object System.Drawing.Font("Consolas",10)
$txtSummary.Visible = $false
$form.Controls.Add($txtSummary)
# OU CSV select button (Page 3)
$btnSelectOUCsv = New-Object System.Windows.Forms.Button
$btnSelectOUCsv.Text = "Select OU CSV"
$btnSelectOUCsv.Location = New-Object System.Drawing.Point(30, 30)
$btnSelectOUCsv.Width = 120
$btnSelectOUCsv.Visible = $false
$btnSelectOUCsv.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $openFileDialog.Title = "Select OU/Security Group Mapping CSV"
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        Load-SecGroupCsv $openFileDialog.FileName
        if (-not $script:secGroupCsvError) {
            Initialize-OUDropdowns
        }
    }
})
$form.Controls.Add($btnSelectOUCsv)

$btnCreateUser = New-Object System.Windows.Forms.Button
$btnCreateUser.Text = "Create User"
$btnCreateUser.Location = New-Object System.Drawing.Point(500,485)
$btnCreateUser.Width = 100
$btnCreateUser.Visible = $false
$btnCreateUser.Enabled = $false
$btnCreateUser.Add_Click({
    if ($btnCreateUser.Visible -and $btnCreateUser.Enabled) {
        Create-ADUserFromForm
    }
})
$form.Controls.Add($btnCreateUser)

$btnBack2 = New-Object System.Windows.Forms.Button
$btnBack2.Text = "Back"
$btnBack2.Location = New-Object System.Drawing.Point(200,485)
$btnBack2.Width = 100
$btnBack2.Visible = $false
$btnBack2.Add_Click({
    $ShowPage2.Invoke()
})
$form.Controls.Add($btnBack2)

# =========================
# OU Tree Structure (supports arbitrary depth)
# =========================
# CSV-driven OU dropdown initialisation
function Initialize-OUDropdowns {
    $cmbOU.Items.Clear()
    $cmbSubOU.Items.Clear()
    $lblSubOU.Visible = $false
    $cmbSubOU.Visible = $false
    $lblMgrRole.Visible = $false
    $cmbMgrRole.Visible = $false

    # Only show first two directory levels in OU dropdown
    $cmbOU.Items.Clear()
    $mainOUs = @()
    foreach ($ouPath in $script:ouPaths) {
        $parts = $ouPath.Split('\')
        if ($parts.Count -ge 3) {
            $mainOU = ($parts[0..1] -join '\')
            if ($mainOUs -notcontains $mainOU) { $mainOUs += $mainOU }
        } else {
            $mainOU = $ouPath
            if ($mainOUs -notcontains $mainOU) { $mainOUs += $mainOU }
        }
    }
    foreach ($mainOU in $mainOUs) {
        $cmbOU.Items.Add($mainOU)
    }
    $cmbSubOU.Items.Clear()
    $lblSubOU.Visible = $false
    $cmbSubOU.Visible = $false
    $lblMgrRole.Visible = $false
    $cmbMgrRole.Visible = $false
}

function Update-SubOUDropdown {
    $cmbSubOU.Items.Clear()
    $lblSubOU.Visible = $false
    $cmbSubOU.Visible = $false
    $lblMgrRole.Visible = $false
    $cmbMgrRole.Visible = $false
    $selectedOU = $cmbOU.SelectedItem
    if (-not $selectedOU) { return }
    # finds all subdirectory options for selected main OU (after second '\')
    $subOUs = @()
    foreach ($ouPath in $script:ouPaths) {
        $parts = $ouPath.Split('\')
        if ($parts.Count -ge 3) {
            $mainOU = ($parts[0..1] -join '\')
            if ($mainOU -eq $selectedOU) {
                if ($parts.Count -gt 2) {
                    $subdir = ($parts[2..($parts.Count-1)] -join '\')
                    if ($subdir -and $subOUs -notcontains $subdir) { $subOUs += $subdir }
                }
            }
        }
    }
    if ($subOUs.Count -gt 0) {
        $lblSubOU.Visible = $true
        $cmbSubOU.Visible = $true
        foreach ($sub in $subOUs) { $cmbSubOU.Items.Add($sub) }
    }
    # displays manager role dropdown if "PDC-MANAGEMENT" in main OU selection
    if ($selectedOU -like '*PDC-MANAGEMENT*') {
        $lblMgrRole.Visible = $true
        $cmbMgrRole.Visible = $true
    }
}

function Get-FullOUPath {
    $selectedOU = $cmbOU.SelectedItem
    $selectedSubOU = $cmbSubOU.SelectedItem
    if (-not $selectedOU) { return $null }

    # finds full OU path from CSV
    $fullPath = $selectedOU
    if ($selectedSubOU -and $selectedSubOU.Trim() -ne '') {
        $fullPath = "$selectedOU\$selectedSubOU"
    }
    # validates against loaded OU paths
    foreach ($ouPath in $script:ouPaths) {
        if ($ouPath -eq $fullPath) { return $ouPath }
    }
    # if cannot be found, fallback to main OU
    return $selectedOU
}
# Event handler for OU selection change
$cmbOU.Add_SelectedIndexChanged({
    Update-SubOUDropdown
    Update-Summary
})
$cmbSubOU.Add_SelectedIndexChanged({
    Update-Summary
})
$cmbMgrRole.Add_SelectedIndexChanged({
    Update-Summary
})

# =========================
# Summary Update Function
# =========================
function Update-Summary {
    $domainRoot = 'paradigmcos.local'
    $selectedOU = $cmbOU.SelectedItem
    $selectedSubOU = $cmbSubOU.SelectedItem
    $mgrRole = if ($cmbMgrRole.Visible) { $cmbMgrRole.SelectedItem } else { "Not a manager" }

    $ouPath = $selectedOU
    if ($selectedSubOU -and $selectedSubOU.Trim() -ne '') {
        $ouPath = "$selectedOU\$selectedSubOU"
    }
    $ldapOUPath = Convert-OUPathToLDAP $ouPath
    $summary = @()
    $summary += 'User Summary:'
    $summary += '============='
    $summary += "Name: $($txtFirstName.Text) $($txtLastName.Text)"
    $summary += "Username: $($txtUsername.Text)"
    $summary += "Company: $($cmbCompany.SelectedItem)"
    $summary += "Office: $($cmbOffice.SelectedItem)"
    $summary += "Department: $($cmbDepartment.SelectedItem)"
    $summary += "Title: $($cmbTitle.SelectedItem)"
    $summary += "Street Address: $($cmbStreetAddress.SelectedItem)"
    $summary += "City: $($cmbCity.SelectedItem)"
    $summary += "State: $($cmbState.SelectedItem)"
    $summary += "Postal Code: $($cmbPostalCode.SelectedItem)"
    $summary += "Telephone: $($txtTelephone.Text)"
    $summary += "Email: $($txtUsername.Text)@paradigmcos.com"
    $summary += "OU Path (CSV): $ouPath"
    $summary += "OU Path (AD): $ldapOUPath"
    $summary += "Manager Role: $mgrRole"
    $summary += ''
    $summary += "Please review the information above and click 'Create User' to proceed."
    $txtSummary.Text = $summary -join "`r`n"
}

# =========================
# Add User Page Switch Logic
# =========================
## contents of page 1 -- add user path
$ShowPage1 = {
    # Hides dashboard controls
    $btnGoToAddUser.Visible = $false
    # Shows add user page 1 controls
    $lblFirstName.Visible = $true
    $txtFirstName.Visible = $true
    $lblLastName.Visible = $true
    $txtLastName.Visible = $true
    $lblUsername.Visible = $true
    $txtUsername.Visible = $true
    $lblCompany.Visible = $true
    $cmbCompany.Visible = $true
    $btnNext.Visible = $true
    $btnSelectCsv.Visible = $true
    $btnBackToDashboardFromAdd.Visible = $true
    
    $btnBack.Visible = $false
    $btnNext2.Visible = $false
    $lblOffice.Visible = $false
    $cmbOffice.Visible = $false
    $lblState.Visible = $false
    $cmbState.Visible = $false
    $lblCity.Visible = $false
    $cmbCity.Visible = $false
    $lblPostalCode.Visible = $false
    $cmbPostalCode.Visible = $false
    $lblStreet.Visible = $false
    $cmbStreetAddress.Visible = $false
    $lblDepartment.Visible = $false
    $cmbDepartment.Visible = $false
    $lblTitle.Visible = $false
    $cmbTitle.Visible = $false
    $lblTelephone.Visible = $false
    $txtTelephone.Visible = $false
    #$btnSubmit.Visible = $false

    $btnBack2.Visible = $false
    $txtSummary.Visible = $false
    $btnCreateUser.Visible = $false
    $lblOU.Visible = $false
    $cmbOU.Visible = $false
    $lblSubOU.Visible = $false
    $cmbSubOU.Visible = $false
    $lblMgrRole.Visible = $false
    $cmbMgrRole.Visible = $false
    $btnSelectOUCsv.Visible = $false
}

## contents of page 2 -- add user path
$ShowPage2 = {
    $lblFirstName.Visible = $false
    $txtFirstName.Visible = $false
    $lblLastName.Visible = $false
    $txtLastName.Visible = $false
    $lblUsername.Visible = $false
    $txtUsername.Visible = $false
    $lblCompany.Visible = $false
    $cmbCompany.Visible = $false
    $btnSelectCsv.Visible = $false
    $btnNext.Visible = $false
    $btnBackToDashboardFromAdd.Visible = $false
    
    $btnBack.Visible = $true
    $btnNext2.Visible = $true

    $btnBack2.Visible = $false
    $txtSummary.Visible = $false
    $btnCreateUser.Visible = $false  
    $lblOU.Visible = $false
    $cmbOU.Visible = $false
    $lblSubOU.Visible = $false
    $cmbSubOU.Visible = $false

    $lblOffice.Visible = $true
    $cmbOffice.Visible = $true
    $lblDepartment.Visible = $true
    $cmbDepartment.Visible = $true
    $lblTitle.Visible = $true
    $cmbTitle.Visible = $true
    $lblStreet.Visible = $true
    $cmbStreetAddress.Visible = $true
    $lblCity.Visible = $true
    $cmbCity.Visible = $true
    $lblState.Visible = $true
    $cmbState.Visible = $true
    $lblPostalCode.Visible = $true
    $cmbPostalCode.Visible = $true
    $lblTelephone.Visible = $true
    $txtTelephone.Visible = $true
    #$btnSubmit.Visible = $true
}

## contents of page 3 -- add user path
$ShowPage3 = {
    # Hides all page 2 controls
    $lblOffice.Visible = $false
    $cmbOffice.Visible = $false
    $lblDepartment.Visible = $false
    $cmbDepartment.Visible = $false
    $lblTitle.Visible = $false
    $cmbTitle.Visible = $false
    $lblStreet.Visible = $false
    $cmbStreetAddress.Visible = $false
    $lblCity.Visible = $false
    $cmbCity.Visible = $false
    $lblState.Visible = $false
    $cmbState.Visible = $false
    $lblPostalCode.Visible = $false
    $cmbPostalCode.Visible = $false
    $lblTelephone.Visible = $false
    $txtTelephone.Visible = $false
    $btnBack.Visible = $false
    $btnNext2.Visible = $false

    # Hides all page 1 controls
    $lblFirstName.Visible = $false
    $txtFirstName.Visible = $false
    $lblLastName.Visible = $false
    $txtLastName.Visible = $false
    $lblUsername.Visible = $false
    $txtUsername.Visible = $false
    $lblCompany.Visible = $false
    $cmbCompany.Visible = $false
    $btnSelectCsv.Visible = $false
    $btnNext.Visible = $false
    $btnBackToDashboardFromAdd.Visible = $false

    # Shows only page 3 controls
    Initialize-OUDropdowns
    $lblOU.Visible = $true
    $cmbOU.Visible = $true
    $btnSelectOUCsv.Visible = $true
    $btnCreateUser.Visible = $true
    $btnCreateUser.Enabled = $true
    $btnBack2.Visible = $true
    $txtSummary.Visible = $true
    Update-Summary
}

$ShowDashboard = {
    # Shows dashboard controls
    $btnGoToAddUser.Visible = $true
    # Hides all add user controls (page 1, 2, 3)
    $lblFirstName.Visible = $false
    $txtFirstName.Visible = $false
    $lblLastName.Visible = $false
    $txtLastName.Visible = $false
    $lblUsername.Visible = $false
    $txtUsername.Visible = $false
    $lblCompany.Visible = $false
    $cmbCompany.Visible = $false
    $btnNext.Visible = $false
    $btnSelectCsv.Visible = $false
    $btnBackToDashboardFromAdd.Visible = $false
    $btnBack.Visible = $false
    $btnNext2.Visible = $false
    $lblOffice.Visible = $false
    $cmbOffice.Visible = $false
    $lblState.Visible = $false
    $cmbState.Visible = $false
    $lblCity.Visible = $false
    $cmbCity.Visible = $false
    $lblPostalCode.Visible = $false
    $cmbPostalCode.Visible = $false
    $lblStreet.Visible = $false
    $cmbStreetAddress.Visible = $false
    $lblDepartment.Visible = $false
    $cmbDepartment.Visible = $false
    $lblTitle.Visible = $false
    $cmbTitle.Visible = $false
    $lblTelephone.Visible = $false
    $txtTelephone.Visible = $false
    $btnBack2.Visible = $false
    $txtSummary.Visible = $false
    $btnCreateUser.Visible = $false
    $lblOU.Visible = $false
    $cmbOU.Visible = $false
    $lblSubOU.Visible = $false
    $cmbSubOU.Visible = $false
}
# =========================
# Dashboard Controls
# =========================
$btnGoToAddUser = New-Object System.Windows.Forms.Button
$btnGoToAddUser.Text = "Add New User"
$btnGoToAddUser.Location = New-Object System.Drawing.Point(300,240)
$btnGoToAddUser.Width = 200
$btnGoToAddUser.Height = 50
$btnGoToAddUser.Visible = $true
$btnGoToAddUser.Add_Click({ $ShowPage1.Invoke() })
$form.Controls.Add($btnGoToAddUser)

# =========================
# OU Path Conversion Utility
# =========================
function Convert-OUPathToLDAP {
    param(
        [string]$ouPath
    )
    if (-not $ouPath -or $ouPath.Trim() -eq '') { return $null }
    # Only processes if starts with paradigmcos.local\
    if ($ouPath -notlike 'paradigmcos.local*') { return $null }
    # Hardcoded exception for paradigmcos.local\Users
    if ($ouPath -eq 'paradigmcos.local\Users') {
        $domain = 'paradigmcos.local'
        $domainDCs = $domain.Split('.') | ForEach-Object { 'DC=' + $_ }
        $dcString = $domainDCs -join ','
        return "CN=Users,$dcString"
    }
    $parts = $ouPath.Split('\')
    if ($parts.Count -lt 2) { return $null }
    # First part is domain
    $domain = $parts[0]
    $ouParts = $parts[1..($parts.Count-1)]
    # Reverses OU parts for correct LDAP order (leaf to root)
    $ouPartsReversed = [System.Collections.ArrayList]::new()
    $ouPartsReversed.AddRange($ouParts)
    [void]$ouPartsReversed.Reverse()
    $ouString = ($ouPartsReversed | ForEach-Object { 'OU=' + $_ }) -join ','
    # Builds DC string
    $domainDCs = $domain.Split('.') | ForEach-Object { 'DC=' + $_ }
    $dcString = $domainDCs -join ','
    # Combines
    $ldapPath = if ($ouString) { "$ouString,$dcString" } else { $dcString }
    return $ldapPath
}
# =========================
# Initial Page Load
# =========================
$ShowDashboard.Invoke()
[void]$form.ShowDialog()