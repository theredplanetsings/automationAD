<#============================================
  Automation AD - Main Interface Script
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
    
    # Check if username already exists
    try {
        $existingUser = Get-ADUser -Filter "SamAccountName -eq '$username'" -ErrorAction SilentlyContinue
        if ($existingUser) {
            [System.Windows.Forms.MessageBox]::Show("Username '$username' already exists. Please choose a different username.", "Username Conflict", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    } catch {
        # Continue if user doesn't exist
    }
    # Check if UPN already exists
    try {
        $existingUPN = Get-ADUser -Filter "UserPrincipalName -eq '$userPrincipalName'" -ErrorAction SilentlyContinue
        if ($existingUPN) {
            [System.Windows.Forms.MessageBox]::Show("User Principal Name '$userPrincipalName' already exists. Please choose a different username.", "UPN Conflict", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    } catch {
        # Continue if UPN doesn't exist
    }
    # retrieve values from dropdowns and textboxes - convert to strings
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
    
    # Debug output - remove this after testing
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
    
    # Validate OU selection
    $ouPath = Get-FullOUPath
    if (-not $ouPath) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid Organizational Unit. If you selected a main OU with sub-directories, you must also select a sub-directory.", "OU Selection Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
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
            -Path $ouPath `
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
## update as needed, determines the columns required in the CSV file to autopopulate dropdowns.
# must add additional dropdown sections to this program if you add more columns to the CSV.
$requiredColumns = @("Office","Company","State","City","PostalCode","StreetAddress","Department","Title")
# Use $script: scope for csvData and csvError so all handlers see the same values
$script:csvData = $null
$script:csvError = $false
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
# OU Selection
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

# Sub-OU Selection (appears when main OU has sub-directories)
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

$txtSummary = New-Object System.Windows.Forms.TextBox
$txtSummary.Multiline = $true
$txtSummary.ReadOnly = $true
$txtSummary.Location = New-Object System.Drawing.Point(175,140)
$txtSummary.Size = New-Object System.Drawing.Size(425, 270)
$txtSummary.ScrollBars = 'Vertical'
$txtSummary.Font = New-Object System.Drawing.Font("Consolas",10)
$txtSummary.Visible = $false
$form.Controls.Add($txtSummary)

$btnCreateUser = New-Object System.Windows.Forms.Button
$btnCreateUser.Text = "Create User"
$btnCreateUser.Location = New-Object System.Drawing.Point(500,450)
$btnCreateUser.Width = 100
$btnCreateUser.Visible = $false
$btnCreateUser.Add_Click({ Create-ADUserFromForm })
$form.Controls.Add($btnCreateUser)

$btnBack2 = New-Object System.Windows.Forms.Button
$btnBack2.Text = "Back"
$btnBack2.Location = New-Object System.Drawing.Point(200,450)
$btnBack2.Width = 100
$btnBack2.Visible = $false
$btnBack2.Add_Click({
    $ShowPage2.Invoke()
})
$form.Controls.Add($btnBack2)



# =========================
# OU Management Functions
# =========================
function Initialize-OUDropdowns {
    $cmbOU.Items.Clear()
    $cmbSubOU.Items.Clear()
    
    # Add main OU options
    $cmbOU.Items.Add("Users")
    $cmbOU.Items.Add("Service Accounts")
    $cmbOU.Items.Add("PDC-CONSTRUCTION\USERS")
    $cmbOU.Items.Add("PDC-HQ\USERS")
    $cmbOU.Items.Add("PDC-SERVICES\USERS")
}

function Update-SubOUDropdown {
    $cmbSubOU.Items.Clear()
    $selectedOU = $cmbOU.SelectedItem
    # Show/hide sub-OU dropdown based on selection
    if ($selectedOU -eq "PDC-CONSTRUCTION\USERS" -or $selectedOU -eq "PDC-HQ\USERS" -or $selectedOU -eq "PDC-SERVICES\USERS") {
        $lblSubOU.Visible = $true
        $cmbSubOU.Visible = $true
        $cmbSubOU.Items.Add("External")
        $cmbSubOU.Items.Add("Internal")
    } else {
        $lblSubOU.Visible = $false
        $cmbSubOU.Visible = $false
    }
}

function Get-FullOUPath {
    $selectedOU = $cmbOU.SelectedItem
    $selectedSubOU = $cmbSubOU.SelectedItem
    
    switch ($selectedOU) {
        "Users" { return "OU=Users,DC=paradigmcos,DC=local" }
        "Service Accounts" { return "OU=Service Accounts,DC=paradigmcos,DC=local" }
        "PDC-CONSTRUCTION\USERS" { 
            if ($selectedSubOU) {
                return "OU=$selectedSubOU,OU=Users,OU=PDC-CONSTRUCTION,DC=paradigmcos,DC=local"
            }
            return $null
        }
        "PDC-HQ\USERS" { 
            if ($selectedSubOU) {
                return "OU=$selectedSubOU,OU=Users,OU=PDC-HQ,DC=paradigmcos,DC=local"
            }
            return $null
        }
        "PDC-SERVICES\USERS" { 
            if ($selectedSubOU) {
                return "OU=$selectedSubOU,OU=Users,OU=PDC-SERVICES,DC=paradigmcos,DC=local"
            }
            return $null
        }
    }
    return $null
}

# Add event handler for OU selection change
$cmbOU.Add_SelectedIndexChanged({
    Update-SubOUDropdown
    Update-Summary
})
$cmbSubOU.Add_SelectedIndexChanged({
    Update-Summary
})

# =========================
# Summary Update Function
# =========================
function Update-Summary {
    $ouPath = Get-FullOUPath
    $ouDisplay = if ($cmbSubOU.Visible -and $cmbSubOU.SelectedItem) {
        "$($cmbOU.SelectedItem)\$($cmbSubOU.SelectedItem)"
    } else {
        $cmbOU.SelectedItem
    }
    
    $summary = @"
User Summary:
=============
Name: $($txtFirstName.Text) $($txtLastName.Text)
Username: $($txtUsername.Text)
Company: $($cmbCompany.SelectedItem)
Office: $($cmbOffice.SelectedItem)
Department: $($cmbDepartment.SelectedItem)
Title: $($cmbTitle.SelectedItem)
Street Address: $($cmbStreetAddress.SelectedItem)
City: $($cmbCity.SelectedItem)
State: $($cmbState.SelectedItem)
Postal Code: $($cmbPostalCode.SelectedItem)
Telephone: $($txtTelephone.Text)
Email: $($txtUsername.Text)@paradigmcos.com
Organizational Unit: $ouDisplay

AD Path: $ouPath

Please review the information above and click 'Create User' to proceed.
"@
    $txtSummary.Text = $summary
}

# =========================
# Add User Page Switch Logic
# =========================
## contents of page 1 -- add user path
$ShowPage1 = {
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

    $btnGoToSearch.Visible = $false
    $btnGoToAddUser.Visible = $false
    $btnGoToLicenses.Visible = $false

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
    # Initialize and show OU controls
    Initialize-OUDropdowns
    $lblOU.Visible = $true
    $cmbOU.Visible = $true
    $txtSummary.Visible = $true
    $btnCreateUser.Visible = $true
    $btnBack2.Visible = $true
    Update-Summary

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

    $btnBack.Visible = $false
    $btnNext2.Visible = $false

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
    #$btnSubmit.Visible = $false
}


# =========================
# Dashboard Controls
# =========================
$btnGoToSearch = New-Object System.Windows.Forms.Button
$btnGoToSearch.Text = "Search/View Existing User"
$btnGoToSearch.Location = New-Object System.Drawing.Point(300,180)
$btnGoToSearch.Width = 200
$btnGoToSearch.Height = 50
$btnGoToSearch.Visible = $true
$btnGoToSearch.Add_Click({ $ShowUserSearch.Invoke() })
$form.Controls.Add($btnGoToSearch)

$btnGoToAddUser = New-Object System.Windows.Forms.Button
$btnGoToAddUser.Text = "Add New User"
$btnGoToAddUser.Location = New-Object System.Drawing.Point(300,260)
$btnGoToAddUser.Width = 200
$btnGoToAddUser.Height = 50
$btnGoToAddUser.Visible = $true
$btnGoToAddUser.Add_Click({ $ShowPage1.Invoke() })
$form.Controls.Add($btnGoToAddUser)

$btnGoToLicenses = New-Object System.Windows.Forms.Button
$btnGoToLicenses.Text = "Licenses"
$btnGoToLicenses.Location = New-Object System.Drawing.Point(300,340)
$btnGoToLicenses.Width = 200
$btnGoToLicenses.Height = 50
$btnGoToLicenses.Visible = $true
$form.Controls.Add($btnGoToLicenses)



# =========================
# User Search/Summary Page Controls
# =========================
$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "Search for User (Username, First, or Last Name):"
$lblSearch.Location = New-Object System.Drawing.Point(200,50)
$lblSearch.AutoSize = $true
$lblSearch.Visible = $false
$form.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(200,70)
$txtSearch.Width = 250
$txtSearch.Visible = $false
$form.Controls.Add($txtSearch)

$btnSearchUser = New-Object System.Windows.Forms.Button
$btnSearchUser.Text = "Search"
$btnSearchUser.Location = New-Object System.Drawing.Point(470,68)
$btnSearchUser.Width = 80
$btnSearchUser.Visible = $false
$form.Controls.Add($btnSearchUser)

$lstSearchResults = New-Object System.Windows.Forms.ListBox
$lstSearchResults.Location = New-Object System.Drawing.Point(200,100)
$lstSearchResults.Size = New-Object System.Drawing.Size(350,150)
$lstSearchResults.Visible = $false
$form.Controls.Add($lstSearchResults)

$txtUserSummary = New-Object System.Windows.Forms.TextBox
$txtUserSummary.Multiline = $true
$txtUserSummary.ReadOnly = $true
$txtUserSummary.Location = New-Object System.Drawing.Point(200,260)
$txtUserSummary.Size = New-Object System.Drawing.Size(350,150)
$txtUserSummary.ScrollBars = 'Vertical'
$txtUserSummary.Font = New-Object System.Drawing.Font("Consolas",10)
$txtUserSummary.Visible = $false
$form.Controls.Add($txtUserSummary)

$btnBackToDashboard = New-Object System.Windows.Forms.Button
$btnBackToDashboard.Text = "Back to Dashboard"
$btnBackToDashboard.Location = New-Object System.Drawing.Point(300,425)
$btnBackToDashboard.Width = 150
$btnBackToDashboard.Visible = $false
$btnBackToDashboard.Add_Click({ $ShowDashboard.Invoke() })
$form.Controls.Add($btnBackToDashboard)


# =========================
# License Management Page Controls
# =========================
$lblLicenseSearch = New-Object System.Windows.Forms.Label
$lblLicenseSearch.Text = "Search for User (Username or Email):"
$lblLicenseSearch.Location = New-Object System.Drawing.Point(200,65)
$lblLicenseSearch.AutoSize = $true
$lblLicenseSearch.Visible = $false
$form.Controls.Add($lblLicenseSearch)

$txtLicenseSearch = New-Object System.Windows.Forms.TextBox
$txtLicenseSearch.Location = New-Object System.Drawing.Point(200,85)
$txtLicenseSearch.Width = 250
$txtLicenseSearch.Visible = $false
$form.Controls.Add($txtLicenseSearch)

$btnLicenseUserSearch = New-Object System.Windows.Forms.Button
$btnLicenseUserSearch.Text = "Search"
$btnLicenseUserSearch.Location = New-Object System.Drawing.Point(470,83)
$btnLicenseUserSearch.Width = 80
$btnLicenseUserSearch.Visible = $false
$form.Controls.Add($btnLicenseUserSearch)

$lstLicenseUserResults = New-Object System.Windows.Forms.ListBox
$lstLicenseUserResults.Location = New-Object System.Drawing.Point(200,120)
$lstLicenseUserResults.Size = New-Object System.Drawing.Size(350,95)
$lstLicenseUserResults.Visible = $false
$form.Controls.Add($lstLicenseUserResults)

$lblCurrentLicenses = New-Object System.Windows.Forms.Label
$lblCurrentLicenses.Text = "Current Licenses:"
$lblCurrentLicenses.Location = New-Object System.Drawing.Point(200,220)
$lblCurrentLicenses.AutoSize = $true
$lblCurrentLicenses.Visible = $falseha
$form.Controls.Add($lblCurrentLicenses)

$txtCurrentLicenses = New-Object System.Windows.Forms.TextBox
$txtCurrentLicenses.Multiline = $true
$txtCurrentLicenses.ReadOnly = $true
$txtCurrentLicenses.Location = New-Object System.Drawing.Point(200,240)
$txtCurrentLicenses.Size = New-Object System.Drawing.Size(350,95)
$txtCurrentLicenses.ScrollBars = 'Vertical'
$txtCurrentLicenses.Font = New-Object System.Drawing.Font("Consolas",10)
$txtCurrentLicenses.Visible = $false
$form.Controls.Add($txtCurrentLicenses)

$lblAvailableLicenses = New-Object System.Windows.Forms.Label
$lblAvailableLicenses.Text = "Available Licenses:"
$lblAvailableLicenses.Location = New-Object System.Drawing.Point(200,340)
$lblAvailableLicenses.AutoSize = $true
$lblAvailableLicenses.Visible = $false
$form.Controls.Add($lblAvailableLicenses)

$clbAvailableLicenses = New-Object System.Windows.Forms.CheckedListBox
$clbAvailableLicenses.Location = New-Object System.Drawing.Point(200,360)
$clbAvailableLicenses.Size = New-Object System.Drawing.Size(350,95)
$clbAvailableLicenses.Visible = $false
$form.Controls.Add($clbAvailableLicenses)

$clbAvailableLicenses.IntegralHeight = $false
$clbAvailableLicenses.TabStop = $false
$clbAvailableLicenses.AllowDrop = $false
$clbAvailableLicenses.Enabled = $true

$clbAvailableLicenses.Add_KeyPress({ $_.Handled = $true })
$clbAvailableLicenses.Add_KeyDown({
    $allowed = @([System.Windows.Forms.Keys]::Up, [System.Windows.Forms.Keys]::Down, [System.Windows.Forms.Keys]::Space, [System.Windows.Forms.Keys]::Return)
    if ($allowed -notcontains $_.KeyCode) { $_.SuppressKeyPress = $true }
})

$btnAssignLicenses = New-Object System.Windows.Forms.Button
$btnAssignLicenses.Text = "Assign Selected Licenses"
$btnAssignLicenses.Location = New-Object System.Drawing.Point(200,480)
$btnAssignLicenses.Width = 200
$btnAssignLicenses.Visible = $false
$form.Controls.Add($btnAssignLicenses)

$btnBackToDashboardFromLicenses = New-Object System.Windows.Forms.Button
$btnBackToDashboardFromLicenses.Text = "Back to Dashboard"
$btnBackToDashboardFromLicenses.Location = New-Object System.Drawing.Point(420,480)
$btnBackToDashboardFromLicenses.Width = 130
$btnBackToDashboardFromLicenses.Visible = $false
$btnBackToDashboardFromLicenses.Add_Click({ $ShowDashboard.Invoke() })
$form.Controls.Add($btnBackToDashboardFromLicenses)


# =========================
# Page Switch Logic
# =========================
$ShowDashboard = {
    $btnGoToSearch.Visible = $true
    $btnGoToAddUser.Visible = $true
    $btnGoToLicenses.Visible = $true
    $lblSearch.Visible = $false
    $txtSearch.Visible = $false
    $btnSearchUser.Visible = $false
    $lstSearchResults.Visible = $false
    $txtUserSummary.Visible = $false
    $btnBackToDashboard.Visible = $false
    $btnBackToDashboardFromAdd.Visible = $false
    $btnBackToDashboardFromLicenses.Visible = $false
    # hides user add controls
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
    # hides licenses page controls
    $lblLicenseSearch.Visible = $false
    $txtLicenseSearch.Visible = $false
    $btnLicenseUserSearch.Visible = $false
    $lstLicenseUserResults.Visible = $false
    $lblCurrentLicenses.Visible = $false
    $txtCurrentLicenses.Visible = $false
    $lblAvailableLicenses.Visible = $false
    $clbAvailableLicenses.Visible = $false
    $btnAssignLicenses.Visible = $false
    # Hide OU controls
    $lblOU.Visible = $false
    $cmbOU.Visible = $false
    $lblSubOU.Visible = $false
    $cmbSubOU.Visible = $false
}
$ShowUserSearch = {
    $btnGoToSearch.Visible = $false
    $btnGoToAddUser.Visible = $false
    $btnGoToLicenses.Visible = $false
    $lblSearch.Visible = $true
    $txtSearch.Visible = $true
    $btnSearchUser.Visible = $true
    $lstSearchResults.Visible = $true
    $txtUserSummary.Visible = $true
    $btnBackToDashboard.Visible = $true
    # hides user add controls
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
}
$ShowLicensesPage = {
    $btnGoToSearch.Visible = $false
    $btnGoToAddUser.Visible = $false
    $btnGoToLicenses.Visible = $false
    $lblLicenseSearch.Visible = $true
    $txtLicenseSearch.Visible = $true
    $btnLicenseUserSearch.Visible = $true
    $lstLicenseUserResults.Visible = $true
    $lblCurrentLicenses.Visible = $true
    $txtCurrentLicenses.Visible = $true
    $lblAvailableLicenses.Visible = $true
    $clbAvailableLicenses.Visible = $true
    $btnAssignLicenses.Visible = $true
    $btnBackToDashboardFromLicenses.Visible = $true
    # Hide other controls
    $lblSearch.Visible = $false
    $txtSearch.Visible = $false
    $btnSearchUser.Visible = $false
    $lstSearchResults.Visible = $false
    $txtUserSummary.Visible = $false
    $btnBackToDashboard.Visible = $false
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
}
$btnGoToLicenses.Add_Click({ $ShowLicensesPage.Invoke() })



# =========================
# User Search Logic
# =========================
$btnSearchUser.Add_Click({
    $lstSearchResults.Items.Clear()
    $txtUserSummary.Text = ""
    $query = $txtSearch.Text.Trim()
    if (-not $query) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a search term.","Input Error")
        return
    }
    try {
        $users = Get-ADUser -Filter {SamAccountName -like "*$query*" -or GivenName -like "*$query*" -or Surname -like "*$query*"} -Properties *
        if ($users) {
            foreach ($user in $users) {
                $lstSearchResults.Items.Add("{0} ({1})" -f $user.Name, $user.SamAccountName)
            }
        } else {
            $lstSearchResults.Items.Add("No users found.")
        }
    } catch {
        $lstSearchResults.Items.Add("Error searching AD: $($_.Exception.Message)")
    }
})
$lstSearchResults.Add_SelectedIndexChanged({
    $txtUserSummary.Text = ""
    if ($lstSearchResults.SelectedItem -and -not ($lstSearchResults.SelectedItem -like "No users found.*")) {
        $selected = $lstSearchResults.SelectedItem
        $sam = $selected -replace ".*\(([^)]+)\)", '$1'
        try {
            $user = Get-ADUser -Identity $sam -Properties *
            $groups = (Get-ADPrincipalGroupMembership -Identity $sam | Select-Object -ExpandProperty Name) -join ", "
            $summary = "Name: $($user.Name)`r`nUsername: $($user.SamAccountName)`r`nEmail: $($user.EmailAddress)`r`nTitle: $($user.Title)`r`nDepartment: $($user.Department)`r`nCompany: $($user.Company)`r`nOffice: $($user.Office)`r`nGroups: $groups"
            $txtUserSummary.Text = $summary
        } catch {
            $txtUserSummary.Text = "Error retrieving user details: $($_.Exception.Message)"
        }
    }
})


# =========================
# License Management Logic
# =========================
$lstLicenseUserResults.Add_SelectedIndexChanged({
    $txtCurrentLicenses.Text = ""
    $clbAvailableLicenses.Items.Clear()
    if ($lstLicenseUserResults.SelectedItem -and -not ($lstLicenseUserResults.SelectedItem -like "No users found.*")) {
        $selected = $lstLicenseUserResults.SelectedItem
        $upn = $selected -replace ".*\(([^)]+)\)", '$1'
        try {
            $userLicenses = Get-EntraUserLicenseDetail -UserId $upn
            $txtCurrentLicenses.Text = ($userLicenses | ForEach-Object { $_.SkuPartNumber }) -join ", "
            $allLicenses = Get-EntraSubscribedSku
            foreach ($sku in $allLicenses) {
                $checked = $userLicenses.SkuId -contains $sku.SkuId
                $idx = $clbAvailableLicenses.Items.Add("{0} ({1})" -f $sku.SkuPartNumber, $sku.SkuId)
                if ($checked) { $clbAvailableLicenses.SetItemChecked($idx, $true) }
            }
        } catch {
            $txtCurrentLicenses.Text = "Error retrieving licenses: $($_.Exception.Message)"
        }
    }
})
$btnAssignLicenses.Add_Click({
    if (-not $lstLicenseUserResults.SelectedItem) {
        [System.Windows.Forms.MessageBox]::Show("Select a user first.","Input Error")
        return
    }
    $upn = $lstLicenseUserResults.SelectedItem -replace ".*\(([^)]+)\)", '$1'
    $checkedLicenses = @()
    foreach ($item in $clbAvailableLicenses.CheckedItems) {
        $skuId = $item -replace ".*\(([^)]+)\)", '$1'
        $checkedLicenses += $skuId
    }
    try {
        $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
        $licenses.AddLicenses = $checkedLicenses
        Set-EntraUserLicense -UserId $upn -AssignedLicenses $licenses
        [System.Windows.Forms.MessageBox]::Show("Licenses updated!","Success")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error assigning licenses: $($_.Exception.Message)","Error")
    }
})

# =========================
# Read-Only/Selection-Only Controls
# =========================
# Make txtCurrentLicenses read-only and prevent typing or pasting
$txtCurrentLicenses.ReadOnly = $true
$txtCurrentLicenses.TabStop = $false
$txtCurrentLicenses.Add_KeyPress({ $_.Handled = $true })
$txtCurrentLicenses.Add_KeyDown({ $_.SuppressKeyPress = $true })
$txtCurrentLicenses.Add_KeyUp({ $_.SuppressKeyPress = $true })
$txtCurrentLicenses.Add_MouseDown({ $_.Handled = $true })

# Make lstLicenseUserResults selection-only (no typing)
$lstLicenseUserResults.SelectionMode = 'One'  # only allow single selection
$lstLicenseUserResults.TabStop = $false
$lstLicenseUserResults.Add_KeyPress({ $_.Handled = $true })
$lstLicenseUserResults.Add_KeyDown({
    # Only allow navigation keys (arrows, enter), block all others
    $allowed = @([System.Windows.Forms.Keys]::Up, [System.Windows.Forms.Keys]::Down, [System.Windows.Forms.Keys]::Return)
    if ($allowed -notcontains $_.KeyCode) { $_.SuppressKeyPress = $true }
})

# =========================
# Initial Page Load
# =========================
$ShowDashboard.Invoke()
[void]$form.ShowDialog()