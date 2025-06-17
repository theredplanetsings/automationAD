<#self-contained with functions to add users, search for existing users/display summary of their account, and handle interface#>
#imports
Get-Module -ListAvailable ActiveDirectory
Import-Module ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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

    # retrieve values from dropdowns and textboxes
    $physicalDeliveryOfficeName = $cmbOffice.SelectedItem
    $company = $cmbCompany.SelectedItem
    $st = $cmbState.SelectedItem
    $l = $cmbCity.SelectedItem
    $postalCode = $cmbPostalCode.SelectedItem
    $streetAddress = $cmbStreetAddress.SelectedItem
    $department = $cmbDepartment.SelectedItem
    $title = $cmbTitle.SelectedItem
    $telephoneNumber = $txtTelephone.Text.Trim()
    $mail = "$username@$domainname"
    $mailNickname = $username
    $proxyAddresses = "smtp:$mail"
    
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
            -Path "OU=Internal,OU=Users,OU=PDC-SERVICES,DC=paradigmcos,DC=local" `
            -Enabled $true

        # sets additional attributes
        Set-ADUser `
            -Identity $username `
            -Office $physicalDeliveryOfficeName `
            -OfficePhone $telephoneNumber `
            -EmailAddress $mail `
            -Replace @{
                streetAddress = $streetAddress
                postalCode = $postalCode
                telephoneNumber = $telephoneNumber
                mail = $mail
                mailNickname = $mailNickname
                proxyAddresses = $proxyAddresses
                title = $title
                department = $department
                company = $company
                l = $l
                st = $st
            }
        
        [System.Windows.Forms.MessageBox]::Show("User created successfully!","Success")
        $ShowPage1.Invoke()
        $form.Show()
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Error")
    }
}

# initialises form and controls
$form = New-Object System.Windows.Forms.Form
$form.Text = "Automation AD - Dashboard"
$form.Size = New-Object System.Drawing.Size(800,600)
$form.MinimumSize = New-Object System.Drawing.Size(800,600)
$form.MaximumSize = New-Object System.Drawing.Size(800,600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# CSV import/validation
## update as needed, determines the columns required in the CSV file to autopopulate dropdowns.
# must add additional dropdown sections to this program if you add more columns to the CSV.
$requiredColumns = @("Office","Company","State","City","PostalCode","StreetAddress","Department","Title")
$csvData = $null
$csvError = $false

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

## page 1 controls, default page
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
            $csvData = Import-Csv -Path $openFileDialog.FileName
            $csvColumns = $csvData | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
            $missing = $requiredColumns | Where-Object { $_ -notin $csvColumns }
            if ($missing.Count -gt 0) {
                [System.Windows.Forms.MessageBox]::Show("CSV missing columns: $($missing -join ', ')", "CSV Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $csvError = $true
            } else {
                $csvError = $false
                Populate-DropdownsFromCsv $csvData
                $cmbCompany.Items.Clear()
                $cmbCompany.Items.AddRange(($csvData | Select-Object -ExpandProperty Company -Unique))
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error reading CSV: $($_.Exception.Message)", "CSV Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $csvError = $true
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

# username auto-population code
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

# next button --> page 2
$btnNext = New-Object System.Windows.Forms.Button
$btnNext.Text = "Next"
$btnNext.Location = New-Object System.Drawing.Point(350,375)
$btnNext.Width = 100
$form.Controls.Add($btnNext)

# page 2, hidden by default
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

## page 3 controls, summary and create user
$txtSummary = New-Object System.Windows.Forms.TextBox
$txtSummary.Multiline = $true
$txtSummary.ReadOnly = $true
$txtSummary.Location = New-Object System.Drawing.Point(175,60)
$txtSummary.Size = New-Object System.Drawing.Size(425, 350)
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

# back button from add user page to dashboard
$btnBackToDashboardFromAdd = New-Object System.Windows.Forms.Button
$btnBackToDashboardFromAdd.Text = "Back to Dashboard"
$btnBackToDashboardFromAdd.Location = New-Object System.Drawing.Point(325,450)
$btnBackToDashboardFromAdd.Width = 150
$btnBackToDashboardFromAdd.Visible = $false
$btnBackToDashboardFromAdd.Add_Click({
    $ShowDashboard.Invoke()
})
$form.Controls.Add($btnBackToDashboardFromAdd)

## page switch logic
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
    $btnSubmit.Visible = $false
    $btnBack.Visible = $false

    $txtSummary.Visible = $false
    $btnCreateUser.Visible = $false
    $btnBack2.Visible = $false
}

$ShowPage2 = {
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
    $btnSubmit.Visible = $true
    $btnBack.Visible = $true

    $txtSummary.Visible = $false
    $btnCreateUser.Visible = $false
    $btnBack2.Visible = $false
}

function Update-Summary {
    $firstName = $txtFirstName.Text.Trim()
    $lastName = $txtLastName.Text.Trim()
    $fullName = "$firstName $lastName"
    $username = $txtUsername.Text.Trim()
    $domainname = "paradigmcos.com"
    $userPrincipalName = "$username@$domainname"
    $physicalDeliveryOfficeName = $cmbOffice.SelectedItem
    $company = $cmbCompany.SelectedItem
    $st = $cmbState.SelectedItem
    $l = $cmbCity.SelectedItem
    $postalCode = $cmbPostalCode.SelectedItem
    $streetAddress = $cmbStreetAddress.SelectedItem
    $department = $cmbDepartment.SelectedItem
    $title = $cmbTitle.SelectedItem
    $telephoneNumber = $txtTelephone.Text.Trim()
    $mail = "$username@$domainname"

    $summaryText = "User Account Summary:" + "`r`n"
    $summaryText += "-------------------------" + "`r`n"
    $summaryText += "Full Name:           $fullName`r`n"
    $summaryText += "Username:            $username`r`n"
    $summaryText += "UserPrincipalName:   $userPrincipalName`r`n"
    $summaryText += "Email:               $mail`r`n"
    $summaryText += "Office:              $physicalDeliveryOfficeName`r`n"
    $summaryText += "Company:             $company`r`n"
    $summaryText += "State:               $st`r`n"
    $summaryText += "City:                $l`r`n"
    $summaryText += "Postal Code:         $postalCode`r`n"
    $summaryText += "Street Address:      $streetAddress`r`n"
    $summaryText += "Department:          $department`r`n"
    $summaryText += "Job Title:           $title`r`n"
    $summaryText += "Telephone:           $telephoneNumber`r`n"
    $txtSummary.Text = $summaryText
}

$ShowPage3 = {
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
    $btnSubmit.Visible = $false
    $btnBack.Visible = $false

    $txtSummary.Visible = $true
    $btnCreateUser.Visible = $true
    $btnBack2.Visible = $true

    Update-Summary
}

#page 2 buttons
$btnSubmit = New-Object System.Windows.Forms.Button
$btnSubmit.Text = "Next"
$btnSubmit.Location = New-Object System.Drawing.Point(500,450)
$btnSubmit.Width = 100
$btnSubmit.Visible = $false
#page 2 next button logic
$btnSubmit.Add_Click({
    # validates page 2 fields
    if (-not $cmbOffice.SelectedItem -or -not $cmbDepartment.SelectedItem -or -not $cmbTitle.SelectedItem -or -not $cmbStreetAddress.SelectedItem -or -not $cmbCity.SelectedItem -or -not $cmbState.SelectedItem -or -not $cmbPostalCode.SelectedItem -or -not $txtTelephone.Text.Trim()) {
        [System.Windows.Forms.MessageBox]::Show("Please fill out all fields before continuing.","Input Error")
        return
    }
    $ShowPage3.Invoke()
})
$form.Controls.Add($btnSubmit)

# back button for page 2
$btnBack = New-Object System.Windows.Forms.Button
$btnBack.Text = "Back"
$btnBack.Location = New-Object System.Drawing.Point(200,450)
$btnBack.Width = 100
$btnBack.Visible = $false
$btnBack.Add_Click({
    $ShowPage1.Invoke()
})
$form.Controls.Add($btnBack)

# page 1 next button logic
$btnNext.Add_Click({
    # validates page 1 fields
    if (-not $txtFirstName.Text.Trim() -or -not $txtLastName.Text.Trim() -or -not $txtUsername.Text.Trim() -or -not $cmbCompany.SelectedItem) {
        [System.Windows.Forms.MessageBox]::Show("Please fill out all fields before continuing.","Input Error")
        return
    }
    $ShowPage2.Invoke()
})

## dashboard page
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

## user search/summary page controls
$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "Search for User (Username, First, or Last Name):"
$lblSearch.Location = New-Object System.Drawing.Point(200,100)
$lblSearch.AutoSize = $true
$lblSearch.Visible = $false
$form.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(200,130)
$txtSearch.Width = 250
$txtSearch.Visible = $false
$form.Controls.Add($txtSearch)

$btnSearchUser = New-Object System.Windows.Forms.Button
$btnSearchUser.Text = "Search"
$btnSearchUser.Location = New-Object System.Drawing.Point(470,128)
$btnSearchUser.Width = 80
$btnSearchUser.Visible = $false
$form.Controls.Add($btnSearchUser)

$lstSearchResults = New-Object System.Windows.Forms.ListBox
$lstSearchResults.Location = New-Object System.Drawing.Point(200,170)
$lstSearchResults.Size = New-Object System.Drawing.Size(350,150)
$lstSearchResults.Visible = $false
$form.Controls.Add($lstSearchResults)

$txtUserSummary = New-Object System.Windows.Forms.TextBox
$txtUserSummary.Multiline = $true
$txtUserSummary.ReadOnly = $true
$txtUserSummary.Location = New-Object System.Drawing.Point(200,340)
$txtUserSummary.Size = New-Object System.Drawing.Size(350,150)
$txtUserSummary.ScrollBars = 'Vertical'
$txtUserSummary.Font = New-Object System.Drawing.Font("Consolas",10)
$txtUserSummary.Visible = $false
$form.Controls.Add($txtUserSummary)

$btnBackToDashboard = New-Object System.Windows.Forms.Button
$btnBackToDashboard.Text = "Back to Dashboard"
$btnBackToDashboard.Location = New-Object System.Drawing.Point(300,510)
$btnBackToDashboard.Width = 150
$btnBackToDashboard.Visible = $false
$btnBackToDashboard.Add_Click({ $ShowDashboard.Invoke() })
$form.Controls.Add($btnBackToDashboard)

## licenses page controls
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

$clbAvailableLicenses.SelectionMode = 'MultiSimple'
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

## page switch logic
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

## user search logic
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

## licenses page functionality
$btnLicenseUserSearch.Add_Click({
    $lstLicenseUserResults.Items.Clear()
    $txtCurrentLicenses.Text = ""
    $clbAvailableLicenses.Items.Clear()
    $query = $txtLicenseSearch.Text.Trim()
    if (-not $query) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a search term.","Input Error")
        return
    }
    try {
        $users = Get-EntraUser -Filter "userPrincipalName eq '$query' or displayName eq '$query' or mail eq '$query'" -ErrorAction Stop
        if ($users) {
            foreach ($user in $users) {
                $lstLicenseUserResults.Items.Add("{0} ({1})" -f $user.DisplayName, $user.UserPrincipalName)
            }
        } else {
            $lstLicenseUserResults.Items.Add("No users found.")
        }
    } catch {
        $lstLicenseUserResults.Items.Add("Error searching Entra: $($_.Exception.Message)")
    }
})

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

## initialises
$ShowDashboard.Invoke()

[void]$form.ShowDialog()