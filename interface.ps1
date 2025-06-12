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
    $username = ("{0}{1}" -f $firstName.Substring(0,1), $lastName).ToLower()
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

        # Optionally add to security groups here if needed:
        # Add-ADGroupMember -Identity "VPN Users" -Members $username

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
        
        # hides the creation form
        $form.Hide()
        
        # creates a new summary form to show the properties created/populated
        $summaryForm = New-Object System.Windows.Forms.Form
        $summaryForm.Text = "User Created Summary"
        $summaryForm.Size = New-Object System.Drawing.Size(400,500)
        $summaryForm.StartPosition = "CenterScreen"
        $summaryForm.FormBorderStyle = 'FixedDialog'
        $summaryForm.MaximizeBox = $false

        # creates a multi-line, read-only textbox to display the summary
        $txtSummary = New-Object System.Windows.Forms.TextBox
        $txtSummary.Multiline = $true
        $txtSummary.ReadOnly = $true
        $txtSummary.Dock = 'Fill'
        $txtSummary.ScrollBars = 'Vertical'
        $txtSummary.Font = New-Object System.Drawing.Font("Consolas",10)
        
        $summaryText = "User Account Created:" + "`r`n"
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

        $summaryForm.Controls.Add($txtSummary)

        # button to close summary form and return to main form
        $btnClose = New-Object System.Windows.Forms.Button
        $btnClose.Text = "Close"
        $btnClose.Dock = 'Bottom'
        $btnClose.Add_Click({ 
            # clears input fields for a new user entry
            $txtFirstName.Text = ""
            $txtLastName.Text = ""
            $cmbOffice.SelectedIndex = -1
            $cmbCompany.SelectedIndex = -1
            $cmbState.SelectedIndex = -1
            $cmbCity.SelectedIndex = -1
            $cmbPostalCode.SelectedIndex = -1
            $cmbStreetAddress.SelectedIndex = -1
            $cmbDepartment.SelectedIndex = -1
            $cmbTitle.SelectedIndex = -1
            $txtTelephone.Text = ""
            
            $summaryForm.Close() 
            $form.Show() 
        })
        $summaryForm.Controls.Add($btnClose)

        [void]$summaryForm.ShowDialog()
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Error")
    }
}

# initialises form and controls
$form = New-Object System.Windows.Forms.Form
$form.Text = "Automation AD - Create User"
$form.Size = New-Object System.Drawing.Size(800,600)
$form.MinimumSize = New-Object System.Drawing.Size(800,600)
$form.MaximumSize = New-Object System.Drawing.Size(800,600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# labels + textboxes for First and Last name
$lblFirstName = New-Object System.Windows.Forms.Label
$lblFirstName.Text = "First Name:"
$lblFirstName.Location = New-Object System.Drawing.Point(175,20)
$lblFirstName.AutoSize = $true
$form.Controls.Add($lblFirstName)

$txtFirstName = New-Object System.Windows.Forms.TextBox
$txtFirstName.Location = New-Object System.Drawing.Point(300,18)
$txtFirstName.Width = 200
$txtFirstName.ShortcutsEnabled = $true 
$form.Controls.Add($txtFirstName)

# enables Ctrl+A for First Name textbox
$txtFirstName.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $sender.SelectAll()
        $e.SuppressKeyPress = $true
    }
})

$lblLastName = New-Object System.Windows.Forms.Label
$lblLastName.Text = "Last Name:"
$lblLastName.Location = New-Object System.Drawing.Point(175,60)
$lblLastName.AutoSize = $true
$form.Controls.Add($lblLastName)

$txtLastName = New-Object System.Windows.Forms.TextBox
$txtLastName.Location = New-Object System.Drawing.Point(300,58)
$txtLastName.Width = 200
$txtLastName.ShortcutsEnabled = $true  
$form.Controls.Add($txtLastName)

# enables Ctrl+A for Last Name textbox
$txtLastName.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $sender.SelectAll()
        $e.SuppressKeyPress = $true
    }
})

# dropdown for office name
$lblOffice = New-Object System.Windows.Forms.Label
$lblOffice.Text = "Office:"
$lblOffice.Location = New-Object System.Drawing.Point(175,100)
$lblOffice.AutoSize = $true
$form.Controls.Add($lblOffice)

$cmbOffice = New-Object System.Windows.Forms.ComboBox
$cmbOffice.Location = New-Object System.Drawing.Point(300,98)
$cmbOffice.Width = 200
$cmbOffice.DropDownStyle = 'DropDownList'
$cmbOffice.Items.AddRange(@("Corporate", "Branch A", "Branch B"))
$form.Controls.Add($cmbOffice)

# dropdown for company
$lblCompany = New-Object System.Windows.Forms.Label
$lblCompany.Text = "Company:"
$lblCompany.Location = New-Object System.Drawing.Point(175,140)
$lblCompany.AutoSize = $true
$form.Controls.Add($lblCompany)

$cmbCompany = New-Object System.Windows.Forms.ComboBox
$cmbCompany.Location = New-Object System.Drawing.Point(300,138)
$cmbCompany.Width = 200
$cmbCompany.DropDownStyle = 'DropDownList'
$cmbCompany.Items.AddRange(@("Paradigm Development", "Paradigm Services","Paradigm Management","Paradigm Construction"))
$form.Controls.Add($cmbCompany)

# dropdown for State (st)
$lblState = New-Object System.Windows.Forms.Label
$lblState.Text = "State:"
$lblState.Location = New-Object System.Drawing.Point(175,180)
$lblState.AutoSize = $true
$form.Controls.Add($lblState)

$cmbState = New-Object System.Windows.Forms.ComboBox
$cmbState.Location = New-Object System.Drawing.Point(300,178)
$cmbState.Width = 200
$cmbState.DropDownStyle = 'DropDownList'
$cmbState.Items.AddRange(@("VA", "MD", "DC"))
$form.Controls.Add($cmbState)

# dropdown for City
$lblCity = New-Object System.Windows.Forms.Label
$lblCity.Text = "City:"
$lblCity.Location = New-Object System.Drawing.Point(175,220)
$lblCity.AutoSize = $true
$form.Controls.Add($lblCity)

$cmbCity = New-Object System.Windows.Forms.ComboBox
$cmbCity.Location = New-Object System.Drawing.Point(300,218)
$cmbCity.Width = 200
$cmbCity.DropDownStyle = 'DropDownList'
$cmbCity.Items.AddRange(@("Arlington", "Washington D.C.", "Alexandria"))
$form.Controls.Add($cmbCity)

# dropdown for Postal Code
$lblPostalCode = New-Object System.Windows.Forms.Label
$lblPostalCode.Text = "Postal Code:"
$lblPostalCode.Location = New-Object System.Drawing.Point(175,260)
$lblPostalCode.AutoSize = $true
$form.Controls.Add($lblPostalCode)

$cmbPostalCode = New-Object System.Windows.Forms.ComboBox
$cmbPostalCode.Location = New-Object System.Drawing.Point(300,258)
$cmbPostalCode.Width = 200
$cmbPostalCode.DropDownStyle = 'DropDownList'
$cmbPostalCode.Items.AddRange(@("12345", "23456", "34567"))
$form.Controls.Add($cmbPostalCode)

# dropdown for Street Address
$lblStreet = New-Object System.Windows.Forms.Label
$lblStreet.Text = "Street Address:"
$lblStreet.Location = New-Object System.Drawing.Point(175,300)
$lblStreet.AutoSize = $true
$form.Controls.Add($lblStreet)

$cmbStreetAddress = New-Object System.Windows.Forms.ComboBox
$cmbStreetAddress.Location = New-Object System.Drawing.Point(300,298)
$cmbStreetAddress.Width = 200
$cmbStreetAddress.DropDownStyle = 'DropDownList'
$cmbStreetAddress.Items.AddRange(@("123 Main St", "456 Secondary St" , "789 Third St"))
$form.Controls.Add($cmbStreetAddress)

# dropdown for Department
$lblDepartment = New-Object System.Windows.Forms.Label
$lblDepartment.Text = "Department:"
$lblDepartment.Location = New-Object System.Drawing.Point(175,340)
$lblDepartment.AutoSize = $true
$form.Controls.Add($lblDepartment)

$cmbDepartment = New-Object System.Windows.Forms.ComboBox
$cmbDepartment.Location = New-Object System.Drawing.Point(300,338)
$cmbDepartment.Width = 200
$cmbDepartment.DropDownStyle = 'DropDownList'
$cmbDepartment.Items.AddRange(@("IT", "HR", "Finance"))
$form.Controls.Add($cmbDepartment)

# dropdown for Job Title
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Job Title:"
$lblTitle.Location = New-Object System.Drawing.Point(175,380)
$lblTitle.AutoSize = $true
$form.Controls.Add($lblTitle)

$cmbTitle = New-Object System.Windows.Forms.ComboBox
$cmbTitle.Location = New-Object System.Drawing.Point(300,378)
$cmbTitle.Width = 200
$cmbTitle.DropDownStyle = 'DropDownList'
$cmbTitle.Items.AddRange(@("IT Specialist", "Manager", "Analyst"))
$form.Controls.Add($cmbTitle)

# dropdown for Telephone Number (manual entry)
$lblTelephone = New-Object System.Windows.Forms.Label
$lblTelephone.Text = "Telephone Number:"
$lblTelephone.Location = New-Object System.Drawing.Point(175,420)
$lblTelephone.AutoSize = $true
$form.Controls.Add($lblTelephone)

$txtTelephone = New-Object System.Windows.Forms.TextBox
$txtTelephone.Location = New-Object System.Drawing.Point(300,418)
$txtTelephone.Width = 200
$txtTelephone.ShortcutsEnabled = $true 
$form.Controls.Add($txtTelephone)

# enables Ctrl+A for Telephone textbox
$txtTelephone.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $sender.SelectAll()
        $e.SuppressKeyPress = $true
    }
})

# button to submit and create the user
$btnSubmit = New-Object System.Windows.Forms.Button
$btnSubmit.Text = "Create User"
$btnSubmit.Location = New-Object System.Drawing.Point(500,450)
$btnSubmit.Width = 100
$btnSubmit.Add_Click({ Create-ADUserFromForm })
$form.Controls.Add($btnSubmit)

# initialises the main form
[void]$form.ShowDialog()
