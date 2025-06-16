<# interface, works with ad-user.ps1, cannot test without having permission to edit AD#>

#imports
Get-Module -ListAvailable ActiveDirectory
Import-Module ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
. "$PSScriptRoot\ad-user.ps1"

# initialises form and controls
$form = New-Object System.Windows.Forms.Form
$form.Text = "Automation AD - Create User"
$form.Size = New-Object System.Drawing.Size(800,600)
$form.MinimumSize = New-Object System.Drawing.Size(800,600)
$form.MaximumSize = New-Object System.Drawing.Size(800,600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# CSV import/validation
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
    
    $cmbOffice.Items.AddRange(($data | Select-Object -ExpandProperty Office -Unique))
    $cmbCompany.Items.AddRange(($data | Select-Object -ExpandProperty Company -Unique))
    $cmbState.Items.AddRange(($data | Select-Object -ExpandProperty State -Unique))
    $cmbCity.Items.AddRange(($data | Select-Object -ExpandProperty City -Unique))
    $cmbPostalCode.Items.AddRange(($data | Select-Object -ExpandProperty PostalCode -Unique))
    $cmbStreetAddress.Items.AddRange(($data | Select-Object -ExpandProperty StreetAddress -Unique))
    $cmbDepartment.Items.AddRange(($data | Select-Object -ExpandProperty Department -Unique))
    $cmbTitle.Items.AddRange(($data | Select-Object -ExpandProperty Title -Unique))
}

## page 1 controls, first page we see
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

# username (auto-populates, but editable)
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

# Company dropdown (populated from CSV)
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

# username auto-population logic
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

# next button to go to page 2
$btnNext = New-Object System.Windows.Forms.Button
$btnNext.Text = "Next"
$btnNext.Location = New-Object System.Drawing.Point(350,375)
$btnNext.Width = 100
$form.Controls.Add($btnNext)

## page 2 controls, hidden by default
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

# post Code
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

# telephone Number (manual entry)
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

# next button for page 2
$btnNext2 = New-Object System.Windows.Forms.Button
$btnNext2.Text = "Next"
$btnNext2.Location = New-Object System.Drawing.Point(500,450)
$btnNext2.Width = 100
$btnNext2.Visible = $false
$btnNext2.Add_Click({ $ShowPage3.Invoke() })
$form.Controls.Add($btnNext2)

# back button for page 2
$btnBack = New-Object System.Windows.Forms.Button
$btnBack.Text = "Back"
$btnBack.Location = New-Object System.Drawing.Point(200,450)
$btnBack.Width = 100
$btnBack.Visible = $false
$btnBack.Add_Click({ $ShowPage1.Invoke() })
$form.Controls.Add($btnBack)

## page 3 controls, hidden by default
# summary label
$lblSummary = New-Object System.Windows.Forms.Label
$lblSummary.Text = "User Details Summary:"
$lblSummary.Location = New-Object System.Drawing.Point(175,60)
$lblSummary.AutoSize = $true
$lblSummary.Visible = $false
$form.Controls.Add($lblSummary)

# summary textbox
$txtSummary = New-Object System.Windows.Forms.TextBox
$txtSummary.Multiline = $true
$txtSummary.ReadOnly = $true
$txtSummary.Location = New-Object System.Drawing.Point(175,90)
$txtSummary.Size = New-Object System.Drawing.Size(400,300)
$txtSummary.ScrollBars = 'Vertical'
$txtSummary.Font = New-Object System.Drawing.Font("Consolas",10)
$txtSummary.Visible = $false
$form.Controls.Add($txtSummary)

# create user button
$btnSubmit = New-Object System.Windows.Forms.Button
$btnSubmit.Text = "Create User"
$btnSubmit.Location = New-Object System.Drawing.Point(500,450)
$btnSubmit.Width = 100
$btnSubmit.Visible = $false
$btnSubmit.Add_Click({
    Create-ADUserFromForm `
        -txtFirstName $txtFirstName `
        -txtLastName $txtLastName `
        -txtUsername $txtUsername `
        -cmbOffice $cmbOffice `
        -cmbCompany $cmbCompany `
        -cmbState $cmbState `
        -cmbCity $cmbCity `
        -cmbPostalCode $cmbPostalCode `
        -cmbStreetAddress $cmbStreetAddress `
        -cmbDepartment $cmbDepartment `
        -cmbTitle $cmbTitle `
        -txtTelephone $txtTelephone `
        -ShowPage1 $ShowPage1 `
        -form $form
})
$form.Controls.Add($btnSubmit)

# back button for page 3
$btnBack3 = New-Object System.Windows.Forms.Button
$btnBack3.Text = "Back"
$btnBack3.Location = New-Object System.Drawing.Point(200,450)
$btnBack3.Width = 100
$btnBack3.Visible = $false
$btnBack3.Add_Click({ $ShowPage2.Invoke() })
$form.Controls.Add($btnBack3)

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
    $btnBack3.Visible = $false
    $btnNext2.Visible = $false
    $lblSummary.Visible = $false
    $txtSummary.Visible = $false
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
    $btnNext2.Visible = $true
    $btnBack.Visible = $true
    $lblSummary.Visible = $false
    $txtSummary.Visible = $false
    $btnSubmit.Visible = $false
    $btnBack3.Visible = $false
}

$ShowPage3 = {
    # hides page 1 and 2 controls
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
    $btnBack.Visible = $false
    $btnNext2.Visible = $false
    # shows summary and create user
    $lblSummary.Visible = $true
    $txtSummary.Visible = $true
    $btnSubmit.Visible = $true
    $btnBack3.Visible = $true
    # populates summary
    $txtSummary.Text = "Full Name: $($txtFirstName.Text.Trim()) $($txtLastName.Text.Trim())`r`n" +
        "Username: $($txtUsername.Text.Trim())`r`n" +
        "Company: $($cmbCompany.SelectedItem)`r`n" +
        "Office: $($cmbOffice.SelectedItem)`r`n" +
        "Department: $($cmbDepartment.SelectedItem)`r`n" +
        "Job Title: $($cmbTitle.SelectedItem)`r`n" +
        "Street Address: $($cmbStreetAddress.SelectedItem)`r`n" +
        "City: $($cmbCity.SelectedItem)`r`n" +
        "State: $($cmbState.SelectedItem)`r`n" +
        "Postal Code: $($cmbPostalCode.SelectedItem)`r`n" +
        "Telephone: $($txtTelephone.Text.Trim())"
}

$btnNext.Add_Click({
    # validates page 1 fields
    if (-not $txtFirstName.Text.Trim() -or -not $txtLastName.Text.Trim() -or -not $txtUsername.Text.Trim() -or -not $cmbCompany.SelectedItem) {
        [System.Windows.Forms.MessageBox]::Show("Please fill out all fields before continuing.","Input Error")
        return
    }
    $ShowPage2.Invoke()
})

# shows only page 1 controls at startup
$ShowPage1.Invoke()

# initialises the main form
[void]$form.ShowDialog()