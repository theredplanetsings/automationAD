<# ad-user, works with interface.ps1#>
function Create-ADUserFromForm {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.TextBox]$txtFirstName,
        [Parameter(Mandatory)]
        [System.Windows.Forms.TextBox]$txtLastName,
        [Parameter(Mandatory)]
        [System.Windows.Forms.TextBox]$txtUsername,
        [Parameter(Mandatory)]
        [System.Windows.Forms.ComboBox]$cmbOffice,
        [Parameter(Mandatory)]
        [System.Windows.Forms.ComboBox]$cmbCompany,
        [Parameter(Mandatory)]
        [System.Windows.Forms.ComboBox]$cmbState,
        [Parameter(Mandatory)]
        [System.Windows.Forms.ComboBox]$cmbCity,
        [Parameter(Mandatory)]
        [System.Windows.Forms.ComboBox]$cmbPostalCode,
        [Parameter(Mandatory)]
        [System.Windows.Forms.ComboBox]$cmbStreetAddress,
        [Parameter(Mandatory)]
        [System.Windows.Forms.ComboBox]$cmbDepartment,
        [Parameter(Mandatory)]
        [System.Windows.Forms.ComboBox]$cmbTitle,
        [Parameter(Mandatory)]
        [System.Windows.Forms.TextBox]$txtTelephone,
        [Parameter(Mandatory)]
        [ScriptBlock]$ShowPage1,
        [Parameter(Mandatory)]
        [System.Windows.Forms.Form]$form
    )

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
        # creates the new user
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

        # hides the creation form
        $form.Hide()

        # shows summary form
        $summaryForm = New-Object System.Windows.Forms.Form
        $summaryForm.Text = "User Created Summary"
        $summaryForm.Size = New-Object System.Drawing.Size(400,500)
        $summaryForm.StartPosition = "CenterScreen"
        $summaryForm.FormBorderStyle = 'FixedDialog'
        $summaryForm.MaximizeBox = $false

        $txtSummary = New-Object System.Windows.Forms.TextBox
        $txtSummary.Multiline = $true
        $txtSummary.ReadOnly = $true
        $txtSummary.Dock = 'Fill'
        $txtSummary.ScrollBars = 'Vertical'
        $txtSummary.Font = New-Object System.Drawing.Font("Consolas",10)

        $summaryText = "User Account Created:`r`n"
        $summaryText += "-------------------------`r`n"
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

        $btnClose = New-Object System.Windows.Forms.Button
        $btnClose.Text = "Close"
        $btnClose.Dock = 'Bottom'
        $btnClose.Add_Click({
            $txtFirstName.Text = ""
            $txtLastName.Text = ""
            $txtUsername.Text = ""
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
            $ShowPage1.Invoke()
            $form.Show()
        })
        $summaryForm.Controls.Add($btnClose)

        [void]$summaryForm.ShowDialog()
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Error")
    }
}

Export-ModuleMember -Function Create-ADUserFromForm