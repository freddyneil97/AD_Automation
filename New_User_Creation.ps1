Import-Module ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms

#Declaring the Form and its Dimensions
$NewUserForm=New-Object System.Windows.Forms.Form
$NewUserForm.ClientSize='950,700'
$NewUserForm.Text='New User Creation'
$NewUserForm.BackColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
$NewUserForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle


#Labels and Textboxes
$labels = @{
        "First Name"     = 25
        "Last Name"      = 60
        "Full Name"      = 95
        "Username"       = 130
        "Display Name"   = 165
        "Email"          = 200
        "Model User"     = 270
        "Line Manager"   = 305
        "H Drive"        = 340
        "Object"         = 375
        "Department"     = 405
        "Job Title"      = 440
        "Company"        = 475
        "Office Location"= 510
        "Account Expiry" = 545
        "Ticket Number"  = 580

}

$textboxes = @{}

#Creating the labels and text boxes until display name

foreach ($label in $labels.Keys){

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "${label}:"
    $lbl.Size = New-Object System.Drawing.Point(100,25)
    $lbl.Font = New-Object System.Drawing.Font("Arial",10)
    $lbl.Location = New-Object System.Drawing.Point(25,$labels[$label])
    $NewUserForm.Controls.Add($lbl)

    if($label -in @("First Name","Last Name","Full Name","Username","Display Name","Department","Job Title","Company","Ticket Number","Office Location")){
        $text = New-Object System.Windows.Forms.TextBox
        $text.Size = New-Object System.Drawing.Point(300,25)
        $text.Font = New-Object System.Drawing.Font("Arial",10)
        $text.Location = New-Object System.Drawing.Point(150,$labels[$label])
        $NewUserForm.Controls.Add($text)
        $textboxes[$label] = $text
    }
    if($label -in "Email"){
        $Email = New-Object System.Windows.Forms.TextBox
        $Email.Size = New-Object System.Drawing.Point(150,25)
        $Email.Font = New-Object System.Drawing.Font("Arial",10)
        $Email.Location = New-Object System.Drawing.Point(150,$labels[$label])
        $NewUserForm.Controls.Add($Email)
        $textboxes["${label}_Textbox"] = $Email
        
        $EmailList = New-Object System.Windows.Forms.ComboBox
        $EmailList.Size = New-Object System.Drawing.Point(150,25)
        $EmailList.Location = New-Object System.Drawing.Point(305,$labels[$label])
        $EmailList.Items.AddRange(@("@option1.com","@option2.com","@option3.com"))
        $NewUserForm.Controls.Add($EmailList)
        $textboxes["${label}_ComboBox"] = $EmailList
    }
    if($label -in "Model User", "Line Manager"){
        $MM = New-Object System.Windows.Forms.TextBox
        $MM.Size = New-Object System.Drawing.Point(150,25)
        $MM.Font = New-Object System.Drawing.Font("Arial",10)
        $MM.Location = New-Object System.Drawing.Point(150,$labels[$label])
        $NewUserForm.Controls.Add($MM)
        $textboxes["${label}_Textbox"] = $MM

        $MMB = New-Object System.Windows.Forms.Button
        $MMB.Text = "Lookup"
        $MMB.Size = New-Object System.Drawing.Point(100,25)
        $MMB.Location = New-Object System.Drawing.Point(320,$labels[$label])
        $NewUserForm.Controls.Add($MMB)
        $textboxes["${label}_Button"] = $MMB
    }
    if($label -in "H Drive"){
        #Creating a groupbox for the radio buttons
        $HDriveGroupBox = New-Object System.Windows.Forms.GroupBox
        $HDriveGroupBox.Size = New-Object System.Drawing.Point(300,35)
        $HDriveGroupBox.Location = New-Object System.Drawing.Point(150,335)
        $NewUserForm.Controls.Add($HDriveGroupBox)

        $HD1 = New-Object System.Windows.Forms.RadioButton
        $HD1.Text = "Yes"
        $HD1.Location = New-Object System.Drawing.Point(5,7)   # relative to groupbox
        $HDriveGroupBox.Controls.Add($HD1)
        $textboxes["${label}_Button1"] = $HD1

        $HD2 = New-Object System.Windows.Forms.RadioButton
        $HD2.Text = "No"
        $HD2.Location = New-Object System.Drawing.Point(120,7)   # relative to groupbox
        $HD2.Checked = $true
        $HDriveGroupBox.Controls.Add($HD2)
        $textboxes["${label}_Button2"] = $HD2
    }
    if($label -in "Object"){
       $obj = New-Object System.Windows.Forms.Label
       $obj.Text = ""
       $obj.AutoSize = "True"
       $obj.Location = New-Object System.Drawing.Point(150,$labels[$label])
       $NewUserForm.Controls.Add($obj)     
       $textboxes["${label}"] = $obj
    }
    if($label -in "Account Expiry"){
        #Creating Groupbox for Radio Buttons and DateTimePicker
        $ExpGroupBox = New-Object System.Windows.Forms.GroupBox
        $ExpGroupBox.Size = New-Object System.Drawing.Point(300,35)
        $ExpGroupBox.Location = New-Object System.Drawing.Point(150,535)
        $NewUserForm.Controls.Add($ExpGroupBox)

        $AE = New-Object System.Windows.Forms.DateTimePicker
        $AE.Format = 'Short'
        $AE.Size = New-Object System.Drawing.Point(140,25)
        $AE.Location = New-Object System.Drawing.Point(155,9)   # relative to groupbox
        $AE.Enabled = $false
        $ExpGroupBox.Controls.Add($AE)
        $textboxes["${label}_DateTimePicker"] = $AE

        $Expbutton1 = New-Object System.Windows.Forms.RadioButton
        $Expbutton1.Text = "Yes"
        $Expbutton1.Location = New-Object System.Drawing.Point(5,7)
        $ExpGroupBox.Controls.Add($Expbutton1)
        $textboxes["${label}_RadioButton1"] = $Expbutton1

        $Expbutton2 = New-Object System.Windows.Forms.RadioButton
        $Expbutton2.Text = "No"
        $Expbutton2.Location = New-Object System.Drawing.Point(120,7)
        $Expbutton2.Checked = $true
        $ExpGroupBox.Controls.Add($Expbutton2)
        $textboxes["${label}_RadioButton2"] = $Expbutton2
    }
}

#Creating tabIndex for easier navigation
$textboxes["First Name"].TabIndex = 0
$textboxes["Last Name"].TabIndex = 1
$textboxes["Model User_Textbox"].TabIndex = 2
$textboxes["Line Manager_Textbox"].TabIndex = 3
$textboxes["Ticket Number"].TabIndex = 4

#Button for Searching existing users in Active Directory
$Search = New-Object System.Windows.Forms.Button
$Search.Text = "Lookup"
$Search.Size = New-Object System.Drawing.Point(100,25)
$Search.Location = New-Object System.Drawing.Point(250,235)
$NewUserForm.Controls.Add($Search)

#Submit Button to create the new user in Active Directory
$Submit = New-Object System.Windows.Forms.Button
$Submit.Text = "Submit"
$Submit.Size = New-Object System.Drawing.Point(100,25)
$Submit.Location = New-Object System.Drawing.Point(200,630)
$NewUserForm.Controls.Add($Submit)

#Output Window for Username, Email, Display Name - to be sent to Line Manager email
$OutUser = New-object System.Windows.Forms.Label
$OutUser.Text = "Username, Email, Display Name will be sent to Line Manager upon submission."
$OutUser.AutoSize = "True"
$OutUser.Location = New-Object System.Drawing.Point(500,15)
$NewUserForm.Controls.Add($OutUser)

#Subject to be sent to Line Manager email
$OutUserInfoS = New-object System.Windows.Forms.TextBox
$OutUserInfoS.Multiline = $true
$OutUserInfoS.Size = New-Object System.Drawing.Point(400,25)
$OutUserInfoS.Location = New-Object System.Drawing.Point(500,40)
$OutUserInfoS.ReadOnly = $true
$NewUserForm.Controls.Add($OutUserInfoS)

#Body to be sent to Line Manager email
$OutUserInfo = New-object System.Windows.Forms.TextBox
$OutUserInfo.Multiline = $true
$OutUserInfo.Size = New-Object System.Drawing.Point(400,200)
$OutUserInfo.Location = New-Object System.Drawing.Point(500,70)
$OutUserInfo.ReadOnly = $true
$NewUserForm.Controls.Add($OutUserInfo)


#Output Window for Password - to be sent to Line Manager email

$OutPass = New-object System.Windows.Forms.Label
$OutPass.Text = "Password will be sent to Line Manager upon submission."
$OutPass.AutoSize = "True"
$OutPass.Location = New-Object System.Drawing.Point(500,280)
$NewUserForm.Controls.Add($OutPass)

#Subject to be sent to Line Manager email
$OutUserPassS = New-object System.Windows.Forms.TextBox
$OutUserPassS.Multiline = $true
$OutUserPassS.Size = New-Object System.Drawing.Point(400,25)
$OutUserPassS.Location = New-Object System.Drawing.Point(500,305)
$OutUserPassS.ReadOnly = $true
$NewUserForm.Controls.Add($OutUserPassS)

$OutPassInfo = New-object System.Windows.Forms.TextBox
$OutPassInfo.Multiline = $true
$OutPassInfo.Size = New-Object System.Drawing.Point(400,200)
$OutPassInfo.Location = New-Object System.Drawing.Point(500,335)
$OutPassInfo.ReadOnly = $true
$NewUserForm.Controls.Add($OutPassInfo)

#Account Expiry Radio Button Logic
$enableExpiry = {
    if ($textboxes["Account Expiry_RadioButton1"].Checked) {
        $textboxes["Account Expiry_DateTimePicker"].Enabled = $true
    } else {
        $textboxes["Account Expiry_DateTimePicker"].Enabled = $false
    }
}
$textboxes["Account Expiry_RadioButton1"].Add_CheckedChanged($enableExpiry)
$textboxes["Account Expiry_RadioButton2"].Add_CheckedChanged($enableExpiry)


#Upon entering the First Name and Last Name, Full Name, Username, Display Name and Email will auto populate
$updatefields = {
    $FirstName = $textboxes["First Name"].Text
    $LastName = $textboxes["Last Name"].Text
    $textboxes["Full Name"].Text = "$FirstName.$LastName"
    $textboxes["Username"].Text = "$FirstName.$LastName"
    $textboxes["Display Name"].Text = "$FirstName $LastName" 
    $textboxes["Email_Textbox"].Text = "$FirstName.$LastName"
}

$textboxes["First Name"].Add_TextChanged($updatefields)
$textboxes["Last Name"].Add_TextChanged($updatefields)

#Upon clicking the Search button, it will search Active Directory if there is an existing user with the same name
$searchuser = {
    $Username = $textboxes["Username"].Text
    $ADUser = Get-ADUser -Filter {SamAccountName -eq $Username} -ErrorAction SilentlyContinue
    $AltUser = $textboxes["First Name"].substring(0,1) + $textboxes["Last Name"].Text
    $ADAltUser = Get-ADUser -Filter {SamAccountName -eq $AltUser} -ErrorAction SilentlyContinue
    if ($ADUser -or $ADAltUser) {
        [System.Windows.Forms.MessageBox]::Show("A user with the username '$Username' already exists in Active Directory.","User Exists",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
    } else {
        [System.Windows.Forms.MessageBox]::Show("No existing user found with the username '$Username'.","User Not Found",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    }
}

$Search.Add_Click($searchuser)

#Upon clicking Enter button next to Model User, it will copy the details from the Model User to the respective fields
$copyfrommodel = {
    $ModelUser = $textboxes["Model User_Textbox"].Text
    $ADModelUser = Get-ADUser -Identity $ModelUser -Properties * -ErrorAction SilentlyContinue
    if ($ADModelUser.Manager) {
        # Manager is a DN; get manager object safely
        $managerObj = Get-ADUser -Identity $ADModelUser.Manager -Properties SamAccountName, GivenName, Surname -ErrorAction SilentlyContinue
        if ($managerObj) {
            $textboxes["Line Manager_Textbox"].Text = $managerObj.SamAccountName
            # or use $managerObj.Name for display name
        }
    }

    if ($ADModelUser) {
        #$textboxes["Line Manager"].Text = $ModelUserMgr
        $textboxes["Department"].Text = $ADModelUser.Department
        $textboxes["Job Title"].Text = $ADModelUser.Title
        $textboxes["Company"].Text = $ADModelUser.Company
        $textboxes["Object"].Text = ($ADModelUser.DistinguishedName -split ",", 2)[1]
        $textboxes["Office Location"].Text = $ADModelUser.Office

        if ($ADModelUser.HomeDrive -and $ADModelUser.HomeDirectory ) {
            <# Action to perform if the condition is true #>
            $textboxes["H Drive_Button1"].Checked = $true
        }
        else {
            $textboxes["H Drive_Button2"].Checked = $true
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Model user '$ModelUser' not found in Active Directory.","Model User Not Found",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }
}
$textboxes["Model User_Button"].Add_Click($copyfrommodel)

#Upon clicking the Enter button next to Line Manager, it will search and validate the Line Manager exists in Active Directory
$validatelineManager = {
    $LineManager = $textboxes["Line Manager_Textbox"].Text
    $ADLineManager = Get-ADUser -Identity $LineManager -ErrorAction SilentlyContinue
    if ($ADLineManager) {
        [System.Windows.Forms.MessageBox]::Show("Line Manager '$LineManager' found in Active Directory.","Line Manager Found",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        [System.Windows.Forms.MessageBox]::Show("Line Manager '$LineManager' not found in Active Directory.","Line Manager Not Found",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }
}
$textboxes["Line Manager_Button"].Add_Click($validatelineManager)

#Upon clicking the Submit button, it will first check all mandatory fields are filled, then create the new user in Active Directory
$submituser = {
    $mandatoryFields = @("First Name", "Last Name", "Username", "Display Name", "Email_Textbox", "Ticket Number")
    foreach ($field in $mandatoryFields) {
        if ([string]::IsNullOrWhiteSpace($textboxes[$field].Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please fill in the mandatory field: '$field'.","Missing Information",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
    }

#Password generation with complexity requirements
    $passwordLength = 16
    $lowercase = 1
    $uppercase = 1
    $digits = 1
    $special = 1
    $allChars = @()
    $allChars += [char[]]'abcdefghijklmnopqrstuvwxyz'
    $allChars += [char[]]'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    $allChars += [char[]]'0123456789'
    $allChars += [char[]]'!@#$%^&*()-_=+[]{}|;:,.<>?'
    $passwordChars = @()
    $passwordChars += (Get-Random -InputObject ([char[]]'abcdefghijklmnopqrstuvwxyz') -Count $lowercase)
    $passwordChars += (Get-Random -InputObject ([char[]]'ABCDEFGHIJKLMNOPQRSTUVWXYZ') -Count $uppercase)
    $passwordChars += (Get-Random -InputObject ([char[]]'0123456789') -Count $digits)
    $passwordChars += (Get-Random -InputObject ([char[]]'!@#$%^&*()-_=+[]{}|;:,.<>?') -Count $special)
    $remainingLength = $passwordLength - $passwordChars.Count
    $passwordChars += (Get-Random -InputObject $allChars -Count $remainingLength)
    $password = ($passwordChars | Sort-Object {Get-Random}) -join ''

    #H Drvive assignment
    if ($textboxes["H Drive_Button1"].Checked) {
        $HDrive = "H:"
        $HDirectory = Get-ADUser -Identity $textboxes["Model User_Textbox"].Text -Properties HomeDirectory | Select-Object -ExpandProperty HomeDirectory
        if ($HDirectory -match "(.*Home Folder\\)") {
            $ExtractedPath = $Matches[1]
            $HDirectory = $ExtractedPath + $textboxes["Username"].Text + "\Home"
            New-Item -Path $HDirectory -ItemType Directory
        }
        else {
            $HDirectory = "\\hydra-dc\hackme\IT\Home Folder\" + $textboxes["Username"].Text + "\Home"
            New-Item -Path $HDirectory -ItemType Directory
        }
    } else {
        $HDrive = $null
        $HDirectory = $null
    }

    # Create the new user in Active Directory
    try {
        $newUserParams = @{
            GivenName             = $textboxes["First Name"].Text
            Surname               = $textboxes["Last Name"].Text
            Name                  = $textboxes["Full Name"].Text
            DisplayName           = $textboxes["Display Name"].Text
            SamAccountName        = $textboxes["Username"].Text
            UserPrincipalName     = $textboxes["Email_Textbox"].Text + $textboxes["Email_ComboBox"].SelectedItem
            Email                 = $textboxes["Email_Textbox"].Text + $textboxes["Email_ComboBox"].SelectedItem
            Department            = $textboxes["Department"].Text
            Title                 = $textboxes["Job Title"].Text
            Company               = $textboxes["Company"].Text
            Enabled               = $true
            Manager               = $textboxes["Line Manager_Textbox"].Text
            Description           = $textboxes["Last Name"].Text + ", " + $textboxes["First Name"].Text
            AccountPassword       = (ConvertTo-SecureString -String $password -AsPlainText -Force)
            HomeDrive             = $HDrive
            HomeDirectory         = $HDirectory
            Path                  = $textboxes["Object"].Text
            AccountExpirationDate = if ($textboxes["Account Expiry_RadioButton1"].Checked) { $textboxes["Account Expiry_DateTimePicker"].Value } else { $null }
            PasswordNeverExpires  = $true
            Office                = $textboxes["Office Location"].Text
        }

        New-ADUser @newUserParams

        $ModelUser = $textboxes["Model User_Textbox"].Text
        $ADModelUser = Get-ADUser -Identity $ModelUser -Properties * -ErrorAction SilentlyContinue

        if($ADModelUser.info){
            Set-ADUser $textboxes["Username"].Text -Replace @{info=$ADModelUser.info + '; Ticket Number: ' + $textboxes["Ticket Number"].Text}
        }
        if($ADModelUser.employeeType){
            Set-ADUser $textboxes["Username"].Text -Replace @{EmployeeType=$ADModelUser.EmployeeType} -ErrorAction SilentlyContinue
        }
        if($ADModelUser.middleName){
            Set-ADUser $textboxes["Username"].Text -Replace @{middleName=$ADModelUser.middleName} -ErrorAction SilentlyContinue
        }
        if($ADModelUser.extensionAttribute14){
            Set-ADUser $textboxes["Username"].Text -Replace @{extensionAttribute14=$ADModelUser.extensionAttribute14} -ErrorAction SilentlyContinue
        }
        if($ADModelUser.extensionAttribute15){
            Set-ADUser $textboxes["Username"].Text -Replace @{extensionAttribute15=$ADModelUser.extensionAttribute15} -ErrorAction SilentlyContinue
        }
        if($ADModelUser.msExchExtensionAttribute20){
            Set-ADUser $textboxes["Username"].Text -Replace @{msExchExtensionAttribute20=$ADModelUser.msExchExtensionAttribute20} -ErrorAction SilentlyContinue
        }
        if($ADModelUser.msExchExtensionAttribute21){
            Set-ADUser $textboxes["Username"].Text -Replace @{msExchExtensionAttribute21=$ADModelUser.msExchExtensionAttribute21} -ErrorAction SilentlyContinue
        }
        if($ADModelUser.ExtensionAttribute10){
            Set-ADUser $textboxes["Username"].Text -Replace @{ExtensionAttribute10=$ADModelUser.ExtensionAttribute10} -ErrorAction SilentlyContinue
        }
        [System.Windows.Forms.MessageBox]::Show("User '${newUserParams.SamAccountName}' created successfully in Active Directory.","Success",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error creating user: $_","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }

    #Copy thhe groups from Model User to the new user
    try {
        $ModelUserGroups = Get-ADUser -Identity $ModelUser -Properties MemberOf | Select-Object -ExpandProperty MemberOf

        $ExcludeGroups = @("Domain Admins",
                           "Enterprise Admins",
                           "Schema Admins",
                           "Administrators",
                           "Account Operators",
                           "Server Operators",
                           "Print Operators",
                           "Backup Operators",
                           "Replicator"
                           )
        foreach ($group in $ModelUserGroups) {
            $groupObj = Get-ADGroup -Identity $group
            if ($groupObj.Name -notin $ExcludeGroups) {
                Add-ADGroupMember -Identity $group -Members $textboxes["Username"].Text
            }
        }
    }
    catch {
        <#Do this if a terminating exception happens#>
        [System.Windows.Forms.MessageBox]::Show("Error copying groups from model user: $_","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }


#Output the Subject to the respective text boxes
$OutUserInfoS.Text = "$($textboxes["Ticket Number"].Text) Credentials for $($textboxes["Full Name"].Text)"
$OutUserPassS.Text = "$($textboxes["Ticket Number"].Text) Credentials for $($textboxes["Full Name"].Text)"
#Output the Username to the respective text boxes
$OutUserInfo.Text = "Hi $($textboxes["Line Manager_Textbox"].Text),`r`n`r`nThe new user has been created with the following details:`r`n`r`nUsername: $($textboxes["Username"].Text)`r`nEmail: $($textboxes["Email_Textbox"].Text)$($textboxes["Email_ComboBox"].SelectedItem)`r`nDisplay Name: $($textboxes["Display Name"].Text)`r`n`r`nRegards,`r`nIT Team" 
# Corrected formatting for Password Info output 
$OutPassInfo.Text = "Hi $($textboxes["Line Manager_Textbox"].Text),`r`n`r`nThe password for the new user '$($textboxes["Username"].Text)' is:`r`n$password`r`n`r`nPlease ensure the user changes their password upon first login.`r`n`r`nRegards,`r`nIT Team"
}


$Submit.Add_Click($submituser)

#Displaying the Form
$NewUserForm.ShowDialog()

#Cleans up the Form
$NewUserForm.Dispose()
