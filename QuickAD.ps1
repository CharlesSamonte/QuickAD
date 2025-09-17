######################################################################
# Author: Charles Samonte
# Something to help me with AD account management
######################################################################

Add-Type -AssemblyName PresentationCore, PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

############################################################################
#HELPER FUNCTIONS START HERE
function userNameExist($queryName) {
    $queryName = $queryName.Trim()
    if ((Get-ADUser -Filter "Name -eq '$queryName'").Count -eq 0) {
        #No user in AD
        $queryName = $queryName.Replace(' ', '.')
        if ((Get-ADUser -Filter "Name -eq '$queryName'").Count -eq 0) {
            return $false
        }
    }
    return $true
}

function GetUserCount($query) {
    $query = $query.Trim();
    $queryResult = ([array](Get-ADUser -Filter "Name -like '*$query*'")).Count
    return $queryResult
}

function ConvertTo-String {
    param (
        [Parameter(Mandatory = $true)]
        $Value
    )
    if ($Value -eq $null) {
        return ""
    }
    elseif ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        return ($Value -join ", ")
    }
    else {
        return $Value.ToString()
    }
}

#Show user info with name
function GetUserWithName($queryName) {
    $queryName = $queryName.Trim()
    if ($queryName -eq "") {
        return $null
    }
    elseif (-not (userNameExist($queryName))) {
        #No user in AD
        return $null
    }

    $UserQuery = Get-ADUser -Filter "Name -eq '$queryName'" -Properties *

    #Get Properties
    #$DisplayName = $UserQuery | Select-Object -expand DisplayName
    $fName = $UserQuery | Select-Object -expand GivenName
    $lName = $UserQuery | Select-Object -expand Surname
    $employeeNo = $UserQuery | Select-Object -expand EmployeeNumber
    $employeeType = $UserQuery | Select-Object -expand extensionAttribute3
    $birthdate = $UserQuery | Select-Object -expand extensionAttribute4
    $grade = $UserQuery | Select-Object -expand extensionAttribute2
    $jobTitle = $UserQuery | Select-Object -expand Title
    $email = $UserQuery | Select-Object -expand UserPrincipalName
    $logonName = $UserQuery | Select-Object -expand sAMAccountName

    $comp = $UserQuery | Select-Object -expand Company
    $dept = $UserQuery | Select-Object -expand Department
    $desc = $UserQuery | Select-Object -expand Description
    $loc = $UserQuery | Select-Object -expand CanonicalName
    $lastLogonDate = $UserQuery | Select-Object -expand LastLogonDate

    if ($null -eq $lastLogonDate) {
        $lastLogonDate = "N/A"
    }
    $accountStatus = $UserQuery | Select-Object -expand enabled

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "User Information"
    $form.Size = New-Object System.Drawing.Size(600, 400)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Dock = [System.Windows.Forms.DockStyle]::Top
    $dataGridView.Height = 310
    $dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $dataGridView.RowHeadersVisible = $false  # Hide the row headers

    # Create columns
    $propertyColumn = $dataGridView.Columns.Add("Property", "Property")
    $dataGridView.Columns.Add("Value", "Value")

    # Set the width of the first column
    $dataGridView.Columns[$propertyColumn].Width = 5

    # Add rows
    $dataGridView.Rows.Add("Name", "$fName $lName")
    $dataGridView.Rows.Add("Logon Name", $logonName)
    $dataGridView.Rows.Add("Email", $email)
    $dataGridView.Rows.Add("Job Title", $jobTitle)
    if ($jobTitle -like "*Student*") {
        $dataGridView.Rows.Add("Grade", $grade)
    }
    else {
        $dataGridView.Rows.Add("Birth Date", $birthdate)
        $dataGridView.Rows.Add("Employee No", $employeeNo)
        $dataGridView.Rows.Add("Employee Level", $employeeType)
    }
    $dataGridView.Rows.Add("Description", $desc)
    $dataGridView.Rows.Add("Department", $dept)
    $dataGridView.Rows.Add("Location", $loc)
    $dataGridView.Rows.Add("Last Logon Date", $lastLogonDate)
    $dataGridView.Rows.Add("Enabled", $accountStatus)

    # Create a smaller button
    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Show All"
    $button.Size = New-Object System.Drawing.Size(100, 30)
    $button.Location = New-Object System.Drawing.Point(250, 315)

    # Add an event handler for the button click event
    $button.Add_Click({
            GetAllUserInfoWithName($queryName)
        })

    $form.Controls.Add($dataGridView)
    $form.Controls.Add($button)
    $form.ShowDialog()
    $form.TopMost = $true
}


function GetAllUserInfoWithName($queryName) {
    $queryName = $queryName.Trim()
    if ($queryName -eq "") {
        #Nothing entered
        [System.Windows.MessageBox]::Show("Nothing entered.") | Out-Null
        return $null
    }
    elseif ((Get-ADUser -Filter "Name -eq '$queryName'").Count -eq 0) {
        #No user in AD
        $queryName = $queryName.replace(' ', '.')
        if ((Get-ADUser -Filter "Name -eq '$queryName'").Count -eq 0) {
            [System.Windows.MessageBox]::Show("No User Found.") | Out-Null
            return $null
        }
    }

    # Get the AD user information
    $users = Get-ADUser -Filter "Name -eq '$queryName'" -Property *

    # Create a new form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "AD User Information"
    $form.Size = New-Object System.Drawing.Size(500, 600)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    # Create a DataGridView
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Dock = "Fill"
    $dataGridView.ReadOnly = $true
    $dataGridView.AllowUserToAddRows = $false
    $dataGridView.AllowUserToDeleteRows = $false
    $dataGridView.AutoSizeColumnsMode = "AllCells"
    $dataGridView.RowHeadersVisible = $false  # Hide the row headers

    # Create columns
    $dataGridView.Columns.Add("Property", "Property")
    $dataGridView.Columns.Add("Value", "Value")

    # Loop through each user and add their information to the DataGridView
    foreach ($user in $users) {
        foreach ($property in $user.PSObject.Properties) {
            if ($null -ne $property.Value) {
                $propertyValue = ConvertTo-String -Value $property.Value
            }
            else {
                $propertyValue = " "
            }
            $dataGridView.Rows.Add($property.Name, $propertyValue) | Out-Null
        }
        # Add a blank row to separate users
        $dataGridView.Rows.Add("", "") | Out-Null
    }

    # Add the DataGridView to the form
    $form.Controls.Add($dataGridView)

    # Show the form
    $form.ShowDialog()
    $form.TopMost = $true
}
############################################################################
#ADDING NEW USER FUNCTIONS START HERE
function AddNotCasual {
    #Spacing
    $elementRow = 10;
    $rowSpace = 30;
    $labelLength = 120;
    $inputColStartPos = $labelLength + $rowSpace;

    $windowSize = '300,280';

    #Create Form
    $Form = New-Object system.Windows.Forms.Form
    $Form.ClientSize = $windowSize
    $Form.text = 'Create a New User'
    $Form.StartPosition = 'CenterScreen'
    $Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

    #First Name Label
    $fNameLabel = New-Object System.Windows.Forms.label
    $fNameLabel.Location = New-Object System.Drawing.Size(7, $elementRow)
    $fNameLabel.Size = New-Object System.Drawing.Size($labelLength, 30)
    $fNameLabel.Text = "First Name:"
    $Form.Controls.Add($fNameLabel)
    #First Name TextBox
    $fNameTextbox = New-Object System.Windows.Forms.TextBox
    $fNameTextbox.Location = New-Object System.Drawing.Size($inputColStartPos, $elementRow)
    $fNameTextbox.Size = New-Object System.Drawing.Size(120, 20)
    $Form.Controls.Add($fNameTextbox)

    #Last Name Label
    $elementRow += $rowSpace
    $lNameLabel = New-Object System.Windows.Forms.label
    $lNameLabel.Location = New-Object System.Drawing.Size(7, $elementRow)
    $lNameLabel.Size = New-Object System.Drawing.Size($labelLength, 30)
    $lNameLabel.Text = "Last Name:"
    $Form.Controls.Add($lNameLabel)
    #Last Name TextBox
    $lNameTextbox = New-Object System.Windows.Forms.TextBox
    $lNameTextbox.Location = New-Object System.Drawing.Size($inputColStartPos, $elementRow)
    $lNameTextbox.Size = New-Object System.Drawing.Size(120, 20)
    $Form.Controls.Add($lNameTextbox)

    #DOB Label
    $elementRow += $rowSpace
    $dobLabel = New-Object System.Windows.Forms.label
    $dobLabel.Location = New-Object System.Drawing.Size(7, $elementRow)
    $dobLabel.Size = New-Object System.Drawing.Size($labelLength, 30)
    $dobLabel.Text = "DOB (yyyy-mm-dd):"
    $Form.Controls.Add($dobLabel)
    #DOB TextBox
    $dobTextbox = New-Object System.Windows.Forms.DateTimePicker
    $dobTextbox.Location = New-Object System.Drawing.Size($inputColStartPos, $elementRow)
    $dobTextbox.Size = New-Object System.Drawing.Size(120, 20)
    $dobTextbox.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
    $Form.Controls.Add($dobTextbox)

    #Employee Number Label
    $elementRow += $rowSpace
    $empNoLabel = New-Object System.Windows.Forms.label
    $empNoLabel.Location = New-Object System.Drawing.Size(7, $elementRow)
    $empNoLabel.Size = New-Object System.Drawing.Size($labelLength, 30)
    $empNoLabel.Text = "Employee No.:"
    $Form.Controls.Add($empNoLabel)
    #Employee Number TextBox
    $empNoTextbox = New-Object System.Windows.Forms.TextBox
    $empNoTextbox.Location = New-Object System.Drawing.Size($inputColStartPos, $elementRow)
    $empNoTextbox.Size = New-Object System.Drawing.Size(120, 20)
    $Form.Controls.Add($empNoTextbox)

    #Job Title Label
    $elementRow += $rowSpace
    $titleLabel = New-Object System.Windows.Forms.label
    $titleLabel.Location = New-Object System.Drawing.Size(7, $elementRow)
    $titleLabel.Size = New-Object System.Drawing.Size($labelLength, 30)
    $titleLabel.Text = "Job Title:"
    $Form.Controls.Add($titleLabel)
    #Add Dropdown list
    $jobList = New-Object system.Windows.Forms.ComboBox
    $jobList.DropDownStyle = [system.Windows.Forms.ComboBoxStyle]::DropDown
    $jobList.text = ""
    $jobList.Size = New-Object System.Drawing.Size(120, 20)
    $jobList.location = New-Object System.Drawing.Point($inputColStartPos, $elementRow)
    # Add the items in the dropdown list
    @('Teacher', 'Admin Assistant', 'Ed. Assistant', 'Caretaker', 'Bus Driver', 'Substitute Staff', 'Substitute Teacher') | ForEach-Object { [void] $jobList.Items.Add($_) }
    $jobList.SelectedIndex = -1
    $Form.Controls.Add($jobList)

    #School Label
    $elementRow += $rowSpace
    $schoolLabel = New-Object System.Windows.Forms.label
    $schoolLabel.Location = New-Object System.Drawing.Size(7, $elementRow)
    $schoolLabel.Size = New-Object System.Drawing.Size($labelLength, 30)
    $schoolLabel.Text = "Department:"
    $Form.Controls.Add($schoolLabel)
    #Add Dropdown list
    $schoolList = New-Object system.Windows.Forms.ComboBox
    $schoolList.DropDownStyle = [system.Windows.Forms.ComboBoxStyle]::DropDownList
    $schoolList.text = ""
    $schoolList.Size = New-Object System.Drawing.Size(120, 20)
    $schoolList.location = New-Object System.Drawing.Point($inputColStartPos, $elementRow)
    # Add the items in the dropdown list
    $Global:allSchools | ForEach-Object { [void] $schoolList.Items.Add($_) }
    $Global:allDeptInGsec | ForEach-Object { [void] $schoolList.Items.Add("GSEC - " + $_) }
    $schoolList.SelectedIndex = -1
    $Form.Controls.Add($schoolList)
    $schoolList.Add_SelectedIndexChanged{
        #Modify WhereInSchoolList with department change
        if ($schoolList.SelectedItem -like '*GSEC*') {
            $pathToOU = "OU=" + $schoolList.SelectedItem.Replace("GSEC - ", "") + ",OU=GSEC,OU=GSSD Network,DC=GSSD,DC=ADS"
            try {
                $OUinSchool = Get-ADOrganizationalUnit -Filter * -SearchBase "$pathToOU" -SearchScope OneLevel | Select-Object Name
                $arrayOfOUInSchool = $OUinSchool.Name
                $WhereInSchoolList.Items.Clear()
                $arrayOfOUInSchool | ForEach-Object { [void] $WhereInSchoolList.Items.Add($_) }
            }
            catch {
                
            }
        }
        else {
            $pathToOU = "OU=" + $schoolList.SelectedItem + ",OU=Schools,OU=GSSD Network,DC=GSSD,DC=ADS"
            $OUinSchool = Get-ADOrganizationalUnit -Filter * -SearchBase "$pathToOU" -SearchScope OneLevel | Select-Object Name
            $arrayOfOUInSchool = $OUinSchool.Name
            $WhereInSchoolList.Items.Clear()
            $arrayOfOUInSchool | ForEach-Object { [void] $WhereInSchoolList.Items.Add($_) }
        }
    }

    #WhereInSchool Label
    $elementRow += $rowSpace
    $WhereInSchoolLabel = New-Object System.Windows.Forms.label
    $WhereInSchoolLabel.Location = New-Object System.Drawing.Size(7, $elementRow)
    $WhereInSchoolLabel.Size = New-Object System.Drawing.Size($labelLength, 30)
    $WhereInSchoolLabel.Text = "OU Path:"
    $Form.Controls.Add($WhereInSchoolLabel)
    #Add Dropdown list
    $WhereInSchoolList = New-Object system.Windows.Forms.ComboBox
    $WhereInSchoolList.DropDownStyle = [system.Windows.Forms.ComboBoxStyle]::DropDownList
    $WhereInSchoolList.text = " " 
    $WhereInSchoolList.Size = New-Object System.Drawing.Size(120, 20)
    $WhereInSchoolList.location = New-Object System.Drawing.Point($inputColStartPos, $elementRow)
    # Add the items in the dropdown list
  
    $WhereInSchoolList.SelectedIndex = -1
    $Form.Controls.Add($WhereInSchoolList)

    #Employee Type Label
    $elementRow += $rowSpace
    $empTypeLabel = New-Object System.Windows.Forms.label
    $empTypeLabel.Location = New-Object System.Drawing.Size(7, $elementRow)
    $empTypeLabel.Size = New-Object System.Drawing.Size($labelLength, 30)
    $empTypeLabel.Text = "Employee Level:"
    $Form.Controls.Add($empTypeLabel)
    #Add Dropdown list
    $employeeTypeList = New-Object system.Windows.Forms.ComboBox
    $employeeTypeList.DropDownStyle = [system.Windows.Forms.ComboBoxStyle]::DropDownList
    $employeeTypeList.text = " " 
    $employeeTypeList.location = New-Object System.Drawing.Point($inputColStartPos, $elementRow)
    $employeeTypeList.Size = New-Object System.Drawing.Size(120, 20)
    $employeeTypeList.SelectedIndex = -1
    # Add the items in the dropdown list
    $employeeTypeList.Items.Add(1) | Out-Null
    $employeeTypeList.Items.Add(2) | Out-Null
    $employeeTypeList.Items.Add(3) | Out-Null
    $employeeTypeList.Items.Add(4) | Out-Null
    $employeeTypeList.Items.Add(5) | Out-Null
    $Form.Controls.Add($employeeTypeList)

    $elementRow += $rowSpace
    $BackButton = New-Object System.Windows.Forms.Button
    $BackButton.Location = New-Object System.Drawing.Point(20, $elementRow)
    $BackButton.Size = New-Object System.Drawing.Size(120, 23)
    $BackButton.Text = "Back"
    $BackButton.DialogResult = [System.Windows.Forms.DialogResult]::Abort
    $Form.Controls.Add($BackButton)

    $DoneButton = New-Object System.Windows.Forms.Button
    $DoneButton.Location = New-Object System.Drawing.Point(150, $elementRow)
    $DoneButton.Size = New-Object System.Drawing.Size(120, 23)
    $DoneButton.Text = "Done"
    $DoneButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Form.Controls.Add($DoneButton)
   
    $res = $Form.ShowDialog()

    if ($res -eq [System.Windows.Forms.DialogResult]::OK) {
        #Get variables
        $fName = ($fNameTextbox.Text).Trim()
        $lName = ($lNameTextbox.Text).Trim()
        $dob = $dobTextbox.Value.ToString("yyyy-MM-dd")
        $loginName = ($fName + "." + $lName).ToLower()
         
        $empNumber = ($empNoTextbox.Text).Trim()
        
        $fullName = $fName + " " + $lName
        $displayEmail = $fName + "." + $lName + "@gssd.ca"
        $loginEmail = $loginName + "@gssd.ca"
        $googleEmail = $loginName + "@gssdschools.ca"

        $dept = $schoolList.SelectedItem
        $desc = $dept + " " + $WhereInSchoolList.SelectedItem
        $jobTitle = $jobList.Text
        $employeeType = $employeeTypeList.SelectedItem
        $defaultPass = (Get-Date $dob).ToString("MMMddyyyy")

        #Find the Base Path
        if ($schoolList.SelectedItem -like "*GSEC*") {
            $dept = "GSEC"
            $desc = $schoolList.SelectedItem.Replace("GSEC - " , "") 
            $OUBasePath = $Global:gsecOUPath
            $OUBasePath = "OU=" + $desc + "," + $OUBasePath
            if ($WhereInSchoolList.SelectedItem -ne "" -and $null -ne $WhereInSchoolList.SelectedItem -and $WhereInSchoolList.SelectedItem -ne " ") {
                $OUBasePath = "OU=" + $WhereInSchoolList.SelectedItem + "," + $OUBasePath  
            }

        }
        else {
            $OUBasePath = $Global:schoolsOUPath
            $OUBasePath = "OU=" + $schoolList.SelectedItem + "," + $OUBasePath
            $OUBasePath = "OU=" + $WhereInSchoolList.SelectedItem + "," + $OUBasePath
        }

        if ($fName -eq '' -or $lName -eq '' -or $null -eq $OUBasePath -or $jobTitle -eq '' -or $null -eq $dept -or $employeeType -eq '') {
            $ButtonType = [System.Windows.Forms.MessageBoxButtons]::OK
            $MessageIcon = [System.Windows.Forms.MessageBoxIcon]::Information
            $MessageBody = "Please fill out the form before pressing done."
            $MessageTitle = "Error"
            [System.Windows.Forms.MessageBox]::Show($MessageBody, $MessageTitle, $ButtonType, $MessageIcon) | Out-Null
            # $Form.Dispose()
            # AddNotCasual
            # exit
            return
        }

        #Test to see if the user already exists
        if (userNameExist($fullName)) {
            [System.Windows.MessageBox]::Show("User already exists!") | Out-Null
            $Form.Dispose()
            AddNotCasual
            exit
        }


        #Add user
        try {
            New-ADUser -Name $fullName -Enabled $true -sAMAccountName $loginName -UserPrincipalName $loginEmail -AccountPassword(ConvertTo-SecureString $defaultPass -AsPlainText -Force) -ChangePasswordAtLogon $true -Path $OUBasePath -EmployeeNumber $empNumber -GivenName $fName -Surname $lName -DisplayName $fullName -Description $desc -EmailAddress $displayEmail -Title $jobTitle -Department $dept -Company $Global:company -OtherAttributes @{'extensionAttribute3' = $employeeType; 'extensionAttribute1' = $googleEmail; 'extensionAttribute4' = $dob }
            if ($null -ne ([ADSISearcher] "(sAMAccountName=$loginName)").FindOne()) {
                [System.Windows.MessageBox]::Show("Succesfully Added!") | Out-Null
            }
        }
        catch {
            Write-Host $_.Exception.Message;
            [System.Windows.MessageBox]::Show("There was an Error creating the user!" + $_.Exception.Message) | Out-Null
        }
        finally {
            $Form.Dispose()
            MainMenu
            exit
        }
    }
    elseif ($res -eq [System.Windows.Forms.DialogResult]::Abort) {
        $Form.Dispose()
        MainMenu 
    }
    exit
}

function AddUser {
    # Create a new form
    $Form1 = New-Object system.Windows.Forms.Form
    $Form1.ClientSize = '250,150'
    $Form1.text = 'Create a New User'
    $Form1.StartPosition = 'CenterScreen'
    $Form1.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

    #Sub Button
    $subButton = New-Object System.Windows.Forms.Button
    $subButton.Location = New-Object System.Drawing.Point(40, 20)
    $subButton.Size = New-Object System.Drawing.Size(75, 50)
    $subButton.Text = 'Sub/Casual'
    $subButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $Form1.Controls.Add($subButton)

    #None Sub Button
    $notSubButton = New-Object System.Windows.Forms.Button
    $notSubButton.Location = New-Object System.Drawing.Point(130, 20)
    $notSubButton.Size = New-Object System.Drawing.Size(75, 50)
    $notSubButton.Text = 'Full-Time'
    $notSubButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $Form1.Controls.Add($notSubButton)

    #Add Next button
    $BackButton = New-Object System.Windows.Forms.Button
    $BackButton.Location = New-Object System.Drawing.Point(60, 100)
    $BackButton.Size = New-Object System.Drawing.Size(130, 33)
    $BackButton.Text = "Back"
    $BackButton.DialogResult = [System.Windows.Forms.DialogResult]::Abort
    $Form1.Controls.Add($BackButton)

    # Display the form
    $SCResult = $Form1.ShowDialog()

    if ($SCResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        AddCasual
    }
    elseif ($SCResult -eq [System.Windows.Forms.DialogResult]::No) {
        AddNotCasual
    }
    elseif ($SCResult -eq [System.Windows.Forms.DialogResult]::Abort) {
        MainMenu
    }
    exit
}

############################################################################
#MARK ON LEAVE FUNCTIONS START HERE
function MarkOnLeave($queryName) {
    #Create the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Mark On Leave'
    $form.Size = New-Object System.Drawing.Size(300, 180)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

    $Name = Get-ADUser -Filter "Name -eq '$queryName'" -Properties * | Select-Object -expand Name
    $oldTitle = Get-ADUser -Filter "Name -eq '$queryName'" -Properties * | Select-Object -expand Title
    $initial = New-Object System.Windows.Forms.Label
    $initial.Location = New-Object System.Drawing.Point(10, 20)
    $initial.Size = New-Object System.Drawing.Size(280, 20)
    $initial.Text = 'Before: ' + $Name + " - " + $oldTitle
    $form.Controls.Add($initial)

    #Remove white spaces from start and end
    $oldTitle = $oldTitle.Trim() 

    #Set new title with postfix $Global:onLeaveMark
    $setTitleAs = "$oldTitle $Global:onLeaveMark"
    Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Set-ADUser $_ -Title ($setTitleAs) } 

    #Get new title and show user difference
    $newTitle = Get-ADUser -Filter "Name -eq '$queryName'" -Properties * | Select-Object -expand Title
    $new = New-Object System.Windows.Forms.Label
    $new.Location = New-Object System.Drawing.Point(10, 50)
    $new.Size = New-Object System.Drawing.Size(280, 20)
    $new.Text = 'After: ' + $Name + " - " + $newTitle
    $form.Controls.Add($new)

    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Location = New-Object System.Drawing.Point(90, 90)
    $backButton.Size = New-Object System.Drawing.Size(95, 30)
    $backButton.Text = 'Main Menu'
    $backButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($backButton)

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        MainMenu
    }
    exit
     
}

function RemoveOnLeave($queryName) {
    $oldTitle = Get-ADUser -Filter "Name -eq '$queryName'" -Properties * | Select-Object -expand Title
    if ($oldTitle -notlike '*(On Leave)*') {
        [System.Windows.MessageBox]::Show("This user is NOT on leave.") | Out-Null
        MainMenu
        exit
    }
    $newTitle = $oldTitle.Replace(" $Global:onLeaveMark", "")
    $newTitle = $newTitle.Trim()
    Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Set-ADUser $_ -Title ($newTitle) }

    #Show user
    $Name = Get-ADUser -Filter "Name -eq '$queryName'" -Properties * | Select-Object -expand Name
    #Create the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Mark On Leave'
    $form.Size = New-Object System.Drawing.Size(300, 180)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

    $initial = New-Object System.Windows.Forms.Label
    $initial.Location = New-Object System.Drawing.Point(10, 20)
    $initial.Size = New-Object System.Drawing.Size(280, 20)
    $initial.Text = 'Before: ' + $Name + " - " + $oldTitle
    $form.Controls.Add($initial)

    $showTitle = Get-ADUser -Filter "Name -eq '$queryName'" -Properties * | Select-Object -expand Title
    $new = New-Object System.Windows.Forms.Label
    $new.Location = New-Object System.Drawing.Point(10, 50)
    $new.Size = New-Object System.Drawing.Size(280, 20)
    $new.Text = 'After: ' + $Name + " - " + $showTitle
    $form.Controls.Add($new)

    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Location = New-Object System.Drawing.Point(90, 90)
    $backButton.Size = New-Object System.Drawing.Size(95, 30)
    $backButton.Text = 'Main Menu'
    $backButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($backButton)

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        MainMenu
    }
    exit
}

function LeaveMenu($queryName) {
    #Create the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Mark On Leave'
    $form.Size = New-Object System.Drawing.Size(300, 220)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

    #Full Name
    $fLabel = New-Object System.Windows.Forms.Label
    $fLabel.Location = New-Object System.Drawing.Point(90, 20)
    $fLabel.Size = New-Object System.Drawing.Size(280, 20)
    $fLabel.Text = 'Enter Full Name:'
    $form.Controls.Add($fLabel)
    $fTextBox = New-Object System.Windows.Forms.TextBox
    $fTextBox.Location = New-Object System.Drawing.Point(40, 40)
    $fTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $fTextBox.Text = $queryName
    $form.Controls.Add($fTextBox)

    $markButton = New-Object System.Windows.Forms.Button
    $markButton.Location = New-Object System.Drawing.Point(35, 70)
    $markButton.Size = New-Object System.Drawing.Size(100, 35)
    $markButton.Text = "Mark`nOn Leave"    
    $markButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $form.Controls.Add($markButton)

    $removeButton = New-Object System.Windows.Forms.Button
    $removeButton.Location = New-Object System.Drawing.Point(145, 70)
    $removeButton.Size = New-Object System.Drawing.Size(100, 35)
    $removeButton.Text = "Remove`nOn Leave"
    $removeButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $form.Controls.Add($removeButton)

    #Horizontal line
    $lineLabel = New-Object System.Windows.Forms.Label
    $lineLabel.Location = New-Object System.Drawing.Point(0, 120)
    $lineLabel.Text = " "
    $lineLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D;
    $lineLabel.AutoSize = $false
    $lineLabel.Height = 2
    $lineLabel.Width = 300
    $form.Controls.Add($lineLabel)

    #Button
    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Location = New-Object System.Drawing.Point(90, 135)
    $backButton.Size = New-Object System.Drawing.Size(100, 35)
    $backButton.Text = 'Cancel'
    $backButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($backButton)

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        #Adding On Leave
        $queryName = $fTextBox.Text
        $queryName = $queryName.Trim()
        #Test if user exists
        if ($queryName -eq "") {
            #Nothing entered
            [System.Windows.MessageBox]::Show("Nothing entered.") | Out-Null
            LeaveMenu
            exit
        }
        elseif (-not (userNameExist($queryName))) {
            #No user in AD
            [System.Windows.MessageBox]::Show("Could not find a user with that name.") | Out-Null
            LeaveMenu
            exit
        }
        $title = Get-ADUser -Filter "Name -eq '$queryName'" -Properties * | Select-Object -expand Title
        #Test if user is on leave
        if ($title -like '*(On Leave)*') {
            [System.Windows.MessageBox]::Show("This user is already on leave.") | Out-Null
            LeaveMenu
            exit
        }
        MarkOnLeave($queryName)

    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::No) {
        #Removing On LEave
        $queryName = $fTextBox.Text
        $queryName = $queryName.Trim()
        #Test if user exists
        if ($queryName -eq "") {
            #Nothing entered
            [System.Windows.MessageBox]::Show("Nothing entered.") | Out-Null
            LeaveMenu
            exit
        }
        elseif (-not (userNameExist($queryName))) {
            #No user in AD
            [System.Windows.MessageBox]::Show("No User with that name.") | Out-Null
            LeaveMenu
            exit
        }
        $title = Get-ADUser -Filter "Name -eq '$queryName'" -Properties * | Select-Object -expand Title
        if ($title -notlike '*(On Leave)*') {
            [System.Windows.MessageBox]::Show("This user is NOT on leave") | Out-Null
            LeaveMenu
            exit
        }
        RemoveOnLeave($queryName)
    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        MainMenu
    }
    exit
}

############################################################################
#DElETE FUNCTIONS START HERE
function MoveToDeletes($queryName) {
    $queryName = $queryName.Trim()
    $del = "DEL"

    #check if user is already in deletes
    $currPath = Get-ADUser -Filter "Name -eq '$queryName'" -Properties * | Select-Object -expand DistinguishedName
    $checkPath = "CN=" + $queryName + "," + $Global:delFolderOU
    if ($currPath -eq $checkPath) {
        [System.Windows.MessageBox]::Show("User is already in " + $Global:delFolderName) | Out-Null
        MainMenu
        exit
    }

    #Move user
    Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Set-ADUser $_ -Department ($del) }
    Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Set-ADUser $_ -Description ($null) }
    Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Set-ADUser $_ -Company ($null) }
    Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Set-ADUser $_ -Title ($null) }
    Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Set-ADUser $_ -Clear "extensionattribute3" }
    Get-ADUser -Filter "Name -like '$queryName'" | Set-ADUser -Replace @{msExchHideFromAddressLists = $true }
    Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Disable-ADAccount $_ } 
    Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Move-ADObject $_ -TargetPath "$Global:delFolderOU" } 

    #check is succesfull
    $newPath = Get-ADUser -Filter "Name -eq '$queryName'" -Properties * | Select-Object -expand DistinguishedName
    $checkPath = "CN=" + $queryName + "," + $Global:delFolderOU
    if ($newPath -ne $checkPath) {
        [System.Windows.MessageBox]::Show("There was an Error moving user to " + $Global:delFolderName) | Out-Null
    }
    else {
        [System.Windows.MessageBox]::Show("Succesfully Moved User to " + $Global:delFolderName) | Out-Null
    }

    MainMenu
}

function MoveToDeletesMenu($queryName) {
    #Create the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Move users to ' + $Global:delFolderName
    $form.Size = New-Object System.Drawing.Size(300, 180)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

    #Full Name
    $fLabel = New-Object System.Windows.Forms.Label
    $fLabel.Location = New-Object System.Drawing.Point(90, 20)
    $fLabel.Size = New-Object System.Drawing.Size(280, 20)
    $fLabel.Text = 'Enter Full Name:'
    $form.Controls.Add($fLabel)
    $fTextBox = New-Object System.Windows.Forms.TextBox
    $fTextBox.Location = New-Object System.Drawing.Point(40, 40)
    $fTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $fTextBox.Text = $queryName
    $form.Controls.Add($fTextBox)

    #Button
    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Location = New-Object System.Drawing.Point(35, 70)
    $backButton.Size = New-Object System.Drawing.Size(100, 35)
    $backButton.Text = 'Back'
    $backButton.DialogResult = [System.Windows.Forms.DialogResult]::Abort
    $form.Controls.Add($backButton)

    $delButton = New-Object System.Windows.Forms.Button
    $delButton.Location = New-Object System.Drawing.Point(145, 70)
    $delButton.Size = New-Object System.Drawing.Size(100, 35)
    $delButton.Text = "Move to Deletes"    
    $delButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $form.Controls.Add($delButton)

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        #Move to Deletes
        $queryName = $fTextBox.Text
        $queryName = $queryName.Trim()
        #Test if user exists
        if ($queryName -eq "") {
            #Nothing entered
            [System.Windows.MessageBox]::Show("Nothing entered.") | Out-Null
            MoveToDeletesMenu
            return
        }
        elseif ((Get-ADUser -Filter "Name -eq '$queryName'").Count -eq 0) {
            #No user in AD
            $queryName = $queryName.replace(' ', '.')
            if ((Get-ADUser -Filter "Name -eq '$queryName'").Count -eq 0) {
                [System.Windows.MessageBox]::Show("No User with that name.") | Out-Null
                MoveToDeletesMenu
                exit
            }
        }
        #User exists
        MoveToDeletes($queryName)

    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Abort) {
        MainMenu
    }
    exit
}

############################################################################
#RESET PASSWORD FUNCTIONS START HERE

function ResetToDefaultPass($queryName) {
    $queryName = $queryName.Trim()
    $dobString = (Get-ADUser -Filter "Name -like '$queryName'" -Properties extensionAttribute4).extensionAttribute4

    # Parse it as a DateTime object
    $dob = [datetime]::ParseExact($dobString, "yyyy-MM-dd", $null)

    # Create default password in "MMMddyyyy" format (e.g., "Sep141999")
    $defaultPass = $dob.ToString("MMMddyyyy")
    #change password
    Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Set-ADAccountPassword $_ -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $defaultPass -Force) }
    #unlock account
    Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Unlock-ADAccount $_ }
    #User must change password at next login
    Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Set-ADUser $_ -ChangePasswordAtLogon $true }
    [System.Windows.MessageBox]::Show("User's Password has been reset to $defaultPass") | Out-Null
    MainMenu
}

function ResetCustomPass($queryName) {
    $queryName = $queryName.Trim()
    #Create Form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Set Custom Password'
    $form.Size = New-Object System.Drawing.Size(350, 250)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.StartPosition = 'CenterScreen'

    #Caption Label
    $name = Get-ADUser -Filter "Name -like '$queryName'" -Properties * | Select-Object -expand Name
    $captionLabel = New-Object System.Windows.Forms.label
    $captionLabel.Location = New-Object System.Drawing.Size(7, 10)
    $captionLabel.Size = New-Object System.Drawing.Size(250, 25)
    $captionLabel.Text = "Set Custom Password for $name"
    $form.Controls.Add($captionLabel)

    #Enter Password Label
    $passLabel = New-Object System.Windows.Forms.label
    $passLabel.Location = New-Object System.Drawing.Size(7, 40)
    $passLabel.Size = New-Object System.Drawing.Size(120, 20)
    $passLabel.Text = "Enter New Password:"
    $form.Controls.Add($passLabel)
    #New Password TextBox
    $passTextbox = New-Object System.Windows.Forms.TextBox
    $passTextbox.Location = New-Object System.Drawing.Size(140, 40)
    $passTextbox.Size = New-Object System.Drawing.Size(180, 20)
    $form.Controls.Add($passTextbox)
    $passTextbox.Add_TextChanged{
        if ($passTextbox.Text -eq $repeatPassTextbox.Text) {
            $form.Controls.Add($doneButton)
        }
        else {
            $form.Controls.Remove($doneButton)
        }
    }

    #Repeat Password Label
    $repeatPassLabel = New-Object System.Windows.Forms.label
    $repeatPassLabel.Location = New-Object System.Drawing.Size(7, 70)
    $repeatPassLabel.Size = New-Object System.Drawing.Size(120, 20)
    $repeatPassLabel.Text = "Repeat Password:"
    $repeatPassLabel.AutoSize = $true
    $form.Controls.Add($repeatPassLabel)
    #Repeat Password TextBox
    $repeatPassTextbox = New-Object System.Windows.Forms.TextBox
    $repeatPassTextbox.Location = New-Object System.Drawing.Size(140, 70)
    $repeatPassTextbox.Size = New-Object System.Drawing.Size(180, 20)
    $form.Controls.Add($repeatPassTextbox)
    $repeatPassTextbox.Add_TextChanged{
        if ($passTextbox.Text -eq $repeatPassTextbox.Text) {
            $form.Controls.Add($doneButton)
        }
        else {
            $form.Controls.Remove($doneButton)
        }
    }

    $changePassCB = New-Object System.Windows.Forms.CheckBox
    $changePassCB.Location = New-Object System.Drawing.Size(80, 90)
    $changePassCB.Size = New-Object System.Drawing.Size(200, 50)
    $changePassCB.Checked = $false
    $changePassCB.Text = "Change password at next logon"
    $form.Controls.Add($changePassCB)

    #Done Button
    $doneButton = New-Object System.Windows.Forms.Button
    $doneButton.Location = New-Object System.Drawing.Point(100, 140)
    $doneButton.Size = New-Object System.Drawing.Size(100, 35)
    $doneButton.Text = 'Done'
    $doneButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    #$form.Controls.Add($doneButton)

    #Button
    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Location = New-Object System.Drawing.Point(220, 140)
    $backButton.Size = New-Object System.Drawing.Size(100, 35)
    $backButton.Text = 'Back'
    $backButton.DialogResult = [System.Windows.Forms.DialogResult]::Abort
    $form.Controls.Add($backButton)

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        #Set default password
        try {
            $newPass = $repeatPassTextbox.Text
            #Set custom password
            Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Set-ADAccountPassword $_ -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $newPass -Force) }
           
            #unlock account
            Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Unlock-ADAccount $_ }
           
            #want to change password at next login
            if ($changePassCB.Checked) {
                Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Set-ADUser $_ -ChangePasswordAtLogon $true } 
            }
            else {
                Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Set-ADUser $_ -ChangePasswordAtLogon $false }
            }
           
            [System.Windows.MessageBox]::Show("$name's Password has been set to $newPass") | Out-Null
            MainMenu
            exit
        }
        catch {
            [System.Windows.MessageBox]::Show("The password does not meet the length, complexity, or history requirement of the domain.", 'ERROR', 'OK', 'Error') | Out-Null
            ResetCustomPass($queryName)
            exit
        }
    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Abort) {
        MainMenu
    }
    exit
}

function ResetPasswordMenu($queryName) {
    #Create the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Reset Password'
    $form.Size = New-Object System.Drawing.Size(300, 200)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.StartPosition = 'CenterScreen'

    #Full Name
    $fLabel = New-Object System.Windows.Forms.Label
    $fLabel.Location = New-Object System.Drawing.Point(90, 20)
    $fLabel.Size = New-Object System.Drawing.Size(280, 20)
    $fLabel.Text = 'Enter Full Name:'
    $form.Controls.Add($fLabel)
    $fTextBox = New-Object System.Windows.Forms.TextBox
    $fTextBox.Location = New-Object System.Drawing.Point(40, 40)
    $fTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $fTextBox.Text = $queryName
    $form.Controls.Add($fTextBox)

    $defaultPassButton = New-Object System.Windows.Forms.Button
    $defaultPassButton.Location = New-Object System.Drawing.Point(35, 70)
    $defaultPassButton.Size = New-Object System.Drawing.Size(100, 35)
    $defaultPassButton.Text = "Staff `nDefault"    
    $defaultPassButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $form.Controls.Add($defaultPassButton)

    $customPassButton = New-Object System.Windows.Forms.Button
    $customPassButton.Location = New-Object System.Drawing.Point(145, 70)
    $customPassButton.Size = New-Object System.Drawing.Size(100, 35)
    $customPassButton.Text = "Custom"
    $customPassButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $form.Controls.Add($customPassButton)

    #Horizontal line
    $lineLabel = New-Object System.Windows.Forms.Label
    $lineLabel.Location = New-Object System.Drawing.Point(0, 110)
    $lineLabel.Text = " "
    $lineLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D;
    $lineLabel.AutoSize = $false
    $lineLabel.Height = 2
    $lineLabel.Width = 300
    $form.Controls.Add($lineLabel)

    #Button
    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Location = New-Object System.Drawing.Point(90, 115)
    $backButton.Size = New-Object System.Drawing.Size(100, 35)
    $backButton.Text = 'Back'
    $backButton.DialogResult = [System.Windows.Forms.DialogResult]::Abort
    $form.Controls.Add($backButton)


    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        #Set default password
        $queryName = $fTextBox.Text
        $queryName = $queryName.Trim()
        #Test if user exists
        if ($queryName -eq "") {
            #Nothing entered
            [System.Windows.MessageBox]::Show("Nothing entered.") | Out-Null
            ResetPasswordMenu
            exit
        }
        elseif ((Get-ADUser -Filter "Name -eq '$queryName'").Count -eq 0) {
            #No user in AD
            $queryName = $queryName.replace(' ', '.')
            if ((Get-ADUser -Filter "Name -eq '$queryName'").Count -eq 0) {
                [System.Windows.MessageBox]::Show("No User with that name.") | Out-Null
                ResetPasswordMenu
                exit
            }
        }

        #User exists, continue
        ResetToDefaultPass($queryName)
    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::No) {
        #Set default password
        $queryName = $fTextBox.Text
        #Test if user exists
        if ($queryName -eq "") {
            #Nothing entered
            [System.Windows.MessageBox]::Show("Nothing entered.") | Out-Null
            ResetPasswordMenu
            exit
        }
        elseif ((Get-ADUser -Filter "Name -eq '$queryName'").Count -eq 0) {
            #No user in AD
            $queryName = $queryName.replace(' ', '.')
            if ((Get-ADUser -Filter "Name -eq '$queryName'").Count -eq 0) {
                [System.Windows.MessageBox]::Show("No User with that name.") | Out-Null
                ResetPasswordMenu
                exit
            }
        }

        #User exists, continue
        ResetCustomPass($queryName)
    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Abort) {
        MainMenu
    }
    exit

}
############################################################################
#MOVE USER FUNCTIONS START HERE
function MoveTo($queryName, $department, $OUPath, $path) {
    $queryName = $queryName.Trim()
    # Create variables for replacing current ADUser Properties
    $description = "$department"
    if ($OUPath -ne "" -and $OUPath -ne " " -and $null -ne $OUPath ) {
        $description = "$department $OUPath"
    }
    #Student
    if ( $path -like "*OU=Students*") {
        $description = "$department Students"
    }

    $title = Get-ADUser -Filter "Name -eq '$queryName'" -Properties * | Select-Object -expand Title
    $empNo = Get-ADUser -Filter "Name -eq '$queryName'" -Properties * | Select-Object -expand EmployeeNumber
    $empType = Get-ADUser -Filter "Name -eq '$queryName'" -Properties * | Select-Object -expand extensionAttribute3
    
    #Show current properties and make some editable
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Check Properties'
    $mainFormSizeX = 310
    $form.Size = New-Object System.Drawing.Size($mainFormSizeX, 320)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.StartPosition = 'CenterScreen'

    #Name Label 
    $nameLabel = New-Object System.Windows.Forms.label
    $nameLabel.Location = New-Object System.Drawing.Size(10, 13)
    $nameLabel.Size = New-Object System.Drawing.Size(90, 20)
    $nameLabel.Text = "Name:"
    $form.Controls.Add($nameLabel)
    #Name textbox
    $nameText = New-Object system.Windows.Forms.TextBox
    $nameText.Size = New-Object System.Drawing.Size(180, 20)
    $nameText.location = New-Object System.Drawing.Point(100, 10)
    $nameText.Text = $queryName
    $nameText.ReadOnly = $true
    $form.Controls.Add($nameText)

    #title Label 
    $titleLabel = New-Object System.Windows.Forms.label
    $titleLabel.Location = New-Object System.Drawing.Size(10, 43)
    $titleLabel.Size = New-Object System.Drawing.Size(90, 20)
    $titleLabel.Text = "Job title:"
    $form.Controls.Add($titleLabel)
    #title textbox
    $titleText = New-Object system.Windows.Forms.TextBox
    $titleText.Size = New-Object System.Drawing.Size(180, 20)
    $titleText.location = New-Object System.Drawing.Point(100, 40)
    $titleText.Text = $jobTitle
    $form.Controls.Add($titleText)

    #Employee number Label 
    $employeeNumberLabel = New-Object System.Windows.Forms.label
    $employeeNumberLabel.Location = New-Object System.Drawing.Size(10, 73)
    $employeeNumberLabel.Size = New-Object System.Drawing.Size(90, 20)
    $employeeNumberLabel.Text = "Employee No.:"
    $form.Controls.Add($employeeNumberLabel)
    #Employee numbe textbox
    $employeeNumberText = New-Object system.Windows.Forms.TextBox
    $employeeNumberText.Size = New-Object System.Drawing.Size(180, 20)
    $employeeNumberText.location = New-Object System.Drawing.Point(100, 70)
    $employeeNumberText.Text = $empNo
    $form.Controls.Add($employeeNumberText)

    #Employee type Label 
    $employeeTypeLabel = New-Object System.Windows.Forms.label
    $employeeTypeLabel.Location = New-Object System.Drawing.Size(10, 103)
    $employeeTypeLabel.Size = New-Object System.Drawing.Size(90, 20)
    $employeeTypeLabel.Text = "Employee Level:"
    $form.Controls.Add($employeeTypeLabel)
    #Employee type textbox
    $employeeTypeText = New-Object system.Windows.Forms.TextBox
    $employeeTypeText.Size = New-Object System.Drawing.Size(180, 20)
    $employeeTypeText.location = New-Object System.Drawing.Point(100, 100)
    $employeeTypeText.Text = $empType
    $form.Controls.Add($employeeTypeText)

    $defaultComp = 'GSSD'
    #Company Label 
    $compLabel = New-Object System.Windows.Forms.label
    $compLabel.Location = New-Object System.Drawing.Size(10, 133)
    $compLabel.Size = New-Object System.Drawing.Size(90, 20)
    $compLabel.Text = "Company:"
    $form.Controls.Add($compLabel)
    #title textbox
    $compText = New-Object system.Windows.Forms.TextBox
    $compText.location = New-Object System.Drawing.Point(100, 130)
    $compText.Size = New-Object System.Drawing.Size(180, 20)
    $compText.ReadOnly = $true
    $compText.Text = $defaultComp
    $form.Controls.Add($compText)
    
    #department Label 
    $departmentLabel = New-Object System.Windows.Forms.label
    $departmentLabel.Location = New-Object System.Drawing.Size(10, 163)
    $departmentLabel.Size = New-Object System.Drawing.Size(90, 20)
    $departmentLabel.Text = "Department:"
    $form.Controls.Add($departmentLabel)
    #description textbox
    $departmentText = New-Object system.Windows.Forms.TextBox
    $departmentText.location = New-Object System.Drawing.Point(100, 160)
    $departmentText.Size = New-Object System.Drawing.Size(180, 20)
    $departmentText.Text = $department
    $form.Controls.Add($departmentText)

    #description Label 
    $descriptionLabel = New-Object System.Windows.Forms.label
    $descriptionLabel.Location = New-Object System.Drawing.Size(10, 193)
    $descriptionLabel.Size = New-Object System.Drawing.Size(90, 20)
    $descriptionLabel.Text = "Description:"
    $form.Controls.Add($descriptionLabel)
    #description textbox
    $descriptionText = New-Object system.Windows.Forms.TextBox
    $descriptionText.location = New-Object System.Drawing.Point(100, 190)
    $descriptionText.Size = New-Object System.Drawing.Size(180, 20)
    $descriptionText.Text = $description
    $form.Controls.Add($descriptionText)

    $NextButton.Size = New-Object System.Drawing.Size(90, 35)

    #Button
    $OkButton = New-Object System.Windows.Forms.Button
    $OkButton.location = New-Object System.Drawing.Point(190, 230)
    $OkButton.Size = New-Object System.Drawing.Size(90, 35)
    $OkButton.Text = 'OK'
    $OkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($OkButton)

    #Button
    $backButton = New-Object System.Windows.Forms.Button
    $backButton.location = New-Object System.Drawing.Point(10, 230)
    $backButton.Size = New-Object System.Drawing.Size(90, 35)
    $backButton.Text = 'Back'
    $backButton.DialogResult = [System.Windows.Forms.DialogResult]::Abort
    $form.Controls.Add($backButton)

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        #Get new variables
        $newJobTitle = $titleText.Text
        $newDepartment = $departmentText.Text
        $newDescription = $descriptionText.Text
        $newEmployeeNumber = $employeeNumberText.Text
        $newEmployeeType = $employeeTypeText.Text

        Try {
            # Change user properties
            if ($newJobTitle -ne '') {
                #Not Students
                Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Set-ADUser $_ -Title ($newJobTitle) }
            }
            Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Set-ADUser $_ -Department ($newDepartment) }
            Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Set-ADUser $_ -Description ($newDescription) }
            Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Set-ADUser $_ -Company ($defaultComp) }
            Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Set-ADUser $_ -EmployeeNumber ($newEmployeeNumber) }
            Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Set-ADUser $_ -replace @{"extensionattribute3" = "$newEmployeeType" } }
            Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Set-ADUser $_ -Clear msExchHideFromAddressLists }
            Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Unlock-ADAccount $_ }
            Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Enable-ADAccount $_ }

            # Move User
            Get-ADUser -Filter "Name -like '$queryName'" | ForEach-Object { Move-ADObject $_ -TargetPath "$path" } 
            
        }
        catch {
            [System.Windows.MessageBox]::Show("Uh oh. There was an error.") | Out-Null
            MoveUserMenu
            exit
        }

        [System.Windows.MessageBox]::Show("Successfully moved $queryName to $description") | Out-Null
        MainMenu
    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Abort) {
        MoveUser($queryName)
    }
    exit
}

function MoveUser($queryName) {

    #Spacing
    $marginX = 10;
    $marginY = 10;
    $elementRow = 10;
    $rowSpace = 30;
    $labelLength = 120;
    $inputColStartPos = $labelLength + $rowSpace;
        
    $queryName = $queryName.Trim()

    #Create Form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Move User'
    $mainFormSizeX = 350
    $form.Size = New-Object System.Drawing.Size($mainFormSizeX, 200)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.StartPosition = 'CenterScreen'

    #Caption Label
    $name = Get-ADUser -Filter "Name -eq '$queryName'" -Properties * | Select-Object -expand Name
    $captionLabel = New-Object System.Windows.Forms.label
    $captionLabel.Location = New-Object System.Drawing.Size($marginX, $elementRow)
    $captionLabel.Size = New-Object System.Drawing.Size(200, 25)
    $captionLabel.Text = "Moving $name to..."
    $form.Controls.Add($captionLabel)
   
    $Script:movingWhere = 'e'

    $isStudent = New-Object System.Windows.Forms.CheckBox
    $isStudent.Location = New-Object System.Drawing.Point(210, 6)
    $isStudent.Size = New-Object System.Drawing.Size(80, 25)
    $isStudent.ForeColor = 'Blue'
    $isStudent.Text = "Student?"
    $form.Controls.Add($isStudent)
    $isStudent.Add_CheckStateChanged{
        if ($isStudent.checked) {
            $deptList.Items.Clear();
            $specificOuList.Items.Clear();
            $specificOuList.SelectedIndex = -1;
            $Global:allSchools | ForEach-Object { [void] $deptList.Items.Add($_) }
        }
        else {
            $deptList.Items.Clear();
            $specificOuList.Items.Clear();
            $specificOuList.SelectedIndex = -1;
            $Global:allSchools | ForEach-Object { [void] $deptList.Items.Add($_) }
            $gsecDeptPrefix = "GSEC - "
            $Global:allDeptInGsec | ForEach-Object { [void] $deptList.Items.Add( $gsecDeptPrefix + $_) }
        }
    }

    $elementRow += $rowSpace;
    #List of Department
    $departmentLabel = New-Object System.Windows.Forms.label
    $departmentLabel.Location = New-Object System.Drawing.Size($marginX, $elementRow)
    $departmentLabel.Size = New-Object System.Drawing.Size($labelLength, 20)
    $departmentLabel.Text = "Choose Department:"
    $form.Controls.Add($departmentLabel);
    
    #Add Dropdown list
    $deptList = New-Object system.Windows.Forms.ComboBox
    $deptList.DropDownStyle = [system.Windows.Forms.ComboBoxStyle]::DropDown
    $deptList.Size = New-Object System.Drawing.Size(180, 20)
    $deptList.location = New-Object System.Drawing.Point($inputColStartPos, $elementRow)
    # Add the items in the dropdown list
    $deptList.SelectedIndex = -1
    $Global:allSchools | ForEach-Object { [void] $deptList.Items.Add($_) }
    $gsecDeptPrefix = "GSEC - "
    $Global:allDeptInGsec | ForEach-Object { [void] $deptList.Items.Add( $gsecDeptPrefix + $_) }
    $deptList.Add_SelectedIndexChanged{
        $specificOuList.SelectedIndex = -1
        if ($isStudent.checked) {
            $specificOuList.Items.Clear()
            $pathToOU = "OU=Students,OU=" + $deptList.SelectedItem + ",$Global:schoolsOUPath"
            $OUinSchool = Get-ADOrganizationalUnit -Filter * -SearchBase "$pathToOU" -SearchScope OneLevel | Select-Object Name
            if ($OUinSchool.Length -gt 0) {
                $arrayOfOU = $OUinSchool.Name
                $arrayOfOU | ForEach-Object { [void] $specificOuList.Items.Add($_) }
            }
        }
        else {
            if ($deptList.SelectedItem -like "*GSEC*") {
                $shortenedOU = $deptList.SelectedItem.Replace( $gsecDeptPrefix, "")
                $pathToOU = "OU=" + $shortenedOU + ",$Global:gsecOUPath"
            }
            else {
                $pathToOU = "OU=" + $deptList.SelectedItem + ",$Global:schoolsOUPath"
            }
            $specificOuList.Items.Clear()
            $OUinSchool = Get-ADOrganizationalUnit -Filter * -SearchBase "$pathToOU" -SearchScope OneLevel | Select-Object Name
            if ($OUinSchool.Length -gt 0) {
                $arrayOfOU = $OUinSchool.Name
                $arrayOfOU | ForEach-Object { [void] $specificOuList.Items.Add($_) }
            }
        }
    }

    $form.Controls.Add($deptList);

    $elementRow += $rowSpace;
    #List of Schools
    $specificOuLabel = New-Object System.Windows.Forms.label
    $specificOuLabel.Location = New-Object System.Drawing.Size($marginX, $elementRow)
    $specificOuLabel.Size = New-Object System.Drawing.Size($labelLength, 25)
    $specificOuLabel.Text = "Select OU:"
    $form.Controls.Add($specificOuLabel)
    #Add Dropdown list
    $specificOuList = New-Object system.Windows.Forms.ComboBox
    $specificOuList.DropDownStyle = [system.Windows.Forms.ComboBoxStyle]::DropDown
    $specificOuList.location = New-Object System.Drawing.Point($inputColStartPos, $elementRow)
    $specificOuList.Size = New-Object System.Drawing.Size(180, 20)
    $form.Controls.Add($specificOuList)
    
    #if student
    #List of Schools
    $whichGradeLabel = New-Object System.Windows.Forms.label
    $whichGradeLabel.Location = New-Object System.Drawing.Size(7, 143)
    $whichGradeLabel.Size = New-Object System.Drawing.Size(90, 25)
    $whichGradeLabel.Text = "Grade:"
    #Add Dropdown list
    $whichGradeList = New-Object system.Windows.Forms.ComboBox
    $whichGradeList.DropDownStyle = [system.Windows.Forms.ComboBoxStyle]::DropDown
    $whichGradeList.Size = New-Object System.Drawing.Size(180, 20)
    $whichGradeList.location = New-Object System.Drawing.Point(100, 140)
    # Add the items in the dropdown list

    #Button
    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Location = New-Object System.Drawing.Point(10, 120)
    $backButton.Size = New-Object System.Drawing.Size(90, 35)
    $backButton.Text = 'Back'
    $backButton.DialogResult = [System.Windows.Forms.DialogResult]::Abort
    $form.Controls.Add($backButton)
    
    #Button
    $NextButton = New-Object System.Windows.Forms.Button
    $NextButton.Location = New-Object System.Drawing.Point(100, 120)
    $NextButton.Size = New-Object System.Drawing.Size(90, 35)
    $NextButton.Text = 'Next'
    $NextButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($NextButton)

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        #Move User
        $department = $deptList.SelectedItem

        $basePath = $Global:schoolsOUPath
        if ($department -like "*GSEC*") {
            $department = $department.Replace( $gsecDeptPrefix, "")
            $basePath = $Global:gsecOUPath
            
        }
        $OUPath = $specificOuList.SelectedItem
        $path = "OU=$OUPath,OU=$department,$basePath"
        if ($OUPath -eq "" -or $OUPath -eq " " -or $null -eq $OUPath) {
            # No specific OU under department.
            $path = "OU=$department,$basePath"
        }

        if ($isStudent.checked) {
            $basePath = $Global:schoolsOUPath
            $OUPath = $specificOuList.SelectedItem
            $path = "OU=$OUPath,OU=Students,OU=$department,$basePath"
        }  

        MoveTo $queryName $department $OUPath $path
        # switch ($movingWhere) {
        #     'Different School' {
        #         $department = $deptList.SelectedItem
        #         $OUPath = $specificOuList.SelectedItem
        #         $path = "OU=$OUPath,OU=$department,OU=Schools,OU=GSSD Network,DC=GSSD,DC=ADS"

        #         MoveTo $queryName $department $OUPath $path 
        #     }
        #     'In GSEC' {
        #         $department = 'GSEC'
        #         $OUPath = $whereInGSECList.SelectedItem
        #         $path = "OU=$OUPath,OU=$department,OU=GSSD Network,DC=GSSD,DC=ADS"

        #         MoveTo $queryName $department $OUPath $path 
        #     }            
        #     'In Transportation' {
        #         $department = $whereInTransList.SelectedItem

        #         if ($department -eq "Spare Drivers") {
        #             #Spare drivers does not have any OU inside it.
        #             $path = "OU=$department,OU=Transportation,OU=Users,OU=Board Office,OU=GSSD Network,DC=GSSD,DC=ADS"
        #             $OUPath = ''
        #         }
        #         else {
        #             $OUPath = $whichGarageList.SelectedItem
        #             $path = "OU=$OUPath,OU=$department,OU=Transportation,OU=Users,OU=Board Office,OU=GSSD Network,DC=GSSD,DC=ADS"
        #         }
        #         MoveTo $queryName $department $OUPath $path 
        #     }
        #     'In Grade' {
        #         $department = $whichSchoolList.SelectedItem
        #         $OUPath = $whichGradeList.SelectedItem
        #         $path = "OU=$OUPath,OU=Students,OU=$department,OU=Schools,OU=GSSD Network,DC=GSSD,DC=ADS"

        #         MoveTo $queryName $department $OUPath $path 
        #     }
        # }

    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Abort) {
        MoveUserMenu
    }
    exit
}

function MoveUserMenu($queryName) {
    #Create the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Move User'
    $form.Size = New-Object System.Drawing.Size(300, 180)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.StartPosition = 'CenterScreen'

    #Full Name
    $fLabel = New-Object System.Windows.Forms.Label
    $fLabel.Location = New-Object System.Drawing.Point(90, 20)
    $fLabel.Size = New-Object System.Drawing.Size(280, 20)
    $fLabel.Text = 'Enter Full Name:'
    $form.Controls.Add($fLabel)
    $fTextBox = New-Object System.Windows.Forms.TextBox
    $fTextBox.Location = New-Object System.Drawing.Point(40, 40)
    $fTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $fTextBox.Text = $queryName
    $form.Controls.Add($fTextBox)

    #Button
    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Location = New-Object System.Drawing.Point(35, 70)
    $backButton.Size = New-Object System.Drawing.Size(100, 35)
    $backButton.Text = 'Back'
    $backButton.DialogResult = [System.Windows.Forms.DialogResult]::Abort
    $form.Controls.Add($backButton)

    $nextButton = New-Object System.Windows.Forms.Button
    $nextButton.Location = New-Object System.Drawing.Point(145, 70)
    $nextButton.Size = New-Object System.Drawing.Size(100, 35)
    $nextButton.Text = "Next"    
    $nextButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $form.Controls.Add($nextButton)

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        #Set default password
        $queryName = $fTextBox.Text
        $queryName = $queryName.Trim()

        #Test if user exists
        if ($queryName -eq "") {
            #Nothing entered
            [System.Windows.MessageBox]::Show("Nothing entered.") | Out-Null
            MoveUserMenu
            exit
        }
        elseif (-not (userNameExist($queryName))) {
            #No user in AD
            [System.Windows.MessageBox]::Show("No User with that name.") | Out-Null
            MoveUserMenu
            exit
        }

        #User exists, continue
        MoveUser($queryName)
        MainMenu
    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Abort) {
        MainMenu
    }
    exit

}
############################################################################
#PROGRAM MAIN MENU
function MainMenu {
    #Create the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Main Menu'
    $form.Size = New-Object System.Drawing.Size(540, 280)
    $form.StartPosition = 'CenterScreen'
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
    #$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

    #Find user stuff
    $findLabel = New-Object System.Windows.Forms.Label
    $findLabel.Location = New-Object System.Drawing.Point(12, 18)
    $findLabel.Size = New-Object System.Drawing.Size(250, 20)
    $findLabel.Font = 'Microsoft Sans Serif,11'
    $findLabel.Text = 'Enter Name:'
    $form.Controls.Add($findLabel)

    $findTextBox = New-Object System.Windows.Forms.TextBox
    $findTextBox.Location = New-Object System.Drawing.Point(15, 40)
    $findTextBox.Size = New-Object System.Drawing.Size(185, 20)
    $form.Controls.Add($findTextBox)

    $getUserDetail = New-Object System.Windows.Forms.Button
    $getUserDetail.Location = New-Object System.Drawing.Point(205, 40)
    $getUserDetail.Size = New-Object System.Drawing.Size(65, 18)
    $getUserDetail.Font = 'Microsoft Sans Serif,10'
    $getUserDetail.Text = 'Search'
    $getUserDetail.Add_Click{
        if ($findTextBox.Text -ne "" -and $findTextBox.TextLength -gt 2) {
            $queryName = $findTextBox.Text
            $ShowNames.Items.Clear()
            $ShowNames.Items.AddRange("Looking for $queryName...")
            Start-Sleep -Seconds 0.1
            
            $userExist = userNameExist($queryName)
            if ($userExist) {
                GetUserWithName($queryName)
                $ShowNames.Items.Clear()
                return  
            } 
            
            $ShowNames.Items.AddRange("Could not find user with the name.")
            $ShowNames.Items.AddRange("Searching possible matches...")
            Start-Sleep -Seconds 0.2
            #Count number of users
            $numberOfUsers = GetUserCount($queryName)

            if ($numberOfUsers -ne 0) {
                $users = Get-ADUser -Filter "Name -like '*$queryName*'" -Property *  | Sort-Object Name | Select-Object -ExpandProperty Name                
                $ShowNames.Items.Clear()
                $ShowNames.Items.AddRange($users)
                $countLabel.Text = "Found $numberOfUsers."                 
            }
            else {
                $countLabel.Text = $numberOfUsers                 
                $ShowNames.Items.Clear()
                $ShowNames.Items.AddRange("No Users Found.")
            }
        }
    }
    $findTextBox.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $getUserDetail.PerformClick()
            }
        })
    # add enter press down
    $form.Controls.Add($getUserDetail)

    #Find user stuff
    $countLabel = New-Object System.Windows.Forms.Label
    $countLabel.Location = New-Object System.Drawing.Point(305, 5)
    $countLabel.Size = New-Object System.Drawing.Size(50, 15)
    $countLabel.ForeColor = 'Gray'
    $countLabel.Text = '0'
    $form.Controls.Add($countLabel)

    $ShowNames = New-Object System.Windows.Forms.ListBox
    $ShowNames.Size = New-Object System.Drawing.Size(225, 215)
    $ShowNames.Location = New-Object System.Drawing.Size(290, 20) 
    $disabledItem1 = "No Users Found."
    $ShowNames.Add_SelectedIndexChanged({
            $selectedItem = $ShowNames.SelectedItem
            if ($selectedItem -ne $disabledItem1) {
                $findTextBox.Text = $selectedItem
            }
        })
    $form.Controls.Add($ShowNames)


    #Horizontal line
    $lineLabel = New-Object System.Windows.Forms.Label
    $lineLabel.Location = New-Object System.Drawing.Point(0, 80)
    $lineLabel.Text = " "
    $lineLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D;
    $lineLabel.AutoSize = $false
    $lineLabel.Height = 2
    $lineLabel.Width = 290
    $form.Controls.Add($lineLabel)

    #Find user stuff
    $otherLabel = New-Object System.Windows.Forms.Label
    $otherLabel.Location = New-Object System.Drawing.Point(15, 90)
    $otherLabel.Size = New-Object System.Drawing.Size(250, 20)
    $otherLabel.Font = 'Microsoft Sans Serif,8'
    $otherLabel.Text = 'Other Actions'
    $form.Controls.Add($otherLabel)

    #Create user stuff
    $newUserButton = New-Object System.Windows.Forms.Button
    $newUserButton.Location = New-Object System.Drawing.Point(15, 115)
    $newUserButton.Size = New-Object System.Drawing.Size(75, 50)
    $newUserButton.Text = 'New User'
    $newUserButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($newUserButton)

    #mark on leave stuff
    $onLeaveButton = New-Object System.Windows.Forms.Button
    $onLeaveButton.Location = New-Object System.Drawing.Point(105, 115)
    $onLeaveButton.Size = New-Object System.Drawing.Size(75, 50)
    $onLeaveButton.Text = 'On-Leave'
    $onLeaveButton.DialogResult = [System.Windows.Forms.DialogResult]::YES
    $form.Controls.Add($onLeaveButton)

    #Move to deletes
    $delButton = New-Object System.Windows.Forms.Button
    $delButton.Location = New-Object System.Drawing.Point(195, 115)
    $delButton.Size = New-Object System.Drawing.Size(75, 50)
    $delButton.Text = 'Move To Deletes'
    $delButton.DialogResult = [System.Windows.Forms.DialogResult]::NO
    $form.Controls.Add($delButton)

    #Reset Password
    $restPassButton = New-Object System.Windows.Forms.Button
    $restPassButton.Location = New-Object System.Drawing.Point(15, 175)
    $restPassButton.Size = New-Object System.Drawing.Size(75, 50)
    $restPassButton.Text = 'Reset Password'
    $restPassButton.DialogResult = [System.Windows.Forms.DialogResult]::RETRY
    $form.Controls.Add($restPassButton)

    #mark on leave stuff
    $moveUserButton = New-Object System.Windows.Forms.Button
    $moveUserButton.Location = New-Object System.Drawing.Point(105, 175)
    $moveUserButton.Size = New-Object System.Drawing.Size(75, 50)
    $moveUserButton.Text = 'Move User'
    $moveUserButton.DialogResult = [System.Windows.Forms.DialogResult]::IGNORE
    $form.Controls.Add($moveUserButton)

    #Nothing
    $justButton = New-Object System.Windows.Forms.Button
    $justButton.Location = New-Object System.Drawing.Point(400, 400)
    $justButton.Size = New-Object System.Drawing.Size(75, 50)
    $justButton.Text = 'Secret'
    $justButton.Add_Click{
        [System.Diagnostics.Process]::Start("chrome", "https://www.youtube.com/watch?v=dQw4w9WgXcQ")
    }
    $form.Controls.Add($justButton)


    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        AddNotCasual
    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::YES) {
        LeaveMenu($findTextBox.Text)
    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::NO) {
        MoveToDeletesMenu($findTextBox.Text)
    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::RETRY) {
        ResetPasswordMenu($findTextBox.Text)
    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::IGNORE) {
        MoveUserMenu($findTextBox.Text)
    }
}

############################################################################
#INITIAL SEQUENCE

#Global variables
$Global:defaultPass = "Gssdstaff1"
$Global:company = "GSSD"
$Global:onLeaveMark = "(On Leave)"

$Global:delFolderName = "Deleted Users";
$Global:delFolderOU = "OU=$Global:delFolderName,OU=GSSD Network,DC=GSSD,DC=ADS"

$Global:schoolsOUPath = "OU=Schools,OU=GSSD Network,DC=GSSD,DC=ADS"
$arrOfSchools = Get-ADOrganizationalUnit -Filter * -SearchBase $Global:schoolsOUPath -SearchScope OneLevel | Select-Object Name
$Global:allSchools = $arrOfSchools.Name

$Global:gsecOUPath = "OU=GSEC,OU=GSSD Network,DC=GSSD,DC=ADS"
$arrOfDepartmentsInGsec = Get-ADOrganizationalUnit -Filter * -SearchBase $Global:gsecOUPath -SearchScope OneLevel | Select-Object Name
$Global:allDeptInGsec = $arrOfDepartmentsInGsec.Name

$Global:transpoOUPath = "OU=Transportation,OU=GSEC,OU=GSSD Network,DC=GSSD,DC=ADS"
$arrOfDepartmentsInTrans = Get-ADOrganizationalUnit -Filter * -SearchBase $Global:transpoOUPath -SearchScope OneLevel | Select-Object Name
$Global:allDeptInTrans = $arrOfDepartmentsInTrans.Name

#Start of Main 
MainMenu