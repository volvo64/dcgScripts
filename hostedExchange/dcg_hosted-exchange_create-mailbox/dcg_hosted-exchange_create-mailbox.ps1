Start-Transcript


$workingdir = Split-Path $MyInvocation.MyCommand.Path -Parent #Get current working directory
$firstName = Read-Host "First Name"
$lastName = Read-Host "Last Name"
$companies = Get-Content "$workingdir\companies.txt"
$companies = $companies | sort

# https://docs.microsoft.com/en-us/powershell/scripting/getting-started/cookbooks/selecting-items-from-a-list-box?view=powershell-6

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form 
$form.Text = "Select a Company"
$form.Size = New-Object System.Drawing.Size(300,200) 
$form.StartPosition = "CenterScreen"

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(75,120)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150,120)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20) 
$label.Size = New-Object System.Drawing.Size(280,20) 
$label.Text = "Please select a company:"
$form.Controls.Add($label) 

$listBox = New-Object System.Windows.Forms.ListBox 
$listBox.Location = New-Object System.Drawing.Point(10,40) 
$listBox.Size = New-Object System.Drawing.Size(260,20) 
$listBox.Height = 80

$i = 0
while ($i -lt $companies.length)
    {
    $c = $companies[$i]
    [void] $listBox.Items.Add($c)
    $i = $i + 1
    }

$form.Controls.Add($listBox) 

$form.Topmost = $True

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $selectedCompany = $listBox.SelectedItem
}

$configFile = "$workingdir\conf\$selectedCompany.conf"
$custAttr1 = Get-Content $configFile | Select-Object -Index 1
$dn =  Get-Content $configFile | Select-Object -Index 2
$emailPrefix = Get-Content $configFile | Select-Object -Index 4
$emailPrefix = Invoke-Expression -Command $emailPrefix
$emailDomain = Get-Content $configFile | Select-Object -Index 3
$emailDatabase = Get-Content $configFile | Select-Object -Index 5
$emailABP = Get-Content $configFile | Select-Object -Index 6
$emailAddress = "$emailPrefix$emailDomain"
Write-Host $emailPrefix
Write-Host $custAttr1 $dn


If ((Read-Host "Create a (1) mailbox or (2) distribution list") -eq 1) {
Write-Host You are going to create a mailbox for $firstName $lastName at $selectedCompany.
$yesNo = Read-Host "Are you sure you want to continue? (y/n)"
Switch ($yesNo) {
    Y {Write-host "Creating mailbox now"}
    N {Write-Host "Cancelling"; Exit}
    Default {Write-Host "Cancelling"; Exit}
    }

# Create the mailbox on this line
$password = Read-Host "Enter password" -AsSecureString; New-Mailbox -UserPrincipalName $emailAddress -Alias $emailPrefix -Database $emailDatabase -Name "$firstName $lastName" -OrganizationalUnit $dn  -Password $password -FirstName $firstName -LastName $lastName -DisplayName "$firstName $lastName" -ResetPasswordOnNextLogon $false -AddressBookPolicy $emailABP
sleep 5
Set-Mailbox $emailaddress -CustomAttribute1 $custAttr1
sleep 5
# Send a test email

$PSEmailServer = "host-exch90.dcgla.com"
$EmailCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $emailaddress,$password
$mailto = Read-Host "Enter your email address to send a test message"
Send-MailMessage -From $emailaddress -To $mailto -Subject "Test Message" -Port 2525 -Credential $EmailCredential -SmtpServer host-exch90.dcgla.com
    } 

    ElseIf (-eq 2) {
        Write-host "Creating a distribution list"
        
        New-DistributionGroup -Name "$selectedCompany - $emailAddress