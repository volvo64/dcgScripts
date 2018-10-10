Start-Transcript
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto

$workingdir = Split-Path $MyInvocation.MyCommand.Path -Parent #Get current working directory
$sharedMailboxName = Read-Host "Shared mailbox name"
$fullAccessUser = Read-Host "Enter the email address of the first full access user"
$companies = Get-Content "$workingdir\companies.txt"
$companies = $companies | sort

# https://docs.microsoft.com/en-us/powershell/scripting/getting-started/cookbooks/selecting-items-from-a-list-box?view=powershell-6

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form 
$form.Text = "Select a Company"
$form.Size = New-Object System.Drawing.Size(300, 200) 
$form.StartPosition = "CenterScreen"

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(75, 120)
$OKButton.Size = New-Object System.Drawing.Size(75, 23)
$OKButton.Text = "OK"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150, 120)
$CancelButton.Size = New-Object System.Drawing.Size(75, 23)
$CancelButton.Text = "Cancel"
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 20) 
$label.Size = New-Object System.Drawing.Size(280, 20) 
$label.Text = "Please select a company:"
$form.Controls.Add($label) 

$listBox = New-Object System.Windows.Forms.ListBox 
$listBox.Location = New-Object System.Drawing.Point(10, 40) 
$listBox.Size = New-Object System.Drawing.Size(260, 20) 
$listBox.Height = 80

$i = 0
while ($i -lt $companies.length) {
    $c = $companies[$i]
    [void] $listBox.Items.Add($c)
    $i = $i + 1
}

$form.Controls.Add($listBox) 

$form.Topmost = $True

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $selectedCompany = $listBox.SelectedItem
}

$configFile = "$workingdir\conf\$selectedCompany.conf"
$custAttr1 = Get-Content $configFile | Select-Object -Index 1
$dn = Get-Content $configFile | Select-Object -Index 2
$emailDomain = Get-Content $configFile | Select-Object -Index 3
$emailDatabase = Get-Content $configFile | Select-Object -Index 5
$emailABP = Get-Content $configFile | Select-Object -Index 6
$publicFolderMailbox = Get-Content $configFile | Select-Object -Index 7
$emailAddress = "$sharedMailboxName$emailDomain"
Write-Host $emailAddress
Write-Host $custAttr1 $dn

Write-Host You are going to create a shared mailbox called $emailAddress at $selectedCompany.
$yesNo = Read-Host "Are you sure you want to continue? (y/n)"
Switch ($yesNo) {
    Y {Write-host "Creating mailbox now"}
    N {Write-Host "Cancelling"; Exit}
    Default {Write-Host "Cancelling"; Exit}
}

New-Mailbox -Shared -Name "$selectedCompany - $sharedMailboxName" -DisplayName "$selectedCompany - $sharedMailboxName" -Alias $sharedMailboxName -Database $emailDatabase -OrganizationalUnit $dn -AddressBookPolicy $emailABP
Start-Sleep 5
Set-Mailbox -Identity "$selectedCompany - $sharedMailboxName" -CustomAttribute1 $custAttr1 
Add-MailboxPermission -Identity "$selectedCompany - $sharedMailboxName" -User $fullAccessUser -AccessRights FullAccess -InheritanceType All
Add-ADPermission -Identity "$selectedCompany - $sharedMailboxName" -User $fullAccessUser -AccessRights ExtendedRight -ExtendedRights "Send As"