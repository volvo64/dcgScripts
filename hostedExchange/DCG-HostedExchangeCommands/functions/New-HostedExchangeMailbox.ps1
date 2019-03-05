[CmdletBinding(DefaultParametersetName='manual')]

Param(
    [Parameter(Mandatory=$true)]
    [string]$firstName,

    [Parameter(Mandatory=$true)]
    [string]$lastName,

    [string]$selectedCompany,

    [Parameter(Mandatory=$true)]
    [string]$mailto,

    [Parameter(Mandatory=$true)]
    [string]$strpassword,

    [Parameter(ParameterSetName='csv')]
    [switch]$ImportFromCsv,

    [Parameter(ParameterSetName='csv',Mandatory=$True)]
    [string]$pathToCsv

    )

Start-Transcript
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto

$administrator = 'hosted1\administrator'

$workingdir = Split-Path $MyInvocation.MyCommand.Path -Parent #Get current working directory

If ([string]::IsNullOrWhiteSpace($selectedCompany)) {
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
}
#this section should be rewritten to use a json and an array
$configFile = "$workingdir\conf\$selectedCompany.conf"
$custAttr1 = Get-Content $configFile | Select-Object -Index 1
$dn =  Get-Content $configFile | Select-Object -Index 2
$emailDomain = Get-Content $configFile | Select-Object -Index 3
$emailDatabase = Get-Content $configFile | Select-Object -Index 5
$emailABP = Get-Content $configFile | Select-Object -Index 6
$publicFolderMailbox = Get-Content $configFile | Select-Object -Index 7
$roleAssignmentPolicy = Get-Content $configFile | Select-Object -Index 14

#manual mailbox creation below this line:

If ($ImportFromCsv -eq $false) {

$emailPrefix = Get-Content $configFile | Select-Object -Index 4
$emailPrefix = Invoke-Expression -Command $emailPrefix
$emailAddress = "$emailPrefix$emailDomain"
Write-Host $emailPrefix
Write-Host $custAttr1 $dn

Write-Host You are going to create a mailbox for $firstName $lastName at $selectedCompany.
sleep 5

#create the optional params list here

If (($firstName -ne "") -and ($lastName -ne "")) {$longName = "$firstName $lastName";$params = @{
    "Name" = $longName
    "DisplayName" = $longName
    "FirstName" = $firstName
    "LastName" = $lastName}}
If (($firstName -ne "") -and ($lastName -eq "")) {$params = @{
    "Name" = $firstName
    "DisplayName" = $firstName
    "FirstName" = $firstName}}
If (($firstName -eq "") -and ($lastName -ne "")) {$params = @{
    "Name" = $lastName
    "DisplayName" = $lastName
    "LastName" = $lastName}}


# Create the mailbox on this line
Write-Host $strpassword
$password = ConvertTo-SecureString -string $strpassword -AsPlainText -Force
New-Mailbox -UserPrincipalName $emailAddress -Database $emailDatabase -OrganizationalUnit $dn  -Password $password -ResetPasswordOnNextLogon $false -AddressBookPolicy $emailABP @params
sleep 5
Set-Mailbox $emailaddress -CustomAttribute1 $custAttr1 -DefaultPublicFolderMailbox $publicFolderMailbox
If ($roleAssignmentPolicy -ne "") {Set-Mailbox $emailAddress -RoleAssignmentPolicy $roleAssignmentPolicy}
sleep 5
Add-MailboxPermission -Identity $emailAddress -User $administrator -AccessRight FullAccess -InheritanceType All -Automapping $false

# Send a test email

$PSEmailServer = "host-exch90.dcgla.com"
$EmailCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $emailaddress,$password
# $mailto = Read-Host "Enter your email address to send a test message"
Send-MailMessage -From $emailaddress -To $mailto -Subject "Test Message" -Port 2525 -Credential $EmailCredential

#end manual mailbox creation
}

#create mailboxes from csv
#this probably doesn't work
If ($ImportFromCsv -eq $true) {
    $csv = Import-Csv -Path $pathToCsv
    foreach ($u in $csv) {
        
Write-Host $custAttr1 $dn


# Create the mailbox on this line
$password = ConvertTo-SecureString -string $u.strpassword -AsPlainText -Force
New-Mailbox -Name $u.DisplayName -UserPrincipalName $u.userPrincipalName -Database $emailDatabase -OrganizationalUnit $dn  -Password $password -ResetPasswordOnNextLogon $false -AddressBookPolicy $emailABP -DisplayName $u.DisplayName -FirstName $u.FirstName -LastName $u.LastName

sleep 5

#set mailbox params:
If ($roleAssignmentPolicy -ne "") {$setParams = @{
    RoleAssignmentPolicy = $roleAssignmentPolicy
    }
Set-Mailbox $u.userPrincipalName -CustomAttribute1 $custAttr1 -DefaultPublicFolderMailbox $publicFolderMailbox @setParams

sleep 5
Add-MailboxPermission -Identity $u.userPrincipalName -User 'domain\Administrator' -AccessRight FullAccess -InheritanceType All -Automapping $false


# Send a test email

$PSEmailServer = "mail.server.com"
$EmailCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $u.userPrincipalName,$password

Send-MailMessage -From $u.userPrincipalName -To $mailto -Subject "Test Message" -Port 2525 -Credential $EmailCredential
}
}
}