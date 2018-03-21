Start-Transcript
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010

$firstName = Read-Host "First Name"
$lastName = Read-Host "Last Name"
$companies = Get-Content companies.txt
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

Write-Host You are going to create a mailbox for $firstName $lastName at $selectedCompany. Are you sure you want to continue?
$yesNo = Read-Host "Are you sure you want to continue? (y/n)"
Switch ($yesNo) {
    Y {Write-host "Creating mailbox now"}
    N {Write-Host "Cancelling"; Exit}
    Default {Write-Host "Cancelling"; Exit}
    }

Write-Host "The script has continued on"