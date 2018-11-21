[CmdletBinding()]
Param (
    [Parameter(Position = 1, Mandatory = $true)]
    [string]$client

)

Start-Transcript
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto

$databaseName = "$client`_2016"
$drive = "F:"
$server = #server name here

New-MailboxDatabase -Server $server -Name $databaseName -EdbFilePath "$drive\ExchDB\$databaseName\$databaseName.edb" -LogFolderPath "$drive\ExchDB\$databaseName" -WarningAction SilentlyContinue
Sleep 60

Set-MailboxDatabase -Identity $databaseName -IssueWarningQuota unlimited -ProhibitSendReceiveQuota unlimited -ProhibitSendQuota unlimited -OfflineAddressBook "$databaseName OAB" -CircularLoggingEnabled:$true

Dismount-Database -Identity $databaseName
sleep 5

Mount-Database -Identity $databaseName
Sleep 60