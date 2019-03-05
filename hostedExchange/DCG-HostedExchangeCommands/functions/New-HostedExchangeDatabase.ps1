[CmdletBinding()]
Param (
	[Parameter(Position=1,Mandatory=$true)]
		[string]$client

	)

Start-Transcript
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto

$databaseName = "$client`_2016"
$drive = "F:"
$server = "main-server"
$copyServer = "copy-server"

New-MailboxDatabase -Server $server -Name $databaseName -EdbFilePath "$drive\ExchDB\$databaseName\$databaseName.edb" -LogFolderPath "$drive\ExchDB\$databaseName"
Sleep 5

#create a DB copy to copy-server and restart the Information Store.
#restarted the IS will disconnect Outlook clients; but this server should only be used for copies.
Add-MailboxDatabaseCopy -Identity $databaseName -MailboxServer $copyServer -ActivationPreference 2
Get-Service -Name MSExchangeIS -ComputerName $copyServer | restart-service

Set-MailboxDatabase -Identity $databaseName -IssueWarningQuota unlimited -ProhibitSendReceiveQuota unlimited -ProhibitSendQuota unlimited -OfflineAddressBook "$databaseName OAB" -CircularLoggingEnabled:$true

Dismount-Database -Identity $databaseName
sleep 5

Mount-Database -Identity $databaseName