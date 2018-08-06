Start-Transcript
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto

$workingdir = Split-Path $MyInvocation.MyCommand.Path -Parent #Get current working directory
$hostedExchchangeConfFile = "$workingdir\HostedExchange.conf"

$emailToExport = Read-Host "What is the email address to export?"
$serviceTicketNumber = Read-Host "What is the service ticket number?"
$exportDirectory = Get-Content $hostedExchchangeConfFile | Select-Object -Index 0
Write-Host "The mailbox will be exported to $exportDirectory

Please delete when finished."

New-MailboxExportRequest -Mailbox $emailToExport -AcceptLargeDataLoss -BadItemLimit 1000 -FilePath "$exportDirectory\$serviceTicketNumber $emailToExport.pst"