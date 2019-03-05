[CmdletBinding()]

Param(
    [Parameter(Mandatory=$true,Position=1)]
        [string]$emailToExport,

    [Parameter(Mandatory=$true,Position=2)]
        [string]$serviceTicketNumber
    )

Start-Transcript
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto

$workingdir = Split-Path $MyInvocation.MyCommand.Path -Parent #Get current working directory
$hostedExchchangeConfFile = "$workingdir\HostedExchange.conf"

# $emailToExport = Read-Host "What is the email address to export?"
$serviceTicketNumber = $serviceTicketNumber -replace '\s',''
$exportDirectory = Get-Content $hostedExchchangeConfFile | Select-Object -Index 0
Write-Host "The mailbox will be exported to $exportDirectory

Please delete when finished."

New-MailboxExportRequest -Mailbox $emailToExport -AcceptLargeDataLoss -BadItemLimit 1000 -FilePath "$exportDirectory\$serviceTicketNumber`_$emailToExport.pst" -DomainController HOST-DC41.Hosted1.local -Verbose