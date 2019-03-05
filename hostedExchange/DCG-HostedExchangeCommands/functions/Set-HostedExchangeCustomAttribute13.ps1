[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True,Position=1)]
        [string]$file,

    [Parameter(Mandatory=$true)]
        [string]$custAttribute13
    )

Start-Transcript
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto



foreach ($r in $(Get-Content $file)) {
    Set-Mailbox $r -CustomAttribute13 $custAttribute13
    Get-Mailbox $r | select name,customattribute13 | ft
    Set-DistributionGroup $r -CustomAttribute13 $custAttribute13
    Get-DistributionGroup $r | select name,customattribute13 | ft
    }