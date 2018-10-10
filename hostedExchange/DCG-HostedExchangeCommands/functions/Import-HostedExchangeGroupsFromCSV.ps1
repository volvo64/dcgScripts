[CmdletBinding()]
Param(
    [Parameter(Mandatory = $True, Position = 1)]
    [string]$csvFile

    
)


Start-Transcript
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto

#This is a work in progress. I changed a bunch of stuff here and have _not_ tested it.

$csv = Import-Csv $csvFile
$company = Read-Host "What is the company name (short)?"


ForEach ($c in $csv) {
    New-DistributionGroup -Name "$company - $($c.DistributionGroup)" -Alias $c.Alias -Type Distribution -Members $c.ForwardingAddress1 -OrganizationalUnit $c.OrganizationalUnit
    If ($c.ForwardingAddress2 -ne "") {
        Sleep 5
        Add-DistributionGroupMember -Identity "$company - $($c.DistributionGroup)" -Member $c.ForwardingAddress2
        Set-DistributionGroup -Identity "$company - $($c.DistributionGroup)" -CustomAttribute1 $c.CustomAttribute -RequireSenderAuthenticationEnabled $False
    }
}