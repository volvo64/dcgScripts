Start-Transcript
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto

$csv = Read-Host "Type path (or drop file) to CSV here"
$company = Read-Host "What is the company name (short)?"

$csv  |

ForEach-Object {
    New-DistributionGroup -Name "$company - $($_.DistributionGroup)" -Alias $_.Alias -Type Distribution -Members $_.ForwardingAddress1 -OrganizationalUnit $_.OrganizationalUnit
    If ($_.ForwardingAddress2 -ne "") {
        Start-Sleep 5
        Add-DistributionGroupMember -Identity "$company - $($_.DistributionGroup)" -Member $_.ForwardingAddress2
    }
}

Start-Sleep 20

$csv  | ForEach-Object {Set-DistributionGroup -Identity "$company - $($_.DistributionGroup)" -CustomAttribute1 $_.CustomAttribute -RequireSenderAuthenticationEnabled $False}