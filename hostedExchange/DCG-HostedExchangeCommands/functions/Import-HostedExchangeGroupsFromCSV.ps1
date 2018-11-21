Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010

$company = Read-Host "What is the company name (short)?"

Import-csv #csv file here |

ForEach {
    New-DistributionGroup -Name "$company - $($_.DistributionGroup)" -Alias $_.Alias -Type Distribution -Members $_.ForwardingAddress1 -OrganizationalUnit $_.OrganizationalUnit
    If ($_.ForwardingAddress2 -ne "") {
        Sleep 5
        Add-DistributionGroupMember -Identity "$company - $($_.DistributionGroup)" -Member $_.ForwardingAddress2
        }
    }

Sleep 20

Import-csv #csv file here | ForEach {Set-DistributionGroup -Identity "$company - $($_.DistributionGroup)" -CustomAttribute1 $_.CustomAttribute -RequireSenderAuthenticationEnabled $False}