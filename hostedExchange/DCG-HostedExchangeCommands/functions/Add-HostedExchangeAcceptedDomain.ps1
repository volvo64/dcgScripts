[CmdletBinding()]
Param(
	[Parameter(Mandatory=$True,Position=1)]
		[string]$companyName,

	[Parameter(Mandatory=$True,Position=2)]
		[string]$emailDomain
	)

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010

Write-Host "Create UPN"

Get-ADForest | Set-ADForest -UPNSuffixes @{add="$emailDomain"}

Write-Host "Create Accepted Domain"

New-AcceptedDomain -Name "$companyName - $emailDomain" -DomainName "$emailDomain" -DomainType Authoritative

Write-Host "Accepted Domain created. Please update the Email Address Policy for the client if necessary" -BackgroundColor Red

sleep 60