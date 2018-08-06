Start-Transcript
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto
Import-Module ActiveDirectory

$companyName = Read-Host "Input Company Name (short)"
$emailDomain = Read-Host "Input email domain without @ symbol"

New-ADOrganizationalUnit -Name "$companyName" -Path "OU=Hosting,DC=Hosted1,DC=local"
New-ADOrganizationalUnit -Name "$companyName - Groups" -Path "OU=$companyName,OU=Hosting,DC=Hosted1,DC=local"
New-ADOrganizationalUnit -Name "$companyName - Users" -Path "OU=$companyName,OU=Hosting,DC=Hosted1,DC=local"
New-ADOrganizationalUnit -Name "$companyName - Contacts" -Path "OU=$companyName,OU=Hosting,DC=Hosted1,DC=local"

Write-Host "Create GAL"

New-GlobalAddressList -Name "$companyName - GAL" -RecipientFilter {CustomAttribute1 -like "$companyName"}

Write-Host "Create UPN"

Get-ADForest | Set-ADForest -UPNSuffixes @{add="$emailDomain"}

Write-Host "Create Accepted Domain"

New-AcceptedDomain -Name "$companyName - $emailDomain" -DomainName "$emailDomain" -DomainType Authoritative

Write-Host "Create and Update EAP"

New-EmailAddressPolicy -Name "$companyName" -RecipientContainer "Hosted1.local/Hosting/$companyName" -IncludedRecipients AllRecipients -ConditionalCustomAttribute1 $companyName -Priority 1 -EnabledEmailAddressTemplates "SMTP:%m@$emailDomain"
Update-EmailAddressPolicy -Identity "$companyName"

Write-Host "Create and Update Address List"
New-AddressList -Name "$companyName Contacts" -RecipientContainer "Hosted1.local/Hosting/$companyName" -IncludedRecipients AllRecipients -ConditionalCustomAttribute1 $companyName -Container "\" -DisplayName "$companyName Contacts"
Update-AddressList -Identity "\$companyName Contacts"

Write-Host "Create Offline Address Book"
New-OfflineAddressBook -Name "$companyName`_2016 OAB" -AddressLists "\$companyName Contacts" -VirtualDirectories "HOST-EXCH90\OAB (Default Web Site)"

Write-Host "Create Address Book Policy"
New-AddressBookPolicy -Name "$companyName ABP" -GlobalAddressList "\$companyName - GAL" -OfflineAddressBook "\$companyName`_2016 OAB" -RoomList "\$companyName Contacts" -AddressLists "\$companyName Contacts"

Write-Host "Company created. Please create the company on Reflexion now and don't forget to review the Email Address Policy" -ForegroundColor Yellow -BackgroundColor Red