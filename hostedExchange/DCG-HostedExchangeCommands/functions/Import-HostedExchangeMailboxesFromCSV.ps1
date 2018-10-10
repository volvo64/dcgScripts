Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010


$adminMailbox = 
$csv = Read-Host "Type path (or drop file) to CSV here"
$Password = Read-Host "Enter password" -AsSecureString 
Import-csv $csv | 
    ForEach {
    New-Mailbox -FirstName $_.FirstName -LastName $_.LastName -Alias $_.alias -Name $_.name -userPrincipalName $_.UserPrincipalName -SamAccountName $_.SamAccountName -Database $_.Database -OrganizationalUnit $_.OrganizationalUnit -Password $Password -AddressBookPolicy $_.AddressBookPolicy -ResetPasswordOnNextLogon $true
}

Sleep 10

Import-csv $csv |
    Foreach {
    Set-Mailbox -Identity $_.UserPrincipalName -CustomAttribute1 $_.CustomAttribute1
    If ($_.smtpaddress1 -ne "") {
        Set-Mailbox -Identity $_.UserPrincipalName -EmailAddresses @{add = ${$._smtpaddress1}}
    }
    Add-MailboxPermission $_.UserPrincipalName -User $adminMailbox -AccessRight FullAccess -InheritanceType All -Automapping $false
}