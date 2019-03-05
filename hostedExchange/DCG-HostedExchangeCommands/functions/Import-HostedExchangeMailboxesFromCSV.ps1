[CmdletBinding()]

Param(
    [Parameter(Mandatory=$True)]
    [string]$csv,

    [Parameter(Mandatory=$True)]
    [string]$strpassword
    )

Start-Transcript
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto

$administrator = 'domain\administrator'
$password = ConvertTo-SecureString -string $strpassword -AsPlainText -Force

$import = Import-csv $csv
 
ForEach ($u in $import) {
        New-Mailbox -FirstName $u.FirstName -LastName $u.LastName -Alias $u.Alias -Name $u.name -userPrincipalName $u.UserPrincipalName -SamAccountName $u.alias -Database $u.Database -OrganizationalUnit $u.OrganizationalUnit -Password $Password -AddressBookPolicy $u.AddressBookPolicy -ResetPasswordOnNextLogon $false
        Set-Mailbox -Identity $u.Alias -CustomAttribute1 $u.CustomAttribute1 -DefaultPublicFolderMailbox $u.PublicFolderMailbox
        Add-MailboxPermission -Identity $u.Alias -User $administrator -AccessRight FullAccess -InheritanceType All -Automapping $false
        }