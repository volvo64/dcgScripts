[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True,Position=1)]
        [string]$baseOU,

    [string]$outFile
    
    )

$workingdir = Split-Path $MyInvocation.MyCommand.Path -Parent
$companyName = ( $baseOU -replace '\s','' -replace ',','')
$companyName = ($companyName -split "OU=")[1]
$now = Get-Date -Format o | foreach {$_ -replace ":", "."}

If (!$outFile) {$outFile = "~\Desktop\$companyName`_$now.csv"}

Start-Transcript
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto

$recipients = Get-Recipient -OrganizationalUnit $baseOU

Foreach ($m in $recipients) {
    If ($m.RecipientType -eq "UserMailbox") {
    Get-Mailbox $m.DistinguishedName | 
     select name,
     RecipientType,
     primarysmtpaddress,
     HiddenFromAddressListsEnabled,
      @{Name='EmailAddresses'; Expression={$_.EmailAddresses -join ","}} |
     sort primarysmtpaddress -unique | Export-Csv $outFile -NoTypeInformation -Append
    }
    If ($m.RecipientType -eq "MailContact") {
        Get-MailContact $m.DistinguishedName | 
        select name,
        RecipientType,
        primarysmtpaddress,
        HiddenFromAddressListsEnabled,
        @{Name='EmailAddresses'; Expression={$_.EmailAddresses -join ","}} |
        sort primarysmtpaddress -unique | Export-Csv $outFile -NoTypeInformation -Append
        }

    <# If (($m.RecipientType -eq "MailUniversalSecurityGroup") -or ($m.RecipientType -eq "MailUniversalSecurityGroup")) {
        
        Get-DistributionGroup $m.DistinguishedName | 
        select name,RecipientType,primarysmtpaddress, @{Name='EmailAddresses'; Expression={$_.EmailAddresses -join ","}} |
        sort primarysmtpaddress -unique | Export-Csv $outFile -NoTypeInformation -Append

        Get-DistributionGroupMember $m.DistinguishedName | 
        select @{Name='Members'; Expression={$_.Name -join ","}} |
        Export-Csv $outFile -NoTypeInformation -Append -Force
        } #>
    }

Write-Host "The report has been saved to $outFile"