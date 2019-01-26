Import-Module ActiveDirectory

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto
$MyDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Month = Get-Date -UFormat %B
$time = Get-Date
$timestamp = get-date -f yyyy-MM-dd
$logfile = "$myDir\logs\logfile.log"
Add-Content $logfile "Script ran on $timestamp"

$filterednames = @("dmarc","mimecast","guest","LDAP","vmware","dss","opendns","sp admin","dcg","qbdataservice","sql","st_bernard","hosted","ldapadmin","spadmin","test","noc","st. bernard","st bernard","managed care","bbadmin","besadmin","compliance","discovery","rmmscan","healthmailbox","sharepoint","windows sbs","qbdata","noc_helpdesk","appassure","support","scanner","ftp","app assure","aspnet","Dependable Computer Guys","efax","exchange","INSTALR","IUSR","IWAM","NAV","Quick Books")
$regex = "(" + ($filterednames -join "|") + ")"
Add-Content $logfile "Filtering out the following names: $filterednames"

$PSEmailServer = "mail.server.com"
$SMTPPort = 2525
$SMTPUsername = "username@domain.com"
$EncryptedPasswordFile = "$mydir\email.securestring"
$SecureStringPassword = Get-Content -Path $EncryptedPasswordFile | ConvertTo-SecureString
$EmailCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SMTPUsername,$SecureStringPassword
$MailTo = "mailto@domain.com"
$MailFrom = "mailfrom@domain.com"
$mailAttachments = @()
$dbcountAttachment = "$MyDir\logs\$timestamp`_UsersCount.csv"
$dbMailboxesAttachment = "$MyDir\logs\$timestamp`_UserBreakdown.csv"

Add-Content $logfile 'Beginning search of Exchange mailboxes.'

$databases = Get-MailboxDatabase | Sort-Object Name

foreach ($d in $databases) {
    $dbmailboxes = Get-Mailbox -Database $d -RecipientTypeDetails UserMailbox,LinkedMailbox  | Where-Object {$_ -notmatch $regex}
    $dbcount = $dbmailboxes.count
    If ($dbcount -ne 0) {
        $CountExport = New-Object PSObject
        $CountExport | Add-Member -MemberType NoteProperty -Name "Database" -Value $d.Name
        $CountExport | Add-Member -MemberType NoteProperty -Name "Count" -Value $dbcount
        $CountExport | Export-Csv $dbcountAttachment -Append -NoTypeInformation
        $dbmailboxes | Select-Object Database,Name,PrimarySmtpAddress,CustomAttribute1,Organizationalunit | 
            Sort-Object Name | Export-Csv $dbMailboxesAttachment -Append -NoTypeInformation
        }
    }


Add-Content $logfile "Sending email now"

$MailSubject = "$Month's  Hosted Exchange audit"
$MailBody = "$Month's SPLA Audit Info

Attached are the CSV files with the customers and number of Hosted Exchange accounts and the breakdown of the individual accounts that has the users for each company.

This message was sent from $env:COMPUTERNAME"

$mailAttachments = $mailAttachments += $dbcountAttachment
$mailAttachments = $mailAttachments += $dbMailboxesAttachment

Send-MailMessage -From $MailFrom -To $MailTo -Subject $MailSubject -Body $MailBody -Port $SMTPPort -Credential $EmailCredential -Attachments $mailAttachments

