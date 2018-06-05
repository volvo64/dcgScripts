#For exporting or saving, this gets the directory from which the script is run.
$MyDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Month = Get-Date -UFormat %B
$time = Get-Date
$logfile = "$myDir\logs\logfile.log"
$confFile = "$myDir\company.conf"
$companyName = Get-Content $confFile | Select-Object -Index 0
$auditType = Get-Content $confFile | Select-Object -Index 1
$companyContact = Get-Content $confFile | Select-Object -Index 2
$companyContactEmail = Get-Content $confFile | Select-Object -Index 3
$rdsGroup = Get-Content $confFile | Select-Object -Index 5

$filterednames = @("dcg","administrator","qbdataservice","sql","st_bernard","hosted","ldapadmin","spadmin","test","noc","st. bernard","st bernard","managed care","bbadmin","besadmin","compliance","discovery","rmmscan","healthmailbox","sharepoint","windows sbs","qbdata","noc_helpdesk","appassure","support","scanner","ftp")
$perEnvFilteredNames = get-content $confFile | Select-Object -Index 4
$filterednames = $filterednames += $perEnvFilteredNames
$regex = "(" + ($filterednames -join "|") + ")"

$PSEmailServer = "host-exch90.dcgla.com"
$SMTPPort = 2525
$SMTPUsername = "scriptsender@dcgla.net"
$EncryptedPasswordFile = "$mydir\scriptsender@dcgla.net.securestring"

$SecureStringPassword = Get-Content -Path $EncryptedPasswordFile | ConvertTo-SecureString
$EmailCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SMTPUsername,$SecureStringPassword
$MailTo = "monitoring@dcgla.com"
$MailFrom = "scriptsender@dcgla.net"
$mailAttachments = ""
$extraMailBodyInfo = Get-Content $confFile | Select-Object -Index 6


If ($auditType -match 1) {
$usernamesraw = (Get-AdUser -filter * |Where {($_.enabled -eq "True")}).name

$usernamesfiltered = $usernamesraw | ? {$_ -notmatch $regex}

Get-Date >> $logfile
Add-Content $logfile 'Names of AD users: '
$usernamesfiltered >> $logfile
Add-Content $logfile 'Count of AD users: '
$usernamesfiltered.count >> $logfile
$usernamesfilteredCount = $usernamesfiltered.count
$usernamesfiltered | Sort-Object > "$MyDir\logs\$time AD Users.txt"
$adUsersAttachment = "$MyDir\logs\$time AD Users.txt"
}

If ($auditType -match 2) {

$rdsUsersRaw = (Get-ADGroupMember -Identity $rdsGroup | Where {($_.enabled -eq "True")}).name

$rdsUsersFiltered = $rdsUsersRaw | ? {$_ -notmatch $regex}
Add-Content $logfile 'Names of RDS Users:'
$rdsUsersFiltered >> $logfile
Add-Content $logfile 'Count of RDS Users:'
$rdsUsersFiltered.count >> $logfile
$rdsUsersFilteredCount = $rdsUsersFiltered.count
$rdsUsersFilteredFormatted = $rdsUsersFiltered -join "`r`n" | Out-String
$rdsUsersFiltered |Sort-Object > "$MyDir\logs\$time Remote Desktop Users.txt"
$rdsUsersAttachment = "$MyDir\logs\$time Remote Desktop Users.txt"
}

If ($auditType -match 3) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
    $mailAccountsRaw = (Get-Mailbox -ResultSize Unlimited -WarningAction SilentlyContinue).DisplayName
    $mailAccountsFiltered = $mailAccountsRaw | ? {$_ -notmatch $regex}
    Add-Content $logfile 'Names of Exchange mailboxes:'
    $mailAccountsFiltered >> $logfile
    Add-Content $logfile 'Count of Exchange mailboxes:'
    $mailAccountsFiltered.count >> $logfile
    $mailAccountsFilteredCount = $mailAccountsFiltered.Count
    $mailAccountsFiltered | Sort-Object > "$MyDir\logs\$(get-date -f yyyy-MM-dd)ExchangeUsers.txt"
    $exchangeUsersAttachment = "$MyDir\logs\$(get-date -f yyyy-MM-dd)ExchangeUsers.txt"
    }

$MailSubject = "$companyContact, Please review $companyName's DCG PrivateCLOUD SPLA counts before $Month 15th"
$MailBody = "DCG strives to maintain an accurate active user list, as it pertains to your PrivateCLOUD server SPLA licensing counts on your server(s).  Attached is your current user count related Remote Desktop Services, MS Office, and SQL services on your PrivateCLOUD server.  This will be referenced in your upcoming Monthly Services invoice that will be emailed to you on the 15th of this month.

"
If ($auditType -match 1) {
$mailBody = $mailBody += "Current AD Users: $usernamesfilteredcount

"
$mailAttachments = $mailAttachments += $adUsersAttachment
}

If ($auditType -match 2) {
$MailBody = $MailBody += "Current RDS Users: $rdsUsersFilteredcount

"
$mailAttachments = $mailAttachments += $rdsUsersAttachment
}

If ($auditType -match 3) {
$MailBody = $MailBody += "Current Exchange Users: $mailAccountsFilteredcount

"
$mailAttachments = $mailAttachments += $exchangeUsersAttachment
}

$MailBody = $MailBody += $extraMailBodyInfo

$MailBody = $MailBody += "

If any users should be removed from any of these lists, please contact DCG Technical Solutions before the 15th of the month.  We'll be sure to disable and remove any users so you will not incur any further licensing charges for them on your upcoming Monthly Services invoice.

Credits and refunds will not be issued after the 15th of this month.

Thank you very much for taking the time to review these reports with us

DCG Accounting"

Send-MailMessage -From $MailFrom -To $MailTo -Subject $MailSubject -Body $MailBody -Port $SMTPPort -Credential $EmailCredential -Attachments $mailAttachments
