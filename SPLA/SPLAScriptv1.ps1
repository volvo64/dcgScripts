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

If ($auditType -match 2) {
    $rdsGroup = @()
    $rdsGroup = Get-Content $confFile | Select-Object -Index 5
    $rdsGroup = $rdsGroup -split ","
    Add-Content $logfile "The RDS Group to search is  $rdsGroup"
    }

If ($auditType -match 7) {
    $sslvpnGroup = Get-Content $confFile | Select-Object -Index 7
    Add-Content $logfile "The SSLVPN group to search is $sslvpnGroup"
    }

If ($auditType -match 5) {
    $officeGroup = Get-Content $confFile | Select-Object -Index 8
    Add-Content $logfile "The Office group to search is $officeGroup"
    }

If ($auditType -match 6) {
    $BlaskGuardGroup = Get-Content $confFile | Select-Object -Index 9
    Add-Content $logfile "The BlaskGuard group to search is $BlaskGuardGroup"
    }

$filterednames = @("mailmonitor","mimecast","guest","LDAP","vmware","dss","opendns","sp admin","dcg","qbdataservice","sql","st_bernard","hosted","ldapadmin","spadmin","test","noc","st. bernard","st bernard","managed care","bbadmin","besadmin","compliance","discovery","rmmscan","healthmailbox","sharepoint","windows sbs","qbdata","noc_helpdesk","appassure","scanner","ftp","app assure","aspnet","Dependable Computer Guys","efax","exchange","INSTALR","IUSR","IWAM","Quick Books")
$perEnvFilteredNames = get-content $confFile | Select-Object -Index 4
$perEnvFilteredNames = -split $perEnvFilteredNames
$filterednames = $filterednames += $perEnvFilteredNames
$regex = "(" + ($filterednames -join "|") + ")"
Add-Content $logfile "Filtering out the following names: $filterednames"

$PSEmailServer = "host-exch90.dcgla.com"
$SMTPPort = 2525
$SMTPUsername = "scriptsender@dcgla.net"
$EncryptedPasswordFile = "$mydir\scriptsender@dcgla.net.securestring"
$SecureStringPassword = Get-Content -Path $EncryptedPasswordFile | ConvertTo-SecureString
$EmailCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SMTPUsername,$SecureStringPassword
$MailTo = "monitoring@dcgla.com"
$MailFrom = "scriptsender@dcgla.net"
$mailAttachments = @()
$extraMailBodyInfo = Get-Content $confFile | Select-Object -Index 6
If ($auditType -match 3) {
    
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010

    }

Get-Date >> $logfile

If ($auditType -match 1) {
    
    Add-Content $logfile 'Beginning search of AD users'

    $usernamesraw = (Get-AdUser -filter * |Where-Object {($_.enabled -eq "True")}).name

    $usernamesfiltered = $usernamesraw | Where-Object {$_ -notmatch $regex}

    Add-Content $logfile 'Names of AD users: '
    $usernamesfiltered >> $logfile
    Add-Content $logfile 'Count of AD users: '
    $usernamesfiltered.count >> $logfile
    $usernamesfilteredCount = $usernamesfiltered.count
    $usernamesfiltered | Sort-Object > "$MyDir\logs\$(get-date -f yyyy-MM-dd)ADUsers.txt"
    $adUsersAttachment = "$MyDir\logs\$(get-date -f yyyy-MM-dd)ADUsers.txt"
    }

If ($auditType -match 2) {
    
    If ((Get-WmiObject -Class Win32_ComputerSystem).PartOfDomain -eq "True") {
        Add-Content $logfile "Server is part of a domain"
    
        Add-Content $logfile 'Beginning search of RDS Users.'
        
        $i = -1
        foreach ($groupname in $rdsGroup) {
        $i ++
        $rdsGroupSelected = $rdsGroup[$i]
    
        
        $rdsUsersRaw = (Get-ADGroupMember -Identity $rdsGroupSelected | Get-ADUser | Where-Object {($_.enabled -eq "True")}).name

        $rdsUsersFiltered = $rdsUsersRaw | Where-Object {$_ -notmatch $regex}
        Add-Content $logfile 'Names of RDS Users:'
        $rdsUsersFiltered >> $logfile
        Add-Content $logfile 'Count of RDS Users:'
        # This count is additive because some clients have multiple RDS groups. The next line adds each iteration of a count to the file.
        $rdsUsersFiltered |Sort-Object >> "$MyDir\logs\$(get-date -f yyyy-MM-dd)RemoteDesktopUsers.txt"
        $rdsUsersAttachment = "$MyDir\logs\$(get-date -f yyyy-MM-dd)RemoteDesktopUsers.txt"
        }
        }

    Else {
        Add-Content $logfile "Server is part of a workgroup"
        Add-Content $logfile "Beginning search for RDS Users."

        $rdsUsersRaw = (Get-LocalUser| Where-Object {($_.enabled -eq "True")}).name
        $rdsUsersFiltered = $rdsUsersRaw | Where-Object {$_ -notmatch $regex}
        Add-Content $logfile 'Names of RDS Users:'
        $rdsUsersFiltered >> $logfile
        Add-Content $logfile 'Count of RDS Users:'
        $rdsUsersFiltered.count >> $logfile
        $rdsUsersFilteredCount = $rdsUsersFiltered.count
        # This count is additive because some clients have multiple RDS groups. The next line adds each iteration of a count to the file.
        $rdsUsersFiltered |Sort-Object >> "$MyDir\logs\$(get-date -f yyyy-MM-dd)RemoteDesktopUsers.txt"
        $rdsUsersAttachment = "$MyDir\logs\$(get-date -f yyyy-MM-dd)RemoteDesktopUsers.txt"
        }
        $rdsUsersFilteredCount = 0 
        get-content $rdsUsersAttachment | foreach-object {$rdsUsersFilteredCount++}
    }
    


If ($auditType -match 3) {
    
    Add-Content $logfile 'Beginning search of Exchange mailboxes.'
    
    $mailAccountsRaw = ((Get-Mailbox -ResultSize Unlimited -WarningAction SilentlyContinue).emailaddresses | Where-Object PrefixString -CEQ SMTP| Select-Object SmtpAddress)
    $mailAccountsFiltered = $mailAccountsRaw | Where-Object {$_ -notmatch $regex}
    Add-Content $logfile 'Names of Exchange mailboxes:'
    $mailAccountsFiltered >> $logfile
    Add-Content $logfile 'Count of Exchange mailboxes:'
    $mailAccountsFiltered.count >> $logfile
    $mailAccountsFilteredCount = $mailAccountsFiltered.Count
    $mailAccountsFiltered | Sort-Object > "$MyDir\logs\$(get-date -f yyyy-MM-dd)ExchangeUsers.txt"
    $exchangeUsersAttachment = "$MyDir\logs\$(get-date -f yyyy-MM-dd)ExchangeUsers.txt"
    }

If ($auditType -match 5) {
    Add-Content $logfile "Beginning search of Office users."

     If ((Get-WmiObject -Class Win32_ComputerSystem).PartOfDomain -eq "True") {
        Add-Content $logfile "Server is part of a domain"
        
        $officeUsersRaw = (Get-ADGroupMember -Identity $officeGroup | Get-ADUser | Where-Object {($_.enabled -eq "True")}).name

        $officeUsersFiltered = $officeUsersRaw | Where-Object {$_ -notmatch $regex}
        Add-Content $logfile 'Names of Office Users:'
        $officeUsersFiltered >> $logfile
        Add-Content $logfile 'Count of Office Users:'
        $officeUsersFiltered.count >> $logfile
        $officeUsersFilteredCount = $officeUsersFiltered.count
        $officeUsersFiltered |Sort-Object > "$MyDir\logs\$(get-date -f yyyy-MM-dd)OfficeUsers.txt"
        $officeUsersAttachment = "$MyDir\logs\$(get-date -f yyyy-MM-dd)OfficeUsers.txt"
        }

    Else {
        Add-Content $logfile "Server is part of a workgroup"
        Add-Content $logfile "Beginning search for Office Users."

        $officeUsersRaw = (Get-LocalUser| Where-Object {($_.enabled -eq "True")}).name
        $officeUsersFiltered = $officeUsersRaw | Where-Object {$_ -notmatch $regex}
        Add-Content $logfile 'Names of Office Users:'
        $officeUsersFiltered >> $logfile
        Add-Content $logfile 'Count of Office Users:'
        $officeUsersFiltered.count >> $logfile
        $officeUsersFilteredCount = $officeUsersFiltered.count
        $officeUsersFiltered |Sort-Object > "$MyDir\logs\$(get-date -f yyyy-MM-dd)OfficeUsers.txt"
        $officeUsersAttachment = "$MyDir\logs\$(get-date -f yyyy-MM-dd)OfficeUsers.txt"
        }
    }



If ($auditType -match 6) {

    Add-Content $logfile 'Beginning search of Blaskguard Users.'

     If ((Get-WmiObject -Class Win32_ComputerSystem).PartOfDomain -eq "True") {
        Add-Content $logfile "Server is part of a domain"
        
        $BlaskGuardUsersRaw = (Get-ADGroupMember -Identity $BlaskGuardGroup | Get-ADUser | Where-Object {($_.enabled -eq "True")}).name

        $BlaskGuardUsersFiltered = $BlaskGuardUsersRaw | Where-Object {$_ -notmatch $regex}
        Add-Content $logfile 'Names of BlaskGuard Users:'
        $BlaskGuardUsersFiltered >> $logfile
        Add-Content $logfile 'Count of BlaskGuard Users:'
        $BlaskGuardUsersFiltered.count >> $logfile
        $BlaskGuardUsersFilteredCount = $BlaskGuardUsersFiltered.count
        $BlaskGuardUsersFiltered |Sort-Object > "$MyDir\logs\$(get-date -f yyyy-MM-dd)BlaskGuardUsers.txt"
        $BlaskGuardUsersAttachment = "$MyDir\logs\$(get-date -f yyyy-MM-dd)BlaskGuardUsers.txt"
        }

    Else {
        Add-Content $logfile "Server is part of a workgroup"
        Add-Content $logfile "Beginning search for BlaskGuard Users."

        $BlaskGuardUsersRaw = (Get-LocalUser| Where-Object {($_.enabled -eq "True")}).name
        $BlaskGuardUsersFiltered = $BlaskGuardUsersRaw | Where-Object {$_ -notmatch $regex}
        Add-Content $logfile 'Names of BlaskGuard Users:'
        $BlaskGuardUsersFiltered >> $logfile
        Add-Content $logfile 'Count of BlaskGuard Users:'
        $BlaskGuardUsersFiltered.count >> $logfile
        $BlaskGuardUsersFilteredCount = $BlaskGuardUsersFiltered.count
        $BlaskGuardUsersFiltered |Sort-Object > "$MyDir\logs\$(get-date -f yyyy-MM-dd)BlaskGuardUsers.txt"
        $BlaskGuardUsersAttachment = "$MyDir\logs\$(get-date -f yyyy-MM-dd)BlaskGuardUsers.txt"
        }
    }
    
If ($auditType -match 7) {
    
    Add-Content $logfile 'Beginning search of SSL VPN Users.'
        
    $sslvpnUsersRaw = (Get-ADGroupMember -Identity $sslvpnGroup | Get-ADUser | Where-Object {($_.enabled -eq "True")}).name

    $sslvpnUsersFiltered = $sslvpnUsersRaw | Where-Object {$_ -notmatch $regex}
    Add-Content $logfile 'Names of SSL VPN Users:'
    $sslvpnUsersFiltered >> $logfile
    Add-Content $logfile 'Count of SSL VPN Users:'
    $sslvpnUsersFiltered.count >> $logfile
    $sslvpnUsersFilteredCount = $sslvpnUsersFiltered.count
    $sslvpnUsersFiltered |Sort-Object > "$MyDir\logs\$(get-date -f yyyy-MM-dd)SSLVPNUsers.txt"
    $sslvpnUsersAttachment = "$MyDir\logs\$(get-date -f yyyy-MM-dd)SSLVPNUsers.txt"
    }

$MailSubject = "$companyContact, Please review $companyName's DCG PrivateCLOUD SPLA counts before $Month 15th"
$MailBody = "DCG strives to maintain an accurate active user list, as it pertains to your PrivateCLOUD server SPLA licensing counts on your server(s).  Attached is your current user count related Remote Desktop Services, MS Office, and SQL services on your PrivateCLOUD server.  This will be referenced in your upcoming Monthly Services invoice that will be emailed to you on the 15th of this month.

"
If ($auditType -match 1) {
$mailBody = $mailBody += "Current Active Directory Users: $usernamesfilteredcount

"
$mailAttachments = $mailAttachments += $adUsersAttachment
}

If ($auditType -match 2) {
$MailBody = $MailBody += "Current Remote Desktop Users: $rdsUsersFilteredcount

"
$mailAttachments = $mailAttachments += $rdsUsersAttachment
}

If ($auditType -match 3) {
    If ((Get-Content $confFile | Select-Object -Index 10) -match '^\d+$') {
        $exchangePlusUsersCount = Get-Content $confFile | Select-Object -Index 10
        #The next line compares the Exchange Plus count (specified on line 11 in the conf file) with the count of current mailboxes
        #If the mailbox count is less than the Plus count then the mailbox count simply becomes the Plus count
        #For example if all mailboxes at an org should be Plus licenses, then put a very high number on line 11 and it will always change the mailbox count to the Plus count
        If ($mailAccountsFilteredCount -le $exchangePlusUsersCount) {
            $MailBody = $MailBody += "Current Exchange Plus Users: $mailAccountsFilteredCount

        "
        } 

        #If the Plus count (specified on line 11 in the conf file) is less than the mailbox count then the Plus count gets subtracted from the mailbox count and both get reported separately.

        else {
        $MailBody = $MailBody += "Current Exchange Users: $($mailAccountsFilteredCount - $exchangePlusUsersCount)

Current Exchange Plus Users: $exchangePlusUsersCount

        "
        }
        }
        Else {
            $MailBody = $MailBody += "Current Exchange Users: $mailAccountsFilteredcount

            "
        }
    $mailAttachments = $mailAttachments += $exchangeUsersAttachment
}

If ($auditType -match 5) {
    $MailBody = $MailBody += "Current Office Users: $officeUsersFilteredCount

    "
    $mailAttachments = $mailAttachments += $officeUsersAttachment
    }

If ($auditType -match 6) {
    $MailBody = $MailBody += "Current BlaskGuard Users: $BlaskguardUsersFilteredCount

    "
    $mailAttachments = $mailAttachments += $BlaskguardUsersAttachment
    }

If ($auditType -match 7) {
    $MailBody = $MailBody += "Current SSL VPN Users: $sslvpnUsersFilteredCount
    
    "
    $mailAttachments = $mailAttachments += $sslvpnUsersAttachment
}

$MailBody = $MailBody += $extraMailBodyInfo

$MailBody = $MailBody += "

If any users should be removed from any of these lists, please contact DCG Technical Solutions before the 15th of the month.  We'll be sure to disable and remove any users so you will not incur any further licensing charges for them on your upcoming Monthly Services invoice.

Credits and refunds will not be issued after the 15th of this month.

Thank you very much for taking the time to review these reports with us

DCG Accounting

This message was sent from $env:COMPUTERNAME"

Send-MailMessage -From $MailFrom -To $MailTo -Subject $MailSubject -Body $MailBody -Port $SMTPPort -Credential $EmailCredential -Attachments $mailAttachments
