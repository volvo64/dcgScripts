[CmdletBinding()]

Param (
    [string]$searchDir = "\\server\data",
    [string]$logDir = "C:\permsReports",
    [string]$reportFile = "PermissionsReport",
    [string]$mailServer = "mail.server.com",
    [int]$mailPort = 2525,
    [string]$mailto = "monitoring@domain.com",
    [string]$encryptedPasswordFile = "C:\scriptsender@domain.net.securestring",
    [string]$SMTPUsername = "scriptsender@domain.net"
    )

$now = Get-Date -f yyyy-MM-dd
$report = "$logDir\$now$reportFile.xlsx"
New-Item -ItemType Directory -Path $logDir -Force


Import-Module ActiveDirectory
Import-Module ImportExcel
Import-Module NTFSSecurity

$excluded = @("Guest","DefaultAccount","krbtgt","LDAP Admin","Proweaver","Net Wrix",
    "DUO Auth User","Salesforce Sync","Administrators","Users","Guests","Performance Log Users","IIS_IUSRS",
    "System Managed Accounts Group","Domain Computers","Domain Controllers","Schema Admins","Enterprise Admins","Domain Admins",
    "Domain Users","Domain Guests","Group Policy Creator Owners","Pre-Windows 2000 Compatible Access",
    "Windows Authorization Access Group","Terminal Server License Servers","Denied RODC Password Replication Group","SSLVPN",
    "Redirected Folders","HelpLibraryUpdaters",
    "Netwrix Auditor Administrators","Netwrix Auditor Client Users","SQLServer2005","SQLRUserGroupNETWRIX","Print Operators",
    "Backup Operators","Replicator","Network Configuration Operators","Cryptographic Operators","Event Log Readers",
    "Certificate Service DCOM Access","RDS Remote Access Servers","RDS Endpoint Servers","RDS Management Servers",
    "Access Control Assistance Operators","Cert Publishers","RAS and IAS Servers","Server Operators","Account Operators",
    "Incoming Forest Trust Builders","Allowed RODC Password Replication Group","Key Admins","Enterprise Key Admins","DnsAdmins",
    "DnsUpdateProxy","CREATOR OWNER","SYSTEM")

$regex = "(" + ($excluded -join "|") + ")"

# I'm certain this can be filtered left but I don't know enough about how to to get it done...
$users = (Get-ADUser -Filter * | Where-Object {($_.enabled -eq "True")}) | Where-Object {$_ -notmatch $regex} | Sort-Object Name
$groups = (Get-ADGroup -Filter *) | Where-Object {$_ -notmatch $regex} | Sort-Object Name

foreach ($u in $users) {
    $userGroups = Get-ADPrincipalGroupMembership $u | Where-Object {$_ -notmatch $regex}
    Foreach ($ug in $userGroups) {
        $ExcelData = [PSCustomObject][Ordered]@{
            Name = $u.Name
            Group = $ug.name
            }
        $ExcelData | Export-Excel $report -WorksheetName "User group memberships" -Append -AutoSize
        If (($ug -notlike "*grp*") -and ($ug -notlike "*imporperly named Share Security Group*") -and ($ug -notlike "*improperly named Securtiy Group*")) {
            $ExcelData | Export-Excel $report -WorksheetName "Users in inappropriate group" -Append -AutoSize -KillExcel
            }
        }
    }
  
foreach ($g in $groups) {
    $groupMembers = (Get-ADGroupMember $g -Recursive | Select-Object -ExpandProperty SamAccountname | Get-ADUser -ErrorAction SilentlyContinue | Where-Object {($_.enabled -eq "True")}) | Where-Object {$_ -notmatch $regex}
    foreach ($gm in $groupMembers) {
        $ExcelData = [PSCustomObject][Ordered]@{
            Name = $g.Name
            Member = $gm.Name
            }
        $ExcelData | Export-Excel $report -WorksheetName "Groups and Ultimate Members" -Append -AutoSize -KillExcel
        }
    }

$directories = Get-ChildItem -LiteralPath $searchDir -Directory -Recurse
foreach ($d in $directories) {
    $directlyAppliedUser = Get-NTFSAccess -Path $d.Fullname -ExcludeInherited | Where-Object AccountType -NE group | Where-Object {$_.Account -notmatch $regex}
    foreach ($dau in $directlyAppliedUser) {
        $ExcelData = [PSCustomObject][Ordered]@{
            Path = $dau.Fullname
            Account = $dau.Account
            AccessRights = $dau.AccessRights
            }
        $ExcelData | Export-Excel $report -WorksheetName "Directly applied ACEs" -Append -AutoSize -KillExcel
        }
    }

foreach ($d in $directories) {
    $acl = Get-NTFSAccess $d.FullName -ExcludeInherited | Where-Object AccountType -EQ group | Where-Object {$_.Account -notmatch $regex}
    foreach ($ace in $acl) {
        $aceAccount = ($ace.Account).ToString().Split("_")[-2]
        $acePath = $ace.FullName.Split("\")[-1]
        If (($aceAccount -ne $acePath) -and (!(($aceAccount -eq "domain\improperly named Securtiy Group") -and ($acePath -eq "\\server\data")))   ) {
            $ExcelData = [PSCustomObject][Ordered]@{
            AccessControlEntry = $ace.Account
            Path = $ace.FullName
            }
        $ExcelData | Export-Excel $report -WorksheetName "Possibly mismatched ACEs" -Append -AutoSize -KillExcel
        }
        }
    }

$SecureStringPassword = Get-Content -Path $EncryptedPasswordFile | ConvertTo-SecureString
$EmailCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SMTPUsername, $SecureStringPassword

$mailSubject = "Company shared drive report $now"

$mailBody = "Attached is the S drive report for Company, run on $now.

Please review the attached file and address any entries on the `"Possibly mismatched ACEs`" and the `"Directly applied ACEs`" workbooks.

Dispatch, if any entries are on the two above mentioned sheets, please find the original engineer who did the work, reassign the ticket back to the engineer and have them do the work properly.

Have the engineer review the OneNote article titled `"How to assign S drive permissions`" under the Company page for reference.

This message was sent from $env:COMPUTERNAME"

Send-MailMessage -From $SMTPUsername -To $mailTo -Subject $mailSubject -Attachments $report -Body $mailBody -SmtpServer $mailServer -Port $mailPort -Credential $EmailCredential