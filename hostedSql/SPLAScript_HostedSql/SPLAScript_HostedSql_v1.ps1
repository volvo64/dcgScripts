$MyDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Month = Get-Date -UFormat %B
$time = Get-Date
$logfile = "$myDir\logs\logfile.log"
Add-Content $logfile "Script ran on $(get-date -f yyyy-MM-dd)"

$filterednames = @("mimecast","guest","LDAP","vmware","dss","opendns","sp admin","dcg","qbdataservice","sql","st_bernard","hosted","ldapadmin","spadmin","test","noc","st. bernard","st bernard","managed care","bbadmin","besadmin","compliance","discovery","rmmscan","healthmailbox","sharepoint","windows sbs","qbdata","noc_helpdesk","appassure","support","scanner","ftp","app assure","aspnet","Dependable Computer Guys","efax","exchange","INSTALR","IUSR","IWAM","NAV","Quick Books")
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

Function Show-DatabasesOnServer ([string]$Server)
{
$Srv = New-Object ('Microsoft.SqlServer.Management.SMO.Server') $Server
Write-Host " The Databases on $Server Are As Follows:"
$Srv.Databases | Select Name
}

$companies = Get-Content $MyDir\companies.conf
$companies = $companies | sort

foreach $company in $companies {
    $confFile = $MyDir\$company.conf
    $selectedCompany = Get-Content $confFile | Select-Object -Index 0
    $selectedCompanySqlName = Get-Content $confFile | Select-Object -Index 1
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
    $Server = New-Object ('Microsoft.SqlServer.Management.Smo.Server') $selectedCompanySqlName


    $SQLAccounts = @()
    ForEach ($User in $Server.Logins)
    {
	    If ($User.IsDisabled -like "False" -AND $User.Name -notlike "DYN*" -AND $User.Name -notlike "HOSTED1*" -AND $User.Name -notlike "NT AUTHORITY*" -AND $User.Name -notlike "NT Service*" -AND $User.Name -notlike "SQLService*" -AND $User.Name -notlike "SA" -AND $User.Name -notlike "SAMW" -AND $User.Name -notlike "MISSIONWELL\*" -AND $User.Name -notlike "CORP\*")
	        {
		        $SQLAccounts = $SQLAccounts += $User.Name
	    }
    }



    

    $SQLAccounts
    $SQLAccounts.Count

    $SQLAccounts | Out-File "$MyDir\logs\$(get-date -f yyyy-MM-dd)SPLA.SQL.$selectedCompany.Users.txt"
    $SQLAccounts.Count | Out-File "$MyDir\logs\$(get-date -f yyyy-MM-dd)SPLA.SQL.$selectedCompany.Count.txt"
    }