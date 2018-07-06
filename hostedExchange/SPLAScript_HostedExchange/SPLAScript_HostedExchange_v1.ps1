Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
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

Add-Content $logfile 'Beginning search of Exchange mailboxes.'

#Updated code to ignore AD accounts that are disabled, sorted by Name
$UsersGroups = Get-Mailbox -Filter {(userAccountControl -ne '514') -AND (userAccountControl -ne '546') -AND (userAccountControl -ne '66050') `
	-AND (userAccountControl -ne '66082')} | Group-Object -Property:OrganizationalUnit | Select-Object Name,Count

#Added to count shared mailboxes
$UsersGroupsShared = Get-Mailbox -Filter {IsLinked -eq $true} | Group-Object -Property:OrganizationalUnit | Select-Object Name,Count

#Combine both arrays
$UsersGroups = $UsersGroups += $UsersGroupsShared
$UsersGroups = $UsersGroups | Sort-Object Name

#Array to hold Accounts Count data
$UserAccountsC = @()

#Loop through each user in the array
ForEach ($User in $UsersGroups)
{
	#Strip local domain name for easier reading
	$User.Name = $User.Name -replace "hosted1.local/Hosting/", ""
	
	#Capture the user name if the User is in an OU with " - Users" in the name
	If ($User.Name -match " - Users" -And $User.Name -notmatch "dcg")
	{
		$UserAccountsC = $UserAccountsC += $User
	}
}

#Export the User Account Count - number of users in each domain - information to a CSV file
$UserAccountsC | Export-CSV "$MyDir\logs\$(get-date -f yyyy-MM-dd)UserAccountCount.csv" -nti
$UserAccountsCountAttachment = "$MyDir\logs\$(get-date -f yyyy-MM-dd)UserAccountCount.csv"




#Updated code to ignore AD accounts that are disabled, sorted by Name
$UsersGroups = Get-Mailbox -Filter {(userAccountControl -ne '514') -AND (userAccountControl -ne '546') -AND (userAccountControl -ne '66050') `
	-AND (userAccountControl -ne '66082')} | Group-Object -Property:OrganizationalUnit | Select-Object Name,Group | Sort-Object Name

#Added to count shared mailboxes
$UsersGroupsShared = Get-Mailbox -Filter {IsLinked -eq $true} | Group-Object -Property:OrganizationalUnit | Select-Object Name,Group

#$UsersGroupsShared
#Read-Host 'Shared - Please press ENTER' | Out-Null

#Combine both arrays
$UsersGroups = $UsersGroups += $UsersGroupsShared
$UsersGroups = $UsersGroups | Sort-Object Name

#$UsersGroups
#Read-Host 'All - Please press ENTER' | Out-Null

#Array to hold Accounts Breakdown data
$UserAccountsB = @()

#Loop through each user in the array
ForEach ($User in $UsersGroups)
{
	ForEach ($Name in $User.Group)
		{ 
		If ($Name.Name -notmatch "test" -And $Name.Name -notmatch "administrator" -And $Name.Name -notmatch "dcg")
		{
		#Create a new custom object to hold our result.
		$UserObject = New-Object PSObject

		#Strip local domain name for easier reading
		$User.Name = $User.Name -replace "hosted1.local/Hosting/", ""
		#Split Company Name and OU
		$SplitString = $User.Name -split '/'
		$Company = $SplitString[0]
		$OU = $SplitString[1]

		#Add data to $UserObject
		$UserObject | add-member -membertype NoteProperty -name "Company" -Value $Company
		$UserObject | add-member -membertype NoteProperty -name "OU" -Value $OU
		$UserObject | add-member -membertype NoteProperty -name "User" -Value $Name.Name

		#Save the current $UserObject by appending it to $UserAccountsB ( += means append a new element to ‘me’)
		$UserAccountsB += $UserObject
		}
		}
}

#Export the User Account Breakdown - information to a CSV file
$UserAccountsB | Sort-Object -Property Company, User | Export-csv "$MyDir\logs\$(get-date -f yyyy-MM-dd)UserAccountsBreakdown.csv" -nti
$UserAccountsBreakdownAttachment = "$MyDir\logs\$(get-date -f yyyy-MM-dd)UserAccountsBreakdown.csv"

Add-Content $logfile "Sending email now"

$MailSubject = "$Month's DCG Hosted Exchange audit"
$MailBody = "$Month's SPLA Audit Info

Attached are the CSV files with the customers and number of Hosted Exchange accounts and the breakdown of the individual accounts that has the users for each company.

This message was sent from $env:COMPUTERNAME"

$mailAttachments = $mailAttachments += $UserAccountsCountAttachment
$mailAttachments = $mailAttachments += $UserAccountsBreakdownAttachment

Send-MailMessage -From $MailFrom -To $MailTo -Subject $MailSubject -Body $MailBody -Port $SMTPPort -Credential $EmailCredential -Attachments $mailAttachments