<# create SPLA script#>

[CmdletBinding()]

Param(
    [Parameter(Mandatory=$true)]
    [string]$companyName,

    [Parameter(Mandatory=$true)]
    [string]$contactName,

    [Parameter(Mandatory=$true)]
    [string]$contactEmailAddress,

    [Parameter()]
    [int]$auditType,

    [Parameter()]
    [string]$adminAcctExclusions = "Administrator",

    [Parameter()]
    [string]$rdsGroup,

    [Parameter()]
    [string]$sslVpnGroup,

    [Parameter()]
    [string]$officeGroup,

    [Parameter()]
    [string]$blaskGuardGroup,

    [Parameter()]
    [string]$exchangePlusUsersCount,

    [Parameter()]
    [string]$insecureScriptSenderPassword
    )

$rootDir = "C:\DCG\SPLA"
$logsDir = "logs"
$companyFile = "$rootDir\company.conf"
$scriptRunner = "$rootDir\Run-SPLAScript.ps1"
$splaXMLfile = "$rootDir\spla.xml"
$scriptSender = "scriptsender@dcgla.net"
$EncryptedPasswordFile = "$rootDir\$scriptSender.securestring"

New-Item -ItemType Directory "$rootDir\$logsDir" -Force
New-Item -ItemType File $companyFile, $scriptRunner,$splaXMLfile -Force

If ($auditType -eq 0) {
    while ($auditType -lt 1) {
[int]$auditType = Read-Host -Prompt "Input audit type:
Audit types (line 2) explained. Use one number for every audit type (i.e. AD and Exchange will be 13):
1- AD
2- RDS
3- Exchange
4- SQL
5- Office
6- Blaskguard
7- SSL VPN

Input a number here"
}
}

If (($auditType -match 2) -and ([string]::IsNullOrWhiteSpace($rdsGroup))) {
    $rdsGroup = Read-Host "Define the RDS Group"
    }

If (($auditType -match 2) -and ([string]::IsNullOrWhiteSpace($rdsGroup))) {
    $sslVpnGroup = Read-Host "Define the SSLVPN Group"
    }

Add-Content $companyFile "$companyName
$auditType
$contactName
$contactEmailAddress
$adminAcctExclusions
$rdsGroup

$sslVpnGroup
$officeGroup
$blaskGuardGroup
$exchangePlusUsersCount"

If ([string]::IsNullOrWhiteSpace($insecureScriptSenderPassword)) {
    Read-Host -Prompt "Input password for $scriptSender" -AsSecureString | 
    ConvertFrom-SecureString | 
    Out-File -FilePath $EncryptedPasswordFile
    }
Else {
    $insecureScriptSenderPassword | 
    ConvertTo-SecureString -Force | 
    Out-File -FilePath $EncryptedPasswordFile
    }

Add-Content $scriptRunner '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ScriptFromGitHub = Invoke-WebRequest https://raw.githubusercontent.com/volvo64/dcgScripts/master/SPLA/SPLAScriptv1.ps1
Invoke-Expression $($ScriptFromGitHub.Content)'

Add-Content $splaXMLfile '<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2019-02-28T14:11:18</Date>
    <Author>dbreeden</Author>
    <URI>\Run SPLA</URI>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>2019-02-28T08:00:00</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByMonth>
        <DaysOfMonth>
          <Day>5</Day>
        </DaysOfMonth>
        <Months>
          <January />
          <February />
          <March />
          <April />
          <May />
          <June />
          <July />
          <August />
          <September />
          <October />
          <November />
          <December />
        </Months>
      </ScheduleByMonth>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>domain\user</UserId>
      <LogonType>Password</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>true</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>false</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>P3D</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>powershell.exe</Command>
      <Arguments>-WindowStyle Hidden -NonInteractive -Executionpolicy unrestricted -file C:\dcg\SPLA\Run-SPLAScript.ps1</Arguments>
      <WorkingDirectory>c:\dcg\spla</WorkingDirectory>
    </Exec>
  </Actions>
</Task>'

$adminPassword = $password = Read-Host -Prompt "Enter the local admin password" -AsSecureString
$UserName = $env:username
$Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $adminPassword
$Password = $Credentials.GetNetworkCredential().Password 

Register-ScheduledTask -Xml (get-content $splaXMLfile | out-string) -TaskName "Run SPLA" -User $UserName -Password $password -Force