$MyDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$EncryptedPasswordFile = "$mydir\scriptsender@dcgla.net.securestring"
Read-Host -AsSecureString | ConvertFrom-SecureString | Out-File -FilePath $EncryptedPasswordFile
