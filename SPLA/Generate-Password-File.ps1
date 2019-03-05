$MyDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$EncryptedPasswordFile = "$mydir\scriptsender@dcgla.net.securestring"
Read-Host "Type the password for scriptsender@dcgla" -AsSecureString | ConvertFrom-SecureString | Out-File -FilePath $EncryptedPasswordFile
