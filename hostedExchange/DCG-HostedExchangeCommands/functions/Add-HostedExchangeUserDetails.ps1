[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True,Position=1)]
        [string]$csvFile

    
    )


Start-Transcript
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto

$csv = Import-Csv $csvFile


foreach ($n in $csv) {

    $params = @{}
    If ($n.samaccountName -eq "") {Write-Host "SamAccountName is a required field"; break}
    If ($n.FirstName -ne "") {$params.GivenName = $n.firstName}
    If ($n.LastName -ne "") {$params.Surname = $n.LastName}
    If ($n.DisplayName -ne "") {$params.DisplayName = $n.DisplayName}
    If ($n.Title -ne "") {$params.Title = $n.Title}
    If ($n.PhoneNumber -ne "") {$params.OfficePhone = $n.PhoneNumber}
    If ($n.Department -ne "") {$params.Department = $n.Department}

    $str = $params | Out-String
    Write-Host $str
    Set-ADUser -Identity $n.SamAccountName -SamAccountName $n.SamAccountName @params
}