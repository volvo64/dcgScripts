[CmdletBinding()]

Param(
    [string]$O365Adminuser,

    [string]$strPassword,

    [string]$csvFile

)

If ($strPassword -eq $true) {$password = ConvertTo-SecureString -string $strpassword -AsPlainText -Force}

If (($password -ne "") -or ($strPassword -eq $true)) {$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $O365Adminuser, $password}
If (($csvFile -eq "" ) -or (!$csvFile)) {$csvFile = Read-Host "Path to CSV file for export"}

try {
    Get-MsolDomain -ErrorAction Stop > $null
}
catch {
    if ($cred -eq $null) {$cred = Get-Credential $O365Adminuser}
    Write-Output "Connecting to Office 365..."
    Connect-MsolService -Credential $cred
}

$tenants = Get-MsolPartnerContract -all | Select-Object tenantid, defaultdomainname, name
$TenantID = @{}

Foreach ($t in $tenants) {
    $TenantID.Add($t.Name, $t.TenantId)
}

$TenantID = $tenantID.GetEnumerator() | sort -Property Name
ForEach ($T in $TenantID.GetEnumerator()) {
    <#Write-Host "Name"
    Write-Host "$($T.Name) `n `n"
    #Write-Host "Name.Value"
    Echo "$($T.Name): $($T.Value)"
    #Write-Host "ST.Value"
    Write-Host "$($T.Value)"
    #Write-Host "Get-Msol"
    Get-MSolAccountSku -TenantID "$($T.Value)" #>
    #Export to CSV
    $products = @(Get-MSolAccountSku -TenantID "$($T.Value)")
    $products | Add-Member -MemberType NoteProperty -Name "Client" -Value $($t.Name)
    $products | select Client, skupartnumber, skuid, activeunits, consumedunits | Export-Csv $csvFile -Append -NoTypeInformation
}
