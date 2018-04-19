Start-Sleep -s 900
$vmwareServices = Get-Service | Where ({$_.DisplayName -like "VMware*"} -and {$_.Status -ne "Running"} -and {$_.StartMode -eq "Auto"})
# Write-Host $vmwareServices.length
$i = 0
while ($i -lt $vmwareServices.lenth) {
    $s = $vmwareServices[$i]
    Start-Service -DisplayName $s
    $i = $i + 1
    }