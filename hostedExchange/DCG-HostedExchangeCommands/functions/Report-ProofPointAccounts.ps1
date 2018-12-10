<# Written November 2018 by Damon Breeden

Requires ImportExcel module which needs Powershell 5 to install easily

Reports all Proofpoint accounts to an Excel document #>

$RESTserver = "server.com/api/v1" #no http
$RESTUser = "user@domain.com"

#fix this password
$RESTpw = "!"

$BaseURL = "https://" + $RESTserver
$Header = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$Header.Add("Authorization", "Basic")
$header.Add("X-user", $RESTUser)
$Header.Add("X-password", $RESTpw)
$Header.Add("domain", "dcgla.com")
$Type = "application/json"
$excel = ".\test.xlsx"
$excelendusers = ".\endusers.xlsx"
$logfile = "~\Desktop\logfile.log"

Invoke-RestMethod  -Uri "$BaseURL/orgs/dcgla.com/orgs" -Headers $Header -Method GET -ContentType $Type | select -ExpandProperty orgs
$companies = Invoke-RestMethod  -Uri "$BaseURL/orgs/dcgla.com/orgs" -Headers $Header -Method GET -ContentType $Type | select -ExpandProperty orgs

foreach ($c in $companies) {
    Write-Host $c.primary_domain
    $worksheetname = $c.name[0..30] -join ""
    $users = Invoke-RestMethod  -Uri "$BaseURL/orgs/$($c.primary_domain)/users" -Headers $Header -Method GET -ContentType $Type | select -ExpandProperty users
    #$users | Export-Excel -Path $excel -WorksheetName $worksheetname
    
    
    Foreach ($u in $users) {
        If (($u.psobject.properties.value[2] -eq "end_user") -or ($u.psobject.properties.value[2] -eq "silent_user")) {
            write-host $u.primary_email; Write-Host $u.type
            #this test doesn't really work, since it fails against mailboxes that don't exist.
            $r = Get-Recipient $u.primary_email
            If ($?) {
                If ($r.RecipientType -ne "UserMailbox") {
                    Add-Content $logfile "Mailbox $($u.primary_email) should be changed"
                    Write-Host "Mailbox $($u.primary_email) should be changed"
                    <# $body = '{
                    "type": "functional_account"
                    }'
                    Invoke-Restmethod -uri "$baseurl/orgs/$($c.primary_domain)/users/$($u.primary_email)" -Headers $Header -Method Put -Body $body -ContentType $Type  #>
                    }
                }
            #$u | Export-Excel -Path $excelendusers -WorksheetName $worksheetname -Append
            }
        }

        
    }
