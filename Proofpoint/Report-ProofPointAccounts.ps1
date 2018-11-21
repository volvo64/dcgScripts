

$RESTserver = "www.server.com" #no http
$RESTUser = ""
$RESTpw = ""

$BaseURL = "https://" + $RESTserver
$Header = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$Header.Add("Authorization", "Basic")
$header.Add("X-user", $RESTUser)
$Header.Add("X-password", $RESTpw)
$Header.Add("domain", "domain.com")
$Type = "application/json"

$json = Invoke-RestMethod  -Uri "http://www.server.com/api/org/domain" -Headers $Header -Method GET -ContentType $Type
ConvertFrom-Json $json | write-host