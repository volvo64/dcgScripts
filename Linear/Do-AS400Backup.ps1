[CmdletBinding()]

Param(

    [string]$fileRoot,

    [int]$daysToCleanUp

)

$now = get-date -UFormat %Y%m%d

If ($fileRoot -eq "") { $fileRoot = "path\to\root"}

If (!$daysToCleanUp) {$daysToCleanUp = "3"}

$errorLog = "$fileRoot\errorlog.log"

$folders = Get-ChildItem -Path $fileRoot -Directory; Get-ChildItem -Path $fileRoot -Name *.script

Write-Host $folders

#clean up old folders here

foreach ($f in $folders) {
    If ((($f.CreationTime  | Get-Date -UFormat %Y%m%d) -lt (Get-Date).AddDays(-$daysToCleanUp).ToString("yyyyMMdd")) -eq $true) {
        Remove-Item $f.FullName -Confirm:$false -Force -Recurse -ErrorAction SilentlyContinue
    }
}

#remove old scripts
Remove-Item "$fileRoot\*.script" -Confirm:$false -Force -Recurse -ErrorAction SilentlyContinue

New-Item -ItemType Directory -Path $fileRoot -Name $now

$destPath = "$fileRoot\$now"

If (!$destPath) {Add-Content $errorLog -Value "$now Destination directory doesn't exist, exiting"; break}

$scriptFile = "$fileRoot\$now.script"

Add-Content $scriptFile "user
username
password
binary
cd /remotedir
lcd `"$destPath`"
prompt
mget *
quit
"

FTP -n -s:$scriptFile ftpserver

