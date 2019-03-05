$info = Get-Content ".conf"

$name = $info[0]
$custAttr1 = $info[1]
$dn =  $info[2]
$emailDomain = $info[3]
$emailPrefix = $info[4]
$emailDatabase = $info[5]
$emailABP = $info[6]
$publicFolderMailbox = $info[7]
$contactsOU = $info[8]
$groupsOU = $info[9]
$roleAssignmentPolicy = $info[14]

$companyData=@"
{
    "company": {
        "name": "$name",
        "CustomAttribute1":"$custAttr1",
        "usersOU": "$dn",
        "emailDomain": "$emailDomain",
        "emailPrefix": "$emailPrefix",
        "database": "$emailDatabase",
        "abp": "$emailABP",
        "pfMailbox": "$publicFolderMailbox",
        "contactsOU": "$contactsOU",
        "groupsOU": "$groupsOU"
    }
}
"@