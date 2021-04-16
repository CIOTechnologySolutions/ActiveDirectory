Import-Module activedirectory 

$domain= get-addomain | Select-Object DistinguishedName
$disabledOU="OU=Disabled,$domain"

$OU = get-adorganizationalunit -filter * | Select-Object -ExpandProperty DistinguishedName
$CleanOU = $OU | where-object {($_ -notlike 'OU=Disabled,DC=*') -and ($_ -notlike 'OU=Users,DC=*') -and ($_ -notlike 'OU=service*')}

$cleanup= foreach ($CleanupOU in $CleanOU)
    {Get-ADUser -Filter {Enabled -eq $TRUE} -SearchBase $cleanupou  -Properties Name,SamAccountName,LastLogonDate,DistinguishedName | Where-Object {($_.DistinguishedName -notlike "*Disabled*") -and ($_.DistinguishedName -notlike "*Service*")} | Where-object {
    ($_.LastLogonDate -lt (Get-Date).AddDays(-30)) -and ($_.LastLogonDate -ne $NULL)
    } | Sort-object | Select-object Name,SamAccountName,LastLogonDate,DistinguishedName
}

 $cleanup | % {Get-ADUser -identity $_.name | Set-ADUser -Enabled $false -verbose} 

 $cleanup | % {get-aduser -identity $_.name | Move-ADObject -TargetPath $disabledOU}