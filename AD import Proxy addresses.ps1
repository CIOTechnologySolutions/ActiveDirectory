Import-Csv "C:\users\cioadmin\desktop\alias.csv" | ForEach-Object { Set-AdUser -Identity $_.SamAccountName -Clear ProxyAddresses } 

Import-Csv "C:\users\cioadmin\desktop\alias.csv" | ForEach-Object {
$username = $_.samaccountname
$userproxy = $_.ProxyAddresses -split ','
Set-ADUser -Identity $username -Add @{proxyAddresses= $userproxy}
}