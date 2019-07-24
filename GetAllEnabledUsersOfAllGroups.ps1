Import-Module ActiveDirectory


$Groups = (Get-AdGroup -filter * | Where {$_.name -like "**"} | select name -ExpandProperty name)

$Table = @()

$Record = @{
  "Group Name" = ""
  "Name" = ""
  "Username" = ""
}


Foreach ($Group in $Groups) {

  $Arrayofmembers = Get-ADGroupMember -identity $Group -recursive | Get-ADUser -properties * | where {$_.enabled -eq $true} | select name,samaccountname

  foreach ($Member in $Arrayofmembers) {
    $Record."Group Name" = $Group
    $Record."Name" = $Member.name
    $Record."UserName" = $Member.samaccountname
    $objRecord = New-Object PSObject -property $Record
    $Table += $objrecord

  }
}
$Table | export-csv "C:\temp\SecurityGroups.csv" -NoTypeInformation
