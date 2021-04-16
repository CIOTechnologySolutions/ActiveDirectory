
$drives= [System.IO.DriveInfo]::GetDrives() |where {$_.driveType -eq "network"} | select name
$drives | ForEach-Object {($_.Name -split "\\")[0]}
 foreach ($drive in $cleaned)
  {
 net use $drive /delete /y
 }

gpupdate /force



$drives=GET-WMIOBJECT -query "SELECT * from win32_logicaldisk where DriveType='4'" -computername "@server@" | select deviceID
foreach ($drive in $drives){net use $drive /delete /y}
gpupdate /force