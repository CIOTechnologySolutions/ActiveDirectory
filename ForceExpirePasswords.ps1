<#Windows PowerShell Code
.description
  Password expiration script - forceexpirepasswords.ps1
.Synopsis
  This script is designed to disable password never expires and force change on
  next login for any user provided in a CSV
.Parameter CSV
  Provide path to CSV File that contains a header of sAMAccountName
.example
  forceexpirepasswords.ps1 -csv c:\Scripts\adusers.csv
.NOTES
  Only the sAMAccountName will be read in any CSV provided
  By Kyle Elliott
  kelliott(at)ciotech(dot)us
  Provided as is, without warranty
#>

#Add parameter for CSV to be passed from PoSH
param {
  [Parameter(Mandatory=$true)]
  [string]$csv
}

#This is an AD script
Import-Module ActiveDirectory

#disable password never expires
Import-CSV $csv | ForEach-Object {
$samAccountName = $_."samAccountName"
 get-aduser -Identity $sAMAccountName |
 Set-ADUser $_ -PasswordNeverExpires $false
 }

#set to force password change
Import-Csv $csv | ForEach-Object {
 $samAccountName = $_."samAccountName"
Get-ADUser -Identity $samAccountName |
 Set-ADUser -ChangePasswordAtLogon:$True
}
