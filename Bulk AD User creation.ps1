<#Windows PowerShell Code
.description
  BulkAdUserCreation.ps1
.Synopsis
  Creates a batch of AD users based on a CSV Provided.
.Parameter CSV
  Pass a CSV file with the following headers
  username,password,Firstname,Lastname,EmailAddress,OU,UPN
.Example
  BulkAdUserCreation.ps1 -csv c:\temp\users.csv
.NOTES
  By Kyle Elliott
  kelliott(at)ciotech(dot)us
  Provided as is, without warranty
  #>
﻿# Import active directory module for running AD cmdlets
Param(
  [Parameter(Mandatory=$True)]
  [string]$csv
)


Import-Module activedirectory

#Store the data from ADUsers.csv in the $ADUsers variable
$ADUsers = Import-csv $csv

#Loop through each row containing user details in the CSV file
foreach ($User in $ADUsers)
{
	#Read user data from each field in each row and assign the data to a variable as below

	$Username 	= $User.username
	$Password 	= $User.Password
	$Firstname 	= $User.FirstName
	$Lastname 	= $User.LastName
    $Email      = $User.EmailAddress
	$OU 		= $User.OU #This field refers to the OU the user account is to be created in
    $UPN        = $User.UPN

	#Check to see if the user already exists in AD
	if (Get-ADUser -F {SamAccountName -eq $Username})
	{
		 #If user does exist, give a warning
		 Write-Warning "A user account with username $Username already exist in Active Directory."
	}
	else
	{
		#User does not exist then proceed to create the new user account

        #Account will be created in the OU provided by the $OU variable read from the CSV file
		New-ADUser `
            -SamAccountName $Username `
            -UserPrincipalName $UPN `
            -Name "$Firstname $Lastname" `
            -GivenName $Firstname `
            -Surname $Lastname `
            -Enabled $True `
            -DisplayName "$Firstname $Lastname" `
            -Path $OU `
            -AccountPassword (convertto-securestring $Password -AsPlainText -Force) `
            -EmailAddress $Email
	}
}
