<#Windows PowerShell Code
.description
  Finds all disabled users groups
.Synopsis
  Used to determine what users who had previously been disabled
  that weren't properly removed from groups
.Parameter Path
  path to place CSV
.Example

.NOTES
  By Kyle Elliott
  kelliott(at)ciotech(dot)us
  Provided as is, without warranty
  #>
Param(
  [Parameter(ParameterSetName="path",Mandatory=$True)]
  [string]$path=""
)

  Import-Module Activedirectory

  Get-ADUser -Credential $credentials -Filter * -Properties DisplayName,EmailAddress,memberof,DistinguishedName,Enabled | %  {
    New-Object PSObject -Property @{
  	UserName = $_.DisplayName
      EmailAddress = $_.EmailAddress
      DistinguishedName = $_.DistinguishedName
      Enabled = $_.Enabled
  # Deliminates the document for easy copy and paste using ";" as the delimiter. Incredibly useful for Copy & Paste of group memberships to new hire employees.
  	Groups = ($_.memberof | Get-ADGroup | Select -ExpandProperty Name) -join ";"
  	}
  # The export path is variable change to desired location on domain controller or end user computer.
  } | Select UserName,EmailAddress,@{l='OU';e={$_.DistinguishedName.split(',')[1].split('=')[1]}},Groups,Enabled | Export-Csv  $path –NTI
