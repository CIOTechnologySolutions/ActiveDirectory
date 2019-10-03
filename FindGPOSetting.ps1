<#Windows PowerShell Code
.description
  hunt for a GPO setting across all GPOs
.Synopsis
  hunt for a GPO setting across all GPOs
.Parameter string
  add GPO entry to look For

.Example
  gpohunt.ps1 -string "bitlocker recovery information"
.NOTES
  By Kyle Elliott
  kelliott(at)ciotech(dot)us
  Provided as is, without warranty
  #>

Param(
  [Parameter(Mandatory=$TRUE)]
  [string]$String
)

$NearestDC = (Get-ADDomainController -Discover -NextClosestSite).Name

#Get a list of GPOs from the domain
$GPOs = Get-GPO -All -Server $NearestDC | sort DisplayName

#Go through each Object and check its XML against $String
Foreach ($GPO in $GPOs)  {

  #Get Current GPO Report (XML)
  $CurrentGPOReport = Get-GPOReport -Guid $GPO.Id -ReportType Xml -Server $NearestDC
  If ($CurrentGPOReport -match $String)  {
	$Output = "A Group Policy matching ""$($String)"" has been found:"
	$Output += "- GPO Name: $($GPO.DisplayName)"
	$Output += "- GPO Status: $($GPO.GpoStatus)"
  }
}
Write-Output $output
