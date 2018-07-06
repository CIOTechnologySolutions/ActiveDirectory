#ChangePermissions.ps1
#By Kyle Elliott
#kelliott@ciotech.us
# CACLS rights are
# F = FullControl
# C = Change
# R = Readonly
# W = Write
# You must have powershell community extensions installed prior to running 
# These extensions require a reboot
# The powershell extensions must match the version of powershell you're running
<#
.Parameter username
    Set the username you would like to grant access to
.parameter StartingDir
    Enter the full path of the directory you want to change access to
.Parameter permissions
    CACLS rights are
    F = FullControl
    C = Change
    R = Readonly
    W = Write
.NOTES
    You must have powershell community extensions installed prior to running
    These extensions require a reboot
    The powershell extensions must match the version of powershell you're running
    Provided without any warranty and entirely as is.
    By Kyle Elliott
    kelliott(at)ciotech(dot)us
#>
{
Param(
    [Parameter(Mandatory=$true,helpmessage="Enter the username you would like to grant access")]
    [string]$Username,
    [Parameter(Mandatory=$True,helpmessage="Enter directory you want to Change")]
    [string]$StartingDir,
    [Parameter(Mandatory=$true,helpmessage="Enter F for full, C for change, R for Read only, and W for write")]
    [string]$Permissions
) #end Param

$domainName = ([ADSI]'').name
$Principal="$domainname\$username"

#Import PSCX - Powershell Community Extensions
Import-Module "PSCX" -ErrorAction Stop

#Set required privliges to bypass NTFS security
Set-Privilege (new-object Pscx.Interop.TokenPrivilege "SeRestorePrivilege", $true) #Necessary to set Owner Permissions
Set-Privilege (new-object Pscx.Interop.TokenPrivilege "SeBackupPrivilege", $true) #Necessary to bypass Traverse Checking
#Set-Privilege (new-object Pscx.Interop.TokenPrivilege "SeSecurityPrivilege", $true) #Optional if you want to manage auditing (SACL) on the objects
Set-Privilege (new-object Pscx.Interop.TokenPrivilege "SeTakeOwnershipPrivilege", $true) #Necessary to override FilePermissions & take Ownership


$Verify=Read-Host `n "You are about to change permissions on all" `
"files starting at"$StartingDir.ToUpper() `n "for security"`
"principal"$Principal.ToUpper() `
"with new right of"$Permission.ToUpper()"."`n `
"Do you want to continue? [Y,N]"
if ($Verify -eq "Y") {
    foreach ($file in $(Get-ChildItem $StartingDir -recurse)) {
        #display filename and old permissions
        write-Host -foregroundcolor Yellow $file.FullName
        #uncomment next line only if you want to see old permissions
        #CACLS $file.FullName
        #ADD new permission with CACLS
        CACLS $file.FullName /E /P "${Principal}:${Permission}" >$NULL
        #display new permissions
        Write-Host -foregroundcolor Green "New Permissions"
        CACLS $file.FullName
    }
}
