#Script to Reset MyDocuments Folder permissions
$domainName = ([ADSI]'').name 
 Import-Module "PSCX" -ErrorAction Stop
 Set-Privilege (new-object Pscx.Interop.TokenPrivilege "SeRestorePrivilege", $true) #Necessary to set Owner Permissions
 Set-Privilege (new-object Pscx.Interop.TokenPrivilege "SeBackupPrivilege", $true) #Necessary to bypass Traverse Checking
 #Set-Privilege (new-object Pscx.Interop.TokenPrivilege "SeSecurityPrivilege", $true) #Optional if you want to manage auditing (SACL) on the objects
 Set-Privilege (new-object Pscx.Interop.TokenPrivilege "SeTakeOwnershipPrivilege", $true) #Necessary to override FilePermissions & take Ownership
 $Directorypath = "D:\Users\jtest" #locked user folders exist under here
 $LockedDirs = Get-ChildItem $Directorypath -force #get all of the locked directories.
 Foreach ($Locked in $LockedDirs) {
  Write-Host "Resetting Permissions for "$Locked.Fullname
 #######Take Ownership of the root directory
  $blankdirAcl = New-Object System.Security.AccessControl.DirectorySecurity
  $blankdirAcl.SetOwner([System.Security.Principal.NTAccount]'BUILTIN\Administrators')
  $Locked.SetAccessControl($blankdirAcl)
  
 ###################### Setup & apply correct folder permissions to the root user folder
  #Using recommendation from Ned Pyle's Ask Directory Services blog:
  #Automatic creation of user folders for home, roaming profile and redirected folders.
  $inherit = [system.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
  $propagation = [system.security.accesscontrol.PropagationFlags]"None"
  $fullrights = [System.Security.AccessControl.FileSystemRights]"FullControl"
  $allowrights = [System.Security.AccessControl.AccessControlType]"Allow"
  $DirACL = New-Object System.Security.AccessControl.DirectorySecurity
  #Administrators: Full Control
  $DirACL.AddAccessRule((new-object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators",$fullrights, $inherit, $propagation, "Allow")))
  #System: Full Control
  $DirACL.AddAccessRule((new-object System.Security.AccessControl.FileSystemAccessRule("NT AUTHORITY\SYSTEM",$fullrights, $inherit, $propagation, "Allow")))
  #Creator Owner: Full Control
  $DirACL.AddAccessRule((new-object System.Security.AccessControl.FileSystemAccessRule("CREATOR OWNER",$fullrights, $inherit, $propagation, "Allow")))
  #Useraccount: Full Control (ideally I would error check the existance of the user account in AD)
  #$DirACL.AddAccessRule((new-object System.Security.AccessControl.FileSystemAccessRule("$domainName\$Locked.name",$fullrights, $inherit, $propagation, "Allow")))
  $DirACL.AddAccessRule((new-object System.Security.AccessControl.FileSystemAccessRule("$domainName\$Locked",$fullrights, $inherit, $propagation, "Allow")))
  #Remove Inheritance from the root user folder
  $DirACL.SetAccessRuleProtection($True, $False) #SetAccessRuleProtection(block inheritance?, copy parent ACLs?)
  #Set permissions on User Directory
  Set-Acl -aclObject $DirACL -path $Locked.Fullname
  Write-Host "commencer" -NoNewLine
 ##############Restore admin access & then restore file/folder inheritance on all subitems
  #create a template ACL with inheritance re-enabled; this will be stamped on each subitem to re-establish the file structure with inherited ACLs only.
  #$NewOwner = New-Object System.Security.Principal.NTAccount("$domainName","$Locked.name") #ideally I would error check this.
  $NewOwner = New-Object System.Security.Principal.NTAccount("$domainName","$Locked") #ideally I would error check this.
  $subFileACL = New-Object System.Security.AccessControl.FileSecurity
  $subDirACL = New-Object System.Security.AccessControl.DirectorySecurity
  $subFileACL.SetOwner($NewOwner)
  $subDirACL.SetOwner($NewOwner)
  ######## Enable inheritance ($False) and not copy of parent ACLs ($False)
  $subFileACL.SetAccessRuleProtection($False, $False) #SetAccessRuleProtection(block inheritance?, copy parent ACLs?)
  $subDirACL.SetAccessRuleProtection($False, $False) #SetAccessRuleProtection(block inheritance?, copy parent ACLs?)
  #####loop through subitems
  $subdirs = Get-ChildItem -path $Locked.Fullname -force -recurse #force is necessary to get hidden files/folders
  foreach ($subitem in $subdirs) {
   #take ownership to insure ability to change permissions
   #Then set desired ACL
   if ($subitem.Attributes -match "Directory") {
    # New, blank Directory ACL with only Owner set
    $blankdirAcl = New-Object System.Security.AccessControl.DirectorySecurity
    $blankdirAcl.SetOwner([System.Security.Principal.NTAccount]'BUILTIN\Administrators')
    #Use SetAccessControl to reset Owner; Set-Acl will not work.
    $subitem.SetAccessControl($blankdirAcl)
    #At this point, Administrators have the ability to change the directory permissions
    Set-Acl -aclObject $subDirACL -path $subitem.Fullname -ErrorAction Stop
   } Else {
    # New, blank File ACL with only Owner set 
    $blankfileAcl = New-Object System.Security.AccessControl.FileSecurity
    $blankfileAcl.SetOwner([System.Security.Principal.NTAccount]'BUILTIN\Administrators')
    #Use SetAccessControl to reset Owner; Set-Acl will not work.
    $subitem.SetAccessControl($blankfileAcl)
    #At this point, Administrators have the ability to change the file permissions
    Set-Acl -aclObject $subFileACL -path $subitem.Fullname -ErrorAction Stop
   }
   Write-Host "." -NoNewline
  }
  Write-Host "fin."
 }
 Write-Host "Script Complete."