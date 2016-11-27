###############################################################################
# Fix COM errors
# Author: Kit Menke
# Version 1.0 11/6/2016
# Run this script as administrator.
###############################################################################

Param(
  [string]$CLSID,
  [string]$APPID
)

# TODO: make these parameters?
#$CLSID = "{D63B10C5-BB46-4990-A94F-E40B9D520160}"
#$APPID = "{9CA88EE3-ACB7-47C8-AFC4-AB702511C276}"
$CLSID = "{8D8F4F83-3594-4F07-8369-FC3C3CAE4919}"
$APPID = "{F72671A9-012C-4725-9D2F-2A4D32D65169}"
$LOG_ONLY = $true

# Checkpoint-Computer -Description "Fix DCOM errors script"

# Source the script for enabling permissions
. .\enable-privilege.ps1
# Source the script for taking ownership of the keys
. .\Set-RegistryKeyPermissions.ps1
# Source the script for changing DCOM Application permissions
. .\Set-DcomAppPermissions.ps1

# Adjust the permissions for these keys
Write-Host "CLSID is $CLSID"
Write-Host "APPID is $APPID"

# to check your priviledges:
# whoami /priv

# To change the owner of a registry key the shell needs additional priviledges
# http://stackoverflow.com/questions/6622124/why-does-set-acl-on-the-drive-root-try-to-set-ownership-of-the-object
try {
  enable-privilege SeTakeOwnershipPrivilege
  enable-privilege SeBackupPrivilege
  enable-privilege SeRestorePrivilege
} catch [System.Management.Automation.MethodInvocationException] {
  Write-Host "Unable to enable priviledges, did you run this script as administrator?"
  exit 1
}

# Change permissions of two registry keys and all child keys:
# HKEY_Classes_Root\CLSID\*CLSID*
# HKEY_LocalMachine\Software\Classes\AppID\*APPID*
# BUILTIN\Administrators (S-1-5-32-544) will be the owner and will have Full Control permissions
if ($LOG_ONLY) {
  Print-RegistryKeyPermissions "HKCR" "CLSID\$CLSID" "S-1-5-32-544" $true
  Print-RegistryKeyPermissions "HKLM" "Software\Classes\AppID\$APPID" "S-1-5-32-544" $true
} else {
  #Set-RegistryKeyPermissions "HKCR" "CLSID\$CLSID" "S-1-5-32-544" $true
  #Set-RegistryKeyPermissions "HKLM" "Software\Classes\AppID\$APPID" "S-1-5-32-544" $true
}

# Add NT AUTHORITY\SYSTEM with Local Launch and Local Activation permissions
Set-DcomAppPermissions $APPID "NT AUTHORITY" "SYSTEM"

Write-Host "Script complete"