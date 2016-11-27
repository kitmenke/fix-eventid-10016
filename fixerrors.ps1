###############################################################################
# Fix COM errors
# Author: Kit Menke
# Version 1.0 11/6/2016
###############################################################################
#
# Overview:
# This script is designed to fix a COM error that occurs repeatedly. The error
# is logged in the event log:
###############################################################################
#    The application-specific permission settings do not grant Local Activation permission for the COM Server application with CLSID
#    {D63B10C5-BB46-4990-A94F-E40B9D520160}
#     and APPID
#    {9CA88EE3-ACB7-47C8-AFC4-AB702511C276}
#     to the user NT AUTHORITY\SYSTEM SID (S-1-5-18) from address LocalHost (Using LRPC) running in the application container Unavailable SID (Unavailable). This security permission can be modified
#    using the Component Services administrative tool.
###############################################################################
#
# This script automates the manual steps documented here:
# http://answers.microsoft.com/en-us/windows/forum/windows_8-performance/event-id-10016-the-application-specific-permission/9ff8796f-c352-4da2-9322-5fdf8a11c81e?auth=1
#
###############################################################################

Param(
  [string]$CLSID,
  [string]$APPID
)

# TODO: make these parameters?
#$CLSID = "{D63B10C5-BB46-4990-A94F-E40B9D520160}"
#$APPID = "{9CA88EE3-ACB7-47C8-AFC4-AB702511C276}"

Checkpoint-Computer -Description "Fix DCOM errors script"

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
# In order for this script to run, you must be an administrator
# To change the owner you need SeRestorePrivilege
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
Set-RegistryKeyPermissions "HKCR" "CLSID\$CLSID" "S-1-5-32-544" $true
Set-RegistryKeyPermissions "HKLM" "Software\Classes\AppID\$APPID" "S-1-5-32-544" $true

# Add NT AUTHORITY\SYSTEM with Local Launch and Local Activation permissions
Set-DcomAppPermissions $APPID "NT AUTHORITY" "SYSTEM"

Write-Host "Script complete"