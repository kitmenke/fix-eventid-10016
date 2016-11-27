# WARNING!!!

Create a system restore point and back up your data before running these scripts!

# Overview

This script is designed to fix a COM error that occurs repeatedly. The error is logged in the event log:

    The application-specific permission settings do not grant Local Activation permission for the COM Server application with CLSID
    {D63B10C5-BB46-4990-A94F-E40B9D520160}
     and APPID
    {9CA88EE3-ACB7-47C8-AFC4-AB702511C276}
     to the user NT AUTHORITY\SYSTEM SID (S-1-5-18) from address LocalHost (Using LRPC) running in the application container Unavailable SID (Unavailable). This security permission can be modified
    using the Component Services administrative tool.

Giving the SYSTEM account access is actually pretty complicated. In order to change the permissions for the COM Server application, you need permissions to two separate registry keys:

 - HKEY_Classes_Root\CLSID\*CLSID*
 - HKEY_LocalMachine\Software\Classes\AppID\*APPID*

First, you have to give yourself permission via the Administrators group to the two registry keys. Then, give the SYSTEM account FULL CONTROL permission to the COM Server. Lastly, reset the two keys owner to SYSTEM and TRUSTED INSTALLER.
    
# Scripts

 - finderrors.ps1 - Search the windows event log for entries matching the above.
 - checkerrors.ps1 - Check to see if the given user has FULL CONTROL to the application (using app id).
 - fixerrors.ps1 - Automatically fix the error. This script automates the manual steps documented here:

# Notes

Manual steps to fix the issue:
http://answers.microsoft.com/en-us/windows/forum/windows_8-performance/event-id-10016-the-application-specific-permission/9ff8796f-c352-4da2-9322-5fdf8a11c81e?auth=1

See ChristopherButterfield's answer here: 
https://answers.microsoft.com/en-us/windows/forum/windows_10-performance/event-log-writes-error-related-to-application/e64806c2-510c-44c6-b22e-257d07d47200?page=2

