# WARNING!!!

These scripts can really mess up your PC! Create a system restore point before running the scripts and back up your data.

# Overview

This script is designed to fix a COM error that occurs repeatedly. The error is logged in the event log:

    The application-specific permission settings do not grant Local Activation permission for the COM Server application with CLSID
    {D63B10C5-BB46-4990-A94F-E40B9D520160}
     and APPID
    {9CA88EE3-ACB7-47C8-AFC4-AB702511C276}
     to the user NT AUTHORITY\SYSTEM SID (S-1-5-18) from address LocalHost (Using LRPC) running in the application container Unavailable SID (Unavailable). This security permission can be modified
    using the Component Services administrative tool.

# Scripts

finderrors.ps1 - Search the windows event log for entries matching the above.
checkerrors.ps1 - Check to see if the given user has FULL CONTROL to the application (using app id).
fixerrors.ps1 - TODO: automatically fix the error.

# Notes
This script automates the manual steps documented here:
http://answers.microsoft.com/en-us/windows/forum/windows_8-performance/event-id-10016-the-application-specific-permission/9ff8796f-c352-4da2-9322-5fdf8a11c81e?auth=1

Source:

https://answers.microsoft.com/en-us/windows/forum/windows_10-performance/event-log-writes-error-related-to-application/e64806c2-510c-44c6-b22e-257d07d47200?page=2

  Believe me it works. I have used this fix on 4 Windows 10 machines. I probably did not explain it in detail.

  Gain Administrator and User (full control) permissions of APPID {F72671A9-012C-4725-9D2F-2A4D32D65169} It is controlled by TRUSTED INSTALLER by default.
  Go to Component Services then DCOM CONFIG and scroll down to F72671A9-012C-4725-9D2F-2A4D32D65169 and click properties then security and customize. Then click edit. Click Add and type SYSTEM and tick all 4 "allow" boxes and click OK.
  Then in regedit, restore CLSID permissions back to SYSTEM and APPID permissions back to TRUSTED INSTALLER
  Then reboot and this error will be gone.
  I have a batch file that clears all old Event viewer logs if anyone wants it to clear irrelevant logs.
  Mine is now clean on every boot on all my machines. This fix works. If not, you are missing something out.
