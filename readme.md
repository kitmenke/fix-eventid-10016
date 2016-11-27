# Overview

This script is designed to fix a COM error that occurs repeatedly. The error is logged in the event log:

    The application-specific permission settings do not grant Local Activation permission for the COM Server application with CLSID
    {D63B10C5-BB46-4990-A94F-E40B9D520160}
     and APPID
    {9CA88EE3-ACB7-47C8-AFC4-AB702511C276}
     to the user NT AUTHORITY\SYSTEM SID (S-1-5-18) from address LocalHost (Using LRPC) running in the application container Unavailable SID (Unavailable). This security permission can be modified
    using the Component Services administrative tool.

This script automates the manual steps documented here:
http://answers.microsoft.com/en-us/windows/forum/windows_8-performance/event-id-10016-the-application-specific-permission/9ff8796f-c352-4da2-9322-5fdf8a11c81e?auth=1
