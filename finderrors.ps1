###############################################################################
# Find COM errors
# Author: Kit Menke
# Version 1.0 11/6/2016
###############################################################################

# Notes:
# Get-EventLog doesn't quite work I guess:
# https://stackoverflow.com/questions/31396903/get-eventlog-valid-message-missing-for-some-event-log-sources#
# Get-EventLog Application -EntryType Error -Source "DistributedCOM"
# The application-specific permission settings do not grant Local Activation permission for the COM Server application with CLSID
#$logs = Get-EventLog -LogName "System" -EntryType Error -Source "DCOM" -Newest 1 -Message "The application-specific permission settings do not grant Local Activation permission for the COM Server application with CLSID*"

$EVT_MSG = "The application-specific permission settings do not grant Local Activation permission for the COM Server application with CLSID"
# Search for System event log ERROR entries starting with the specified EVT_MSG
# Level 2 is error, 3 is warning
$logEntry = Get-WinEvent -FilterHashTable @{LogName='System'; Level=2; Id=10016} | Where-Object { $_.Message -like "$EVT_MSG*" } | Select-Object -First 1
if (!$logEntry) {
  Write-Host "No event log entries found."
  exit 1
}

Write-Host "Found an event log entry :"
Write-Host ($logEntry | Format-List | Out-String)
#Write-Host ($logEntry.Properties | Format-List | Out-String)

# Get CLSID and APPID from the event log entry
# which we'll use to look up keys in the registry
$CLSID = $logEntry.Properties[3].Value
Write-Host "CLSID=$CLSID"
$APPID = $logEntry.Properties[4].Value
Write-Host "APPID=$APPID"
$USERDOMAIN = $logEntry.Properties[5].Value
Write-Host "USERDOMAIN=$USERDOMAIN"
$USERNAME = $logEntry.Properties[6].Value
Write-Host "USERNAME=$USERNAME"
$USERSID = $logEntry.Properties[7].Value
Write-Host "USERSID=$USERSID"

Write-Host ".\checkerrors.ps1 ""$APPID"" ""$USERSID"""
Write-Host ".\fixerrors.ps1 ""$APPID"" ""$CLSID"""