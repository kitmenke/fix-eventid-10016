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
$logEntry = Get-WinEvent -FilterHashTable @{LogName='System'; Level=2} | Where-Object { $_.Message -like "$EVT_MSG*" } | Select-Object -First 1
if ($logEntry -eq $null) {
    Write-Host "No event log entries found."
    exit 1
}

Write-Host "Log entry is:"
Write-Host ($logEntry | Format-List | Out-String)


# Get CLSID and APPID from the event log entry
# which we'll use to look up keys in the registry
$CLSID = $logEntry.Properties[3].Value
Write-Host "CLSID is $CLSID"
$APPID = $logEntry.Properties[4].Value
Write-Host "APPID is $APPID"
