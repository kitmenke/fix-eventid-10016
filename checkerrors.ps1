Param(
  [string]$APPID,
  [string]$USERSID
)

$app = get-wmiobject -query ('SELECT * FROM Win32_DCOMApplicationSetting WHERE AppId = "' + $APPID + '"')
if (!$app) {
  Write-Host "No DCOM application was found for $APPID!"
  exit 1
}

$aclList = $app.GetLaunchSecurityDescriptor().Descriptor.DACL
Write-Host ($aclList | Format-List | Out-String)

# check whether the SYSTEM user has access
$acl = $aclList | Where-Object {$_.Trustee.SIDString -eq $USERSID}

$sid = [System.Security.Principal.SecurityIdentifier]$USERSID
$username = $sid.Translate([System.Security.Principal.NTAccount])
if (!$acl) {
  Write-Host "Error! $username ($USERSID) does not have any permission."
  exit 1
} elseif ($acl.AccessMask -ne 31) {
  Write-Host "Error! $username ($USERSID) does not have FULL CONTROL to $APPID. AccessMask is $($acl.AccessMask)."
  exit 1
}

Write-Host "Success! $username ($USERSID) has FULL CONTROL to $APPID."
exit 0