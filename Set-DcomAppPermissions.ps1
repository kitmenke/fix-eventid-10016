# Adapted from: Change DCOM config security settings using Powershell (http://stackoverflow.com/a/22104787/98933)
# Overwrites current permissions
# Other sources:
# https://rkeithhill.wordpress.com/2013/07/25/using-powershell-to-modify-dcom-launch-activation-settings/
function Set-DcomAppPermissions {
  param(
    [string]$appid,
    [string]$domain = "NT AUTHORITY",
    [string]$username = "SYSTEM" 
  )
  $app = get-wmiobject -query ('SELECT * FROM Win32_DCOMApplicationSetting WHERE AppId = "' + $appid + '"') -enableallprivileges
  $sdRes = $app.GetLaunchSecurityDescriptor()
  $sd = $sdRes.Descriptor
  $trustee = ([wmiclass] 'Win32_Trustee').CreateInstance()
  $trustee.Domain = $domain
  $trustee.Name = $username
  $fullControl = 31
  $localLaunchActivate = 11
  $ace = ([wmiclass] 'Win32_ACE').CreateInstance()
  #$ace.AccessMask = $localLaunchActivate
  $ace.AccessMask = $fullControl
  $ace.AceFlags = 0
  $ace.AceType = 0
  $ace.Trustee = $trustee
  $sd.DACL | Format-List | Out-File before.txt
  $sd.DACL += $ace
  $app.SetLaunchSecurityDescriptor($sd)
}
