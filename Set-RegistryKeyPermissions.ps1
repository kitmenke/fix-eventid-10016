# Adapted from: http://stackoverflow.com/a/35844259/98933
# 
# Notes:
# List of well-known SIDs (https://support.microsoft.com/en-us/kb/243330)
# S-1-5-32-544 BUILTIN\Administrators
# S-1-5-32-545 BUILTIN\Users
# S-1-5-18     NT AUTHORITY\SYSTEM
# S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464  NT SERVICE\TrustedInstaller
#$computerName = $env:computername
#$admin = [System.Security.Principal.NTAccount]"$computerName\Administrator"
#$account = [System.Security.Principal.NTAccount]"BUILTIN\Administrators"
#$account = [System.Security.Principal.NTAccount]"NT SERVICE\TrustedInstaller"
#$account = [System.Security.Principal.NTAccount]"BUILTIN\SYSTEM"
# get the sid
# $account.Translate([System.Security.Principal.SecurityIdentifier])
# $sid = [System.Security.Principal.SecurityIdentifier]"S-1-5-18"
# $sid = [System.Security.Principal.SecurityIdentifier]"S-1-5-10"
# NT AUTHORITY\SELF
# $sid.Translate([System.Security.Principal.NTAccount])
# 

function Set-RegistryKeyPermissions {
    # Developed for PowerShell v4.0
    # Required Admin privileges
    # Links:
    #   http://shrekpoint.blogspot.ru/2012/08/taking-ownership-of-dcom-registry.html
    #   http://www.remkoweijnen.nl/blog/2012/01/16/take-ownership-of-a-registry-key-in-powershell/
    #   https://powertoe.wordpress.com/2010/08/28/controlling-registry-acl-permissions-with-powershell/

    param(
    [string]$rootKey,
    [string]$key, 
    [System.Security.Principal.SecurityIdentifier]$sid,
    $recurse = $false)

    switch -regex ($rootKey) {
        'HKCU|HKEY_CURRENT_USER'    { $rootKey = 'CurrentUser' }
        'HKLM|HKEY_LOCAL_MACHINE'   { $rootKey = 'LocalMachine' }
        'HKCR|HKEY_CLASSES_ROOT'    { $rootKey = 'ClassesRoot' }
        'HKCC|HKEY_CURRENT_CONFIG'  { $rootKey = 'CurrentConfig' }
        'HKU|HKEY_USERS'            { $rootKey = 'Users' }
    }

    # !!! I've commented this out since I do this outside of this script !!!
    
    ### Step 1 - escalate current process's privilege
    # get SeTakeOwnership, SeBackup and SeRestore privileges before executes next lines, script needs Admin privilege
    #$import = '[DllImport("ntdll.dll")] public static extern int RtlAdjustPrivilege(ulong a, bool b, bool c, ref bool d);'
    #$ntdll = Add-Type -Member $import -Name NtDll -PassThru
    #$privileges = @{ SeTakeOwnership = 9; SeBackup =  17; SeRestore = 18 }
    #foreach ($i in $privileges.Values) {
    #    $null = $ntdll::RtlAdjustPrivilege($i, 1, 0, [ref]0)
    #}

    function Set-RegistryKeyPermissionsHelper {
        param($rootKey, $key, $sid, $recurse, $recurseLevel = 0)
        
        ### Step 2 - get ownerships of key - it works only for current key
        $regKey = [Microsoft.Win32.Registry]::$rootKey.OpenSubKey($key, 'ReadWriteSubTree', 'TakeOwnership')
        Write-Host $regKey.Name
        
        $acl = New-Object System.Security.AccessControl.RegistrySecurity
        $acl.SetOwner($sid)
        $regKey.SetAccessControl($acl)

        ### Step 3 - enable inheritance of permissions (not ownership) for current key from parent
        $acl.SetAccessRuleProtection($false, $false)
        $regKey.SetAccessControl($acl)

        ### Step 4 - only for top-level key, change permissions for current key and propagate it for subkeys
        # to enable propagations for subkeys, it needs to execute Steps 2-3 for each subkey (Step 5)
        if ($recurseLevel -eq 0) {
            $regKey = $regKey.OpenSubKey('', 'ReadWriteSubTree', 'ChangePermissions')
            $rule = New-Object System.Security.AccessControl.RegistryAccessRule($sid, 'FullControl', 'ContainerInherit', 'None', 'Allow')
            $acl.ResetAccessRule($rule)
            $regKey.SetAccessControl($acl)
        }

        ### Step 5 - recursively repeat steps 2-5 for subkeys
        if ($recurse) {
            foreach($subKey in $regKey.OpenSubKey('').GetSubKeyNames()) {
                Set-RegistryKeyPermissionsHelper $rootKey ($key+'\'+$subKey) $sid $recurse ($recurseLevel+1)
            }
        }
    }

    Set-RegistryKeyPermissionsHelper $rootKey $key $sid $recurse
}