' Constants (taken from WinReg.h)
'
Const HKEY_CLASSES_ROOT   = &H80000000
Const HKEY_CURRENT_USER   = &H80000001
Const HKEY_LOCAL_MACHINE  = &H80000002
Const HKEY_USERS          = &H80000003

' Chose computer name, registry tree and key path
'
strComputer = "." ' Use . for current machine
hDefKey = HKEY_LOCAL_MACHINE
strKeyPath = "SOFTWARE\Microsoft\Cryptography\Defaults\Provider"

' Connect to registry provider on target machine with current user
'
Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

' Try to enum the subkeys of the key path we've chosen. We can't if the key doesn't exist
'
If oReg.EnumKey(hDefKey, strKeyPath, arrSubKeys) = 0 Then
  wscript.echo "Key exists!"
Else
  wscript.echo "Key does not exists!"
End If