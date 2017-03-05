On Error Resume Next
Const HKEY_LOCAL_MACHINE = &H80000002

set objShell = wscript.createobject("wscript.shell")

strComputer = InputBox("Enter PC name:", "Written by Mike Holmes")
If strComputer <> "" Then
    Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
	    strComputer & "\root\default:StdRegProv")
    strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
    strEntryName = "SCForceOption"
    dwValue = 0
    objReg.SetDWORDValue HKEY_LOCAL_MACHINE, strKeyPath, strEntryName, dwValue
    
    If Not Err.Number = 0 Then 
	    wscript.Echo "Error: "& vbCrLf & strComputer & " not found or no access."
    Else
	    wscript.echo "SUCCESS! "& vbCrLf & "You may now log in one time without your CAC." & _
            vbCrLf & "Upon login a GPO will reset the computer to CAC only." 
    End If
End If