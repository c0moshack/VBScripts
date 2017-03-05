On Error Resume Next
 
strComputer = "."
nbSerial = ""
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colBIOS = objWMIService.ExecQuery _
    ("Select * from Win32_BIOS")
For each objBIOS in colBIOS
	nbSerial = objBIOS.SerialNumber
Next

Set objRootDSE = GetObject("LDAP://RootDSE")
Set objNetwork = WScript.CreateObject("WScript.Network")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"

If Err.Number <> 0 Then
	 MsgBox "Connect Error: " & Err.Description
	 WScript.Quit
End If
wscript.echo "<LDAP://OU=NGWI,OU=States," & objRootDSE.Get("defaultNamingContext") & ">;(&(objectClass=computer)(Name=" & objNetwork.Computername & "));adspath;subtree"
Set adoRecordset = adoConnection.Execute("<LDAP://OU=NGWI,OU=States," & objRootDSE.Get("defaultNamingContext") & ">;(&(objectClass=computer)(Name=" & objNetwork.Computername & "));adspath;subtree")

If Err.Number <> 0 Then
	 MsgBox "Query Error: " & Err.Description
	 WScript.Quit
End If

If Not adoRecordset.EOF Then
	 Set objComputer = GetObject(adoRecordset.Fields(0).Value)
	 objComputer.Put "serialNumber", nbSerial
	 objComputer.SetInfo
End If

If Err.Number <> 0 Then
	 MsgBox "Write Error: " & Err.Description
	 WScript.Quit
End If