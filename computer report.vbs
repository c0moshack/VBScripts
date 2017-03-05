Option Explicit
On Error Resume Next

Dim objRootDSE, objNetwork, objWMIService, objComputer
Dim strComputer, strMake, strModel, strSerialNumber, strMacAddresses, strMemory, intMemory
Dim colBIOS, objBIOS
Dim colNetworkAdapterConfiguration, objNetworkAdapterConfiguration
Dim colComputerSystem, objComputerSystem
Dim colPhysicalMemory, objPhysicalMemory
Dim adoConnection, adoRecordset
dim mydatestring
dim colOperatingSystems,objOperatingSystem,osver
dim oReg, lastuser, snode, spath, svaluename, svalue
dim colcpu, objItem, cpuspec
dim objDisk, colDisks , cspace

strComputer = "NGWINB-DISC4-20"
strMemory = ""
strMake = ""
strModel = ""
strSerialNumber = ""
strMacAddresses = ""
myDateString = Date()

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")
Set colDisks = objWMIService.ExecQuery _
    ("Select * from Win32_LogicalDisk")
For Each objDisk in colDisks
	 If objDisk.DeviceID = "C:" Then
		 Wscript.Echo "DeviceID: " & objDisk.DeviceID       
		 cspace = int(objDisk.Size/1073741824) & " GB"
	 end if
Next

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colcpu = objWMIService.ExecQuery("Select * from Win32_Processor")
For Each objItem in colcpu
    cpuspec = objItem.Name
Next

' use "." for local computer
sNode = "NGWINB-DISC4-20"

Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE

On Error Resume Next
Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _
& sNode & "/root/default:StdRegProv")

If Err.Number <> 0 Then
On Error Goto 0
WScript.Echo "Could not connect to computer " & sNode
Else
On Error Goto 0
sPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI"
sValueName = "LastLoggedonUser"

If oReg.GetStringValue(HKLM, sPath, sValueName, sValue) = 0 Then
lastuser = sValue
Else
lastuser = "unknown" & sNode
End If
End If


Set objRootDSE = GetObject("LDAP://RootDSE")
Set objNetwork = WScript.CreateObject("WScript.Network")
Set objWMIService = GetObject("Winmgmts:\\" & strComputer & "\root\cimv2") 

Set colComputerSystem = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem") 
Set colOperatingSystems = objWMIService.ExecQuery _
("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colOperatingSystems
    osver = objOperatingSystem.Caption & " sp " & objOperatingSystem.ServicePackMajorVersion'& "  " & objOperatingSystem.Version
Next

If Not colComputerSystem Is Nothing Then
	 For Each objComputerSystem In colComputerSystem 
		strMake = objComputerSystem.Manufacturer 
		strModel = objComputerSystem.Model 
	Next
End If

Set colBIOS = objWMIService.ExecQuery("Select * From Win32_BIOS")

If Not colBIOS Is Nothing Then
	 For Each objBIOS in colBIOS
		 strSerialNumber = objBIOS.SerialNumber
	 Next
End If

Set colNetworkAdapterConfiguration = objWMIService.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

strMacAddresses = ""

If Not colNetworkAdapterConfiguration Is Nothing Then
	 For Each objNetworkAdapterConfiguration in colNetworkAdapterConfiguration
		 If strMacAddresses <> "" Then
			 strMacAddresses = strMacAddresses & ", "
		 End If
		 strMacAddresses = strMacAddresses & objNetworkAdapterConfiguration.MACAddress
	 Next
End If

Set colPhysicalMemory = objWMIService.ExecQuery("Select * From Win32_PhysicalMemory")

If Not colPhysicalMemory Is Nothing Then
	 intMemory = 0
	 For Each objPhysicalMemory In colPhysicalMemory
		 intMemory = intMemory + Int(objPhysicalMemory.Capacity)
	 Next
	 strMemory = (intMemory / 1024 / 1024) & " MB"
End If

Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"

If Err.Number <> 0 Then
	 MsgBox "Connect Error: " & Err.Description
	 WScript.Quit
End If

Set adoRecordset = adoConnection.Execute("<LDAP://" & objRootDSE.Get("defaultNamingContext") & ">;(&(objectCategory=Computer)(name=" & objNetwork.Computername & "));adspath;subtree")

If Err.Number <> 0 Then
	 MsgBox "Query Error: " & Err.Description
	 WScript.Quit
End If

If Not adoRecordset.EOF Then
	 Set objComputer = GetObject(adoRecordset.Fields(0).Value)
	 objComputer.Put "description", "Make, " & strmake & " ,Model, " &strModel & ",cpu," & cpuspec &  ",Ram, " & strMemory & ",hdd size," & cspace & " ,Serial No, " & strSerialNumber & " ,MAC Addresses, " & strMACAddresses & "," & osver & ", Last booted, " & mydatestring & ", lastuser, " &lastuser
	 objComputer.SetInfo
End If

If Err.Number <> 0 Then
	 MsgBox "Write Error: " & Err.Description
	 WScript.Quit
End If