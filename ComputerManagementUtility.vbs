<html>
<head>

<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="description" content="Created by Stuart Barrett">
<meta name="description" content="Last Update: 04/08/11">

<script Language="VBScript">
	On Error Resume Next
	Set objWMIService = GetObject("Winmgmts:\\.\root\cimv2")
	Set colItems = objWMIService.ExecQuery _
		("Select * From Win32_DesktopMonitor where DeviceID = 'DesktopMonitor1'")
	For Each objItem in colItems
		intHorizontal = objItem.ScreenWidth
		intVertical = objItem.ScreenHeight
	Next
	intLeft = (intHorizontal - 1024) / 2
	intTop = (intVertical - 625) / 2
	self.ResizeTo 1024,625
	window.moveTo intLeft, intTop
</script>

<HTA:APPLICATION
  APPLICATIONNAME="PCManagementUtil"
  ID="objPCManage"
  VERSION="3.8"
  BORDER="dialog"
  MAXIMIZEBUTTON="no"
  SCROLL="no"
  CONTEXTMENU="no"
  SINGLEINSTANCE="no"/>

<script Language="VBScript">
	On Error Resume Next
	Set objShell = CreateObject("WScript.Shell")
		strRunAs = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\Setup\strRunAs")
		strTempLoc = objShell.ExpandEnvironmentStrings("%TEMP%")
		If InStr(LCase(objPCManage.commandLine), "true") = 0 Then
			strRunAs = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\Setup\strRunAs")
			booRunAs = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\Setup\booRunAs")
			If booRunAs = 1 Then
				objShell.Run(strRunAs & "*" & strTempLoc & Chr(34))
				Window.Close
			End If
		End If
</script>

<title>PC Management Utility</title>
 
<style type="text/css">
body {
	font-family: times new roman, arial;
	font-size: 0.9em;
}
hr {
	color: #a5a5a5;
	height: 2px;
}
table.htmltable {
	border-width: 2px;
	border-spacing: 2px;
	border-style: outset;
	border-color: gray;
	border-collapse: separate;
	background-color: white;
}
table.htmltable th {
	border-width: 1px;
	padding: 2px;
	border-style: inset;
	border-color: gray;
}
table.htmltable td {
	border-width: 1px;
	padding: 2px;
	border-style: inset;
	border-color: gray;
}
table.menutable {
	background-color: "#EEEEEE";
)
table.menutable th {
}
table.menutable td {
	padding: 0px,20px;
	font-weight:bold;
}
table.processtable {
	table-layout: fixed;
	border-collapse: collapse;
	bordercolor: "#111111";
	width: "100%";
}
table.processtable th {
	background-color: "black";
	color: "white";
	height: 47px;
	cursor: "default";
}
table.processtable td {
	background-color: "white";
	color: "black";
	cursor: "default";
}
table.prodkeystable {
	width:"100%";
}
table.prodkeystable th {
	background-color: white;
	border: 1px solid black;
	width: 295px;
	padding: 1px;
	font-family: 'Courier New';
	width:"50%";
}
#menubar {
  box-shadow:rgb(136, 136, 136) 0px 5px 5px;
  background-color:#EEEEEE;
  border:1px solid #A5A5A5;
  font-color: #A5A5A5;
  font-size: 70%;
  border-bottom:0px;
  bottom:0px;
  height:25px;
  margin:auto;
  padding:5px 5px 5px 5px;
  position:absolute;
  width:100%;
  z-index:500;
}
#tab7, #tab8 {
  cursor:pointer;
}
</style>
  
</head>


<script Language="VBScript">

	Const HKEY_USERS = &H80000003
	Const HKEY_LOCAL_MACHINE = &H80000002
	Const ADS_SCOPE_SUBTREE = 2
	Const ForWriting = 2
	Const ForAppending = 8
	Const adVarChar = 200
	Const MaxCharacters = 255
	Dim strCopy, strPC, strGlobalLogFileName, strPCInfoTab, strLocalSID, strRemoteSID
	Dim strRemoteLoggedOn, strLocalLoggedOn, strRegStart
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	strTemp = objShell.ExpandEnvironmentStrings("%TEMP%")
	strHTMLTempDir = Replace(LCase(strTemp), "c:", "file:///c:")
	strHTMLTempDir = Replace(strHTMLTempDir, "\", "/")
	strFileDate = Right("00" & Day(Now()), 2) & Right("0" & Month(now()), 2) & Year(Now())
	arriValues = Array()
	strTempTitle = "Clean all temp files, which will clean each profile's Temp folder, " & _
	"Hotfix uninstallers over 90 days old (if required), IE history, Firefox Cache, " & _
	"Temp Internet Files, Run / Recent Docs / Regedit Last Key / Paint Recent / " & _
	"Wordpad Recent / Common Dialog Open / Save As MRUs and WMP Recent URLS / Recent Files"
	
	'#--------------------------------------------------------------------------
	'#	If you require more actions in the Action List first amend the
	'#	strActionList variable to allow them to show in the list
	'#--------------------------------------------------------------------------
	strActionList =	"<option value=""9"" title=""Change the Computer Description"">" & _
					"Change Computer Description</option>" & _		
					"<option value=""10"" title=""Change the local Admin password"">" & _
					"Change Local Admin Password</option>" & _	
					"<option value=""12"" title=" & Chr(34) & strTempTitle & _
					Chr(34) & ">Clean Temp Files</option>" & _
					"<option value=""11"" title=""Backup and clear the Application " & _
					"Event Log"">Clear Application Event Log</option>" & _	
					"<option value=""1"" title=""Copy a user profile from / to the PC. " & _
					"This can be copied to / from another PC or to a USB drive"">" & _
					"Copy Profile</option>" & _	
					"<option value=""18"" title=""Delete old user profiles of a specified age"">" & _
					"Delete Old User Profiles</option>" & _						
					"<option value=""2"" title=""Enable / Disable IE Proxy"">" & _
					"Enable / Disable IE Proxy</option>" & _
					"<option value=""3"" title = ""Enable / Disable RDP"">" & _
					"Enable / Disable RDP</option>" & _
					"<option value=""19"" title=""Export Full Inventory"">" & _
					"Export Full Inventory</option>" & _
					"<option value=""14"" title = ""Get Microsoft Product Keys"">" & _
					"Get Microsoft Product Keys</option>" & _
					"<option value=""4"" title = ""Launch Computer Management"">" & _
					"Launch Computer Management</option>" & _
					"<option value=""17"" title=""List updates / hotfixes"">" & _
					"List Updates and Hotfixes</option>" & _
					"<option value=""5"" title=""Open any available share on the remote PC"">Open Share</option>" & _
					"<option value=""6"" title=""Ping Machine"">Ping Machine</option>" & _
					"<option value=""21"" title=""Run PSExec based command"">Run PSExec Command</option>" & _
					"<option value=""22"" title=""Run custome command"">Run Custom Command</option>" & _
					"<option value=""16"" title=""List the current System Restore points and " & _
					"optionally create a new one"">System Restore</option>" & _
					"<option value=""7"" title=""Shutdown / Restart / Log Off"">" & _
					"Shutdown / Restart / Log Off</option>" & _
					"<option value=""23"" title=""View the local user accounts on the PC"">" & _
					"User Accounts</option>" & _
					"<option value=""8"" title=""View user profiles"">View Profiles</option>"
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ExecuteAction()
    '#	PURPOSE........:	Determines what action is performed on choosing an 
    '#						item from the Action List
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	To add actions to the list you must first amend 
	'#						the strActionList variable
    '#--------------------------------------------------------------------------
	Sub ExecuteAction()
		btnStop.Disabled = True
		btnStop.style.cursor = "default"
		btnStop.title = ""
		Select Case Actions.Value
			Case 1
				WaitMessage.InnerHTML = "<hr>"
				CopyProfile()
			Case 2
				WaitMessage.InnerHTML = "<hr>"
				EnableDisableIEProxy()
			Case 3
				WaitMessage.InnerHTML = "<hr>"
				EnableDisableRDP()
			Case 4
				WaitMessage.InnerHTML = "<hr>"
				ManagePC()
			Case 5
				WaitMessage.InnerHTML = "<hr>"
				OpenShare()
			Case 6
				WaitMessage.InnerHTML = "<hr>"
				PingMachine , True, False
			Case 7
				WaitMessage.InnerHTML = "<hr>"
				ShutdownRestartPC()
			Case 8
				WaitMessage.InnerHTML = "<hr>"
				ViewProfiles()
			Case 9
				WaitMessage.InnerHTML = "<hr>"
				ChangePCDescription()
			Case 10
				WaitMessage.InnerHTML = "<hr>"
				ChangeLocalAdminPassword()
			Case 11
				WaitMessage.InnerHTML = "<hr>"
				ClearAppEventLog()
			Case 12
				WaitMessage.InnerHTML = "<hr>"
				DeleteTempFiles()
			Case 13	'WR Specific
				WaitMessage.InnerHTML = "<hr>"
				TMobileReg()	'WR Specific
			Case 14
				WaitMessage.InnerHTML = "<hr>"
				GetMSProductKeys()
			Case 15
				WaitMessage.InnerHTML = "<hr>"
				ProcessAudit()
			Case 16
				WaitMessage.InnerHTML = "<hr>"
				SystemRestore()
			Case 17
				ListUpdates2()
			Case 18
				CleanProfiles2()
			Case 19
				WaitMessage.InnerHTML = "<hr>"
				ExpInventory()
			Case 20
				WaitMessage.InnerHTML = "<hr>"
				DeleteStartupItems()
			Case 21
				WaitMessage.InnerHTML = "<hr>"
				RunPSExecCommand()
			Case 22
				WaitMessage.InnerHTML = "<hr>"
				RunCustomCommand()
			Case 23
				ShowUserAccountsInfo()
		End Select
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ChangeButtonColour(strbtnName)
    '#	PURPOSE........:	Changes button colour on click
    '#	ARGUMENTS......:	strbtnName = Name of the button
    '#	EXAMPLE........:	ChangeButtonColour(btnCheck)
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ChangeButtonColour(strbtnName)
		On Error Resume Next
		strbtnName.style.backgroundcolor="#CCCCCC"
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	RevertButtonColour(strbtnName)
    '#	PURPOSE........:	Changes button colour back after click
    '#	ARGUMENTS......:	strbtnName = Name of the button
    '#	EXAMPLE........:	RevertButtonColour(btnCheck)
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub RevertButtonColour(strbtnName)
		On Error Resume Next
		strbtnName.style.backgroundcolor="#DDDDDD"
	End Sub

	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ClearPCs()
    '#	PURPOSE........:	Clears all of the PCs from the AvailablePCs listbox
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------	
	Sub ClearPCs()
		For Each objOption in AvailablePCs.Options
			objOption.RemoveNode
		Next 
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	PopulateOUs()
    '#	PURPOSE........:	Auto-populates the AvailableOUs list box with all
	'#						of the OUs in the Forest
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	Does not seem to work on Windows 7 PCs
    '#--------------------------------------------------------------------------
	Sub PopulateOUs()
		On Error Resume Next
		
		strRootOU = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\Setup\strRootOU")
		If strRootOU = "" Then
			Set objRootDSE = GetObject("LDAP://RootDSE")
			strDNSDomain = objRootDSE.Get("defaultNamingContext")
			strRootOU = strDNSDomain
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\Setup\strRootOU", strRootOU, "REG_SZ"
		End If
		
		Set objConnection = CreateObject("ADODB.Connection")
		Set objCommand =   CreateObject("ADODB.Command")
		objConnection.Provider = "ADsDSOObject"
		objConnection.Open "Active Directory Provider"
		Set objCommand.ActiveConnection = objConnection

		strFilter = "(&(objectCategory=organizationalUnit))"
		strQuery = "<LDAP://" & strRootOU & ">;" & strFilter & _
		";distinguishedName, ADsPath;subtree" 
				
		objCommand.CommandText = strQuery
		objCommand.Properties("Page Size") = 750
		objCommand.Properties("Timeout") = 60
		objCommand.Properties("Cache Results") = False

		Set objRecordSet = objCommand.Execute

		objRecordSet.MoveFirst
		Do Until objRecordSet.EOF
			Set objOption = Document.createElement("OPTION")
			objOption.Text = objRecordSet.Fields("distinguishedName").Value
			objOption.Title = objRecordSet.Fields("distinguishedName").Value
			objOption.Value = objRecordSet.Fields("ADsPath").Value
			AvailableOUs.Add(objOption)
			objRecordSet.MoveNext
		Loop
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ShowOUPCs()
    '#	PURPOSE........:	Populates the AvailablePCs listbox with the PCs in
    '#						the OU selected in the AvailableOUs listbox
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ShowOUPCs()
		On Error Resume Next
		
		DataArea.InnerHTML = ""
		ClearPCs
		
		btnSelectPC.Disabled = False
		btnSelectPC.Title = "Select highlighted PC"
		
		AvailablePCs.Disabled = False
		
		strRootOU = AvailableOUs.Value
		Set objConnection = CreateObject("ADODB.Connection")
		Set objCommand =   CreateObject("ADODB.Command")
		objConnection.Provider = "ADsDSOObject"
		objConnection.Open "Active Directory Provider"
		Set objCommand.ActiveConnection = objConnection

		strQuery = "Select Name, ADsPath FROM '" & strRootOU & _
		"' WHERE objectCategory='Computer' ORDER BY Name"
				
		objCommand.CommandText = strQuery
		objCommand.Properties("Page Size") = 750
		objCommand.Properties("Timeout") = 60
		objCommand.Properties("Cache Results") = False

		Set objRecordSet = objCommand.Execute
		
		objRecordSet.MoveFirst
		
		If (objRecordset.EOF = True) Then
			btnSelectPC.Disabled = True
			btnSelectPC.style.cursor = "default"
			btnSelectPC.Title = ""
			MsgBox "No computers have been found in the OU '" & strRootOU & _
			"'.",48,"PC Management Utility"
			CleanUp
		End If
		
		Do Until objRecordSet.EOF
			Set objOption = Document.createElement("OPTION")
			objOption.Text = objRecordSet.Fields("Name").Value
			objOption.Value = objRecordSet.Fields("Name").Value	
			AvailablePCs.Add(objOption)
			objRecordSet.MoveNext
		Loop
		objRecordset.Close
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ChangeSearchView()
    '#	PURPOSE........:	Sets the default Search View (AD / IP Range search)
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ChangeSearchView(intSearchView)
		If intSearchView = 1 Then
			SearchArea.InnerHTML = "<div  align=""center""><b>" & _
			"<input type=""text"" name=""IP1A"" id=""IP1A"" size=""2"" " & _
			"maxLength=""3"" onKeyUp=""IPCheck1A"">&nbsp;.&nbsp;" & _
			"<input type=""text"" name=""IP2A"" id=""IP2A"" size=""2"" " & _
			" maxLength=""3"" onKeyUp=""IPCheck2A"">&nbsp;.&nbsp;" & _
			"<input type=""text"" name=""IP3A"" id=""IP3A"" size=""2"" " & _
			"maxLength=""3"" onKeyUp=""IPCheck3A"">&nbsp;.&nbsp;" & _
			"<input type=""text"" name=""IP4A"" id=""IP4A"" size=""2"" " & _
			"maxLength=""3"" onKeyUp=""IPCheck4A"">" & _
			"<br>to<br>" & _
			"<input type=""text"" name=""IP1B"" id=""IP1B"" size=""2"" " & _
			"maxLength=""3"" onKeyUp=""IPCheck1B"">&nbsp;.&nbsp;" & _
			"<input type=""text"" name=""IP2B"" id=""IP2B"" size=""2"" " & _
			"maxLength=""3"" onKeyUp=""IPCheck2B"">&nbsp;.&nbsp;" & _
			"<input type=""text"" name=""IP3B"" id=""IP3B"" size=""2"" " & _
			"maxLength=""3"" onKeyUp=""IPCheck3B"">&nbsp;.&nbsp;" & _
			"<input type=""text"" name=""IP4B"" id=""IP4B"" size=""2"" " & _
			"maxLength=""3"" onKeyUp=""IPCheck4B""></b><br>" & _
			"<input id=runbutton type=""button"" value=""Search Range"" " & _
			"name=""btnSearchRange"" title=""Search for IPs in IP range"" " & _
			"onClick=""SearchPCRange"" onMouseDown=""ChangeButtonColour(btnSearchRange)"" " & _
			"onMouseUp=""RevertButtonColour(btnSearchRange)""> &nbsp;&nbsp; <span id=""Last5"" " & _
			"title=""View the last 5 used IP Ranges"" " & _
			"style=""text-decoration:underline;color:blue;cursor:pointer"" " & _
			"onclick=""LastFive"">Last 5</span>" & _
			"<p><select size=""16"" name=""AvailablePCs"" style=""width:300"" " & _
			"onDblClick=""ShowPCInfo AvailablePCs.Value, False"" onChange=""AddHostName()"">" & _
			"</select></div>"
			Else
				SearchArea.InnerHTML = "<div align=""center""><select size=""10"" " & _
				"name=""AvailableOUs"" style=""width:300"" onDblClick=""ShowOUPCs"">" & _
				"</select><br><input id=runbutton " & _
				"type=""button"" value=""Show PCs in OU"" name=""btnSelectOU"" " & _
				"title=""Show PCs in highlighted OU"" onClick=""ShowOUPCs"" " & _
				"onMouseDown=""ChangeButtonColour(btnSelectOU)"" " & _
				"onMouseUp=""RevertButtonColour(btnSelectOU)"">&nbsp;&nbsp;" & _
				"<p><select size=""10"" name=""AvailablePCs"" style=""width:300"" " & _
				"onDblClick=""ShowPCInfo AvailablePCs.Value, False"">" & _
				"</select></div>"
				PopulateOUs
		End If
		CleanUp
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	IPCheck1A() - IPCheck4B()
    '#	PURPOSE........:	Validates the IP entered as it is typed
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub IPCheck1A()
		On Error Resume Next
		If IP1A.Value > 255 OR IsNumeric(IP1A.Value) = False  Then
			IP1A.Value = ""
		End If
		If Len(IP1A.Value) = 3 OR InStr(IP1A.Value, ".") > 0 Then
			IP1A.Value=Replace(IP1A.Value,".","")
			IP2A.Focus
			IP2A.Select
		End If
	End Sub

	Sub IPCheck2A()
		On Error Resume Next
		If IP2A.Value > 255 OR IsNumeric(IP2A.Value) = False Then
			IP2A.Value = ""
		End If
		If Len(IP2A.Value) = 3 OR InStr(IP2A.Value, ".") > 0 Then
			IP2A.Value=Replace(IP2A.Value,".","")
			IP3A.Focus
			IP3A.Select
		End If
	End Sub

	Sub IPCheck3A()
		On Error Resume Next
		If IP3A.Value > 255 OR IsNumeric(IP3A.Value) = False Then
			IP3A.Value = ""
		End If
		If Len(IP3A.Value) = 3 OR InStr(IP3A.Value, ".") > 0 Then
			IP3A.Value=Replace(IP3A.Value,".","")
			IP4A.Focus
			IP4A.Select
		End If
	End Sub

	Sub IPCheck4A()
		On Error Resume Next
		If IP4A.Value > 255 OR IsNumeric(IP4A.Value) = False Then
			IP4A.Value = ""
		End If
		If Len(IP4A.Value) = 3 OR InStr(IP4A.Value, ".") > 0 Then
			IP4A.Value=Replace(IP4A.Value,".","")
			IP1B.Focus
			IP1B.Select
		End If
	End Sub

	Sub IPCheck1B()
		On Error Resume Next
		If IP1B.Value > 255 OR IsNumeric(IP1B.Value) = False  Then
			IP1B.Value = ""
		End If
		If Len(IP1B.Value) = 3 OR InStr(IP1B.Value, ".") > 0 Then
			If CInt(IP1B.Value) < CInt(IP1A.Value) Then
				IP1B.Value=Replace(IP1B.Value,".","")
				MsgBox "'" & IP1B.Value & "' is not a valid entry", _
				vbExclamation, "PC Management Utility"
				IP1B.Value = ""
				Else
					IP1B.Value=Replace(IP1B.Value,".","")
					IP2B.Focus
					IP2B.Select
			End If
		End If
	End Sub

	Sub IPCheck2B()
		On Error Resume Next
		If IP2B.Value > 255 OR IsNumeric(IP2B.Value) = False Then
			IP2B.Value = ""
		End If
		If Len(IP2B.Value) = 3 OR InStr(IP2B.Value, ".") > 0 Then
			If IP1A.Value = IP1B.Value AND CInt(IP2B.Value) < CInt(IP2A.Value) Then
				IP2B.Value=Replace(IP2B.Value,".","")
				MsgBox "'" & IP2B.Value & "' is not a valid entry", _
				vbExclamation, "PC Management Utility"
				IP2B.Value = ""
				Else
					IP2B.Value=Replace(IP2B.Value,".","")
					IP3B.Focus
					IP3B.Select
			End If
		End If
	End Sub

	Sub IPCheck3B()
		On Error Resume Next
		If IP3B.Value > 255 OR IsNumeric(IP3B.Value) = False Then
			IP3B.Value = ""
		End If
		If Len(IP3B.Value) = 3 OR InStr(IP3B.Value, ".") > 0 Then
			If IP1A.Value = IP1B.Value AND IP2A.Value = IP2B.Value AND _
			CInt(IP3B.Value) < CInt(IP3A.Value) Then
				IP3B.Value=Replace(IP3B.Value,".","")
				MsgBox "'" & IP3B.Value & "' is not a valid entry", _
				vbExclamation, "PC Management Utility"
				IP3B.Value = ""
				Else
					IP3B.Value=Replace(IP3B.Value,".","")
					IP4B.Focus
					IP4B.Select
			End If
		End If
	End Sub
	
	Sub IPCheck4B()
		On Error Resume Next
		If IP4B.Value > 255 OR IsNumeric(IP4B.Value) = False Then
			IP4B.Value = ""
		End If
		If Len(IP4B.Value) = 3 OR InStr(IP4B.Value, ".") > 0 Then
			If IP1A.Value = IP1B.Value AND IP2A.Value = IP2B.Value AND _
			IP3A.Value = IP3B.Value AND CInt(IP4B.Value) < CInt(IP4A.Value) Then
				IP4B.Value=Replace(IP4B.Value,".","")
				MsgBox "'" & IP4B.Value & "' is not a valid entry", _
				vbExclamation, "PC Management Utility"
				IP4B.Value = ""
				Else
					IP4B.Value=Replace(IP4B.Value,".","")
					btnSearchRange.Focus
			End If
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	UpdateLastFive()
    '#	PURPOSE........:	Adds the current IP range into the Last 5 IP ranges
	'#						list
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub UpdateLastFive()
		On Error Resume Next
		strIPAddress1 = IP1A.Value & "." & IP2A.Value & "." & _
		IP3A.Value & "." & IP4A.Value
		strIPAddress2 = IP1B.Value & "." & IP2B.Value & "." & _
		IP3B.Value & "." & IP4B.Value
		
		strLastIP1A = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IPRanges\IP1A")
		strLastIP1B = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IPRanges\IP1B")
		strLastIP2A = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IPRanges\IP2A")
		strLastIP2B = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IPRanges\IP2B")
		strLastIP3A = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IPRanges\IP3A")
		strLastIP3B = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IPRanges\IP3B")
		strLastIP4A = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IPRanges\IP4A")
		strLastIP4B = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IPRanges\IP4B")
		strLastIP5A = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IPRanges\IP5A")
		strLastIP5B = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IPRanges\IP5B")
		
		If strLastIP5A = "" Then
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP5A", strIPAddress1, "REG_SZ"
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP5B", strIPAddress2, "REG_SZ"
			Exit Sub
		End If
		If strLastIP4A = "" Then
			If strLastIP5A <> strIPAddress1 AND strLastIP5B <> strIPAddress2 Then
				objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP4A", strIPAddress1, "REG_SZ"
				objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP4B", strIPAddress2, "REG_SZ"
				Exit Sub
			End If
		End If
		If strLastIP3A = "" Then
			If strLastIP5A <> strIPAddress1 AND strLastIP5B <> strIPAddress2 Then
				If strLastIP4A <> strIPAddress1 AND strLastIP4B <> strIPAddress2 Then
					objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP3A", _
					strIPAddress1, "REG_SZ"
					objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP3B", _
					strIPAddress2, "REG_SZ"
					Exit Sub
				End If
			End If
		End If
		If strLastIP2A = "" Then
			If strLastIP5A <> strIPAddress1 AND strLastIP5B <> strIPAddress2 Then
				If strLastIP4A <> strIPAddress1 AND strLastIP4B <> strIPAddress2 Then
					If strLastIP3A <> strIPAddress1 AND strLastIP3B <> strIPAddress2 Then
						objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP2A", _
						strIPAddress1, "REG_SZ"
						objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP2B", _
						strIPAddress2, "REG_SZ"
						Exit Sub
					End If
				End If
			End If
		End If
		If strLastIP1A = "" Then
			If strLastIP5A <> strIPAddress1 AND strLastIP5B <> strIPAddress2 Then
				If strLastIP4A <> strIPAddress1 AND strLastIP4B <> strIPAddress2 Then
					If strLastIP3A <> strIPAddress1 AND strLastIP3B <> strIPAddress2 Then
						If strLastIP2A <> strIPAddress1 AND strLastIP2B <> strIPAddress2 Then
							objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP1A", _
							strIPAddress1, "REG_SZ"
							objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP1B", _
							strIPAddress2, "REG_SZ"
							Exit Sub
						End If
					End If
				End If
			End If
		End If
		
		If strLastIP1A = strIPAddress1 AND strLastIP1B = strIPAddress2 Then
			Exit Sub
		End If
		If strLastIP2A = strIPAddress1 AND strLastIP2B = strIPAddress2 Then
			Exit Sub
		End If
		If strLastIP3A = strIPAddress1 AND strLastIP3B = strIPAddress2 Then
			Exit Sub
		End If
		If strLastIP4A = strIPAddress1 AND strLastIP4B = strIPAddress2 Then
			Exit Sub
		End If
		If strLastIP5A = strIPAddress1 AND strLastIP5B = strIPAddress2 Then
			Exit Sub
		End If
		
		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP1A", strIPAddress1, "REG_SZ"
		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP1B", strIPAddress2, "REG_SZ"
		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP2A", strLastIP1A, "REG_SZ"
		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP2B", strLastIP1B, "REG_SZ"
		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP3A", strLastIP2A, "REG_SZ"
		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP3B", strLastIP2B, "REG_SZ"
		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP4A", strLastIP3A, "REG_SZ"
		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP4B", strLastIP3B, "REG_SZ"
		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP5A", strLastIP4A, "REG_SZ"
		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP5B", strLastIP4B, "REG_SZ"
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	LastFive()
    '#	PURPOSE........:	Allows user to choose from last 5 IP ranges
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub LastFive()
		On Error Resume Next
		strLastIP1A = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IPRanges\IP1A")
		strLastIP1B = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IPRanges\IP1B")
		strLastIP2A = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IPRanges\IP2A")
		strLastIP2B = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IPRanges\IP2B")
		strLastIP3A = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IPRanges\IP3A")
		strLastIP3B = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IPRanges\IP3B")
		strLastIP4A = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IPRanges\IP4A")
		strLastIP4B = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IPRanges\IP4B")
		strLastIP5A = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IPRanges\IP5A")
		strLastIP5B = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IPRanges\IP5B")
		
		If strLastIP1A = "" Then 
			strLastIP1A = "None"
			strLastIP1B = "None"
		End If
		If strLastIP2A = "" Then 
			strLastIP2A = "None"
			strLastIP2B = "None"
		End If
		If strLastIP3A = "" Then 
			strLastIP3A = "None"
			strLastIP3B = "None"
		End If
		If strLastIP4A = "" Then 
			strLastIP4A = "None"
			strLastIP4B = "None"
		End If
		If strLastIP5A = "" Then 
			strLastIP5A = "None"
			strLastIP5B = "None"
		End If
		
		strLastFiveMsg = "1. " & strLastIP1A & " to " & strLastIP1B & vbCrLf & _
		"2. " & strLastIP2A & " to " & strLastIP2B & vbCrLf & _
		"3. " & strLastIP3A & " to " & strLastIP3B & vbCrLf & _
		"4. " & strLastIP4A & " to " & strLastIP4B & vbCrLf & _
		"5. " & strLastIP5A & " to " & strLastIP5B & vbCrLf & _
		vbCrLf & "Please choose an IP range from the list"
		
		intLast5 = InputBox(strLastFiveMsg,"Last 5 IP Ranges")
		
		If IsNull(intLast5) OR intLast5 = "" Then
			Exit Sub
		End If
		
		InArray = False
		For a = 1 to 5
		If InStr(a,intLast5) = 1 Then
				InArray = True
			End If
		Next
		If InArray = False Then 
			MsgBox "'" & intLast5 & "' is not a valid response", _
			vbExclamation, "PC Management Utility"
			Exit Sub
		End If
		Select Case intLast5
			Case 1
				If strLastIP1A <> "None" Then
					CleanUp
					arrIP1 = Split(strLastIP1A,".")
					arrIP2 = Split(strLastIP1B,".")
					IP1A.Value = arrIP1(0)
					IP2A.Value = arrIP1(1)
					IP3A.Value = arrIP1(2)
					IP4A.Value = arrIP1(3)
					IP1B.Value = arrIP2(0)
					IP2B.Value = arrIP2(1)
					IP3B.Value = arrIP2(2)
					IP4B.Value = arrIP2(3)
					Else
						MsgBox "This IP range is empty", _
						vbExclamation, "PC Management Utility"
				End If
			Case 2
				CleanUp
				If strLastIP2A <> "None" Then
					arrIP1 = Split(strLastIP2A,".")
					arrIP2 = Split(strLastIP2B,".")
					IP1A.Value = arrIP1(0)
					IP2A.Value = arrIP1(1)
					IP3A.Value = arrIP1(2)
					IP4A.Value = arrIP1(3)
					IP1B.Value = arrIP2(0)
					IP2B.Value = arrIP2(1)
					IP3B.Value = arrIP2(2)
					IP4B.Value = arrIP2(3)
					Else
						MsgBox "This IP range is empty", _
						vbExclamation, "PC Management Utility"
				End If
			Case 3
				CleanUp
				If strLastIP3A <> "None" Then
					arrIP1 = Split(strLastIP3A,".")
					arrIP2 = Split(strLastIP3B,".")
					IP1A.Value = arrIP1(0)
					IP2A.Value = arrIP1(1)
					IP3A.Value = arrIP1(2)
					IP4A.Value = arrIP1(3)
					IP1B.Value = arrIP2(0)
					IP2B.Value = arrIP2(1)
					IP3B.Value = arrIP2(2)
					IP4B.Value = arrIP2(3)
					Else
						MsgBox "This IP range is empty", _
						vbExclamation, "PC Management Utility"
				End If
			Case 4
				CleanUp
				If strLastIP4A <> "None" Then
					arrIP1 = Split(strLastIP4A,".")
					arrIP2 = Split(strLastIP4B,".")
					IP1A.Value = arrIP1(0)
					IP2A.Value = arrIP1(1)
					IP3A.Value = arrIP1(2)
					IP4A.Value = arrIP1(3)
					IP1B.Value = arrIP2(0)
					IP2B.Value = arrIP2(1)
					IP3B.Value = arrIP2(2)
					IP4B.Value = arrIP2(3)
					Else
						MsgBox "This IP range is empty", _
						vbExclamation, "PC Management Utility"
				End If
			Case 5
				CleanUp
				If strLastIP5A <> "None" Then
					arrIP1 = Split(strLastIP5A,".")
					arrIP2 = Split(strLastIP5B,".")
					IP1A.Value = arrIP1(0)
					IP2A.Value = arrIP1(1)
					IP3A.Value = arrIP1(2)
					IP4A.Value = arrIP1(3)
					IP1B.Value = arrIP2(0)
					IP2B.Value = arrIP2(1)
					IP3B.Value = arrIP2(2)
					IP4B.Value = arrIP2(3)
					Else
						MsgBox "This IP range is empty", _
						vbExclamation, "PC Management Utility"
				End If
		End Select
		SearchPCRange
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	SearchPCRange()
    '#	PURPOSE........:	Searches through the entered IP range and adds PCs
	'#						to AvailablePCs listbox
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	Adds PCs in format {IP ADDRESS} during first pass
	'#						and then {IP ADDRESS} - {COMPUTER NAME} after
	'#						second (much quicker) pass
    '#--------------------------------------------------------------------------
	Sub SearchPCRange()
		On Error Resume Next
		Dim arrOfflinePCs()
		n = 0
		DataArea.InnerHTML = ""
		ClearPCs
		
		booResolveHostNames = objShell.RegRead(strRegStart & "\Software\SKB\PCManagementUtil\Others\booResolveHostNames")
		If booResolveHostNames = "" OR IsNull(booResolveHostNames) Then 
			objShell.RegWrite strRegStart & "\Software\SKB\PCManagementUtil\Others\booResolveHostNames", _
			"1", "REG_SZ"
			booResolveHostNames = "1"
		End If
		
		AvailablePCs.Disabled = False	
		strIPAddress1 = IP1A.Value & "." & IP2A.Value & "." & _
		IP3A.Value & "." & IP4A.Value
		If IP1A.Value = "" OR IP2A.Value = "" _
		OR IP3A.Value = "" OR IP4A.Value = "" Then
			MsgBox "You have entered an invalid IP Range", _
			vbExclamation, "PC Management Utility"
			CleanUp
			IP1A.Focus
			Exit Sub
		End If
		If IP1B.Value = "" OR IP2B.Value = "" _
		OR IP3B.Value = "" OR IP4B.Value = "" Then
			IP1B.Value = IP1A.Value
			IP2B.Value = IP2A.Value
			IP3B.Value = IP3A.Value
			IP4B.Value = IP4A.Value
			strIPAddress2 = strIPAddress1
			Else
				strIPAddress2 = IP1B.Value & "." & IP2B.Value & "." & _
				IP3B.Value & "." & IP4B.Value
				If CheckIP(strIPAddress1,strIPAddress2) = False Then
					MsgBox strIPAddress1 & " to " & strIPAddress2 & _
					" is not a valid IP Range", vbExclamation, "PC Management Utility"
					CleanUp
					IP1A.Focus
					Exit Sub
				End If
		End If
		UpdateLastFive

		If CInt(IP1B.Value) > CInt(IP1A.Value) OR CInt(IP2B.Value) > CInt(IP2A.Value) Then
			MsgBox "This IP address range is too large!",vbExclamation, "PC Management Utility"
			Exit Sub
			ElseIf CInt(IP3B.Value) > CInt(IP3A.Value) Then
				If booResolveHostNames = "1" Then
					LongSearchPrompt = MsgBox("This operation will take a long time" & _
					vbCrLf & vbCrLf & "Do you wish to continue?",vbYesNo+vbExclamation, _
					"PC Management Utility")
					If LongSearchPrompt = vbNo Then
						Exit Sub
					End If
				End If
				y = (CInt(IP3B.Value) - CInt(IP3A.Value)) * 255
				x = IP4B.Value + (y - CInt(IP4A.Value))
				Else
					If IP4B.Value - IP4A.Value > 100 Then
						LongSearchPrompt = MsgBox("This operation will take a long time" & _
						vbCrLf & vbCrLf & "Do you wish to continue?",vbYesNo+vbExclamation, _
						"PC Management Utility")
						If LongSearchPrompt = vbNo Then
							Exit Sub
						End If
				End If
					x = IP4B.Value - IP4A.Value
		End If	
		
		If strIPAddress1 <> strIPAddress2 Then
			strCurrentIP = strIPAddress1
			For i = 1 to x + 1
				On Error Resume Next
				If strPC <> "" OR IP1A.Value = "" Then
					Exit Sub
				End If
				Set objOption = Document.createElement("OPTION")
				objOption.Text = strCurrentIP
				objOption.Title = strCurrentIP
				objOption.Value = strCurrentIP
				AvailablePCs.Add(objOption)
				strTempIP = strCurrentIP
				strCurrentIP = NewIP(strTempIP)
			Next
			
			If booResolveHostNames = "1" Then
				On Error Goto 0
				Set objWMIService2 = GetObject("winmgmts:\\.\root\cimv2")
				For Each objItem In AvailablePCs.Options
					strCurrentIP = objItem.Value
					DataArea.InnerHTML = "<h3><i>Adding Computer Name for " & strCurrentIP & "...</i></h3>"
					PauseScript(1)
					
					If Reachable(strCurrentIP) Then
						Set colPing = objWMIService2.ExecQuery _
							("Select * from Win32_PingStatus Where Address = '" & strCurrentIP & _
							"' AND ResolveAddressNames = TRUE")
						For Each objItem3 in colPing
							strOptionValue = objItem3.ProtocolAddressResolved
						Next
						
						objItem.Text = strCurrentIP & " - " & strOptionValue
						objItem.Title = strCurrentIP & " - " & strOptionValue
						objItem.Value = strOptionValue	
						Else
							objItem.RemoveNode
					End If
				Next
			End If
			DataArea.InnerHTML = "<h3><i>Finished Scan!</i></h3>"
			Else
				If Reachable(strIPAddress1) Then
					ShowPCInfo strIPAddress1, False
					Else
						ErrConnectPrompt = MsgBox("Error connecting to " & strIPAddress1 & vbCrLf & _
						vbCrLf & "Would you like to run a continuous ping to this machine?", _
						vbExclamation+vbYesNo,"Error")
						Err.Clear
						If ErrConnectPrompt = vbYes Then
							btnRescan.Disabled = False
							btnRescan.Title = "Rescan the selected PC"
							DataArea.InnerHTML = "<div align=""center"">" & _
							"<span id=""WaitMessage""></span></div>" & _
							"<input id=runbutton  class=""button"" type=""button"" value=""Stop"" " & _
							"name=""btnStop"" onclick=""StopAction"">"
							PingMachine strIPAddress1,False,True
							Else
								CleanUp
								Exit Sub
						End If
				End If
		End If
	End Sub
	
	Sub AddHostName()
		On Error Resume Next
		booResolveHostNames = objShell.RegRead(strRegStart & "\Software\SKB\PCManagementUtil\Others\booResolveHostNames")

		If booResolveHostNames = 0 Then
			For Each objItem In AvailablePCs
				If objItem.Value = AvailablePCs.Value AND InStr(objItem.Text, "-") = 0 Then
					strCurrentIP = objItem.Value
					DataArea.InnerHTML = "<h3><i>Adding Computer Name for " & strCurrentIP & "...</i></h3>"
					PauseScript(1)
					Set objWMIService2 = GetObject("winmgmts:\\.\root\cimv2")
					If Reachable(strCurrentIP) Then
						Set colPing = objWMIService2.ExecQuery _
						("Select * from Win32_PingStatus Where Address = '" & strCurrentIP & _
						"' AND ResolveAddressNames = TRUE")
						For Each objItem2 in colPing
							strOptionValue = objItem2.ProtocolAddressResolved
						Next
						objItem.Text = strCurrentIP & " - " & strOptionValue
						objItem.Title = strCurrentIP & " - " & strOptionValue
						objItem.Value = strOptionValue
						Else
							objItem.RemoveNode
					End If
				End If
				DataArea.InnerHTML = "<h3><i>Finished!</i></h3>"
			Next
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ShowPCInfo(strComp, booChoice)
    '#	PURPOSE........:	Shows details of selected PC along with Action List
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	ShowPCInfo(PC1, True)
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ShowPCInfo(strComp, booChoice)
		If booChoice = False Then
			On Error Resume Next			
			booRescan = False
			strError = ""
			If IsNull(strComp) OR strComp = "" OR strComp = "." Then
				strComp = objShell.ExpandEnvironmentStrings("%ComputerName%")
			End If
			If UCase(strPC) = UCase(strComp) Then
				booRescan = True
				Else
					strPC = Trim(strComp)
			End If
			If InStr(strPC, ".") > 0 Then
				strPC = GetPCName()
			End If
			strPC = UCase(strPC)
			PCSearch.Value = strPC
			AvailablePCs.style.backgroundcolor = "#EEEEEE"
			If SearchView.Value = 1 Then
				IP1A.Disabled = True
				IP1A.Style.backgroundcolor = "#EEEEEE"
				IP2A.Disabled = True
				IP2A.Style.backgroundcolor = "#EEEEEE"
				IP3A.Disabled = True
				IP3A.Style.backgroundcolor = "#EEEEEE"
				IP4A.Disabled = True
				IP4A.Style.backgroundcolor = "#EEEEEE"
				IP1B.Disabled = True
				IP1B.Style.backgroundcolor = "#EEEEEE"
				IP2B.Disabled = True
				IP2B.Style.backgroundcolor = "#EEEEEE"
				IP3B.Disabled = True
				IP3B.Style.backgroundcolor = "#EEEEEE"
				IP4B.Disabled = True
				IP4B.Style.backgroundcolor = "#EEEEEE"
				
				btnSearchRange.Disabled = True
				btnSearchRange.style.cursor = "default"
				btnSearchRange.Title = ""
				
				Else
					btnSelectOU.Disabled = True
					btnSelectOU.style.cursor = "default"
					btnSelectOU.Title = ""
					
					AvailableOUs.Disabled = True
					AvailableOUs.Style.backgroundcolor = "#EEEEEE"
			End If
			AvailablePCs.Disabled = True
			
			btnSearch.Disabled = True
			btnSearch.style.cursor = "default"
			btnSearch.Title = ""
			
			PCSearch.style.backgroundcolor = "#EEEEEE"
			PCSearch.Disabled = True
			
			btnSelectPC.Disabled = True
			btnSelectPC.style.cursor = "default"
			btnSelectPC.Title = ""
			
			If Reachable(strPC) Then
				SoftwareTab.InnerHTML = ""
				ProcessesTab.InnerHTML = ""
				StartupTab.InnerHTML = ""
				ServicesTab.InnerHTML = ""
				
				If booRescan = False Then
					DataArea.InnerHTML = "<h3><i>Fetching info for " & strPC & ", please wait.</i></h3>"
					Else
						DataArea.InnerHTML = "<h3><i>Rescanning " & strPC & ", please wait.</i></h3>"
				End If
				PauseScript(1)
				btnRescan.Disabled = False
				btnRescan.Title = "Rescan the selected PC"
				
				strQueryChoices = objShell.RegRead(strRegStart & "\Software\SKB\PCManagementUtil\Setup\strQueryChoices")
				If strQueryChoices = "" Then
					objShell.RegWrite strRegStart & "\Software\SKB\PCManagementUtil\Setup\strQueryChoices", _
					"1,2,3,4,5,", "REG_SZ"
					strQueryChoices = "1,2,3,4,5,"
				End If
				
				If InStr(strQueryChoices, "1") > 0 Then booPCInfo = True
				If InStr(strQueryChoices, "2") > 0 Then booSoftware = True
				If InStr(strQueryChoices, "3") > 0 Then booProcesses = True
				If InStr(strQueryChoices, "4") > 0 Then booStartup = True
				If InStr(strQueryChoices, "5") > 0 Then booServices = True
				
				Err.Clear
				
				Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
				strPC & "\root\cimv2") 
				Set objWMIService2 = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2") 
				
				If Err.Number <> 0 Then
					strError = "Unable to access WMI Repository for " & strPC & vbCrLf & _
					vbCrLf & " Please make sure you have the required privileges to access this PC"
					MsgBox strError,vbExclamation,"Access Error"
					CleanUp
					Exit Sub
				End If
				
				booWMIPrompt = objShell.RegRead(strRegStart & "\Software\SKB\PCManagementUtil\Others\booWMIPrompt")
				If booWMIPrompt = "" OR IsNull(booWMIPrompt) Then 
					objShell.RegWrite strRegStart & "\Software\SKB\PCManagementUtil\Others\booWMIPrompt", _
					"1", "REG_SZ"
					booWMIPrompt = "1"
				End If
				Set colPing = objWMIService2.ExecQuery _
					("Select * from Win32_PingStatus Where Address = '" & strPC & "'")
						
				For Each objItem in colPing
					intResponseTime = objItem.ResponseTime
					If intResponseTime = 0 Then
						strResponseTime = "<1ms"
						Else
							strResponseTime = intResponseTime & "ms"
					End If
				Next
				
				If booWMIPrompt = "1" AND booRescan = False AND intResponseTime > 20 Then
					booSlowWMIPrompt = MsgBox("The WMI connection to " & strPC & " is quite slow." & vbCrLf & vbCrLf & _
					"Do you wish to continue?",vbInformation+vbYesNo,"PC Management Utility")
					If booSlowWMIPrompt = vbNo Then
						CleanUp
						Exit Sub
					End If
				End If
				
				Set colComputer = objWMIService.ExecQuery _
					("Select * from Win32_ComputerSystem")
			
				For Each objItem In colComputer
					strRemoteLoggedOn = objItem.UserName
					strLoggedOn = objItem.UserName
					strManufacturer = objItem.Manufacturer
					strModel = objItem.Model
					intMemSize = Round(objItem.TotalPhysicalMemory / 1073741824,2)	
				Next
				
				strRemoteSID = GetSIDFromUser(strLoggedOn)

				Set colNAC = objWMIService.ExecQuery _
					("Select * from Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
				
				intIpCount = 0
					
				For Each objItem in colNAC
					If IsNull(objItem.IPAddress) Then
						strIP = ""
						Else
							strIP = objItem.IPAddress(0)
					End If
					
					If strIP <> "0.0.0.0" AND strIP <> "" Then
						If intIpCount > 0  Then
							strResolveIP = strResolveIP & "; " & strIP
							If objItem.DHCPEnabled = "True" Then
								strResolveIP = strResolveIP & " (DHCP enabled)"
								Else
									strResolveIP = strResolveIP & " (Static IP)"
							End If
							Else
								strResolveIP = strIP
								If objItem.DHCPEnabled = "True" Then
									strResolveIP = strResolveIP & " (DHCP enabled)"
									Else
										strResolveIP = strResolveIP & " (Static IP)"
								End If
						End If
						intIpCount = intIpCount + 1
					End If
					strIP = ""
				Next

				If booRescan = False Then
					DataArea.InnerHTML = "<h3><i>Fetching info for " & strPC & ", please wait..</i></h3>"
					Else
						DataArea.InnerHTML = "<h3><i>Rescanning " & strPC & ", please wait..</i></h3>"
				End If
				PauseScript(1)
				
				If IsNull(strRemoteLoggedOn) OR strRemoteLoggedOn = "" Then
					strRemoteLoggedOn = "Not logged in"
				End If
			
				Set colOS = objWMIService.ExecQuery _
					("Select * from Win32_OperatingSystem")
			
				For Each objItem In colOS
					strOS = objItem.Caption
					intServicePackMajor = objItem.ServicePackMajorVersion
					dtmBootDate = FormatDate(objItem.LastBootUpTime)
					strUptime = TimeSpan(dtmBootDate,Now)
					strDescription = Trim(objItem.Description)
					strOSSerial = objItem.SerialNumber
					dtmInstallDate = FormatDate(objItem.InstallDate)
				Next
				
				Set colBIOS = objWMIService.ExecQuery _
					("Select * from Win32_BIOS")
		
				For Each objItem In colBIOS
					strSerial = objItem.SerialNumber
				Next
				
				If booRescan = False Then
					DataArea.InnerHTML = "<h3><i>Fetching info for " & strPC & ", please wait...</i></h3>"
					Else
						DataArea.InnerHTML = "<h3><i>Rescanning " & strPC & ", please wait...</i></h3>"
				End If
				PauseScript(1)
			
				strArchitecture = CheckWinArchitecture()
				
				If booRescan = False Then
					DataArea.InnerHTML = "<h3><i>Fetching info for " & strPC & ", please wait....</i></h3>"
					Else
						DataArea.InnerHTML = "<h3><i>Rescanning " & strPC & ", please wait....</i></h3>"
				End If
				PauseScript(1)
				
				strPCInfoTab = "<b><u>" & strPC & "</u></b><p>" & _
				"<table width=""100%"">" & _
					"<tr><td width=""30%"" style=""padding-bottom:15px"">Logged On User:</td><td style=""padding-bottom:15px"">" & _
					strRemoteLoggedOn & "</td></tr>" & _
					"<tr><td>IP Address(es):</td><td>" & strResolveIP & "</td></tr>" & _
					"<tr><td style=""padding-bottom:15px"">Response Time:</td><td style=""padding-bottom:15px"">" & _
					strResponseTime & "</td></tr>" & _
					"<tr><td>Description:</td><td>" & strDescription & "</td></tr>" & _
					"<tr><td>Manufacturer:</td><td>" & strManufacturer & "</td></tr>" & _
					"<tr><td>Model:</td><td>" & strModel & "</td></tr>" & _
					"<tr><td>Serial Number:</td><td>" & strSerial & "</td></tr>" & _
					"<tr><td style=""padding-bottom:15px"">RAM:</td><td style=""padding-bottom:15px"">" & _
					intMemSize & " GB</td></tr>" & _
					"<tr><td>OS:</td><td>" & strOS & "</td></tr>" & _
					"<tr><td>Product ID:</td><td>" & strOSSerial & "</td></tr>" & _
					"<tr><td>Architecture:</td><td>" & strArchitecture & "</td></tr>" & _
					"<tr><td>Service Pack:</td><td>" & intServicePackMajor & "</td></tr>" & _
					"<tr><td style=""padding-bottom:15px"">Install Date:</td><td style=""padding-bottom:15px"">" & _
					dtmInstallDate & "</td></tr>" & _
					"<tr><td>Last Reboot:</td><td>" & dtmBootDate & "</td></tr>" & _
					"<tr><td>System Uptime:</td><td>" & strUptime & "</td></tr>" & _
				"</table>" & _
				"<hr><input id=""btnCopyPCInfo"" style=""width:75px;"" class=""button"" type=""button"" value=""Copy"" " & _
				"name=""btnCopyPCInfo"" title=""Copy PC Info to clipboard"" onclick=""CopyPCInfo()"">" & _
				"<input id=""PrintButton"" style=""width:75px;"" class=""button"" type=""button"" value=""Print"" " & _
				"name=""PrintButton"" title=""Print Window"" onclick=""Window.Print()"">"
			
				strCopy = strPC & "}{" & strRemoteLoggedOn & "}{" & strResolveIP & "}{" & strDescription & _
				"}{" & strManufacturer & "}{" & strModel & "}{" & strSerial & "}{" & intMemSize & "}{" & strOS & _
				"}{" & strOSSerial & "}{" & strArchitecture & "}{" & intServicePackMajor & "}{" & dtmInstallDate & _
				"}{" & dtmBootDate & "}{" & strUptime & "}{" & strResponseTime
								
				tab1.Disabled = False
				tab1.Title = "View information about the PC"
				
				If booSoftware = True Then ShowSoftwareInfo False
				tab2.Disabled = False
				tab2.Title = "View the software items installed on the PC and uninstall any as required"
				If booProcesses = True Then ShowProcessInfo False
				tab3.Disabled = False
				tab3.Title = "View the processes running on the PC and kill any as required"
				If booServices = True Then ShowServiceInfo False
				tab4.Disabled = False
				tab4.Title = "View the Services running on the PC and stop / start any as required"
				If booStartup = True Then ShowStartupInfo False
				tab5.Disabled = False
				tab5.Title = "View the startup items on the PC and remove any as required"
				Else
					ErrConnectPrompt = MsgBox("Error connecting to " & strPC & vbCrLf & vbCrLf & _
					"Would you like to run a continuous ping to this machine?",vbExclamation+vbYesNo,"Error")
					Err.Clear
					If ErrConnectPrompt = vbYes Then
						btnRescan.Disabled = False
						btnRescan.Title = "Rescan the selected PC"
						DataArea.InnerHTML = "<b><u>" & strPC & "</u></b><p><div align=""center"">" & _
						"<span id=""WaitMessage""></span></div>" & _
						"<input id=runbutton  class=""button"" type=""button"" value=""Stop"" " & _
						"name=""btnStop"" title=""Stop running action"" onclick=""StopAction()"">"
						PingMachine False, True
						Else
							CleanUp
							Exit Sub
					End If
			End If
		End If
		DataArea.InnerHTML = strPCInfoTab
		tab1.bgcolor="#cccccc"
		tab2.bgcolor="#eeeeee"
		tab3.bgcolor="#eeeeee"
		tab4.bgcolor="#eeeeee"
		tab5.bgcolor="#eeeeee"
		tab6.bgcolor="#eeeeee"
		tab6.Disabled = False
		tab6.Title = "Perform actions on the PC"
		tab1.style.cursor = "pointer"
		tab2.style.cursor = "pointer"
		tab3.style.cursor = "pointer"
		tab4.style.cursor = "pointer"
		tab5.style.cursor = "pointer"
		tab6.style.cursor = "pointer"
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	CopyPCInfo()
    '#	PURPOSE........:	Copies the PC Info to the clipboard
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub CopyPCInfo()
		arrCopyInfo = Split(strCopy, "}{")
		strCopyString = arrCopyInfo(0) & vbCrLf & vbCrLf & _
		"Logged On User:" & String(20 - Len("Logged On User:"), " ") & arrCopyInfo(1) & vbCrlf & vbCrLf & _
		"IP Address(es):" & String(20 - Len("IP Address(es):"), " ") & arrCopyInfo(2) & vbCrlf & _
		"Response Time:" & String(20 - Len("Response Time:"), " ") & arrCopyInfo(15) & vbCrlf & vbCrLf & _
		"Description:" & String(20 - Len("Description:"), " ") & arrCopyInfo(3) & vbCrlf & _
		"Manufacturer:" & String(20 - Len("Manufacturer:"), " ") & arrCopyInfo(4) & vbCrlf & _
		"Model:" & String(20 - Len("Model:"), " ") & arrCopyInfo(5) & vbCrlf & _
		"Serial Number:" & String(20 - Len("Serial Number:"), " ") & arrCopyInfo(6) & vbCrlf & _
		"RAM:" & String(20 - Len("RAM:"), " ") & arrCopyInfo(7) & " GB" & vbCrlf & vbCrLf & _
		"OS:" & String(20 - Len("OS:"), " ") & arrCopyInfo(8) & vbCrlf & _
		"Product ID:" & String(20 - Len("Product ID:"), " ") & arrCopyInfo(9) & vbCrlf & _
		"Architecture:" & String(20 - Len("Architecture:"), " ") & arrCopyInfo(10) & vbCrlf & _
		"Service Pack:" & String(20 - Len("Service Pack:"), " ") & arrCopyInfo(11) & vbCrlf & _
		"Install Date:" & String(20 - Len("Install Date:"), " ") & arrCopyInfo(12) & vbCrlf & vbCrlf & _
		"Last Reboot:" & String(20 - Len("Last Reboot:"), " ") & arrCopyInfo(13) & vbCrlf & _
		"System Uptime:" & String(20 - Len("System Uptime:"), " ") & arrCopyInfo(14)

		Document.parentwindow.clipboardData.SetData "text", strCopyString
		MsgBox "The info has now been copied to the clipboard", vbInformation, "PC Management Utility"
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ShowSoftwareInfo(booChoice)
    '#	PURPOSE........:	Displays the Software tab
    '#	ARGUMENTS......:	booChoice = boolean value to determine whether it
	'#						was called via Tab (True) or initial query (False)
    '#	EXAMPLE........:	ShowSoftwareInfo(True)
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ShowSoftwareInfo(booChoice)
		If booChoice = False Then
			On Error Resume Next
			
			DataArea.InnerHTML = "<h3><i>Fetching Software info for " & strPC & ", please wait.</i></h3>"
			PauseScript(1)
			
			booSoftwareVersion = objShell.RegRead(strRegStart & "\Software\SKB\PCManagementUtil\Tabs\booSoftwareVersion")
			booSoftwareVendor = objShell.RegRead(strRegStart & "\Software\SKB\PCManagementUtil\Tabs\booSoftwareVendor")
			booSoftwareInstallDate = objShell.RegRead(strRegStart & _
			"\Software\SKB\PCManagementUtil\Tabs\booSoftwareInstallDate")
			If booSoftwareVersion = "" OR IsNull(booSoftwareVersion) Then 
				objShell.RegWrite strRegStart & "\Software\SKB\PCManagementUtil\Tabs\booSoftwareVersion", _
				"1", "REG_SZ"
				booSoftwareVersion = "1"
			End If
			If booSoftwareVendor = "" OR IsNull(booSoftwareVendor) Then 
				objShell.RegWrite strRegStart & "\Software\SKB\PCManagementUtil\Tabs\booSoftwareVendor", _
				"1", "REG_SZ"
				booSoftwareVendor = "1"
			End If		
			If booSoftwareInstallDate = "" OR IsNull(booSoftwareInstallDate) Then 
				objShell.RegWrite strRegStart & "\Software\SKB\PCManagementUtil\Tabs\booSoftwareInstallDate", _
				"1", "REG_SZ"
				booSoftwareInstallDate = "1"
			End If			
			
			strPath = objShell.ExpandEnvironmentStrings("%path%")
			
			strHTML = "<b><u>" & strPC & "</u></b><br>"
			strHTML = strHTML & "<span id=""NumSoftwareItems"" style=""width:40%;font-size:0.9em;text-align:right;font-style:italic;font-weight:bold;""></span><br>" 
			strHTML = strHTML & "<table width=""100%"">" 
			strHTML = strHTML & 	"<tr>" 
			strHTML = strHTML & 		"<td width=""40%"">" 
			strHTML = strHTML & 			"<select size=""28"" name=""SoftwareItems"" style=""width:250"" " 
			strHTML = strHTML & 			"onChange=""GetSoftwareDetails()""></select><br>" 
			strHTML = strHTML & 			"<span style=""float:left;"">"
			strHTML = strHTML & 			"<input id=""SWRefreshButton"" class=""button"" type=""button"" " 
			strHTML = strHTML & 			"title=""Refresh software list"" value=""Refresh"" " 
			strHTML = strHTML & 			"name=""SWRefreshButton"" onclick=""RefreshSoftware()""></span>" 
			strHTML = strHTML & 			"<span style=""float:right;"">"
			strHTML = strHTML & "			<select name=""SWExport"" "
			strHTML = strHTML & "			title=""Export the software list"" onChange=""ExportSoftwareDetails()"">"
			strHTML = strHTML & "				<option value=""0"">Export to:</option>"
			strHTML = strHTML & "				<option value=""1"" title=""Export the software list to a Comma " & _
			"Seperated Values (csv) file"")>Export to csv</option>"
			strHTML = strHTML & "				<option value=""2"" title=""Export the software list to a formatted Excel " & _
			"(xls) spreadsheet"">Export to xls</option>"
			strHTML = strHTML & "				<option value=""3"" title=""Export the software list to a Web " & _
			"page (html) file"">Export to html</option>"
			strHTML = strHTML & "				<option value=""4"" title=""Export the software list to a Text " & _
			"(txt) file"">Export to txt</option>"
			strHTML = strHTML & "			</select></span>"
			strHTML = strHTML & 		"</td>" 
			strHTML = strHTML & 		"<td style=""vertical-align:top;"">" 
			strHTML = strHTML & 			"<table width=""100%"" style=""table-layout:fixed;"">" 
			strHTML = strHTML & 				"<tr>" 
			strHTML = strHTML & 					"<td width=""35%"">" 
			strHTML = strHTML & 						"<b>Software Name: </b>"  
			strHTML = strHTML & 					"</td>" 
			strHTML = strHTML & 					"<td width=""65%"">" 
			strHTML = strHTML & 						"<span id=""SoftwareName""></span><br>" 
			strHTML = strHTML & 					"</td>" 
			strHTML = strHTML & 				"</tr>" 
			
			If booSoftwareVersion = "1" Then
				strHTML = strHTML & 				"<tr>" 
				strHTML = strHTML & 					"<td>" 
				strHTML = strHTML & 						"<b>Software Version: </b>" 
				strHTML = strHTML & 					"</td>" 
				strHTML = strHTML & 					"<td>" 
				strHTML = strHTML & 						"<span id=""SoftwareVersion""></span><br>" 
				strHTML = strHTML & 					"</td>" 
				strHTML = strHTML & 				"</tr>" 
			End If
			
			If booSoftwareVendor = "1" Then
				strHTML = strHTML & 				"<tr>" 
				strHTML = strHTML & 					"<td>" 
				strHTML = strHTML & 						"<b>Vendor: </b>" 
				strHTML = strHTML & 					"</td>" 
				strHTML = strHTML & 					"<td>" 
				strHTML = strHTML & 						"<span id=""SoftwareVendor""></span><br>" 
				strHTML = strHTML & 					"</td>" 
				strHTML = strHTML & 				"</tr>" 
			End If
			
			If booSoftwareInstallDate = "1" Then
				strHTML = strHTML & 				"<tr>" 
				strHTML = strHTML & 					"<td>" 
				strHTML = strHTML & 						"<b>Install Date:</b>" 
				strHTML = strHTML & 					"</td>" 
				strHTML = strHTML & 					"<td>" 
				strHTML = strHTML & 						"<span id=""SoftwareDate""></span>" 
				strHTML = strHTML & 					"</td>" 
				strHTML = strHTML & 				"</tr>" 
			End If
			
			strHTML = strHTML & 				"<tr>"
			strHTML = strHTML & 					"<td colspan=""2"">" 
			strHTML = strHTML & 						"&nbsp;" 
			strHTML = strHTML & 					"</td>" 
			strHTML = strHTML & 				"</tr>" 
			strHTML = strHTML & 				"<tr>" 
			strHTML = strHTML & 					"<td colspan=""2"">" 
			strHTML = strHTML & 						"<b>Uninstall String:</b>" 
			strHTML = strHTML & 					"</td>" 
			strHTML = strHTML & 				"</tr>" 
			strHTML = strHTML & 				"<tr>" 
			strHTML = strHTML & 					"<td colspan=""2"" style=""word-break:break-all;"">"
			strHTML = strHTML & 						"<span id=""SoftwareUninstallString""></span>" 
			strHTML = strHTML & 					"</td>" 
			strHTML = strHTML & 				"</tr>" 
			strHTML = strHTML & 				"<tr>"
			strHTML = strHTML & 					"<td colspan=""2"">" 
			strHTML = strHTML & 						"&nbsp;" 
			strHTML = strHTML & 					"</td>" 
			strHTML = strHTML & 				"</tr>" 
			strHTML = strHTML & 				"<tr>" 
			strHTML = strHTML & 					"<td colspan=""2"">" 
			strHTML = strHTML & 						"<b>Silent String:</b>" 
			strHTML = strHTML & 					"</td>" 
			strHTML = strHTML & 				"</tr>" 
			strHTML = strHTML & 				"<tr>" 
			strHTML = strHTML & 					"<td colspan=""2"" style=""word-break:break-all;"">"
			strHTML = strHTML & 						"<span id=""SoftwareSilentString""></span>" 
			strHTML = strHTML & 					"</td>" 
			strHTML = strHTML & 				"</tr>" 
			strHTML = strHTML & 				"<tr>"
			strHTML = strHTML & 					"<td colspan=""2"">" 
			strHTML = strHTML & 						"&nbsp;" 
			strHTML = strHTML & 					"</td>" 
			strHTML = strHTML & 				"</tr>" 
			strHTML = strHTML & 				"<tr>" 
			strHTML = strHTML & 					"<td>" 
			strHTML = strHTML & 						"<input id=""btnUninstall"" class=""button"" type=""button"" " 
			strHTML = strHTML & 						"value=""Uninstall App"" " 
			strHTML = strHTML & 						"name=""btnUninstall"" disabled=true onclick=""UninstallSoftware False"">" 
			strHTML = strHTML & 					"</td>" 
			strHTML = strHTML & 					"<td>" 
			strHTML = strHTML & 						"<input id=""btnSilentUninstall"" class=""button"" type=""button"" " 
			strHTML = strHTML & 						"value=""Silent Uninstall"" " 
			strHTML = strHTML & 						"name=""btnSilentUninstall"" disabled=true onclick=""UninstallSoftware True"">" 
			strHTML = strHTML & 					"</td>" 
			strHTML = strHTML & 				"</tr>" 
			strHTML = strHTML & 				"<tr>"
			strHTML = strHTML & 					"<td colspan=""2"">" 
			strHTML = strHTML & 						"&nbsp;" 
			strHTML = strHTML & 					"</td>" 
			strHTML = strHTML & 				"</tr>" 
			strHTML = strHTML & 				"<tr>"
			strHTML = strHTML & 					"<td colspan=""2"">" 
			strHTML = strHTML & 						"&nbsp;" 
			strHTML = strHTML & 					"</td>" 
			strHTML = strHTML & 				"</tr>" 
			strHTML = strHTML & 				"<tr>" 
			strHTML = strHTML & 					"<td colspan=""2"">" 
			strHTML = strHTML & 						"<i>Please note, you must have either <a href=""http://sites.google.com/site/4utils/projects"" " 
			strHTML = strHTML & 						"target=""_blank"" title=""Download Rctrlx""><b>Rctrlx</b></a> or <a href=""" 
			strHTML = strHTML & 						"http://technet.microsoft.com/en-us/sysinternals/bb897553.aspx"" title=""Download PSExec"" " 
			strHTML = strHTML & 						"target=""_blank""><b>PSExec</b></a> in your <u><a title=""System Path: " & strPath 
			strHTML = strHTML & 						""">system path</a></u> to be able to remotely uninstall software.</i>" 
			strHTML = strHTML & 					"</td>" 
			strHTML = strHTML & 				"</tr>" 
			strHTML = strHTML & 				"<tr>" 
			strHTML = strHTML & 			"</table>" 
			strHTML = strHTML & 		"</td>" 
			strHTML = strHTML & 	"</tr>" 
			strHTML = strHTML & "</table>"

			SoftwareTab.InnerHTML = strHTML
			
			PopulateSoftware()
			Else
				If IsNull(SoftwareTab.InnerHTML) OR SoftwareTab.InnerHTML = "" Then
					ShowSoftwareInfo False
				End If
				
				tab1.bgcolor="#eeeeee"
				tab2.bgcolor="#cccccc"
				tab3.bgcolor="#eeeeee"
				tab4.bgcolor="#eeeeee"
				tab5.bgcolor="#eeeeee"
				tab6.bgcolor="#eeeeee"
				DataArea.InnerHTML = SoftwareTab.InnerHTML
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	PopulateSoftware()
    '#	PURPOSE........:	Populates the software items in the Software tab
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub PopulateSoftware()
		On Error Resume Next
		
		Set DataList = CreateObject("ADOR.Recordset")
		DataList.Fields.Append "Text", adVarChar, MaxCharacters
		DataList.Fields.Append "Value", adVarChar, MaxCharacters
		DataList.Open
		
		booSortSoftware = objShell.RegRead(strRegStart & "\Software\SKB\PCManagementUtil\Tabs\booSortSoftware")
		booShowSoftware = objShell.RegRead(strRegStart & "\Software\SKB\PCManagementUtil\Tabs\booShowSoftware")
		If booShowSoftware = "" OR IsNull(booShowSoftware) Then 
			objShell.RegWrite strRegStart & "\Software\SKB\PCManagementUtil\Tabs\booShowSoftware", _
			"0", "REG_SZ"
			booShowSoftware = "0"
		End If
		If booSortSoftware = "" OR IsNull(booSortSoftware) Then 
			objShell.RegWrite strRegStart & "\Software\SKB\PCManagementUtil\Tabs\booSortSoftware", _
			"1", "REG_SZ"
			booSortSoftware = "1"
		End If
		
		DataArea.InnerHTML = "<h3><i>Fetching Software info for " & strPC & ", please wait..</i></h3>"
		PauseScript(1)
		
		Set objReg = GetObject("winmgmts://" & strPC & "/root/default:StdRegProv")
		strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
		objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys
		For Each objItem In arrSubKeys
			strValueName = "DisplayName"
			strSubPath = strKeyPath & "\" & objItem
			objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath,strValueName,strValue
			
			If strValue <> "" AND InStr(strValue, "Hotfix") = 0 AND _
			InStr(strValue, "Security Update") = 0 AND _
			InStr(strValue, "Update for Windows") = 0 Then
				booHide = 0
				objReg.GetDwordValue HKEY_LOCAL_MACHINE,strSubPath, _
				"SystemComponent",booHide
				If booHide <> 1 OR IsNull(booHide) OR booHide = "" OR booShowSoftware = "1" Then
					strName = strValue
					objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath, _
					"DisplayVersion",strVersion
					objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath, _
					"InstallDate",intInstallDate
					objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath, _
					"Publisher",strVendor
					objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath, _
					"UninstallString",strUninstallString
					objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath, _
					"QuietUninstallString",strSilentString
					
					If IsNull(intInstallDate) OR intInstallDate = "" Then
						dtmInstallDate = "&nbsp;"
						Else 
							dtmInstallDate = MID(intInstallDate,7,2) & _
							"/" & MID(intInstallDate,5,2) & "/" & _
							LEFT(intInstallDate,4)
							If NOT IsDate(dtmInstallDate) Then
								dtmInstallDate = "&nbsp;"
							End If
					End If
					If IsNull(strName) OR strName = "" Then
						strSoftwareName = "&nbsp;"
					End If
					If IsNull(strVendor) OR strVendor = "" Then
						strVendor = "&nbsp;"
					End If
					If IsNull(strVersion) OR strVersion = "" Then
						strVersion = "&nbsp;"
					End If
					If IsNull(strUninstallString) OR strUninstallString = "" Then
						strUninstallString = "&nbsp;"
					End If
					
					If InStr(Lcase(strUninstallString), "msiexec.exe") > 0 Then
						strSilentString = strUninstallString & " /qn /norestart"
					End If
					
					If IsNull(strSilentString) OR strSilentString = "" Then
						strSilentString = "&nbsp;"
					End If

					DataList.AddNew
					DataList("Text") = strName
					DataList("Value") = strName & "|" & strVendor & _
					"|" & strVersion & "|" & dtmInstallDate & _
					"|" & strUninstallString & "|" & strSilentString
					
					If Err.Number <> 0 Then
						DataList("Value") = strName & "|" & strVendor & _
						"|" & strVersion & "|" & dtmInstallDate & _
						"|&nbsp;|&nbsp;"
						Err.Clear
					End If
					
					DataList.Update 
				End If
			End If
		Next
		
		DataArea.InnerHTML = "<h3><i>Fetching Software info for " & strPC & ", please wait...</i></h3>"
		PauseScript(1)
		
		strKeyPath = strRemoteSID & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
		objReg.EnumKey HKEY_USERS, strKeyPath, arrSubkeys
		For Each objItem In arrSubKeys
			strValueName = "DisplayName"
			strSubPath = strKeyPath & "\" & objItem
			objReg.GetExpandedStringValue HKEY_USERS,strSubPath,strValueName,strValue
			
			If strValue <> "" AND InStr(strValue, "Hotfix") = 0 AND _
			InStr(strValue, "Security Update") = 0 AND _
			InStr(strValue, "Update for Windows") = 0 Then
				booHide = 0
				objReg.GetDwordValue HKEY_LOCAL_MACHINE,strSubPath, _
				"SystemComponent",booHide
				If booHide <> 1 OR IsNull(booHide) OR booHide = "" OR booShowSoftware = "1" Then
					strName = strValue
					objReg.GetExpandedStringValue HKEY_USERS,strSubPath, _
					"DisplayVersion",strVersion
					objReg.GetExpandedStringValue HKEY_USERS,strSubPath, _
					"InstallDate",intInstallDate
					objReg.GetExpandedStringValue HKEY_USERS,strSubPath, _
					"Publisher",strVendor
					objReg.GetExpandedStringValue HKEY_USERS,strSubPath, _
					"UninstallString",strUninstallString
					objReg.GetExpandedStringValue HKEY_USERS,strSubPath, _
					"QuietUninstallString",strSilentString
					If IsNull(intInstallDate) OR intInstallDate = "" Then
						dtmInstallDate = "&nbsp;"
						Else 
							dtmInstallDate = MID(intInstallDate,7,2) & _
							"/" & MID(intInstallDate,5,2) & "/" & _
							LEFT(intInstallDate,4)
							If NOT IsDate(dtmInstallDate) Then
								dtmInstallDate = "&nbsp;"
							End If
					End If
					If IsNull(strName) OR strName = "" Then
						strSoftwareName = "&nbsp;"
					End If
					If IsNull(strVendor) OR strVendor = "" Then
						strVendor = "&nbsp;"
					End If
					If IsNull(strVersion) OR strVersion = "" Then
						strVersion = "&nbsp;"
					End If
					If IsNull(strUninstallString) OR strUninstallString = "" Then
						strUninstallString = "&nbsp;"
					End If
					
					If InStr(Lcase(strUninstallString), "msiexec.exe") > 0 Then
						strSilentString = strUninstallString & " /qn /norestart"
					End If
					
					If IsNull(strSilentString) OR strSilentString = "" Then
						strSilentString = "&nbsp;"
					End If
					
					DataList.AddNew
					DataList("Text") = strName
					DataList("Value") = strName & "|" & strVendor & _
					"|" & strVersion & "|" & dtmInstallDate & _
					"|" & strUninstallString & "|" & strSilentString
					
					If Err.Number <> 0 Then
						DataList("Value") = strName & "|" & strVendor & _
						"|" & strVersion & "|" & dtmInstallDate & _
						"|&nbsp;|&nbsp;"
						Err.Clear
					End If
					
					DataList.Update 
				End If
			End If
		Next

		If booSortSoftware = "1" Then DataList.Sort = "Text"
		
		DataArea.InnerHTML = "<h3><i>Fetching Software info for " & strPC & ", please wait....</i></h3>"
		PauseScript(1)
		
		DataList.MoveFirst
		Do Until DataList.EOF
			Set objOption = Document.createElement("OPTION")
			objOption.Text = DataList.Fields.Item("Text")
			objOption.Title = DataList.Fields.Item("Text")
			objOption.Value = DataList.Fields.Item("Value")
			SoftwareItems.Add(objOption)
			DataList.MoveNext
		Loop

		NumSoftwareItems.InnerHTML = SoftwareItems.Options.Length & " Items"
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	GetSoftwareDetails()
    '#	PURPOSE........:	Displays the details for the Software Item selected
	'#						in the Software tab
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub GetSoftwareDetails()
		On Error Resume Next
		SoftwareTab.InnerHTML = Nothing
		btnUninstall.Disabled = True
		btnUninstall.style.cursor = "default"
		btnUninstall.Title = ""
		btnSilentUninstall.Disabled = True
		btnSilentUninstall.style.cursor = "default"
		btnSilentUninstall.Title = ""
		arrSoftwareDetails = Split(SoftwareItems.Value, "|")
		SoftwareName.InnerHTML = arrSoftwareDetails(0)
		SoftwareVendor.InnerHTML = arrSoftwareDetails(1)
		SoftwareVersion.InnerHTML = arrSoftwareDetails(2)
		SoftwareDate.InnerHTML =  arrSoftwareDetails(3)
		strUninstallString = arrSoftwareDetails(4)
		strSilentString = arrSoftwareDetails(5)
		If strUninstallString <> "&nbsp;" AND strUninstallString <> "" Then
			If InStr(LCase(strUninstallString), "msiexec") > 0 Then _
				strUninstallString = Replace(LCase(strUninstallString), "/i", "/x")
			btnUninstall.Disabled = False
			btnUninstall.Title = "Uninstall selected application on PC interactively"
		End If
		
		If strSilentString <> "&nbsp;" AND strSilentString <> "" Then
			If InStr(LCase(strSilentString), "msiexec") > 0 Then _
				strSilentString = Replace(LCase(strSilentString), "/i", "/x")
			btnSilentUninstall.Disabled = False
			btnSilentUninstall.Title = "Uninstall selected application on PC silently"
		End If
		
		SoftwareSilentString.InnerHTML = strSilentString
		SoftwareUninstallString.InnerHTML = strUninstallString
		SoftwareTab.InnerHTML = DataArea.InnerHTML
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ExportSoftwareDetails()
    '#	PURPOSE........:	Export the details for the Software Items
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ExportSoftwareDetails()
		On Error Resume Next
		SoftwareTab.InnerHTML = Nothing
		
		Select Case SWExport.Value
			Case 1
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\SoftwareDetails" & strPC & ".csv",True)
				objFile.WriteLine "Software Items on " & strPC
				objFile.WriteLine ""
				objFile.WriteLine "Total: " & NumSoftwareItems.InnerHTML
				objFile.WriteLine ""
				objFile.WriteLine "Name,Vendor,Version,Install Date,Uninstall String,Silent String"
			Case 2
				Const xlContinuous = 1
				Const xlThin = 2
				Const xlAutomatic = -4105
				
				strExcelPath = objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe\")
			   
				If strExcelPath = "" Then
					MsgBox "Unable to export. Excel does not appear to be installed.", vbExclamation, "PC Management Utility"
					Exit Sub
				End If
				
				Set objExcel = CreateObject("Excel.Application")
				objExcel.Visible = False
				Set objWorkBook = objExcel.WorkBooks.Add
				Set objWorksheet = objWorkbook.Worksheets(1)
				objExcel.DisplayAlerts = False
				For i = 1 to 3
					objWorkbook.Worksheets(2).Delete
				Next
				objExcel.DisplayAlerts = True
				objWorksheet.Name = "Software Details"
				
				objWorkSheet.Cells(1, 1) = "Software Items on " & strPC
				objWorkSheet.Cells(3, 1) = "Total: " & NumSoftwareItems.InnerHTML

				intStartRow = 6
				
				objWorkSheet.Cells(5, 1) = "Name"
				objWorkSheet.Cells(5, 2) = "Vendor"
				objWorkSheet.Cells(5, 3) = "Version"
				objWorkSheet.Cells(5, 4) = "Install Date"
				objWorkSheet.Cells(5, 5) = "Uninstall String"
				objWorkSheet.Cells(5, 6) = "Silent String"
			Case 3
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\SoftwareDetails" & strPC & ".htm",True)
				objFile.WriteLine "<style type=""text/css"">"
				objFile.WriteLine "body{background-color:#CEF0FF;}"
				objFile.WriteLine "table.export{border-width:1px;border-spacing:1px;border-style:solid;border-color:gray;border-collapse:collapse;}"
				objFile.WriteLine "table.export th{border-width:1px;padding:1px;border-style:solid;border-color:gray;padding:2px 7px 2px 7px;}"
				objFile.WriteLine "table.export td{border-width:1px;padding:1px;border-style:dotted;border-color:gray;padding:2px 7px 2px 7px;}"
				objFile.WriteLine ".backtotop a {font-size:0.9em;}"
				objFile.WriteLine "</style>"
				objFile.WriteLine "<div style=""font-weight:bold;""><a name =""top"">Software Items on " & strPC & "</a><p>"
				objFile.WriteLine "Total: " & NumSoftwareItems.InnerHTML & "<p></div>"
				objFile.WriteLine "<table class=""export"">"
				objFile.WriteLine "	<tr>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Name"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th>"
				objFile.WriteLine "			Google"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Vendor"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Version"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th>"
				objFile.WriteLine "			Install Date"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Uninstall String"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Silent String"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "	</tr>"
			Case 4
				intColumnIndex = 9
				intColumnIndex2 = 11
				intColumnIndex3 = 12
				For Each objOption in SoftwareItems.Options
					arrSoftwareDetails = Split(objOption.Value, "|")
					strSoftwareName = arrSoftwareDetails(0)
					If strSoftwareName = "&nbsp;" Then strSoftwareName = ""
					strSoftwareVendor = arrSoftwareDetails(1)
					If strSoftwareVendor = "&nbsp;" Then strSoftwareVendor = ""
					strSoftwareVersion = arrSoftwareDetails(2)
					If strSoftwareVersion = "&nbsp;" Then strSoftwareVersion = ""
					If Len(strSoftwareName) > intColumnIndex - 5 Then intColumnIndex = Len(strSoftwareName) + 5
					If Len(strSoftwareVendor) > intColumnIndex2 - 5 Then intColumnIndex2 = Len(strSoftwareVendor) + 5
					If Len(strSoftwareVersion) > intColumnIndex3 - 5 Then intColumnIndex3 = Len(strSoftwareVersion) + 5
				Next
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\SoftwareDetails" & strPC & ".txt",True)
				objFile.WriteLine "Software Items on " & strPC
				objFile.WriteLine ""
				objFile.WriteLine "Total: " & NumSoftwareItems.InnerHTML
				objFile.WriteLine ""
				objFile.WriteLine "Name" & _
				String(intColumnIndex - 4, " ") & "Vendor" & _
				String(intColumnIndex2 - 6, " ") & "Version" & _
				String(intColumnIndex3 - 7, " ") & "Install Date"
		End Select
		
		For Each objOption in SoftwareItems.Options
			arrSoftwareDetails = Split(objOption.Value, "|")
			strSoftwareName = arrSoftwareDetails(0)
			strSoftwareVendor = arrSoftwareDetails(1)
			strSoftwareVersion = arrSoftwareDetails(2)
			dtmSoftwareDate =  arrSoftwareDetails(3)
			strUninstallString =  arrSoftwareDetails(4)
			strSilentString =  arrSoftwareDetails(5)

			If strSoftwareName = "&nbsp;" Then strSoftwareName = ""
			If strSoftwareVendor = "&nbsp;" Then strSoftwareVendor = ""
			If strSoftwareVersion = "&nbsp;" Then strSoftwareVersion = ""
			If dtmSoftwareDate = "&nbsp;" Then dtmSoftwareDate = ""
			If strUninstallString = "&nbsp;" Then strUninstallString = ""
			If strSilentString = "&nbsp;" Then strSilentString = ""
			
			If IsDate(dtmSoftwareDate) Then dtmSoftwareDate = CDate(dtmSoftwareDate)
			
			Select Case SWExport.Value
				Case 1
					strSoftwareName = EncodeCsv(strSoftwareName)
					strSoftwareVendor = EncodeCsv(strSoftwareVendor)
					strSoftwareVersion = EncodeCsv(strSoftwareVersion)
					dtmSoftwareDate = EncodeCsv(dtmSoftwareDate)
					strUninstallString = EncodeCsv(strUninstallString)
					strSilentString = EncodeCsv(strSilentString)
					
					strCSV = strCSV & strSoftwareName & "," & _
					strSoftwareVendor & "," & strSoftwareVersion & "," & _
					dtmSoftwareDate & "," & strUninstallString & "," & _
					strSilentString & vbCrLf
				Case 2
					objWorkSheet.Cells(intStartRow, 1) = strSoftwareName
					objWorkSheet.Cells(intStartRow, 2) = strSoftwareVendor
					objWorkSheet.Cells(intStartRow, 3) = strSoftwareVersion
					objWorkSheet.Cells(intStartRow, 4) = dtmSoftwareDate
					objWorkSheet.Cells(intStartRow, 5) = strUninstallString
					objWorkSheet.Cells(intStartRow, 6) = strSilentString
					intStartRow = intStartRow + 1
				Case 3
					objFile.WriteLine "	<tr>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strSoftwareName
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "		 	<a target=_blank href=""http://www.google.com/search?q=" & _
						strSoftwareName & """>Search</a>" 
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strSoftwareVendor
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strSoftwareVersion
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & dtmSoftwareDate
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strUninstallString
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strSilentString
					objFile.WriteLine "		</td>"
					objFile.WriteLine "	</tr>"
				Case 4
					strTxt = strTxt & strSoftwareName & _
					String(intColumnIndex - Len(strSoftwareName), " ") & strSoftwareVendor & _
					String(intColumnIndex2 - Len(strSoftwareVendor), " ") & strSoftwareVersion & _
					String(intColumnIndex3 - Len(strSoftwareVersion), " ") & dtmSoftwareDate & vbCrLf
			End Select
		Next		

		Select Case SWExport.Value
			Case 1
				objFile.WriteLine strCSV
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\SoftwareDetails" & strPC & ".csv"
			Case 2
				Set objRange = objWorkSheet.Range("A1:Z5")
				Set objRange2 = objWorkSheet.Range("A5:F" & intStartRow - 1)
				Set objRange3 = objWorkSheet.Range("E:F")
				Set objRangeH = objWorkSheet.Range("A5:F5")
				
				objRange.Font.Bold = True
				objRange2.Borders.LineStyle = xlContinuous
				objRange2.Borders.Weight = xlThin
				objRange2.Borders.ColorIndex = xlAutomatic
				objRange3.ColumnWidth = 75
				objRange3.WrapText = True
				objRangeH.AutoFilter
				
				objWorksheet.Range("A6").Select
				objExcel.ActiveWindow.FreezePanes = "True"
				objWorksheet.Range("A1").Select
				
				objWorkSheet.Columns("A:ZZ").EntireColumn.AutoFit
				objExcel.DisplayAlerts = False
				objExcel.ActiveWorkbook.SaveAs(strTemp & "\SoftwareDetails" & strPC & ".xls")
				objExcel.Visible = True
				Set objExcel = Nothing
			Case 3
				objFile.WriteLine "</table>"
				objFile.WriteLine "<p class=""backtotop""><a href=""" & strHTMLTempDir & "/SKB/SoftwareDetails" & _
				strPC & ".htm#top"">[..back to top..]</a></p>"
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\SoftwareDetails" & strPC & ".htm"
			Case 4
				objFile.WriteLine strTxt
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\SoftwareDetails" & strPC & ".txt"
			End Select
		
		SWExport.Value = 0

		SoftwareTab.InnerHTML = DataArea.InnerHTML
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	UninstallSoftware(booSilent)
    '#	PURPOSE........:	Remotely uninstalls software
    '#	ARGUMENTS......:	booSilent = boolean value to determine whether it
	'#						is a regular uninstall (False) or a silent uninstall
	'#						(True)
    '#	EXAMPLE........:	UninstallSoftware(True)
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub UninstallSoftware(booSilent)
		On Error Resume Next
		SoftwareTab.InnerHTML = Nothing
		arrSoftwareDetails = Split(SoftwareItems.Value, "|")
		strName = arrSoftwareDetails(0)
		strUninstallString = arrSoftwareDetails(4)
		strSilentString = arrSoftwareDetails(5)

		strPath = objShell.ExpandEnvironmentStrings("%path%")
		arrPaths = Split(strPath, ";")
		For i = 0 To UBound(arrPaths)
			strPathFolder = arrPaths(i) & "\"
			strPathFolder = Replace(strPathFolder, "\\", "\")
			strPathFolder = Replace(LCase(strPathFolder), "%systemroot%", _
			objShell.ExpandEnvironmentStrings("%systemroot%"))
			If objFSO.FileExists(strPathFolder & "psexec.exe") Then strPSExecInPath = 1
			If objFSO.FileExists(strPathFolder & "rctrlx.exe") Then strRctrlxInPath = 1
		Next

		If strPSExecInPath = 0 AND strRctrlxInPath = 0 Then
			MsgBox "Neither Rctrlx nor PSExec can be found in System Path.", vbExclamation, "PC Management Utility"
			Exit Sub
		End If
		
		booUseInstallMonitor = objShell.RegRead(strRegStart & "\Software\SKB\PCManagementUtil\Tabs\booUseInstallMonitor")
	
		strUninstallString = Replace(strUninstallString, Chr(34), "")
		strSilentString = Replace(strSilentString, Chr(34), "")
		strUninstallString = Replace(LCase(strUninstallString), "/i", "/X")
		strSilentString = Replace(LCase(strSilentString), "/i", "/X")
		
		Err.Clear
		
		If booSilent = True Then
			If strRctrlxInPath = 1 Then
				objShell.Run "%COMSPEC% /c rctrlx " & strPC & " /i /app " & _
				strSilentString, 0
				Else
					objShell.Run "%COMSPEC% /c psexec -i \\" & strPC & " " & _
					strSilentString, 0
			End If
			Else
				If strRctrlxInPath = 1 Then
					objShell.Run "%COMSPEC% /c rctrlx " & strPC & " /i /app " & _
					strUninstallString, 0
					Else
						objShell.Run "%COMSPEC% /c psexec -i \\" & strPC & " " & _
						strUninstallString, 0
				End If
		End If
		
		MsgBox "Starting uninstall of " & strName & " on " & strPC, vbInformation, "PC Management Utility"
		
		If booUseInstallMonitor = "1" Then
			PauseScript(3000)
			x = 0
			Dim arrProcess()
			arrUninstallString = Split(strUninstallString, "\")

			For i = 0 To UBound(arrUninstallString)
				If InStr(arrUninstallString(i),".exe") > 0 Then
					arrExeString = Split(arrUninstallString(i), " ")
						For j = 0 To UBound(arrExeString)
							If InStr(arrExeString(j),".exe") > 0 Then
								ReDim Preserve arrProcess(x)
								arrProcess(x) = arrExeString(j)
								arrProcess(x) = Replace(arrProcess(x), Chr(34), "")
								x = x + 1
							End If
						Next
				End If
			Next
			If x > 1 Then 
				MsgBox "More than one executable file in uninstall command, unable to monitor uninstall.",_
				vbExclamation, "PC Management Utility"
				Else
					CheckUninstallProcess(strName)
					objShell.Run Chr(34) & strTemp & "\SKB\CheckProcess" & strName & ".vbs" & Chr(34) & _
					" " & arrProcess(0) & " " & strPC, 1
			End If
		End If
		
		SoftwareTab.InnerHTML = DataArea.InnerHTML
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	RefreshSoftware()
    '#	PURPOSE........:	Refreshes software in Software tab
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub RefreshSoftware()
		ShowSoftwareInfo False
		ShowSoftwareInfo True
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	CheckUninstallProcess(strName)
    '#	PURPOSE........:	Checks to see if uninstaller is running on remote
	'#						PC and reports when it is finished
    '#	ARGUMENTS......:	strName = Software Name to append to CheckProcess
	'#						in vbs filename in case 2 uninstallers are running
    '#	EXAMPLE........:	CheckUninstallProcess(Microsoft Office)
    '#	NOTES..........:	Only creates vbs file to accomplish purpose in
	'#						%temp% directory. Will not work if uninstall
	'#						process name is diffrent to exe file in uninstall
	'#						string.
    '#--------------------------------------------------------------------------
	Sub CheckUninstallProcess(strName)
		Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\CheckProcess" & strName & ".vbs",True)
		objFile.WriteLine "Set objShell = CreateObject(""WScript.Shell"")"
		objFile.WriteLine "strUser = objShell.ExpandEnvironmentStrings(""%USERNAME%"")"
		objFile.WriteLine ""
		objFile.WriteLine "If (Wscript.Arguments.Count < 2) Then"
		objFile.WriteLine "	MsgBox ""Required parameter missing"", vbExclamation, ""PC Management Utility"""
		objFile.WriteLine "	Wscript.Quit"
		objFile.WriteLine "End If"
		objFile.WriteLine ""
		objFile.WriteLine "strUninstallString = Wscript.Arguments(0)"
		objFile.WriteLine "strPC = Wscript.Arguments(1)"
		objFile.WriteLine "strUninstallName = " & Chr(34) & strName & Chr(34)
		objFile.WriteLine ""
		objFile.WriteLine "Set objWMIService = GetObject(""winmgmts:{impersonationLevel=impersonate}!\\"" & strPC & ""\root\cimv2"")"
		objFile.WriteLine ""
		objFile.WriteLine "intProcStartCount = CountProcesses"
		objFile.WriteLine "intProcCurrentCount = CountProcesses"
		objFile.WriteLine ""
		objFile.WriteLine "If intProcStartCount = 0 Then "
		objFile.WriteLine "	WScript.Sleep 3000"
		objFile.WriteLine "	intProcStartCount = CountProcesses"
		objFile.WriteLine "	intProcCurrentCount = CountProcesses"
		objFile.WriteLine "	If intProcStartCount = 0 Then"
		objFile.WriteLine "		MsgBox ""The install has not started correctly or the uninstaller is not running "" & _"
		objFile.WriteLine "		""under the '"" & UCase(strUninstallString) & ""' process and can therefore not be monitored."", _"
		objFile.WriteLine "		vbExclamation, ""PC Management Utility"""
		objFile.WriteLine "		WScript.Quit"
		objFile.WriteLine "	End If"
		objFile.WriteLine "End If"
		objFile.WriteLine ""
		objFile.WriteLine "Do While intProcStartCount = intProcCurrentCount"
		objFile.WriteLine "	WScript.Sleep 5000"
		objFile.WriteLine "	intProcCurrentCount = CountProcesses"
		objFile.WriteLine "Loop"
		objFile.WriteLine ""
		objFile.WriteLine "If intProcCurrentCount < intProcStartCount Then MsgBox ""The uninstallation of "" & strUninstallName & _"
		objFile.WriteLine """ on "" & strPC & "" has now completed."", vbInformation, ""PC Management Utility"""
		objFile.WriteLine ""
		objFile.WriteLine "Function CountProcesses"
		objFile.WriteLine "	x = 0"
		objFile.WriteLine "	Set colProcess = objWMIService.ExecQuery _"
		objFile.WriteLine "		(""Select * from Win32_Process Where Name='"" & strUninstallString & ""'"")"
		objFile.WriteLine "	For Each objItem in colProcess"
		objFile.WriteLine "		colProperties = objItem.GetOwner(strNameOfUser,strUserDomain)"
		objFile.WriteLine "		If LCase(strNameOfUser) = LCase(strUser) Then x = x + 1"
		objFile.WriteLine "	Next"
		objFile.WriteLine "	CountProcesses = x"
		objFile.WriteLine "End Function"
		objFile.Close
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ShowProcessInfo(booChoice)
    '#	PURPOSE........:	Displays the Processes tab
    '#	ARGUMENTS......:	booChoice = boolean value to determine whether it
	'#						was called via Tab (True) or initial query (False)
    '#	EXAMPLE........:	ShowProcessInfo(True)
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ShowProcessInfo(booChoice)
		If booChoice = False Then
			On Error Resume Next
			
			booSortProcess = objShell.RegRead(strRegStart & "\Software\SKB\PCManagementUtil\Tabs\booSortProcess")
			If booSortProcess = "" OR IsNull(booSortProcess) Then 
				objShell.RegWrite strRegStart & "\Software\SKB\PCManagementUtil\Tabs\booSortProcess", _
				"1", "REG_SZ"
				booSortProcess = "1"
			End If
			
			Set DataList = CreateObject("ADOR.Recordset")
			DataList.Fields.Append "Process", adVarChar, MaxCharacters
			DataList.Fields.Append "ProcessID", adVarChar, MaxCharacters
			DataList.Fields.Append "MemUsage", adVarChar, MaxCharacters
			DataList.Fields.Append "ProcessUser", adVarChar, MaxCharacters
			DataList.Open

			
			DataArea.InnerHTML = "<h3><i>Fetching Process info for " & strPC & ", please wait.</i></h3>"
			PauseScript(1)
			
			Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
			strPC & "\root\cimv2") 

			Set colProcesses = objWMIService.ExecQuery _
				("Select * From Win32_Process")
				
			strHTML = "<b><u>" & strPC & "</u><p>"	
			strHTML = strHTML & "<div style=""overflow:auto;width:100%;height:458;" & _
            "border:1px solid #black;border-bottom-color:#a5a5a5;border-right:0px;padding:0px;margin:0px"">" 
			strHTML = strHTML & "<table class=""processtable"" id=""ProcessTable"" border=""1"">" 
			strHTML = strHTML & "<form Name = ""ProcessForm"" Method = ""post"">" 
			strHTML = strHTML & "<tr>" 
			strHTML = strHTML & "	<th style=""width:40%;"" title=""Process - The name of the process"">Process</th>" 
			strHTML = strHTML & "	<th style=""width:13%;"" title=""Process ID - A numerical identifier that " & _
			"uniquely distinguishes a process while it runs"">Process ID</th>"
			strHTML = strHTML & "	<th style=""width:16%;"" title=""User Name - The name of the user whose Terminal " & _
			"Services session owns the process"">User Name</th>"
			strHTML = strHTML & "	<th style=""width:13%;"" title=""Mem Usage - The current working set of a process, " & _
			"in kilobytes. The current working set is the number of pages currently resident in memory"">Mem Usage</th>" 
			strHTML = strHTML & "	<th style=""width:10%;"" title=""Google - Search Google for the process"">Google</th>" 
			strHTML = strHTML & "	<th style=""width:8%;"" title=""Kill - Tickbox to select individual " & _
			"processes to 'kill'"">Kill</th>" 
			strHTML = strHTML & "</tr>" 
			
			DataArea.InnerHTML = "<h3><i>Fetching Process info for " & strPC & ", please wait..</i></h3>"
			PauseScript(1)
			
			For Each objItem in colProcesses
				strProcessName = objItem.Caption
				intProcessID = objItem.ProcessID
				intMemUsage = objItem.WorkingSetSize
				If IsNull(intMemUsage) OR intMemUsage = "" Then intMemUsage = 0

				colProperties = objItem.GetOwner _
					(strProcessUser,strProcessUserDomain)
				
				DataList.AddNew
				DataList("Process") = strProcessName
				DataList("ProcessID") = intProcessID
				DataList("MemUsage") = intMemUsage
				DataList("ProcessUser") = strProcessUser
				DataList.Update 
			Next
			
			If booSortProcess = "1" Then DataList.Sort = "Process"
			
			DataList.MoveFirst
			Do Until DataList.EOF
				strProcessName = DataList.Fields.Item("Process")
				intProcessID = DataList.Fields.Item("ProcessID")
				intMemUsage = DataList.Fields.Item("MemUsage")
				strProcessUser = DataList.Fields.Item("ProcessUser")
				
				If intMemUsage <> 0 Then
					strMemUsage = Round(intMemUsage / 1024,2)
					strMemUsage = FormatNumber(strMemUsage,0) & " KB"
					Else
						strMemUsage = "0 KB"
				End If
				
				
				strHTML = strHTML & "<tr>"
				strHTML = strHTML & "	<td style=""width:40%;word-break:break-all;"" title=""" & strProcessName & _
				""">" & strProcessName & "</td>" 
				strHTML = strHTML & "	<td style=""width:13%;"">" & intProcessID & "</td>" 
				strHTML = strHTML & "	<td style=""width:16%;"" title=""" & strProcessUser & """>" & strProcessUser & "</td>" 
				strHTML = strHTML & "	<td style=""width:13%;"">" & strMemUsage & "</td>" 
				strHTML = strHTML & "	<td style=""width:10%;text-align:center;"">" & _
				"<a target=_blank href=""http://www.google.com/search?q=" & _
				strProcessName & """ title=""Search Google"">Search</a></td>" 
				strHTML = strHTML & "	<td style=""width:8%;text-align:center;"">" 
				strHTML = strHTML & "		<input type=""checkbox"" name=""cbxProcKill"" Value="""  & strProcessName & _
				"|" & intProcessID & "|" & intMemUsage & "|" & strProcessUser & """" & _
				"title=""" & strProcessName & """>"
				strHTML = strHTML & "	</td>"
				strHTML = strHTML & "</tr>"
				DataList.MoveNext
			Loop
				
			DataArea.InnerHTML = "<h3><i>Fetching Process info for " & strPC & ", please wait...</i></h3>"
			
			strHTML = strHTML & "</form>" 
			strHTML = strHTML & "</table>" 
			strHTML = strHTML & "</div>" 
			strHTML = strHTML & "<span style=""float:left;"">"
			strHTML = strHTML & "<input id=""RefreshButton"" class=""button"" type=""button"" value=""Refresh Processes"" " & _
			"name=""RefreshButton"" onclick=RefreshProcesses()>"
			strHTML = strHTML & "<input id=""KillButton"" class=""button"" type=""button"" value=""Kill Checked Process(es)"" " & _
			"name=""KillButton"" onclick=KillProcess()>"
			strHTML = strHTML & "</span>"
			strHTML = strHTML & "<span style=""float:right;"">"
			strHTML = strHTML & "<select name=""ProcessExport"" title=""Export the process list"" onChange=""ExportProcessDetails()"">"
			strHTML = strHTML & "	<option value=""0"">Export to:</option>"
			strHTML = strHTML & "	<option value=""1"" title=""Export the process list to a Comma " & _
			"Seperated Values (csv) file"")>Export to csv</option>"
			strHTML = strHTML & "	<option value=""2"" title=""Export the process list to an Excel " & _
			"(xls) file"">Export to xls</option>"
			strHTML = strHTML & "	<option value=""3"" title=""Export the process list to a Web " & _
			"page (html) file"">Export to html</option>"
			strHTML = strHTML & "	<option value=""4"" title=""Export the process list to a Text " & _
			"(txt) file"">Export to txt</option>"
			strHTML = strHTML & "</select></span>"	

			ProcessesTab.InnerHTML = strHTML
			
			DataArea.InnerHTML = "<h3><i>Fetching Process info for " & strPC & ", please wait....</i></h3>"
			PauseScript(1)
			
			Else
				If IsNull(ProcessesTab.InnerHTML) OR ProcessesTab.InnerHTML = "" Then
					ShowProcessInfo False
				End If
				tab1.bgcolor="#eeeeee"
				tab2.bgcolor="#eeeeee"
				tab3.bgcolor="#cccccc"
				tab4.bgcolor="#eeeeee"
				tab5.bgcolor="#eeeeee"
				tab6.bgcolor="#eeeeee"
				DataArea.InnerHTML = ProcessesTab.InnerHTML
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	KillProcess()
    '#	PURPOSE........:	Kills selected process(es) in Processes tab
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub KillProcess()
		On Error Resume Next
		ProcessesTab.InnerHTML = Nothing
		strMsg3 = ""
		strInput = Document.ProcessForm.cbxProcKill
		For Each strInput in ProcessForm
			If strInput.Checked = True Then
				arrInput = Split(strInput.Value, "|")
				strProcessName = arrInput(0)
				strMsg = strMsg & vbCrLf & strProcessName
			End If
		Next
		KillProcPrompt = MsgBox("Are you sure you wish to kill the following process(es) on " & strPC & _
		": " & vbCrLf & strMsg, vbQuestion+vbYesNo, "PC Management Utility")
		If KillProcPrompt = vbYes Then
			For Each strInput in ProcessForm
				If strInput.Checked = True Then
					arrInput = Split(strInput.Value, "|")
					strProcessName = arrInput(0)
					
					Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
					strPC & "\root\cimv2") 
					
					Set colProcesses = objWMIService.ExecQuery _
						("Select * from Win32_Process Where Name = " & Chr(34) & strProcessName & Chr(34))
					intProcStartCount = colProcesses.Count

					For Each objItem In colProcesses
						objItem.Terminate()
						PauseScript(100)
					Next
					Err.Clear
					Set colProcesses = objWMIService.ExecQuery _
						("Select * from Win32_Process Where Name = " & Chr(34) & strProcessName & Chr(34))
					intProcCurrentCount = colProcesses.Count
							
					If intProcCurrentCount = intProcStartCount Then
						objShell.Run "pskill.exe \\" & strPC & " " & strProcessName, 0, True
						PauseScript(100)
					End If
					
					Set colProcesses = objWMIService.ExecQuery _
						("Select * from Win32_Process Where Name = " & Chr(34) & strProcessName & Chr(34))
					intProcCurrentCount = colProcesses.Count
					
					If intProcCurrentCount = intProcStartCount Then
						strMsg3 = strMsg3 & vbCrLf & strProcessName
						Else
							strMsg2 = strMsg2 & vbCrLf & strProcessName
					End If
				End If
			Next
			If strMsg3 <> "" Then
				MsgBox "You killed the following process(es) on " & strPC & ": " & vbCrLf & _
				strMsg2 & vbCrLf & vbCrLf & "The following process(es) could not be killed: " & _
				vbCrLf & strMsg3, vbInformation, "PC Management Utility"
				Else
					MsgBox "You killed the following process(es) on " & strPC & ": " & _
					vbCrLf & strMsg2, vbInformation, "PC Management Utility"
					RefreshProcesses()
			End If
			Else
				For Each strInput in ProcessForm
					strInput.Checked = False
				Next
				ProcessesTab.InnerHTML = DataArea.InnerHTML
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	RefreshProcesses()
    '#	PURPOSE........:	Refreshes processes in Processes tab
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub RefreshProcesses()
		ShowProcessInfo False
		ShowProcessInfo True
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ExportProcessDetails()
    '#	PURPOSE........:	Export the details for the Processes
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ExportProcessDetails()
		On Error Resume Next
		ProcessesTab.InnerHTML = Nothing
		intProcesses = 0
		
		strInput = Document.ProcessForm.cbxProcKill
		
		For Each strInput in ProcessForm
			intProcesses = intProcesses + 1
		Next
		
		Select Case ProcessExport.Value
			Case 1
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\ProcessDetails" & strPC & ".csv",True)
				objFile.WriteLine "Processes on " & strPC
				objFile.WriteLine ""
				objFile.WriteLine "Running Processes: " & intProcesses
				objFile.WriteLine ""
				objFile.WriteLine "Process,Process ID,User Name,Mem Usage (KB)"
			Case 2
				Const xlContinuous = 1
				Const xlThin = 2
				Const xlAutomatic = -4105
				Const xlCenter = -4108
				
				strExcelPath = objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe\")
			   
				If strExcelPath = "" Then
					MsgBox "Unable to export. Excel does not appear to be installed.", vbExclamation, "PC Management Utility"
					Exit Sub
				End If
				
				Set objExcel = CreateObject("Excel.Application")
				objExcel.Visible = False
				Set objWorkBook = objExcel.WorkBooks.Add
				Set objWorksheet = objWorkbook.Worksheets(1)
				objExcel.DisplayAlerts = False
				For i = 1 to 3
					objWorkbook.Worksheets(2).Delete
				Next
				objExcel.DisplayAlerts = True
				objWorksheet.Name = "Process Details"
				
				objWorkSheet.Cells(1, 1) = "Processes on " & strPC
				objWorkSheet.Cells(3, 1) = "Running Processes: " & intProcesses

				intStartRow = 6
				
				objWorkSheet.Cells(5, 1) = "Process"
				objWorkSheet.Cells(5, 2) = "Process ID"
				objWorkSheet.Cells(5, 3) = "User Name"
				objWorkSheet.Cells(5, 4) = "Mem Usage (KB)"
			Case 3
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\ProcessDetails" & strPC & ".htm",True)
				objFile.WriteLine "<style type=""text/css"">"
				objFile.WriteLine "body{background-color:#CEF0FF;}"
				objFile.WriteLine "table.export{border-width:1px;border-spacing:1px;border-style:solid;border-color:gray;border-collapse:collapse;}"
				objFile.WriteLine "table.export th{border-width:1px;padding:1px;border-style:solid;border-color:gray;padding:2px 7px 2px 7px;}"
				objFile.WriteLine "table.export td{border-width:1px;padding:1px;border-style:dotted;border-color:gray;padding:2px 7px 2px 7px;}"
				objFile.WriteLine ".backtotop a {font-size:0.9em;}"
				objFile.WriteLine "</style>"
				objFile.WriteLine "<div style=""font-weight:bold;""><a name =""top"">Processes on " & strPC & "</a><p>"
				objFile.WriteLine "Running Processes: " & intProcesses & "<p></div>"
				objFile.WriteLine "<table class=""export"">"
				objFile.WriteLine "	<tr>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Process"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Process ID"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			User Name"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Mem Usage"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Process Library"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Google"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "	</tr>"
			Case 4
				intColumnIndex = 12
				intColumnIndex2 = 15
				intColumnIndex3 = 14
				For Each strInput in ProcessForm
					arrInput = Split(strInput.Value, "|")
					strProcessName = arrInput(0)
					If strProcessName = "&nbsp;" Then strProcessName = ""
					strProcessID = arrInput(1)
					If strProcessID = "&nbsp;" Then strProcessID = ""
					strProcessUser = arrInput(3)
					
					If Len(strProcessName) > intColumnIndex - 5 Then intColumnIndex = Len(strProcessName) + 5
					If Len(strProcessID) > intColumnIndex2 - 5 Then intColumnIndex2 = Len(strProcessID) + 5
					If Len(strProcessUser) > intColumnIndex3 - 5 Then intColumnIndex3 = Len(strProcessUser) + 5
				Next

				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\ProcessDetails" & strPC & ".txt",True)
				objFile.WriteLine "Processes on " & strPC
				objFile.WriteLine ""
				objFile.WriteLine "Running Processes: " & intProcesses
				objFile.WriteLine ""
				objFile.WriteLine "Process" & _
				String(intColumnIndex - 7, " ") & "Process ID" & _
				String(intColumnIndex2 - 10, " ") & "User Name" & _
				String(intColumnIndex3 - 9, " ") & "Mem Usage"
		End Select
		
		For Each strInput in ProcessForm
			On Error Resume Next
			arrInput = Split(strInput.Value, "|")
			strProcessName = arrInput(0)
			strProcessID = arrInput(1)
			If strProcessID = "" Then strProcessID = "0"

			intMemUsage = arrInput(2)
			If intMemUsage <> 0 Then
				strMemUsage = Round(intMemUsage / 1024,2)
				intMemUsage = Round(intMemUsage / 1024,2)
				strMemUsage = FormatNumber(strMemUsage,0) & " KB"
				Else
					strMemUsage = "0 KB"
			End If
			
			strProcessUser = arrInput(3)
			
			Select Case ProcessExport.Value
				Case 1
					strProcessName = EncodeCsv(strProcessName)

					strCSV = strCSV & strProcessName & "," & _
					strProcessID & "," & strProcessUser & "," & _
					intMemUsage & vbCrLf
				Case 2
					objWorkSheet.Cells(intStartRow, 1) = strProcessName
					objWorkSheet.Cells(intStartRow, 2) = strProcessID
					objWorkSheet.Cells(intStartRow, 3) = strProcessUser
					objWorkSheet.Cells(intStartRow, 4) = intMemUsage
					intStartRow = intStartRow + 1
				Case 3
					objFile.WriteLine "	<tr>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strProcessName
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strProcessID
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strProcessUser
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strMemUsage
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td style=""text-align:center;"">"
					objFile.WriteLine "		 	<a target=_blank href=""http://www.processlibrary.com/directory/files/" & _
						strProcessName & """>Search</a>" 
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td style=""text-align:center;"">"
					objFile.WriteLine "		 	<a target=_blank href=""http://www.google.com/search?q=" & _
						strProcessName & """>Search</a>" 
					objFile.WriteLine "		</td>"					
					objFile.WriteLine "	</tr>"
				Case 4
					strTxt = strTxt & strProcessName & _
					String(intColumnIndex - Len(strProcessName), " ") & strProcessID & _
					String(intColumnIndex2 - Len(strProcessID), " ") & strProcessUser & _
					String(intColumnIndex3 - Len(strProcessUser), " ") & strMemUsage & vbCrLf
			End Select
			strProcessName = ""
			strProcessID = ""
			strMemUsage = ""
		Next
			
		Select Case ProcessExport.Value
			Case 1
				objFile.WriteLine strCSV
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\ProcessDetails" & strPC & ".csv"
			Case 2
				Set objRangeH = objWorkSheet.Range("A5:D5")
				Set objRange = objWorkSheet.Range("A1:Z5")
				Set objRange2 = objWorkSheet.Range("A5:D" & intStartRow - 1)
				Set objRange3 = objWorkSheet.Range("B:B")
				Set objRange4 = objWorkSheet.Range("D:D")
				
				objRange.Font.Bold = True
				objRange2.Borders.LineStyle = xlContinuous
				objRange2.Borders.Weight = xlThin
				objRange2.Borders.ColorIndex = xlAutomatic
				objRange3.HorizontalAlignment = xlCenter
				objRange4.HorizontalAlignment = xlCenter
				objRange4.ColumnWidth = 10.86
				objRange4.NumberFormat = "#,##0"
				objRangeH.WrapText = True
				objRangeH.EntireRow.Autofit
				objRangeH.AutoFilter
				
				objWorksheet.Range("A6").Select
				objExcel.ActiveWindow.FreezePanes = "True"
				objWorksheet.Range("A1").Select
				
				objWorkSheet.Columns("A:C").EntireColumn.AutoFit
				objWorkSheet.Columns("B:B").ColumnWidth = 14.29
				objExcel.DisplayAlerts = False
				objExcel.ActiveWorkbook.SaveAs(strTemp & "\SKB\ProcessDetails" & strPC & ".xls")
				objExcel.Visible = True
				Set objExcel = Nothing
			Case 3
				objFile.WriteLine "</table>"
				objFile.WriteLine "<p class=""backtotop""><a href=""" & strHTMLTempDir & "/SKB/ProcessDetails" & _
				strPC & ".htm#top"">[..back to top..]</a></p>"
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\ProcessDetails" & strPC & ".htm"
			Case 4
				objFile.WriteLine strTxt
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\ProcessDetails" & strPC & ".txt"
			End Select
		
		ProcessExport.Value = 0

		ProcessesTab.InnerHTML = DataArea.InnerHTML
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ShowServiceInfo(booChoice)
    '#	PURPOSE........:	Displays the Services tab
    '#	ARGUMENTS......:	booChoice = boolean value to determine whether it
	'#						was called via Tab (True) or initial query (False)
    '#	EXAMPLE........:	ShowServiceInfo(True)
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ShowServiceInfo(booChoice)
		If booChoice = False Then
			On Error Resume Next
			x = 0
			
			booSortServices = objShell.RegRead(strRegStart & "\Software\SKB\PCManagementUtil\Tabs\booSortServices")
			If booSortServices = "" OR IsNull(booSortServices) Then 
				objShell.RegWrite strRegStart & "\Software\SKB\PCManagementUtil\Tabs\booSortServices", _
				"1", "REG_SZ"
				booSortServices = "1"
			End If
			
			Set DataList = CreateObject("ADOR.Recordset")
			DataList.Fields.Append "ServiceFull", adVarChar, MaxCharacters
			DataList.Fields.Append "Service", adVarChar, MaxCharacters
			DataList.Fields.Append "Description", adVarChar, MaxCharacters
			DataList.Fields.Append "Status", adVarChar, MaxCharacters
			DataList.Fields.Append "StartupType", adVarChar, MaxCharacters
			DataList.Fields.Append "LogOnAs", adVarChar, MaxCharacters
			DataList.Open

			
			DataArea.InnerHTML = "<h3><i>Fetching Services info for " & strPC & ", please wait.</i></h3>"
			PauseScript(1)
			
			Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
			strPC & "\root\cimv2") 

			Set colService = objWMIService.ExecQuery _
				("Select * From Win32_Service")
				
			strHTML = "<b><u>" & strPC & "</u><p>"	
			strHTML = strHTML & "<div style=""overflow:auto;width:100%;height:458;" & _
            "border:1px solid #black;border-bottom-color:#a5a5a5;border-right:0px;padding:0px;margin:0px"">" 
			strHTML = strHTML & "<table class=""processtable"" id=""ProcessTable"" border=""1"">" 
			strHTML = strHTML & "<form Name = ""ServiceForm"" Method = ""post"">" 
			strHTML = strHTML & "<tr>" 
			strHTML = strHTML & "	<th style=""width:42%"">Service</th>" 
			strHTML = strHTML & "	<th style=""width:11%"">Status</th>" 
			strHTML = strHTML & "	<th style=""width:11%"">Startup Type</th>" 
			strHTML = strHTML & "	<th style=""width:19%"">Log On As</th>" 
			strHTML = strHTML & "	<th style=""width:17%"">Action</th>" 
			strHTML = strHTML & "</tr>"
			
			DataArea.InnerHTML = "<h3><i>Fetching Services info for " & strPC & ", please wait..</i></h3>"
			PauseScript(1)
			
			For Each objItem in colService
				strServiceName = objItem.Caption
				strService = objItem.Name
				strStatus = objItem.State
				strStartupType = objItem.StartMode
				strDescription = objItem.Description
				strLogOnAs = objItem.StartName
				
				If InStr(LCase(strLogOnAs), "networkservice") > 0 Then strLogOnAs = "Network Service"
				If InStr(LCase(strLogOnAs), "localsystem") > 0 Then strLogOnAs = "Local System"
				If InStr(LCase(strLogOnAs), "localservice") > 0 Then strLogOnAs = "Local Service"
				
				DataList.AddNew
				DataList("ServiceFull") = strServiceName
				DataList("Service") = strService
				DataList("Description") = strDescription
				DataList("Status") = strStatus
				DataList("StartupType") = strStartupType
				DataList("LogOnAs") = strLogOnAs
				DataList.Update 
			Next
			
			If booSortServices = "1" Then DataList.Sort = "ServiceFull"
			
			DataList.MoveFirst
			Do Until DataList.EOF
					strServiceName = DataList.Fields.Item("ServiceFull")
					strService = DataList.Fields.Item("Service")
					strDescription = DataList.Fields.Item("Description")
					strStatus = DataList.Fields.Item("Status")
					strStartupType = DataList.Fields.Item("StartupType")
					strLogOnAs = DataList.Fields.Item("LogOnAs")
			
				strHTML = strHTML & "<tr>"
				strHTML = strHTML & "	<td style=""width:42%;"" "
				
				If LCase(strStartupType) = "disabled" Then
					strHTML = strHTML & "style=color:gray; "
					Else
						Select Case LCase(strStatus)
							Case "running", "start pending", "continue pending", _
							"pause pending", "paused"
								strHTML = strHTML & "style=color:green; "
							Case "stopped", "stop pending", "unknown"
								strHTML = strHTML & "style=color:red; "
						End Select
				End If
				
				If strDescription <> "" Then
					strHTML = strHTML & "title="" " & strServiceName & " - " & _
					strDescription & " "" "
				End If
				strHTML = strHTML & ">" & strServiceName & "</td>"
				strHTML = strHTML & "	<td style=""width:11%;"" "

				If LCase(strStartupType) = "disabled" Then
					strHTML = strHTML & "style=color:gray; "
					Else
						Select Case LCase(strStatus)
							Case "running", "start pending", "continue pending", _
							"pause pending", "paused"
								strHTML = strHTML & "style=color:green; "
							Case "stopped", "stop pending", "unknown"
								strHTML = strHTML & "style=color:red; "
						End Select
				End If

				strHTML = strHTML & ">" & strStatus & "</td>" 
				strHTML = strHTML & "	<td style=""width:11%;"" "
				
				If LCase(strStartupType) = "disabled" Then
					strHTML = strHTML & "style=color:gray; "
					Else
						Select Case LCase(strStatus)
							Case "running", "start pending", "continue pending", _
							"pause pending", "paused"
								strHTML = strHTML & "style=color:green; "
							Case "stopped", "stop pending", "unknown"
								strHTML = strHTML & "style=color:red; "
						End Select
				End If
				
				strHTML = strHTML & ">" & strStartupType & "</td>" 
				strHTML = strHTML & "	<td style=""width:19%;word-break:break-all;"" "
				
				If LCase(strStartupType) = "disabled" Then
					strHTML = strHTML & "style=color:gray; "
					Else
						Select Case LCase(strStatus)
							Case "running", "start pending", "continue pending", _
							"pause pending", "paused"
								strHTML = strHTML & "style=color:green; "
							Case "stopped", "stop pending", "unknown"
								strHTML = strHTML & "style=color:red; "
						End Select
				End If
				
				strHTML = strHTML & ">" & strLogOnAs & "</td>" 
				
				strHTML = strHTML & "	<td style=""width:17%;"">" 
				strHTML = strHTML & "		<select style=""width:100%;"" name=""ChangeService"" " & _
				"title=""Change the service state"" onChange=""ChangeServiceState()"">"
				strHTML = strHTML & "				<option value=""0||" & strServiceName & _
				"||" & DataList("Description") & "||" & DataList("Status") & "||" & DataList("StartupType") & "||" & _
				DataList("LogOnAs") & """ title="""">Select Action</option>"
				
				If LCase(strStartupType) <> "disabled" Then	
					Select Case LCase(strStatus)
						Case "running", "paused"
							strHTML = strHTML & "				<option value=""1||" & strService & _
							""" title=""Stop the service"")>Stop Service</option>"
							strHTML = strHTML & "				<option value=""2||" & strService &  _
							""" title=""Restart the service"">Restart Service</option>"
						Case "stopped"
							strHTML = strHTML & "				<option value=""3||" & strService & _
							""" title=""Start the service"">Start Service</option>"
						Case "stop pending", "start pending", "continue pending", "pause pending", "unknown"
							strHTML = strHTML & "				<option value=""3||" & strService & _
							""" title=""Start the service"">Start Service</option>"
							strHTML = strHTML & "				<option value=""1||" & strService & _
							""" title=""Stop the service"")>Stop Service</option>"
							strHTML = strHTML & "				<option value=""2||" & strService &  _
							""" title=""Restart the service"">Restart Service</option>"
					End Select
					
					strHTML = strHTML & "					<option value=""4||" & strService & _
					""" title=""Disable the service"">Disable Service</option>"
					
					Else
						strHTML = strHTML & "				<option value=""5||" & strService & _
						""" title=""Enable the service"">Enable Service</option>"
				End If
				
				strHTML = strHTML & "			</select>"	
				
				strHTML = strHTML & "	</td>" 
				strHTML = strHTML & "</tr>" 
				DataList.MoveNext
			Loop
				
			DataArea.InnerHTML = "<h3><i>Fetching Services info for " & strPC & ", please wait...</i></h3>"
			
			strHTML = strHTML & "</form>" 
			strHTML = strHTML & "</table>" 
			strHTML = strHTML & "</div>" 
			strHTML = strHTML & "<span style=""float:left;"">"
			strHTML = strHTML & "<input id=""RefreshButton"" class=""button"" type=""button"" " & _
			"value=""Refresh Services"" name=""RefreshButton"" onclick=RefreshServices()>"
			strHTML = strHTML & "</span>"
			strHTML = strHTML & "<span style=""float:right;"">"
			strHTML = strHTML & "<select name=""ServicesExport"" title=""Export the services list"" " & _
			"onChange=""ExportServiceDetails()"">"
			strHTML = strHTML & "	<option value=""0"">Export to:</option>"
			strHTML = strHTML & "	<option value=""1"" title=""Export the services list to a Comma " & _
			"Seperated Values (csv) file"")>Export to csv</option>"
			strHTML = strHTML & "	<option value=""2"" title=""Export the services list to an Excel " & _
			"(xls) file"">Export to xls</option>"
			strHTML = strHTML & "	<option value=""3"" title=""Export the services list to a Web " & _
			"page (html) file"">Export to html</option>"
			strHTML = strHTML & "	<option value=""4"" title=""Export the services list to a Text " & _
			"(txt) file"">Export to txt</option>"
			strHTML = strHTML & "</select></span>"	

			ServicesTab.InnerHTML = strHTML
			
			DataArea.InnerHTML = "<h3><i>Fetching Services info for " & strPC & ", please wait....</i></h3>"
			PauseScript(1)
			
			Else
				If IsNull(ServicesTab.InnerHTML) OR ServicesTab.InnerHTML = "" Then
					ShowServiceInfo False
				End If
				tab1.bgcolor="#eeeeee"
				tab2.bgcolor="#eeeeee"
				tab3.bgcolor="#eeeeee"
				tab4.bgcolor="#cccccc"
				tab5.bgcolor="#eeeeee"
				tab6.bgcolor="#eeeeee"
				DataArea.InnerHTML = ServicesTab.InnerHTML
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ExportServiceDetails()
    '#	PURPOSE........:	Export the details for the Services
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ExportServiceDetails()
		On Error Resume Next
		
		ServicesTab.InnerHTML = Nothing
		intServices = 0
		
		strInput = Document.ServiceForm.ChangeService
		
		For Each strInput in ServiceForm
			intServices = intServices + 1
		Next
		
		Select Case ServicesExport.Value
			Case 1
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\ServicesInfo" & strPC & ".csv",True)
				objFile.WriteLine "Services on " & strPC
				objFile.WriteLine ""
				objFile.WriteLine "Total Services: " & intServices
				objFile.WriteLine ""
				objFile.WriteLine "Service,Description,Status,Startup Type,Log On As"
			Case 2
				Const xlContinuous = 1
				Const xlThin = 2
				Const xlAutomatic = -4105
				Const xlCenter = -4108
				
				strExcelPath = objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe\")
			   
				If strExcelPath = "" Then
					MsgBox "Unable to export. Excel does not appear to be installed.", vbExclamation, "PC Management Utility"
					Exit Sub
				End If
				
				Set objExcel = CreateObject("Excel.Application")
				objExcel.Visible = False
				Set objWorkBook = objExcel.WorkBooks.Add
				Set objWorksheet = objWorkbook.Worksheets(1)
				objExcel.DisplayAlerts = False
				For i = 1 to 3
					objWorkbook.Worksheets(2).Delete
				Next
				objExcel.DisplayAlerts = True
				objWorksheet.Name = "Service Details"
				
				objWorkSheet.Cells(1, 1) = "Services on " & strPC
				objWorkSheet.Cells(3, 1) = "Total Services: " & intServices

				intStartRow = 6
				
				objWorkSheet.Cells(5, 1) = "Service"
				objWorkSheet.Cells(5, 2) = "Description"
				objWorkSheet.Cells(5, 3) = "Status"
				objWorkSheet.Cells(5, 4) = "Startup Type"
				objWorkSheet.Cells(5, 5) = "Log On As"
			Case 3
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\ServicesInfo" & strPC & ".htm",True)
				objFile.WriteLine "<style type=""text/css"">"
				objFile.WriteLine "body{background-color:#CEF0FF;}"
				objFile.WriteLine "table.export{border-width:1px;border-spacing:1px;border-style:solid;border-color:gray;border-collapse:collapse;}"
				objFile.WriteLine "table.export th{border-width:1px;padding:1px;border-style:solid;border-color:gray;padding:2px 7px 2px 7px;}"
				objFile.WriteLine "table.export td{border-width:1px;padding:1px;border-style:dotted;border-color:gray;padding:2px 7px 2px 7px;}"
				objFile.WriteLine ".backtotop a {font-size:0.9em;}"
				objFile.WriteLine "</style>"
				objFile.WriteLine "<div style=""font-weight:bold;""><a name =""top"">Services on " & strPC & "</a><p>"
				objFile.WriteLine "Total Services: " & intServices & "<p></div>"
				objFile.WriteLine "<table class=""export"">"
				objFile.WriteLine "	<tr>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Service"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Description"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Google"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Status"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Startup Type"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Log On As"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "	</tr>"
			Case 4
				intColumnIndex = 12
				intColumnIndex2 = 11
				intColumnIndex3 = 17
				For Each strInput in ServiceForm
					arrInput = Split(strInput.Value, "||")
					strServiceName = arrInput(1)
					strStatus = arrInput(3)
					If Len(strServiceName) > intColumnIndex - 5 Then intColumnIndex = Len(strServiceName) + 5
					If Len(strStatus) > intColumnIndex2 - 5 Then intColumnIndex2 = Len(strStatus) + 5
				Next

				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\ServicesInfo" & strPC & ".txt",True)
				objFile.WriteLine "Services on " & strPC
				objFile.WriteLine ""
				objFile.WriteLine "Total Services: " & intServices
				objFile.WriteLine ""
				objFile.WriteLine "Service" & _
				String(intColumnIndex - 7, " ") & "Status" & _
				String(intColumnIndex2 - 6, " ") & "Startup Type" & _
				String(intColumnIndex3 - 12, " ") & "Log On As"
		End Select
		
		For Each strInput in ServiceForm
			arrInput = Split(strInput.Value, "||")
			strServiceName = arrInput(1)
			strDescription = arrInput(2)
			strDescription = Replace(strDescription, vbCrLf, " ")
			strStatus = arrInput(3)
			strStartupType = arrInput(4)
			strLogOnAs = arrInput(5)
			
			Select Case ServicesExport.Value
				Case 1
					strDescription = Replace(strDescription, ",", "")
					strCSV = strCSV & strServiceName & "," & _
					strDescription & "," & strStatus & "," & _
					strStartupType & "," & strLogOnAs & vbCrLf
				Case 2
					objWorkSheet.Cells(intStartRow, 1) = strServiceName
					objWorkSheet.Cells(intStartRow, 2) = strDescription
					objWorkSheet.Cells(intStartRow, 3) = strStatus
					objWorkSheet.Cells(intStartRow, 4) = strStartupType
					objWorkSheet.Cells(intStartRow, 5) = strLogOnAs
					intStartRow = intStartRow + 1
				Case 3
					objFile.WriteLine "	<tr>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strServiceName
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strDescription
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			<a target=_blank href=""http://www.google.com/search?q=" & _
					strServiceName & " Service"">Search</a>"
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strStatus
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "		 	" & strStartupType
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "		 	" & strLogOnAs
					objFile.WriteLine "		</td>"	
					objFile.WriteLine "	</tr>"
				Case 4
					strTxt = strTxt & strServiceName & _
					String(intColumnIndex - Len(strServiceName), " ") & strStatus & _
					String(intColumnIndex2 - Len(strStatus), " ") & strStartupType & _
					String(intColumnIndex3 - Len(strStartupType), " ") & strLogOnAs & vbCrLf
			End Select
		Next
			
		Select Case ServicesExport.Value
			Case 1
				objFile.WriteLine strCSV
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\ServicesInfo" & strPC & ".csv"
			Case 2
				Set objRange = objWorkSheet.Range("A1:Z5")
				Set objRange2 = objWorkSheet.Range("A5:E" & intStartRow - 1)
				Set objRange3 = objWorkSheet.Range("B:B")
				Set objRangeH = objWorksheet.Range("A5:E5")
				
				objRange.Font.Bold = True
				objRange2.Borders.LineStyle = xlContinuous
				objRange2.Borders.Weight = xlThin
				objRange2.Borders.ColorIndex = xlAutomatic
				objRange3.ColumnWidth = 50
				objRange3.WrapText = True
				objRangeH.AutoFilter
				
				objWorksheet.Range("A6").Select
				objExcel.ActiveWindow.FreezePanes = "True"
				objWorksheet.Range("A1").Select
				
				objWorkSheet.Columns("A:A").EntireColumn.AutoFit
				objWorkSheet.Columns("C:Z").EntireColumn.AutoFit
				objExcel.DisplayAlerts = False
				objExcel.ActiveWorkbook.SaveAs(strTemp & "\SKB\ServicesInfo" & strPC & ".xls")
				objExcel.Visible = True
				Set objExcel = Nothing
			Case 3
				objFile.WriteLine "</table>"
				objFile.WriteLine "<p class=""backtotop""><a href=""" & strHTMLTempDir & "/SKB/ServicesInfo" & _
				strPC & ".htm#top"">[..back to top..]</a></p>"
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\ServicesInfo" & strPC & ".htm"
			Case 4
				objFile.WriteLine strTxt
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\ServicesInfo" & strPC & ".txt"
			End Select
		
		ServicesExport.Value = 0

		ServicesTab.InnerHTML = DataArea.InnerHTML
	End Sub

	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ChangeServiceState()
    '#	PURPOSE........:	Change the state of / enable or disable the service
	'#						depending on the current state
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ChangeServiceState()
		On Error Resume Next

		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\cimv2")
	
		ServicesTab.InnerHTML = Nothing
		
		strInput = Document.ServiceForm.ChangeService
		
		For Each strInput in ServiceForm
			intAction = Left(strInput.Value, 1)
			If intAction <> 0 Then
				arrInput = Split(strInput.Value, "||")
				intAction = arrInput(0)
				strService = arrInput(1)
				For Each objOption in strInput.Options
					If objOption.Title = "" Then strInput.Value = objOption.Value
				Next
				
				Set colService = objWMIService.ExecQuery _
					("Select * From Win32_Service Where Name='" & strService & "'")

				Select Case intAction
					Case 1
						For Each objItem in colService
							errReturn = objItem.StopService()
							If errReturn = 0 Then
								MsgBox "The '" & strService & "' Service has been " & _
								"Stopped successfully.", vbInformation, _
								"PC Management Utility"
								RefreshServices()
								Else
									MsgBox "There was an error Stopping the '" & _
									strService & "' Service", vbExclamation, _
									"PC Management Utility"
									ServicesTab.InnerHTML = DataArea.InnerHTML
							End If
						Next
					Case 2
						For Each objItem in colService
							errReturn = objItem.StopService()
							If errReturn = 0 Then
								PauseScript(2500)
								errReturn = objItem.StartService()
								If errReturn = 0 Then
									MsgBox "The '" & strService & "' Service has been " & _
									"Restarted successfully.", vbInformation, _
									"PC Management Utility"
									RefreshServices()
									Else
										MsgBox "There was an error Restarting the '" & _
										strService & "' Service", vbExclamation, _
										"PC Management Utility"
										ServicesTab.InnerHTML = DataArea.InnerHTML
								End If
								Else
									MsgBox "There was an error Restarting the '" & _
									strService & "' Service", vbExclamation, _
									"PC Management Utility"
									ServicesTab.InnerHTML = DataArea.InnerHTML
							End If
						Next
					Case 3
						For Each objItem in colService
							If LCase(objItem.Name) = LCase(strService) Then
								errReturn = objItem.StartService()
								If errReturn = 0 Then
									MsgBox "The '" & strService & "' Service has been " & _
									"Started successfully.", vbInformation, _
									"PC Management Utility"
									RefreshServices()
									Else
										MsgBox "There was an error Starting the '" & _
										strService & "' Service", vbExclamation, _
										"PC Management Utility"
										ServicesTab.InnerHTML = DataArea.InnerHTML
								End If
							End If
						Next
					Case 4
						For Each objItem in colService
							If LCase(objItem.Name) = LCase(strService) Then
								errReturn = objItem.ChangeStartMode("Disabled")
								If errReturn = 0 Then
									MsgBox "The '" & strService & "' Service has been " & _
									"Disabled successfully.", vbInformation, _
									"PC Management Utility"
									RefreshServices()
									Else
										MsgBox "There was an error Disabling the '" & _
										strService & "' Service", vbExclamation, _
										"PC Management Utility"
										ServicesTab.InnerHTML = DataArea.InnerHTML
								End If
							End If
						Next
					Case 5
						For Each objItem in colService
							If LCase(objItem.Name) = LCase(strService) Then
								errReturn = objItem.ChangeStartMode("Manual")
								If errReturn = 0 Then
									MsgBox "The '" & strService & "' Service has been " & _
									"Enabled successfully.", vbInformation, _
									"PC Management Utility"
									RefreshServices()
									Else
										MsgBox "There was an error Enabling the '" & _
										strService & "' Service", vbExclamation, _
										"PC Management Utility"
										ServicesTab.InnerHTML = DataArea.InnerHTML
								End If
							End If
						Next
				End Select
				Exit Sub
			End If
		Next
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	RefreshServices()
    '#	PURPOSE........:	Refreshes services in Services tab
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub RefreshServices()
		ShowServiceInfo False
		ShowServiceInfo True
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ShowStartupInfo(booChoice)
    '#	PURPOSE........:	Displays the Startup Items tab
    '#	ARGUMENTS......:	booChoice = boolean value to determine whether it
	'#						was called via Tab (True) or initial query (False)
    '#	EXAMPLE........:	ShowStartupInfo(True)
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ShowStartupInfo(booChoice)
		If booChoice = False Then
			On Error Resume Next
			
			booSortStartup = objShell.RegRead(strRegStart & "\Software\SKB\PCManagementUtil\Tabs\booSortStartup")
			If booSortStartup = "" OR IsNull(booSortStartup) Then 
				objShell.RegWrite strRegStart & "\Software\SKB\PCManagementUtil\Tabs\booSortStartup", _
				"1", "REG_SZ"
				booSortStartup = "1"
			End If
			
			Set DataList = CreateObject("ADOR.Recordset")
			DataList.Fields.Append "StartupItem", adVarChar, MaxCharacters
			DataList.Fields.Append "Command", adVarChar, MaxCharacters
			DataList.Fields.Append "User", adVarChar, MaxCharacters
			DataList.Fields.Append "StartupLocation", adVarChar, MaxCharacters
			DataList.Open

			
			DataArea.InnerHTML = "<h3><i>Fetching Startup info for " & strPC & ", please wait.</i></h3>"
			PauseScript(1)
			
			Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
			strPC & "\root\cimv2") 
			
			Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
			strPC & "\root\default:StdRegProv") 
	
			strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\S-1-5-18\"
			strKeyValue = "ProfileImagePath"
			objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strKeyValue, strValue
			arrStringValue = Split(strValue,":\")
			strSystemDrive =  arrStringValue(0)
			strNetworkServicePath = arrStringValue(1)
			
			strRoot = "\\" & strPC & "\" & strSystemDrive & "$"
			
			Err.Clear
			strNetworkServiceFullPath = strSystemDrive & ":\" & strNetworkServicePath & "\Start Menu\Programs\Startup"
			
			Set colStartup	= objWMIService.ExecQuery _
				("Select * From Win32_StartupCommand")
				
			strHTML = "<b><u>" & strPC & "</u><p>"
			strHTML = strHTML & "<div style=""overflow:auto;width:100%;height:458;" & _
            "border:1px solid #black;border-bottom-color:#a5a5a5;border-right:0px;padding:0px;margin:0px"">" 
			strHTML = strHTML & "<table class=""processtable"" id=""ProcessTable"" border=""1"">" 
			strHTML = strHTML & "<form Name = ""StartupForm"" Method = ""post"">" 
			strHTML = strHTML & "<tr>" 
			strHTML = strHTML & "<th style=""width:23%"">Startup Item</th>" 
			strHTML = strHTML & "<th style=""width:20%"">User</th>" 
			strHTML = strHTML & "<th style=""width:49%"">Startup Location</th>" 
			strHTML = strHTML & "<th style=""width:8%"">Delete</th>" 
			strHTML = strHTML & "</tr>" 
			
			DataArea.InnerHTML = "<h3><i>Fetching Startup info for " & strPC & ", please wait..</i></h3>"
			PauseScript(1)
			
			For Each objItem in colStartup
				strStartupName = Trim(objItem.Caption)
				strStartupUser = Trim(objItem.User)
				strStartupCommand = Trim(objItem.Command)
				strStartupLocation = Trim(objItem.Location)
				If UCase(strStartupName) <> UCase("CTFMON.EXE") AND UCase(strStartupName) <> UCase("DESKTOP") _
				AND UCase(strStartupName) <> "" Then
					If UCase(strStartupUser) = ".DEFAULT" Then strStartupFileUser = "Default User"
					If UCase(strStartupUser) = "NT AUTHORITY\SYSTEM" Then strStartupLocation = strNetworkServiceFullPath
					If UCase(strStartupUser) = UCase(strRemoteLoggedOn) _
					AND InStr(UCase(strStartupLocation), "HKU\" & UCase(strRemoteSID)) > 0 Then
						arrStartupLocation = Split(strStartupLocation, "\")
							strStartupLocation = "HKCU"
						For i = 2 to UBound(arrStartupLocation)
							strStartupLocation = strStartupLocation & "\" & arrStartupLocation(i)
						Next
					End If
					
					Select Case strStartupLocation
						Case "Common Startup"
							If objFSO.FolderExists(strRoot & "\Documents and Settings\All Users") Then
								strStartupLocation = _
								strSystemDrive & ":\Documents and Settings\All Users\Start Menu\Programs\Startup"
								ElseIf objFSO.FolderExists(strRoot & "\ProgramData\Start Menu\Programs\Startup") Then
									strStartupLocation = strSystemDrive & ":\ProgramData\Start Menu\Programs\Startup"
									ElseIf objFSO.FolderExists(strRoot & _
									"\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup") Then
										strStartupLocation = strSystemDrive & ":\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
							End If
						Case "Startup"
							arrStartupUser = Split(strStartupUser, "\")
							strStartupFileUser = UCase(arrStartupUser(UBound(arrStartupUser)))
							If objFSO.FolderExists(strRoot & "\Documents and Settings\" & strStartupFileUser) Then
								strStartupLocation = _
								strSystemDrive & ":\Documents and Settings\" & strStartupFileUser & "\Start Menu\Programs\Startup"
								ElseIf objFSO.FolderExists(strRoot & "\Users\" & strStartupFileUser & _
								"\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup") Then
									strStartupLocation = strSystemDrive & ":\Users\" & strStartupFileUser & _
									"\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
							End If
					End Select	
					
					If strStartupUser = "Default User" Then strStartupUser = ".DEFAULT"
					
					DataList.AddNew
					DataList("StartupItem") = strStartupName
					DataList("User") = strStartupUser
					DataList("StartupLocation") = strStartupLocation
					DataList("Command") = strStartupCommand
					DataList.Update 
				End If
			Next
			
			If booSortStartup = "1" Then DataList.Sort = "StartupItem"
			
			DataList.MoveFirst
			Do Until DataList.EOF
				strStartupName = DataList.Fields.Item("StartupItem")
				strStartupUser = DataList.Fields.Item("User")
				strStartupLocation = DataList.Fields.Item("StartupLocation")
				strStartupCommand = DataList.Fields.Item("Command")
			
				strHTML = strHTML & "<tr>"
				strHTML = strHTML & "<td style=""width:23%"">" & strStartupName & "</td>"
				strHTML = strHTML & "<td style=""width:20%;word-break:break-all;"">" & strStartupUser & "</td>"
				strHTML = strHTML & "<td style=""width:49%;word-break:break-all;"">" & strStartupLocation & "</td>"
				strHTML = strHTML & "<td style=""width:8%"" align=""center"">" 
				strHTML = strHTML & "<input type=""checkbox"" name=""cbxStartupDelete"" Value="" " & _
				strStartupName & "|" & strStartupUser & "|" & _
				strStartupLocation & "|" & strStartupCommand & """" & _
				"title=""" & strStartupName & """>"
				strHTML = strHTML & "</td>" 
				strHTML = strHTML & "</tr>" 
				DataList.MoveNext
			Loop

			DataArea.InnerHTML = "<h3><i>Fetching Startup info for " & strPC & ", please wait...</i></h3>"
			
			strHTML = strHTML & "</form>" 
			strHTML = strHTML & "</table>" 
			strHTML = strHTML & "</div>" 
			strHTML = strHTML & "<span style=""float:left;"">"
			strHTML = strHTML & "<input id=""btnStartupRefresh"" class=""button"" type=""button"" " & _
			"style=""width:100px"" title=""Refresh Startup items""" & _
			"value=""Refresh"" name=""btnStartupRefresh"" onclick=RefreshStartupItems()>"
			strHTML = strHTML & "<input id=""btnStartupDelete"" class=""button"" type=""button"" " & _
			"style=""width:175px"" title=""Delete checked Startup item(s)""" & _
			"value=""Delete Checked Item(s)"" name=""btnStartupDelete"" onclick=DeleteTabStartupItems()>"
			strHTML = strHTML & "<input id=""btnStartupAdd"" class=""button"" type=""button"" " & _
			"style=""width:150px"" title=""Add new Startup item""" & _
			"value=""Add Startup Item"" name=""btnStartupAdd"" onclick=""ChooseStartupFile()"">"
			strHTML = strHTML & "</span>"
			strHTML = strHTML & "<span style=""float:right;"">"
			strHTML = strHTML & "<select name=""StartupExport"" title=""Export the startup list"" onChange=""ExportStartupDetails()"">"
			strHTML = strHTML & "	<option value=""0"">Export to:</option>"
			strHTML = strHTML & "	<option value=""1"" title=""Export the startup list to a Comma " & _
			"Seperated Values (csv) file"")>Export to csv</option>"
			strHTML = strHTML & "	<option value=""2"" title=""Export the startup list to an Excel " & _
			"(xls) file"">Export to xls</option>"
			strHTML = strHTML & "	<option value=""3"" title=""Export the startup list to a Web " & _
			"page (html) file"">Export to html</option>"
			strHTML = strHTML & "	<option value=""4"" title=""Export the startup list to a Text " & _
			"(txt) file"">Export to txt</option>"
			strHTML = strHTML & "</select></span>"

			StartupTab.InnerHTML = strHTML
			
			DataArea.InnerHTML = "<h3><i>Fetching Startup info for " & strPC & ", please wait....</i></h3>"
			PauseScript(1)
			
			Else
				If IsNull(StartupTab.InnerHTML) OR StartupTab.InnerHTML = "" Then
					ShowStartupInfo False
				End If
				tab1.bgcolor="#eeeeee"
				tab2.bgcolor="#eeeeee"
				tab3.bgcolor="#eeeeee"
				tab4.bgcolor="#eeeeee"
				tab5.bgcolor="#cccccc"
				tab6.bgcolor="#eeeeee"
				DataArea.InnerHTML = StartupTab.InnerHTML
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	DeleteTabStartupItems()
    '#	PURPOSE........:	Deletes selected Startup Items
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	Will delete either from Registry or file location
    '#--------------------------------------------------------------------------
	Sub DeleteTabStartupItems()
		On Error Resume Next
		StartupTab.InnerHTML = Nothing

		strInputDelete = Document.StartupForm.cbxStartupDelete
	
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\cimv2") 
		
		Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\default:StdRegProv") 
	
		strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\S-1-5-18\"
		strKeyValue = "ProfileImagePath"
		objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strKeyValue, strValue
		arrStringValue = Split(strValue,":\")
		strSystemDrive =  arrStringValue(0)
		strNetworkServicePath = arrStringValue(1)
		
		strRoot = "\\" & strPC & "\" & strSystemDrive & "$"
		For Each strInputDelete in StartupForm
			If strInputDelete.Checked = True Then
				arrInputDelete = Split(strInputDelete.Value, "|")
				strInputName = arrInputDelete(0)
				strDeleteMsg = strDeleteMsg & vbCrLf & strInputName
			End If
		Next
		If strDeleteMsg = "" Then Exit Sub
		StartupPrompt = MsgBox("Are you sure you wish to remove the following Startup Item(s) on " & _
		strPC & ": " & vbCrLf & strDeleteMsg, vbQuestion+vbYesNo, "PC Management Utility")
		If StartupPrompt = vbYes Then
			For Each strInputDelete in StartupForm
				If strInputDelete.Checked = True Then
					arrInputDelete = Split(strInputDelete.Value, "|")
					strInputName = Trim(arrInputDelete(0))
					strInputUser = Trim(arrInputDelete(1))
					strInputUserWMI = Replace(strInputUser,"\","\\")
					strNetworkServiceFullPath = strRoot & "\" & _
					strNetworkServicePath & "\Start Menu\Programs\Startup\"
					Set colStartup = objWMIService.ExecQuery _
						("Select * From Win32_StartupCommand Where Caption = '" & _
						strInputName & "' AND User = '" & strInputUserWMI & "'")
					For Each objItem in colStartup
						strStartupName = objItem.Caption
						strStartupCommand = objItem.Command
						strStartupUser = objItem.User
						strStartupLocation = objItem.Location
						booFile = True
						Select Case strStartupLocation
							Case "Common Startup"
								If objFSO.FolderExists(strRoot & "\Documents and Settings\All Users") Then
									strStartupLocation = strRoot & "\Documents and Settings\All Users\Start Menu\Programs\Startup"
									ElseIf objFSO.FolderExists(strRoot & _
									"\ProgramData\Start Menu\Programs\Startup") Then
										strStartupLocation = strRoot & "\ProgramData\Start Menu\Programs\Startup"
										ElseIf objFSO.FolderExists(strRoot & _
										"\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup") Then
											strStartupLocation = strRoot & "\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
								End If
							Case "Startup"
								arrStartupUser = Split(strStartupUser, "\")
								strStartupFileUser = UCase(arrStartupUser(UBound(arrStartupUser)))
								If UCase(strStartupFileUser) = ".DEFAULT" Then strStartupFileUser = "Default User"
								If UCase(strStartupFileUser) = "NT AUTHORITY\SYSTEM" Then 
									strStartupLocation = strNetworkServicePath
								End If
								If objFSO.FolderExists(strRoot & "\Documents and Settings\" & strStartupFileUser) Then
									strStartupLocation = strRoot & _
									"\Documents and Settings\" & strStartupFileUser & "\Start Menu\Programs\Startup"
									ElseIf objFSO.FolderExists(strRoot & "\Users\" & _
									strStartupFileUser & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup") Then
										strStartupLocation = strRoot & "\Users\" & strStartupFileUser & _
										"\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
								End If
						End Select
						If inStr(strStartupLocation,"HKU") > 0 Then
							strStartupLocation = Replace(strStartupLocation,"HKU\","")
							booFile = False
							intRegType = 1
						End If
						If inStr(strStartupLocation,"HKLM") > 0 Then
							strStartupLocation = Replace(strStartupLocation,"HKLM\","")
							booFile = False
							intRegType = 2
						End If
						strMsg2 = strMsg2 & vbCrLf & strInputName
					Next
					If booFile = True Then
						arrCheckFile = Split(strStartupCommand,"\")
						strCheckFile = arrCheckFile(UBound(arrCheckFile))
						If objFSO.FileExists(strStartupLocation & "\" & strStartupName) Then
							ElseIf objFSO.FileExists(strStartupLocation & "\" & strStartupName & ".lnk") Then
								objFSO.DeleteFile(strStartupLocation & "\" & strStartupName & ".lnk")
								ElseIf objFSO.FileExists(strStartupLocation & "\" & strCheckFile) Then
									objFSO.DeleteFile(strStartupLocation & "\" & strCheckFile)
						End If
						Else
							Select Case intRegType
								Case 1
									objReg.DeleteValue HKEY_USERS, strStartupLocation, strStartupName
								Case 2
									objReg.DeleteValue HKEY_LOCAL_MACHINE, strStartupLocation, strStartupName
							End Select
					End If
				End If
			Next
			MsgBox "You have removed the following Startup Item(s) on " & strPC & ": " & vbCrLf & _
			strMsg2,vbInformation,"PC Management Utility" 
			RefreshStartupItems()
			Else
				For Each strInputDelete in StartupForm
					strInputDelete.Checked = False
				Next
				StartupTab.InnerHTML = DataArea.InnerHTML
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ChooseStartupFile()
    '#	PURPOSE........:	Allows selection of new Startup Item
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	Will add entry to registry at:
	'#						HKLM\Software\Micrsoft\Windows\CurrentVersion\Run
    '#--------------------------------------------------------------------------
	Sub ChooseStartupFile()
		intWinVer = CheckWinVer(".")
		Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\default:StdRegProv")
		
		If intWinVer = 2 Then
			Set objCD = CreateObject("UserAccounts.CommonDialog")
			objCD.Filter = "All Files (*.*)|*.*|Executables (*.exe)|*.exe"
			objCD.FilterIndex = 2
			objCD.InitialDir = "\\" & strPC & "\c$"
			
			initCD = objCD.ShowOpen
			
			If initCD = False Then
				Exit Sub
				Else
					strAddStartupFile = objCD.FileName
			End If
			Else
				strAddStartupFile = InputBox("Please enter the required Startup item path and filename in the format:" & _
				vbCrLf & vbCrLf & "C:\Folder\Subfolder\Program.exe", "PC Management Utility")
				If strAddStartupFile = "" Then Exit Sub
				strAddStartupFileCheck = Replace(strAddStartupFile, "c:", "\\" & strPC & "\c$")
				If NOT objFSO.FileExists(strAddStartupFileCheck) Then
					AddStartupPrompt = MsgBox("This file does not exist on the target machine." & vbCrLf & vbCrLf & _
					"Would you like to add this startup entry anyway?", _
					vbInformation+vbYesNo,"PC Management Utility")
					If AddStartupPrompt = vbNo Then Exit Sub
				End If
		End If
		
		strAddStartupFile = Replace(UCase(strAddStartupFile), "\\" & strPC & "\C$", "C:")
		arrAddStartupFile = Split(strAddStartupFile, "\")
		strAddStartup = arrAddStartupFile(UBound(arrAddStartupFile))
		
		If InStr(strAddStartup, " ") > 0 Then
			arrAddStartup = Split(strAddStartup, " ")
			strAddStartup = ""
			For i = 0 to UBound(arrAddStartup)
				strUCaseLetter = Left(arrAddStartup(i), 1)
				strUCaseLetter = UCase(strUCaseLetter)
				intEndLength = Len(arrAddStartup(i)) - 1
				strEndString = Right(arrAddStartup(i), intEndLength)
				strEndString = LCase(strEndString)
				arrAddStartup(i) = strUCaseLetter & strEndString
				If i <> 0 Then
					strAddStartup = strAddStartup & " " & arrAddStartup(i)
					Else
						strAddStartup = arrAddStartup(i)
				End If
			Next
			Else
				strUCaseLetter = Left(strAddStartup, 1)
				strUCaseLetter = UCase(strUCaseLetter)
				intEndLength = Len(strAddStartup) - 1
				strEndString = Right(strAddStartup, intEndLength)
				strEndString = LCase(strEndString)
				strAddStartup = strUCaseLetter & strEndString
		End If
		If InStr(strAddStartup, ".") > 0 Then
			arrAddStartup = Split(strAddStartup, ".")
			strAddStartup = ""
			For i = 0 To UBound(arrAddStartup) - 1
				strAddStartup = strAddStartup & arrAddStartup(i)
			Next
		End If
		
		strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run\"
		strKeyValue = strAddStartup
		strValue = strAddStartupFile
		Err.Clear
		objReg.SetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strKeyValue, strValue
		If Err.Number = 0 Then
			MsgBox "You have added the following Startup Item on " & strPC & ": " & vbCrLf & vbCrLf & _
			strAddStartup & " (" & LCase(strAddStartupFile) & ")", vbInformation,"PC Management Utility"
			RefreshStartupItems()
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	RefreshStartupItems()
    '#	PURPOSE........:	Refreshes item in Startup Items tab
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub RefreshStartupItems()
		ShowStartupInfo False
		ShowStartupInfo True
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ExportStartupDetails()
    '#	PURPOSE........:	Export the details for the Startup Items
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ExportStartupDetails()
		On Error Resume Next
		StartupTab.InnerHTML = Nothing
		intStartup = 0
		
		strInput = Document.StartupForm.cbxStartupDelete
		
		For Each strInput in StartupForm
			intStartup = intStartup + 1
		Next
		
		Select Case StartupExport.Value
			Case 1
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\StartupDetails" & strPC & ".csv",True)
				objFile.WriteLine "Startup Items on " & strPC
				objFile.WriteLine ""
				objFile.WriteLine "Total: " & intStartup & " Items"
				objFile.WriteLine ""
				objFile.WriteLine "Startup Item,User,Command,Startup Location"
			Case 2
				Const xlContinuous = 1
				Const xlThin = 2
				Const xlAutomatic = -4105
				
				strExcelPath = objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe\")
			   
				If strExcelPath = "" Then
					MsgBox "Unable to export. Excel does not appear to be installed.", vbExclamation, "PC Management Utility"
					Exit Sub
				End If
				
				Set objExcel = CreateObject("Excel.Application")
				objExcel.Visible = False
				Set objWorkBook = objExcel.WorkBooks.Add
				Set objWorksheet = objWorkbook.Worksheets(1)
				objExcel.DisplayAlerts = False
				For i = 1 to 3
					objWorkbook.Worksheets(2).Delete
				Next
				objExcel.DisplayAlerts = True
				objWorksheet.Name = "Startup Details"
				
				objWorkSheet.Cells(1, 1) = "Startup Items on " & strPC
				objWorkSheet.Cells(3, 1) = "Total: " & intStartup & " Items"

				intStartRow = 6
				
				objWorkSheet.Cells(5, 1) = "Startup Item"
				objWorkSheet.Cells(5, 2) = "User"
				objWorkSheet.Cells(5, 3) = "Command"
				objWorkSheet.Cells(5, 4) = "Startup Location"
			Case 3
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\StartupDetails" & strPC & ".htm",True)
				objFile.WriteLine "<style type=""text/css"">"
				objFile.WriteLine "body{background-color:#CEF0FF;}"
				objFile.WriteLine "table.export{border-width:1px;border-spacing:1px;border-style:solid;border-color:gray;border-collapse:collapse;}"
				objFile.WriteLine "table.export th{border-width:1px;padding:1px;border-style:solid;border-color:gray;padding:2px 7px 2px 7px;}"
				objFile.WriteLine "table.export td{border-width:1px;padding:1px;border-style:dotted;border-color:gray;padding:2px 7px 2px 7px;}"
				objFile.WriteLine ".backtotop a {font-size:0.9em;}"
				objFile.WriteLine "</style>"
				objFile.WriteLine "<div style=""font-weight:bold;""><a name =""top"">Startup Items on " & strPC & "</a><p>"
				objFile.WriteLine "Total: " & intStartup & " Items<p></div>"
				objFile.WriteLine "<table class=""export"">"
				objFile.WriteLine "	<tr>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Startup Item"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Google"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			User"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Command"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Startup Location"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "	</tr>"
			Case 4
				intColumnIndex = 17
				intColumnIndex2 = 9
				For Each strInput in StartupForm
					arrInput = Split(strInput.Value, "|")
					strStartupName = Trim(arrInput(0))
					strStartupUser = arrInput(1)
					strStartupLocation = arrInput(2)
					If Len(strStartupName) > intColumnIndex - 5 Then intColumnIndex = Len(strStartupName) + 5
					If Len(strStartupUser) > intColumnIndex2 - 5 Then intColumnIndex2 = Len(strStartupUser) + 5
				Next
				
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\StartupDetails" & strPC & ".txt",True)
				objFile.WriteLine "Startup Items on " & strPC
				objFile.WriteLine ""
				objFile.WriteLine "Total: " & intStartup & " Items"
				objFile.WriteLine ""
				objFile.WriteLine "Startup Item" & _
				String(intColumnIndex - 12, " ") & "User" & _
				String(intColumnIndex2 - 4, " ") & "Startup Location"
		End Select
		
		For Each strInput in StartupForm
			On Error Resume Next
			arrInput = Split(strInput.Value, "|")
			strStartupName = Trim(arrInput(0))
			strStartupUser = arrInput(1)
			strStartupLocation = arrInput(2)
			strStartupCommand = arrInput(3)
			
			Select Case StartupExport.Value
				Case 1
					strStartupName = EncodeCsv(strStartupName)
					strStartupUser = EncodeCsv(strStartupUser)
					strStartupLocation = EncodeCsv(strStartupLocation)
					strStartupCommand = EncodeCsv(strStartupCommand)
				
					strCSV = strCSV & strStartupName & "," & _
					strStartupUser & "," & strStartupCommand & "," & _
					strStartupLocation & vbCrLf
				Case 2
					objWorkSheet.Cells(intStartRow, 1) = strStartupName
					objWorkSheet.Cells(intStartRow, 2) = strStartupUser
					objWorkSheet.Cells(intStartRow, 3) = strStartupCommand
					objWorkSheet.Cells(intStartRow, 4) = strStartupLocation
					intStartRow = intStartRow + 1
				Case 3
					objFile.WriteLine "	<tr>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strStartupName
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			<a target=_blank href=""http://www.google.com/search?q=" & _
					strStartupName & """>Search</a>"
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strStartupUser
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strStartupCommand
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strStartupLocation
					objFile.WriteLine "		</td>"					
					objFile.WriteLine "	</tr>"
				Case 4
					strTxt = strTxt & strStartupName & _
					String(intColumnIndex - Len(strStartupName), " ") & strStartupUser & _
					String(intColumnIndex2 - Len(strStartupUser), " ") & strStartupLocation & vbCrLf
			End Select
			strStartupName = ""
			strStartupUser = ""
			strStartupLocation = ""
			strStartupCommand = ""
		Next
			
		Select Case StartupExport.Value
			Case 1
				objFile.WriteLine strCSV
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\StartupDetails" & strPC & ".csv"
			Case 2
				Set objRange = objWorkSheet.Range("A1:Z5")
				Set objRange2 = objWorkSheet.Range("A5:D" & intStartRow - 1)
				Set objRangeH = objWorksheet.Range("A5:D5")
				
				objRange.Font.Bold = True
				objRange2.Borders.LineStyle = xlContinuous
				objRange2.Borders.Weight = xlThin
				objRange2.Borders.ColorIndex = xlAutomatic
				objRangeH.AutoFilter
				
				objWorksheet.Range("A6").Select
				objExcel.ActiveWindow.FreezePanes = "True"
				objWorksheet.Range("A1").Select
				
				objWorkSheet.Columns("A:ZZ").EntireColumn.AutoFit
				objExcel.DisplayAlerts = False
				objExcel.ActiveWorkbook.SaveAs(strTemp & "\SKB\StartupDetails" & strPC & ".xls")
				objExcel.Visible = True
				Set objExcel = Nothing
			Case 3
				objFile.WriteLine "</table>"
				objFile.WriteLine "<p class=""backtotop""><a href=""" & strHTMLTempDir & "/SKB/StartupDetails" & _
				strPC & ".htm#top"">[..back to top..]</a></p>"
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\StartupDetails" & strPC & ".htm"
			Case 4
				objFile.WriteLine strTxt
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\StartupDetails" & strPC & ".txt"
		End Select
		
		StartupExport.Value = 0

		StartupTab.InnerHTML = DataArea.InnerHTML
	End Sub

	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ShowPCActions()
    '#	PURPOSE........:	Displays the Actions tab
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	ShowPCActions()
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ShowPCActions()
		tab1.bgcolor="#eeeeee"
		tab2.bgcolor="#eeeeee"
		tab3.bgcolor="#eeeeee"
		tab4.bgcolor="#eeeeee"
		tab5.bgcolor="#eeeeee"
		tab6.bgcolor="#cccccc"
		DataArea.InnerHTML = "<b><u>" & strPC & "</u></b><p>" & _
		"<select size=""15"" name=""Actions"" style=""width:100%"" onChange=""PreExecutionAction()"" onDblClick=""DblClickAction()"">" & _
		strActionList & _
		"</select><br>" & _
		"<input id=runbutton  class=""button"" type=""button"" value=""Execute"" " & _
		"name=""btnExecute"" title=""Execute highlighted action"" " & _
		"onClick=""ExecuteAction()"" onMouseDown=""ChangeButtonColour(btnExecute)"" onMouseUp=""RevertButtonColour(btnExecute)"">" & _
		"<input id=runbutton  class=""button"" type=""button"" value=""Stop"" " & _
		"name=""btnStop"" style=""cursor:default""" & _
		"Disabled=""True"" onclick=""StopAction"" onMouseDown=""ChangeButtonColour(btnStop)"" " & _
		"onMouseUp=""RevertButtonColour(btnStop)""><p>" & _
		"<div align=""center""><span id=""WaitMessage""><hr></span></div>"	
	End Sub
	
	Sub PreExecutionAction()
		Select Case Actions.Value
			Case 2
				btnExecute.Disabled = True
				btnExecute.style.cursor = "default"
				btnExecute.title = "Complete acton below"
				EnableDisableIEProxy()
			Case 3
				btnExecute.Disabled = True
				btnExecute.style.cursor = "default"
				btnExecute.title = "Complete acton below"
				EnableDisableRDP()
			Case 5
				btnExecute.Disabled = True
				btnExecute.style.cursor = "default"
				btnExecute.title = "Complete acton below"
				OpenShare()
			Case 6
				btnExecute.Disabled = True
				btnExecute.style.cursor = "default"
				btnExecute.title = "Complete acton below"
				PrePingMachine()
			Case 7
				btnExecute.Disabled = True
				btnExecute.style.cursor = "default"
				btnExecute.title = "Complete acton below"
				ShutdownRestartPC()
			Case 9
				btnExecute.Disabled = True
				btnExecute.style.cursor = "default"
				btnExecute.title = "Complete acton below"
				ChangePCDescription()
			Case 10
				btnExecute.Disabled = True
				btnExecute.style.cursor = "default"
				btnExecute.title = "Complete acton below"
				ChangeLocalAdminPassword()
			Case 14
				btnExecute.Disabled = True
				btnExecute.style.cursor = "default"
				btnExecute.title = "Complete acton below"
				GetMSProductKeys()
			Case 16
				btnExecute.Disabled = True
				btnExecute.style.cursor = "default"
				btnExecute.title = "Complete acton below"
				SystemRestore()
			Case 17
				btnExecute.Disabled = True
				btnExecute.style.cursor = "default"
				btnExecute.title = "Complete acton below"
				ListUpdates()
			Case 18
				btnExecute.Disabled = True
				btnExecute.style.cursor = "default"
				btnExecute.title = "Complete acton below"
				CleanProfiles()
			Case 23
				btnExecute.Disabled = True
				btnExecute.style.cursor = "default"
				btnExecute.title = "Complete acton below"
				ShowUserAccountsInfo()
			Case Else
				btnExecute.Disabled = False
				btnExecute.style.cursor = "default"
				btnExecute.title = "Execute highlighted action"
				WaitMessage.InnerHTML = "<hr>"
		End Select
	End Sub
	
	Sub DblClickAction()
		Select Case Actions.Value
			Case 1
				WaitMessage.InnerHTML = "<hr>"
				CopyProfile()
			Case 2
				WaitMessage.InnerHTML = "<hr>"
				EnableDisableIEProxy()
			Case 3
				WaitMessage.InnerHTML = "<hr>"
				EnableDisableRDP()
			Case 4
				WaitMessage.InnerHTML = "<hr>"
				ManagePC()
			Case 8
				WaitMessage.InnerHTML = "<hr>"
				ViewProfiles()
			Case 11
				WaitMessage.InnerHTML = "<hr>"
				ClearAppEventLog()
			Case 12
				WaitMessage.InnerHTML = "<hr>"
				DeleteTempFiles()
			Case 14
				WaitMessage.InnerHTML = "<hr>"
				GetMSProductKeys()
			Case 15
				WaitMessage.InnerHTML = "<hr>"
				ProcessAudit()
			Case 16
				WaitMessage.InnerHTML = "<hr>"
				SystemRestore()
			Case 19
				WaitMessage.InnerHTML = "<hr>"
				ExpInventory()
			Case 20
				WaitMessage.InnerHTML = "<hr>"
				DeleteStartupItems()
			Case 21
				WaitMessage.InnerHTML = "<hr>"
				RunPSExecCommand()
			Case 22
				WaitMessage.InnerHTML = "<hr>"
				RunCustomCommand()
		End Select
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	Setup()
    '#	PURPOSE........:	Allows user to define setup options for application
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	Dynamically creates new temporary HTA
    '#--------------------------------------------------------------------------
	Sub Setup()
		On Error Resume Next
		booRunAs = objShell.RegRead(strRegStart & "\Software\SKB\PCManagementUtil\Setup\booRunAs")
		If booRunAs = 1 Then
			arrCommands = Split(objPCManage.commandLine, "*")
			strTempLoc = arrCommands(UBound(arrCommands))
			Else
				strTempLoc = strTemp
		End If
		If NOT objFSO.FolderExists(strTempLoc & "\SKB") Then
			Set objFolder = objFSO.CreateFolder(strTempLoc & "\SKB")
		End If
		Set objFile = objFSO.CreateTextFile(strTempLoc & "\SKB\Setup.hta",True)
		objFile.WriteLine "<html>"
		objFile.WriteLine "<head>"
		objFile.WriteLine "<script language=""VBScript"">"
		objFile.WriteLine "	intLeft = window.screenLeft"
		objFile.WriteLine "	intTop = window.screenTop"
		objFile.WriteLine "	window.moveTo -2000,-2000"
		objFile.WriteLine "	window.ResizeTo 1,1"
		objFile.WriteLine "<" & Chr(47) & "Script>"
		objFile.WriteLine ""
		objFile.WriteLine "<title>Setup Options</title>"
		objFile.WriteLine ""
		objFile.WriteLine "<HTA:APPLICATION"
		objFile.WriteLine "  APPLICATIONNAME=""Setup"""
		objFile.WriteLine "  ID=""objSetup"""
		objFile.WriteLine "  VERSION=""1.0"""
		objFile.WriteLine "  BORDER=""dialog"""
		objFile.WriteLine "  SYSMENU=""no"""
		objFile.WriteLine "  MAXIMIZEBUTTON=""no"""
		objFile.WriteLine "  SINGLEINSTANCE=""yes"""
		objFile.WriteLine "  CONTEXTMENU=""no"""
		objFile.WriteLine "  SCROLL=""no""/>"
		objFile.WriteLine "  "
		objFile.WriteLine "<style type=""text/css"" media=""all"">"
		objFile.WriteLine ""
		objFile.WriteLine "body {"
		objFile.WriteLine "	font-size: 0.8em;"
		objFile.WriteLine "}"
		objFile.WriteLine ""
		objFile.WriteLine "#tab1 {"
		objFile.WriteLine "	text-decoration: none;"
		objFile.WriteLine "	display: block;"
		objFile.WriteLine "	padding: 0.24em 1em;"
		objFile.WriteLine "	width: 25%;"
		objFile.WriteLine "	text-align: center;"
		objFile.WriteLine "	cursor: hand;"
		objFile.WriteLine "}"
		objFile.WriteLine ""
		objFile.WriteLine "#tab2 {"
		objFile.WriteLine "	text-decoration: none;"
		objFile.WriteLine "	display: block;"
		objFile.WriteLine "	padding: 0.24em 1em;"
		objFile.WriteLine "	width: 25%;"
		objFile.WriteLine "	text-align: center;"
		objFile.WriteLine "	cursor: hand;"
		objFile.WriteLine "}"
		objFile.WriteLine ""
		objFile.WriteLine "#tab3 {"
		objFile.WriteLine "	text-decoration: none;"
		objFile.WriteLine "	display: block;"
		objFile.WriteLine "	padding: 0.24em 1em;"
		objFile.WriteLine "	width: 25%;"
		objFile.WriteLine "	text-align: center;"
		objFile.WriteLine "	cursor: hand;"
		objFile.WriteLine "}"
		objFile.WriteLine ""
		objFile.WriteLine "#tab4 {"
		objFile.WriteLine "	text-decoration: none;"
		objFile.WriteLine "	display: block;"
		objFile.WriteLine "	padding: 0.24em 1em;"
		objFile.WriteLine "	width: 25%;"
		objFile.WriteLine "	text-align: center;"
		objFile.WriteLine "	cursor: hand;"
		objFile.WriteLine "}"
		objFile.WriteLine ""
		objFile.WriteLine "#header ul {"
		objFile.WriteLine "	list-style: none;"
		objFile.WriteLine "	padding: 0;"
		objFile.WriteLine "	margin: 0;"
		objFile.WriteLine "}"
		objFile.WriteLine ""
		objFile.WriteLine "#header li {"
		objFile.WriteLine "	float: left;"
		objFile.WriteLine "	border: 1px solid #bbb;"
		objFile.WriteLine "	margin: 0;"
		objFile.WriteLine "}"
		objFile.WriteLine ""
		objFile.WriteLine "#dataarea {"
		objFile.WriteLine "	border: 1px solid black;"
		objFile.WriteLine "	border-top-width: 0;"
		objFile.WriteLine "	width:100%;"
		objFile.WriteLine "	height:267px;"
		objFile.WriteLine "	position:absolute;"
		objFile.WriteLine "	top:37px;"
		objFile.WriteLine " }"
		objFile.WriteLine ""
		objFile.WriteLine "</style>"
		objFile.WriteLine ""
		objFile.WriteLine "  "
		objFile.WriteLine "</head>"
		objFile.WriteLine ""
		objFile.WriteLine "<script language=""VBScript"">"
		objFile.WriteLine ""
		objFile.WriteLine "	Set objShell = CreateObject(""WScript.Shell"")"
		objFile.WriteLine "	Set objFSO = CreateObject(""Scripting.FileSystemObject"")"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub tab1Click()"
		objFile.WriteLine "		If tab1.style.fontweight <> ""bold"" Then "
		objFile.WriteLine "			On Error Resume Next"
		objFile.WriteLine "			ChangeColours(1)"
		objFile.WriteLine "			strHTML = ""<br><table width=""""100%"""">"""
		objFile.WriteLine "			strHTML = strHTML & ""	<tr>"""
		objFile.WriteLine "			strHTML = strHTML & ""		<td style=""""vertical-align:top;width:50%;"""">"""
		objFile.WriteLine "			strHTML = strHTML & ""			<b>Initial Queries</b><br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""checkbox"""" name=""""cbxPCInfo"""" title=""""Query PC Info on initial PC Search"""" checked=true disabled=true> """
		objFile.WriteLine "			strHTML = strHTML & ""			PC Info<br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""checkbox"""" name=""""cbxSoftware"""" title=""""Query Software Info on initial PC Search"""" onClick=InitialQueries>"""
		objFile.WriteLine "			strHTML = strHTML & ""			Software<br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""checkbox"""" name=""""cbxProcesses"""" title=""""Query Processes Info on initial PC Search"""" onClick=InitialQueries>"""
		objFile.WriteLine "			strHTML = strHTML & ""			Processes<br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""checkbox"""" name=""""cbxServices"""" title=""""Query Services Info on initial PC Search"""" onClick=InitialQueries>"""
		objFile.WriteLine "			strHTML = strHTML & ""			Services<br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""checkbox"""" name=""""cbxStartup"""" title=""""Query Startup Info on initial PC Search"""" onClick=InitialQueries>"""
		objFile.WriteLine "			strHTML = strHTML & ""			Startup Items<p>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<b>Default Search View</b><br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<select name=""""SearchView"""" style=""""width:227"""" title=""""Change the Default Search View"""" onChange=""""DefaultSearchView()"""">"""
		objFile.WriteLine "			strHTML = strHTML & ""				<option value=""""1"""">IP Range Search</option>"""
		objFile.WriteLine "			strHTML = strHTML & ""				<option value=""""2"""">Active Directory Search</option>"""
		objFile.WriteLine "			strHTML = strHTML & ""			</select><p>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<b>Delete Files on Exit?</b><br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""radio"""" name=""""DeleteFiles"""" title=""""Delete all temp / inventory files on exit as created by the PC Management Utility"""" value=""""1"""" onClick=""""DeleteFilesOnExit()""""> Yes&nbsp;&nbsp;&nbsp;"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""radio"""" name=""""DeleteFiles""""  title=""""Do not delete all temp / inventory files on exit as created by the PC Management Utility"""" value=""""0"""" onClick=""""DeleteFilesOnExit()""""> No"""
		objFile.WriteLine "			strHTML = strHTML & ""		</td>"""
		objFile.WriteLine "			strHTML = strHTML & ""		<td style=""""vertical-align:top;width:50%;"""">"""
		objFile.WriteLine "			strHTML = strHTML & ""			<b>Root OU</b><br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type = """"text"""" name = """"txtRootOU"""" size=""""33"""" onKeyUp = """"ChangeRootOUTitle()""""><br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input style=""""width:55px;"""" type = """"button"""" value = """"Set"""" title=""""Set the Root OU for all searches of PCs / Users (where required)"""" onClick=""""SetRootOU()""""><p>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<b>RunAs Different User</b><br> """
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""checkbox"""" name=""""cbxRunAs"""" title=""""RunAs different user"""" onClick=RunAs>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""text"""" name=""""txtRunAs"""" size=""""29"""" onKeyUp=RunAs title=""""Username to Run As on opening app""""><br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			Use Savecred? &nbsp;<input type=""""checkbox"""" name=""""cbxSavecred"""" title=""""Save password after first run"""" onClick=RunAs>"""
		objFile.WriteLine "			strHTML = strHTML & ""		</td>"""
		objFile.WriteLine "			strHTML = strHTML & ""	</tr>"""
		objFile.WriteLine "			strHTML = strHTML & ""</table>"""
		objFile.WriteLine "			DataArea.InnerHTML = strHTML"
		objFile.WriteLine "			PopulateGeneralData()"
		objFile.WriteLine "		End If"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub PopulateGeneralData()"
		objFile.WriteLine "		strQueryChoices = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strQueryChoices"")"
		objFile.WriteLine "		If strQueryChoices = """" Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strQueryChoices"", _"
		objFile.WriteLine "			""1,2,3,4,5,"", ""REG_SZ"""
		objFile.WriteLine "			strQueryChoices = ""1,2,3,4,5,"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If InStr(strQueryChoices, ""2"") > 0 Then cbxSoftware.Checked = True"
		objFile.WriteLine "		If InStr(strQueryChoices, ""3"") > 0 Then cbxProcesses.Checked = True"
		objFile.WriteLine "		If InStr(strQueryChoices, ""4"") > 0 Then cbxStartup.Checked = True"
		objFile.WriteLine "		If InStr(strQueryChoices, ""5"") > 0 Then cbxServices.Checked = True"
		objFile.WriteLine "		intSearchView = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\intSearchView"")"
		objFile.WriteLine "		If intSearchView = """" Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\intSearchView"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			intSearchView = 1"
		objFile.WriteLine "		End If"
		objFile.WriteLine "		SearchView.Value = intSearchView"
		objFile.WriteLine "		booDeleteTemp = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\booDeleteTemp"")"
		objFile.WriteLine "		If booDeleteTemp = """" Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\booDeleteTemp"", _"
		objFile.WriteLine "			""0"", ""REG_SZ"""
		objFile.WriteLine "			booDeleteTemp = ""0"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		For Each objButton in DeleteFiles"
		objFile.WriteLine "			If objButton.Value = booDeleteTemp Then"
		objFile.WriteLine "				objButton.Checked = True"
		objFile.WriteLine "			End If"
		objFile.WriteLine "		Next"
		objFile.WriteLine "		"
		objFile.WriteLine "		strRootOU = objShell.RegRead(""HKCU\Software\SKB\PCManagementUtil\Setup\strRootOU"")"
		objFile.WriteLine "		If strRootOU = """" Then"
		objFile.WriteLine "			Set objRootDSE = GetObject(""LDAP://RootDSE"")"
		objFile.WriteLine "			strDNSDomain = objRootDSE.Get(""defaultNamingContext"")"
		objFile.WriteLine "			strRootOU = strDNSDomain"
		objFile.WriteLine "			objShell.RegWrite ""HKCU\Software\SKB\PCManagementUtil\Setup\strRootOU"", strRootOU, ""REG_SZ"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		txtRootOU.Value = strRootOU"
		objFile.WriteLine "		txtRootOU.Title = strRootOU"
		objFile.WriteLine "		"
		objFile.WriteLine "		strRunAsUser = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strRunAsUser"")"
		objFile.WriteLine "		strRunAs = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strRunAs"")"
		objFile.WriteLine "		setbooRunAs = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\booRunAs"")"
		objFile.WriteLine "		"
		objFile.WriteLine "		If IsNull(booRunAs) OR booRunAs = """" Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\booRunAs"", _"
		objFile.WriteLine "			""0"", ""REG_SZ"""
		objFile.WriteLine "			booRunAs = 0"
		objFile.WriteLine "		End If"
		objFile.WriteLine "		"
		objFile.WriteLine "		If booRunAs = 1 Then"
		objFile.WriteLine "			cbxRunAs.Checked = True"
		objFile.WriteLine "			txtRunAs.Value = strRunAsUser"
		objFile.WriteLine "			Else"
		objFile.WriteLine "				txtRunAs.Style.backgroundcolor = ""#dddddd"""
		objFile.WriteLine "				txtRunAs.Disabled = True"
		objFile.WriteLine "				txtRunAs.Value = strRunAsUser"
		objFile.WriteLine "				cbxRunAs.Checked = False"
		objFile.WriteLine "				cbxSaveCred.Checked = False"
		objFile.WriteLine "				cbxSaveCred.Disabled = True"
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If InStr(LCase(strRunAs), ""/savecred"") > 0 Then cbxSaveCred.Checked = True"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub InitialQueries()"
		objFile.WriteLine "		On Error Resume Next"
		objFile.WriteLine "		booSoftware = False"
		objFile.WriteLine "		booProcesses = False"
		objFile.WriteLine "		booServices = False"
		objFile.WriteLine "		booStartup = False"
		objFile.WriteLine "		strQueryChoices = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strQueryChoices"")"
		objFile.WriteLine "		If strQueryChoices = """" Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strQueryChoices"", _"
		objFile.WriteLine "			""1,2,3,4,5,"", ""REG_SZ"""
		objFile.WriteLine "			strQueryChoices = ""1,2,3,4,5,"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If InStr(strQueryChoices, ""2"") > 0 Then booSoftware = True"
		objFile.WriteLine "		If InStr(strQueryChoices, ""3"") > 0 Then booProcesses = True"
		objFile.WriteLine "		If InStr(strQueryChoices, ""4"") > 0 Then booStartup = True"
		objFile.WriteLine "		If InStr(strQueryChoices, ""5"") > 0 Then booServices = True"
		objFile.WriteLine "		"
		objFile.WriteLine "		If cbxSoftware.Checked AND booSoftware = False Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strQueryChoices"", _"
		objFile.WriteLine "			strQueryChoices & ""2,"", ""REG_SZ"""
		objFile.WriteLine "			ElseIf cbxSoftware.Checked = False AND booSoftware = True Then"
		objFile.WriteLine "				strQueryChoices = Replace(strQueryChoices, ""2,"", """")"
		objFile.WriteLine "				objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strQueryChoices"", _"
		objFile.WriteLine "				strQueryChoices, ""REG_SZ"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If cbxProcesses.Checked AND booProcesses = False Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strQueryChoices"", _"
		objFile.WriteLine "			strQueryChoices & ""3,"", ""REG_SZ"""
		objFile.WriteLine "			ElseIf cbxProcesses.Checked = False AND booProcesses = True Then"
		objFile.WriteLine "				strQueryChoices = Replace(strQueryChoices, ""3,"", """")"
		objFile.WriteLine "				objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strQueryChoices"", _"
		objFile.WriteLine "				strQueryChoices, ""REG_SZ"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If cbxStartup.Checked AND booStartup = False Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strQueryChoices"", _"
		objFile.WriteLine "			strQueryChoices & ""4,"", ""REG_SZ"""
		objFile.WriteLine "			ElseIf cbxStartup.Checked = False AND booStartup = True Then"
		objFile.WriteLine "				strQueryChoices = Replace(strQueryChoices, ""4,"", """")"
		objFile.WriteLine "				objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strQueryChoices"", _"
		objFile.WriteLine "				strQueryChoices, ""REG_SZ"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If cbxServices.Checked AND booServices = False Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strQueryChoices"", _"
		objFile.WriteLine "			strQueryChoices & ""5,"", ""REG_SZ"""
		objFile.WriteLine "			ElseIf cbxServices.Checked = False AND booServices = True Then"
		objFile.WriteLine "				strQueryChoices = Replace(strQueryChoices, ""5,"", """")"
		objFile.WriteLine "				objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strQueryChoices"", _"
		objFile.WriteLine "				strQueryChoices, ""REG_SZ"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub DefaultSearchView()"
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\intSearchView"", _"
		objFile.WriteLine "		SearchView.Value, ""REG_SZ"""
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub DeleteFilesOnExit()"
		objFile.WriteLine "		For Each objButton in DeleteFiles"
		objFile.WriteLine "			If objButton.Checked Then"
		objFile.WriteLine "				objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\booDeleteTemp"", _"
		objFile.WriteLine "				objButton.Value, ""REG_SZ"""
		objFile.WriteLine "			End If"
		objFile.WriteLine "		Next"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub SetRootOU()"
		objFile.WriteLine "		On Error Resume Next"
		objFile.WriteLine "		Set objRootDSE = GetObject(""LDAP://RootDSE"")"
		objFile.WriteLine "		"
		objFile.WriteLine "		strRootOU = txtRootOU.Value"
		objFile.WriteLine "		"
		objFile.WriteLine "		Set adoCommand = CreateObject(""ADODB.Command"")"
		objFile.WriteLine "		Set adoConnection = CreateObject(""ADODB.Connection"")"
		objFile.WriteLine "		adoConnection.Provider = ""ADsDSOObject"""
		objFile.WriteLine "		adoConnection.Open ""Active Directory Provider"""
		objFile.WriteLine "		adoCommand.ActiveConnection = adoConnection"
		objFile.WriteLine ""
		objFile.WriteLine "		strFilter = ""(&(cn="" & strPC & ""))"""
		objFile.WriteLine "		strQuery = ""<LDAP://"" & strRootOU & "">;"" & strFilter & _"
		objFile.WriteLine "		"";distinguishedName,objectCategory;subtree"""
		objFile.WriteLine "		"
		objFile.WriteLine "		adoCommand.CommandText = strQuery"
		objFile.WriteLine "		adoCommand.Properties(""Page Size"") = 750"
		objFile.WriteLine "		adoCommand.Properties(""Timeout"") = 60"
		objFile.WriteLine "		adoCommand.Properties(""Cache Results"") = False"
		objFile.WriteLine "		"
		objFile.WriteLine "		Err.Clear"
		objFile.WriteLine "		Set adoRecordset = adoCommand.Execute"
		objFile.WriteLine "		If Err.Number <> 0 Then"
		objFile.WriteLine "			strRootOU = objShell.RegRead(""HKCU\Software\SKB\PCManagementUtil\Setup\strRootOU"")"
		objFile.WriteLine "			MsgBox ""'"" & txtRootOU.Value & ""' is not a valid Root OU."" & vbCrLf & vbCrLf & _"
		objFile.WriteLine "			""An example of a valid Root OU would be 'ou=uk,dc=acme,dc=group'. If you are having difficulties "" & _"
		objFile.WriteLine "			""determining the required Root OU then a good utility to use would be ADExplorer from Sysinternals, "" & _"
		objFile.WriteLine "			""which will provide you with the distinguished name of any OU / object in your AD forest."", _"
		objFile.WriteLine "			vbExclamation, ""PC Management Utility"""
		objFile.WriteLine "			txtRootOU.Value = strRootOU"
		objFile.WriteLine "			Else"
		objFile.WriteLine "				objShell.RegWrite ""HKCU\Software\SKB\PCManagementUtil\Setup\strRootOU"", strRootOU, ""REG_SZ"""
		objFile.WriteLine "				MsgBox ""The Root OU has now been set to '"" & strRootOU & ""'"", vbInformation, ""PC Management Utility""	"
		objFile.WriteLine "		End If"
		objFile.WriteLine "		adoRecordSet.Close"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub ChangeRootOUTitle()"
		objFile.WriteLine "		txtRootOU.Title = txtRootOU.Value"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub RunAs()"
		objFile.WriteLine "		If cbxRunAs.Checked Then"
		objFile.WriteLine "			cbxSaveCred.Disabled = False"
		objFile.WriteLine "			arrCommands = Split(objSetup.commandLine, ""*"")"
		objFile.WriteLine "			strScriptLocation = arrCommands(1)"
		objFile.WriteLine "			strScriptLocation = Replace(strScriptLocation, ""file:///"", """")"
		objFile.WriteLine "			strScriptLocation = Replace(strScriptLocation, ""%20"", "" "")"
		objFile.WriteLine "			strScriptLocation = Replace(strScriptLocation, ""/"", ""\"")"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\booRunAs"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			txtRunAs.Style.backgroundcolor = ""white"""
		objFile.WriteLine "			txtRunAs.Disabled = False"
		objFile.WriteLine "			strRunAsUser = txtRunAs.Value"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strRunAsUser"", _"
		objFile.WriteLine "			strRunAsUser, ""REG_SZ"""
		objFile.WriteLine "			If cbxSaveCred.Checked Then"
		objFile.WriteLine "				strRunAs = ""%windir%\system32\runas.exe /user:"" & strRunAsUser & _"
		objFile.WriteLine "				"" /savecred """"mshta \"" & Chr(34) & strScriptLocation & ""\"" & Chr(34) & "" True"""
		objFile.WriteLine "				Else"
		objFile.WriteLine "					strRunAs = ""%windir%\system32\runas.exe /user:"" & strRunAsUser & _"
		objFile.WriteLine "					"" /savecred """"mshta \"" & Chr(34) & strScriptLocation & ""\"" & Chr(34) & "" True"""
		objFile.WriteLine "			End If"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strRunAs"", _"
		objFile.WriteLine "			strRunAs, ""REG_SZ"""
		objFile.WriteLine "			Else"
		objFile.WriteLine "				objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\booRunAs"", _"
		objFile.WriteLine "				""0"", ""REG_SZ"""
		objFile.WriteLine "				txtRunAs.Style.backgroundcolor = ""#dddddd"""
		objFile.WriteLine "				txtRunAs.Disabled = True"
		objFile.WriteLine "				cbxSaveCred.Disabled = True"
		objFile.WriteLine "		End If"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub tab2Click()"
		objFile.WriteLine "		If tab2.style.fontweight <> ""bold"" Then "
		objFile.WriteLine "			ChangeColours(2)"
		objFile.WriteLine "			strHTML = ""<br><table width=""""100%"""">"""
		objFile.WriteLine "			strHTML = strHTML & ""	<tr>"""
		objFile.WriteLine "			strHTML = strHTML & ""		<td style=""""vertical-align:top;width:50%;"""">"""
		objFile.WriteLine "			strHTML = strHTML & ""			<b>Software</b><br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""checkbox"""" name=""""cbxSortSoftware"""" title=""""Sort software list alphabetically"""" onClick=ApplyTabAction>"""
		objFile.WriteLine "			strHTML = strHTML & ""			Sort Software list alphabetically<br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""checkbox"""" name=""""cbxShowSoftware"""" title=""""Show hidden software in software list, note ticking this box will show some duplicate software in list"""" onClick=ApplyTabAction>"""
		objFile.WriteLine "			strHTML = strHTML & ""			Show hidden software in software list<br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""checkbox"""" name=""""cbxUseInstallMonitor"""" title=""""Use software install monitor"""" onClick=ApplyTabAction>"""
		objFile.WriteLine "			strHTML = strHTML & ""			Use software uninstall monitor<p>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<b>Data Displayed</b><br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""checkbox"""" name=""""cbxSoftwareName"""" title=""""Show Software name in Software tab"""" checked=true disabled=true> """
		objFile.WriteLine "			strHTML = strHTML & ""			<i>Software Name</i><br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""checkbox"""" name=""""cbxSoftwareVersion"""" title=""""Show Software version in Software tab"""" onClick=ApplyTabAction>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<i>Software Version</i><br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""checkbox"""" name=""""cbxSoftwareVendor"""" title=""""Show Software vendor in Software tab"""" onClick=ApplyTabAction>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<i>Software Vendor</i><br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""checkbox"""" name=""""cbxSoftwareInstallDate"""" title=""""Show Software install date in Software tab"""" onClick=ApplyTabAction>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<i>Software Install Date</i>"""
		objFile.WriteLine "			strHTML = strHTML & ""		</td>"""
		objFile.WriteLine "			strHTML = strHTML & ""		<td style=""""vertical-align:top;width:50%;"""">"""
		objFile.WriteLine "			strHTML = strHTML & ""			<b>Processes</b><br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""checkbox"""" name=""""cbxSortProcess"""" title=""""Sort process list alphabetically"""" onClick=ApplyTabAction>"""
		objFile.WriteLine "			strHTML = strHTML & ""			Sort Process list alphabetically<p>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<b>Services</b><br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""checkbox"""" name=""""cbxSortServices"""" title=""""Sort services list alphabetically"""" onClick=ApplyTabAction>"""
		objFile.WriteLine "			strHTML = strHTML & ""			Sort Services list alphabetically<p>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<b>Startup Items</b><br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""checkbox"""" name=""""cbxSortStartup"""" title=""""Sort startup items list alphabetically"""" onClick=ApplyTabAction>"""
		objFile.WriteLine "			strHTML = strHTML & ""			Sort Startup Items list alphabetically<p>"""
		objFile.WriteLine "			strHTML = strHTML & ""		</td>"""
		objFile.WriteLine "			strHTML = strHTML & ""	</tr>"""
		objFile.WriteLine "			strHTML = strHTML & ""	</table>"""
		objFile.WriteLine "			DataArea.InnerHTML = strHTML"
		objFile.WriteLine "			PopulateTabData()"
		objFile.WriteLine "		End If"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub PopulateTabData()"
		objFile.WriteLine "		On Error Resume Next"
		objFile.WriteLine "		booSoftwareVersion = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSoftwareVersion"")"
		objFile.WriteLine "		booSoftwareVendor = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSoftwareVendor"")"
		objFile.WriteLine "		booSoftwareInstallDate = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSoftwareInstallDate"")"
		objFile.WriteLine "		booSortSoftware = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortSoftware"")"
		objFile.WriteLine "		booShowSoftware = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booShowSoftware"")"
		objFile.WriteLine "		booSortProcess = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortProcess"")"
		objFile.WriteLine "		booSortServices = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortServices"")"
		objFile.WriteLine "		booSortStartup = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortStartup"")"
		objFile.WriteLine "		booUseInstallMonitor = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booUseInstallMonitor"")"
		objFile.WriteLine "		"
		objFile.WriteLine "		If booSoftwareVersion = """" OR IsNull(booSoftwareVersion) Then "
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSoftwareVersion"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			booSoftwareVersion = ""1"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If booSoftwareVendor = """" OR IsNull(booSoftwareVendor) Then "
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSoftwareVendor"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			booSoftwareVendor = ""1"""
		objFile.WriteLine "		End If		"
		objFile.WriteLine "		If booSoftwareInstallDate = """" OR IsNull(booSoftwareInstallDate) Then "
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSoftwareInstallDate"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			booSoftwareInstallDate = ""1"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If booSortSoftware = """" OR IsNull(booSortSoftware) Then "
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortSoftware"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			booSortSoftware = ""1"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If booShowSoftware = """" OR IsNull(booShowSoftware) Then "
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booShowSoftware"", _"
		objFile.WriteLine "			""0"", ""REG_SZ"""
		objFile.WriteLine "			booShowSoftware = ""0"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If booSortProcess = """" OR IsNull(booSortProcess) Then "
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortProcess"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			booSortProcess = ""1"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If booSortServices = """" OR IsNull(booSortServices) Then "
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortServices"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			booSortServices = ""1"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If booSortStartup = """" OR IsNull(booSortStartup) Then "
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortStartup"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			booSortStartup = ""1"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If booUseInstallMonitor = """" OR IsNull(booUseInstallMonitor) Then "
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booUseInstallMonitor"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			booUseInstallMonitor = ""1"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		"
		objFile.WriteLine "		If booSoftwareVersion = ""1"" Then"
		objFile.WriteLine "			cbxSoftwareVersion.Checked = True"
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If booSoftwareVendor = ""1"" Then"
		objFile.WriteLine "			cbxSoftwareVendor.Checked = True"
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If booSoftwareInstallDate = ""1"" Then"
		objFile.WriteLine "			cbxSoftwareInstallDate.Checked = True"
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If booSortSoftware = ""1"" Then"
		objFile.WriteLine "			cbxSortSoftware.Checked = True"
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If booSortServices = ""1"" Then"
		objFile.WriteLine "			cbxSortServices.Checked = True"
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If booShowSoftware = ""1"" Then"
		objFile.WriteLine "			cbxShowSoftware.Checked = True"
		objFile.WriteLine "			Else"
		objFile.WriteLine "				cbxShowSoftware.Checked = False"
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If booSortProcess = ""1"" Then"
		objFile.WriteLine "			cbxSortProcess.Checked = True"
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If booSortStartup = ""1"" Then"
		objFile.WriteLine "			cbxSortStartup.Checked = True"
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If booUseInstallMonitor = ""1"" Then"
		objFile.WriteLine "			cbxUseInstallMonitor.Checked = True"
		objFile.WriteLine "		End If"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub ApplyTabAction()"
		objFile.WriteLine "		On Error Resume Next"
		objFile.WriteLine "		booSoftwareVersion = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSoftwareVersion"")"
		objFile.WriteLine "		booSoftwareVendor = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSoftwareVendor"")"
		objFile.WriteLine "		booSoftwareInstallDate = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSoftwareInstallDate"")"
		objFile.WriteLine "		booSortSoftware = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortSoftware"")"
		objFile.WriteLine "		booShowSoftware = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booShowSoftware"")"
		objFile.WriteLine "		booSortProcess = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortProcess"")"
		objFile.WriteLine "		booSortServices = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortServices"")"
		objFile.WriteLine "		booSortStartup = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortStartup"")"
		objFile.WriteLine "		booUseInstallMonitor = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booUseInstallMonitor"")"
		objFile.WriteLine "				"
		objFile.WriteLine "		If cbxSoftwareVersion.Checked AND booSoftwareVersion = ""0"" Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSoftwareVersion"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			ElseIf cbxSoftwareVersion.Checked = False AND booSoftwareVersion = ""1"" Then"
		objFile.WriteLine "				objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSoftwareVersion"", _"
		objFile.WriteLine "				""0"", ""REG_SZ"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If cbxSoftwareVendor.Checked AND booSoftwareVendor = ""0"" Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSoftwareVendor"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			ElseIf cbxSoftwareVendor.Checked = False AND booSoftwareVendor = ""1"" Then"
		objFile.WriteLine "				objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSoftwareVendor"", _"
		objFile.WriteLine "				""0"", ""REG_SZ"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If cbxSoftwareInstallDate.Checked AND booSoftwareInstallDate = ""0"" Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSoftwareInstallDate"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			ElseIf cbxSoftwareInstallDate.Checked = False AND booSoftwareInstallDate = ""1"" Then"
		objFile.WriteLine "				objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSoftwareInstallDate"", _"
		objFile.WriteLine "				""0"", ""REG_SZ"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If cbxSortSoftware.Checked AND booSortSoftware = ""0"" Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortSoftware"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			ElseIf cbxSortSoftware.Checked = False AND booSortSoftware = ""1"" Then"
		objFile.WriteLine "				objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortSoftware"", _"
		objFile.WriteLine "				""0"", ""REG_SZ"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If cbxShowSoftware.Checked AND booShowSoftware = ""0"" Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booShowSoftware"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			ElseIf cbxShowSoftware.Checked = False AND booShowSoftware = ""1"" Then"
		objFile.WriteLine "				objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booShowSoftware"", _"
		objFile.WriteLine "				""0"", ""REG_SZ"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If cbxSortProcess.Checked AND booSortProcess = ""0"" Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortProcess"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			ElseIf cbxSortProcess.Checked = False AND booSortProcess = ""1"" Then"
		objFile.WriteLine "				objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortProcess"", _"
		objFile.WriteLine "				""0"", ""REG_SZ"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If cbxSortServices.Checked AND booSortServices = ""0"" Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortServices"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			ElseIf cbxSortServices.Checked = False AND booSortServices = ""1"" Then"
		objFile.WriteLine "				objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortServices"", _"
		objFile.WriteLine "				""0"", ""REG_SZ"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If cbxSortStartup.Checked AND booSortStartup = ""0"" Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortStartup"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			ElseIf cbxSortStartup.Checked = False AND booSortStartup = ""1"" Then"
		objFile.WriteLine "				objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortStartup"", _"
		objFile.WriteLine "				""0"", ""REG_SZ"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If cbxUseInstallMonitor.Checked AND booUseInstallMonitor = ""0"" Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booUseInstallMonitor"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			ElseIf cbxUseInstallMonitor.Checked = False AND booUseInstallMonitor = ""1"" Then"
		objFile.WriteLine "				objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booUseInstallMonitor"", _"
		objFile.WriteLine "				""0"", ""REG_SZ"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub tab3Click()"
		objFile.WriteLine "		If tab3.style.fontweight <> ""bold"" Then "
		objFile.WriteLine "			ChangeColours(3)"
		objFile.WriteLine "			strHTML = ""<br><table width=""""100%"""">"""
		objFile.WriteLine "			strHTML = strHTML & ""	<tr>"""
		objFile.WriteLine "			strHTML = strHTML & ""		<td style=""""vertical-align:top;width:50%;"""">"""
		objFile.WriteLine "			strHTML = strHTML & ""			<form name=""""InventoryFolder"""">"""
		objFile.WriteLine "			strHTML = strHTML & ""				<b>File Location for Inventory Files</b><br>"""
		objFile.WriteLine "			strHTML = strHTML & ""				<input type = """"text"""" name = """"txtFolder"""" style=""""width:97%"""" onBlur=""""ChooseManualInventoryFolder()"""">"""
		objFile.WriteLine "			strHTML = strHTML & ""				<br><input type = """"button"""" value = """"Browse ..."""" title=""""Browse for a folder in which to save your Exported Inventory files"""" onClick=""""ChooseInventoryFolder()""""></form>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<p><b>Change Computer Description</b><br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""checkbox"""" name=""""cbxQueryAD"""" title=""""Query AD for PC Info / AD Description"""" onClick=ApplyActionsAction>"""
		objFile.WriteLine "			strHTML = strHTML & ""			Query AD for PC Info"""
		objFile.WriteLine "			strHTML = strHTML & ""			<p><br><div align=""""center"""">"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input style=""""width:75%;"""" type=""""button"""" value=""""Clear IP Ranges"""" title=""""Clear all Last 5 IP Ranges"""" onClick=""""ClearIPRanges()""""><br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input style=""""width:75%;"""" type=""""button"""" value=""""Clear PSExec Commands"""" title=""""Clear all PSExec commands"""" onClick=""""ClearPSExecCommands()"""">"""
		objFile.WriteLine "			strHTML = strHTML & ""			</div>"""
		objFile.WriteLine "			strHTML = strHTML & ""		</td>"""
		objFile.WriteLine "			strHTML = strHTML & ""		<td style=""""vertical-align:top;width:50%;"""">"""
		objFile.WriteLine "			strHTML = strHTML & ""			<b>Delete Old User Profiles</b><p>"""
		objFile.WriteLine "			strHTML = strHTML & ""			Protected Profiles:<br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<select name=""""ProtectedProfiles"""" style=""""width:70%"""" title=""""Profiles protected from deletion by the Clean Profiles routine""""></select>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type = """"button"""" style=""""width:27%"""" value=""""Remove"""" title=""""Remove username from protected profiles list"""" onClick=""""RemoveProtectedProfile()""""><br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type = """"text"""" name = """"txtProfile"""" style=""""width:70%"""">"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type = """"button"""" style=""""width:27%"""" value=""""Add"""" title=""""Add username to protected profiles list"""" onClick=""""AddProtectedProfile()""""><hr>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type = """"text"""" name = """"txtProfAge"""" size=""""1"""" style=""""text-align:center;""""> """
		objFile.WriteLine "			strHTML = strHTML & ""			Days to keep old user profiles<br>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type = """"button"""" style=""""width:27%"""" value=""""Set"""" title=""""Set number of days to keep old user profiles"""" onClick=""""ApplyActionsAction()"""">"""
		objFile.WriteLine "			strHTML = strHTML & ""		</td>"""
		objFile.WriteLine "			strHTML = strHTML & ""	</tr>"""
		objFile.WriteLine "			strHTML = strHTML & ""	</table>"""
		objFile.WriteLine "			DataArea.InnerHTML = strHTML"
		objFile.WriteLine "			PopulateActionsData()"
		objFile.WriteLine "		End If"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub PopulateActionsData()"
		objFile.WriteLine "		On Error Resume Next"
		objFile.WriteLine "		strInvDirectory = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strInvDirectory"")"
		objFile.WriteLine "		If strInvDirectory = """" Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strInvDirectory"", _"
		objFile.WriteLine "			""C:\SKB\Inventory"", ""REG_SZ"""
		objFile.WriteLine "			strInvDirectory = ""C:\SKB\Inventory"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		InventoryFolder.txtFolder.Value = strInvDirectory"
		objFile.WriteLine "		strCurrentProfiles = objShell.RegRead(""HKEY_USERS\"" & strSID & _"
		objFile.WriteLine "		""\Software\SKB\PCManagementUtil\Setup\strCurrentProfiles"")"
		objFile.WriteLine "		arrCurrentProfiles = Split(strCurrentProfiles, "";"")"
		objFile.WriteLine "		For i = 0 To UBound(arrCurrentProfiles) - 1"
		objFile.WriteLine "			Set objOption = Document.createElement(""OPTION"")"
		objFile.WriteLine "			objOption.Text = arrCurrentProfiles(i)"
		objFile.WriteLine "			objOption.Value = arrCurrentProfiles(i)"
		objFile.WriteLine "			objOption.Title = ""Username: "" & arrCurrentProfiles(i)"
		objFile.WriteLine "			ProtectedProfiles.Add(objOption)"
		objFile.WriteLine "		Next"
		objFile.WriteLine "		booQueryAD = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booQueryAD"")"
		objFile.WriteLine "		If booQueryAD = """" OR IsNull(booQueryAD) Then "
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booQueryAD"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			booQueryAD = ""1"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		If booQueryAD = ""1"" Then"
		objFile.WriteLine "			cbxQueryAD.Checked = True"
		objFile.WriteLine "		End If"
		objFile.WriteLine "		intProfAge = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\intProfAge"")"
		objFile.WriteLine "		If intProfAge = """" OR IsNull(intProfAge) Then "
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\intProfAge"", _"
		objFile.WriteLine "			""90"", ""REG_SZ"""
		objFile.WriteLine "			intProfAge = ""90"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		txtProfAge.Value = intProfAge"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub ApplyActionsAction()"
		objFile.WriteLine "		On Error Resume Next"
		objFile.WriteLine "		booQueryAD = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booQueryAD"")"
		objFile.WriteLine "		If cbxQueryAD.Checked AND booQueryAD = ""0"" Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booQueryAD"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			ElseIf cbxQueryAD.Checked = False AND booQueryAD = ""1"" Then"
		objFile.WriteLine "				objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booQueryAD"", _"
		objFile.WriteLine "				""0"", ""REG_SZ"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		intProfAge = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\intProfAge"")"
		objFile.WriteLine "		If IsNumeric(txtProfAge.Value) Then"
		objFile.WriteLine "			intValue = Round(txtProfAge.Value)"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\intProfAge"", _"
		objFile.WriteLine "			intValue, ""REG_SZ"""
		objFile.WriteLine "			txtProfAge.Value = intValue"
		objFile.WriteLine "			Else"
		objFile.WriteLine "				txtProfAge.Value = intProfAge"
		objFile.WriteLine "		End If"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub ChooseInventoryFolder()"
		objFile.WriteLine "		strStartDir = ""c:\"""
		objFile.WriteLine "		InventoryFolder.txtFolder.Value = PickFolder(strStartDir)"
		objFile.WriteLine "		If objFSO.FolderExists(InventoryFolder.txtFolder.Value) Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strInvDirectory"", _"
		objFile.WriteLine "			InventoryFolder.txtFolder.Value, ""REG_SZ"""
		objFile.WriteLine "			Else InventoryFolder.txtFolder.Value = objShell.RegRead _"
		objFile.WriteLine "				(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strInvDirectory"")"
		objFile.WriteLine "		End If"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub ChooseManualInventoryFolder()"
		objFile.WriteLine "		If objFSO.FolderExists(InventoryFolder.txtFolder.Value) Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strInvDirectory"", _"
		objFile.WriteLine "			InventoryFolder.txtFolder.Value, ""REG_SZ"""
		objFile.WriteLine "			Else InventoryFolder.txtFolder.Value = objShell.RegRead _"
		objFile.WriteLine "				(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strInvDirectory"")"
		objFile.WriteLine "		End If"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub ClearIPRanges()"
		objFile.WriteLine "		On Error Resume Next"
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\IPRanges\IP1A"""
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\IPRanges\IP1B"""
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\IPRanges\IP2A"""
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\IPRanges\IP2B"""
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\IPRanges\IP3A"""
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\IPRanges\IP3B"""
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\IPRanges\IP4A"""
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\IPRanges\IP4B"""
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\IPRanges\IP5A"""
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\IPRanges\IP5B"""
		objFile.WriteLine "		MsgBox ""All of the IP Ranges have now been cleared"", vbInformation, ""PC Management Utility"""
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub ClearPSExecCommands()"
		objFile.WriteLine "		On Error Resume Next"
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave01"""
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave02"""
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave03"""
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave04"""
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave05"""
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave06"""
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave07"""
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave08"""
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave09"""
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave10"""
		objFile.WriteLine "		MsgBox ""All of the PSExec Commands have now been cleared"", vbInformation, ""PC Management Utility"""
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub AddProtectedProfile()"
		objFile.WriteLine "		On Error Resume Next"
		objFile.WriteLine "		If txtProfile.Value = """" Then Exit Sub"
		objFile.WriteLine "		strCurrentProfiles = objShell.RegRead(""HKEY_USERS\"" & strSID & _"
		objFile.WriteLine "		""\Software\SKB\PCManagementUtil\Setup\strCurrentProfiles"")"
		objFile.WriteLine "		If strCurrentProfiles = """" Then"
		objFile.WriteLine "			For Each objOption in ProtectedProfiles.Options"
		objFile.WriteLine "				objOption.RemoveNode"
		objFile.WriteLine "			Next"
		objFile.WriteLine "		End If"
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strCurrentProfiles"", _"
		objFile.WriteLine "		strCurrentProfiles & txtProfile.Value & "";"", ""REG_SZ"""
		objFile.WriteLine "		Set objOption = Document.createElement(""OPTION"")"
		objFile.WriteLine "		objOption.Text = txtProfile.Value"
		objFile.WriteLine "		objOption.Value = txtProfile.Value"
		objFile.WriteLine "		objOption.Title = ""Username: "" & txtProfile.Value"
		objFile.WriteLine "		ProtectedProfiles.Add(objOption)"
		objFile.WriteLine "		txtProfile.Value = """""
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub RemoveProtectedProfile()"
		objFile.WriteLine "		On Error Resume Next"
		objFile.WriteLine "		strCurrentProfiles = objShell.RegRead(""HKEY_USERS\"" & strSID & _"
		objFile.WriteLine "		""\Software\SKB\PCManagementUtil\Setup\strCurrentProfiles"")"
		objFile.WriteLine "		strCurrentProfiles = Replace(LCase(strCurrentProfiles), LCase(ProtectedProfiles.Value) & "";"", """")"
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strCurrentProfiles"", _"
		objFile.WriteLine "		strCurrentProfiles, ""REG_SZ"""
		objFile.WriteLine "		For Each objOption in ProtectedProfiles.Options"
		objFile.WriteLine "			 If objOption.Value = ProtectedProfiles.Value Then"
		objFile.WriteLine "				objOption.RemoveNode"
		objFile.WriteLine "				Exit Sub"
		objFile.WriteLine "			End If"
		objFile.WriteLine "		Next"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub tab4Click()"
		objFile.WriteLine "		If tab4.style.fontweight <> ""bold"" Then "
		objFile.WriteLine "			ChangeColours(4)"
		objFile.WriteLine "			strHTML = ""<br><table width=""""100%"""">"""
		objFile.WriteLine "			strHTML = strHTML & ""	<tr>"""
		objFile.WriteLine "			strHTML = strHTML & ""		<td style=""""vertical-align:top;width:50%;"""">"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""checkbox"""" name=""""cbxResolveHostNames"""" title=""""Resolve Host Names on searching via IP search. Note if this is not selected then the Available PCs list will be populated with machines that are not available"""" onClick=ApplyOtherAction>"""
		objFile.WriteLine "			strHTML = strHTML & ""			Resolve Host Names in IP Search<p>"""
		objFile.WriteLine "			strHTML = strHTML & ""			<input type=""""checkbox"""" name=""""cbxWMIPrompt"""" title=""""Show 'Slow WMI' prompt on searching for a machine"""" onClick=ApplyOtherAction>"""
		objFile.WriteLine "			strHTML = strHTML & ""			Show 'Slow WMI' prompt"""
		objFile.WriteLine "			strHTML = strHTML & ""		</td>"""
		objFile.WriteLine "			strHTML = strHTML & ""		<td style=""""vertical-align:top;width:50%;"""">"""
		objFile.WriteLine "			strHTML = strHTML & ""		</td>"""
		objFile.WriteLine "			strHTML = strHTML & ""	</tr>"""
		objFile.WriteLine "			strHTML = strHTML & ""	<tr><td>&nbsp;</td></tr>"""
		objFile.WriteLine "			strHTML = strHTML & ""	<tr>"""
		objFile.WriteLine "			strHTML = strHTML & ""		<td style=""""vertical-align:top;width:50%;"""" colspan=""""2"""">"""
		objFile.WriteLine "			strHTML = strHTML & ""			What...? I like um... grey space!!<p>"""
		objFile.WriteLine "			strHTML = strHTML & ""			I am sure this will be filled with more Setup loveliness at a later date!"""
		objFile.WriteLine "			strHTML = strHTML & ""		</td>"""
		objFile.WriteLine "			strHTML = strHTML & ""	</tr>"""
		objFile.WriteLine "			strHTML = strHTML & ""	</table>"""
		objFile.WriteLine "			DataArea.InnerHTML = strHTML"
		objFile.WriteLine "			PopulateOtherData()"
		objFile.WriteLine "		End If"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub PopulateOtherData()"
		objFile.WriteLine "		On Error Resume Next"
		objFile.WriteLine "		booWMIPrompt = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Others\booWMIPrompt"")"
		objFile.WriteLine "		"
		objFile.WriteLine "		If booWMIPrompt = """" OR IsNull(booWMIPrompt) Then "
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Others\booWMIPrompt"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			booWMIPrompt = ""1"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		"
		objFile.WriteLine "		If booWMIPrompt = ""1"" Then"
		objFile.WriteLine "			cbxWMIPrompt.Checked = True"
		objFile.WriteLine "		End If"
		objFile.WriteLine "		booResolveHostNames = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Others\booResolveHostNames"")"
		objFile.WriteLine "		"
		objFile.WriteLine "		If booResolveHostNames = """" OR IsNull(booResolveHostNames) Then "
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Others\booResolveHostNames"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			booResolveHostNames = ""1"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		"
		objFile.WriteLine "		If booResolveHostNames = ""1"" Then"
		objFile.WriteLine "			cbxResolveHostNames.Checked = True"
		objFile.WriteLine "		End If"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub ApplyOtherAction()"
		objFile.WriteLine "		On Error Resume Next"
		objFile.WriteLine "		booWMIPrompt = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Others\booWMIPrompt"")"
		objFile.WriteLine "		If cbxWMIPrompt.Checked AND booWMIPrompt = ""0"" Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Others\booWMIPrompt"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			ElseIf cbxWMIPrompt.Checked = False AND booWMIPrompt = ""1"" Then"
		objFile.WriteLine "				objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Others\booWMIPrompt"", _"
		objFile.WriteLine "				""0"", ""REG_SZ"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "		booResolveHostNames = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Others\booResolveHostNames"")"
		objFile.WriteLine "		If cbxResolveHostNames.Checked AND booResolveHostNames = ""0"" Then"
		objFile.WriteLine "			objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Others\booResolveHostNames"", _"
		objFile.WriteLine "			""1"", ""REG_SZ"""
		objFile.WriteLine "			ElseIf cbxResolveHostNames.Checked = False AND booResolveHostNames = ""1"" Then"
		objFile.WriteLine "				objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Others\booResolveHostNames"", _"
		objFile.WriteLine "				""0"", ""REG_SZ"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub SetDefaults()"
		objFile.WriteLine "		On Error Resume Next"
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strQueryChoices"", _"
		objFile.WriteLine "		""1,2,3,4,5,"", ""REG_SZ"""
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\intSearchView"", _"
		objFile.WriteLine "		""1"", ""REG_SZ"""
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strInvDirectory"", _"
		objFile.WriteLine "		""C:\SKB\Inventory"", ""REG_SZ"""
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\booDeleteTemp"", _"
		objFile.WriteLine "		""0"", ""REG_SZ"""
		objFile.WriteLine "		Set objRootDSE = GetObject(""LDAP://RootDSE"")"
		objFile.WriteLine "		strDNSDomain = objRootDSE.Get(""defaultNamingContext"")"
		objFile.WriteLine "		strRootOU = strDNSDomain"
		objFile.WriteLine "		objShell.RegWrite ""HKCU\Software\SKB\PCManagementUtil\Setup\strRootOU"", strRootOU, ""REG_SZ"""
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\booRunAs"", _"
		objFile.WriteLine "		""0"", ""REG_SZ"""
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strRunAs"", _"
		objFile.WriteLine "		"""", ""REG_SZ"""
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strRunAsUser"", _"
		objFile.WriteLine "		"""", ""REG_SZ"""
		objFile.WriteLine "		"
		objFile.WriteLine "		If tab1.style.fontweight = ""bold"" Then "
		objFile.WriteLine "			PopulateGeneralData()"
		objFile.WriteLine "		End If"
		objFile.WriteLine "		"
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSoftwareVersion"", _"
		objFile.WriteLine "		""1"", ""REG_SZ"""
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSoftwareVendor"", _"
		objFile.WriteLine "		""1"", ""REG_SZ"""
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSoftwareInstallDate"", _"
		objFile.WriteLine "		""1"", ""REG_SZ"""
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortSoftware"", _"
		objFile.WriteLine "		""1"", ""REG_SZ"""
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booShowSoftware"", _"
		objFile.WriteLine "		""0"", ""REG_SZ"""
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortProcess"", _"
		objFile.WriteLine "		""1"", ""REG_SZ"""
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booSortStartup"", _"
		objFile.WriteLine "		""1"", ""REG_SZ"""
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booUseInstallMonitor"", _"
		objFile.WriteLine "		""1"", ""REG_SZ"""
		objFile.WriteLine "		"
		objFile.WriteLine "		If tab2.style.fontweight = ""bold"" Then "
		objFile.WriteLine "			PopulateTabData()"
		objFile.WriteLine "		End If"
		objFile.WriteLine "		"
		objFile.WriteLine "		strInvDirectory = ""C:\SKB\Inventory"""
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\strInvDirectory"", _"
		objFile.WriteLine "		strInvDirectory, ""REG_SZ"""
		objFile.WriteLine "		objShell.RegDelete ""HKEY_USERS\"" & strSID & _"
		objFile.WriteLine "		""\Software\SKB\PCManagementUtil\Setup\strCurrentProfiles"""
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\booQueryAD"", _"
		objFile.WriteLine "		""1"", ""REG_SZ"""
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Tabs\intProfAge"", _"
		objFile.WriteLine "		""90"", ""REG_SZ"""
		objFile.WriteLine "		"
		objFile.WriteLine "		If tab3.style.fontweight = ""bold"" Then"
		objFile.WriteLine "			PopulateActionsData()"
		objFile.WriteLine "			For Each objOption in ProtectedProfiles.Options"
		objFile.WriteLine "				objOption.RemoveNode"
		objFile.WriteLine "			Next"
		objFile.WriteLine "			Set objOption = Document.createElement(""OPTION"")"
		objFile.WriteLine "			objOption.Text = """""
		objFile.WriteLine "			ProtectedProfiles.Add(objOption)"
		objFile.WriteLine "		End If"
		objFile.WriteLine "		"
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Others\booWMIPrompt"", _"
		objFile.WriteLine "		""1"", ""REG_SZ"""
		objFile.WriteLine "		objShell.RegWrite ""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Others\booResolveHostNames"", _"
		objFile.WriteLine "		""1"", ""REG_SZ"""
		objFile.WriteLine "		"
		objFile.WriteLine "		If tab4.style.fontweight = ""bold"" Then"
		objFile.WriteLine "			PopulateOtherData()"
		objFile.WriteLine "		End If"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub tabMouseOver(strTab)"
		objFile.WriteLine "		If strTab.style.fontweight <> ""bold"" Then "
		objFile.WriteLine "			strTab.style.background = ""#ddf"""
		objFile.WriteLine "			Else"
		objFile.WriteLine "				strTab.style.cursor = ""pointer"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub tabMouseOut(strTab)"
		objFile.WriteLine "		If strTab.style.fontweight <> ""bold"" Then "
		objFile.WriteLine "			strTab.style.background = ""#EEEEEE"""
		objFile.WriteLine "			Else "
		objFile.WriteLine "				strTab.style.background = ""#dddddd"""
		objFile.WriteLine "		End If"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub CloseSetup()"
		objFile.WriteLine "		arrCommands = Split(objSetup.commandLine, ""*"")"
		objFile.WriteLine "		strScriptLocation = arrCommands(1)"
		objFile.WriteLine "		strScriptLocation = Replace(strScriptLocation, ""file:///"", """")"
		objFile.WriteLine "		strScriptLocation = Replace(strScriptLocation, ""%20"", "" "")"
		objFile.WriteLine "		strScriptLocation = Replace(strScriptLocation, ""/"", ""\"")"
		objFile.WriteLine "		booRunAs = objShell.RegRead(""HKEY_USERS\"" & strSID & ""\Software\SKB\PCManagementUtil\Setup\booRunAs"")"
		objFile.WriteLine "		If booRunAs = 1 Then objShell.Run(strScriptLocation)"
		objFile.WriteLine "		Window.Close"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Sub Window_onLoad()"
		objFile.WriteLine "		self.ResizeTo 546,350"
		objFile.WriteLine "		On Error Resume Next"
		objFile.WriteLine "		tab1Click()"
		objFile.WriteLine "		window.moveTo intLeft,intTop"
		objFile.WriteLine "	End Sub"
		objFile.WriteLine ""
		objFile.WriteLine "	Function ChangeColours(intTab)"
		objFile.WriteLine "		document.getElementById(""tab"" & intTab).style.fontweight = ""bold"""
		objFile.WriteLine "		document.getElementById(""tab"" & intTab).style.backgroundcolor = ""#dddddd"""
		objFile.WriteLine "		document.getElementById(""tab"" & intTab).style.bordercolor = ""black"""
		objFile.WriteLine "		document.getElementById(""tab"" & intTab).style.borderbottomwidth = ""0"""
		objFile.WriteLine "		"
		objFile.WriteLine "		For i = 1 to 4"
		objFile.WriteLine "			If i <> intTab Then"
		objFile.WriteLine "				document.getElementById(""tab"" & i).style.fontweight = ""normal"""
		objFile.WriteLine "				document.getElementById(""tab"" & i).style.backgroundcolor = ""#eeeeee"""
		objFile.WriteLine "				document.getElementById(""tab"" & i).style.bordercolor = """""
		objFile.WriteLine "				document.getElementById(""tab"" & i).style.borderbottomcolor = ""black"""
		objFile.WriteLine "				document.getElementById(""tab"" & i).style.borderbottomwidth = """""
		objFile.WriteLine "			End If"
		objFile.WriteLine "		Next"
		objFile.WriteLine "	End Function"
		objFile.WriteLine ""
		objFile.WriteLine "	Function GetSIDFromUser(UserName)"
		objFile.WriteLine "		If InStr(UserName, ""\"") > 0 Then"
		objFile.WriteLine "			DomainName = Mid(UserName, 1, InStr(UserName, ""\"") - 1)"
		objFile.WriteLine "			UserName = Mid(UserName, InStr(UserName, ""\"") + 1)"
		objFile.WriteLine "			Else"
		objFile.WriteLine "				DomainName = CreateObject(""WScript.Network"").UserDomain"
		objFile.WriteLine "		End If"
		objFile.WriteLine ""
		objFile.WriteLine "		On Error Resume Next"
		objFile.WriteLine "		Set WMIUser = GetObject(""winmgmts:{impersonationlevel=impersonate}!"" _"
		objFile.WriteLine "		& ""/root/cimv2:Win32_UserAccount.Domain='"" & DomainName & ""'"" _"
		objFile.WriteLine "		& "",Name='"" & UserName & ""'"")"
		objFile.WriteLine "		If Err = 0 Then Result = WMIUser.SID Else Result = """""
		objFile.WriteLine "		On Error GoTo 0"
		objFile.WriteLine ""
		objFile.WriteLine "		GetSIDFromUser = Result"
		objFile.WriteLine "	End Function"
		objFile.WriteLine ""
		objFile.WriteLine "	Function PickFolder(strStartDir)"
		objFile.WriteLine "		Set objShellApp = CreateObject(""Shell.Application"")"
		objFile.WriteLine "		Set objFolder = objShellApp.BrowseForFolder(0, ""Choose a folder"", 0, strStartDir)"
		objFile.WriteLine "		If (NOT objFolder Is Nothing) Then"
		objFile.WriteLine "		  PickFolder = objFolder.Items.Item.path"
		objFile.WriteLine "		End If"
		objFile.WriteLine "	End Function	"
		objFile.WriteLine ""
		objFile.WriteLine "<" & Chr(47) & "Script>"
		objFile.WriteLine ""
		objFile.WriteLine "<body style=""background-color:#dddddd"">"
		objFile.WriteLine ""
		objFile.WriteLine "<span id=""header"">"
		objFile.WriteLine "	<ul>"
		objFile.WriteLine "		<li id=""tab1"" style=""background-color:#eeeeee"" onClick=""tab1Click()"" onMouseOver=""tabMouseOver(tab1)"" onMouseOut=""tabMouseOut(tab1)"">GENERAL</li>"
		objFile.WriteLine "		<li id=""tab2"" style=""background-color:#eeeeee"" onClick=""tab2Click()"" onMouseOver=""tabMouseOver(tab2)"" onMouseOut=""tabMouseOut(tab2)"">TAB OPTIONS</a></li>"
		objFile.WriteLine "		<li id=""tab3"" style=""background-color:#eeeeee"" onClick=""tab3Click()"" onMouseOver=""tabMouseOver(tab3)"" onMouseOut=""tabMouseOut(tab3)"">ACTIONS</a></li>"
		objFile.WriteLine "		<li id=""tab4"" style=""background-color:#eeeeee"" onClick=""tab4Click()"" onMouseOver=""tabMouseOver(tab4)"" onMouseOut=""tabMouseOut(tab4)"">OTHER</a></li>"
		objFile.WriteLine "	</ul>"
		objFile.WriteLine "</span>"
		objFile.WriteLine ""
		objFile.WriteLine "<span id=""DataArea"">"
		objFile.WriteLine "</span>"
		objFile.WriteLine ""
		objFile.WriteLine "<span style=""position:absolute;right:20px;bottom:20px;"">"
		objFile.WriteLine "	<input type=""button"" value=""Set Defaults"" title=""Set all values to default"" onClick=""SetDefaults()"">"
		objFile.WriteLine "	<input type=""button"" value=""Close"" title=""Close"" onClick=""Window.Close"">"
		objFile.WriteLine "</span>"
		objFile.WriteLine ""
		objFile.WriteLine "<script language=""VBScript"">"
		objFile.WriteLine "	Set objWMIService = GetObject(""winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2"")"
		objFile.WriteLine ""
		objFile.WriteLine "	Set colComputer = objWMIService.ExecQuery _"
		objFile.WriteLine "		(""Select * from Win32_ComputerSystem"")"
		objFile.WriteLine ""
		objFile.WriteLine "	For Each objItem In colComputer"
		objFile.WriteLine "		strLoggedOn = objItem.UserName"
		objFile.WriteLine "	Next"
		objFile.WriteLine "	strSID = GetSIDFromUser(strLoggedOn)"
		objFile.WriteLine "<" & Chr(47) & "Script>"
		objFile.WriteLine ""
		objFile.WriteLine "</body>"
		objFile.WriteLine "</html>"
		objFile.Close
		objShell.Run(Chr(34) & strTempLoc & "\SKB\Setup.hta" & Chr(34) & "*" & document.location.href)
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	About()
    '#	PURPOSE........:	Displays information about application
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub About()
		Msgbox "PC Management Utility" & vbCrLf & "Version: " & objPCManage.Version & vbCrLf & vbCrLf & _
		"This utility was created by Stuart Barrett to manage Local or Remote PCs and has grown into the " & _
		"script you see today but it would not have done so without the wonderful Spiceworks Community, " & _
		"Gods bless ye all!" & vbCrLf & vbCrLf & "This utility is Open Source so you may chop it and change it as " & _
		"you wish, but please credit me as the creator in your amended script!", vbInformation, "About"
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	CleanUp()
    '#	PURPOSE........:	Resets the form to select new PC
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub CleanUp()
		DataArea.InnerHTML = ""
		strPC = ""
		strPCInfoTab = ""
		PCSearch.Disabled = False
		
		btnSearch.Disabled = False
		btnSearch.Title = "Search PC"
		
		btnSelectPC.Disabled = False
		btnSelectPC.Title = "Select highlighted PC"
		
		btnRescan.Disabled = True
		btnRescan.style.cursor = "default"
		btnRescan.Title = ""
		
		PCSearch.Value = ""
		PCSearch.style.backgroundcolor = "white"
		AvailablePCs.Disabled = False
		AvailablePCs.style.backgroundcolor = "white"
		tab1.Disabled = True
		tab1.Title = ""
		tab2.Disabled = True
		tab2.Title = ""
		tab3.Disabled = True
		tab3.Title = ""
		tab4.Disabled = True
		tab4.Title = ""
		tab5.Disabled = True
		tab5.Title = ""
		tab6.Disabled = True
		tab6.Title = ""
		tab1.style.cursor = "default"
		tab2.style.cursor = "default"
		tab3.style.cursor = "default"
		tab4.style.cursor = "default"
		tab5.style.cursor = "default"
		tab6.style.cursor = "default"
		tab1.bgcolor="#eeeeee"
		tab2.bgcolor="#eeeeee"
		tab3.bgcolor="#eeeeee"
		tab4.bgcolor="#eeeeee"
		tab5.bgcolor="#eeeeee"
		tab6.bgcolor="#eeeeee"
		PCSearch.focus()
		If SearchView.Value = 2 Then
			AvailableOUs.Disabled = False
			
			btnSelectOU.Disabled = False
			btnSelectOU.Title = "Select highlighted PC"
			
			AvailableOUs.Style.backgroundcolor = "white"
			ClearPCs
			Else
				IP1A.Disabled = False
				IP1A.Style.backgroundcolor = "white"
				IP2A.Disabled = False
				IP2A.Style.backgroundcolor = "white"
				IP3A.Disabled = False
				IP3A.Style.backgroundcolor = "white"
				IP4A.Disabled = False
				IP4A.Style.backgroundcolor = "white"
				IP1B.Disabled = False
				IP1B.Style.backgroundcolor = "white"
				IP2B.Disabled = False
				IP2B.Style.backgroundcolor = "white"
				IP3B.Disabled = False
				IP3B.Style.backgroundcolor = "white"
				IP4B.Disabled = False
				IP4B.Style.backgroundcolor = "white"
				IP1A.Value = ""
				IP2A.Value = ""
				IP3A.Value = ""
				IP4A.Value = ""
				IP1B.Value = ""
				IP2B.Value = ""
				IP3B.Value = ""
				IP4B.Value = ""
				btnSearchRange.Disabled = False
				btnSearchRange.Title = "Search for IPs in IP range"
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ExpInventory()
    '#	PURPOSE........:	Exports an inventory of the PC to a HTML file
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ExpInventory()
		On Error Resume Next
		
		Const SEARCH_KEY = "DigitalProductID"
		
		strInvDirectory = objShell.RegRead(strRegStart & "\Software\SKB\PCManagementUtil\Setup\strInvDirectory")
		
		If strInvDirectory = "" Then
			strInvDirectory = "C:\SKB\Inventory"
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\Setup\strInvDirectory", _
			strInvDirectory, "REG_SZ"
		End If
		
		If strInvDirectory = "C:\SKB\Inventory" Then
			If NOT objFSO.FolderExists("c:\SKB") Then
				objFSO.CreateFolder("c:\SKB")
			End If
		End If
		
		strInvDirectory2 = Replace(strInvDirectory, "C:", "file:///c:")
		strInvDirectory2 = Replace(strInvDirectory2, "\", "/")
	
		intProgTotal = 17
		intProgMult = 100 / intProgTotal
		intProgDone = 0
		intProfileCount = 0
		intColumnIndex = 0
		window.clearTimeout(idTimer)
		
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\cimv2") 
		
		Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\default:StdRegProv") 
		
		intProgDone = intProgDone + 1	'1
		UpdateProgressBar intProgMult,intProgDone,intProgTotal
		
		Set colComputer = objWMIService.ExecQuery _
			("Select * from Win32_ComputerSystem")
		
		For Each objItem In colComputer
			strManufacturer = objItem.Manufacturer
			strModel = objItem.Model
			strDomainRole = objItem.DomainRole
			intMemSize = round(objItem.TotalPhysicalMemory / 1073741824,2)	
		Next
		
		intProgDone = intProgDone + 1	'2
		UpdateProgressBar intProgMult,intProgDone,intProgTotal
		
		Set objWMIService2 = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2") 
		
		Set colNAC = objWMIService.ExecQuery _
			("Select * from Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
		
		intIpCount = 0
			
		For Each objItem in colNAC
			If IsNull(objItem.IPAddress) Then
				strIP = ""
				Else
					strIP = objItem.IPAddress(0)
			End If
			
			If strIP <> "0.0.0.0" AND strIP <> "" Then
				If intIpCount > 0  Then
					strResolveIP = strResolveIP & "; " & strIP
					If objItem.DHCPEnabled = "True" Then
						strResolveIP = strResolveIP & " (DHCP enabled)"
						Else
							strResolveIP = strResolveIP & " (Static IP)"
					End If
					Else
						strResolveIP = strIP
						If objItem.DHCPEnabled = "True" Then
							strResolveIP = strResolveIP & " (DHCP enabled)"
							Else
								strResolveIP = strResolveIP & " (Static IP)"
						End If
				End If
				intIpCount = intIpCount + 1
			End If
			strIP = ""
		Next
		
		intProgDone = intProgDone + 1	'3
		UpdateProgressBar intProgMult,intProgDone,intProgTotal
		
		Set colOS = objWMIService.ExecQuery _
			("Select * from Win32_OperatingSystem")
		
		For Each objItem In colOS
			strOS = objItem.Caption
			intServicePackMajor = objItem.ServicePackMajorVersion
			dtmBootDate = FormatDate(objItem.LastBootUpTime)
			strUptime = TimeSpan(dtmBootDate,Now)
			strDescription = objItem.Description
			strInstallDate = FormatDate(objItem.InstallDate)
			strBuildNumber = objItem.BuildNumber
			strPID = objItem.SerialNumber
		Next
					
		strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
		objReg.GetBinaryValue HKEY_LOCAL_MACHINE, strKeyPath, _
		SEARCH_KEY, arrDPIDBytes
		If Not IsNull(arrDPIDBytes) Then
			strOSKey = DecodeKey(arrDPIDBytes)
		End If
		
		intProgDone = intProgDone + 1	'4
		UpdateProgressBar intProgMult,intProgDone,intProgTotal

		strArchitecture = CheckWinArchitecture()

		Set colBIOS = objWMIService.ExecQuery _
			("Select * from Win32_BIOS")
		
		For Each objItem In colBIOS
			strSerial = objItem.SerialNumber
		Next
		
		intProgDone = intProgDone + 1	'5
		UpdateProgressBar intProgMult,intProgDone,intProgTotal
		
		strKeyPath = strRemoteSID & "\SessionInformation"
		objReg.GetDWORDValue HKEY_USERS,strKeyPath,"ProgramCount",intProgCount
		
		strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
		objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys
		For Each objItem In arrSubkeys
			strValueName = "ProfileImagePath"
			strSubPath = strKeyPath & "\" & objItem
			objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath,strValueName,strValue
			
			If strValue <> "" Then
				arrPath = Split(strValue, "\")
				strProfName = arrPath(Ubound(arrPath))
				strProfileInfo = strProfileInfo & "<tr><td><b>" & strProfName & ":&nbsp;&nbsp;</b></td><td>" & _
				strValue & "</td></tr>"
				intProfileCount = intProfileCount + 1
			End If
		Next

		If intProfileCount > 1 Then
			strProfileFooter = "<span style=""font-size: 0.9em;""><b>" & _
			intProfileCount & " Profiles</b></span>"
			Else
				strProfileFooter = "<span style=""font-size: 0.9em;""><b>" & intProfileCount & " Profile</b></span>"
		End If
		intProgDone = intProgDone + 1	'6
		UpdateProgressBar intProgMult,intProgDone,intProgTotal
		
		Set colShares = objWMIService.ExecQuery("Select * from Win32_Share Where Type='2147483648' OR Type='0'")
		
		For Each objItem In colShares
			strShareName = objItem.Name
			strShareCaption = objItem.Caption
			strSharePath = objItem.Path
			strShareInfo = strShareInfo & "<tr><td>" & strShareName & "</td><td>" & _
			strShareCaption & "</td><td>" & strSharePath & "</td></tr>"
		Next
		
		intProgDone = intProgDone + 1	'7
		UpdateProgressBar intProgMult,intProgDone,intProgTotal
		
		Set colUserAccounts = objWMIService.ExecQuery _
			("Select * from Win32_UserAccount Where LocalAccount = True")
 
		For Each objItem in colUserAccounts
			strUserAccName = objItem.Name
			strUserAccDescription =  objItem.Description
			If IsNull (strUserAccDescription) OR strUserAccDescription = "" Then
				strUserAccDescription = "&nbsp;"
			End If
			booUserAccDisabled = objItem.Disabled
			If booUserAccDisabled = True Then
				strUserAccDisabled = "Yes"
				Else
					strUserAccDisabled = "No"
			End If
			strUserAccountInfo = strUserAccountInfo & "<tr><td>" & strUserAccName & _
			"</td><td>" & strUserAccDescription & _
			"</td><td>" & strUserAccDisabled & "</td></tr>"
		Next
		
		intProgDone = intProgDone + 1	'8
		UpdateProgressBar intProgMult,intProgDone,intProgTotal
		
		Set colUserGroups = objWMIService.ExecQuery _
			("SELECT * FROM Win32_Group Where LocalAccount = True")
		
		For Each objItem In colUserGroups
			strGroup = objItem.Name
			Set objGroup = GetObject("WinNT://" & strPC & "/" & strGroup)
			For Each Member In objGroup.Members
				strGroupMember = Member.Name
				strUserGroupInfo = strUserGroupInfo & "<tr><td>" & strGroup & _
				"</td><td>" & strGroupMember & "</td></tr>"
			Next
		Next
		
		intProgDone = intProgDone + 1	'9
		UpdateProgressBar intProgMult,intProgDone,intProgTotal
		
		If strDomainRole < 2 Then
			Set objWMIService2 = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
			strPC & "\root\securitycenter")
			
			Set colAV = objWMIService2.ExecQuery _
				("Select * from AntiVirusProduct")
			
			For Each objItem in colAV
				strAVName = objItem.displayName
				booAVUpToDate = objItem.productUptoDate
				strAVVersion = objItem.versionNumber
				If booAVUpToDate = True Then
					strAVUpToDate = "Yes"
					Else
						strAVUpToDate = "No"
				End If
				strAVInfo = strAVInfo & "<tr><td>" & strAVName & "</td>" & _
				"<td>" & strAVUpToDate & "</td>" & _
				"<td>" & strAVVersion & "</td>"
			Next
		End If

		intProgDone = intProgDone + 1	'10
		UpdateProgressBar intProgMult,intProgDone,intProgTotal
		
		Set colDisk = objWMIService.ExecQuery _
			("Select * from Win32_LogicalDisk")
		
		For Each objItem In colDisk
			strDriveType = objItem.DriveType
			If strDriveType <> 0 AND strDriveType <> 4 AND strDriveType <> 5 Then
				intSize = objItem.Size
				intFreeSpace = objItem.FreeSpace
				strDiskName = objItem.Name
				strVolumeName = objItem.VolumeName
				strFileSystem = objItem.FileSystem
				strVolumeSerialNumber = objItem.VolumeSerialNumber
				strDiskDescription = objItem.Description
			
				If IsNumeric(intFreeSpace) = False Then
					intFreeSpace=0
				End If
				If IsNumeric(intSize) = False Then
					intSize = 0
				End If
				If objItem.FreeSpace > 0 Then
					pctFreeSpace = round(((intFreeSpace / intSize) * 100),0)
				Else
					pctFreeSpace=0
				End If
				strDiskSize = ConvertToDiskSize(intSize) 
				strFreeSpace = ConvertToDiskSize(intFreeSpace)
				strUsedSpace = ConvertToDiskSize(intSize-intFreeSpace)
			
				strDiskInfo = strDiskInfo & "<tr><td>" & strDiskName & "</td><td>" & _
				strFileSystem & "</td></td><td>" & strDiskDescription & "</td></td><td>" & _
				strVolumeSerialNumber & "</td></td><td>" & strDiskSize & "</td>" & _
				"</td><td>" & strFreeSpace & "</td>" & _
				"</td><td>" & pctFreeSpace & "</td>"
			End If
		Next
		
		intProgDone = intProgDone + 1	'11
		UpdateProgressBar intProgMult,intProgDone,intProgTotal
		
		Set colPrinter = objWMIService.ExecQuery _
			("Select * from Win32_Printer")
		
		For Each objItem In colPrinter
			strPrinter = objItem.Caption
			booLocal = objItem.Local
			strPort = objItem.PortName
			booShared = objItem.Shared
			strShareName = objItem.ShareName
			
			If booLocal = True Then
				strLocal = "Local"
				Else
					strLocal = "Network"
			End If
			
			If IsNull(strShareName) OR strShareName = "" Then
				strShareName = "None"
			End If

			strPrinters = strPrinters & "<tr><td>" & strPrinter & _
			"</td><td>" & strPort & "</td><td>" & strLocal & "</td><td>" & _
			booShared & "</td><td>" & strShareName & "</td></tr>"
		Next
		
		intProgDone = intProgDone + 1	'12
		UpdateProgressBar intProgMult,intProgDone,intProgTotal
		
		Set colProcesses = objWMIService.ExecQuery _
			("Select * from Win32_Process")
		
		For Each objItem In colProcesses
			strProcess = objItem.Caption
			intMemUsage = objItem.WorkingSetSize
			strProcessID = objItem.ProcessID
			strGoogle = "<a target=_blank href=""http://www.google.com/search?q=" & _
			strProcess & """>Search</a>"
			
			If IsNull(intMemUsage) OR intMemUsage = "" Then
				strMemUsage = "0 MB"
					Else
						strMemUsage = ConvertToDiskSize(intMemUsage)
			End If
			
			strProcesses = strProcesses & "<tr><td>" & strProcess & _
			"</td><td>" & strProcessID & _
			"</td><td>" & strMemUsage & "</td>" & _
			"<td>" & strGoogle & "</td></tr>"
		Next
		
		intProgDone = intProgDone + 1	'13
		UpdateProgressBar intProgMult,intProgDone,intProgTotal
		
		strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
		objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys

		For Each objItem In arrSubkeys
			strValueName = "DisplayName"
			strSubPath = strKeyPath & "\" & objItem
			objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath,strValueName,strValue
		
			If strValue <> "" AND InStr(strValue, "Hotfix") = 0 AND _
			InStr(strValue, "Security Update") = 0 AND _
			InStr(strValue, "Update for Windows") = 0 Then
				strSoftwareName = strValue
				strGoogleSW = "<a target=_blank href=""http://www.google.com/search?q=" & _
				strSoftwareName & """>" & strSoftwareName & "</a>"
				objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath, _
				"DisplayVersion",strSoftwareVersion
				objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath, _
				"InstallDate",intSoftwareInstallDate
				objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath, _
				"Publisher",strSoftwareVendor
				If IsNull(intSoftwareInstallDate) OR intSoftwareInstallDate = "" Then
					strSoftwareInstallDate = "&nbsp;"
					Else 
						strSoftwareInstallDate = MID(intSoftwareInstallDate,7,2) & _
						"/" & MID(intSoftwareInstallDate,5,2) & "/" & _
						LEFT(intSoftwareInstallDate,4)
						If NOT IsDate(strSoftwareInstallDate) Then
							strSoftwareInstallDate = "&nbsp;"
						End If
				End If
				If IsNull(strSoftwareName) OR strSoftwareName = "" Then
					strSoftwareName = "&nbsp;"
				End If
				If IsNull(strSoftwareVendor) OR strSoftwareVendor = "" Then
					strSoftwareVendor = "&nbsp;"
				End If
				If IsNull(strSoftwareVersion) OR strSoftwareVersion = "" Then
					strSoftwareVersion = "&nbsp;"
				End If
				strSoftwareInfo = strSoftwareInfo & "<tr><td>" & strGoogleSW & _
				"</td><td>" & strSoftwareVendor & _
				"</td><td>" & strSoftwareVersion & "</td>" & _
				"<td>" & strSoftwareInstallDate & "</td></tr>"
			End If
		Next
		
		intProgDone = intProgDone + 1	'14
		UpdateProgressBar intProgMult,intProgDone,intProgTotal
		
		Set colStartup = objWMIService.ExecQuery _
			("Select * from Win32_StartupCommand")
		
		For Each objItem In colStartup
			strStartupName = objItem.Caption
			strStartupCommand = objItem.Command
			strStartupUser = objItem.User
			strStartupLocation = objItem.Location
			
			If strStartupLocation = "Common Startup" Then
				If objFSO.FolderExists("C:\Documents and Settings\All Users") Then
					strStartupLocation = _
					"C:\Documents and Settings\All Users\Start Menu\Programs\Startup\"
					ElseIf objFSO.FolderExists("C:\ProgramData") Then
						strStartupLocation = "C:\ProgramData\Start Menu\Programs\Startup\"
				End If
			End If

			If LCase(strStartupCommand) <> "desktop.ini" AND LCase(strStartupUser) <> ".default" _
			AND strStartupName <> "" AND InStr(LCase(strStartupUser), "nt authority") = 0 _
			AND LCase(strStartupName) <> LCase("ctfmon.exe") AND LCase(strStartupName) <> LCase("desktop") Then
				strStartupInfo = strStartupInfo & "<tr><td>" & strStartupName & "</td>" & _
				"<td>" & strStartupCommand & "</td><td>" & strStartupUser & "</td>" & _
				"<td>" & strStartupLocation & "</td>"
			End If
		Next
		
		intProgDone = intProgDone + 1	'15
		UpdateProgressBar intProgMult,intProgDone,intProgTotal
		
		strBackToTop = "<p class=""backtotop""><a href=""" & strInvDirectory2 & _
		"/" & strPC & "_" & strFileDate & ".html#top"">[..back to top..]</a></p>"
		
		strInventoryInfo = "<u><h2><a name =""top"">Inventory for " & _
		strPC & "</a></h2></u><p><br><h3>Main Info</h3>" & _
		"<table cellpadding=""1"">" & _
			"<tr><td><b>Computer Name:&nbsp;&nbsp;</b></td><td>" & strPC & "</td></tr>" & _
			"<tr><td><b>IP Address(es):&nbsp;&nbsp;</b></td><td>" & strResolveIP & "</td></tr>" & _
			"<tr><td><b>Logged On User:&nbsp;&nbsp;</b></td><td>" & strRemoteLoggedOn & "</td></tr>" & _
			"<tr><td><b>Running Programs:&nbsp;&nbsp;</b></td><td>" & intProgCount & "</td></tr>" & _
			"<tr><td>&nbsp;</td><td>&nbsp;</td></tr>" & _
			"<tr><td><b>Description:&nbsp;&nbsp;</b></td><td>" & strDescription & "</td></tr>" & _
			"<tr><td><b>Manufacturer:&nbsp;&nbsp;</b></td><td>" & strManufacturer & "</td></tr>" & _
			"<tr><td><b>Model:&nbsp;&nbsp;</b></td><td>" & strModel & "</td></tr>" & _
			"<tr><td><b>Serial Number:&nbsp;&nbsp;</b></td><td>" & strSerial & "</td></tr>" & _
			"<tr><td><b>RAM:&nbsp;&nbsp;</b></td><td>" & intMemSize & " GB</td></tr>" & _
			"<tr><td>&nbsp;</td><td>&nbsp;</td></tr>" & _
			"<tr><td><b>OS:&nbsp;&nbsp;</b></td><td>" & strOS & " (Build: " & strBuildNumber & ")</td></tr>" & _
			"<tr><td><b>Product ID:&nbsp;&nbsp;</b></td><td>" & strPID & "</td></tr>" & _
			"<tr><td><b>Product Key:&nbsp;&nbsp;</b></td><td>" & strOSKey & "</td></tr>" & _
			"<tr><td>&nbsp;</td><td>&nbsp;</td></tr>" & _
			"<tr><td><b>Architecture:&nbsp;&nbsp;</b></td><td>" & strArchitecture & "</td></tr>" & _
			"<tr><td><b>Service Pack:&nbsp;&nbsp;</b></td><td>" & intServicePackMajor & "</td></tr>" & _
			"<tr><td><b>Install Date:&nbsp;&nbsp;</b></td><td>" & strInstallDate & "</td></tr>" & _
			"<tr><td>&nbsp;</td><td>&nbsp;</td></tr>" & _
			"<tr><td><b>Last Reboot:&nbsp;&nbsp;</b></td><td>" & dtmBootDate & "</td></tr>" & _
			"<tr><td><b>System Uptime:&nbsp;&nbsp;</b></td><td>" & strUptime & "</td></tr>" & _
		"</table>" & strBackToTop & _
		"<p><br><h3>Profile Paths</h3>" & _
		"<table cellpadding=""1"">" & _
			strProfileInfo & _
		"</table><p>" & strProfileFooter & strBackToTop & _
		"<p><br><h3>Local User Accounts</h3>" & _
		"<table class=""htmltable"">" & _
			"<tr><th>Username</th><th>Description</th><th>Disabled</th>" & _ 
			strUserAccountInfo & _
		"</table>" & strBackToTop & _
		"<p><br><h3>Local User Groups</h3>" & _
		"<table class=""htmltable"">" & _
			"<tr><th>Group</th><th>Username</th>" & _
			strUserGroupInfo & _
		"</table>" & strBackToTop & _
		"<p><br><h3>Antivirus Info</h3>" & _
		"<table class=""htmltable"">" & _
			"<tr><th>Antivirus Name</th><th>Up To Date</th><th>Version</th>" & _
			strAVInfo & _
		"</table>" & strBackToTop & _
		"<p><br><h3>Disk Info</h3>" & _
		"<table class=""htmltable"">" & _
			"<tr><th>Disk Drive</th><th>File System</th><th>Description</th>" & _
			"<th>Serial Number</th><th>Disk Size</th><th>Free Space</th><th>% Free</th></tr>" & _
			strDiskInfo & _
		"</table>" & strBackToTop & _
		"<p><br><h3>Remote Shares</h3>" & _
		"<table class=""htmltable"">" & _
			"<tr><th>Share Name</th><th>Description</th><th>Path</th>" & strShareInfo & _
		"</table>" & strBackToTop & _
		"<p><br><h3>Printers</h3>" & _
		"<table class=""htmltable"">" & _
			"<tr><th>Printer</th><th>Port</th><th>Local?</th><th>Shared?</th>" & _
			"<th>Share Name</th></tr>" & _
			strPrinters & _
		"</table>" & strBackToTop & _
		"<p><br><h3>Running Processes</h3>" & _
		"<table class=""htmltable"">" & _
			"<tr><th>Process Name</th><th>Process ID</th><th>Mem Usage</th>" & _
			"<th>Google Search</th></tr>" & _
			strProcesses & _
		"</table>" & strBackToTop & _
		"<p><br><h3>Installed Software</h3>" & _
		"<table class=""htmltable"">" & _
			"<tr><th>Software Name</th><th>Vendor</th><th>Version</th><th>Install Date</th></tr>" & _
			strSoftwareInfo & _
		"</table>" & strBackToTop & _
		"<p><br><h3>Startup Items</h3>" & _
		"<table class=""htmltable"">" & _
			"<tr><th>Startup Item</th><th>Command</th><th>User</th><th>Startup Location</th></tr>" & _
			strStartupInfo & _
		"</table>" & strBackToTop
		
		intProgDone = intProgDone + 1	'16
		UpdateProgressBar intProgMult,intProgDone,intProgTotal

		If NOT objFSO.FolderExists(strInvDirectory) Then
			Set objFolder = objFSO.CreateFolder(strInvDirectory)
		End If
		
		Set objFile = objFSO.CreateTextFile(strInvDirectory & "\" & _
		strPC & "_" & strFileDate & ".html",True)
		objFile.WriteLine "<html>"
		objFile.WriteLine "<head>"
		objFile.WriteLine "<style type=""text/css"">"
		objFile.WriteLine "body {"
		objFile.WriteLine "	font-family: verdana, arial;"
		objFile.WriteLine "	background-color:#CEF0FF"
		objFile.WriteLine "}"
		objFile.WriteLine "table.htmltable {"
		objFile.WriteLine "	border-width: 2px;"
		objFile.WriteLine "	border-spacing:2px;"
		objFile.WriteLine "	font-size: 0.9em;"
		objFile.WriteLine "	border-style:outset;"
		objFile.WriteLine "	border-color: rgb(0, 102, 204);"
		objFile.WriteLine "	border-collapse: separate;"
		objFile.WriteLine "	background-color: rgb(0, 153, 204);}"
		objFile.WriteLine "table.htmltable th {"
		objFile.WriteLine "	border-width: 1px;"
		objFile.WriteLine "	padding: 3px;"
		objFile.WriteLine "	border-style: inset;"
		objFile.WriteLine "	border-color: rgb(0, 102, 204);"
		objFile.WriteLine "	background-color: #CEF0FF;"
		objFile.WriteLine "}"
		objFile.WriteLine "table.htmltable td {"
		objFile.WriteLine "	border-width: 1px;"
		objFile.WriteLine "	padding: 3px;"
		objFile.WriteLine "	border-style: inset;"
		objFile.WriteLine "	border-color: rgb(0, 102, 204);"
		objFile.WriteLine "	background-color: #CEF0FF;"
		objFile.WriteLine "}"
		objFile.WriteLine ".backtotop A {"
		objFile.WriteLine "	font-size: 0.9em;"
		objFile.WriteLine "}"
		objFile.WriteLine ".def h3 {"
		objFile.WriteLine "	color: rgb(0, 102, 204);"
		objFile.WriteLine "}"
		objFile.WriteLine "</style>"
		objFile.WriteLine "<title>Inventory for " & strPC & "</title>"
		objFile.WriteLine "</head>"
		objFile.WriteLine "<body>"
		objFile.WriteLine strInventoryInfo
		objFile.WriteLine "</body>"
		objFile.WriteLine "</html>" 
		objFile.Close
		
		intProgDone = intProgDone + 1	'17
		UpdateProgressBar intProgMult,intProgDone,intProgTotal

		document.body.style.cursor = "default"

		strOpen = MsgBox("The inventory for " & strPC & " has been exported to " & _
		strInvDirectory & "\" & strPC & "_" & strFileDate & ".html" & _
		vbCrLf & vbCrLf & "Would you like to view the file now?",vbYesNo+vbQuestion,"Exported")
		If strOpen = vbYes Then
			objShell.Run strInvDirectory & "\" & strPC & "_" & strFileDate & ".html"
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	UpdateProgressBar(intProgMult,intProgDone,intProgTotal)
    '#	PURPOSE........:	Determines what action is performed on choosing an 
    '#						item from the Action List
    '#	ARGUMENTS......:	intProgMult = 100 / intProgTotal to determine 
	'#						percentage
	'#						intProgDone = amount of actions performed so far
	'#						intProgTotal = total actions performed
    '#	EXAMPLE........:	UpdateProgressBar intProgMult,intProgDone,intProgTotal
    '#	NOTES..........:	For this Sub to work you must Update the bar more
	'#						than once, and first declare the below variables
	'#						intProgTotal = 16 (or however many actions required)
	'#						intProgMult = 100 / intProgTotal
	'#						intProgDone = 0
    '#--------------------------------------------------------------------------
	Sub UpdateProgressBar(intProgMult,intProgDone,intProgTotal)
		If intProgDone = intProgTotal Then
			WaitMessage.InnerHTML = "<hr><p><div><span style=""padding:5px;margin:0px;width=80%;" & _
			"background-color:blue;text-align='right';font-size:0.9em;color:'white';""></span></div>"
			Else
				pctDone = intProgMult * intProgDone
				WaitMessage.InnerHTML = "<hr><p><div><span style=""padding:5px;margin:0px;width=" & _
				pctDone / 1.25 & "%;background-color:blue;text-align:'right';font-size:0.9em;color:'white';"">" & _
				"</span><span style=""padding:5px;margin:0px;width=" & _
				(100 - pctDone) / 1.25 & "%;background-color:#3366CC;font-size:0.9em;"">&nbsp;</span></div>"
		End If
		PauseScript(100)
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	StopAction()
    '#	PURPOSE........:	Stops the current running action
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub StopAction()
		document.body.style.cursor = "default"
		btnStop.Disabled = True
		btnStop.style.cursor = "default"
		btnStop.title = ""
		intPings = 0
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ChangePCDescription()
    '#	PURPOSE........:	Changes the description of the PC
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------	
	Sub ChangePCDescription()
		On Error Resume Next

		booQueryAD = objShell.RegRead(strRegStart & "\Software\SKB\PCManagementUtil\Tabs\booQueryAD")
		If booQueryAD = "" OR IsNull(booQueryAD) Then 
			objShell.RegWrite strRegStart & "\Software\SKB\PCManagementUtil\Tabs\booQueryAD", _
			"1", "REG_SZ"
			booQueryAD = "1"
		End If
		
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\cimv2") 
		
		Set colOS = objWMIService.ExecQuery _
			("Select * from Win32_OperatingSystem")
			
		For Each objItem In ColOS
			strDescription = Trim(objItem.Description)
		Next 

		Set colComputer = objWMIService.ExecQuery _
			("Select * from Win32_ComputerSystem")
		
		For Each objItem In colComputer
			strManufacturer = Trim(objItem.Manufacturer)
			strModel = Trim(objItem.Model)
		Next
		If InStr(strRemoteLoggedOn, "\") > 0 Then
			arrLoggedOn = Split(strRemoteLoggedOn, "\")
			strLoggedOn = UCase(arrLoggedOn(UBound(arrLoggedOn)))
			Else
				strLoggedOn = strRemoteLoggedOn
		End If
		strArchitecture = CheckWinArchitecture()
		
		If booQueryAD = "1" Then
			Set objRootDSE = GetObject("LDAP://RootDSE")
			
			strRootOU = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\Setup\strRootOU")
			If strRootOU = "" Then
				strDNSDomain = objRootDSE.Get("defaultNamingContext")
				strRootOU = strDNSDomain
				objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\Setup\strRootOU", strRootOU, "REG_SZ"
			End If
			
			Set adoCommand = CreateObject("ADODB.Command")
			Set adoConnection = CreateObject("ADODB.Connection")
			adoConnection.Provider = "ADsDSOObject"
			adoConnection.Open "Active Directory Provider"
			adoCommand.ActiveConnection = adoConnection

			strFilter = "(&(cn=" & strPC & "))"
			strQuery = "<LDAP://" & strRootOU & ">;" & strFilter & _
			";distinguishedName,objectCategory;subtree"
			
			adoCommand.CommandText = strQuery
			adoCommand.Properties("Page Size") = 750
			adoCommand.Properties("Timeout") = 60
			adoCommand.Properties("Cache Results") = False

			Set adoRecordset = adoCommand.Execute

			Do Until adoRecordset.EOF
				strPCDN = adoRecordset.Fields("distinguishedName").Value
				adoRecordset.MoveNext
			Loop
			adoRecordset.Close

			strFilter = "(&(sAMAccountName=" & strLoggedOn & "))"
			strQuery = "<LDAP://" & strRootOU & ">;" & strFilter & _
			";distinguishedName,objectCategory;subtree"
			
			adoCommand.CommandText = strQuery

			Set adoRecordset = adoCommand.Execute

			Do Until adoRecordset.EOF
				strUserDN = adoRecordset.Fields("distinguishedName").Value
				adoRecordset.MoveNext
			Loop
			adoRecordset.Close
			
			Set objPC = GetObject("LDAP://" & strPCDN)
			strADDescription = objPC.Description
			Set objUser = GetObject("LDAP://" & strUserDN)
			strFirstName = objUser.Get("givenName")
			strSurname = objUser.Get("sn")
			
			strPreDefDesc1 = strModel & " - " & strFirstName & " " & strSurname
			strPreDefDesc2 = strModel & " - " & strArchitecture
			strPreDefDesc3 = strManufacturer & " " & strModel
			strPreDefDesc4 = strManufacturer & " " & strModel & " - " & strFirstName & " " & strSurname
			Else
				strADDescription = "<i>Query AD disabled in Setup</i>"
				strPreDefDesc1 = strModel & " - " & strLoggedOn
				strPreDefDesc2 = strModel & " - " & strArchitecture
				strPreDefDesc3 = strManufacturer & " " & strModel
				strPreDefDesc4 = strManufacturer & " " & strModel & " - " & strLoggedOn
		End If
		
		strHTML = "<hr><div style=""text-align:left;"">" 
		strHTML = strHTML & "Please use one of the following methods to change the description:" 
		strHTML = strHTML & "<table width=""100%"">" 
		strHTML = strHTML & 	"<tr>" 
		strHTML = strHTML & 		"<td width=""50%"">" 
		strHTML = strHTML & 			"<b><i>Current Description: " & strDescription & "</i></b><p>" 
		strHTML = strHTML & 			"<input type=""checkbox"" name=""cbxADDesc"" title=""Set the PC description as per the " 
		strHTML = strHTML & 			"Active Directory description"" onclick=""ChangePCDescription2 False"" " 
		strHTML = strHTML & 			"value=""" & strADDescription & """>" 
		strHTML = strHTML & 			"AD Description: " & strADDescription & "<p>" 
		strHTML = strHTML & 			"<input id=""RunButton"" class=""button"" type=""button"" value=""Set Description"" " 
		strHTML = strHTML & 			"name=""btnSetDescription"" onclick=""ChangePCDescription2 True"" " 
		strHTML = strHTML & 			"title=""Set the description"">" 
		strHTML = strHTML & 		"</td>" 
		strHTML = strHTML & 		"<td style=""vertical-align:text-top;"">" 
		strHTML = strHTML & 			"Predefined Descriptions:<br>" 
		strHTML = strHTML & 			"<select size=""4"" name=""DescChooser"" style=""width:100%;"" " 
		strHTML = strHTML & 			"onChange=""ChangePCDescription2 False"">" 
		strHTML = strHTML & 				"<option value=""" & strPreDefDesc1 & """ title=""" & strPreDefDesc1 & """>" 
		strHTML = strHTML & 				strPreDefDesc1 & "</option>" 
		strHTML = strHTML & 				"<option value=""" & strPreDefDesc2 & """ title=""" & strPreDefDesc2 & """>" 
		strHTML = strHTML & 				strPreDefDesc2 & "</option>" 
		strHTML = strHTML & 				"<option value=""" & strPreDefDesc3 & """ title=""" & strPreDefDesc3 & """>" 
		strHTML = strHTML & 				strPreDefDesc3 & "</option>" 
		strHTML = strHTML & 				"<option value=""" & strPreDefDesc4 & """ title=""" & strPreDefDesc4 & """>" 
		strHTML = strHTML & 				strPreDefDesc4 & "</option>" 
		strHTML = strHTML & 			"</select><br>"
		strHTML = strHTML & 			"Custom Description:<br>"
		strHTML = strHTML & 			"<input type=""text"" name=""txtCustomDesc"" size=""30"" " 
		strHTML = strHTML & 			"onKeyUp=""ChangePCDescription2 False"" "
		strHTML = strHTML & 			"title=""Custom Description"">"
		strHTML = strHTML & 			"<input id=""RunButton"" class=""button"" type=""button"" value=""Clear"" "
		strHTML = strHTML & 			"name=""btnClearDescription"" onclick=""ClearCustomDesc()"" title=""Clear the Custom Description"">"
		strHTML = strHTML & 		"</td>"
		strHTML = strHTML & 	"</tr>"
		strHTML = strHTML & "</table>"
		strHTML = strHTML & "</div>"
		
		WaitMessage.InnerHTML = strHTML
		If booQueryAD = "0" Then cbxADDesc.Disabled = True
	End Sub
	
	Sub ChangePCDescription2(booExec)
		On Error Resume Next
		
		If cbxADDesc.Checked = True Then
			DescChooser.Value = ""
			DescChooser.Disabled = True
			txtCustomDesc.Value = ""
			txtCustomDesc.Disabled = True
			btnClearDescription.Disabled = True
			txtCustomDesc.style.backgroundcolor = "#dddddd"
			If booExec = True Then
				strADDesc = cbxADDesc.Value
				
				Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
				strPC & "\root\cimv2") 
				Set colOS = objWMIService.ExecQuery _
					("Select * from Win32_OperatingSystem")
				For Each objItem In ColOS
					objItem.Description = strADDesc
					objItem.Put_
					If IsNull(objItem.Description) = False Then
						MsgBox "The description on " & strPC & _
						" has now been changed to:" & vbCrLf & vbCrLf & _
						strADDesc, vbInformation, "Change Computer Description"
					End If
				Next
			End If
			Else
				DescChooser.Disabled = False
				txtCustomDesc.Disabled = False
				btnClearDescription.Disabled = False
				txtCustomDesc.style.backgroundcolor = "white"
		End If
		
		booCompare = False
		If DescChooser.Value <> "" Then
			For Each objOption in DescChooser.Options
				If objOption.Value = txtCustomDesc.Value Then booCompare = True
			Next 
			If booCompare = True OR txtCustomDesc.Value = "" Then
				txtCustomDesc.Value = DescChooser.Value
			End If
			If booExec = True Then
				Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
				strPC & "\root\cimv2") 
				Set colOS = objWMIService.ExecQuery _
					("Select * from Win32_OperatingSystem")
				For Each objItem In ColOS
					objItem.Description = DescChooser.Value
					objItem.Put_
					If IsNull(objItem.Description) = False Then
						MsgBox "The description on " & strPC & _
						" has now been changed to:" & vbCrLf & vbCrLf & _
						DescChooser.Value, vbInformation, "Change Computer Description"
					End If
				Next
			End If
		End If
		
		If txtCustomDesc.Value <> "" AND txtCustomDesc.Value <> DescChooser.Value  Then
			cbxADDesc.Checked = False
			DescChooser.Value = ""
			DescChooser.Disabled = True
			If booExec = True Then
				Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
				strPC & "\root\cimv2") 
				Set colOS = objWMIService.ExecQuery _
					("Select * from Win32_OperatingSystem")
				For Each objItem In ColOS
					objItem.Description = txtCustomDesc.Value
					objItem.Put_
					If IsNull(objItem.Description) = False Then
						MsgBox "The description on " & strPC & _
						" has now been changed to:" & vbCrLf & vbCrLf & _
						txtCustomDesc.Value, vbInformation, "Change Computer Description"
					End If
				Next
			End If
			ElseIf txtCustomDesc.Value = "" AND cbxADDesc.Checked = False Then
				DescChooser.Disabled = False
		End If
		
		If booExec = True Then
			ChangePCDescription()
		End If
	End Sub
	
	Sub ClearCustomDesc()
		txtCustomDesc.Value = ""
		DescChooser.Value = ""
		ChangePCDescription2 False
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ChangeLocalAdminPassword()
    '#	PURPOSE........:	Change the local Admin password for the PC
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	ChangeLocalAdminPassword()
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ChangeLocalAdminPassword()
		strHTML = "<hr><div style=""text-align:left;"">" & _
		"Please use the form below to change the local admin password:<p>" & _
		"Please enter the required password below:<br>" & _
		"<input type=""password"" name=""PasswordArea"" size=""30"" onKeyUp=""CheckStrength(PasswordArea.value)"">" & _
		"<input type=""text"" name=""txtPasswordArea"" size=""30"" onKeyUp=""CheckStrength(txtPasswordArea.value)"" " & _
		"style=""visibility:hidden;display:none;"">" & _
		"&nbsp;&nbsp;&nbsp;<span style=""font-weight:bold;"" id=""CheckStrengthArea""></span>" & _
		"<p>Please verify the password below:<br>" & _
		"<input type=""password"" name=""VerifyPasswordArea"" size=""30"">" & _
		"<input type=""text"" name=""txtVerifyPasswordArea"" size=""30"" " & _
		"style=""visibility:hidden;display:none;""><br>" & _
		"<input id=""RunButton"" class=""button"" type=""button"" value=""Set Password"" " & _
		"name=""btnSetPassword"" onclick=""ChangeLocalAdminPassword2()"" " & _
		"title=""Set the local admin password"">" & _
		"<span style=width:85px;text-align:right>Show" & _
		"<input type=""checkbox"" name=""cbxShowPassword"" onClick=ShowPasswordField() " & _
		"title=""Show the text entered"">" & _
		"</span></div>"
		WaitMessage.InnerHTML = strHTML
	End Sub
	
	Sub CheckStrength(strPassword)
		Dim arrStrength(6)
		arrStrength(0) = ""
		arrStrength(1) = "<font color=""Red"">Very Weak</font>"
		arrStrength(2) = "<font color=""#FF9900"">Weak</font>"
		arrStrength(3) = "<font color=""#FFFF33"">Medium</font>"
		arrStrength(4) = "<font color=""Green"">Strong</font>"
		arrStrength(5) = "Very Strong"

		intScore = 1
		x = 0

		If Len(strPassword) < 1 Then
			CheckStrengthArea.InnerHTML = arrStrength(0)
			Exit Sub
		End If

		If Len(strPassword) <= 4 Then
			CheckStrengthArea.InnerHTML = arrStrength(1)
			Exit Sub
		End If

		If Len(strPassword) >= 8 Then
			intScore = intScore + 1
		End If
		
		If Len(strPassword) >= 12 Then
			intScore = intScore + 1
		End If

		Set regex = New RegExp

		regex.Pattern = "\d+"
		If regex.Test(strPassword) Then
			x = x + 1
		End If

		regex.Pattern = "[a-z]"
		If regex.Test(strPassword) Then
			x = x + 1
		End If
		
		regex.Pattern = "[A-Z]"
		If regex.Test(strPassword) Then
			x = x + 1
		End If
		
		regex.Pattern = ".[!,@,#,$,%,^,&,*,?,_,~,-,�,(,)]"
		If regex.Test(strPassword) Then
			intScore = intScore + 1
			x = x + 1
		End If
		
		For i = 1 to Len(strPassword)
			z = 0
			strTestLetter = LCase(Mid(strPassword, i, 1))
			For y = 1 To Len(strPassword)
				If strTestLetter = LCase(Mid(strPassword, y, 1)) Then 
					z = z + 1
				End If
			Next
			pctTest = z / Len(strPassword) * 100
			If pctTest > 40 Then
				If intScore <> 1 Then 
					intScore = intScore - 1
					Else
						CheckStrengthArea.InnerHTML = arrStrength(1)
						Exit Sub
				End If
			End If
			If pctTest = 100 Then
				CheckStrengthArea.InnerHTML = arrStrength(1)
				Exit Sub
			End If
		Next

		If x >= 2 Then intScore = intScore + 1
		If intScore >= 5 Then intScore = 5
		If Len(strPassword) < 12 AND intScore = 5 Then intScore = 4
		If Len(strPassword) < 8 AND intScore > 3 Then intScore = 3
			
		CheckStrengthArea.InnerHTML = arrStrength(intScore)
	End Sub

	
	Sub ChangeLocalAdminPassword2()
		On Error Resume Next
		Set objUser = GetObject("WinNT://" & strPC & "/Administrator")
		If cbxShowPassword.Checked Then
			strPassword = txtPasswordArea.Value
			strPassword2 = txtVerifyPasswordArea.Value
			Else
				strPassword = PasswordArea.Value
				strPassword2 = VerifyPasswordArea.Value
		End If
		If strPassword <> strPassword2 Then
			MsgBox "The passwords do not match!", vbExclamation,"Change Local Admin Password"
			Exit Sub
		End If
		objUser.SetPassword(strPassword)
		If Err.Number <> 0 Then
			Select Case Err.Number
				Case -2147022651
					MsgBox "There was an error setting the local Admin password for " & strPC & "." & vbCrLf & _
					vbCrLf & "The password does not meet the password policy requirements. " & _
					"Check the minimum password length, password complexity and password history requirements.", _
					vbExclamation, "Change Local Admin Password"
				Case Else
					MsgBox "There was an unspecified error setting the local Admin password for " & strPC, _
					vbExclamation,"Change Local Admin Password"
			End Select
			Err.Clear
			Else
				MsgBox "The local Admin password has now been changed on " & strPC, vbInformation, _
				"Change Local Admin Password"
		End If
		PasswordArea.Value = ""
		txtPasswordArea.Value = ""
		VerifyPasswordArea.Value = ""
		txtVerifyPasswordArea.Value = ""
		CheckStrengthArea.InnerHTML = ""
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ShowPasswordField()
    '#	PURPOSE........:	Show / Hide password text in Change Local Admin
	'#						Password routine
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ShowPasswordField()
		If cbxShowPassword.Checked Then
			txtPasswordArea.Value = PasswordArea.Value
			txtVerifyPasswordArea.Value = VerifyPasswordArea.Value
			PasswordArea.style.visibility = "hidden"
			PasswordArea.style.display = "none"
			txtPasswordArea.style.visibility = "visible"
			txtPasswordArea.style.display = "inline"
			VerifyPasswordArea.style.visibility = "hidden"
			VerifyPasswordArea.style.display = "none"
			txtVerifyPasswordArea.style.visibility = "visible"
			txtVerifyPasswordArea.style.display = "inline"
			Else
				PasswordArea.Value = txtPasswordArea.Value
				VerifyPasswordArea.Value = txtVerifyPasswordArea.Value
				txtPasswordArea.style.visibility = "hidden"
				txtPasswordArea.style.display = "none"
				PasswordArea.style.visibility = "visible"
				PasswordArea.style.display = "inline"
				txtVerifyPasswordArea.style.visibility = "hidden"
				txtVerifyPasswordArea.style.display = "none"
				VerifyPasswordArea.style.visibility = "visible"
				VerifyPasswordArea.style.display = "inline"
		End If
	End Sub
	
	Sub CleanProfiles()
		On Error Resume Next
		
		intProfAge = objShell.RegRead(strRegStart & "\Software\SKB\PCManagementUtil\Tabs\intProfAge")
		If intProfAge = "" OR IsNull(intProfAge) Then 
			objShell.RegWrite strRegStart & "\Software\SKB\PCManagementUtil\Tabs\intProfAge", _
			"90", "REG_SZ"
			intProfAge = "90"
		End If
		strHTML = "<hr><div style=""text-align:left;"">" & _
		"<p><input type=""text"" name=""txtProfAge"" value=""" & intProfAge & """ size=""1"" style=""text-align:center;""> " & _
		"&nbsp;Number of days to keep the user profiles on " & strPC & "<p>" & _
		"<input id=""RunButton"" class=""button"" type=""button"" value=""Begin Deletion"" " & _
		"name=""btnDeleteProf"" onclick=""CleanProfiles2()"" " & _
		"title=""Delete all profiles older than the specified number of days"">" & _
		"</div>"
		WaitMessage.InnerHTML = strHTML
		txtProfAge.Focus()
		txtProfAge.Select()
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	CleanProfiles()
    '#	PURPOSE........:	Deletes old user profiles of a specified age
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	Will delete profile folders and reg keys
    '#--------------------------------------------------------------------------
	Sub CleanProfiles2()
		intAge = txtProfAge.Value
		intFolderCount = 0
		strRoot = GetRoot()
		If IsNull(intAge) OR intAge = "" Then
			Exit Sub
		End If
		If IsNumeric(intAge) = False Then
			MsgBox "'" & intAge & "' is not a valid response", vbExclamation, "Delete Old User Profiles"
			Exit Sub
		End If
		dtmEarliestDate = DateAdd ("d", -intAge, Date)
		If objFSO.FolderExists(strRoot & "\documents and settings") Then
			strProfile = strRoot & "\documents and settings"
			ElseIf objFSO.FolderExists(strRoot & "\users") Then
				strProfile = strRoot & "\users"
		End If
		
		strCurrentProfiles = objShell.RegRead(strRegStart & _
		"\Software\SKB\PCManagementUtil\Setup\strCurrentProfiles")
		
		Set strHoldProf = objFSO.getFolder(strProfile)
		Set arrProfFolders = strHoldProf.SubFolders
		For Each objItem in arrProfFolders
			If objItem.DateLastModified < dtmEarliestDate _
			AND InStr(LCase(objItem.Name) & ";", LCase(strCurrentProfiles)) = 0 _
			AND LCase(objItem.Name) <> "administrator" _
			AND LCase(objItem.Name) <> "all users" _
			AND LCase(objItem.Name) <> "localservice" _
			AND LCase(objItem.Name) <> "default user" _
			AND LCase(objItem.Name) <> "networkservice" Then
				intFolderCount = intFolderCount + 1
				strMsg = strMsg & objItem.Name & "; "
			End If
		Next	
		If intFolderCount > 0 Then
			CleanProfPrompt = MsgBox("WARNING: You are about to delete the " & _
			"following " & intFolderCount & " user profiles on " & strPC & ":  " & _
			vbCrLf & vbCrLf & strMsg & vbCrLf & vbCrLf & _
			"Do you wish to continue?", vbExclamation+vbYesNo,"Delete Old User Profiles")
			If CleanProfPrompt = vbYes Then
				On Error Resume Next		
				intProgDone = 0
				intProgTotal = intFolderCount * 2
				intProgMult = 100 / intProgTotal
				For Each objItem in arrProfFolders
					If objItem.DateLastModified < dtmEarliestDate _
					AND InStr(LCase(objItem.Name) & ";", LCase(strCurrentProfiles)) = 0 _
					AND LCase(objItem.Name) <> "administrator" _
					AND LCase(objItem.Name) <> "all users" _
					AND LCase(objItem.Name) <> "localservice" _
					AND LCase(objItem.Name) <> "default user" _
					AND LCase(objItem.Name) <> "networkservice" Then
						On Error Resume Next
						intProgDone = intProgDone + 1
						UpdateProgressBar intProgMult,intProgDone,intProgTotal
						DeleteFolderContents objItem
						intProgDone = intProgDone + 1
						UpdateProgressBar intProgMult,intProgDone,intProgTotal
					End If
				Next
				
				intProgDone = 0
				intProgTotal = 3
				intProgMult = 100 / intProgTotal
				
				intProgDone = intProgDone + 1
				UpdateProgressBar intProgMult,intProgDone,intProgTotal
				
				Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
				strPC & "\root\default:StdRegProv") 
						
				intProgDone = intProgDone + 1
				UpdateProgressBar intProgMult,intProgDone,intProgTotal
						
				arrProfValues = Split(strMsg, "; ")
				For i = 0 to UBound(arrProfValues) - 1
					ReDim Preserve arrSIDValues(i)
					arrSIDValues(i) = GetSIDFromUser(arrProfValues(i))
				Next
				
				intProgDone = intProgDone + 1
				UpdateProgressBar intProgMult,intProgDone,intProgTotal
				
				For i = 0 to UBound(arrSIDValues) - 1
					strGuid = ""
					strSubKey = arrSIDValues(i)
					If strSubKey <> "" Then
						strUserName = arrProfValues(i)
						strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" & strSubKey
						objReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"Guid", strGuid
						strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" & strSubKey
						DeleteAllSubkeysMACHINE strKeyPath
						strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\" & strSubKey
						DeleteAllSubkeysMACHINE strKeyPath
						strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\" & strSubKey
						DeleteAllSubkeysMACHINE strKeyPath
						strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\" & strSubKey
						DeleteAllSubkeysMACHINE strKeyPath
						If strGuid <> "" Then
							strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\PolicyGuid\" & strGuid
							DeleteAllSubkeysMACHINE strKeyPath
							strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileGuid\" & strGuid
							DeleteAllSubkeysMACHINE strKeyPath
						End If
					End If
				Next
				
				strOrphanMsg = ""
				strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
				objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys
		
				For Each objItem In arrSubkeys
					strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
					strValueName = "ProfileImagePath"
					strSubPath = strKeyPath & "\" & objItem
					objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath,strValueName,strValue
					strGuid = ""
					strSubKey = objItem
					
					If strValue <> "" Then
						strRoot = GetRoot()
						strProfPath = Right(strValue,Len(strValue) - 2)
						strProfPath = strRoot & strProfPath
						If objFSO.FolderExists(strProfPath) = False Then
							arrUserName = Split(strProfPath, "\")
							strUserName = arrUserName(UBound(arrUserName))
							strOrphanMsg = strOrphanMsg & strUserName & "; "
							strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" & strSubKey
							objReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"Guid", strGuid
							strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" & strSubKey
							DeleteAllSubkeysMACHINE strKeyPath
							strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\" & strSubKey
							DeleteAllSubkeysMACHINE strKeyPath
							strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\" & strSubKey
							DeleteAllSubkeysMACHINE strKeyPath
							strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\" & strSubKey
							DeleteAllSubkeysMACHINE strKeyPath
							If strGuid <> "" Then
								strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\PolicyGuid\" & strGuid
								DeleteAllSubkeysMACHINE strKeyPath
								strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileGuid\" & strGuid
								DeleteAllSubkeysMACHINE strKeyPath
							End If
						End If
					End If
				Next
				If strOrphanMsg = "" Then				
					MsgBox "The following " & intFolderCount & " user profile(s) on " & strPC & " have now been deleted:  " & _
					vbCrLf & vbCrLf & strMsg, vbInformation,"Delete Old User Profiles"		
					Else
						MsgBox "The following " & intFolderCount & " user profile(s) on " & strPC & " have now been deleted:  " & _
						vbCrLf & vbCrLf & strMsg & vbCrLf & vbCrLf & "Deleted the orphaned reg keys for the following user(s): " & _
						vbCrLf & vbCrLf & strOrphanMsg, vbInformation,"Delete Old User Profiles"
				End If
			End If
			Else
				
				Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
				strPC & "\root\default:StdRegProv") 
				strOrphanMsg = ""
				strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
				objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys
				For Each objItem In arrSubkeys
					strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
					strValueName = "ProfileImagePath"
					strSubPath = strKeyPath & "\" & objItem
					objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath,strValueName,strValue
					strGuid = ""
					strSubKey = objItem
					
					If strValue <> "" Then
						strRoot = GetRoot()
						strProfPath = Right(strValue,Len(strValue) - 2)
						strProfPath = strRoot & strProfPath
						If objFSO.FolderExists(strProfPath) = False Then
							arrUserName = Split(strProfPath, "\")
							strUserName = arrUserName(UBound(arrUserName))
							strOrphanMsg = strOrphanMsg & strUserName & "; "
							strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" & strSubKey
							objReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"Guid", strGuid
							strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" & strSubKey
							DeleteAllSubkeysMACHINE strKeyPath
							strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\" & strSubKey
							DeleteAllSubkeysMACHINE strKeyPath
							strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\" & strSubKey
							DeleteAllSubkeysMACHINE strKeyPath
							strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\" & strSubKey
							DeleteAllSubkeysMACHINE strKeyPath
							If strGuid <> "" Then
								strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\PolicyGuid\" & strGuid
								DeleteAllSubkeysMACHINE strKeyPath
								strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileGuid\" & strGuid
								DeleteAllSubkeysMACHINE strKeyPath
							End If
						End If
					End If
				Next
				If strOrphanMsg = "" Then
					MsgBox "No user profiles found of the specified age on " & strPC, vbInformation, "Delete Old User Profiles"
					Else
						MsgBox "No user profiles found of the specified age on " & strPC & _
						vbCrLf & vbCrLf & "Deleted the orphaned reg keys for the following users: " & _
						vbCrLf & vbCrLf & strOrphanMsg, vbInformation, "Delete Old User Profiles"
				End If
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	DeleteAllSubkeysMACHINE(strKeyPath)
    '#	PURPOSE........:	Deletes all the subkeys in the specfied Reg keypath
    '#	ARGUMENTS......:	strKeyPath = the full path to the reg key to remove
    '#	EXAMPLE........:	DeleteAllSubkeysMACHINE("\SOFTWARE\Microsoft")
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------	
	Sub DeleteAllSubkeysMACHINE(strKeyPath)
		On Error Goto 0
		Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\default:StdRegProv")
		objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys 
		If IsArray(arrSubkeys) Then 
			For Each objItem In arrSubkeys
				objReg.DeleteKey HKEY_LOCAL_MACHINE, strKeyPath & "\" & objItem
				DeleteAllSubkeysMACHINE strKeyPath & "\" & objItem
			Next 
		End If 
		objReg.DeleteKey HKEY_LOCAL_MACHINE, strKeyPath
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	DeleteFolderContents(strFolder)
    '#	PURPOSE........:	Deletes all the contents of the specified folder
    '#	ARGUMENTS......:	strFolder = full path to the folder
    '#	EXAMPLE........:	DeleteFolderContents("c:\Temp")
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub DeleteFolderContents(strFolder)
		On Error Resume Next
		Set objFolder=objFSO.GetFolder(strFolder)
		If Err.Number <> 0 Then
			Err.Clear
			Exit Sub
		End If
		For Each objItem In objFolder.SubFolders
			Err.Clear
			objItem.Delete True
			If Err.Number <> 0 Then
				objShell.Run "%COMSPEC% /c del /f /q " & Chr(34) & objItem & Chr(34),0,True
			End If
			Err.Clear
		Next
		For Each objItem In objFolder.Files
			Err.Clear
			objItem.Delete True
			If Err.Number <> 0 Then
				objShell.Run "%COMSPEC% /c del /f /q " & Chr(34) & objItem & Chr(34),0,True
			End If
		Next
		objFolder.Delete True
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ClearAppEventLog()
    '#	PURPOSE........:	Backs up and clears the Application Event log of
	'#						the PC
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ClearAppEventLog()
		On Error Resume Next
		
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate,(Backup)}!\\" & _
		strPC & "\root\cimv2") 
		
		If NOT objFSO.FolderExists("\\" & strPC & "\C$\SKB") Then
			Set objFolder = objFSO.CreateFolder("\\" & strPC & "\C$\SKB")
		End If
		
		strBackupEventLog = "\\" & strPC & "\C$\SKB\EventLogBackup\"
		
		If NOT objFSO.FolderExists(strBackupEventLog) Then
			Set objFolder = objFSO.CreateFolder(strBackupEventLog)
		End If
		
		strFileName = "application.evt"
		
		If objFSO.FileExists(strBackupEventLog & strFileName) Then
			If objFSO.FileExists(strBackupEventLog & "application" & strFileDate & ".evt") Then
				For i = 1 to 5
					If NOT objFSO.FileExists(strBackupEventLog & "application" & _
					strFileDate & i & ".evt") Then
						strFileName = "application" & strFileDate & i & ".evt"
						Exit For
					End If
				Next
				Else
					strFileName = "application" & strFileDate & ".evt"
			End If
			
			Set objFile = objFSO.GetFile(strBackupEventLog & "application.evt")
			objFile.Name = strFileName

			Set objFile = Nothing
		End If
		
		Set colLogFiles = objWMIService.ExecQuery _
			("Select * from Win32_NTEventLogFile where LogFileName='Application'")

		For Each objItem in colLogFiles
			errBackupLog = objItem.BackupEventLog(strBackupEventLog & "application.evt")
			If errBackupLog <> 0 Then        
				MsgBox "The Application event log could not be backed up",vbExclamation,"Clear Application Event Log"
			Else
				objItem.ClearEventLog()
				MsgBox "The Application event log has now been backed up and cleared on " & GetPCName(), vbInformation, _
				"Clear Application Event Log"
			End If
		Next
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	CopyProfile()
    '#	PURPOSE........:	Copies a user profile from / to the PC. This
	'#						can be copied to / from another PC or to a USB drive
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub CopyProfile()		
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\cimv2") 
		
		Set objWMIService2 = GetObject("winmgmts:\\.\root\cimv2")
		
		strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
		
		ProfilePromptMedia = ""
		strDestinationPC = ""
		strSourcePC = ""
		booProfileUserExists = False
		booSrcProfileUserExists = False
		booDestProfileUserExists = False
		OverwriteExisting = True
		iSUCCESS = 0
		iERROR = 1 
		iWARNING = 2
		iINFORMATION = 4
		Dim arrDrive()
		Dim arrUserAcc
		ReDim arrUserAcc(5)
		InArray = False
		i = 1
		x = 1

		ProfilePrompt = InputBox ("1. Source PC, ie. copy profile FROM this PC" & vbCrLf &_
		"2. Destination PC, ie. copy profile TO this PC" & vbCrLf & vbCrLf & _
		"Is " & strPC & " the Source or Destination PC?", "Copy Profile")
		If ProfilePrompt = "1" Then
			strSource = "\\" & strPC & "\C$"
			ProfilePromptMedia = InputBox ("1. Another PC" & vbCrLf &_
			"2. A USB drive or other removable media" & vbCrLf & vbCrLf & _
			"Is the destination another PC or a form of removable media?", "Copy Profile")
			If ProfilePromptMedia = "1" Then
				strDestination = InputBox("Please enter the Computer Name of the " & _
				"Destination PC","Copy Profile")
				If Reachable(strDestination) Then
					strDestinationPC = strDestination
					strDestination = "\\" & strDestination & "\C$"
					Else
					MsgBox "Error connecting to " & strDestination & ".", vbExclamation,"Copy Profile"
						Exit Sub
				End If
				ElseIf ProfilePromptMedia = "2" Then
					MsgBox "Please make sure the USB drive is currently plugged into" & _
					"your computer and then click OK",vbInformation,"Copy Profile"
					Set colDisk = objWMIService2.ExecQuery _
						("Select * from Win32_LogicalDisk")
							
					For Each objItem In colDisk
						strDriveType = objItem.DriveType
						If strDriveType = 2 Then
							strDriveName = objItem.Caption
							ReDim Preserve arrDrive(i)
							arrDrive(i) = strDriveName
							i = i + 1
						End If
					Next
					If i > 1 Then
						For i = 1 to UBound(arrDrive)
							strProfileMediaMsg = strProfileMediaMsg & _
							i & ". " & arrDrive(i) & " Drive" & vbCrLf
						Next
						intProfileRem = InputBox(strProfileMediaMsg & vbCrLf & _
						"Please choose the drive you would like to use.","Copy Profile")
						
						For a = 1 to UBound(arrDrive)
							If InStr(a,intProfileRem) = 1 Then
								InArray = True
							End If
						Next

						If InArray = False Then 
							MsgBox "'" & intProfileRem & "' is not a valid response", _
							vbExclamation, "Copy Profile"
							Exit Sub
						End If
					
						If IsNull(intProfileRem) OR intProfileRem = "" Then
							Exit Sub
						End If
						strDestination = arrDrive(intProfileRem)
						Else
							MsgBox "No USB drives have been found",vbExclamation, "Copy Profile"
							Exit Sub
					End If
					ElseIf ProfilePromptMedia = "" Then
						Exit Sub
						Else
							MsgBox "'" & ProfilePromptMedia & "' is not a valid response", _
							vbExclamation, "Copy Profile"
							Exit Sub
			End If
			ElseIf ProfilePrompt = "2" Then
				strDestination = "\\" & strPC & "\C$"
				strSource = InputBox("Please enter the Computer Name of the Source PC","Copy Profile")
				If IsNull(strSource) OR strSource = "" OR strSource = "." Then
					Exit Sub
				End If
				If Reachable(strSource) Then
					strSourcePC = strSource
					strSource = "\\" & strSource & "\C$"
					Else
						MsgBox "Error connecting to " & strSource, vbExclamation,"Copy Profile"
						Exit Sub
				End If
				ElseIf ProfilePrompt = "" Then
					Exit Sub
					Else
						MsgBox "'" & ProfilePrompt & "' is not a valid response", _
						vbExclamation, "Copy Profile"
						Exit Sub
		End If

		Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\default:StdRegProv") 
		strProfilePath = "\Documents and Settings\"
		objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys
		
		For Each objItem In arrSubkeys
			strValueName = "ProfileImagePath"
			strSubPath = strKeyPath & "\" & objItem
			objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath,strValueName,strValue
			If InStr(LCase(strValue), "c:\documents and settings\") = 1 Then
				ReDim Preserve arrUserAcc(x)
				arrUserAcc(x) = Replace(strValue,"c:\documents and settings\","")
				x = x + 1
			End If
		Next
		For x = 1 to UBound(arrUserAcc)
			strProfileUserAccMsg = strProfileUserAccMsg & _
			x & ". " & arrUserAcc(x) & vbCrLf
		Next
		intProfileUser = InputBox(strProfileUserAccMsg & vbCrLf & _ 
		"Please choose a user account to copy", "Copy Profile")
		If IsNull(intProfileUser) OR intProfileUser = "" Then
			Exit Sub
		End If
		
		InArray = False
		For a = 1 to UBound(arrUserAcc)
		If InStr(a,intProfileUser) = 1 Then
				InArray = True
			End If
		Next
		If InArray = False Then 
			MsgBox "'" & intProfileUser & "' is not a valid response", _
			vbExclamation, "Copy Profile"
			Exit Sub
		End If
		If InStr(LCase(arrUserAcc(intProfileUser)), "c:\documents and settings\") > 0 Then
			strUserProf = arrUserAcc(intProfileUser)
			arrUserAccSplit = Split(strUserProf, "\")
			strUserAcc = arrUserAccSplit(2)
			strGlobalLogFileName = strUserAcc
		End If
		For Each objItem In arrSubkeys
			strValueName = "ProfileImagePath"
			strSubPath = strKeyPath & "\" & objItem
			objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath,strValueName,strValue
			If InStr(LCase(strValue), "c:\documents and settings\") > 0 Then
				strValue = Replace(strValue, "C:", "\\" & strPC & "\C$")
				If UCase(strValue) = "\\" & strPC & "\C$" & UCase(strProfilePath) & UCase(strUserAcc) Then
					booProfileUserExists = True
				End If
			End If
		Next
		If strSourcePC <> "" Then
			Set objReg2 = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
			strSourcePC & "\root\default:StdRegProv") 
 
			objReg2.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys
		
			For Each objItem In arrSubkeys
				strValueName = "ProfileImagePath"
				strSubPath = strKeyPath & "\" & objItem
				objReg2.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath,strValueName,strValue
				If InStr(LCase(strValue), "c:\documents and settings\") > 0 Then
					strValue=Replace(strValue,"C:","\\" & strSourcePC & "\C$")
					If strValue = "\\" & strSourcePC & "\C$" & strProfilePath & strUserAcc Then
						booSrcProfileUserExists = True
					End If
				End If
			Next
		End If		
		If strDestinationPC <> "" Then
			Set objReg2 = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
			strDestinationPC & "\root\default:StdRegProv") 
 
			objReg2.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys
		
			For Each objItem In arrSubkeys
				strValueName = "ProfileImagePath"
				strSubPath = strKeyPath & "\" & objItem
				objReg2.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath,strValueName,strValue
				If InStr(LCase(strValue), "c:\documents and settings\") = 1 Then
					strValue=Replace(strValue,"C:","\\" & strDestinationPC & "\C$")
					If strValue = "\\" & strDestinationPC & "\C$" & strProfilePath & strUserAcc Then
						booDestProfileUserExists = True
					End If
				End If
			Next
		End If
		If booProfileUserExists = True Then
			If booSrcProfileUserExists = True OR booDestProfileUserExists = True _
			OR ProfilePromptMedia = "2" Then
				strCopySource = strSource & strProfilePath & strUserAcc
				strCopyDestination = strDestination & strProfilePath & strUserAcc
				ProfileCont = MsgBox("The script will now copy all of the required files from " & _
				strCopySource & " to " & strCopyDestination & vbCrLf & vbCrLf & _
				"Would you like to continue?",vbQuestion+vbYesNo,"Copy Profile")
				
				On Error Resume Next
				
				If ProfileCont = vbYes Then
					intProgDone = 0
					intProgTotal = 9
					intProgMult = 100 / intProgTotal
					
					strLogEntry = "Started Script"
					Call ScriptLog(strLogEntry, iINFORMATION)
					Set SourceFolder = objFSO.GetFolder(strCopySource & _
					"\Application Data\Microsoft\Outlook\")
					strExtension = "nk2"
					booExtExists = False
					intProgDone = intProgDone + 1
					UpdateProgressBar intProgMult,intProgDone,intProgTotal	'1
					strSourceCopyFile = strCopySource & "\Application Data\Microsoft\Outlook\*.nk2"
					strDestinationFolder = strCopyDestination & "\Application Data\Microsoft\Outlook\"
					Set checkFolder = SourceFolder.Files
					For Each objItem in checkFolder
						Extension = objFSO.GetExtensionName(LCase(objItem.name))
						If Extension = strExtension Then
							booExtExists = True
						End If
					Next
					intProgDone = intProgDone + 1
					UpdateProgressBar intProgMult,intProgDone,intProgTotal	'2
					If booExtExists = True Then
						If NOT objFSO.FolderExists(strDestination & strProfilePath) Then
							Set objFolder = objFSO.CreateFolder(strDestination & strProfilePath)
						End If
						If NOT objFSO.FolderExists(strDestination & strProfilePath & _
						strUserAcc) Then
							Set objFolder = objFSO.CreateFolder(strDestination & _
							strProfilePath & strUserAcc)
						End If
						If NOT objFSO.FolderExists(strDestination & strProfilePath & _
						strUserAcc & "\Application Data") Then
							Set objFolder = objFSO.CreateFolder(strDestination & _
							strProfilePath & strUserAcc & "\Application Data")
						End If
						If NOT objFSO.FolderExists(strDestination & strProfilePath & _
						strUserAcc & "\Application Data\Microsoft") Then
							Set objFolder = objFSO.CreateFolder(strDestination & _
							strProfilePath & strUserAcc & "\Application Data\Microsoft")
						End If
						If NOT objFSO.FolderExists(strDestination & strProfilePath & _
						strUserAcc & "\Application Data\Mozilla") Then
							Set objFolder = objFSO.CreateFolder(strDestination & _
							strProfilePath & strUserAcc & "\Application Data\Mozilla")
						End If
						If NOT objFSO.FolderExists(strDestination & strProfilePath & _
						strUserAcc & "\Application Data\Mozilla\Firefox") Then
							Set objFolder = objFSO.CreateFolder(strDestination & _
							strProfilePath & strUserAcc & "\Application Data\Mozilla\Firefox")
						End If
						If NOT objFSO.FolderExists(strDestinationFolder) Then
							Set objFolder = objFSO.CreateFolder(strDestinationFolder)
						End If
						strLogEntry = "Copying Outlook NK2 (Auto-complete) files"
						Call ScriptLog(strLogEntry, iINFORMATION)
						objFSO.CopyFile strSourceCopyFile, strDestinationFolder, OverwriteExisting
						If Err.Number <> 0 Then
							strLogEntry = "Problem copying files from " & strSourceFolder & _
							" to " & strDestinationFolder
							strErr = Err.Number & " : " & Err.Description
							  
							Call ScriptLog(strLogEntry, iERROR)
							Call ScriptLog(strErr, iERROR)
							Else
								strLogEntry = "Succeeded copying files"
								Call ScriptLog(strLogEntry, iSUCCESS)
						End If 
					End If
					
					intProgDone = intProgDone + 1
					UpdateProgressBar intProgMult,intProgDone,intProgTotal	'3
					
					strSourceFolder = strCopySource & "\Desktop"
					strDestinationFolder = strCopyDestination & "\Desktop"
					If objFSO.FolderExists(strSourceFolder) Then
						If NOT objFSO.FolderExists(strDestinationFolder) Then
							Set objFolder = objFSO.CreateFolder(strDestinationFolder)
						End If
						strLogEntry = "Copying Desktop items"
						Call ScriptLog(strLogEntry, iINFORMATION)
						objFSO.CopyFolder strSourceFolder, strDestinationFolder
						If Err.Number <> 0 Then
							strLogEntry = "Problem copying files from " & strSourceFolder & _
							" to " & strDestinationFolder
							strErr = Err.Number & " : " & Err.Description
							Call ScriptLog(strLogEntry, iERROR)
							Call ScriptLog(strErr, iERROR)
							Else
								strLogEntry = "Succeeded copying files"
								Call ScriptLog(strLogEntry, iSUCCESS)
						End If 
					End If
					
					intProgDone = intProgDone + 1
					UpdateProgressBar intProgMult,intProgDone,intProgTotal	'4
					
					strSourceFolder = strCopySource & "\Application Data\Mozilla\Firefox\Profiles"
					strDestinationFolder = strCopyDestination & "\Application Data\Mozilla\Firefox\Profiles"
					If objFSO.FolderExists(strSourceFolder) Then
						If NOT objFSO.FolderExists(strDestinationFolder) Then
							Set objFolder = objFSO.CreateFolder(strDestinationFolder)
						End If
						objFSO.CopyFolder strSourceFolder, strDestinationFolder
						strLogEntry = "Copying Firefox profile"
						Call ScriptLog(strLogEntry, iINFORMATION)
						objFSO.CopyFile strSourceCopyFile, strDestinationFolder, _
						OverwriteExisting
						If Err.Number <> 0 Then
							strLogEntry = "Problem copying files from " & strSourceFolder & _
							" to " & strDestinationFolder
							strErr = Err.Number & " : " & Err.Description
							Call ScriptLog(strLogEntry, iERROR)
							Call ScriptLog(strErr, iERROR)
							Else
								strLogEntry = "Succeeded copying files"
								Call ScriptLog(strLogEntry, iSUCCESS)
						End If 
					End If
					
					intProgDone = intProgDone + 1
					UpdateProgressBar intProgMult,intProgDone,intProgTotal	'5
					
					Set SourceFolder = objFSO.GetFolder(strCopySource & _
					"\Local Settings\Application Data\Microsoft\Outlook\")
					strExtension = "pst"
					booExtExists = False
					strSourceCopyFile = strCopySource & _
					"\Local Settings\Application Data\Microsoft\Outlook\*.pst"
					strSourceFolder = strCopySource & _
					"\Local Settings\Application Data\Microsoft\Outlook\"
					strDestinationFolder = strCopyDestination & _
					"\Local Settings\Application Data\Microsoft\Outlook\"
					Set checkFolder = SourceFolder.Files
					For Each objItem in checkFolder
						Extension = objFSO.GetExtensionName(LCase(objItem.name))
						If Extension = strExtension Then
							booExtExists = True
						End If
					Next
					If booExtExists = True Then
						If NOT objFSO.FolderExists(strDestination & strProfilePath & _
						strUserAcc & "\Local Settings") Then
							Set objFolder = objFSO.CreateFolder(strDestination & _
							strProfilePath & strUserAcc & "\Local Settings")
						End If
						If NOT objFSO.FolderExists(strDestination & strProfilePath & _
						strUserAcc & "\Local Settings\Application Data") Then
							Set objFolder = objFSO.CreateFolder(strDestination & _
							strProfilePath & strUserAcc & "\Local Settings\Application Data")
						End If
						If NOT objFSO.FolderExists(strDestination & strProfilePath & _
						strUserAcc & "\Local Settings\Application Data\Microsoft") Then
							Set objFolder = objFSO.CreateFolder(strDestination & _
							strProfilePath & strUserAcc & "\Local Settings\Application Data\Microsoft")
						End If
						If NOT objFSO.FolderExists(strDestinationFolder) Then
							Set objFolder = objFSO.CreateFolder(strDestinationFolder)
						End If
						strLogEntry = "Copying Outlook PST files"
						Call ScriptLog(strLogEntry, iINFORMATION)
						objFSO.CopyFile strSourceCopyFile, strDestinationFolder, OverwriteExisting
						If Err.Number <> 0 Then
							strLogEntry = "Problem copying files from " & strSourceFolder & _
							" to " & strDestinationFolder
							strErr = Err.Number & " : " & Err.Description
							Call ScriptLog(strLogEntry, iERROR)
							Call ScriptLog(strErr, iERROR)
							Else
								strLogEntry = "Succeeded copying files"
								Call ScriptLog(strLogEntry, iSUCCESS)
						End If 
					End If
					
					intProgDone = intProgDone + 1
					UpdateProgressBar intProgMult,intProgDone,intProgTotal	'6
					
					strSourceFolder = strCopySource & "\Application Data\Microsoft\Signatures"
					strDestinationFolder = strCopyDestination & "\Application Data\Microsoft\Signatures"
					If objFSO.FolderExists(strSourceFolder) Then
						If NOT objFSO.FolderExists(strDestinationFolder) Then
							Set objFolder = objFSO.CreateFolder(strDestinationFolder)
						End If
						strLogEntry = "Copying Outlook Signatures"
						Call ScriptLog(strLogEntry, iINFORMATION)
						objFSO.CopyFolder strSourceFolder, strDestinationFolder
						If Err.Number <> 0 Then
							strLogEntry = "Problem copying files from " & strSourceFolder & _
							" to " & strDestinationFolder
							strErr = Err.Number & " : " & Err.Description
							Call ScriptLog(strLogEntry, iERROR)
							Call ScriptLog(strErr, iERROR)
							Else
								strLogEntry = "Succeeded copying files"
								Call ScriptLog(strLogEntry, iSUCCESS)
						End If 
					End If
					
					intProgDone = intProgDone + 1
					UpdateProgressBar intProgMult,intProgDone,intProgTotal	'7
					
					strSourceFolder = strCopySource & "\Favorites"
					strDestinationFolder = strCopyDestination & "\Favorites"
					If objFSO.FolderExists(strSourceFolder) Then
						If NOT objFSO.FolderExists(strDestinationFolder) Then
							Set objFolder = objFSO.CreateFolder(strDestinationFolder)
						End If
						strLogEntry = "Copying IE Favorites"
						Call ScriptLog(strLogEntry, iINFORMATION)
						objFSO.CopyFolder strSourceFolder, strDestinationFolder
						If Err.Number <> 0 Then
							strLogEntry = "Problem copying files from " & strSourceFolder & _
							" to " & strDestinationFolder
							strErr = Err.Number & " : " & Err.Description
							Call ScriptLog(strLogEntry, iERROR)
							Call ScriptLog(strErr, iERROR)
							Else
								strLogEntry = "Succeeded copying files"
								Call ScriptLog(strLogEntry, iSUCCESS)
						End If
					End If					

					intProgDone = intProgDone + 1
					UpdateProgressBar intProgMult,intProgDone,intProgTotal	'8
					
					strSourceFolder = strCopySource & "\My Documents"
					strDestinationFolder = strCopyDestination & "\My Documents"
					If objFSO.FolderExists(strSourceFolder) Then
						If NOT objFSO.FolderExists(strDestinationFolder) Then
							Set objFolder = objFSO.CreateFolder(strDestinationFolder)
						End If
						strLogEntry = "Copying My Documents folder"
						Call ScriptLog(strLogEntry, iINFORMATION)
						objFSO.CopyFolder strSourceFolder, strDestinationFolder
						If Err.Number <> 0 Then
							strLogEntry = "Problem copying files from " & strSourceFolder & _
							" to " & strDestinationFolder
							strErr = Err.Number & " : " & Err.Description 
							Call ScriptLog(strLogEntry, iERROR)
							Call ScriptLog(strErr, iERROR)
							Else
								strLogEntry = "Succeeded copying files"
								Call ScriptLog(strLogEntry, iSUCCESS)
						End If
					End If
					
					intProgDone = intProgDone + 1
					UpdateProgressBar intProgMult,intProgDone,intProgTotal	'9
					
					strLogEntry = "Ended Script"
					Call ScriptLog(strLogEntry, iINFORMATION)
					MsgBox "The files have now been copied from " & strCopySource & _
					" to " & strCopyDestination & vbCrLf & vbCrLf & _
					"A report can be viewed at \\" & strPC & _
					"\C$\SKB\ProfileCopy\ProfileCopy-" & strGlobalLogFileName & "-" & _
					strFileDate & ".log",vbInformation, "Copy Profile"
				End If
				On Error GoTo 0
				Else
					MsgBox "The user account " & strUserAcc & _
					" does not exist on at least one of the PCs",vbExclamation, "Copy Profile"
			End If
			Else
				MsgBox "The user account " & strUserAcc & _
				" does not exist on at least one of the PCs",vbExclamation, "Copy Profile"
		End If	
	End Sub

	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	DeleteTempFiles()
    '#	PURPOSE........:	Deletes all of the Temp files from the PC
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	Will delete the following files:
	'#						Hotfix uninstallers over 90 days old (if required)
	'#						Temporary files folder from all profiles
	'#						IE History from all profiles
	'#						Temporary Internet Files from all profiles
	'#						Mozilla Firefox Cache from all profiles 
	'#						Will delete the following reg entries:
	'#						Run / Recent Docs / Regedit Last Key MRUs
	'#						Paint Recent / Wordpad Recent MRUs
	'#						Windows Media Player URLs / Files and
	'#						Common Dialog Open / Save As MRUs
    '#--------------------------------------------------------------------------
	Sub DeleteTempFiles()
		Const DeleteReadOnly = True
		intTotalSize = 0
		i = 0
		intProgTotal = 6
		intProgMult = 100 / intProgTotal
		intProgDone = 0		
		
		If Reachable(strPC) Then
			CleanUninstPrompt = MsgBox ("Would you like to clean Hotfix Uninstallers " & _
			"greater than 90 days old?",vbQuestion+vbYesNo, "Clean Temp Files")
			document.body.style.cursor = "wait"
			PauseScript(50)
		
			Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
			strPC & "\root\default:StdRegProv")
				
			intProgDone = intProgDone + 1	'1
			UpdateProgressBar intProgMult,intProgDone,intProgTotal

			strPCName = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
			strUserProf = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
			strRoot = GetRoot()
			strDelTempFile = strRoot & "\deletetempfiles.bat"
			strDelTempFile2 = strDelTempFile2 = strTemp & "\SKB\deletedfiles-" & _
			strFileDate & "-" & strPC & ".txt"
			Set objFile = objFSO.CreateTextFile(strDelTempFile, True)
			objFile.WriteLine "@ECHO OFF"
			objFile.WriteLine "cls"			
			objFile.Close
			Set objFile = objFSO.CreateTextFile(strDelTempFile2, True)
			objFile.WriteLine "Deleted Files"
			objFile.WriteLine ""
			objFile.WriteLine "Computer Name: " & UCase(strPC)
			objFile.WriteLine "Date: " & Date
			objFile.WriteLine ""
			objFile.Close
			objShell.Run "%COMSPEC% /c cacls " & strDelTempFile & _
			" /c /e /g " & strRemoteLoggedOn & ":F", 2, True
			
			intProgDone = intProgDone + 1	'2
			UpdateProgressBar intProgMult,intProgDone,intProgTotal
			
			'#--------------------------------------------------------------------------
			'#	Delete all the Hotfix uninstallers (if required)
			'#--------------------------------------------------------------------------
			strWindows = strRoot & "\windows"
			If objFSO.FolderExists(strRoot & "\documents and settings") Then
				strProfileRoot = strRoot & "\documents and settings"
				ElseIf objFSO.FolderExists(strRoot & "\users") Then
					strProfileRoot = strRoot & "\users"
			End If
			If CleanUninstPrompt = vbYes Then
				dtmEarliestDate = DateAdd ("d", -90, Date)
				Set strHoldWinXP = objFSO.getFolder(strWindows)
				Set arrWinXPFolders = strHoldWinXP.SubFolders
				intFolderCount = 0
				For Each objItem in arrWinXPFolders
					If InStr (1, objItem.Name, "$NT", 1) > 0 _
					AND objItem.DateCreated < dtmEarliestDate Then
						intFolderCount = intFolderCount + 1
					End If
				Next
				If intFolderCount > 0 Then
					intProgDone2 = 0
					intProgTotal2 = intFolderCount * 3
					intProgMult2 = 100 / intProgTotal2
					For Each objItem in arrWinXPFolders
						If InStr (1, objItem.Name, "$NT", 1) > 0 _
						AND objItem.DateCreated < dtmEarliestDate Then
							On Error Resume Next
							intTotalSize = intTotalSize + GetSize(strWindows & "\" & objItem.Name)
							intProgDone2 = intProgDone2 + 1
							UpdateProgressBar intProgMult2,intProgDone2,intProgTotal2
							objFSO.DeleteFolder(strWindows & "\" & objItem.Name), DeleteReadOnly
							Set objFile = objFSO.OpenTextFile(strDelTempFile2, ForAppending, True)
							objFile.WriteLine ""
							objFile.WriteLine "Deleting Old Hotfix Uninstallers"
							objFile.WriteLine "---------------------------"
							objFile.Close
							intProgDone2 = intProgDone2 + 1
							UpdateProgressBar intProgMult2,intProgDone2,intProgTotal2
							intTotalSize = intTotalSize - GetSize(strWindows & "\" & objItem.Name)
							intProgDone2 = intProgDone2 + 1
							UpdateProgressBar intProgMult2,intProgDone2,intProgTotal2
						End If
					Next
				End If
			End If
			
			intProgDone = intProgDone + 1	'3
			UpdateProgressBar intProgMult,intProgDone,intProgTotal
			
			On Error GoTo 0
			
			'#--------------------------------------------------------------------------
			'#	Delete all the files within the Profile folders
			'#--------------------------------------------------------------------------
			intFolderCount = 0
			Set objProfileFolder = objFSO.GetFolder(strProfileRoot)
			For Each objItem In objProfileFolder.Subfolders
				If LCase(objItem.Name) <> "all users" _
				AND LCase(objItem.Name) <> "localservice" _
				AND LCase(objItem.Name) <> "default user" _
				AND LCase(objItem.Name) <> "networkservice" Then
					intFolderCount = intFolderCount + 1
				End If
			Next
			intProgDone3 = 0
			intProgTotal3 = intFolderCount * 5
			intProgMult3 = 100 / intProgTotal3
			For Each objItem In objProfileFolder.SubFolders
				If LCase(objItem.Name) <> "all users" _
				AND LCase(objItem.Name) <> "localservice" _
				AND LCase(objItem.Name) <> "default user" _
				AND LCase(objItem.Name) <> "networkservice" Then
					strProfile = objItem
					Set objFile = objFSO.OpenTextFile(strDelTempFile2, ForAppending, True)
					objFile.WriteLine ""
					objFile.WriteLine "Profile: " & UCase(objItem.Name)
					objFile.WriteLine ""
					objFile.Close
					strTempFilesXP = strProfile & "\Local Settings\Temp"
					strTempFiles7 = strProfile & "\AppData\Local\Temp"
					
					If objFSO.FolderExists(strTempFilesXP) Then
						strTempFiles = strTempFilesXP
						ElseIf objFSO.FolderExists(strTempFiles7) Then
							strTempFiles = strTempFiles7
					End If
					
					strHistoryXP = strProfile & "\Local Settings\History"
					strHistory7 = strProfile & "\AppData\Local\Microsoft\Windows\History"
					
					If objFSO.FolderExists(strHistoryXP) Then
						strHistory = strHistoryXP
						ElseIf objFSO.FolderExists(strHistory7) Then
							strHistory = strHistory7
					End If
					
					strCookiesXP = strProfile & "\Cookies"
					strCookies7 = strProfile & "\AppData\Roaming\Microsoft\Windows\Cookies"
					
					If objFSO.FolderExists(strCookiesXP) Then
						strCookies = strCookiesXP
						ElseIf objFSO.FolderExists(strCookies7) Then
							strCookies = strCookies7
					End If
					
					strTempInternetFilesXP = strProfile & "\Local Settings\Temporary Internet Files"
					strTempInternetFiles7 = strProfile & "\AppData\Local\Microsoft\Windows\Temporary Internet Files"
					
					If objFSO.FolderExists(strTempInternetFilesXP) Then
						strTempInternetFiles = strTempInternetFilesXP
						ElseIf objFSO.FolderExists(strTempInternetFiles7) Then
							strTempInternetFiles = strTempInternetFiles7
					End If
					
					strFirefoxProfileXP = strProfile & "\Local Settings\Application Data\Mozilla\Firefox\Profiles"
					strFirefoxProfile7 = strProfile & "\AppData\Local\Mozilla\Firefox\Profiles"
					
					If objFSO.FolderExists(strFirefoxProfileXP) Then
						strFirefoxProfile = strFirefoxProfileXP
						ElseIf objFSO.FolderExists(strFirefoxProfile7) Then
							strFirefoxProfile = strFirefoxProfile7
					End If
					
					strHistoryContentXP = strProfile & "\Local Settings\Temporary Internet Files\Content.IE5\"
					
					If objFSO.FolderExists(strHistoryContentXP) Then
						strHistoryContent = strHistoryContentXP
						Else
							strHistoryContent = ""
					End If
					
					intTotalSize = intTotalSize + GetSize(strHistory)
					intTotalSize = intTotalSize + GetSize(strFirefoxProfile)
					intTotalSize = intTotalSize + GetSize(strCookies)
					intProgDone3 = intProgDone3 + 1
					UpdateProgressBar intProgMult3,intProgDone3,intProgTotal3
					intTotalSize = intTotalSize + GetSize(strTempFiles)
					intTotalSize = intTotalSize + GetSize(strTempInternetFiles)
					intProgDone3 = intProgDone3 + 1
					UpdateProgressBar intProgMult3,intProgDone3,intProgTotal3
					Set objFile = objFSO.OpenTextFile(strDelTempFile2, ForAppending, True)
					objFile.WriteLine ""
					objFile.WriteLine "Deleting Temp Folder"
					objFile.WriteLine "---------------------------"
					objFile.Close
					DeleteProfFolderContents strTempFiles,strDelTempFile
					Set objFile = objFSO.OpenTextFile(strDelTempFile2, ForAppending, True)
					objFile.WriteLine ""
					objFile.WriteLine "Deleting IE History Folder"
					objFile.WriteLine "---------------------------"
					objFile.Close
					DeleteProfFolderContents strHistory,strDelTempFile
					intProgDone3 = intProgDone3 + 1
					UpdateProgressBar intProgMult3,intProgDone3,intProgTotal3
					Set objFile = objFSO.OpenTextFile(strDelTempFile2, ForAppending, True)
					objFile.WriteLine ""
					objFile.WriteLine "Deleting Cookies Folder"
					objFile.WriteLine "---------------------------"
					objFile.Close
					DeleteProfFolderContents strCookies,strDelTempFile
					Set objFile = objFSO.OpenTextFile(strDelTempFile2, ForAppending, True)
					objFile.WriteLine ""
					objFile.WriteLine "Deleting Firefox Cache Folder"
					objFile.WriteLine "---------------------------"
					objFile.Close
					DeleteProfFolderContents strFirefoxProfile,strDelTempFile
					Set objFile = objFSO.OpenTextFile(strDelTempFile2, ForAppending, True)
					objFile.WriteLine ""
					objFile.WriteLine "Deleting History Content"
					objFile.WriteLine "---------------------------"
					objFile.Close
					DeleteProfFolderContents strHistoryContent,strDelTempFile
					intProgDone3 = intProgDone3 + 1
					UpdateProgressBar intProgMult3,intProgDone3,intProgTotal3
					Set objFile = objFSO.OpenTextFile(strDelTempFile2, ForAppending, True)
					objFile.WriteLine ""
					objFile.WriteLine "Deleting Temporary Internet Files Folder"
					objFile.WriteLine "---------------------------"
					objFile.Close
					DeleteProfFolderContents strTempInternetFiles,strDelTempFile
					intTotalSize = intTotalSize - (GetSize(strHistory) + _
					GetSize(strTempFiles) + GetSize(strTempInternetFiles) + _
					GetSize(strCookies) + GetSize(strFirefoxProfile))
					intProgDone3 = intProgDone3 + 1
					UpdateProgressBar intProgMult3,intProgDone3,intProgTotal3
				End If
			Next
			
			strHistoryContent7 = strProfileRoot & _
			"\Owner\AppData\Local\Microsoft\Windows\Temporary Internet Files\Content.IE5"
			If objFSO.FolderExists(strHistoryContent7) Then
				Set objFile = objFSO.OpenTextFile(strDelTempFile2, ForAppending, True)
				objFile.WriteLine ""
				objFile.WriteLine "History Content"
				objFile.WriteLine "---------------------------"
				objFile.Close
				DeleteProfFolderContents strHistoryContent7,strDelTempFile
			End If
			
			intProgDone = intProgDone + 1	'4
			UpdateProgressBar intProgMult,intProgDone,intProgTotal
			
			'#--------------------------------------------------------------------------
			'#	Clear the MRU info from the registry
			'#--------------------------------------------------------------------------
			Set objFile = objFSO.OpenTextFile(strDelTempFile2, ForAppending, True)
			objFile.WriteLine ""
			objFile.WriteLine "Clearing MRU Lists"
			objFile.WriteLine "---------------------------"
			objFile.Close
			strKeyPath = (strRemoteSID & _
			"\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\RunMRU")
			DeleteAllKeys strKeyPath, "String"
			Set objFile = objFSO.OpenTextFile(strDelTempFile2, ForAppending, True)
			objFile.WriteLine "Cleared RunMRU"
			objFile.Close
			strKeyPath = (strRemoteSID & _
			"\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs")
			DeleteAllSubkeysUSERS strKeyPath
			DeleteAllKeys strKeyPath, "Binary"
			Set objFile = objFSO.OpenTextFile(strDelTempFile2, ForAppending, True)
			objFile.WriteLine "Cleared Recent Docs"
			objFile.Close
			strKeyPath = (strRemoteSID & _
			"\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\LastKey")
			DeleteAllKeys strKeyPath, "Regedit"
			Set objFile = objFSO.OpenTextFile(strDelTempFile2, ForAppending, True)
			objFile.WriteLine "Cleared Regedit Last Key"
			objFile.Close
			strKeyPath = (strRemoteSID & _
			"\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Paint\Recent File List")
			DeleteAllSubkeysUSERS strKeyPath
			Set objFile = objFSO.OpenTextFile(strDelTempFile2, ForAppending, True)
			objFile.WriteLine "Cleared Paint Recent File List"
			objFile.Close
			strKeyPath = (strRemoteSID & _
			"\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Wordpad\Recent File List")
			DeleteAllSubkeysUSERS strKeyPath
			Set objFile = objFSO.OpenTextFile(strDelTempFile2, ForAppending, True)
			objFile.WriteLine "Cleared Wordpad Recent File List"
			objFile.Close
			strKeyPath = (strRemoteSID & _
			"\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\LastVisitedMRU")
			DeleteAllKeys strKeyPath, "String"
			strKeyPath = (strRemoteSID & _
			"\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\OpenSaveMRU")
			DeleteAllSubkeysUSERS strKeyPath
			Set objFile = objFSO.OpenTextFile(strDelTempFile2, ForAppending, True)
			objFile.WriteLine "Cleared Common Dialog MRUs"
			objFile.Close
			strKeyPath = (strRemoteSID & _
			"\SOFTWARE\Microsoft\MediaPlayer\Player\RecentFileList")
			DeleteAllSubkeysUSERS strKeyPath
			strKeyPath = (strRemoteSID & _
			"\SOFTWARE\Microsoft\MediaPlayer\Player\RecentURLList")
			DeleteAllSubkeysUSERS strKeyPath
			Set objFile = objFSO.OpenTextFile(strDelTempFile2, ForAppending, True)
			objFile.WriteLine "Cleared Windows Media Player Recent File / URL List"
			objFile.Close
			
			intProgDone = intProgDone + 1	'5
			UpdateProgressBar intProgMult,intProgDone,intProgTotal

			Set objFile = objFSO.OpenTextFile(strDelTempFile, ForAppending, True)
			objFile.WriteLine "del %0"		
			objFile.Close
			strKeyPath = (strRemoteSID & _
			"\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce")
			strValueName = "DeleteTempFiles"
			strValue = "c:\deletetempfiles.bat"
			objReg.CreateKey HKEY_USERS,strKeyPath
			objReg.SetStringValue HKEY_USERS,strKeyPath,strValueName,strValue

			strDeleted = ConvertToDiskSize(intTotalSize)
			intProgDone = intProgDone + 1	'6
			UpdateProgressBar intProgMult,intProgDone,intProgTotal
			document.body.style.cursor = "default"
			PauseScript(50)
			TempFilePrompt = MsgBox("Temporary files have been deleted for " & UCase(strPC) & _
			vbCrLf & vbCrLf & strDeleted & " of data was deleted" & vbCrLf & vbCrLf & _
			"After the PC is restarted a script will run to delete any files that were " & _
			"open at the time" & vbCrLf & vbCrLf & "Would you like to view a file showing the " & _
			"deleted files and folders?",vbInformation+vbYesNo, "Clean Temp Files")
			If TempFilePrompt = vbYes Then
				objShell.Run strDelTempFile2
			End If
			If NOT objFSO.FolderExists(strTemp & "\SKB") Then
				Set objFolder = objFSO.CreateFolder(strTemp & "\SKB")
			End If
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	DeleteAllKeys(strKeyPath, strType)
    '#	PURPOSE........:	Deletes all the Reg keys in the specfied keypath
    '#	ARGUMENTS......:	strKeyPath = the full path to the reg key to remove
	'#						strType = the key type (eg. "String") for the
	'#						MRUList value
    '#	EXAMPLE........:	DeleteAllKeys("\SOFTWARE\Microsoft", "String")
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub DeleteAllKeys(strKeyPath, strType)
		Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\default:StdRegProv")
		For intAscVal = 97 to 122
			strValueName = (Chr(intAscVal))
			objReg.DeleteValue HKEY_USERS, strKeyPath, strValueName
		Next
		For i = 0 to 28
			strValueName = i
			objReg.DeleteValue HKEY_USERS, strKeyPath, strValueName
		Next
		strValue = ""
		Select Case strType
			Case "String"
				strValueName = "MRUList"
				objReg.SetStringValue HKEY_USERS, strKeyPath, strValueName,strValue
			Case "Binary"
				arrValue = Array()
				strValueName = "MRUListEx"
				objReg.SetBinaryValue HKEY_USERS, strKeyPath, strValueName,arrValue
			Case "Regedit"
				strValueName = "LastKey"
				objReg.SetStringValue HKEY_USERS, strKeyPath, strValueName,strValue
		End Select
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	DeleteAllSubkeysUSERS(strKeyPath)
    '#	PURPOSE........:	Deletes all the subkeys in the specfied Reg keypath
    '#	ARGUMENTS......:	strKeyPath = the full path to the reg key to remove
    '#	EXAMPLE........:	DeleteAllSubkeysUSERS("\SOFTWARE\Microsoft")
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------	
	Sub DeleteAllSubkeysUSERS(strKeyPath)
		Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\default:StdRegProv")
		objReg.EnumKey HKEY_USERS, strKeyPath, arrSubkeys 
		If IsArray(arrSubkeys) Then 
			For Each objItem In arrSubkeys 
				objReg.DeleteKey HKEY_USERS, strKeyPath & "\" & objItem
			Next 
		End If 
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	DeleteProfFolderContents(strFolder, strDelTmpFile)
    '#	PURPOSE........:	Deletes all the contents of the specified folder
    '#	ARGUMENTS......:	strFolder = full path to the folder
	'#						strDelTmpFile = full path to deletetempfiles.bat
	'#						file
    '#	EXAMPLE........:	DeleteProfFolderContents("c:\Temp","c:\deletetempfiles.bat")
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub DeleteProfFolderContents(strFolder,strDelTmpFile)
		strDelTempFile2 = strTemp & "\SKB\deletedfiles-" & strFileDate & _
			"-" & strPC & ".txt"
		On Error Resume Next
		Set objFolder=objFSO.GetFolder(strFolder)
		If Err.Number <> 0 Then
			Err.Clear
			Exit Sub
		End If
		For Each objItem In objFolder.SubFolders
			strCurrFolder = objItem
			objItem.Delete True
			If Err.Number = 0 Then
				Set objFile = objFSO.OpenTextFile(strDelTempFile2, ForAppending, True)
				objFile.WriteLine strCurrFolder
				objFile.Close
				Else
					Err.Clear
					
					objShell.Run "%COMSPEC% /c del /f /q " & Chr(34) & objItem & Chr(34),0,True
					If objFSO.FolderExists(strCurrFolder) Then
						Set objFile = objFSO.OpenTextFile(strDelTmpFile, ForAppending, True)
						objFile.WriteLine "del /f /q " & Chr(34) & objItem & Chr(34)
						objFile.Close
					End If
				End If
			Err.Clear
		Next
		For Each objItem In objFolder.Files
			strCurrFile = objItem
			objItem.Delete True
			If Err.Number = 0 Then
				Set objFile = objFSO.OpenTextFile(strDelTempFile2, ForAppending, True)
				objFile.WriteLine strCurrFile
				objFile.Close
				Else
					Err.Clear
					objShell.Run "%COMSPEC% /c del /f /q " & Chr(34) & objItem & Chr(34),0,True
					If objFSO.FileExists(strCurrFile) Then
						Set objFile = objFSO.OpenTextFile(strDelTmpFile, ForAppending, True)
						objFile.WriteLine "del /f /q " & Chr(34) & objItem & Chr(34)
						objFile.Close
					End If
			End If
			Err.Clear
		Next
	End Sub

	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	EnableDisableIEProxy()
    '#	PURPOSE........:	Enables / Disables the IE proxy
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	Will only work if a user is currently logged onto
	'#						PC as this is set in HKCU
    '#--------------------------------------------------------------------------
	Sub EnableDisableIEProxy()
		WaitMessage.InnerHTML = "<hr><p>IE Proxy is <span id=IEProxyStatus></span>&nbsp;on " & _
		strPC & "<p><input id=""RunButton"" class=""button"" type=""button"" " & _
		"name=""btnIEProxy"" onclick=""EnableDisableIEProxy2()"">"
	
		Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\default:StdRegProv") 
		
		strKeyPath = (strRemoteSID & _
		"\Software\Microsoft\Windows\CurrentVersion\Internet Settings")
		strValueName = "ProxyEnable"

		objReg.GetDWORDValue HKEY_USERS,strKeyPath,strValueName,dwValue

		If dwValue = 1 Then
			IEProxyStatus.InnerHTML = "<b>ENABLED</b>"
			IEProxyStatus.style.color = "green"
			btnIEProxy.Value = "Disable"
			btnIEProxy.Title = "Disable IE Proxy on PC"
			ElseIf dwValue = 0 Then
				IEProxyStatus.InnerHTML = "<b>DISABLED</b>"
				IEProxyStatus.style.color = "red"
				btnIEProxy.Value = "Enable"
				btnIEProxy.Title = "Enable IE Proxy on PC"
		End If
	End Sub
	
	Sub EnableDisableIEProxy2()
		Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\default:StdRegProv") 
		
		strKeyPath = (strRemoteSID & _
		"\Software\Microsoft\Windows\CurrentVersion\Internet Settings")
		strValueName = "ProxyEnable"
		
		If InStr(UCase(IEProxyStatus.InnerHTML), "ENABLED") > 0 Then
			dwValue = 0
			objReg.SetDWORDValue HKEY_USERS,strKeyPath,strValueName,dwValue 
			MsgBox "IE Proxy is now DISABLED on " & strPC & ".",vbInformation, "Disable IE Proxy"
			Else
				dwValue = 1
				objReg.SetDWORDValue HKEY_USERS,strKeyPath,strValueName,dwValue 
				MsgBox "IE Proxy is now ENABLED on " & strPC & ".",vbInformation, "Enable IE Proxy"
		End If
		EnableDisableIEProxy()
	End Sub

	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	EnableDisableRDP()
    '#	PURPOSE........:	Enables / Disables the RDP
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub EnableDisableRDP()
		WaitMessage.InnerHTML = "<hr><p>Remote Desktop is <span id=RDPStatus></span>&nbsp;on " & _
		strPC & "<p><input id=""RunButton"" class=""button"" type=""button"" " & _
		"name=""btnRDP"" onclick=""EnableDisableRDP2()"">"
	
		Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
		strPC & "\root\default:StdRegProv")

		strKeyPath = "SYSTEM\CurrentControlSet\Control\Terminal Server"
		strValueName = "fDenyTSConnections"

		objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue

		If dwValue = 1 Then
			RDPStatus.InnerHTML = "<b>DISABLED</b>"
			RDPStatus.style.color = "red"
			btnRDP.Value = "Enable"
			btnRDP.Title = "Enable Remote Desktop on PC"
			ElseIf dwValue = 0 then
				RDPStatus.InnerHTML = "<b>ENABLED</b>"
				RDPStatus.style.color = "green"
				btnRDP.Value = "Disable"
				btnRDP.Title = "Disable Remote Desktop on PC"
		End If
	End Sub
	
	Sub EnableDisableRDP2()
		Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\default:StdRegProv") 
		
		strKeyPath = "SYSTEM\CurrentControlSet\Control\Terminal Server"
		strValueName = "fDenyTSConnections"
		
		If InStr(UCase(RDPStatus.InnerHTML), "ENABLED") > 0 Then
			dwValue = 1
			objReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue 
			MsgBox "Remote Desktop is now DISABLED on " & strPC & ".",vbInformation, "Disable RDP"
			Else
				dwValue = 0
				objReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue 
				MsgBox "Remote Desktop is now ENABLED on " & strPC & ".",vbInformation, "Enable RDP"
		End If
		EnableDisableRDP()
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	GetMSProductKeys()
    '#	PURPOSE........:	Retrieves the Microsoft Product Keys for Windows /
	'#						Office XP / Office 2003 / Office 2007 / Exchange
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	GetMSProductKeys()
    '#	NOTES..........:	Props go to Parabellum for this Sub, although have
	'#						also added 64-bit support and  expanded on initial 
	'#						idea
    '#--------------------------------------------------------------------------
	Sub GetMSProductKeys()
		strHTML = "<hr><div style=""text-align:left;"">" & _
		"Which product key would you like to retrieve?" & _
		"<p><input type=""radio"" name=""MSProdKeyType"" value=""1"" " & _
		"title=""Retrieve Microsoft Windows product key"">Microsoft Windows<br>" & _
		"<input type=""radio"" name=""MSProdKeyType"" value=""2"" " & _
		"title=""Retrieve Microsoft Office product key"">Microsoft Office<br>" & _
		"<input type=""radio"" name=""MSProdKeyType"" value=""3"" " & _
		"title=""Retrieve Microsoft Exchange product key"">Microsoft Exchange<p>" & _
		"<input type=""radio"" name=""MSProdKeyType"" value=""4"" " & _
		"title=""Retrieve all Microsoft product keys"">All<p>" & _
		"<input id=""RunButton"" class=""button"" type=""button"" value=""Retrieve"" " & _
		"name=""btnMSProdKey"" onclick=""GetMSProductKeys2()"" " & _
		"title=""Retrieve selected product key from PC"">" & _
		"</div>"
		WaitMessage.InnerHTML = strHTML
	End Sub
	
	Sub GetMSProductKeys2()
		booOfficeMsg = False
		booMSOffice = False
		Const SEARCH_KEY = "DigitalProductID"
		Dim arrSubKeys(4,1)
		strWinVer = CheckWinArchitecture()
		intMSSPan = 0
		j = 0
				
		Select Case strWinVer
			Case "64-bit"
				arrSubKeys(0,0) = "Microsoft Windows Product Key"
				arrSubKeys(0,1) = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
				arrSubKeys(1,0) = "Microsoft Office 2003"
				arrSubKeys(1,1) = "SOFTWARE\Wow6432Node\Microsoft\Office\11.0\Registration"
				arrSubKeys(2,0) = "Microsoft Office XP"
				arrSubKeys(2,1) = "SOFTWARE\Wow6432Node\Microsoft\Office\10.0\Registration"
				arrSubKeys(3,0) = "Microsoft Office 2007"
				arrSubKeys(3,1) = "SOFTWARE\Wow6432Node\Microsoft\Office\12.0\Registration"
				arrSubKeys(4,0) = "Microsoft Exchange Product Key"
				arrSubKeys(4,1) = "SOFTWARE\Microsoft\Exchange\Setup"				
			Case Else
				arrSubKeys(0,0) = "Microsoft Windows Product Key"
				arrSubKeys(0,1) = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
				arrSubKeys(1,0) = "Microsoft Office 2003"
				arrSubKeys(1,1) = "SOFTWARE\Microsoft\Office\11.0\Registration"
				arrSubKeys(2,0) = "Microsoft Office XP"
				arrSubKeys(2,1) = "SOFTWARE\Microsoft\Office\10.0\Registration"
				arrSubKeys(3,0) = "Microsoft Office 2007"
				arrSubKeys(3,1) = "SOFTWARE\Microsoft\Office\12.0\Registration"
				arrSubKeys(4,0) = "Microsoft Exchange Product Key"
				arrSubKeys(4,1) = "SOFTWARE\Microsoft\Exchange\Setup"
		End Select

		Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\default:StdRegProv") 
	
		For Each objButton in MSProdKeyType
			If objButton.Checked Then MSPrompt = objButton.Value
		Next
		Select Case MSPrompt
			Case 1
				a = 0
				z = 0
			Case 2
				a = 1
				z = 3
			Case 3
				a = 4
				z = 4
			Case 4
				a = 0
				z = 4
				strHTML = "<hr><p><div style=""text-align:right;"">"
				strHTML = strHTML & "<select name=""MSProdKeyExport"" " & _
				"title=""Export the product keys list"" onChange=""ExportMSProdKeyDetails()"">"
				strHTML = strHTML & "<option value=""0"">Export to:</option>"
				strHTML = strHTML & "<option value=""1"" title=""Export the product keys to a Comma " & _
				"Seperated Values (csv) file"")>Export to csv</option>"
				strHTML = strHTML & "<option value=""2"" title=""Export the product keys to a formatted Excel " & _
				"(xls) spreadsheet"">Export to xls</option>"
				strHTML = strHTML & "<option value=""3"" title=""Export the product keys to a Web " & _
				"page (html) file"">Export to html</option>"
				strHTML = strHTML & "<option value=""4"" title=""Export the product keys to a Text " & _
				"(txt) file"">Export to txt</option>"
				strHTML = strHTML & "</select>"
				strHTML = strHTML & "<div style=""overflow:auto;width:100%;height:125;" & _
				"border:1px solid #a5a5a5;border-right:0px;padding:0px;margin: 0px"">" 
				strHTML = strHTML & "<table class=""prodkeystable"">"
		End Select
 
		For x = a to z
			Select Case x
				Case 0
					strProduct = "Microsoft Windows"
				Case 4
					strProduct = "Microsoft Exchange"
			End Select
					
			objReg.GetBinaryValue HKEY_LOCAL_MACHINE, arrSubKeys(x,1), _
			SEARCH_KEY, arrDPIDBytes
			If Not IsNull(arrDPIDBytes) Then
				If x = 0 Then
					Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
					strPC & "\root\cimv2")
					Set colOS = objWMIService.ExecQuery _
						("Select * from Win32_OperatingSystem")
				
					For Each objItem In colOS
						strOS = objItem.Caption
						strBuild = objItem.BuildNumber
						strSerial = objItem.SerialNumber
						strRegistered = objItem.RegisteredUser
					Next
					strCopyProdKey = DecodeKey(arrDPIDBytes)
					intMSSPan = intMSSPan + 1
					
					If MSPrompt <> 4 Then
						WaitMessage.InnerHTML = "<hr><b><span id=""MSProd" & intMSSpan & """>" & strOS & "</span></b><p>" & _
						"<table><tr><td><b>Build Number: </b></td><td>" & strBuild & "</td></tr>" & _
						"<tr><td><b>PID: </b></td><td>" & strSerial & "</td></tr>" & _
						"<tr><td><b>Registered to: </b></td><td>" & strRegistered & "</td></tr>" & _
						"<tr><td><b>Windows Product Key: &nbsp;</b></td><td>" & _
						"<span id=""MSProdKey" & intMSSpan & """ style=""font-family:'Courier New';font-weight:bold;"">" & _
						strCopyProdKey & "</span></td></tr>" & _
						"<tr><td>&nbsp;</td><td><input id=""RunButton"" class=""button"" type=""button"" " & _
						"value=""Copy Key"" name=""btnCopyProdKey"" onclick=""CopyKeyInfo(" & intMSSpan & ")"" " & _
						"title=""Copy product key to clipboard""></td></tr></table>"
						Else
							strHTML = strHTML & "<tr><td id=""MSProd" & intMSSpan & """ style=width:""40%"">" & strOS & _
							"</td><th id=""MSProdKey" & intMSSpan & """>" & _
							strCopyProdKey & "</th><td style=width:""10%"">" & _
							"<input id=""RunButton"" class=""button"" type=""button"" " & _
							"value=""Copy Key"" name=""btnCopyProdKey"" onclick=""CopyKeyInfo(" & intMSSpan & ")"" " & _
							"title=""Copy product key to clipboard""></td></tr>"
					End If
					Else
						strCopyProdKey = DecodeKey(arrDPIDBytes)
						intMSSPan = intMSSPan + 1
						If MSPrompt <> 4 Then
							WaitMessage.InnerHTML = "<hr><p><b><span id=""MSProd" & intMSSpan & """ style=width:""40%"">" & arrSubKeys(x,0) & _
							"</span></b><p>" & _
							"<span id=""MSProdKey" & intMSSpan & """ style=""font-family:'Courier New';font-weight:bold;"">" & _
							strCopyProdKey & "</span>" & _
							"<p><input id=""RunButton"" class=""button"" type=""button"" value=""Copy Key"" " & _
							"name=""btnCopyProdKey"" onclick=""CopyKeyInfo(" & intMSSpan & ")"" " & _
							"title=""Copy product key to clipboard"">"
							Else
								strHTML = strHTML & "<tr><td id=""MSProd" & intMSSpan & """>" & _
								arrSubKeys(x,0)  & "</td><th id=""MSProdKey" & intMSSpan & """>" & _
								strCopyProdKey & "</th><td style=width:""10%"">" & _
								"<input id=""RunButton"" class=""button"" type=""button"" " & _
								"value=""Copy Key"" name=""btnCopyProdKey"" onclick=""CopyKeyInfo(" & intMSSpan & ")"" " & _
								"title=""Copy product key to clipboard""></td></tr>"
						End If
				End If
				Else
					objReg.EnumKey HKEY_LOCAL_MACHINE, arrSubKeys(x,1), arrGUIDKeys
					If Not IsNull(arrGUIDKeys) Then
						For Each GUIDKey In arrGUIDKeys
							objReg.GetBinaryValue HKEY_LOCAL_MACHINE, _
							arrSubKeys(x,1) & "\" & GUIDKey, SEARCH_KEY, arrDPIDBytes
							If Not IsNull(arrDPIDBytes) Then
								If x > 0 AND x < 4 Then 
									booMSOffice = True
									j = j + 1
								End If
								strCopyProdKey = DecodeKey(arrDPIDBytes)
								intMSSPan = intMSSPan + 1
								If MSPrompt <> 4 Then
									If j > 1 AND x > 0 AND x < 4 Then
										strHTML = Replace(strHTML, "<p>", "<br>")
										strHTML = strHTML & "<p><b><span id=""MSProd" & _
										intMSSpan & """ style=width:""40%"">" & arrSubKeys(x,0) & _
										"</span></b><br><span id=""MSProdKey" & intMSSpan & _
										""" style=""font-family:'Courier New';font-weight:bold;"">" & _
										strCopyProdKey & "</span>" & _
										"<br><input id=""RunButton"" class=""button"" type=""button"" " & _
										"value=""Copy Key"" name=""btnCopyProdKey"" " & _
										"onclick=""CopyKeyInfo(" & intMSSpan & ")"" " & _
										"title=""Copy product key to clipboard"">"
										Else
											strHTML = "<hr><p><b><span id=""MSProd" & _
											intMSSpan & """ style=width:""40%"">" & arrSubKeys(x,0) & _
											"</span></b><p>" & _
											"<span id=""MSProdKey" & intMSSpan & _
											""" style=""font-family:'Courier New';font-weight:bold;"">" & _
											strCopyProdKey & "</span>" & _
											"<p><input id=""RunButton"" class=""button"" type=""button"" " & _
											"value=""Copy Key"" name=""btnCopyProdKey"" " & _
											"onclick=""CopyKeyInfo(" & intMSSpan & ")"" " & _
											"title=""Copy product key to clipboard"">"
									End If
									Else
										strHTML = strHTML & "<tr><td id=""MSProd" & intMSSpan & """>" & _
										arrSubKeys(x,0)  & "</td><th id=""MSProdKey" & intMSSpan & """>" & _
										strCopyProdKey & "</th><td  style=width:""10%"">" & _
										"<input id=""RunButton"" class=""button"" type=""button"" " & _
										"value=""Copy Key"" name=""btnCopyProdKey"" onclick=""CopyKeyInfo(" & intMSSpan & ")"" " & _
										"title=""Copy product key to clipboard""></td></tr>"
								End If
							End If
						Next
						WaitMessage.InnerHTML = strHTML
						Else
							If x = 0 OR x = 4 Then
								If MSPrompt <> 4 Then
									For Each objButton in MSProdKeyType
										objButton.Checked = False
									Next
									MsgBox strProduct & " cannot be found on " & _
									strPC, vbInformation, "PC Management Utility"
								End If
							End If
							If x = 3 AND booMSOffice = False AND MSPrompt <> 4 Then
								For Each objButton in MSProdKeyType
									objButton.Checked = False
								Next
								MsgBox "Microsoft Office cannot be found on " & _
								strPC, vbInformation, _
								"PC Management Utility"
							End If
					End If
			End If
		Next
		If MSPrompt = 4 Then
			strHTML = strHTML & "</table></div></div>"
			WaitMessage.InnerHTML = strHTML
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ExportMSProdKeyDetails()
    '#	PURPOSE........:	Export the Microsoft Product Keys
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ExportMSProdKeyDetails()
		On Error Resume Next
		
		strProduct1 = MSProd1.InnerHTML
		strKey1 = MSProdKey1.InnerHTML
		strProduct2 = MSProd2.InnerHTML
		strKey2 = MSProdKey2.InnerHTML
		strProduct3 = MSProd3.InnerHTML
		strKey3 = MSProdKey3.InnerHTML
		strProduct4 = MSProd4.InnerHTML
		strKey4 = MSProdKey4.InnerHTML
		strProduct5 = MSProd5.InnerHTML
		strKey5 = MSProdKey5.InnerHTML
		strProduct6 = MSProd6.InnerHTML
		strKey6 = MSProdKey6.InnerHTML

		Select Case MSProdKeyExport.Value
			Case 1
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\MSProdKeyDetails" & strPC & ".csv",True)
				objFile.WriteLine "Microsoft Product Keys on " & strPC
				objFile.WriteLine ""
				objFile.WriteLine "Product,Product Key"
			Case 2
				Const xlContinuous = 1
				Const xlThin = 2
				Const xlAutomatic = -4105
				
				strExcelPath = objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe\")
			   
				If strExcelPath = "" Then
					MsgBox "Unable to export. Excel does not appear to be installed.", vbExclamation, "PC Management Utility"
					Exit Sub
				End If
				
				Set objExcel = CreateObject("Excel.Application")
				objExcel.Visible = False
				Set objWorkBook = objExcel.WorkBooks.Add
				Set objWorksheet = objWorkbook.Worksheets(1)
				objExcel.DisplayAlerts = False
				For i = 1 to 3
					objWorkbook.Worksheets(2).Delete
				Next
				objExcel.DisplayAlerts = True
				objWorksheet.Name = "MS Product Keys"
				
				objWorkSheet.Cells(1, 1) = "Microsoft Product Keys on " & strPC
				intStartRow = 4
				
				objWorkSheet.Cells(3, 1) = "Product"
				objWorkSheet.Cells(3, 2) = "Product Key"
			Case 3
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\MSProdKeyDetails" & strPC & ".htm",True)
				objFile.WriteLine "<style type=""text/css"">"
				objFile.WriteLine "body{background-color:#CEF0FF;}"
				objFile.WriteLine "table.export{border-width:1px;border-spacing:1px;border-style:solid;border-color:gray;border-collapse:collapse;}"
				objFile.WriteLine "table.export th{border-width:1px;padding:1px;border-style:solid;border-color:gray;padding:2px 7px 2px 7px;}"
				objFile.WriteLine "table.export td{border-width:1px;padding:1px;border-style:dotted;border-color:gray;padding:2px 7px 2px 7px;}"
				objFile.WriteLine "</style>"
				objFile.WriteLine "<div style=""font-weight:bold;""><a name =""top"">Microsoft Product Keys on " & _
				strPC & "</a><p>"
				objFile.WriteLine "</div>"
				objFile.WriteLine "<table class=""export"">"
				objFile.WriteLine "	<tr>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Product"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Product Key"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "	</tr>"
			Case 4
				intColumnIndex = 12
				If Len(strProduct1) + 5 > intColumnIndex Then _
					intColumnIndex = Len(MSProd1.InnerHTML) + 5
				If Len(strProduct2) + 5 > intColumnIndex Then _
					intColumnIndex = Len(MSProd2.InnerHTML) + 5
				If Len(strProduct3) + 5 > intColumnIndex Then _
					intColumnIndex = Len(MSProd3.InnerHTML) + 5
				If Len(strProduct4) + 5 > intColumnIndex Then _
					intColumnIndex = Len(MSProd4.InnerHTML) + 5
				If Len(strProduct5) + 5 > intColumnIndex Then _
					intColumnIndex = Len(MSProd5.InnerHTML) + 5
				If Len(strProduct6) + 5 > intColumnIndex Then _
					intColumnIndex = Len(MSProd6.InnerHTML) + 5
				
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\MSProdKeyDetails" & strPC & ".txt",True)
				objFile.WriteLine "Microsoft Product Keys on " & strPC
				objFile.WriteLine ""
				objFile.WriteLine "Product" & String(intColumnIndex - 7, " ") & "Product Key"
		End Select
			
		If MSProdKeyExport.Value = 1 Then 
			strProduct1 = Replace(strProduct1, ",", ".")
			strProduct2 = Replace(strProduct2, ",", ".")
			strProduct3 = Replace(strProduct3, ",", ".")
			strProduct4 = Replace(strProduct4, ",", ".")
			strProduct5 = Replace(strProduct5, ",", ".")
			strProduct6 = Replace(strProduct6, ",", ".")
		End If
		
		Select Case MSProdKeyExport.Value
			Case 1
				If NOT IsNull(strProduct1) AND strProduct1 <> "" Then
					strCSV = strCSV & strProduct1 & "," & _
					strKey1 & vbCrLf
				End If
				If NOT IsNull(strProduct2) AND strProduct2 <> "" Then
					strCSV = strCSV & strProduct2 & "," & _
					strKey2 & vbCrLf
				End If
				If NOT IsNull(strProduct3) AND strProduct3 <> "" Then
					strCSV = strCSV & strProduct3 & "," & _
					strKey3 & vbCrLf
				End If
				If NOT IsNull(strProduct4) AND strProduct4 <> "" Then
					strCSV = strCSV & strProduct4 & "," & _
					strKey4 & vbCrLf
				End If
				If NOT IsNull(strProduct5) AND strProduct5 <> "" Then
					strCSV = strCSV & strProduct5 & "," & _
					strKey5 & vbCrLf
				End If
				If NOT IsNull(strProduct6) AND strProduct6 <> "" Then
					strCSV = strCSV & strProduct6 & "," & _
					strKey6 & vbCrLf
				End If
			Case 2
				If NOT IsNull(strProduct1) AND strProduct1 <> "" Then
					objWorkSheet.Cells(intStartRow, 1) = strProduct1
					objWorkSheet.Cells(intStartRow, 2) = strKey1
					intStartRow = intStartRow + 1
				End If
				If NOT IsNull(strProduct2) AND strProduct2 <> "" Then
					objWorkSheet.Cells(intStartRow, 1) = strProduct2
					objWorkSheet.Cells(intStartRow, 2) = strKey2
					intStartRow = intStartRow + 1
				End If
				If NOT IsNull(strProduct3) AND strProduct3 <> "" Then
					objWorkSheet.Cells(intStartRow, 1) = strProduct3
					objWorkSheet.Cells(intStartRow, 2) = strKey3
					intStartRow = intStartRow + 1
				End If
				If NOT IsNull(strProduct4) AND strProduct4 <> "" Then
					objWorkSheet.Cells(intStartRow, 1) = strProduct4
					objWorkSheet.Cells(intStartRow, 2) = strKey4
					intStartRow = intStartRow + 1
				End If
				If NOT IsNull(strProduct5) AND strProduct5 <> "" Then
					objWorkSheet.Cells(intStartRow, 1) = strProduct5
					objWorkSheet.Cells(intStartRow, 2) = strKey5
					intStartRow = intStartRow + 1
				End If
				If NOT IsNull(strProduct6) AND strProduct6 <> "" Then
					objWorkSheet.Cells(intStartRow, 1) = strProduct6
					objWorkSheet.Cells(intStartRow, 2) = strKey6
					intStartRow = intStartRow + 1
				End If
			Case 3
				If NOT IsNull(strProduct1) AND strProduct1 <> "" Then
					objFile.WriteLine "	<tr>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strProduct1
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td style=""font-family:'Courier New';"">"
					objFile.WriteLine "			" & strKey1
					objFile.WriteLine "		</td>"
					objFile.WriteLine "	</tr>"
				End If
				If NOT IsNull(strProduct2) AND strProduct2 <> "" Then
					objFile.WriteLine "	<tr>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strProduct2
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td style=""font-family:'Courier New';"">"
					objFile.WriteLine "			" & strKey2
					objFile.WriteLine "		</td>"
					objFile.WriteLine "	</tr>"
				End If
				If NOT IsNull(strProduct3) AND strProduct3 <> "" Then
					objFile.WriteLine "	<tr>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strProduct3
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td style=""font-family:'Courier New';"">"
					objFile.WriteLine "			" & strKey3
					objFile.WriteLine "		</td>"
					objFile.WriteLine "	</tr>"
				End If
				If NOT IsNull(strProduct4) AND strProduct4 <> "" Then
					objFile.WriteLine "	<tr>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strProduct4
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td style=""font-family:'Courier New';"">"
					objFile.WriteLine "			" & strKey4
					objFile.WriteLine "		</td>"
					objFile.WriteLine "	</tr>"
				End If
				If NOT IsNull(strProduct5) AND strProduct5 <> "" Then
					objFile.WriteLine "	<tr>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strProduct5
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td style=""font-family:'Courier New';"">"
					objFile.WriteLine "			" & strKey5
					objFile.WriteLine "		</td>"
					objFile.WriteLine "	</tr>"
				End If
				If NOT IsNull(strProduct6) AND strProduct6 <> "" Then
					objFile.WriteLine "	<tr>"
					objFile.WriteLine "		<td>"
					objFile.WriteLine "			" & strProduct6
					objFile.WriteLine "		</td>"
					objFile.WriteLine "		<td style=""font-family:'Courier New';"">"
					objFile.WriteLine "			" & strKey6
					objFile.WriteLine "		</td>"
					objFile.WriteLine "	</tr>"
				End If
			Case 4
				If NOT IsNull(strProduct1) AND strProduct1 <> "" Then
					strTxt = strTxt & strProduct1 & _
					String(intColumnIndex - Len(strProduct1), " ") & _
					strKey1 & vbCrLf
				End If
				If NOT IsNull(strProduct2) AND strProduct2 <> "" Then
					strTxt = strTxt & strProduct2 & _
					String(intColumnIndex - Len(strProduct2), " ") & _
					strKey2 & vbCrLf
				End If
				If NOT IsNull(strProduct3) AND strProduct3 <> "" Then
					strTxt = strTxt & strProduct3 & _
					String(intColumnIndex - Len(strProduct3), " ") & _
					strKey3 & vbCrLf
				End If
				If NOT IsNull(strProduct4) AND strProduct4 <> "" Then
					strTxt = strTxt & strProduct4 & _
					String(intColumnIndex - Len(strProduct4), " ") & _
					strKey4 & vbCrLf
				End If
				If NOT IsNull(strProduct5) AND strProduct5 <> "" Then
					strTxt = strTxt & strProduct5 & _
					String(intColumnIndex - Len(strProduct5), " ") & _
					strKey5 & vbCrLf
				End If
				If NOT IsNull(strProduct6) AND strProduct6 <> "" Then
					strTxt = strTxt & strProduct6 & _
					String(intColumnIndex - Len(strProduct6), " ") & _
					strKey6 & vbCrLf
				End If
		End Select		

		Select Case MSProdKeyExport.Value
			Case 1
				objFile.WriteLine strCSV
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\MSProdKeyDetails" & strPC & ".csv"
			Case 2
				Set objRange = objWorkSheet.Range("A1:Z3")
				Set objRange2 = objWorkSheet.Range("A3:B" & intStartRow - 1)
				Set objRange3 = objWorkSheet.Range("B4:B" & intStartRow - 1)
				
				objRange.Font.Bold = True
				objRange2.Borders.LineStyle = xlContinuous
				objRange2.Borders.Weight = xlThin
				objRange2.Borders.ColorIndex = xlAutomatic
				objRange3.Font.Name = "Courier New"
				objWorksheet.Range("A1").Select
				
				objWorkSheet.Columns("A:ZZ").EntireColumn.AutoFit
				objExcel.DisplayAlerts = False
				objExcel.ActiveWorkbook.SaveAs(strTemp & "\SKB\MSProdKeyDetails" & strPC & ".xls")
				objExcel.Visible = True
				Set objExcel = Nothing
			Case 3
				objFile.WriteLine "</table>"
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\MSProdKeyDetails" & strPC & ".htm"
			Case 4
				objFile.WriteLine strTxt
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\MSProdKeyDetails" & strPC & ".txt"
			End Select
		
		MSProdKeyExport.Value = 0
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	CopyKeyInfo()
    '#	PURPOSE........:	Copies the product key to the clipboard
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub CopyKeyInfo(intChoice)
		Select Case intChoice
			Case 1
				Document.parentwindow.clipboardData.SetData "text", MSProdKey1.InnerHTML
			Case 2
				Document.parentwindow.clipboardData.SetData "text", MSProdKey2.InnerHTML
			Case 3
				Document.parentwindow.clipboardData.SetData "text", MSProdKey3.InnerHTML
			Case 4
				Document.parentwindow.clipboardData.SetData "text", MSProdKey4.InnerHTML
			Case 5
				Document.parentwindow.clipboardData.SetData "text", MSProdKey5.InnerHTML
			Case 6
				Document.parentwindow.clipboardData.SetData "text", MSProdKey6.InnerHTML
		End Select
		
		MsgBox "The product key has now been copied to the clipboard", vbInformation, "PC Management Utility"
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ListUpdates()
    '#	PURPOSE........:	Lists all security updates / hotfixes
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ListUpdates()
		x = 0
		Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strPC & "\root\cimv2")

		Set colQuickFixes = objWMIService.ExecQuery _
			("Select * from Win32_QuickFixEngineering")

		For Each objItem in colQuickFixes
			strDescription = objItem.Description
			strHotfixID = objItem.HotFixID
			dtmInstallDate = objItem.InstallDate
			strInstalledBy = objItem.InstalledBy
			If IsNull(dtmInstallDate) OR dtmInstallDate = "" Then
				dtmInstallDate = objItem.InstalledOn
			End If
			If strHotfixID <> "File 1" Then
				strTxt = strTxt & "Description: " & strDescription & vbCrLf & _
				"Hotfix ID: " & strHotfixID & vbCrLf & _
				"Installation Date: " & dtmInstallDate & vbCrLf & _
				"Installed By: " & strInstalledBy & vbCrLf & _
				"--------------" & vbCrLf
				x = x + 1
			End If
		Next
		
		strHTML = "<textarea name=""UpdatesListTextArea"" rows=""10"" cols=""77""></textarea>"
		strHTML = strHTML & "<br><div style=""float:right;"">"
		strHTML = strHTML & "<select name=""UpdatesExport"" "
		strHTML = strHTML & "title=""Export the list of security updates / hotfixes"" onChange=""ExportUpdateInfo()"">"
		strHTML = strHTML & "	<option value=""0"">Export to:</option>"
		strHTML = strHTML & "	<option value=""1"" title=""Export the security updates / hotfixes to a Comma " & _
		"Seperated Values (csv) file"")>Export to csv</option>"
		strHTML = strHTML & "	<option value=""2"" title=""Export the security updates / hotfixes to a formatted Excel " & _
		"(xls) spreadsheet"">Export to xls</option>"
		strHTML = strHTML & "	<option value=""3"" title=""Export the security updates / hotfixes to a Web " & _
		"page (html) file"">Export to html</option>"
		strHTML = strHTML & "	<option value=""4"" title=""Export the security updates / hotfixes to a Text " & _
		"(txt) file"">Export to txt</option>"
		strHTML = strHTML & "</select>"
		
		WaitMessage.InnerHTML = strHTML
		
		UpdatesListTextArea.Value = "Updates on " & strPC & vbCrLf & vbCrLf & _
		strTxt & vbCrLf & vbCrLf & "Total Updates: " & x
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ExportUpdateInfo()
    '#	PURPOSE........:	Exports the list of security updates / hotfixes
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------	
	Sub ExportUpdateInfo()
		On Error Resume Next
		intColumnIndex = 16
		intColumnIndex2 = 14
		intColumnIndex3 = 22
		arrTextArea = Split(UpdatesListTextArea.Value, vbCrLf)
		strTotal = arrTextArea(UBound(arrTextArea))
		
		Select Case UpdatesExport.Value
			Case 1
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\ListUpdates" & strPC & ".csv",True)
				objFile.WriteLine "Updates on " & strPC
				objFile.WriteLine ""
				objFile.WriteLine strTotal
				objFile.WriteLine ""
				objFile.WriteLine "Description,Hotfix ID,Installation Date,Installed By"
			Case 2
				Const xlContinuous = 1
				Const xlThin = 2
				Const xlAutomatic = -4105
				Const xlCenter = -4108
				
				strExcelPath = objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe\")
			   
				If strExcelPath = "" Then
					MsgBox "Unable to export. Excel does not appear to be installed.", vbExclamation, "PC Management Utility"
					Exit Sub
				End If
				
				Set objExcel = CreateObject("Excel.Application")
				objExcel.Visible = False
				Set objWorkBook = objExcel.WorkBooks.Add
				Set objWorksheet = objWorkbook.Worksheets(1)
				objExcel.DisplayAlerts = False
				For i = 1 to 3
					objWorkbook.Worksheets(2).Delete
				Next
				objExcel.DisplayAlerts = True
				objWorksheet.Name = "Updates"
				
				objWorkSheet.Cells(1, 1) = "Updates on " & strPC
				objWorkSheet.Cells(3, 1) = strTotal

				intStartRow = 6
				
				objWorkSheet.Cells(5, 1) = "Description"
				objWorkSheet.Cells(5, 2) = "Hotfix ID"
				objWorkSheet.Cells(5, 3) = "Installation Date"
				objWorkSheet.Cells(5, 4) = "Installed By"
			Case 3
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\ListUpdates" & strPC & ".htm",True)
				objFile.WriteLine "<style type=""text/css"">"
				objFile.WriteLine "body{background-color:#CEF0FF;}"
				objFile.WriteLine "table.export{border-width:1px;border-spacing:1px;border-style:solid;border-color:gray;border-collapse:collapse;}"
				objFile.WriteLine "table.export th{border-width:1px;padding:1px;border-style:solid;border-color:gray;padding:2px 7px 2px 7px;}"
				objFile.WriteLine "table.export td{border-width:1px;padding:1px;border-style:dotted;border-color:gray;padding:2px 7px 2px 7px;}"
				objFile.WriteLine ".backtotop a {font-size:0.9em;}"
				objFile.WriteLine "</style>"
				objFile.WriteLine "<div style=""font-weight:bold;""><a name =""top"">Updates on " & strPC & "</a><p>"
				objFile.WriteLine strTotal & "</div><p>"
				objFile.WriteLine "<table class=""export"">"
				objFile.WriteLine "	<tr>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Description"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th>"
				objFile.WriteLine "			Hotfix ID"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th>"
				objFile.WriteLine "			Installation Date"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th>"
				objFile.WriteLine "			Installed By"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "	</tr>"
			Case 4		
				For i = 0 to UBound(arrTextArea)
					strLine = arrTextArea(i)
					If InStr(strLine, "Description:") > 0 Then
						strLine = Replace(strLine, "Description: ", "")
						strDescription = strLine
						If Len(strDescription) > intColumnIndex - 5 Then intColumnIndex = Len(strDescription) + 5
					End If
					If InStr(strLine, "Hotfix ID:") > 0 Then
						strLine = Replace(strLine, "Hotfix ID: ", "")
						strHotfixID = strLine
						If Len(strHotfixID) > intColumnIndex2 - 5 Then intColumnIndex2 = Len(strHotfixID) + 5
					End If
				Next

				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\ListUpdates" & strPC & ".txt",True)
				objFile.WriteLine "Updates on " & strPC
				objFile.WriteLine ""
				objFile.WriteLine strTotal
				objFile.WriteLine ""
				objFile.WriteLine "Description" & _
				String(intColumnIndex - 11, " ") & "Hotfix ID" & _
				String(intColumnIndex2 - 9, " ") & "Installation Date" & _
				String(5, " ") & "Installed By"
		End Select
		
		For i = 0 to UBound(arrTextArea)
			strLine = arrTextArea(i)
			If strLine <> "--------------" _
			AND strLine <> "" AND InStr(strLine, "Updates on ") = 0 Then
				If InStr(strLine, "Description:") > 0 Then
					strLine = Replace(strLine, "Description: ", "")
					strDescription = strLine
				End If
				If InStr(strLine, "Hotfix ID:") > 0 Then
					strLine = Replace(strLine, "Hotfix ID: ", "")
					strHotfixID = strLine
				End If
				If InStr(strLine, "Installation Date:") > 0 Then
					strLine = Replace(strLine, "Installation Date: ", "")
					dtmInstallDate = strLine
					If IsNull(dtmInstallDate) OR dtmInstallDate = "" Then
						dtmInstallDate = objItem.InstalledOn
					End If
					If IsDate(dtmInstallDate) Then dtmInstallDate = CDate(dtmInstallDate)
				End If
				If InStr(strLine, "Installed By:") > 0 Then
					strLine = Replace(strLine, "Installed By: ", "")
					strInstalledBy = strLine
					
					If strHotfixID <> "File 1" Then
						Select Case UpdatesExport.Value
							Case 1
								strCSV = strCSV & strDescription & "," & _
								strHotfixID & "," & dtmInstallDate & "," & _
								strInstalledBy & vbCrLf
							Case 2
								objWorkSheet.Cells(intStartRow, 1) = strDescription
								objWorkSheet.Cells(intStartRow, 2) = strHotfixID
								objWorkSheet.Cells(intStartRow, 3) = dtmInstallDate
								objWorkSheet.Cells(intStartRow, 4) = strInstalledBy
								intStartRow = intStartRow + 1
							Case 3
								objFile.WriteLine "	<tr>"
								objFile.WriteLine "		<td>"
								objFile.WriteLine "			" & strDescription
								objFile.WriteLine "		</td>"
								objFile.WriteLine "		<td>"
								objFile.WriteLine "			" & strHotfixID
								objFile.WriteLine "		</td>"
								objFile.WriteLine "		<td style=""text-align:center;"">"
								objFile.WriteLine "			" & dtmInstallDate
								objFile.WriteLine "		</td>"
								objFile.WriteLine "		<td>"
								objFile.WriteLine "			" & strInstalledBy
								objFile.WriteLine "		</td>"
								objFile.WriteLine "	</tr>"
							Case 4
								strTxt = strTxt & strDescription & _
								String(intColumnIndex - Len(strDescription), " ") & _
								strHotfixID & String(intColumnIndex2 - Len(strHotfixID), " ") & _
								dtmInstallDate & String(intColumnIndex3 - Len(dtmInstallDate), " ") & _
								strInstalledBy & vbCrLf
						End Select
					End If
				End If
			End If
		Next
		
		Select Case UpdatesExport.Value
			Case 1
				objFile.WriteLine strCSV
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\ListUpdates" & strPC & ".csv"
			Case 2
				Set objRange = objWorkSheet.Range("A1:Z5")
				Set objRange2 = objWorkSheet.Range("A5:D" & intStartRow - 1)
				Set objRange3 = objWorkSheet.Range("C:C")
				
				objRange.Font.Bold = True
				objRange2.Borders.LineStyle = xlContinuous
				objRange2.Borders.Weight = xlThin
				objRange2.Borders.ColorIndex = xlAutomatic
				objRange3.HorizontalAlignment = xlCenter
				objWorksheet.Range("A6").Select
				objExcel.ActiveWindow.FreezePanes = "True"
				objWorksheet.Range("A1").Select
				
				objWorkSheet.Columns("A:ZZ").EntireColumn.AutoFit
				objExcel.DisplayAlerts = False
				objExcel.ActiveWorkbook.SaveAs(strTemp & "\SKB\ListUpdates" & strPC & ".xls")
				objExcel.Visible = True
				Set objExcel = Nothing
			Case 3
				objFile.WriteLine "</table>"
				objFile.WriteLine "<p class=""backtotop""><a href=""" & strHTMLTempDir & "/SKB/ListUpdates" & _
				strPC & ".htm#top"">[..back to top..]</a></p>"
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\ListUpdates" & strPC & ".htm"
			Case 4
				objFile.WriteLine strTxt
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\ListUpdates" & strPC & ".txt"
		End Select
		
		UpdatesExport.Value = 0
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ManagePC()
    '#	PURPOSE........:	Opens Computer Management for PC
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ManagePC()
		command = "mmc %windir%\system32\compmgmt.msc -s /computer:\\" & strPC
		objShell.Run command,1,False
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	OpenShare()
    '#	PURPOSE........:	Opens the requested share for the PC
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub OpenShare()
		WaitMessage.InnerHTML = "<div style=""text-align:left;font-weight:bold;"">Please select a share:<br>" & _
		"<select size=""9"" name=""ShareChooser"" style=""width:100%"" " & _
		"onDblClick=""OpenShare2"">" & _
		"</select><br>" & _
		"<input id=runbutton type=""button"" value=""Open Share"" name=""btnOpenShare"" title=""Open selected share on remote PC"" onClick=""OpenShare2()""></div>"
		
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\cimv2")
		
		Set colShares = objWMIService.ExecQuery _
			("Select * from Win32_Share Where Type='2147483648'")
					
		For Each objOption in ShareChooser.Options
			objOption.RemoveNode
		Next 
		
		For Each objItem In colShares
			strShareName = objItem.Name
			strCaption = objItem.Caption
			strPath = objItem.Path
			strShare = "\\" & strPC & "\" & strShareName			
			Set objOption = Document.createElement("OPTION")
			objOption.Text = strShare & " (" & strCaption & ")"
			objOption.Title = "Path: " & strPath
			objOption.Value = "explorer.exe /e, " & strShare & ", /separate"
			ShareChooser.Add(objOption)
		Next
	End Sub
	
	Sub OpenShare2()
		If ShareChooser.Value <> "" Then
			objShell.Run ShareChooser.Value
		End If
	End Sub
	
	Sub PrePingMachine()
		strHTML = "<hr><div style=""text-align:left;"">" & _
		"<p><input type=""text"" name=""txtPingNum"" value=""5"" size=""1"" onKeyUp=PingCheck() " & _
		"style=""text-align:center;""> " & _
		"&nbsp;Number of pings required (type 0 for unlimited pings)<p>" & _
		"<input id=""RunButton"" class=""button"" type=""button"" value=""Start Ping"" " & _
		"name=""btnStartPing"" onclick=""PingMachine True, False"" " & _
		"title=""Ping machine using specified number of pings""> &nbsp;&nbsp;" & _
		"<input type=""checkbox"" name=""cbxExportPingInfo"" " & _
		"title=""Export info to text file on completion"">Export" & _
		"</div>"
		WaitMessage.InnerHTML = strHTML
		txtPingNum.Focus()
		txtPingNum.Select()
	End Sub
	
	Sub PingCheck()
		If IsNull(txtPingNum.Value) OR txtPingNum.Value = "" Then 
			cbxExportPingInfo.Disabled = False
			Exit Sub
		End If
		If IsNumeric(txtPingNum.Value) = False Then
			txtPingNum.Value = ""
			Exit Sub
		End If
		If txtPingNum.Value = 0 Then
			cbxExportPingInfo.Checked = False
			cbxExportPingInfo.Disabled = True
			Else
				cbxExportPingInfo.Disabled = False
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	PingMachine(booAction,booErr)
    '#	PURPOSE........:	Pings a machine and outputs the data to the screen
	'#						in the same format as DOS Ping
    '#	ARGUMENTS......:	booAction = boolean value to determine whether the
	'#						ping was started using the Ping command from the
	'#						Action List (True) or not (False)
	'#						booErr = boolean value to determine whether the
	'#						was started due to an unreachable PC (True) Or
	'#						not (False)
    '#	EXAMPLE........:	PingMachine(PC1,True,False)
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub PingMachine(booAction,booErr)
		On Error Resume Next
		Dim arrPing()
		Dim arrPingResponseTime()
		minPingResponseTime = 999999
		maxPingResponseTime = 0
		intRequestTimedOut = 0
		intPings = txtPingNum.Value
		If intPings = "" Then intPings = 0
		If cbxExportPingInfo.Checked = True Then ViewPingPrompt = True
		
		If booAction = True Then
			booAction2 = False
			If intPings = 0 Then
				PingMachine False, False
				Exit Sub
			End If
			If intPings > 60 Then
				LongPingPrompt = MsgBox("This operation will take longer than 1 minute" & _
				vbCrLf & vbCrLf & "Do you wish to continue?",vbYesNo+vbExclamation, "Ping")
				If LongPingPrompt = vbNo Then
					Exit Sub
				End If
			End If	
			document.body.style.cursor = "wait"
			Else
				intPings = 999999
		End If
		btnStop.Disabled = False
		btnStop.title = "Stop running action"
		Set objWMIService2 = GetObject("winmgmts:\\.\root\cimv2")
		
		intPingResponse = 0	
		For i = 1 to intPings
			If btnStop.Disabled = True Then
					If booAction = False OR booAction2 = False Then
						If intPingResponse = 0 Then
								intPingsLost = intPingsSent
								pctPingsLost = 100
								objTextArea.Value = objTextArea.Value & vbCrLf & "Packets: Sent = " & _
								intPingsSent & ", Received = " & intPingResponse & ", Lost = " & _
								intPingsLost & " (" & pctPingsLost & "% loss)"
								Else
									intAverageResponseTime = Round(intAverageResponseTime / intPingResponse)
									intPingsLost = intPingsSent - intPingResponse
									pctPingsLost = Round(intPingsLost / intPingResponse * 100,2)
									objTextArea.Value = objTextArea.Value & vbCrLf & "Packets: Sent = " & _
									intPingsSent & ", Received = " & intPingResponse & ", Lost = " & _
									intPingsLost & " (" & pctPingsLost & "% loss)" & vbCrLf & _
									"Approximate round trip times in milliseconds:" & _
									vbCrLf & "  Minimum = " & minPingResponseTime & "ms, Maximum = " & _
									maxPingResponseTime & "ms, Average = " & intAverageResponseTime & "ms"
									objTextArea.scrollTop = objTextArea.scrollTop + objTextArea.scrollHeight
							End If
						Exit Sub
					End If
			End If
			ReDim Preserve arrPing(i)
			ReDim Preserve arrPingResponseTime(i)
			Set colIP = objWMIService2.ExecQuery _
					("Select * From Win32_PingStatus Where Address = '" & _
					strPC & "'")
			For Each objItem in colIP
				strPingIP = objItem.ProtocolAddress
				strPingReplySize = objItem.ReplySize
				strPingTTL = objItem.TimeToLive
				If intRequestTimedOut = 3 Then
					intRequestTimedOut = 0
				End If
				intPingStatusCode = objItem.StatusCode
				If IsNull(intPingStatusCode) Then
					intRequestTimedOut = intRequestTimedOut + 1
					If intRequestTimedOut = 3 Then
						If intPingsSent < 10 Then
								arrPing(i) = "0" & intPingsSent + 1 & ": Request Timed Out"
								Else
									arrPing(i) = intPingsSent + 1 & ": Request Timed Out"
						End If
					End If
					ElseIf intPingStatusCode <> 0 Then
						intRequestTimedOut = intRequestTimedOut + 1
						If intRequestTimedOut = 3 Then
							If intPingsSent < 10 Then
								arrPing(i) = "0" & intPingsSent + 1 & ": " & _
								GetPingStatus(intPingStatusCode)
								Else
									arrPing(i) = intPingsSent + 1 & ": " & _
									GetPingStatus(intPingStatusCode)
							End If
						End If
						Else
							intRequestTimedOut = 0
							arrPingResponseTime(i) = objItem.ResponseTime
							If minPingResponseTime > arrPingResponseTime(i) Then
								minPingResponseTime = arrPingResponseTime(i)
							End If
							If  maxPingResponseTime < arrPingResponseTime(i) Then
								maxPingResponseTime = arrPingResponseTime(i)
							End If
							intAverageResponseTime = intAverageResponseTime + arrPingResponseTime(i)
							If arrPingResponseTime(i) < 1 Then
								arrPingResponseTime(i) = "<1"
								Else
									arrPingResponseTime(i) = "=" & arrPingResponseTime(i)
							End If
							
							If i < 10 Then
								arrPing(i) = "0" & intPingsSent + 1 & ": Reply from " & strPingIP & _
								": bytes=" & strPingReplySize & " time" & arrPingResponseTime(i) & _
								"ms TTL=" & strPingTTL
								Else
									arrPing(i) = intPingsSent + 1 & ": Reply from " & strPingIP & _
									": bytes=" & strPingReplySize & " time" & _
									arrPingResponseTime(i) & "ms TTL=" & strPingTTL
							End If
							intPingResponse = intPingResponse + 1
				End If
				If i = 1 Then
					If booErr = False Then
						WaitMessage.InnerHTML = "<textarea name=""PingTextArea"" rows=""11"" " & _
						"cols=""77""></textarea>"
						Else
							WaitMessage.InnerHTML = "<textarea name=""PingTextArea"" rows=""26"" " & _
							"cols=""77""></textarea>"
					End If
				End If
				If intRequestTimedOut = 0 OR intRequestTimedOut = 3 Then
					intPingsSent = intPingsSent + 1
					strPingMsg = strPingMsg & arrPing(i) & vbCrLf
					
					Dim objTextArea : Set objTextArea = window.document.getElementById("PingTextArea")
					objTextArea.Value = strPingMsg
					objTextArea.scrollTop = objTextArea.scrollTop + objTextArea.scrollHeight
				End If
				PauseScript(1000)
			Next
		Next
		If booAction = True Then
			intAverageResponseTime = Round(intAverageResponseTime / intPingResponse)
			intPingsLost = intPingsSent - intPingResponse
			pctPingsLost = Round(intPingsLost / intPingResponse * 100,2)
			objTextArea.Value = objTextArea.Value & vbCrLf & "Packets: Sent = " & _
			intPingsSent & ", Received = " & intPingResponse & ", Lost = " & _
			intPingsLost & " (" & pctPingsLost & "% loss)" & vbCrLf & _
			"Approximate round trip times in milliseconds:" & vbCrLf & _
			"  Minimum = " & minPingResponseTime & "ms, Maximum = " & _
			maxPingResponseTime & "ms, Average = " & intAverageResponseTime & "ms"
			objTextArea.scrollTop = objTextArea.scrollTop + objTextArea.scrollHeight
			document.body.style.cursor = "default"
			If ViewPingPrompt = True Then
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\ping.txt",True)
				objFile.WriteLine "Pinging " & strPC & " [" & strPingIP & _
				"] with " & strPingReplySize & " bytes of data: "
				objFile.WriteLine ""
				objFile.WriteLine strPingMsg
				objFile.WriteLine "Ping statistics for " & _
				strPingIP & ":"
				objFile.WriteLine vbTab & "Packets: Sent = " & intPingsSent & ", Received = " & _
				intPingResponse & ", Lost = " & intPingsLost & " (" & pctPingsLost& "% loss)"
				objFile.WriteLine "Approximate round trip times in milliseconds:"
				objFile.WriteLine vbTab & "Minimum = " & minPingResponseTime & _
				"ms, Maximum = " & maxPingResponseTime & "ms, Average = " & _
				intAverageResponseTime & "ms"
				objFile.Close
				objShell.Run strTemp & "\SKB\ping.txt"
			End If
		End If
		btnStop.Disabled = True
		btnStop.style.cursor = "default"
		btnStop.title = ""
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	RunPSExecCommands()
    '#	PURPOSE........:	Allows a user to build and perform PSExec commands 
	'#						upon a machine using a form and also offers the 
	'#						ability to save / open commands
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub RunPSExecCommand()
		On Error Resume Next
		booPSExecInPath = False
		strPath = objShell.ExpandEnvironmentStrings("%path%")
		arrPaths = Split(strPath, ";")
		For i = 0 To UBound(arrPaths)
			strPathFolder = arrPaths(i) & "\"
			strPathFolder = Replace(strPathFolder, "\\", "\")
			strPathFolder = Replace(LCase(strPathFolder), "%systemroot%", _
			objShell.ExpandEnvironmentStrings("%systemroot%"))
			If objFSO.FileExists(strPathFolder & "psexec.exe") Then 
				booPSExecInPath = True
				strPSExecLocation = strPathFolder & "psexec.exe"
			End If
		Next
		
		strFooter = "<span id=""BackToAction"" title=""Back to Actions list"" " & _
		"style=""text-decoration:underline;color:blue;cursor:pointer;position:absolute;bottom:50px"" " & _
		"onclick=""ShowPCActions()"">[..back to action list..]</span>"	

		If booPSExecInPath = False Then
			strHTML = "<script language=""VBScript"">" & _
			"On Error Resume Next" & _
			"<" & Chr(47) & "script>" & _
			"<b><u>" & strPC & "</u></b><p>" & _
			"PSExec cannot be found in the System Path. " & _
			"The folders within your System Path are as below:<p>"
			If UBound(arrPaths) < 21 Then
				For i = 0 To UBound(arrPaths)
					strPathFolder = arrPaths(i) & "\"
					strPathFolder = Replace(strPathFolder, "\\", "\")
					strPathFolder = Replace(LCase(strPathFolder), "%systemroot%", _
					objShell.ExpandEnvironmentStrings("%systemroot%"))
					strHTML = strHTML & LCase(strPathFolder) & "<br>"
				Next
				Else strHTML = strHTML & strPath & "<br>"
			End If
			strHTML = strHTML & "<br>" & _
			"You can download this utility from the following location:<p>" & _
			"<a href=""http://technet.microsoft.com/en-us/sysinternals/bb897553.aspx"" " & _
			"title=""Download PSExec"" target=""_blank"">" & _
			"http://technet.microsoft.com/en-us/sysinternals/bb897553.aspx</a></p><br>" & _
			"<br>Once you have downloaded PSExec you must place it in one of the above mentioned folders.</p>"
			
			DataArea.InnerHTML = strHTML & strFooter
			Exit Sub
		End If
		strHTML = "<script language=""VBScript"">On Error Resume Next<" & Chr(47) & "script><b><u>" & strPC & "</u></b><p>" & _
		"Use the form below to perform commands on " & strPC & " using PSExec." & _
		"<form Name = ""Commands"" Method = ""post"">" & _
		"<table id=""CommandsTable"" width=""90%"">" & _
			"<tr><td>" & _
				"<table width=""100%""><tr>" & _
				"<td title=""Command to be run"">" & _
					"Command: " & _
				"</td title=""Command to be run"">" & _
				"<td>" & _
					"<input type=""text"" name=""txtCommand"" size=""20"" " & _
					"onKeyUp=""ExecutePSExecCommand False, False, False"">&nbsp;&nbsp;&nbsp;" & _
				"</td>" & _
				"<td title=""Command variables"">" & _
					"Command Variables: " & _
				"</td>" & _
				"<td title=""Command variables"">" & _
					"<input type=""text"" name=""txtVariables"" size=""20"" " & _
					"onKeyUp=""ExecutePSExecCommand False, False, False"">" & _
				"</td></tr>" & _
				"<tr>" & _
				"<td title=""Optional user name for login to remote system"">" & _
					"Username: " & _
				"</td>" & _
				"<td title=""Optional user name for login to remote system"">" & _
					"<input type=""text"" name=""txtUsername"" size=""20"" " & _
					"style=""background-color:#dddddd;"" " & _
					"disabled=true onKeyUp=""ExecutePSExecCommand False, False, False"">" & _
				"</td>" & _
				"<td title=""Optional password for username, if left blank where a username is specified " & _
				"then you will be asked to supply one on executing the command"">" & _
					"Password: " & _
				"</td>" & _
				"<td title=""Optional password for username, if left blank where a username is specified " & _
				"then you will be asked to supply one on executing the command"">" & _
					"<input type=""password"" name=""txtPassword"" size=""20"" " & _
					"style=""background-color:#dddddd;"" " & _
					"disabled=true onKeyUp=""ExecutePSExecCommand False, False, False"">" & _
				"</td>" & _
				"</tr>" & _
				"</tr></table>" & _
			"</td></tr>" & _
			"<tr><td>&nbsp;</td></tr>" & _
			"<tr><td>" & _
				"<table width=""100%""><tr>" & _
				"<td title=""Run the program so that it interacts with the desktop of the  specified " & _
				"session on the remote system"">" & _
					"Interactive" & _
				"</td>" & _
				"<td title=""Run the program so that it interacts with the desktop of the  specified " & _
				"session on the remote system"">" & _
					"<input type=""checkbox"" name=""cbxInteractive"" " & _
					"onClick=""ExecutePSExecCommand False, False, False"">" & _
				"</td>" & _
				"<td title=""Does not load the specified account's profile"">" & _
					"Do not load user profile" & _
				"</td>" & _
				"<td title=""Does not load the specified account's profile"">" & _
					"<input type=""checkbox"" name=""cbxNotLoadProf"" disabled=true " & _
					"onClick=""ExecutePSExecCommand False, False, False"">" & _
				"</td>" & _
				"<td title=""Run the remote process in the System account"">" & _
					"Run as System" & _
				"</td>" & _
				"<td title=""Run the remote process in the System account"">" & _
					"<input type=""checkbox"" name=""cbxSystem"" onClick=""ExecutePSExecCommand False, False, False"">" & _
				"</td></tr>" & _
				"<tr>" & _
				"<td title=""Copy the specified program to the remote system for execution"">" & _
					"Copy file to remote host" & _
				"</td>" & _
				"<td title=""Copy the specified program to the remote system for execution"">" & _
					"<input type=""checkbox"" name=""cbxCopyFile"" onClick=""ExecutePSExecCommand False, False, False"">" & _
				"</td>" & _
				"<td title=""Copy the specified program even if the file already exists on the remote system"">" & _
					"Force copy" & _
				"</td>" & _
				"<td title=""Copy the specified program even if the file already exists on the remote system"">" & _
					"<input type=""checkbox"" name=""cbxForceCopy"" disabled=true " & _
					"onClick=""ExecutePSExecCommand False, False, False"">" & _
				"</td>" & _
				"<td title=""Copy the specified file only if it has a higher version number or is " & _
				"newer than the one on the remote system"">" & _
					"Copy if file is newer" & _
				"</td>" & _
				"<td title=""Copy the specified file only if it has a higher version number or is " & _
				"newer than the one on the remote system"">" & _
					"<input type=""checkbox"" name=""cbxNewerFile"" disabled=true " & _
					"onClick=""ExecutePSExecCommand False, False, False"">" & _
				"</td></tr>" & _
				"<td title=""Optional user name for login to remote system"">" & _
					"Alternate Username" & _
				"</td>" & _
				"<td title=""Optional user name for login to remote system"">" & _
					"<input type=""checkbox"" name=""cbxUsername"" " & _
					"onClick=""ExecutePSExecCommand False, False, False"">" & _
				"</td>" & _
				"<td title=""Carries out the command specified by string and then " & _
				"the command Window remains (eg. for ipconfig)"">" & _
					"Keep CMD on screen" & _
				"</td>" & _
				"<td title=""Carries out the command specified by string and then " & _
				"the command Window remains (eg. for ipconfig)"">" & _
					"<input type=""checkbox"" name=""cbxCMD"" " & _
					"onClick=""ExecutePSExecCommand False, False, False"">" & _
				"</td>" & _
				"<td title=""Change the priority at which the process is run"" " & _
				"colspan=""2"">" & _
					"Priority &nbsp;&nbsp;" & _
					"<select size=""1"" name=""PriorityChooser"" style=""width:139px"" " & _
					"onChange=""ExecutePSExecCommand False, False, False"">" & _
						"<option value=""0""></option>" & _
						"<option value=""low"">Low</option>" & _
						"<option value=""belownormal"">Below Normal</option>" & _
						"<option value=""abovenormal"">Above Normal</option>" & _
						"<option value=""high"">High</option>" & _
						"<option value=""realtime"">Realtime</option>" & _
					"</select>" & _
				"</td>" & _
				"</tr>" & _
				"</table>" & _
			"</td></tr>" & _
		"</table></form><p>" & _
		"Preview:" & _
		"<div style=background-color:white;width:580px;height:50px; id=""PreviewCommand"">" & _
		"psexec.exe \\" & strPC & "</div>" & _
		"<div>" & _
		"<form Name = ""PSExecSaves"" Method = ""post"">" & _
		"<input id=runbutton  class=""button"" type=""button"" value=""Clear"" name=""btnPSClear"" " & _
		"title=""Clear all PS Exec Commands"" onClick=""ClearPSExecCommands()"">" & _
		"<input id=runbutton  class=""button"" type=""button"" value=""Save"" name=""btnPSSave"" " & _
		"title=""Save PSExec command for future use"" onClick=""ExecutePSExecCommand False, True, False"">" & _
		"<input id=runbutton  class=""button"" type=""button"" value=""Execute"" name=""btnPSExecute"" " & _
		"title=""Execute command on PC"" onClick=""ExecutePSExecCommand True, False, False"">" & _
		"<span style=""position:absolute;right:73px;"">" & _
		"Saved Commands: " & _
		"<select size=""1"" name=""SaveList"" style=""width:250px;"" " & _
		"onChange=""OpenPSExecCommand()""></select></span></form></div>" & _
		"<p><br><i><b>PSExec Location: </b>" & strPSExecLocation & "</i></p>"
		
		DataArea.InnerHTML = strHTML & strFooter
		PopulatePSExecSaveList()
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	PopulatePSExecSaveList()
    '#	PURPOSE........:	Populates the PSExec save list drop down box
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub PopulatePSExecSaveList()
		On Error Resume Next
		Dim arrPSExecSaves(9)
		arrPSExecSaves(0) = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave01")
		arrPSExecSaves(1) = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave02")
		arrPSExecSaves(2) = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave03")
		arrPSExecSaves(3) = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave04")
		arrPSExecSaves(4) = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave05")
		arrPSExecSaves(5) = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave06")
		arrPSExecSaves(6) = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave07")
		arrPSExecSaves(7) = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave08")
		arrPSExecSaves(8) = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave09")
		arrPSExecSaves(9) = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave10")
		
		For Each objOption in PSExecSaves.SaveList.Options
			objOption.RemoveNode
		Next 
		
		Set objOption = Document.createElement("OPTION")
		objOption.Text = ""
		objOption.Value = ""
		PSExecSaves.SaveList.Add(objOption)
		
		For i = 0 To 9
			If arrPSExecSaves(i) <> "" Then
				strCheckString = ""
				strPSExecSaveString = Replace(arrPSExecSaves(i), "strPC", strPC)
				arrPSExecSaveName = Split(strPSExecSaveString, "}{")
				strPSExecSaveName = arrPSExecSaveName(0)
				strPSExecSaveCommand = arrPSExecSaveName(1)
				strPSExecSaveVariables = arrPSExecSaveName(2)
				
				If strPSExecSaveVariables <> "" AND strPSExecSaveVariables <> " " Then _
					strPSExecSaveName = strPSExecSaveName & " " & strPSExecSaveVariables
				
				If InStr(LCase(strPSExecSaveCommand), "%comspec% /k") > 0 Then _
					strPSExecSaveName = strPSExecSaveName & " (Keep CMD)"
					
				strTitle = strPSExecSaveName
				
				Set objOption = Document.createElement("OPTION")
				objOption.Text = strPSExecSaveName
				objOption.Value = strPSExecSaveString
				objOption.Title = strTitle
				PSExecSaves.SaveList.Add(objOption)
			End If
		Next
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ClearPSExecCommands()
    '#	PURPOSE........:	Clears the PSExec form / preview
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ClearPSExecCommands()
		With Commands
			.txtCommand.Value = ""
			.txtVariables.Value = ""
			.txtUsername.Value = ""
			.txtPassword.Value = ""
			.txtUsername.style.backgroundcolor = "#dddddd"
			.txtPassword.style.backgroundcolor = "#dddddd"
			.txtUsername.Disabled = True
			.txtPassword.Disabled = True
			.cbxInteractive.Checked = False
			.cbxNotLoadProf.Disabled = True
			.cbxNotLoadProf.Checked = False
			.cbxSystem.Disabled = False
			.cbxSystem.Checked = False
			.cbxCopyFile.Checked = False
			.cbxForceCopy.Checked = False
			.cbxForceCopy.Disabled = True
			.cbxNewerFile.Checked = False
			.cbxNewerFile.Disabled = True
			.cbxUsername.Checked = False
			.cbxUsername.Disabled = False
			.cbxCMD.Checked = False
			.PriorityChooser.Value = 0
		End With
		PreviewCommand.InnerHTML = "psexec.exe \\" & strPC
	End Sub

	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	OpenPSExecCommand()
    '#	PURPOSE........:	Opens the command specified in the PSExec save list
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub OpenPSExecCommand()
		On Error Resume Next
		strCheckCommand = ""
		ClearPSExecCommands()
		If PSExecSaves.SaveList.Value <> "" Then
			strSaveString = PSExecSaves.SaveList.Value
			arrSaveString = Split(strSaveString, "}{")
			strPreviewSaveCommand = arrSaveString(0)
			strExecuteSaveCommand = LCase(arrSaveString(1))
			strVariables = LCase(arrSaveString(2))
			arrCheckCommand = Split(strExecuteSaveCommand, " ")
			For i = 0 To UBound(arrCheckCommand)
				strCheckCommand = strCheckCommand & arrCheckCommand(i) & " "
				If LCase(arrCheckCommand(i)) = "\\" & LCase(strPC) Then Exit For
			Next
			
			arrUserValue = Split(strPreviewSaveCommand, " ")
			If inStr(strCheckCommand, "-i") > 0 Then Commands.cbxInteractive.Checked = True
			If inStr(LCase(strExecuteSaveCommand), "%comspec% /k") > 0 Then Commands.cbxCMD.Checked = True
			If inStr(strCheckCommand, "-e") > 0 Then Commands.cbxNotLoadProf.Checked = True
			If inStr(strCheckCommand, "-s") > 0 Then Commands.cbxSystem.Checked = True
			If inStr(strCheckCommand, "-s") > 0 Then 
				Commands.cbxSystem.Checked = True
				Commands.cbxUsername.Disabled = True
			End If
			If inStr(strCheckCommand, "-c") > 0 Then 
				Commands.cbxCopyFile.Checked = True
				Commands.cbxForceCopy.Disabled = False
				Commands.cbxNewerFile.Disabled = False
			End If
			If inStr(strCheckCommand, "-f") > 0 Then Commands.cbxForceCopy.Checked = True
			If inStr(strCheckCommand, "-v") > 0 Then Commands.cbxNewerFile.Checked = True
			If inStr(strCheckCommand, "-low") > 0 Then _
				Commands.PriorityChooser.Value = "low"
			If inStr(strCheckCommand, "-belownormal") > 0 Then _
				Commands.PriorityChooser.Value = "belownormal"
			If inStr(strCheckCommand, "-abovenormal") > 0 Then _
				Commands.PriorityChooser.Value = "abovenormal"
			If inStr(strCheckCommand, "-high") > 0 Then _
				Commands.PriorityChooser.Value = "high"
			If inStr(strCheckCommand, "-realtime") > 0 Then _
				Commands.PriorityChooser.Value = "realtime"
			If inStr(strCheckCommand, "-u") > 0 Then 
				Commands.cbxUsername.Checked = True
				For i = 0 to UBound(arrUserValue)
					If arrUserValue(i) = "-u" Then Commands.txtUsername.Value = arrUserValue(i + 1)

				Next
				Commands.txtPassword.Value = "******"
				Commands.txtUsername.style.backgroundcolor = "white"
				Commands.txtPassword.style.backgroundcolor = "white"
				Commands.txtUsername.Disabled = False
				Commands.txtPassword.Disabled = False
				Commands.cbxNotLoadProf.Disabled = False
				Commands.cbxSystem.Disabled = True
			End If
			For i = 0 to UBound(arrUserValue)
				If LCase(arrUserValue(i)) = "\\" & LCase(strPC) Then
					For x = i + 1 To UBound(arrUserValue)
						Commands.txtCommand.Value = Commands.txtCommand.Value & arrUserValue(x) & " "
					Next
					Exit For
				End If
			Next
			Commands.txtVariables.Value = strVariables
			ExecutePSExecCommand False, False, True
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	UpdatePSExecSaveCommands(strSaveString)
    '#	PURPOSE........:	Adds the current PSExec command into save list
	'#						registry value
    '#	ARGUMENTS......:	strSaveString = string value to be saved in
	'#						registry
    '#	EXAMPLE........:	UpdatePSExecSaveCommands("Preview}{Execute}")
    '#	NOTES..........:	Allows 10 commands to be saved
    '#--------------------------------------------------------------------------
	Sub UpdatePSExecSaveCommands(strSaveString)
		On Error Resume Next
		strSaveString = Trim(strSaveString)
		strPSExecSave01 = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave01")
		strPSExecSave02 = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave02")
		strPSExecSave03 = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave03")
		strPSExecSave04 = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave04")
		strPSExecSave05 = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave05")
		strPSExecSave06 = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave06")
		strPSExecSave07 = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave07")
		strPSExecSave08 = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave08")
		strPSExecSave09 = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave09")
		strPSExecSave10 = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave10")
		
		If strPSExecSave10 = "" Then
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave10", strSaveString, "REG_SZ"
			PopulatePSExecSaveList()
			Exit Sub
		End If
		If strPSExecSave09 = "" Then
			If strPSExecSave10 <> strSaveString Then
				objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave09", strSaveString, "REG_SZ"
				PopulatePSExecSaveList()
				Exit Sub
			End If
		End If
		If strPSExecSave08 = "" Then
			If strPSExecSave10 <> strSaveString Then
				If strPSExecSave09 <> strSaveString Then
					objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave08", _
					strSaveString, "REG_SZ"
					PopulatePSExecSaveList()
					Exit Sub
				End If
			End If
		End If
		If strPSExecSave07 = "" Then
			If strPSExecSave10 <> strSaveString Then
				If strPSExecSave09 <> strSaveString Then
					If strPSExecSave08 <> strSaveString Then
						objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave07", _
						strSaveString, "REG_SZ"
						PopulatePSExecSaveList()
						Exit Sub
					End If
				End If
			End If
		End If
		If strPSExecSave06 = "" Then
			If strPSExecSave10 <> strSaveString Then
				If strPSExecSave09 <> strSaveString Then
					If strPSExecSave08 <> strSaveString Then
						If strPSExecSave07 <> strSaveString Then
							objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave06", _
							strSaveString, "REG_SZ"
							PopulatePSExecSaveList()
							Exit Sub
						End If
					End If
				End If
			End If
		End If
		If strPSExecSave05 = "" Then
			If strPSExecSave10 <> strSaveString Then
				If strPSExecSave09 <> strSaveString Then
					If strPSExecSave08 <> strSaveString Then
						If strPSExecSave07 <> strSaveString Then
							If strPSExecSave06 <> strSaveString Then
								objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave05", _
								strSaveString, "REG_SZ"
								PopulatePSExecSaveList()
								Exit Sub
							End If
						End If
					End If
				End If
			End If
		End If
		If strPSExecSave04 = "" Then
			If strPSExecSave10 <> strSaveString Then
				If strPSExecSave09 <> strSaveString Then
					If strPSExecSave08 <> strSaveString Then
						If strPSExecSave07 <> strSaveString Then
							If strPSExecSave06 <> strSaveString Then
								If strPSExecSave05 <> strSaveString Then
									objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave04", _
									strSaveString, "REG_SZ"
									PopulatePSExecSaveList()
									Exit Sub
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		If strPSExecSave03 = "" Then
			If strPSExecSave10 <> strSaveString Then
				If strPSExecSave09 <> strSaveString Then
					If strPSExecSave08 <> strSaveString Then
						If strPSExecSave07 <> strSaveString Then
							If strPSExecSave06 <> strSaveString Then
								If strPSExecSave05 <> strSaveString Then
									If strPSExecSave04 <> strSaveString Then
										objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave03", _
										strSaveString, "REG_SZ"
										PopulatePSExecSaveList()
										Exit Sub
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		If strPSExecSave02 = "" Then
			If strPSExecSave10 <> strSaveString Then
				If strPSExecSave09 <> strSaveString Then
					If strPSExecSave08 <> strSaveString Then
						If strPSExecSave07 <> strSaveString Then
							If strPSExecSave06 <> strSaveString Then
								If strPSExecSave05 <> strSaveString Then
									If strPSExecSave04 <> strSaveString Then
										If strPSExecSave03 <> strSaveString Then
											objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave02", _
											strSaveString, "REG_SZ"
											PopulatePSExecSaveList()
											Exit Sub
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		If strPSExecSave01 = "" Then
			If strPSExecSave10 <> strSaveString Then
				If strPSExecSave09 <> strSaveString Then
					If strPSExecSave08 <> strSaveString Then
						If strPSExecSave07 <> strSaveString Then
							If strPSExecSave06 <> strSaveString Then
								If strPSExecSave05 <> strSaveString Then
									If strPSExecSave04 <> strSaveString Then
										If strPSExecSave03 <> strSaveString Then
											If strPSExecSave02 <> strSaveString Then
												objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave01", _
												strSaveString, "REG_SZ"
												PopulatePSExecSaveList()
												Exit Sub
											End If
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		
		If strPSExecSave01 = strSaveString Then Exit Sub
		If strPSExecSave02 = strSaveString Then Exit Sub
		If strPSExecSave03 = strSaveString Then Exit Sub
		If strPSExecSave04 = strSaveString Then Exit Sub
		If strPSExecSave05 = strSaveString Then Exit Sub
		If strPSExecSave06 = strSaveString Then Exit Sub
		If strPSExecSave07 = strSaveString Then Exit Sub
		If strPSExecSave08 = strSaveString Then Exit Sub
		If strPSExecSave09 = strSaveString Then Exit Sub
		If strPSExecSave10 = strSaveString Then Exit Sub
		
		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave01", strSaveString, "REG_SZ"
		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave02", strPSExecSave01, "REG_SZ"
		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave03", strPSExecSave02, "REG_SZ"
		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave04", strPSExecSave03, "REG_SZ"
		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave05", strPSExecSave04, "REG_SZ"
		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave06", strPSExecSave05, "REG_SZ"
		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave07", strPSExecSave06, "REG_SZ"
		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave08", strPSExecSave07, "REG_SZ"
		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave09", strPSExecSave08, "REG_SZ"
		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave10", strPSExecSave09, "REG_SZ"
		
		PopulatePSExecSaveList()
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ExecutePSExecSaveCommands(booRun, booSave, booOpen)
    '#	PURPOSE........:	Executes the PSExec command specified
    '#	ARGUMENTS......:	booRun = boolean value to determine whether to
	'#						execute command (True) or not (False)
	'#						booSave = boolean value to determine whether to
	'#						Save command (True) or not (False)
	'#						booOpen = boolean value to determine whether to
	'#						Open command (True) or not (False)
    '#	EXAMPLE........:	ExecutePSExecSaveCommands PC1, True, False, False
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ExecutePSExecCommand(booRun, booSave, booOpen)
		strPSExec = "psexec.exe "
		strComputer = "\\" & strPC & " "
		If Commands.txtCommand.Value <> "" Then
			strCommand = Trim(Commands.txtCommand.Value)
			If InStr(strCommand, " ") > 0 AND InStr(strCommand,Chr(34)) = 0 Then
				strCommand = Chr(34) & strCommand & Chr(34) & " "
				Else
					strCommand = strCommand & " "
			End If
		End If
		If Commands.txtVariables.Value <> "" Then _
			strVariables =  " " & Commands.txtVariables.Value
		If Commands.cbxInteractive.Checked Then
			If InStr(strSwitches, "-i") = 0 Then strSwitches = strSwitches & "-i "
			Else
				strSwitches = Replace(strSwitches, "-i", "")
		End If
		If Commands.cbxCMD.Checked Then
			strCMD = "%COMSPEC% /k "
			Else 
				strCMD = ""
		End If
		If Commands.cbxNotLoadProf.Checked Then
			If InStr(strSwitches, "-e") = 0 Then strSwitches = strSwitches & "-e "
			Else
				strSwitches = Replace(strSwitches, "-e", "")
		End If
		If Commands.cbxSystem.Checked Then
			If InStr(strSwitches, "-s") = 0 Then strSwitches = strSwitches & "-s "
			Commands.cbxUsername.Checked = False
			Commands.cbxUsername.Disabled = True
			Else
				Commands.cbxUsername.Disabled = False
				strSwitches = Replace(strSwitches, "-s", "")
		End If
		If Commands.cbxCopyFile.Checked Then
			If InStr(strSwitches, "-c") = 0 Then strSwitches = strSwitches & "-c "
			Commands.cbxForceCopy.Disabled = False
			Commands.cbxNewerFile.Disabled = False
			Else
				strSwitches = Replace(strSwitches, "-c", "")
				Commands.cbxForceCopy.Disabled = True
				Commands.cbxForceCopy.Checked = False
				Commands.cbxNewerFile.Disabled = True
				Commands.cbxNewerFile.Checked = False
		End If
		If Commands.cbxForceCopy.Checked Then
			If InStr(strSwitches, "-f") = 0 Then strSwitches = strSwitches & "-f "
			Else
				strSwitches = Replace(strSwitches, "-f", "")
		End If
		If Commands.cbxNewerFile.Checked Then
			If InStr(strSwitches, "-v") = 0 Then strSwitches = strSwitches & "-v "
			Else
				strSwitches = Replace(strSwitches, "-v", "")
		End If
		If Commands.PriorityChooser.Value <> "0" then 
			strPriority = "-" & Commands.PriorityChooser.Value & " "
		End If
		If Commands.cbxUsername.Checked Then
			Commands.txtUsername.style.backgroundcolor = "white"
			Commands.txtPassword.style.backgroundcolor = "white"
			Commands.txtUsername.Disabled = False
			Commands.txtPassword.Disabled = False
			Commands.cbxNotLoadProf.Disabled = False
			Commands.cbxSystem.Checked = False
			Commands.cbxSystem.Disabled = True
			On Error Goto 0
			If PSExecSaves.SaveList.Value <> "" Then
				strSaveString = PSExecSaves.SaveList.Value
				arrSaveString = Split(strSaveString, "}{")
				strExecuteCommand = arrSaveString(1)
				arrExecuteCommand = Split(strExecuteCommand, " ")
				For i = 0 To UBound(arrExecuteCommand)
					If arrExecuteCommand(i) = "-p" Then
						strPassword = arrExecuteCommand(i + 1)
						Exit For
					End If
				Next
				Commands.txtPassword.Value = strPassword
			End If
			If Commands.txtUsername.Value = "" Then
				Commands.txtUsername.Value = strLocalLoggedOn
			End If
			strUserName = "-u " & Commands.txtUsername.Value & " "
			If Commands.txtPassword.Value <> "" Then
				strUserVisiblePass = "-p ****** "
				strUserPass = "-p " & Commands.txtPassword.Value & " "
			End If
			Else
				Commands.txtUsername.style.backgroundcolor = "#dddddd"
				Commands.txtPassword.style.backgroundcolor = "#dddddd"
				Commands.txtUsername.Disabled = True
				Commands.txtPassword.Disabled = True
				Commands.cbxNotLoadProf.Disabled = True
				Commands.cbxSystem.Disabled = False
				Commands.txtUsername.Value = ""
				Commands.txtPassword.Value = ""
		End If

	
		strVisibleCommand = strPSExec & strSwitches & strUserName & strUserVisiblePass & _
		strPriority & strComputer & strCommand & strVariables
	
		strExecuteCommand = strPSExec & strSwitches & strUserName & _
		strUserPass & strPriority & strComputer & strCMD & strCommand & strVariables
		
		PreviewCommand.InnerHTML = strVisibleCommand
		
		If booOpen = False Then	PSExecSaves.SaveList.Value = ""
	
		If booSave = True Then
			If strUserPass <> "" Then
				PWPrompt = MsgBox("The password provided will be stored in clear text." & _
				vbCrLf & vbCrLf & "Are you sure you wish to do this?",vbYesNo+vbExclamation, _
				"PC Management Utility")
				If PWPrompt = vbNo Then
					Exit Sub
				End If
			End If
			strPreviewSaveCommand = Replace(strVisibleCommand, strPC, "strPC")
			strExecuteSaveCommand = Replace(strExecuteCommand, strPC, "strPC")
			
			If strVariables <> "" AND strVariables <> " " Then
				strPreviewSaveCommand = Replace(strPreviewSaveCommand, strVariables, "")
				strExecuteSaveCommand = Replace(strExecuteSaveCommand, strVariables, "")
				Else
					strVariables = ""
			End If
			
			strSaveString = Trim(strPreviewSaveCommand) & "}{" & Trim(strExecuteSaveCommand) & _
			"}{" & Trim(strVariables)
			
			UpdatePSExecSaveCommands(strSaveString)
		End If
		
		If booRun = True Then
			objShell.Run "%COMSPEC% /c " & strExecuteCommand
		End If
		
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	RunCustomCommands()
    '#	PURPOSE........:	Allows a user to perform custom commands upon a
	'#						machine using a number of variables
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub RunCustomCommand()
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\cimv2")
		Set colOS = objWMIService.ExecQuery _
			("Select * from Win32_OperatingSystem")
		For Each objItem In colOS
			strWin = objItem.WindowsDirectory
		Next
		arrWin = Split(strWin, ":\")
		strWin = arrWin(UBound(arrWin))
		strRoot = GetRoot()
		
		strUsername = LoggedOnUser(strPC)
		arrUsername = Split(strUsername, "\")
		strUsername = LCase(arrUsername(UBound(arrUsername)))
		
		strHTML = "<script language=""VBScript"">" & _
		"On Error Resume Next" & _
		"<" & Chr(47) & "script>" & _
		"<b><u>" & strPC & "</u></b><p>" & _
		"Use the form below to perform custom commands on " & strPC & ".<p>" & _
		"You can use a number of variables to run commands as below:<p>" & _
		"<table>" & _
			"<tr>" & _
				"<td>" & _
					"Computer Name&nbsp;&nbsp;" & _
				"</td>" & _
				"<td>" & _
					strPC & "&nbsp;&nbsp;" & _
				"</td>" & _
				"<td>" & _
					"%C%" & _
				"</td>" & _
			"</tr>" & _
			"<tr>" & _
				"<td>" & _
					"System Root&nbsp;&nbsp;" & _
				"</td>" & _
				"<td>" & _
					"<span id=""Root"">" & strRoot & "&nbsp;&nbsp;</span>" & _
				"</td>" & _
				"<td>" & _
					"%R%" & _
				"</td>" & _
			"</tr>" & _
			"<tr>" & _
				"<td>" & _
					"Logged On User&nbsp;&nbsp;" & _
				"</td>" & _
				"<td>" & _
					"<span id=""UserLoggedOn"">" & strUsername & "&nbsp;&nbsp;</span>" & _
				"</td>" & _
				"<td>" & _
					"%U%" & _
				"</td>" & _
			"</tr>" & _
			"<tr>" & _
				"<td>" & _
					"Windows Directory&nbsp;&nbsp;" & _
				"</td>" & _
				"<td>" & _
					"<span id=""Windir"">" & strWin & "&nbsp;&nbsp;</span>" & _
				"</td>" & _
				"<td>" & _
					"%WIN%" & _
				"</td>" & _
			"</tr>" & _
		"</table>" & _
		"<form Name = ""CustomCommands"" Method = ""post"">" & _
		"<input type=""text"" name=""txtCommand"" size=""65"" " & _
		"onKeyUp=""ExecuteCustomCommand False""><p>" & _
		"Preview:" & _
		"<div style=background-color:white;width:580px;height:50px; id=""PreviewCommand""></div>" & _
		"<input id=runbutton class=""button"" type=""button"" value=""Clear"" name=""btnClear"" " & _
		"title=""Clear commands"" onClick=""ClearCustomCommand()"" " & _
		"onMouseDown=""ChangeButtonColour(btnClear)"" onMouseUp=""RevertButtonColour(btnClear)"">" & _
		"<input id=runbutton class=""button"" type=""button"" value=""Execute"" name=""btnCusExecute"" " & _
		"title=""Execute command on PC"" onClick=""ExecuteCustomCommand True""" & _
		"onMouseDown=""ChangeButtonColour(btnCusExecute)"" onMouseUp=""RevertButtonColour(btnCusExecute)""><br>"

		strFooter = "<span id=""BackToAction"" title=""Back to Actions list"" " & _
		"style=""text-decoration:underline;color:blue;cursor:pointer;position:absolute;bottom:50px"" " & _
		"onclick=""ShowPCActions()"">[..back to action list..]</span>"	
		
		DataArea.InnerHTML = strHTML & strFooter
		CustomCommands.txtCommand.Focus()
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ClearCustomCommands()
    '#	PURPOSE........:	Clears the Custom Commands form / preview
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ClearCustomCommand()
		PreviewCommand.InnerHTML = ""
		CustomCommands.txtCommand.Value = ""
		CustomCommands.txtCommand.Focus()
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ExecuteCustomCommands(booRun)
    '#	PURPOSE........:	Executes the custom command specified
    '#	ARGUMENTS......:	booRun = boolean value to determine whether to
	'#						execute command (True) or not (False)
    '#	EXAMPLE........:	RunCustomCommands(True)
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ExecuteCustomCommand(booRun)
		strCommand = CustomCommands.txtCommand.Value

		strCommand = Replace(LCase(strCommand), "%c%", strPC)
		strCommand = Replace(LCase(strCommand), "%u%", LCase(UserLoggedOn.InnerHTML))
		strCommand = Replace(LCase(strCommand), "%win%", LCase(Windir.InnerHTML))
		strCommand = Replace(LCase(strCommand), "%r%", LCase(Root.InnerHTML))
		
		strCommand = Replace(LCase(strCommand), "&nbsp;", "")
		PreviewCommand.InnerHTML = strCommand
		
		If booRun = True Then
			objShell.Run "%COMSPEC% /c " & strCommand
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	SystemRestore()
    '#	PURPOSE........:	Shows a list of the System Restore points 
	'#						and gives option of creating a new Restore point,
	'#						exporting all restore points and/or enabling /
	'#						disabling system restore
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub SystemRestore()
		Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strPC & "\root\default")		

		Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
		
		Set colRestore = objWMIService.ExecQuery _
			("Select * from SystemRestore")
			
		If colRestore.Count > 0 Then
			For Each objItem in colRestore
				strDescription = objItem.Description
				intSequenceNumber = objItem.SequenceNumber
				
				Select Case objItem.RestorePointType
					Case 0 
						strRestoreType = "Application installation"
					Case 1 
						strRestoreType = "Application uninstall"
					Case 6 
						strRestoreType = "Restore"
					Case 7 
						strRestoreType = "Checkpoint"
					Case 10 
						strRestoreType = "Device drive installation"
					Case 11 
						strRestoreType = "First run"
					Case 12 
						strRestoreType = "Modify settings"
					Case 13 
						strRestoreType = "Cancelled operation"
					Case 14 
						strRestoreType = "Backup recovery"
					Case Else 
						strRestoreType = "Unknown"
				End Select
				
				dtmConvertedDate.Value = objItem.CreationTime
				dtmCreationTime = dtmConvertedDate.GetVarDate
				
				strTxt =  strTxt & "Name: " & objItem.Description & vbCrLf & _
				"Number: " & objItem.SequenceNumber & vbCrLf & _
				"Restore Point Type: " & strRestoreType & vbCrLf & _
				"Time: " & dtmCreationTime & vbCrLf & _
				"--------------" & vbCrLf
			Next
			Else
				Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
				strPC & "\root\default:StdRegProv") 
				
				strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore"
				strValueName = "DisableSR"
				
				objReg.GetDWORDValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, dwValue

				If dwValue = 1 Then
					RestorePrompt = MsgBox("System Restore is Disabled on " & strPC & vbCrLf & vbCrLf & _
					"Would you like to Enable it?", vbQuestion+vbYesNo, _
					"PC Management Utility")
						If RestorePrompt = vbYes Then
							Err.Clear
							Set objItem = objWMIService.Get("SystemRestore")
							objItem.Enable("")
							If Err.Number <> 0 Then 
								MsgBox "There was an error enabling System Restore on " & _
								strPC, vbExclamation, "PC Management Utility"
								Exit Sub
								Else
									WaitMessage.InnerHTML = "<hr><div style=""text-align:left;font-weight=bold"">Please wait...</div>"
									MsgBox "System Restore is now Enabled on " & _
									strPC, vbInformation, "PC Management Utility"
									PauseScript(5000)
									SystemRestore()
									Exit Sub
							End If
						End If
					ElseIf dwValue = 0 Then
						MsgBox "System Restore is Enabled but there are currently no " & _
						"Restore Points on " & strPC, vbInformation, "PC Management Utility"
				End If
		End If
		
		strHTML = "<textarea name=""SystemRestoreTextArea"" rows=""10"" cols=""77""></textarea>"
		strHTML = strHTML & "<br><div style=""float:left;"">"
		strHTML = strHTML & "<input id=""RunButton"" class=""button"" type=""button"" value=""Create Restore Point"" " & _
		"name=""btnCreateRestorePoint"" style=""width:150px;"" onclick=""CreateSysRestorePoint()"" " & _
		"title=""Create a new System Restore Point"">"
		strHTML = strHTML & "<input id=""RunButton"" class=""button"" type=""button"" value=""Disable System Restore"" " & _
		"name=""btnDisableSysRestore"" style=""width:175px;"" onclick=""DisableSysRestore()"" " & _
		"title=""Disable System Restore on PC""></div>"
		strHTML = strHTML & "<div style=""float:right;"">"
		strHTML = strHTML & "<select name=""SysRestoreExport"" "
		strHTML = strHTML & "title=""Export the list of System Restore Points"" onChange=""ExportSysRestoreInfo()"">"
		strHTML = strHTML & "	<option value=""0"">Export to:</option>"
		strHTML = strHTML & "	<option value=""1"" title=""Export the System Restore Points to a Comma " & _
		"Seperated Values (csv) file"")>Export to csv</option>"
		strHTML = strHTML & "	<option value=""2"" title=""Export the System Restore Points to a formatted Excel " & _
		"(xls) spreadsheet"">Export to xls</option>"
		strHTML = strHTML & "	<option value=""3"" title=""Export the System Restore Points to a Web " & _
		"page (html) file"">Export to html</option>"
		strHTML = strHTML & "	<option value=""4"" title=""Export the System Restore Points to a Text " & _
		"(txt) file"">Export to txt</option>"
		strHTML = strHTML & "</select>"
		
		WaitMessage.InnerHTML = strHTML
		
		SystemRestoreTextArea.Value = "System Restore Points on " & strPC & vbCrLf & vbCrLf & _
		strTxt & vbCrLf & vbCrLf & "Total Restore Points: " & colRestore.Count
		
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	CreateSysRestorePoint()
    '#	PURPOSE........:	Creates a new Restore point
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub CreateSysRestorePoint()
		CONST MODIFY_SETTINGS = 12
		CONST BEGIN_SYSTEM_CHANGE = 100

		Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strPC & "\root\default")

		Set objItem = objWMIService.Get("SystemRestore")
		errResults = objItem.CreateRestorePoint _
		("Scripted restore point", MODIFY_SETTINGS, BEGIN_SYSTEM_CHANGE)
		If Err.Number = 0 Then
			MsgBox "A new System Restore Point has been created successfully.", vbInformation, "PC Management Utility"
			Else
				MsgBox "There was an error creating a new System Restore Point.", vbExclamation, "PC Management Utility"
		End If
		SystemRestore()
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	DisableSysRestore()
    '#	PURPOSE........:	Disables System Restore
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub DisableSysRestore()
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\default")
		Err.Clear
		
		Set objItem = objWMIService.Get("SystemRestore")
		objItem.Disable("")
		If Err.Number <> 0 Then 
			MsgBox "There was an error disabling System Restore on " & _
			strPC, vbExclamation, "PC Management Utility"
			Else
				MsgBox "System Restore is now Disabled on " & _
				strPC, vbInformation, "PC Management Utility"
				WaitMessage.InnerHTML = "<hr>"
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ExportSysRestoreInfo()
    '#	PURPOSE........:	Exports the list of System Restore Points
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------	
	Sub ExportSysRestoreInfo()
		On Error Resume Next
		Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
		'x = 0
		intColumnIndex = 9
		intColumnIndex2 = 11
		intColumnIndex3 = 30
		arrTextArea = Split(SystemRestoreTextArea.Value, vbCrLf)
		strTotal = arrTextArea(UBound(arrTextArea))
		
		Select Case SysRestoreExport.Value
			Case 1
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\RestorePoints" & strPC & ".csv",True)
				objFile.WriteLine "System Restore Points on " & strPC
				objFile.WriteLine ""
				objFile.WriteLine strTotal
				objFile.WriteLine ""
				objFile.WriteLine "Name,Number,Restore Point Type,Created On"
			Case 2
				Const xlContinuous = 1
				Const xlThin = 2
				Const xlAutomatic = -4105
				Const xlCenter = -4108
				
				strExcelPath = objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe\")
			   
				If strExcelPath = "" Then
					MsgBox "Unable to export. Excel does not appear to be installed.", vbExclamation, "PC Management Utility"
					Exit Sub
				End If
				
				Set objExcel = CreateObject("Excel.Application")
				objExcel.Visible = False
				Set objWorkBook = objExcel.WorkBooks.Add
				Set objWorksheet = objWorkbook.Worksheets(1)
				objExcel.DisplayAlerts = False
				For i = 1 to 3
					objWorkbook.Worksheets(2).Delete
				Next
				objExcel.DisplayAlerts = True
				objWorksheet.Name = "System Restore Points"
				
				objWorkSheet.Cells(1, 1) = "System Restore Points on " & strPC
				objWorkSheet.Cells(3, 1) = strTotal

				intStartRow = 6
				
				objWorkSheet.Cells(5, 1) = "Name"
				objWorkSheet.Cells(5, 2) = "Number"
				objWorkSheet.Cells(5, 3) = "Restore Point Type"
				objWorkSheet.Cells(5, 4) = "Created On"
			Case 3
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\RestorePoints" & strPC & ".htm",True)
				objFile.WriteLine "<style type=""text/css"">"
				objFile.WriteLine "body{background-color:#CEF0FF;}"
				objFile.WriteLine "table.export{border-width:1px;border-spacing:1px;border-style:solid;border-color:gray;border-collapse:collapse;}"
				objFile.WriteLine "table.export th{border-width:1px;padding:1px;border-style:solid;border-color:gray;padding:2px 7px 2px 7px;}"
				objFile.WriteLine "table.export td{border-width:1px;padding:1px;border-style:dotted;border-color:gray;padding:2px 7px 2px 7px;}"
				objFile.WriteLine ".backtotop a {font-size:0.9em;}"
				objFile.WriteLine "</style>"
				objFile.WriteLine "<div style=""font-weight:bold;""><a name =""top"">System Restore Points on " & strPC & "</a><p>"
				objFile.WriteLine strTotal & "</div><p>"
				objFile.WriteLine "<table class=""export"">"
				objFile.WriteLine "	<tr>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Name"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th>"
				objFile.WriteLine "			Number"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th>"
				objFile.WriteLine "			Restore Point Type"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th>"
				objFile.WriteLine "			Created On"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "	</tr>"
			Case 4		
				For i = 0 to UBound(arrTextArea)
					strLine = arrTextArea(i)
					If InStr(strLine, "Name:") > 0 Then
						strLine = Replace(strLine, "Name: ", "")
						strDescription = strLine
						If Len(strDescription) > intColumnIndex - 5 Then intColumnIndex = Len(strDescription) + 5
					End If
				Next

				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\RestorePoints" & strPC & ".txt",True)
				objFile.WriteLine "System Restore Points on " & strPC
				objFile.WriteLine ""
				objFile.WriteLine strTotal
				objFile.WriteLine ""
				objFile.WriteLine "Name" & _
				String(intColumnIndex - 4, " ") & "Number" & _
				String(5, " ") & "Restore Point Type" & _
				String(12, " ") & "Created On"
		End Select
		
		For i = 0 to UBound(arrTextArea)
			strLine = arrTextArea(i)
			If strLine <> "--------------" _
			AND strLine <> "" AND InStr(strLine, "System Restore Points on ") = 0 Then
				If InStr(strLine, "Name:") > 0 Then
					strLine = Replace(strLine, "Name: ", "")
					strDescription = strLine
				End If
				If InStr(strLine, "Number:") > 0 Then
					strLine = Replace(strLine, "Number: ", "")
					intSequenceNumber = strLine
				End If
				If InStr(strLine, "Restore Point Type:") > 0 Then
					strLine = Replace(strLine, "Restore Point Type: ", "")
					strRestoreType = strLine
				End If
				If InStr(strLine, "Time:") > 0 Then
					strLine = Replace(strLine, "Time: ", "")
					dtmCreationTime = strLine
					If IsDate(dtmCreationTime) Then dtmCreationTime = CDate(dtmCreationTime)
					
					Select Case SysRestoreExport.Value
						Case 1
							strCSV = strCSV & strDescription & "," & _
							intSequenceNumber & "," & strRestoreType & "," & _
							dtmCreationTime & vbCrLf
						Case 2
							objWorkSheet.Cells(intStartRow, 1) = strDescription
							objWorkSheet.Cells(intStartRow, 2) = intSequenceNumber
							objWorkSheet.Cells(intStartRow, 3) = strRestoreType
							objWorkSheet.Cells(intStartRow, 4) = dtmCreationTime
							intStartRow = intStartRow + 1
						Case 3
							objFile.WriteLine "	<tr>"
							objFile.WriteLine "		<td>"
							objFile.WriteLine "			" & strDescription
							objFile.WriteLine "		</td>"
							objFile.WriteLine "		<td style=""text-align:center;"">"
							objFile.WriteLine "			" & intSequenceNumber
							objFile.WriteLine "		</td>"
							objFile.WriteLine "		<td>"
							objFile.WriteLine "			" & strRestoreType
							objFile.WriteLine "		</td>"
							objFile.WriteLine "		<td>"
							objFile.WriteLine "			" & dtmCreationTime
							objFile.WriteLine "		</td>"
							objFile.WriteLine "	</tr>"
						Case 4
							strTxt = strTxt & strDescription & _
							String(intColumnIndex - Len(strDescription), " ") & _
							intSequenceNumber & String(intColumnIndex2 - Len(intSequenceNumber), " ") & _
							strRestoreType & String(intColumnIndex3 - Len(strRestoreType), " ") & _
							dtmCreationTime & vbCrLf
					End Select
				End If
			End If
		Next
		
		Select Case SysRestoreExport.Value
			Case 1
				objFile.WriteLine strCSV
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\RestorePoints" & strPC & ".csv"
			Case 2
				Set objRange = objWorkSheet.Range("A1:Z5")
				Set objRange2 = objWorkSheet.Range("A5:D" & intStartRow - 1)
				Set objRange3 = objWorkSheet.Range("B:B")
				
				objRange.Font.Bold = True
				objRange2.Borders.LineStyle = xlContinuous
				objRange2.Borders.Weight = xlThin
				objRange2.Borders.ColorIndex = xlAutomatic
				objRange3.HorizontalAlignment = xlCenter
				objWorksheet.Range("A6").Select
				objExcel.ActiveWindow.FreezePanes = "True"
				objWorksheet.Range("A1").Select
				
				objWorkSheet.Columns("A:ZZ").EntireColumn.AutoFit
				objExcel.DisplayAlerts = False
				objExcel.ActiveWorkbook.SaveAs(strTemp & "\SKB\RestorePoints" & strPC & ".xls")
				objExcel.Visible = True
				Set objExcel = Nothing
			Case 3
				objFile.WriteLine "</table>"
				objFile.WriteLine "<p class=""backtotop""><a href=""" & strHTMLTempDir & "/SKB/RestorePoints" & _
				strPC & ".htm#top"">[..back to top..]</a></p>"
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\RestorePoints" & strPC & ".htm"
			Case 4
				objFile.WriteLine strTxt
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\RestorePoints" & strPC & ".txt"
		End Select
		
		SysRestoreExport.Value = 0
	End Sub
	
	Sub ShutdownRestartPC()
		strHTML = "<hr><div style=""text-align:left;"">" & _
		"Which action would you like to perform?" & _
		"<p><input type=""radio"" name=""ShutdownType"" value=""1"" " & _
		"title=""Shutdown machine (" & strPC & ")"">Shutdown machine<br>" & _
		"<input type=""radio"" name=""ShutdownType"" value=""2"" " & _
		"title=""Restart machine (" & strPC & ")"">Restart machine<br>" & _
		"<input type=""radio"" name=""ShutdownType"" value=""0"" " & _
		"title=""Log Off machine (" & strPC & ")"">Log Off machine<p>" & _
		"<input id=""RunButton"" class=""button"" type=""button"" value=""Execute"" " & _
		"name=""btnShutdown"" onclick=""ShutdownRestartPC2()"" " & _
		"title=""Commence selected action on PC"">" & _
		"</div>"
		WaitMessage.InnerHTML = strHTML
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ShutdownRestartPC2()
    '#	PURPOSE........:	Shuts down / restarts / logs off PC
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	Will also run continuous ping to machine if
	'#						shutdown / restart requested
    '#--------------------------------------------------------------------------
	Sub ShutdownRestartPC2()
		For Each objButton in ShutdownType
			If objButton.Checked Then intSD = objButton.Value
		Next
		If IsNull(intSD) OR intSD = "" Then
			Exit Sub
		End If
		
		Set objWMIService = GetObject("winmgmts:{(Shutdown)}\\" & _ 
		strPC & "\root\cimv2")
		
		Set colOS = objWMIService.ExecQuery _
			("Select * from Win32_OperatingSystem")
			
		For Each objItem in colOS
			objItem.Win32Shutdown(intSD)
		Next
		
		If intSD = 1 Then
			MsgBox strPC & " is now shutting down",vbInformation, _
			"Shutdown / Restart / Log Off"
			ElseIf intSD = 2 Then
				MsgBox strPC & " is now restarting",vbInformation, _
				"Shutdown / Restart / Log Off"
		End If
		If intSD <> 0 Then
			On Error Resume Next
			Err.Clear
			'PingMachine False, False
			WaitMessage.InnerHTML = "<hr><p><br>Computer " & strPC & _
			" is <span id=OnlineStatus style=""font-weight:bold;""></span>"
			For i = 1 to 999999
				If Err.Number <> 0 Then Exit For
				If Reachable(strPC) = False Then
					OnlineStatus.style.color = "red"
					OnlineStatus.InnerHTML = "OFFLINE"
					Else
						OnlineStatus.style.color = "green"
						OnlineStatus.InnerHTML = "ONLINE"
				End If
				PauseScript(1000)
			Next
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ShowUserAccountsInfo()
    '#	PURPOSE........:	Displays the User Accounts tab
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ShowUserAccountsInfo()
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\cimv2") 
			
		Set colUserGroups = objWMIService.ExecQuery _
			("SELECT * FROM Win32_Group Where LocalAccount = True")
			
		For Each objItem In colUserGroups
			i = 0
			strGroup = objItem.Name
			Set objGroup = GetObject("WinNT://" & strPC & "/" & strGroup)
			For Each Member In objGroup.Members
				i = i + 1
				strGroupMember = Member.Name
				If i > 1 Then
					strTxt = strTxt & vbCrLf & vbTab & strGroupMember
					Else
						strTxt = strTxt & vbCrLf & "------------------------------------------" & _
						vbCrLf & strGroup & vbCrLf & vbTab & strGroupMember
				End If
			Next
		Next
		
		strTxt = strTxt & vbCrLf & "------------------------------------------"
		
		strHTML = "<textarea name=""UserAccountsListTextArea"" rows=""10"" " & _
		"cols=""77""></textarea>"
		
		strHTML = strHTML & "<br><div style=""text-align:right;"">"
		strHTML = strHTML & "<select name=""UserAccsExport"" "
		strHTML = strHTML & "title=""Export the User Accounts list"" onChange=""ExportUserAccountsInfo()"">"
		strHTML = strHTML & "	<option value=""0"">Export to:</option>"
		strHTML = strHTML & "	<option value=""1"" title=""Export the User Accounts list to a Comma " & _
		"Seperated Values (csv) file"")>Export to csv</option>"
		strHTML = strHTML & "	<option value=""2"" title=""Export the User Accounts list to a formatted Excel " & _
		"(xls) spreadsheet"">Export to xls</option>"
		strHTML = strHTML & "	<option value=""3"" title=""Export the User Accounts list to a Web " & _
		"page (html) file"">Export to html</option>"
		strHTML = strHTML & "	<option value=""4"" title=""Export the User Accounts list to a Text " & _
		"(txt) file"">Export to txt</option>"
		strHTML = strHTML & "	</select></div>"
		
		WaitMessage.InnerHTML = strHTML
		
		UserAccountsListTextArea.Value = "User Accounts on " & strPC & vbCrLf & vbCrLf & _
		strTxt
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ExportUserAccountsInfo()
    '#	PURPOSE........:	Export the details for the User Accounts
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ExportUserAccountsInfo()
		On Error Resume Next
		
		Select Case UserAccsExport.Value
			Case 1
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\UserAccountsDetails" & strPC & ".csv",True)
				objFile.WriteLine "User Accounts Items on " & strPC
				objFile.WriteLine ""
				objFile.WriteLine "Group Name,User Name"
			Case 2
				Const xlContinuous = 1
				Const xlThin = 2
				Const xlAutomatic = -4105
				
				strExcelPath = objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe\")
			   
				If strExcelPath = "" Then
					MsgBox "Unable to export. Excel does not appear to be installed.", vbExclamation, "PC Management Utility"
					Exit Sub
				End If
				
				Set objExcel = CreateObject("Excel.Application")
				objExcel.Visible = False
				Set objWorkBook = objExcel.WorkBooks.Add
				Set objWorksheet = objWorkbook.Worksheets(1)
				objExcel.DisplayAlerts = False
				For i = 1 to 3
					objWorkbook.Worksheets(2).Delete
				Next
				objExcel.DisplayAlerts = True
				objWorksheet.Name = "User Accounts Details"
				
				objWorkSheet.Cells(1, 1) = "User Accounts on " & strPC

				intStartRow = 4
				
				objWorkSheet.Cells(3, 1) = "Group Name"
				objWorkSheet.Cells(3, 2) = "User Name"
			Case 3
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\UserAccountsDetails" & strPC & ".htm",True)
				objFile.WriteLine "<style type=""text/css"">"
				objFile.WriteLine "body{background-color:#CEF0FF;}"
				objFile.WriteLine "table.export{border-width:1px;border-spacing:1px;border-style:solid;border-color:gray;border-collapse:collapse;}"
				objFile.WriteLine "table.export th{border-width:1px;padding:1px;border-style:solid;border-color:gray;padding:2px 7px 2px 7px;}"
				objFile.WriteLine "table.export td{border-width:1px;padding:1px;border-style:dotted;border-color:gray;padding:2px 7px 2px 7px;}"
				objFile.WriteLine ".backtotop a {font-size:0.9em;}"
				objFile.WriteLine "</style>"
				objFile.WriteLine "<div style=""font-weight:bold;""><a name =""top"">User Accounts on " & strPC & "</a><p></div>"
				objFile.WriteLine "<table class=""export"">"
				objFile.WriteLine "	<tr>"
				objFile.WriteLine "		<th style=""text-align:left;"">"
				objFile.WriteLine "			Group Name"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "		<th>"
				objFile.WriteLine "			User Name"
				objFile.WriteLine "		</th>"
				objFile.WriteLine "	</tr>"
			Case 4
				intColumnIndex = 15
				intLineLength = 0
				booNewGroup = False
		
				arrTextArea = Split(UserAccountsListTextArea.Value, vbCrLf)
				For i = 0 to UBound(arrTextArea)
					strLine = arrTextArea(i)
					If strLine <> "------------------------------------------" _
					AND strLine <> "" AND InStr(strLine, "User Accounts on ") = 0 Then
						If booNewGroup = True Then
							strGroupName = strLine
							If Len(strGroupName) > intColumnIndex - 5 Then intColumnIndex = Len(strGroupName) + 5
							booNewGroup = False
							Else
								strGroupUser = Trim(strLine)
								strGroupUser = Replace(strGroupUser, vbTab, "")
						End If
						If Len(strGroupUser) > intLineLength Then intLineLength = Len(strGroupUser)
						Else
							booNewGroup = True
					End If			
				Next
				
				intLineLength = intLineLength + intColumnIndex + 1
				
				Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\UserAccountsDetails" & strPC & ".txt",True)
				objFile.WriteLine "User Accounts on " & strPC
				objFile.WriteLine ""
				objFile.WriteLine String(intLineLength, "-")
				objFile.WriteLine "Group Name" & _
				String(intColumnIndex - 10, " ") & "User Name"
		End Select
		
		arrTextArea = Split(UserAccountsListTextArea.Value, vbCrLf)
		booNewGroup = False
		
		For i = 0 to UBound(arrTextArea)
			strLine = arrTextArea(i)
			If strLine <> "------------------------------------------" _
			AND strLine <> "" AND InStr(strLine, "User Accounts on ") = 0 Then
				If booNewGroup = False Then
					strGroupUser = Trim(strLine)
					strGroupUser = Replace(strGroupUser, vbTab, "")
					
					Select Case UserAccsExport.Value
						Case 1
							strCSV = strCSV & strGroupName & "," & _
							strGroupUser & vbCrLf
						Case 2
							objWorkSheet.Cells(intStartRow, 1) = strGroupName
							objWorkSheet.Cells(intStartRow, 2) = strGroupUser
							intStartRow = intStartRow + 1
						Case 3
							objFile.WriteLine "	<tr>"
							objFile.WriteLine "		<td>"
							objFile.WriteLine "			" & strGroupName
							objFile.WriteLine "		</td>"
							objFile.WriteLine "		<td>"
							objFile.WriteLine "			" & strGroupUser
							objFile.WriteLine "		</td>"
							objFile.WriteLine "	</tr>"
						Case 4
							strTxt = strTxt & strGroupName & _
							String(intColumnIndex - Len(strGroupName), " ") & _
							strGroupUser & vbCrLf
					End Select
					
					Else
						If UserAccsExport.Value = 4 Then _
							strTxt = strTxt & String(intLineLength, "-") & vbCrLf
						
						strGroupName = strLine
						booNewGroup = False
				End If
				Else
					booNewGroup = True
			End If
		Next

		Select Case UserAccsExport.Value
			Case 1
				objFile.WriteLine strCSV
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\UserAccountsDetails" & strPC & ".csv"
			Case 2
				Set objRange = objWorkSheet.Range("A1:Z3")
				Set objRange2 = objWorkSheet.Range("A3:B" & intStartRow - 1)
				
				objRange.Font.Bold = True
				objRange2.Borders.LineStyle = xlContinuous
				objRange2.Borders.Weight = xlThin
				objRange2.Borders.ColorIndex = xlAutomatic
				objWorksheet.Range("A4").Select
				objExcel.ActiveWindow.FreezePanes = "True"
				objWorksheet.Range("A1").Select
				
				objWorkSheet.Columns("A:ZZ").EntireColumn.AutoFit
				objExcel.DisplayAlerts = False
				objExcel.ActiveWorkbook.SaveAs(strTemp & "\SKB\UserAccountsDetails" & strPC & ".xls")
				objExcel.Visible = True
				Set objExcel = Nothing
			Case 3
				objFile.WriteLine "</table>"
				objFile.WriteLine "<p class=""backtotop""><a href=""" & strHTMLTempDir & "/SKB/UserAccountsDetails" & _
				strPC & ".htm#top"">[..back to top..]</a></p>"
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\UserAccountsDetails" & strPC & ".htm"
			Case 4
				strTxt = strTxt & String(intLineLength, "-")
				objFile.WriteLine strTxt
				objFile.Close
				Set objFile = Nothing
				objShell.Run strTemp & "\SKB\UserAccountsDetails" & strPC & ".txt"
		End Select
		
		UserAccsExport.Value = 0
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ViewProfiles()
    '#	PURPOSE........:	Shows a list of user profiles on the PC and
	'#						optionally exports list to text file
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ViewProfiles()
		On Error Resume Next
		intProfileCount = 0
		intColumnIndex = 17
		intColumnIndex2 = 17
		intColumnIndex3 = 16
		intProgDone = 0
		
		Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\default:StdRegProv") 
		
		strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
		objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys
		
		For Each objItem In arrSubkeys
			strValueName = "ProfileImagePath"
			strSubPath = strKeyPath & "\" & objItem
			objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath,strValueName,strValue
			If strValue <> "" Then
				intProfileCount = intProfileCount + 1
			End If
		Next
		
		intProgTotal = intProfileCount * 7
		intProgMult = 100 / intProgTotal
 
		For Each objItem In arrSubkeys
			strValueName = "ProfileImagePath"
			strSubPath = strKeyPath & "\" & objItem
			objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath,strValueName,strValue
			If strValue <> "" Then
				strRoot = GetRoot()
				strProfPath = Right(strValue,Len(strValue) - 2)
				strProfPath = strRoot & strProfPath
				intProgDone = intProgDone + 1	'1
				UpdateProgressBar intProgMult,intProgDone,intProgTotal
				arrPath = Split(strValue, "\")
				strProfName = arrPath(Ubound(arrPath))
				If Len(strProfName) > intColumnIndex Then 
					intColumnIndex = Len(strProfName) + 5
				End If
				intProgDone = intProgDone + 1	'2
				UpdateProgressBar intProgMult,intProgDone,intProgTotal
				If objFSO.FolderExists(strProfPath) Then
					Set objProfileFolder = objFSO.GetFolder(strProfPath)
					If objFSO.FileExists(objProfileFolder & "\NTUSER.DAT") Then
						strNTUser = strProfPath & "\NTUSER.DAT"
						Set objNTUserFile = objFSO.GetFile(strNTUser)
						dtmProfLastUsed = objNTUserFile.DateLastModified
						Else
							dtmProfLastUsed = objProfileFolder.DateLastModified
					End If
					intProgDone = intProgDone + 1	'3
					UpdateProgressBar intProgMult,intProgDone,intProgTotal
					strFolderSize = ConvertToDiskSize(objProfileFolder.Size)
					If Len(strFolderSize) > intColumnIndex3 Then 
						intColumnIndex3 = Len(strFolderSize) + 5
					End If
					If InStr(LCase(strRemoteLoggedOn),LCase(strProfName)) > 0 Then
						dtmProfLastUsed = Now
					End If
					intProgDone = intProgDone + 1	'4
					UpdateProgressBar intProgMult,intProgDone,intProgTotal
					If DatePart("d", dtmProfLastUsed) < 10 Then
						strProfLastUsedD = "0" & DatePart("d", dtmProfLastUsed)
						Else
							strProfLastUsedD = DatePart("d", dtmProfLastUsed)
					End If
					intProgDone = intProgDone + 1	'5
					UpdateProgressBar intProgMult,intProgDone,intProgTotal
					If DatePart("m", dtmProfLastUsed) < 10 Then
						strProfLastUsedM = "0" & DatePart("m", dtmProfLastUsed)
						Else
							strProfLastUsedM = DatePart("m", dtmProfLastUsed)
					End If
					intProgDone = intProgDone + 1	'6
					UpdateProgressBar intProgMult,intProgDone,intProgTotal
					dtmProfLastUsed = strProfLastUsedD & "/" & _
					strProfLastUsedM & "/" & DatePart("yyyy", dtmProfLastUsed)
					Else
						intProgDone = intProgDone + 4	'6
						UpdateProgressBar intProgMult,intProgDone,intProgTotal
						strValue = "Path Deleted - Please run Delete Old User Profiles routine"
						strFolderSize = "n/a"
						dtmProfLastUsed = "n/a"
				End If
				If Len(strValue) > intColumnIndex2 Then 
					intColumnIndex2 = Len(strValue) + 5
				End If
				intProgDone = intProgDone + 1	'7
				UpdateProgressBar intProgMult,intProgDone,intProgTotal
				strCSV = strCSV & strProfName & "," & _
				strValue & "," & strFolderSize & "," & _
				dtmProfLastUsed & vbCrLf
				strMsg = strMsg & strProfName & ":  " & strValue & vbCrLf
			End If
			dtmProfLastUsed = ""
			strProfName = ""
			strFolderSize = ""
		Next
		strMsg = strMsg & vbCrLf & intProfileCount & " profiles"
		MsgBox strMsg, vbInformation, "Profile Paths"
		ViewProfPrompt = MsgBox("Would you like to export this info to a csv file?", vbQuestion+vbYesNo, _
		"Profile Paths")
		If ViewProfPrompt = vbYes Then
			Set objFile = objFSO.CreateTextFile(strTemp & "\SKB\profiles.csv",True)
			objFile.WriteLine "Profiles on " & strPC
			objFile.WriteLine ""
			objFile.WriteLine "Total Profiles: " & intProfileCount
			objFile.WriteLine ""
			objFile.WriteLine "Profile Name,Profile Path,Folder Size,Last Modified"
			objFile.WriteLine strCSV
			objShell.Run strTemp & "\SKB\profiles.csv"
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	ScriptLog(strLogDetails,intLogType)
    '#	PURPOSE........:	Exports profile copy log into formatted text file
    '#	ARGUMENTS......:	strLogDetails = Text to show on log
	'#						intLogType = log type (eg. 0 = SUCCESS)
    '#	EXAMPLE........:	ScriptLog("Copied ok",0)
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Sub ScriptLog(strLogDetails,intLogType)
		On Error Resume Next
		Const intColumnIndex1 = 30
		Const intColumnIndex2 = 23 
 
		If IsNumeric(intLogType) Then 
			Select Case CInt(intLogType)
				Case 0
					strLogType = "SUCCESS"
				Case 1
					strLogType = "ERROR"
				Case 2
					strLogType = "WARNING"
				Case 4
					strLogType = "INFORMATION"
				Case Else
					strLogType = "UNKNOWN"
			End Select
			Else
				strLogType = "INCORRECT_TYPE"
		End If       
 
		If NOT objFSO.FolderExists("\\" & strPC & "\C$\SKB") Then
			Set objFolder = objFSO.CreateFolder("\\" & strPC & "\C$\SKB")
		End If
		
		strLogFileLocation = "\\" & strPC & "\C$\SKB\ProfileCopy\"
		strLogFileName = "ProfileCopy-" & strGlobalLogFileName & "-" & strFileDate & ".log"
 
		If NOT objFSO.FolderExists(strLogFileLocation) Then
			Set objFolder = objFSO.CreateFolder(strLogFileLocation)
		End If

		strLogDateTime = "[" & Now() & "]"
      
		strError = strError & strLogDateTime & String(intColumnIndex1 - Len(strLogDateTime), " ")
		strError = strError & strLogType & String(intColumnIndex2 - Len(strLogType), " ")
		strError = strError & strLogDetails & vbNewLine
      
		If strLogType = "INCORRECT_TYPE" Then 
			strError = strError & strLogDateTime & String(intColumnIndex1 - Len(strLogDateTime), " ")
			strError = strError & strLogType & String(intColumnIndex2 - Len(strLogType), " ")
			strError = strError & "An Incorrect_Type means an invalid Logtype number " & _
			"was passed to the ScriptLog procedure" & vbNewLine
		End If 
 
		If NOT objFSO.FileExists(strLogFileLocation & strLogFileName) Then 
			Set objFile = objFSO.OpenTextFile(strLogFileLocation & strLogFileName, ForWriting, True)
 
			If Err.Number <> 0 Then
			MsgBox "Error opening file: " & strLogFileLocation & _
			strLogFileName,vbExclamation, "Copy Profile"
			Exit Sub
			End If
			Else
				Set objFile = objFSO.OpenTextFile(strLogFileLocation & strLogFileName, ForAppending, False)
 
				If Err.Number <> 0 Then
					MsgBox "Error opening file: " & strLogFileLocation & _
					strLogFileName,vbExclamation, "Copy Profile"
					Exit Sub
				End If
			End If

		objFile.Write(strError)
		objFile.Close
	
		Set objFile = Nothing
		Err.Clear 
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	PauseScript(intPause)
    '#	PURPOSE........:	Pauses the script
    '#	ARGUMENTS......:	intPause = number of milliseconds to pause
    '#	EXAMPLE........:	PauseScript(1000)
    '#	NOTES..........:	Above example will pause script for 1 second
    '#--------------------------------------------------------------------------
	Sub PauseScript(intPause)
		objShell.Run "%COMSPEC% /c ping -w " & intPause & " -n 1 1.0.0.0", 0, True
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	MigrateData()
    '#	PURPOSE........:	Migrates data from version 3.1 to 3.2
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	Preserves Setup options, PSExec commands and 
	'#						IP ranges. Will only run once.
    '#--------------------------------------------------------------------------	
	Sub MigrateData()
		On Error Resume Next
		booDeleteTemp = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\booDeleteTemp")
		intSearchView = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\intSearchView")
		strQueryChoices = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\strQueryChoices")
		strInvDirectory = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\strInvDirectory")
		IP1A = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IP1A")
		IP2A = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IP2A")
		IP3A = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IP3A")
		IP4A = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IP4A")
		IP5A = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IP5A")
		IP1B = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IP1B")
		IP2B = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IP2B")
		IP3B = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IP3B")
		IP4B = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IP4B")
		IP5B = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\IP5B")
		strPSExecSave01 = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\strPSExecSave01")
		strPSExecSave02 = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\strPSExecSave02")
		strPSExecSave03 = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\strPSExecSave03")
		strPSExecSave04 = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\strPSExecSave04")
		strPSExecSave05 = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\strPSExecSave05")
		strPSExecSave06 = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\strPSExecSave06")
		strPSExecSave07 = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\strPSExecSave07")
		strPSExecSave08 = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\strPSExecSave08")
		strPSExecSave09 = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\strPSExecSave09")
		strPSExecSave10 = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\strPSExecSave10")

		If booDeleteTemp <> "" Then
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\Setup\booDeleteTemp", _
			booDeleteTemp, "REG_SZ"
		End If
		If intSearchView <> "" Then
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\Setup\intSearchView", _
			intSearchView, "REG_SZ"
		End If
		If strQueryChoices <> "" Then
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\Setup\strQueryChoices", _
			strQueryChoices, "REG_SZ"
		End If
		If strInvDirectory <> "" Then
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\Setup\strInvDirectory", _
			strInvDirectory, "REG_SZ"
		End If
		If IP1A <> "" Then
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP1A", _
			IP1A, "REG_SZ"
		End If
		If IP2A <> "" Then
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP2A", _
			IP2A, "REG_SZ"
		End If
		If IP3A <> "" Then
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP3A", _
			IP3A, "REG_SZ"
		End If
		If IP4A <> "" Then
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP4A", _
			IP4A, "REG_SZ"
		End If
		If IP5A <> "" Then
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP5A", _
			IP5A, "REG_SZ"
		End If
		If IP1B <> "" Then
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP1B", _
			IP1B, "REG_SZ"
		End If
		If IP2B <> "" Then
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP2B", _
			IP2B, "REG_SZ"
		End If
		If IP3B <> "" Then	
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP3B", _
			IP3B, "REG_SZ"
		End If
		If IP4B <> "" Then	
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP4B", _
			IP4B, "REG_SZ"
		End If
		If IP5B <> "" Then	
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\IPRanges\IP5B", _
			IP5B, "REG_SZ"
		End If
		If strPSExecSave01 <> "" Then	
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave01", _
			strPSExecSave01, "REG_SZ"
		End If
		If strPSExecSave02 <> "" Then	
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave02", _
			strPSExecSave02, "REG_SZ"
		End If
		If strPSExecSave03 <> "" Then	
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave03", _
			strPSExecSave03, "REG_SZ"
		End If
		If strPSExecSave04 <> "" Then	
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave04", _
			strPSExecSave04, "REG_SZ"
		End If
		If strPSExecSave05 <> "" Then	
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave05", _
			strPSExecSave05, "REG_SZ"
		End If
		If strPSExecSave06 <> "" Then	
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave06", _
			strPSExecSave06, "REG_SZ"
		End If
		If strPSExecSave07 <> "" Then	
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave07", _
			strPSExecSave07, "REG_SZ"
		End If
		If strPSExecSave08 <> "" Then	
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave08", _
			strPSExecSave08, "REG_SZ"
		End If
		If strPSExecSave09 <> "" Then	
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave09", _
			strPSExecSave09, "REG_SZ"
		End If
		If strPSExecSave10 <> "" Then	
			objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\PSExecSaves\strPSExecSave10", _
			strPSExecSave10, "REG_SZ"
		End If

		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\booDeleteTemp")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\intSearchView")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\strQueryChoices")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\strInvDirectory")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\IP1A")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\IP2A")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\IP3A")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\IP4A")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\IP5A")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\IP1B")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\IP2B")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\IP3B")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\IP4B")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\IP5B")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\strPSExecSave01")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\strPSExecSave02")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\strPSExecSave03")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\strPSExecSave04")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\strPSExecSave05")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\strPSExecSave06")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\strPSExecSave07")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\strPSExecSave08")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\strPSExecSave09")
		objShell.RegDelete("HKCU\Software\SKB\PCManagementUtil\strPSExecSave10")

		objShell.RegWrite "HKCU\Software\SKB\PCManagementUtil\FirstRun", _
		"1", "REG_SZ"
	End Sub
	
	'#--------------------------------------------------------------------------
    '#	SUBROUTINE.....:	Window_onLoad()
    '#	PURPOSE........:	Runs on application opening
    '#	ARGUMENTS......:	
    '#	EXAMPLE........:	
    '#	NOTES..........:	Sets Window size and sets focus to PCSearch box,
	'#						creates %TEMP%\SKB folder and reads / applies app
	'#						defaults from registry
    '#--------------------------------------------------------------------------
	Sub Window_onLoad()
		On Error Resume Next
		PauseScript(1)
		intFirstRun = objShell.RegRead("HKCU\Software\SKB\PCManagementUtil\FirstRun")
		If intFirstRun <> 1 Then MigrateData()
		strRunAsUser = objShell.ExpandEnvironmentStrings("%USERNAME%")
		RunAs.InnerHTML = "<span style=""cursor:default"">Running as user: " & strRunAsUser & "</span>"
		UtilVersion.InnerHTML = "<span style=""cursor:default"">Version: "  & objPCManage.Version & _
		"&nbsp;&nbsp;&nbsp;</span><span title=""View information about the Computer Management Utility"" " & _
		"style=""color:black;cursor:pointer;font-style:normal;"" onclick=""About()"">?</span>"
		strLocalLoggedOn = LoggedOnUser(".")
		Err.Clear
		strLocalSID = GetSIDFromUser(strLocalLoggedOn)
		
		If NOT objFSO.FolderExists(strTemp & "\SKB") Then
			objFSO.CreateFolder(strTemp & "\SKB")
		End If
		intSearchView = objShell.RegRead("HKEY_USERS\" & strLocalSID & "\Software\SKB\PCManagementUtil\Setup\intSearchView")
		If Err.Number = 0 Then
			strRegStart = "HKEY_USERS\" & strLocalSID
			Else
				strRegStart = "HKCU\"
		End If
		
		If intSearchView = "" Then
			objShell.RegWrite strRegStart & "\Software\SKB\PCManagementUtil\Setup\intSearchView", _
			"1", "REG_SZ"
			intSearchView = "1"
		End If
		SearchView.Value = intSearchView
		ChangeSearchView(intSearchView)
		LoadingUtil.style.visibility = "hidden"
		LoadingUtil.style.display = "none"
		UtilLayout.style.visibility = "visible"
		UtilLayout.style.display = "inline"
		PCSearch.focus()
    End Sub
	
	Sub Window_onunLoad
		On Error Resume Next
		
		booDeleteTemp = objShell.RegRead(strRegStart & "\Software\SKB\PCManagementUtil\Setup\booDeleteTemp")
		If booDeleteTemp = "" Then
			objShell.RegWrite strRegStart & "\Software\SKB\PCManagementUtil\Setup\booDeleteTemp", _
			"0", "REG_SZ"
			booDeleteTemp = 0
		End If
		If booDeleteTemp = 1 Then
			If objFSO.FolderExists ("c:\skb") Then
				objFSO.DeleteFolder "c:\skb", True
			End If
			strInvDirectory = objShell.RegRead(strRegStart & "\Software\SKB\PCManagementUtil\Setup\strInvDirectory")
			
			If strInvDirectory <> "" Then
				If objFSO.FolderExists(strInvDirectory) Then
					objFSO.DeleteFolder(strInvDirectory)
				End If
			End If
			objFSO.DeleteFolder(strTemp & "\SKB")
		End If
	End Sub
	
	'#--------------------------------------------------------------------------
	'#  FUNCTION.......:	Reachable(strComp)
	'#  PURPOSE........:	Checks whether the remote PC is online
	'#  ARGUMENTS......:	
	'#  EXAMPLE........:	Reachable(PC1)
	'#  NOTES..........:  
	'#--------------------------------------------------------------------------
	Function Reachable(strComp)
		Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
		Set colPing = objWMIService.ExecQuery _
			("Select * from Win32_PingStatus Where Address = '" & strComp & "'")
		For Each objItemR in colPing
			If IsNull(objItemR.StatusCode) Or objItemR.StatusCode <> 0 Then
				Reachable = False
				Else
					Reachable = True
			End If
		Next
	End Function
	
	'#--------------------------------------------------------------------------
	'#  FUNCTION.......:	GetSize(strFolder)
	'#  PURPOSE........:	Gets the size of specified folder
	'#  ARGUMENTS......:	strFolder = full path to the folder
	'#  EXAMPLE........:	GetSize("c:\Temp")
	'#  NOTES..........:	
	'#--------------------------------------------------------------------------
	Function GetSize(strFolder)
		On Error Resume Next
		Set objFolder = objFSO.GetFolder(strFolder)
		GetSize = objFolder.Size
	End Function

	'#--------------------------------------------------------------------------
	'#  FUNCTION.......:	GetSIDFromUser(strUserName)
	'#  PURPOSE........:	Gets the SID from the specified user
	'#  ARGUMENTS......:	strUserName = Username for which to retrieve SID
	'#  EXAMPLE........:	GetSIDFromUser("acmegroup\jimbob")
	'#  NOTES..........:  
	'#--------------------------------------------------------------------------
	Function GetSIDFromUser(strUserName)
		If InStr(strUserName, "\") > 0 Then
			arrUserName = Split(strUserName, "\")
			strDomainName = arrUserName(LBound(arrUserName))
			strUserName = arrUserName(UBound(arrUserName))
			Else
				strDomainName = CreateObject("WScript.Network").UserDomain
		End If

		On Error Resume Next
		
		Set objWMIService2 = GetObject("winmgmts:\\.\root\cimv2")
		
		Set objAccount = objWMIService2.Get _
			("Win32_UserAccount.Name='" & strUserName & "',Domain='" & _
			strDomainName & "'")
			
		If Err = 0 Then 
			Result = objAccount.SID 
			Else 
				Result = ""
		End If
		
		On Error GoTo 0

		GetSIDFromUser = Result
	End Function

	'#--------------------------------------------------------------------------
	'#  FUNCTION.......:	DecodeKey(arriValues)
	'#  PURPOSE........:	Decodes any Microsoft Product Key
	'#  ARGUMENTS......:	arriValues = key values to be decoded
	'#  EXAMPLE........:	DecodeKey(arriValues)
	'#  NOTES..........:	Props go to Parabellum for this Function, though
	'#						have modified to only return Product Key for
	'#						versatility
	'#--------------------------------------------------------------------------
	Function DecodeKey(arriValues)
		arrFoundKeys = Array()
		arrDPID = Array()
		For i = 52 to 66
			ReDim Preserve arrDPID(UBound(arrDPID) + 1)
			arrDPID(UBound(arrDPID)) = arriValues(i)
		Next
		arrChars = Array("B","C","D","F","G","H","J","K","M","P", _
		"Q","R","T","V","W","X","Y","2","3","4","6","7","8","9")
		For i = 24 To 0 Step -1
			k = 0
			For j = 14 To 0 Step -1
				k = k * 256 Xor arrDPID(j)
				arrDPID(j) = Int(k / 24)
				k = k Mod 24
			Next
			strProductKey = arrChars(k) & strProductKey
			If i Mod 5 = 0 And i <> 0 Then strProductKey = "-" & strProductKey
		Next
		
		ReDim Preserve arrFoundKeys(UBound(arrFoundKeys) + 1)
		arrFoundKeys(UBound(arrFoundKeys)) = strProductKey
		strKey = UBound(arrFoundKeys)
		DecodeKey = arrFoundKeys(strKey)	
	End Function
	
	'#--------------------------------------------------------------------------
	'#  FUNCTION.......:	CheckIP(strIP,strIP2)
	'#  PURPOSE........:	Validates both entered IP addresses
	'#  ARGUMENTS......:	strIP = IP Address entered in top 4 boxes
	'#						strIP2 = IP Address entered in bottom 4 boxes
	'#  EXAMPLE........:	Check IP(200.200.10.4,200.200.10.5)
	'#  NOTES..........:	
	'#--------------------------------------------------------------------------
	Function CheckIP(strIP,strIP2)
		CheckIP = True
		arrCheckIP = Split(strIP,".")
		arrCheckIP2 = Split(strIP2,".")
		int1A = CInt(arrCheckIP(0))
		int1B = CInt(arrCheckIP2(0))
		int2A = CInt(arrCheckIP(1))
		int2B = CInt(arrCheckIP2(1))
		int3A = CInt(arrCheckIP(2))
		int3B = CInt(arrCheckIP2(2))
		int4A = CInt(arrCheckIP(3))
		int4B = CInt(arrCheckIP2(3))
		
		If int1B > 255 OR int2B > 255 OR int3B > 255 OR int4B > 255 _
		OR int1A > 255 OR int2A > 255 OR int3A > 255 OR int4A > 255 Then
			CheckIP = False
			Else
				If int4A > int4B Then
					If int1A < int1B OR int2A < int2B OR int3A < int3B Then
						Exit Function
						Else
							CheckIP = False
							Exit Function
					End If
				End If
				If int3A > int3B Then
					If int1A < int1B OR int2A < int2B Then
						Exit Function
						Else
							CheckIP = False
							Exit Function
					End If
				End If
				If int2A > int2B Then
					If int1A < int1B OR int2A < int2B Then
						Exit Function
						Else
							CheckIP = False
							Exit Function
					End If
				End If
				If int1A > int1B Then
					CheckIP = False
				End If
		End If
	End Function
	
	'#--------------------------------------------------------------------------
	'#  FUNCTION.......:	NewIP(strIP)
	'#  PURPOSE........:	Increments the IP address
	'#  ARGUMENTS......:	strIP = IP address to be incremented
	'#  EXAMPLE........:	NewIP(200.200.10.4)
	'#  NOTES..........:  
	'#--------------------------------------------------------------------------
	Function NewIP(strIP)
		arrNewIP = Split(strIP,".")
		If arrNewIP(3) <> 255 Then
			NewIP = arrNewIP(0) & "." & arrNewIP(1) & "." & _
			arrNewIP(2) & "." & CInt(arrNewIP(3)) + 1
			Exit Function
			Else arrNewIP(3) = 1
		End If
		If arrNewIP(2) <> 255 Then
			NewIP = arrNewIP(0) & "." & arrNewIP(1) & "." & _
			arrNewIP(2) + 1 & "." & arrNewIP(3)
			Exit Function
			Else arrNewIP(2) = 1
		End If
		If arrNewIP(1) <> 255 AND arrNewIP(2) = 0 Then
			NewIP = arrNewIP(0) & "." & arrNewIP(1) + 1 & "." & _
			arrNewIP(2) & "." & arrNewIP(3)
			Exit Function
			Else arrNewIP(3) = 1
		End If
		If arrNewIP(0) <> 255 AND arrNewIP(1) = 0 Then
			NewIP = arrNewIP(0) + 1 & "." & arrNewIP(1)  & "." & _
			arrNewIP(2) & "." & arrNewIP(3)
			Else MsgBox "Error"
		End If
	End Function

	'#--------------------------------------------------------------------------
	'#  FUNCTION.......:	GetPingStatus(intStatusCode)
	'#  PURPOSE........:	Gets Ping response when unable to ping
	'#  ARGUMENTS......:	intStatusCode = status code as returned from WMI
	'#  EXAMPLE........:	GetPingStatus(11010)
	'#  NOTES..........:  
	'#--------------------------------------------------------------------------
	Function GetPingStatus(intStatusCode)
		Select Case intStatusCode
			Case 11001
				strStatus = "Buffer Too Small"
			Case 11002
				strStatus = "Destination Net Unreachable"
			Case 11003
				strStatus = "Destination Host Unreachable"
			Case 11004
				strStatus = "Destination Protocol Unreachable"
			Case 11005
				strStatus = "Destination Port Unreachable"
			Case 11006
				strStatus = "No Resources"
			Case 11007
				strStatus = "Bad Option"
			Case 11008
				strStatus = "Hardware Error"
			Case 11009
				strStatus = "Packet Too Big"
			Case 11010
				strStatus = "Request Timed Out"
			Case 11011
				strStatus = "Bad Request"
			Case 11012
				strStatus = "Bad Route"
			Case 11013
				strStatus = "TimeToLive Expired Transit"
			Case 11014
				strStatus = "TimeToLive Expired Reassembly"
			Case 11015
				strStatus = "Parameter Problem"
			Case 11016
				strStatus = "Source Quench"
			Case 11017
				strStatus = "Option Too Big"
			Case 11018
				strStatus = "Bad Destination"
			Case 11032
				strStatus = "Negotiating IPSEC"
			Case 11050
				strStatus = "General Failure"
		End Select
		GetPingStatus = strStatus
	End Function

	'#--------------------------------------------------------------------------
	'#  FUNCTION.......:	CheckWinArchitecture()
	'#  PURPOSE........:	Check Windows Architecture on PC
	'#  ARGUMENTS......:	
	'#  EXAMPLE........:	
	'#  NOTES..........:	Will return "32-bit" or "64-bit", etc
	'#--------------------------------------------------------------------------
	Function CheckWinArchitecture()
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
			strPC & "\root\cimv2") 
		Set colProcessor = objWMIService.ExecQuery _
				("Select * from Win32_Processor")
			For Each objItem In ColProcessor
				intArchitecture = objItem.Architecture
				Select Case intArchitecture
					Case 0
						strArchitecture = "32-bit"
					Case 1
						strArchitecture = "MIPS"
					Case 2
						strArchitecture = "Alpha"
					Case 3
						strArchitecture = "PowerPC"
					Case 6
						strArchitecture = "IPF"
					Case 9
						strArchitecture = "64-bit"
				End Select
			Next
		CheckWinArchitecture = strArchitecture
	End Function
	
	'#--------------------------------------------------------------------------
	'#  FUNCTION.......:	CheckWinVer()
	'#  PURPOSE........:	Check Windows Version on PC
	'#  ARGUMENTS......:	
	'#  EXAMPLE........:	
	'#  NOTES..........:	
	'#--------------------------------------------------------------------------
	Function CheckWinVer(strComp)
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strComp & "\root\cimv2") 
		Set colOS = objWMIService.ExecQuery _
			("Select Version from Win32_OperatingSystem")
		For Each objItem In colOS
			intVersion = objItem.Version
			arrVersion = Split(intVersion, ".")
			If arrVersion(1) = "" Then arrVersion(1) = 0
			intVersion = arrVersion(0) & "." & arrVersion(1)
			Select Case intVersion
				Case 5.0
					CheckWinVer = 1
				Case 5.1
					CheckWinVer = 2
				Case 5.2
					CheckWinVer = 2
				Case Else
					CheckWinVer = 3
			End Select
		Next
	End Function
	
	'#--------------------------------------------------------------------------
	'#  FUNCTION.......:	ConvertToDiskSize(intValue)
	'#  PURPOSE........:	Gets disk size string (eg. 1 MB)
	'#  ARGUMENTS......:	intValue = number of bytes to convert
	'#  EXAMPLE........:	ConvertToDiskSize(1024)
	'#  NOTES..........:  
	'#--------------------------------------------------------------------------
	Function ConvertToDiskSize(intValue)
		If (intValue / 1099511627776) > 1 Then
            ConvertToDiskSize = Round(intValue / 1099511627776,1) & " TB "
			ElseIf (intValue / 1073741824) > 1 Then
				ConvertToDiskSize = Round(intValue / 1073741824,1) & " GB "
				ElseIf (intValue / 1048576) > 1 Then
					ConvertToDiskSize = Round(intValue / 1048576,2) & " MB "
					ElseIf (intValue / 1024) > 1 Then
						ConvertToDiskSize = Round(intValue / 1024,2) & " KB "
						Else
							ConvertToDiskSize = Round(intValue) & " Bytes "
		End If
    End Function
	
	'#--------------------------------------------------------------------------
	'#  FUNCTION.......:	TimeSpan(Date1,Date2)
	'#  PURPOSE........:	Works out the difference between 2 dates
	'#  ARGUMENTS......:	Date1 = first date
	'#						Date2 = second date
	'#  EXAMPLE........:	TimeSpan(Date1,Now)
	'#  NOTES..........:	Dates can be specified in either order
	'#--------------------------------------------------------------------------
	Function TimeSpan(Date1,Date2)
		If (IsDate(Date1) And IsDate(Date2)) = False Then
			TimeSpan = "00:00:00"
			Exit Function
		End If
 
		intSeconds = Abs(DateDiff("S", Date1, Date2))
		intMinutes = intSeconds \ 60
		intHours = intMinutes \ 60
		intDays = intHours \ 24
		intHours = intHours MOD 24
		intMinutes = intMinutes MOD 60
		intSeconds = intSeconds MOD 60

		If intDays = 1 Then
			strDay = " day, "
			Else strDay = " days, "
		End If
		If intHours = 1 Then
			strHour = " hour, "
			Else strHour = " hours, "
		End If
		If intMinutes = 1 Then
			strMinute = " minute and "
			Else strMinute = " minutes and "
		End If
		If intSeconds = 1 Then
			strSecond = " second"
			Else strSecond = " seconds"
		End If
		
		If intDays = 0 Then
			If intHours = 0 Then
				TimeSpan = intMinutes & strMinute & intSeconds & strSecond
				Else
					TimeSpan = intHours & strHour & _
					intMinutes & strMinute & intSeconds & strSecond
			End If
			Else
				TimeSpan = intDays & strDay & intHours & strHour & _
				intMinutes & strMinute & intSeconds & strSecond
		End If
	End Function
	
	'#--------------------------------------------------------------------------
	'#  FUNCTION.......:	LoggedOnUser()
	'#  PURPOSE........:	Get the name of the logged on user as per WMI
	'#  ARGUMENTS......:	
	'#  EXAMPLE........:	
	'#  NOTES..........:	
	'#--------------------------------------------------------------------------
	Function LoggedOnUser(strComp)
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strComp & "\root\cimv2")
		Set colComputer = objWMIService.ExecQuery _
			("Select * from Win32_ComputerSystem")
		For Each objItem In colComputer
			strLoggedOn = objItem.UserName
		Next
		LoggedOnUser = strLoggedOn
	End Function
	
	'#--------------------------------------------------------------------------
	'#  FUNCTION.......:	GetPCName()
	'#  PURPOSE........:	Get the name of the PC as per WMI
	'#  ARGUMENTS......:	
	'#  EXAMPLE........:	
	'#  NOTES..........:	
	'#--------------------------------------------------------------------------
	Function GetPCName()
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\cimv2") 

		Set colComputer = objWMIService.ExecQuery _
			("Select * from Win32_ComputerSystem")
	
		For Each objItem In colComputer
			strPCName = objItem.Caption
		Next
		GetPCName = UCase(strPCName)
	End Function
	
	'#--------------------------------------------------------------------------
	'#  FUNCTION.......:	GetRoot()
	'#  PURPOSE........:	Retrieves the root share name 
	'#  ARGUMENTS......:	
	'#  EXAMPLE........:	
	'#  NOTES..........:	Retrieves root share in format: \\PC1\C$
	'#--------------------------------------------------------------------------
	Function GetRoot()	
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strPC & "\root\cimv2") 

		Set colShares = objWMIService.ExecQuery("Select * from Win32_Share Where Type='2147483648' AND Caption='Default share'")
		
		For Each objItem In colShares
			strRootShare = objItem.Name
		Next
		GetRoot = "\\" & strPC & "\" & strRootShare
	End Function
	
	'#--------------------------------------------------------------------------
	'#  FUNCTION.......:	FormatDate(strValue)
	'#  PURPOSE........:	Function to format dates in the format YYYYMMDD to 
	'#						DD/MM/YYYY
	'#  ARGUMENTS......:	strValue = date value in the format YYYYMMDD
	'#  EXAMPLE........:	FormatDate(20100914)
	'#  NOTES..........:	
	'#--------------------------------------------------------------------------
	Function FormatDate(strValue)
		If IsNull(strValue) OR strValue = "" Then
			strDate = ""
		Else
    		strDate = MID(strValue,7,2) & "/" & MID(strValue,5,2) & "/" & _
			LEFT(strValue,4) & " " & _ 
			MID(strValue,9,2) & ":" &  MID(strValue,11,2)
		End If
		FormatDate = strDate
	End Function
	
	'#--------------------------------------------------------------------------
    '#	FUNCTION.......:	EncodeCsv(strText)
    '#	PURPOSE........:	Encode provided text for CSV export
    '#	ARGUMENTS......:	strText = text to encode
    '#	EXAMPLE........:	EncodeCsv("Some text, etc.")
    '#	NOTES..........:	
    '#--------------------------------------------------------------------------
	Function EncodeCsv(strText)
		strText = Replace(strText, Chr(34), "")
		strText = Replace(strText, vbCrLf, " ")
		strText = Chr(34) & strText & Chr(34)
		EncodeCsv = strText
	End Function

	Sub Document_onContextMenu
		About()
	End Sub
	
</script>

<body style="background-color:#dddddd;">

	<span id="LoadingUtil"><h3>Loading, please wait...<h3>
	</span>
	<span id="UtilVersion" style="position:absolute;top:5px;right:10px;font-style:italic;font-weight:bold;font-size:0.9em;color:'#888888';">
	</span>
	
	<table width="100%" cellpadding="0" style="visibility:hidden;display:none;" id="UtilLayout">
		<tr>
			<td width="35%" valign="top">
				<span id=SearchArea>
				</span>
				<div style="text-align:center;">
					<input id=runbutton type="button" value="Select PC" name="btnSelectPC" title="Select highlighted PC" onClick="ShowPCInfo AvailablePCs.Value, False" onMouseDown="ChangeButtonColour(btnSelectPC)" onMouseUp="RevertButtonColour(btnSelectPC)">
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					
					<input type="text" name="PCSearch" size="15">
					<input id=runbutton type="button" value="Search" name="btnSearch" title="Search PC" onClick="ShowPCInfo PCSearch.Value, False" onMouseDown="ChangeButtonColour(btnSearch)" onMouseUp="RevertButtonColour(btnSearch)">
					<p><select name="SearchView" style="width:227" onChange="ChangeSearchView(SearchView.Value)">
					<option value="1">IP Range Search</option>
					<option value="2">Active Directory Search</option>
					</select><p>
					<input id=runbutton  class="button" type="button" value="Rescan" name="btnRescan" title="Rescan the selected PC" onClick="ShowPCInfo strPC, False" onMouseDown="ChangeButtonColour(btnRescan)" onMouseUp="RevertButtonColour(btnRescan)" disabled="true">
					<input id=runbutton  class="button" type="button" value="Reset Form" name="btnCleanUp" title="Reset the form to select a new PC" onClick="CleanUp" onMouseDown="ChangeButtonColour(btnCleanUp)" onMouseUp="RevertButtonColour(btnCleanUp)"><p>
					<div style="font-size: 0.9em;cursor:default;"><b>Created by Stuart Barrett | </b><a href="http://community.spiceworks.com/scripts/show/585" title="Download the latest version of the Computer Management Utility" target="_blank">Latest Version</a></div>
				</div>
			</td>
			<td width="65%" valign="top">
				<span id=DataArea>
				</span>
				<span id=RunAs style="position:absolute;bottom:35px;right:10px;font-style:italic;font-weight:bold;font-size:0.9em;color:'#888888';">
				</span>
			</td>
		</tr>
	</table>
	<p>
	<div id="menubar" style="padding-top:5px;" align="center">
		<table class="menutable">
			<tr>
				<td id="tab1" onclick="ShowPCInfo strPC, True" disabled=true>
					PC INFO
				</td>
				<td id="tab2" onclick="ShowSoftwareInfo True" disabled=true>
					SOFTWARE
				</td>
				<td id="tab3" onclick="ShowProcessInfo True" disabled=true>
					PROCESSES
				</td>
				<td id="tab4" onclick="ShowServiceInfo True" disabled=true>
					SERVICES
				</td>
				<td id="tab5" onclick="ShowStartupInfo True" disabled=true>
					STARTUP ITEMS
				</td>
				<td id="tab6" onclick="ShowPCActions()" disabled=true>
					ACTIONS
				</td>
				<td id="tab7" title="Change setup options for the PC Management Utility" onclick="Setup()">
					SETUP
				</td>
				<td id="tab8" title="Quit application" onclick="Window.Close">
					QUIT
				</td>
			</tr>
		</table>
	</div>
	
	<span style ="visibility:hidden; display:none;" id=SoftwareTab>
	</span>
	<span style ="visibility:hidden; display:none;" id=ProcessesTab>
	</span>
	<span style ="visibility:hidden; display:none;" id=ServicesTab>
	</span>
	<span style ="visibility:hidden; display:none;" id=StartupTab>
	</span>
	
</body>

</html>