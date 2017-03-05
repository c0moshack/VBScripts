Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile("C:\Users\paul.j.brown\Documents\J6\VBS\software_test.tsv", True)
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSoftware = objWMIService.ExecQuery _
 ("SELECT * FROM Win32_Product")

For Each objSoftware in colSoftware
	If objSoftware.Name = "DameWare" Then
		objTextFile.WriteLine objSoftware.Name & vbtab & objSoftware.Version
	Else
		objTextFile.Writeline "Done"
	End If
Next
objTextFile.Close