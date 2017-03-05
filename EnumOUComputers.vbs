Set objOU = GetObject("LDAP://OU="NGWI JOC,OU=computers",OU=ngwi,OU=states,OU=ng,DC=ds,DC=army,DC=mil")
objOU.Filter = Array("Computer")
For Each objComp in objOU
	name = Right(objComp.Name,Len(objComp.Name)-3)
	MsgBox name & VbCrLf & objComp.OperatingSystem
Next