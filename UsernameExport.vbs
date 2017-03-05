On Error Resume Next

Const ADS_SCOPE_SUBTREE = 2

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
objCommandProperties("Sort On") = "SN"

objCommand.CommandText = _
    "SELECT SN FROM 'LDAP://ou=ngwi,dc=ng,dc=ds,dc=army,dc=mil' WHERE objectCategory='user'"  
Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst
Do Until objRecordSet.EOF
    Wscript.Echo objRecordSet.Fields("SN").Value
    objRecordSet.MoveNext
Loop