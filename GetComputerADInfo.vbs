On Error Resume Next

set fso = CreateObject("Scripting.FileSystemObject")
set ts = fso.CreateTextFile(".\computers.csv", true)

ts.writeline("Name,Serial Number,Description")
count = 0

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Sort On") = "Name"

objCommand.CommandText = _
    "<LDAP://OU=NGWI,OU=States,DC=ng,DC=ds,DC=army,DC=mil>;" & _
        "(&(objectCategory=computer));info,name,description;Subtree"
Set objRs = objCommand.Execute

objRs.MoveFirst
Do Until objRs.EOF
compName = objRs.Fields("name").Value
compInfo = objRs.Fields("info").Value
arrDesc = objRs.Fields("description").Value
compSn = split(compInfo,";")

ts.writeline(compName & "," & compSn(1) & "," & arrDesc(0))

objRs.MoveNext
Loop