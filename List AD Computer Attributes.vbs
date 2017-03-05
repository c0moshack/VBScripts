'List Selected Computer Account Attributes
'Demonstration script that retrieves the location and description attributes for a computer account in Active Directory. 

On Error Resume Next

strComputer = InputBox("Enter PC name:", "Written by Mike Holmes")
Set objComputer = GetObject _
    ("LDAP://CN="&strComputer&",OU=NGWI DISC4-J6,OU=Computers,OU=NGWI,OU=States,DC=ng,DC=ds,DC=army,DC=mil")

objProperty = objComputer.Get("Location")
If IsNull(objProperty) Then
    Wscript.Echo "The location has not been set."
Else
    Wscript.Echo "Location: " & objProperty
    objProperty = Null
End If

objProperty = objComputer.Get("Description")
If IsNull(objProperty) Then
    Wscript.Echo "The description has not been set."
Else
    Wscript.Echo "Description: " & objProperty
    objProperty = Null
End If

