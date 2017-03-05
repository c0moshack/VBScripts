strComputer = INPUTBOX("Please enter the name of the computer:")

' Check that user entered a value
IF strComputer = "" THEN
	wscript.quit
END IF

ON ERROR RESUME NEXT ' Handle errors connecting to the computer (Not switched on, permissions error etc)
SET objWMI = GETOBJECT("winmgmts:{impersonationLevel=impersonate}!//" & strComputer & "")

IF err <> 0 THEN ' Check for error
	wscript.echo "Error connecting to specified computer: " & err.description
	wscript.quit
END IF
ON ERROR GOTO 0 ' Turn off resume next error handling

SET colOS = objWMI.ExecQuery("Select * from Win32_ComputerSystem")

FOR EACH objItem In colOS
	IF strUsers <> "" THEN
		strUsers = strUsers & ", " & objItem.UserName
	ELSE
		strUsers = objItem.UserName
	END IF
NEXT

wscript.echo "The following user(s) are logged on to " & strComputer & ":" & strUsers