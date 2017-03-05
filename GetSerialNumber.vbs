Sub ReturnSerialNumber
	'''strComputer = INPUTBOX("Please enter the name of the computer:")
	strComputer = ComputerSerial.Value
	
	' Check that user entered a value
	IF strComputer = "" THEN
		wscript.quit
	END IF

	ON ERROR RESUME NEXT ' Handle errors connecting to the computer (Not switched on, permissions error etc)
	SET objWMI = GETOBJECT("winmgmts:{impersonationLevel=impersonate}!//" & strComputer & "")
	
	IF err <> 0 THEN ' Check for error
		'''wscript.echo "Error connecting to specified computer: " & err.description
		ComputerSerial.Value = "Error connecting to specified computer: " & err.description
		wscript.quit
	END IF
	ON ERROR GOTO 0 ' Turn off resume next error handling
	
	SET colOS = objWMI.ExecQuery("Select * from Win32_BIOS")
	
	FOR EACH objItem In colOS
		IF strSerial <> "" THEN
			strSerial = strSerial & ", " & objItem.SerialNumber
		ELSE
			strSerial = objItem.SerialNumber
		END IF
	NEXT
	
	ComputerSerial.Value = "Serial Number: " & strSerial 
	'''Wscript.Echo "Serial Number: " & strSerial 
End Sub

ReturnSerialNumber