

For i = asc("K") To asc("M") step 2
	msgbox removeNetDrive(i)
Next

'msgbox removeNetDrive("M")




Function removeNetDrive(rmDriveLetter)
	dim FSO, wNet, execReturn, vType, letter, errorFlag
	Set wNet  	= CreateObject("WScript.Network")
	set FSO 	= CreateObject("Scripting.FileSystemObject")
	vType	    = typeName(rmDriveLetter)
	execReturn  = -1 'default exec value'
	select case vType
		case "Integer" 
			letter = chr(rmDriveLetter) & ":"
		case "Long"
			letter = chr(rmDriveLetter) & ":"
	 	case "String"
	 		if Instr(rmDriveLetter, ":") Then
	 			letter = rmDriveLetter
	 		else
	 			letter = rmDriveLetter & ":"
	 		end if
	 	case else
	 		execReturn = 2 'Type non valide'
	 		errorFlag = true
	End select

	If Not errorFlag Then
		If FSO.DriveExists(letter) Then
			wNet.RemoveNetworkDrive letter, True, True
			execReturn = 0 'sucsess'
		else
			execReturn = 1 'la lettre " & letter & " N'est pas utilisee'
		End If
	End if
	set FSO  = nothing
	set wNet = nothing
	removeNetDrive = execReturn
End Function