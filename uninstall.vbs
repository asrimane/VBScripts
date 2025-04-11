'Declaration variable'
dim objIEDebugWindow,objDocument,DebugTrigger
Dim eraseFlag, unmountFlag
Dim execStatus, errorFlag
Dim wshShell, wshNetwork, fso, scriptObject, objShell
Dim appdataPath, desktopPath, installPath
Dim appdataFolderName, installFolderName
Dim shortcutAtStartPath,shortcutDesktopPath, umShortcutName, mShortcutName
Dim mScriptName, umScriptName, uninScriptName
Dim CurrentScriptName, CurrentScriptPath
Dim username
Dim selfDelete
dim exitFlag, taskPos, retryAllowed, nbOfRetryAllowed, retryCounter, retryTaskId, noError
'###[ First config ]#######################################################'
Set WshShell 	 	= CreateObject("WScript.Shell")
Set wshNetwork  	= CreateObject("WScript.Network")
Set fso			 	= CreateObject("Scripting.FileSystemObject")
set objShell		= CreateObject("Shell.Application")
Set scriptObject 	= fso.GetFile(Wscript.scriptfullname)
appdataPath			= WshShell.expandEnvironmentStrings("%APPDATA%") & "\"
shortcutAtStartPath	= WshShell.SpecialFolders("Startup") & "\"
desktopPath			= WshShell.SpecialFolders("Desktop") & "\"
username 			= wshNetwork.username
shortcutDesktopPath = desktopPath
CurrentScriptName 	= Wscript.ScriptName
CurrentScriptPath   = split(scriptObject, CurrentScriptName)
installPath			= "C:\"
appdataFolderName	= "MountDriveScript"
installFolderName	= "Scripts"
umShortcutName		= "UnMountDrive.LNK"
mShortcutName		= "MountDrive.LNK"
mScriptName			= "MountDrive.vbs"
umScriptName		= "UnmountDrive.bat"
uninScriptName 		= "uninstall.vbs"
eraseFlag			= true
unmountFlag			= false
DebugTrigger		= false
selfDelete 			= false
execStatus 			= 0
taskPos 			= 0
exitFlag 			= false
retryAllowed 		= true
nbOfRetryAllowed 	= 2
retryCounter 		= 0
retryTaskId 		= 0
'##############################################################################'

Debug "Starting uninstall"

'###[ Launch parameters ]#######################################################'
if WScript.Arguments.Count > 0 then
	set objArg = wscript.Arguments.Named
	if objArg.Exists("debug") then
		DebugUser()
	end if
	If objArg.Exists("noErase") Then
		eraseFlag = false
	End If
	If objArg.Exists("unmount") Then
		unmountFlag = true
	End If
end if
'##############################################################################'
If Not msgBox("Voulez-vous desinstaller le script ainsi que tout ses composants ?", _
		 vbYesNo + vbInformation + vbDefaultButton1, "MountDrive uninstall") = vbYes Then Wscript.Quit 1
	

If fso.FolderExists(installPath & installFolderName) OR eraseFlag Then
	Debug "Installation folder exists"
	If Not CurrentScriptPath(0) <> installPath & installFolderName & "\" Then selfDelete = true
	If eraseFlag Then
		on error resume next
		Do Until exitFlag
			Err.clear
			noError = false
			select case taskPos
				case 0
					Debug "Task 1"
					If Not fso.FolderExists(appdataPath & appdataFolderName) Then fso.CreateFolder appdataPath & appdataFolderName
					fso.CopyFile CurrentScriptPath(0) & CurrentScriptName, appdataPath & appdataFolderName & "\" & CurrentScriptName
				case 1
					Debug "Task 2"
					If fso.FileExists(shortcutDesktopPath & mShortcutName) 	Then 
						fso.DeleteFile(shortcutDesktopPath & mShortcutName)
					else
						debug "noShortcut"
					End If
				case 2
					Debug "Task 3"
					If fso.FileExists(shortcutDesktopPath & umShortcutName) Then 
						fso.DeleteFile(shortcutDesktopPath & umShortcutName)
					else
						debug "noShortcut"
					End If
				case 3
					Debug "Task 4"
					If fso.FileExists(shortcutAtStartPath & mShortcutName) 	Then 
						fso.DeleteFile(shortcutAtStartPath & mShortcutName)
					else
						debug "noShortcut"
					End If
				case 4
					Debug "Task 5"
					'If selfDelete Then scriptObject.Delete
					Debug "erase current script"
					If fso.FolderExists(installPath & installFolderName) Then fso.DeleteFolder(installPath & installFolderName)
				case 5
					Debug "Task 6"
					If fso.FolderExists(appdataPath & appdataFolderName) 	Then 
						fso.DeleteFolder(appdataPath & appdataFolderName)
					else
						debug "noAppdatafolder"
					End If
				case else
					Debug "All task done !"
					exitFlag = true
			end select		

			select case Err.Number
				case 0
					taskPos = taskPos + 1
					retryCounter = 0
					retryTaskId = 0
					noError = true
					Debug "No error... continue"
				case 70
					Debug "Error permission denied"
					Debug "Trying to run current script as administrator"
					Debug "Path = " & scriptObject
					Debug objShell.ShellExecute( "wscript.exe", _
									appdataPath & appdataFolderName & "\" & CurrentScriptName & " /debug", "", "runas", 1)
					set objShell = nothing
					wscript.Quit 1
				case 2
				case else
					Debug "Error case else"
					debug "Erreure non geree [" & Err.Number & "]" & Err.Description
			End Select
			If Not noError Then
				Debug "Error -- 1"
				If retryAllowed Then
					Debug "ok for retry"
					If retryCounter >= nbOfRetryAllowed  AND Not retryTaskId <> taskPos  Then 
						Debug "Retry over reset value"
						taskPos = taskPos + 1
						retryCounter = 0
						retryTaskId = 0
					else
						retryTaskId = taskPos	
						retryCounter = retryCounter + 1
					end if
					Debug "retryTaskID : " & retryTaskId & " " & "retryCounter =  " & retryCounter 
				else
					taskPos = taskPos + 1
				End If		
			End If
		Loop







	End If

	If unmountFlag Then execute("net use * /delete /yes")
else
	Debug "Installation folder doesn't exists... abort"
	execStatus = 1
End If
Debug "uninstall finish ExecReturn = " & execStatus


msgBox "Desinstallation terminee", vbInformation + vbDefaultButton1, "MountDrive uninstall"
set wshNetwork = nothing 
set wshNetwork = nothing 
set fso 	   = nothing 


'###[ IE_onQuit ]#########################################################################'
Sub IE_onQuit()
	DebugTrigger = false
	set objIEDebugWindow = nothing
	set objDocument = nothing
End Sub

'###[ Debug ]#########################################################################'
Sub Debug (Text)
	if Not DebugTrigger then
		Exit sub
  	end if
  	If Not IsObject(objIEDebugWindow) Then
		Set objIEDebugWindow = WScript.CreateObject("InternetExplorer.Application", "IE_") 
		with objIEDebugWindow
			.Navigate 	"about:blank"
			.Visible 	= true
			.ToolBar 	= false
			.StatusBar 	= false
			.Width 		= 600
			.Height 	= 500
			.Left 		= 10
			.Top 		= 10
			.Document.title = "Mount Drive Debug windows"
		end with

		Do While objIEDebugWindow.Busy
		     WScript.Sleep 100
		Loop
		Set objDocument = objIEDebugWindow.Document
		objDocument.Open
		objDocument.writeln("<b>" & Now & "</b></br>"	)
	else
		If InStr(TypeName(objIEDebugWindow),"IWebBrowser2") Then
	  	  	objDocument.writeln(Text & "<br>" & vbCrLf)
	  	else
	  		IE_onQuit()
	  		Exit Sub
		End If
  	End If	
End Sub

'###[ DebugUser ]#########################################################################'
Sub DebugUser ()
	DebugTrigger = true
	debug "  "
	debug "  "
End Sub

'###[ WRLogMAJ ]#########################################################################'
Function WRLogMAJ(opt, strLogTxt)
	dim msg
	select case opt
		case "open"

			msg = 	"# [Open]:{Log for " & CurrentScriptName & "} - [Computer] = {" & ComputerName & "} - " & _
					"[UserName] = {" & winUser & "} - [DATE-TIME] = {" & DT_TM & "}"
			msg = strFill(msg, "#") & vbCrLf & msg
			wrlogSub(msg)
			wrlogSub("# " & strFill(strLogTxt, "-"))
			wrlogSub("# " & strLogTxt)

		case "log"
			wrlogSub("# " & strFill(strLogTxt, "-"))
			wrlogSub("# " & strLogTxt )
			wrlogSub("# " & strFill(strLogTxt, "-"))

		case "close"
			wrlogSub("# " & strLogTxt)
			wrlogSub("# " & strFill(strLogTxt, "-"))
			msg = 	"# [Close]:{Log for " & CurrentScriptName & "} - [Computer] = {" & ComputerName & "} - " & _
					"[UserName] = {" & winUser & "} - [DATE-TIME] = {" & DT_TM & "}"
			msg = msg & vbCrLf & strFill(msg, "#")
			wrlogSub(msg)
	End select


end Function

'###[ wrlogSub ]#########################################################################'
Sub wrlogSub(strtext)
	select case WriteNewLineToFile(LocalDrive & "\" & "Log-Mise-A-Jour.txt",strtext)
		case 0
		case 1
	end select
End Sub

'###[ strFill ]#########################################################################'
Function strFill(strMsg, strChar)
	dim execRet, tmpVar, vType, intSize
	execRet = -1
	intSize = 0
	vType = typeName(strMsg)
	select case vType
		case "Integer" 
			intSize = strMsg
		case "Long"
			intSize = strMsg
	 	case "String"
	 		intSize = Len(strMsg)
	 	case else
	 		execRet = -4
	End select
	If intSize > 0 Then
		If Not strChar = vbEmpty Then
			For i = 0 To intSize step 1 
				tmpVar = tmpVar & strChar
				execRet = tmpVar
			Next
		else
			execRet = -3
		End If
	else
		execRet = -2
	End If
strFill = execRet
End Function

'###[ execute ]#########################################################################'
Function execute(strCommand)
	Dim oShell,execReturn
	execReturn = -1
	if strCommand = vbEmpty then exit function
	Set oShell = WScript.CreateObject ("WScript.Shell")
	Debug "EXEC " & strCommand
	execReturn = oShell.run(strCommand, 0, true)
	Debug "EXEC-RETURN " & execReturn
	Debug "EXEC END"
	set oShell = Nothing
	execute = execReturn
End function

'###[  ]#########################################################################'

'###[  ]#########################################################################'

'###[  ]#########################################################################'

'###[  ]#########################################################################'

'###[  ]#########################################################################'
