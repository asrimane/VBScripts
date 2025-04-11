const version = "1.0.0"
'##Variable globale##
dim strScriptDestinationPath,strFinalScriptName,CurrentScriptName,CurrentScriptPath, WshShell, deleteScriptFlag
dim Tarray,strUpdateName,winUser,DT_TM,LocalDrive,strSharedUpdateFolder,strHostUpdate,ComputerName, noDownloadFlag
Dim objIEDebugWindow, objDocument, usingGetNextDrive, debugTrigger, AutoAllocStartDrive,AutoAllocEndDrive,objArg
Dim createInstallationFolder

Set WshShell = CreateObject("WScript.Shell")
set fso=CreateObject("Scripting.FileSystemObject")
set f=fso.GetFile(Wscript.scriptfullname)
Set wshNetwork  = CreateObject("WScript.Network")


'################################################################################################################################################
'#[DEBUG TRIGGER]##
debugTrigger 			 = false
deleteScriptFlag		 = false
usingGetNextDrive 		 = false 
noDownloadFlag			 = false
createInstallationFolder = false
AutoAllocStartDrive 	 = 75
AutoAllocEndDrive 		 = 90
strSharedUpdateFolder 	 = "\c_Public\VBscript\UPDATE-MOUNT-SCRIPT"
strScriptDestinationPath = "C:\Scripts"
strFinalScriptName 		 = "MountDrive.vbs"
strUninstallScriptName	 = "uninstall.vbs"
strUpdateName 			 = "UPDATE.420"
strHostUpdate 			 = "\\lowrynas"
winUser 				 = wshNetwork.username
ComputerName 			 = wshNetwork.ComputerName
LocalDrive 				 = AutoAlloc("raw")
CurrentScriptName 		 = WScript.ScriptName
Tarray 					 = split(f,CurrentScriptName)
CurrentScriptPath 		 = Tarray(0)
DT_TM 					 = Day(Date)&":"& Month(date) &":"& Year(Date) & "-@-" & Time


if WScript.Arguments.Count > 0 then
	set objArg = wscript.Arguments.Named
   	If objArg.Exists("d") Then
   			debugTrigger = true
   	End If
   	If objArg.Exists("nd") Then
   			noDownloadFlag = true
   	End If
	if objArg.Exists("user") then
		winUser = objArg.Item("user")
	end if	
end if

on error resume next
dim attempCounter
attempCounter = 0
Do
	Err.clear
	wshNetwork.MapNetworkDrive LocalDrive, strHostUpdate & strSharedUpdateFolder, false,strUser,strPassword	
	If Err.Number = -2147024810 Then
		subCredRequest()
	End If
	attempCounter = attempCounter +1
	If attempCounter > 5 Then
		msgbox "Annulation de la mise a jour"
		wscript.Quit 1		
	End If

Loop While Err.Number <> 0

Call WRLogMAJ("open", "Installation Begin")

Debug	"##########################################################" & vbCrLf 
Debug 	"# Error Number : " & Err.Number
Debug 	"# LocalDrive : " & LocalDrive
Debug 	"# strHostUpdate : " & strHostUpdate
Debug 	"# strSharedUpdateFolder : " & strSharedUpdateFolder
debug 	"# noDownloadFlag = " & noDownloadFlag
Debug	"##########################################################" & vbCrLf 
Debug	"# Mise a jour du script : " & strFinalScriptName & vbCrLf 
Debug	"# Destination de la mise a jour : " & strScriptDestinationPath & vbCrLf
Debug	"# Nom du fichier de mise a jour : " & strUpdateName & vbCrLf
Debug	"# Nom du script executer pour la mise à jour : " & CurrentScriptName & vbCrLf 
Debug	"# Chemin d'acces du dossier du quel le script a ete lance " & CurrentScriptPath & vbCrLf 
Debug 	"##############################################################" & vbCrLf 
Debug 	"# Tache executee :" & vbCrLf 
	
If Not noDownloadFlag Then
	debug "Telechargement de la mise a jour"
	If fso.FileExists(LocalDrive & "\" & strUpdateName) Then
		debug "Telechargement du fichier de mise a jour " & strUpdateName
		call WRLogMAJ("log", "Telechargement du fichier de mise a jour " & strUpdateName)
		fso.CopyFile LocalDrive & "\" & strUpdateName, CurrentScriptPath
		If FileExists(LocalDrive & "\" & strUninstallScriptName) Then
			fso.CopyFile LocalDrive & "\" & strUninstallScriptName, CurrentScriptPath
		else
			call WRLogMAJ("Log", "uninstall file download failed")
			debug "uninstall Telechargement echouer"
		End If
		If fso.FileExists(CurrentScriptPath & strUpdateName) AND fso.FileExists(CurrentScriptPath & strUninstallScriptName) Then
			call WRLogMAJ("log", "Telechargement reussi")
			debug "Telechargement reussi"
		else
			call WRLogMAJ("Log", "Telechargement echouer")
			debug "Telechargement echouer"
		End If
	End If
End If

if Not fso.FileExists(strScriptDestinationPath & "\" & strFinalScriptName) then
	if fso.FileExists(CurrentScriptPath & strUpdateName) then
		Debug "# Deplacement de la mise a jour : " & strUpdateName & " Dans le dossier : " & strScriptDestinationPath & vbCrLf 
		if Not fso.FolderExists(strScriptDestinationPath) then createInstallationFolder = true
		If createInstallationFolder Then fso.CreateFolder strScriptDestinationPath
		fso.MoveFile CurrentScriptPath & strUpdateName, strScriptDestinationPath & "\" & strFinalScriptName
		fso.MoveFile CurrentScriptPath & strUninstallScriptName, strScriptDestinationPath & "\" & strUninstallScriptName

		Call WRLogMAJ("log", "Install Update")
	else
		msgbox 	"Le script de mise a jour n'a trouver aucune mise a jour dans : " & CurrentScriptPath & strUpdateName & VbCrLf & _
				" La mise a jour va etre annulee.",vbCritical
		Call WRLogMAJ("log", "Error No Update")
		'Wscript.Quit 1
	end if
else
	if fso.FileExists(CurrentScriptPath & strUpdateName) then
		Debug "# Supression du script : " & strFinalScriptName & " Dans le dossier : " & strScriptDestinationPath & vbCrLf 
		fso.DeleteFile(strScriptDestinationPath & "\" & strFinalScriptName)
		Debug "# Deplacement de la mise à jour : " & strUpdateName & "  Dans le dossier : " & strScriptDestinationPath & vbCrLf
		fso.MoveFile CurrentScriptPath & strUpdateName, strScriptDestinationPath & "\" & strFinalScriptName
		Call WRLogMAJ("log", "Install Update")
	else
		msgbox 	"Le script de mise a jour n'a trouver aucune mise a jour dans : " & CurrentScriptPath & strUpdateName & VbCrLf & _
				" La mise a jour va etre annulee.",vbCritical
		Call WRLogMAJ("log", "Error No Update")
		'Wscript.Quit 1
	end if
end if


Call WRLogMAJ("close", "Installation Finish")
wshNetwork.RemoveNetworkDrive  LocalDrive, True, True
WshShell.Popup "Mise a  jour du script : MountDrive " & vbcrlf & "effectuee", 3, "Update script for MountDrive ", vbInformation
if InStr(CurrentScriptPath,"AppData") then
	deleteScriptFlag = true
end if
If deleteScriptFlag Then
	Debug "# Supression du script de mise à jour" & vbCrLf
	f.Delete
End If
If debugTrigger Then
	WshShell.Run "wscript " & strScriptDestinationPath & "\" & strFinalScriptName & " /debug" & " /user:" & winUser
else
	WshShell.Run "wscript " & strScriptDestinationPath & "\" & strFinalScriptName & " /user:" & winUser
End If
'##############################################################################################
'WriteNewLineToFile
'##############################################################################################

Function WriteNewLineToFile(strPath,strTxt)    
	dim oFSO,oTxtFile,intExecStatus                       
	Set oFSO = WScript.CreateObject("Scripting.FileSystemObject")
	intExecStatus = 0
	if oFSO.FileExists(strPath) then
		set oTxtFile = oFSO.OpenTextFile(strPath,8) 
		oTxtFile.WriteLine(strTxt)
		oTxtFile.Close
		intExecStatus = 0
	else 
		intExecStatus = 1
	end if	

	'*** Destruction des objets
	Set oFSO = Nothing
	Set oTxtFile = nothing
	WriteNewLineToFile = intExecStatus
End Function
'##############################################################################################
'##############################################################################################

'##############################################################################################
'WriteNewLineToFile
'##############################################################################################
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
'##############################################################################################
'##############################################################################################
Sub wrlogSub(strtext)
	select case WriteNewLineToFile(LocalDrive & "\" & "Log-Mise-A-Jour.txt",strtext)
		case 0
		case 1
	end select
End Sub

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
'##############################################################################################
'AutoAlloc
'##############################################################################################
Function AutoAlloc(parameter)
	dim returnValue
	Select Case parameter

		case "oneTime"
			Debug "oneTime"
			tempArray = split(Asc(GetNextDrive(AutoAllocStartDrive, AutoAllocEndDrive)),":")
			returnValue = tempArray(0) 
		case "auto"
			Debug "Auto"
			If Not usingGetNextDrive Then
				usingGetNextDrive = true
			End If
			tempArray = split(Asc(GetNextDrive(AutoAllocStartDrive, AutoAllocEndDrive)),":")
			returnValue = tempArray(0)
		case "raw"
			returnValue = GetNextDrive(AutoAllocStartDrive, AutoAllocEndDrive)
		case ""
			Debug "No argument"
			returnValue = AutoAllocEndDrive
		case else
			returnValue = AutoAllocEndDrive

	End Select
	AutoAlloc = returnValue
End Function
'##############################################################################################
'GetNextDrive
'##############################################################################################
Function GetNextDrive(START_DRIVE, END_DRIVE)                              
	Dim FSO, letter, drv	
	Set FSO = CreateObject("Scripting.FileSystemObject")	

	For letter = START_DRIVE To END_DRIVE	
		drv = Chr(letter) & ":"	
	    If Not FSO.DriveExists(drv) Then	
	    	GetNextDrive = drv
			Exit Function
	    End If
	Next
	Set FSO = nothing
End Function


'##############################################################################################
'[DEBUG]
'##############################################################################################
Sub IE_onQuit()
	debugTrigger = false
	set objIEDebugWindow = nothing
	set objDocument = nothing
End Sub
Sub Debug (Text)
	if Not debugTrigger then
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
			.Document.title = "Mount Drive DEBUG windows"
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

'##############################################################################################
Function requestCredential(strUsername, strPassword)
	dim objIE, strHTML, wshShell, bCancelButton, bOkButton, execStatus, windowsTitle
	dim bFavoritesBar, bLinksExplorer, strRegValFavBar, strRegValLinksExp	

	bFavoritesBar     = False
	strRegValFavBar   = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\MINIE\LinksBandEnabled"
	bLinksExplorer 	  = False
	strRegValLinksExp = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\LinksExplorer\Docked"
	windowsTitle = WScript.ScriptName '& ": Nom d'utilisateur et mot de passe requis"
	execStatus = -1	

	Set wshShell = CreateObject("WScript.Shell")	

	On Error Resume Next	

	If wshShell.RegRead( strRegValFavBar ) = 1 Then
		bFavoritesBar = True 'flag is disable'
		wshShell.RegWrite strRegValFavBar, 0, "REG_DWORD"
	End If
	' Temporarily hide IE's Links Explorer if it is visible
	If wshShell.RegRead( strRegValLinksExp ) = 1 Then
		bLinksExplorer = True 'flag is disable'
		wshShell.RegWrite strRegValLinksExp, 0, "REG_DWORD"
	End If
	On Error Goto 0	

	Set objIE  = CreateObject("InternetExplorer.Application")	

	objIE.Navigate "about:blank"
	' Add string of "invisible" characters (500 tabs) to clear the title bar
	objIE.Document.title = windowsTitle & String( 500, 7 )
	objIE.AddressBar     = False
	objIE.Resizable      = false
	objIE.StatusBar      = False
	objIE.ToolBar        = False
	objIE.Width          = 340
	objIE.Height         = 180
	' Center the dialog window on the screen
	With objIE.Document.parentWindow.screen
		objIE.Left = (.availWidth  - objIE.Width ) \ 2
		objIE.Top  = (.availheight - objIE.Height) \ 2
	End With
	' Wait till IE is ready
	Do While objIE.Busy
		WScript.Sleep 200
	Loop
	' Insert the HTML code to prompt for a password
	strHTML = "<div style=""text-align: center;"">" _
	        & "<p>" & myPrompt & "</p>" _
	        & "<p>" _ 
	        & "<label for=""Username"">Username</label> " _
	        &"<input type=""textfield"" size=""20"" id=""Username"" onkeyup=" _
	        & """if(event.keyCode==13){document.all.OKButton.click();}"" /></p>" _
	        & "<p>" _ 
	        & "<label for=""Password"">Password</label> " _
	        &"<input type=""password"" size=""20"" id=""Password"" onkeyup=" _
	        & """if(event.keyCode==13){document.all.OKButton.click();}"" /></p>" _
	        & "<p>" _ 
	        &"<input type=""hidden"" id=""OK"" name=""OK"" value=""0"" /> " _
	        & "<input type=""hidden"" id=""CANCEL"" name=""CANCEL"" value=""0"" /> " _
	        & "<input type=""submit"" value="" OK "" id=""OKButton"" " _
	        & "onclick=""document.all.OK.value=1"" /> " _
	        & "<input type=""submit"" value="" CANCEL "" id=""CancelButton"" " _
	        & "onclick=""document.all.CANCEL.value=1"" /> </p>" _
	        & "</div>"
	objIE.Document.body.innerHTML = strHTML
	' Hide the scrollbars
	objIE.Document.body.style.overflow = "auto"
	' Make the window visible
	objIE.Visible = True
	' Set focus on password input field
	objIE.Document.all.Username.focus	

	On Error Resume Next
	Do While (execStatus = -1)
		WScript.Sleep 200
		If Err.Number = 424 Then	' User clicked red X (or Alt+F4) to close IE window
			msgbox Err.Number
			strUsername = ""
			strPassword = ""
			execStatus = 1
			objIE.Quit
			Set objIE = Nothing
			exit do
		elseIf objIE.Document.all.CANCEL.value = 1 Then ' User clicked cancel button
			strUsername = ""
			strPassword = ""
			execStatus = 2
			objIE.Quit
			Set objIE = Nothing
		elseIf objIE.Document.all.OK.value = 1 Then
			execStatus = 0
		End If
		wscript.Sleep(1)
	Loop
	On Error Goto 0
	If Not execStatus <> 0 Then  'changement fait entre if not >0 pour if not <> 0'
		strUsername = Trim(objIE.Document.all.Username.value)
		strPassword = Trim(objIE.Document.all.Password.value)
		objIE.Quit
		Set objIE = Nothing
		execStatus = 0
	End If	
	

	On Error Resume Next
	' Restore IE's Favorites Bar if applicable
	If bFavoritesBar Then wshShell.RegWrite strRegValFavBar, 1, "REG_DWORD"
	' Restore IE's Links Explorer if applicable
	If bLinksExplorer Then wshShell.RegWrite strRegValLinksExp, 1, "REG_DWORD"
	On Error Goto 0	

	Set wshShell = Nothing	

	requestCredential = execStatus
End Function

Sub subCredRequest ()
	Select Case requestCredential(strUser, strPassword)
		case 1
			msgbox "Demande de mot de passe annulee, le script s'arrete" & vbcrlf & "Arret du script",vbCritical
			WScript.Quit 1
		case 2
			msgbox "Demande de mot de passe annulee, le script s'arrete" & vbcrlf & "Arret du script",vbCritical
			WScript.Quit 1
	end Select 
End Sub
'##############################################################################################


