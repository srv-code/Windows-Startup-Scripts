Option Explicit
Dim objFSO, objFolder, objShell, objTextFile, objFile
Dim strDirectory, strFile, strText, strUsername, str, logEntriesCount, strDate, strTime
strDate = Date
strTime = Time
strDirectory = "C:\Users\sourav\AppData\Local\AutoLogger\Logs"
strFile = "\login.log"

' Create the File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get username and create log string to write
strUsername = CreateObject("WScript.Network").UserName
strText = "User '" & strUsername & "' logged in on " & strDate & " at " & strTime

' Check that the strDirectory folder exists
If objFSO.FolderExists(strDirectory) Then
   Set objFolder = objFSO.GetFolder(strDirectory)
Else
   Set objFolder = objFSO.CreateFolder(strDirectory)
   WScript.Echo "Just created " & strDirectory
End If

If objFSO.FileExists(strDirectory & strFile) Then
   Set objFolder = objFSO.GetFolder(strDirectory)
Else
   Set objFile = objFSO.CreateTextFile(strDirectory & strFile)
   Wscript.Echo "Just created " & strDirectory & strFile
End If 

set objFile = nothing
set objFolder = nothing
' OpenTextFile Method needs a Const value
' ForAppending = 8 ForReading = 1, ForWriting = 2
Const ForAppending = 8

Set objTextFile = objFSO.OpenTextFile _
(strDirectory & strFile, ForAppending, True)

' Writes strText every time you run this VBScript
objTextFile.WriteLine(strText)
objTextFile.Close

'Showing confirmation message to user about successful log write
CreateObject("SAPI.SpVoice").Speak("Welcome " & strUsername & "!")
CreateObject("SAPI.SpVoice").Speak("Login successfully logged on " & strDate & " at " & strTime & "!")
str = msgbox("Welcome " & strUsername & "!" & (Chr(10)) & "Login successfully logged!" & (Chr(10)) & "Logged on " & strDate & " at " & strTime & "!", vbInformation, "AutoLogger: Welcome wisher")

set objTextFile = nothing
set objFSO = nothing
' Bonus or cosmetic section to open file with Notepad for user viewing
If err.number = vbEmpty then
	' Counting the total number of log entries
	set objFSO = CreateObject("Scripting.FileSystemObject") 
	set objTextFile = objFSO.OpenTextFile(strDirectory & strFile, ForAppending, True) 
	logEntriesCount = objTextFile.Line - 1
	Set objFSO = nothing
	Set objTextFile = nothing
	
	If (msgbox("Want to view all " & logEntriesCount & " log entries?", vbYesNo, "AutoLogger: View log?") = vbYes) Then
		Set objShell = CreateObject("WScript.Shell")
		objShell.run("Notepad" & " " & strDirectory & strFile)
	End If
Else WScript.echo "AutoLogger: VBScript Error: " & err.number
End If

WScript.Quit