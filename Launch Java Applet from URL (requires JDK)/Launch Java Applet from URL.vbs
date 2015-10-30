Const ForReading = 1
Const ForWriting = 2
Set fso = CreateObject("Scripting.FileSystemObject")

' User input
Set oWS = WScript.CreateObject("WScript.Shell")
userProfile = oWS.ExpandEnvironmentStrings("%userprofile%")
name = InputBox("Raw code URL: ", "URL", "")
' Exit if there's no input
If (name <> "") Then
	' Set variables
	path = fso.GetParentFolderName(name)
	name = fso.GetFileName(name)
	short = Left (name, Len(name)-5)
	
	' Create folder if folder doesn't exist
	Set fso = CreateObject("Scripting.FileSystemObject")
	If Not (fso.FolderExists(userprofile &"\AppData\Roaming\JavaApplets")) Then
		fso.CreateFolder(userprofile &"\AppData\Roaming\JavaApplets")
	End If
	
	' Download from server & Copy file to folder
	Set fso = CreateObject("Scripting.FileSystemObject")
	strFile = (userprofile &"\AppData\Roaming\JavaApplets\" &short &".java")
	Set file = fso.OpenTextFile(strFile, ForWriting, True)
	Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
	http.Open "GET", (path &"/" &short &".java"), False
	http.Send
	For i = 1 To LenB(http.ResponseBody)
		file.Write Chr(AscB(MidB(http.ResponseBody, i, 1)))
	Next
	file.Close()
	
	' Delete line with first occurrence of "package " in .java file
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set file = fso.OpenTextFile(userprofile &"\AppData\Roaming\JavaApplets\" &short &".java", ForReading)
	For count = 0 to 1
		Do Until file.AtEndOfStream
			strLine = file.ReadLine
			If InStr(strLine, "package ") = 0 Then
				strNewContents = strNewContents &strLine &vbCrLf
			End If
		Loop
	Next
	file.Close
	Set file = fso.OpenTextFile(userprofile &"\AppData\Roaming\JavaApplets\" &short &".java", ForWriting)
	file.Write strNewContents
	file.Close
	
	' Create .html file
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set file = fso.CreateTextFile(userprofile &"\AppData\Roaming\JavaApplets\" &short &".html", True)
	file.WriteLine("<applet code=""" &short &""" width=400 height=400></applet>")
	file.Close
	
	' Run applet
	CreateObject("Wscript.Shell").Run "cmd.exe /c start /min javac " &userprofile &"\AppData\Roaming\JavaApplets\" &short &".java", 0, True
	WScript.Sleep 5000
	CreateObject("Wscript.Shell").Run "cmd.exe /c start /min appletviewer " &userprofile &"\AppData\Roaming\JavaApplets\" &short &".html", 0, True
	
	' Delete files
	WScript.Sleep 2000
	fso.DeleteFile(userprofile &"\AppData\Roaming\JavaApplets\*.*")
End If