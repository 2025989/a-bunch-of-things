input = InputBox("Type the text you'd like to convert to a sound file.")

Const SAFT48kHz16BitStereo = 39
Const SSFMCreateForWrite = 3 'Creates file even if file exists and so destroys or overwrites the existing file
Set oWS = WScript.CreateObject("WScript.Shell")
userProfile = oWS.ExpandEnvironmentStrings("%userprofile%")

Dim oFileStream, oVoice

Set oFileStream = CreateObject("SAPI.SpFileStream")
oFileStream.Format.Type = SAFT48kHz16BitStereo
oFileStream.Open userprofile &"\Desktop\" &input &".mp3", SSFMCreateForWrite

Set oVoice = CreateObject("SAPI.SpVoice")
with oVoice
	Set .voice = .getvoices.item(0) 'there may be multiple voices installed on your system. try changing the int.
end with
oVoice.Speak input
Set oVoice.AudioOutputStream = oFileStream
oVoice.Speak input
Msgbox("Saved to Desktop!")
oFileStream.Close