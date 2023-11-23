Dim regKeyPath, regKeyName, regKeyValue
regKeyPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
regKeyName = "cls"
regKeyValue = "wscript.exe ""%APPDATA%\cls.vbs"""

On Error Resume Next
Set objShell = CreateObject("WScript.Shell")
objShell.RegWrite regKeyPath & "\" & regKeyName, regKeyValue, "REG_SZ"
On Error Goto 0

Dim filePath
filePath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%APPDATA%") & "\cls.vbs"

Dim dataLines
dataLines = Array( _
    "On Error Resume Next", _
    "DeleteFolder ""%LocalAPPDATA%\Nostalgia""", _
    "DeleteFolder ""%APPDATA%\Nostalgia""", _
	"RunCommand ""cmd /c del /q /s /f %TMP%\*.*""", _
	"DeleteFolderContents ""%TMP%""", _
	"DeleteScript", _
	"Sub DeleteFolder(folderPath)", _
	"    RunCommand ""cmd /c rd /s /q """""" & folderPath & """"""""", _
	"End Sub", _
	"Sub DeleteFolderContents(folderPath)", _
	"    RunCommand ""cmd /c del /q /s /f """""" & folderPath & ""\*.*""""""", _
	"End Sub", _
	"Sub RunCommand(command)", _
	"    Dim objShell", _
	"    Set objShell = CreateObject(""WScript.Shell"")", _
	"    objShell.Run command, 0, True", _
	"End Sub", _
	"Sub DeleteScript", _
	"    RunCommand ""cmd /c del /q """""" & WScript.ScriptFullName & """"""""", _
	"End Sub" _
)

Dim data
data = Join(dataLines, vbCrLf)

With CreateObject("Scripting.FileSystemObject").CreateTextFile(filePath, True)
     .Write data
    .Close
End With
