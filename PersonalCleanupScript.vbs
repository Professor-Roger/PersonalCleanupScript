Option Explicit
On Error Resume Next

Dim objFSO, objShell, objWMIService, colProcesses, objProcess, objFile, objFolder
Dim strFolderPath, strFolderPathTMP, strScriptPath
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

CheckAndStopProcess "Xenon.exe"

CleanFolderContents ExpandEnvironment("%LocalAPPDATA%\Nostalgia")

CleanFolderContentsWithAdminRights ExpandEnvironment("%TMP%")

strScriptPath = WScript.ScriptFullName
If objFSO.FileExists(strScriptPath) Then
    objFSO.DeleteFile strScriptPath
End If

Sub CheckAndStopProcess(processName)
    Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & processName & "'")
    If colProcesses.Count > 0 Then
        For Each objProcess In colProcesses
            objProcess.Terminate()
            WScript.Sleep 1000
        Next
    End If
End Sub

Sub CleanFolderContents(folderPath)
    If objFSO.FolderExists(folderPath) Then
        Set objFolder = objFSO.GetFolder(folderPath)
        For Each objFile In objFolder.Files
            objFile.Delete
        Next
        For Each objSubfolder In objFolder.Subfolders
            objSubfolder.Delete
        Next
    End If
End Sub

Sub CleanFolderContentsWithAdminRights(folderPath)
    objShell.Run "cmd /c rmdir /s /q """ & folderPath & """", 0, True
End Sub
Function ExpandEnvironment(strPath)
    ExpandEnvironment = objShell.ExpandEnvironmentStrings(strPath)
End Function
