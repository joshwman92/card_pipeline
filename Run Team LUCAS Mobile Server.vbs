Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

appDir = fso.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = appDir

Set env = shell.Environment("PROCESS")
env("LUCAS_SETTINGS_PATH") = appDir & "\lucas_settings.json"
env("LUCAS_ASSIGNMENT_CONFIG_PATH") = appDir & "\assignment_companies.json"
env("LUCAS_MOBILE_PORT") = "8765"

pythonwPath = appDir & "\.venv\Scripts\pythonw.exe"
pythonPath = appDir & "\.venv\Scripts\python.exe"

If fso.FileExists(pythonwPath) Then
    shell.Run """" & pythonwPath & """ """ & appDir & "\app.py"" --mobile-server", 0, False
ElseIf fso.FileExists(pythonPath) Then
    shell.Run """" & pythonPath & """ """ & appDir & "\app.py"" --mobile-server", 0, False
Else
    MsgBox "L.U.C.A.S could not find the local Python environment." & vbCrLf & vbCrLf & _
           "Run install_dependencies.bat, then open this launcher again.", _
           vbExclamation, "L.U.C.A.S Mobile Server"
End If
