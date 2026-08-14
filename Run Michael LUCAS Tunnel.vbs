Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

appDir = fso.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = appDir

scriptPath = appDir & "\scripts\windows\run_mobile_tunnel.ps1"
If fso.FileExists(scriptPath) Then
    shell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """ -Profile personal", 0, False
Else
    MsgBox "Could not find " & scriptPath, vbExclamation, "L.U.C.A.S Tunnel"
End If
