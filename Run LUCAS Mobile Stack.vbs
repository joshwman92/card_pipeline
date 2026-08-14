Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

appDir = fso.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = appDir

launchers = Array( _
    "Run Team LUCAS Mobile Server.vbs", _
    "Run Michael LUCAS Mobile Server.vbs", _
    "Run Team LUCAS Tunnel.vbs", _
    "Run Michael LUCAS Tunnel.vbs" _
)

For Each launcher In launchers
    path = appDir & "\" & launcher
    If fso.FileExists(path) Then
        shell.Run """" & path & """", 0, False
        WScript.Sleep 1500
    Else
        MsgBox "Could not find " & path, vbExclamation, "L.U.C.A.S Mobile Stack"
    End If
Next
