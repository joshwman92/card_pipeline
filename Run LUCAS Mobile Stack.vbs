Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

appDir = fso.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = appDir

Sub LaunchMobileServer(settingsFile, assignmentFile, port, publicUrlName, publicUrl)
    Set env = shell.Environment("PROCESS")
    env("LUCAS_SETTINGS_PATH") = appDir & "\" & settingsFile
    env("LUCAS_ASSIGNMENT_CONFIG_PATH") = appDir & "\" & assignmentFile
    env("LUCAS_MOBILE_PORT") = port
    env(publicUrlName) = publicUrl

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
End Sub

Sub LaunchTunnel(profile)
    scriptPath = appDir & "\scripts\windows\run_mobile_tunnel.ps1"
    If fso.FileExists(scriptPath) Then
        shell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """ -Profile " & profile, 0, False
    Else
        MsgBox "Could not find " & scriptPath, vbExclamation, "L.U.C.A.S Tunnel"
    End If
End Sub

Sub LaunchEbayBroker()
    scriptPath = appDir & "\deploy\ebay-broker\windows\run-ebay-broker.ps1"
    If fso.FileExists(scriptPath) Then
        shell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """", 0, False
    Else
        MsgBox "Could not find " & scriptPath, vbExclamation, "L.U.C.A.S eBay Broker"
    End If
End Sub

LaunchMobileServer "lucas_settings.json", "assignment_companies.json", "8765", "LUCAS_TEAM_MOBILE_PUBLIC_URL", "https://team-lucas.mikeyscards.com"
WScript.Sleep 1500
LaunchMobileServer "lucas_settings.michael.json", "assignment_companies.michael.json", "8766", "LUCAS_PERSONAL_MOBILE_PUBLIC_URL", "https://lucas.mikeyscards.com"
WScript.Sleep 1500
LaunchEbayBroker
WScript.Sleep 1500
LaunchTunnel "team"
WScript.Sleep 1500
LaunchTunnel "personal"
