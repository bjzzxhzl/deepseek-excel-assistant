Option Explicit

Dim fso, ws, appDir, manifestPath, manifestFile, manifestText
Dim useLocalServer, nodeCmd, excelRunning, process, carrierArgument, carrierPath

Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("WScript.Shell")
appDir = fso.GetParentFolderName(WScript.ScriptFullName)
ws.CurrentDirectory = appDir

' Public HTTPS mode does not need Node.js. Start the local server only when
' manifest.xml explicitly points to localhost.
useLocalServer = False
manifestPath = fso.BuildPath(appDir, "manifest-standalone.xml")
If fso.FileExists(manifestPath) Then
  Set manifestFile = fso.OpenTextFile(manifestPath, 1, False)
  manifestText = LCase(manifestFile.ReadAll)
  manifestFile.Close
  useLocalServer = (InStr(manifestText, "http://localhost:8090") > 0 Or _
                    InStr(manifestText, "http://127.0.0.1:8090") > 0)
End If

If useLocalServer Then
  If ws.Run("cmd /c where node >nul 2>nul", 0, True) = 0 Then
    nodeCmd = "node"
  ElseIf fso.FileExists("C:\Program Files\nodejs\node.exe") Then
    nodeCmd = """C:\Program Files\nodejs\node.exe"""
  Else
    MsgBox "Node.js was not found. Install the Node.js LTS release, then try again.", 48, "ExcelAI"
    WScript.Quit 1
  End If
  ws.Run "cmd /c " & nodeCmd & " ""server.mjs""", 0, False
End If

' A normal Excel process loads the registered carrier through OPEN/OPENn.
' If Excel was already running before installation, explicitly open the carrier
' so this shortcut can repair the current session without a background service.
excelRunning = False
On Error Resume Next
For Each process In GetObject("winmgmts:\\.\root\cimv2").ExecQuery("Select ProcessId from Win32_Process where Name='EXCEL.EXE'")
  excelRunning = True
  Exit For
Next
carrierArgument = ws.RegRead("HKCU\Software\DeepSeekExcelAssistant\ExcelStartupArgument")
On Error GoTo 0

carrierPath = Trim(carrierArgument)
If LCase(Left(carrierPath, 3)) = "/r " Then
  carrierPath = Trim(Mid(carrierPath, 4))
End If
carrierPath = Replace(carrierPath, """", "")
If excelRunning And Len(carrierPath) > 0 And fso.FileExists(carrierPath) Then
  ws.Run "excel.exe /r """ & carrierPath & """", 1, False
Else
  ws.Run "excel.exe", 1, False
End If
