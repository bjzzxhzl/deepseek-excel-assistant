' ExcelAI - 隐藏窗口启动服务器（无黑窗，空闲自动退出）
Set fso = CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
Set ws = CreateObject("WScript.Shell")
ws.CurrentDirectory = dir

' 探测 node
If ws.Run("cmd /c where node >nul 2>nul", 0, True) = 0 Then
  nodeCmd = "node"
ElseIf fso.FileExists("C:\Program Files\nodejs\node.exe") Then
  nodeCmd = """C:\Program Files\nodejs\node.exe"""
Else
  MsgBox "未检测到 Node.js 运行环境，请先到 https://nodejs.org 安装 LTS 版本。", 48, "ExcelAI"
  WScript.Quit 1
End If

' 隐藏窗口启动（0 = 隐藏），空闲 15 分钟无任务窗格心跳时服务器自动退出
ws.Run "cmd /c " & nodeCmd & " ""server.mjs""", 0, False
