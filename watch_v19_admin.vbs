' V19 Admin dotnet watch 启动器
' 以窗口模式启动，可看到 dotnet watch 热重载输出
Dim shell, fso, scriptDir, psPath
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
psPath = scriptDir & "\watch_v19_admin.ps1"
' 1=正常窗口, 显示控制台用于查看热重载状态
shell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & psPath & """", 1, False
