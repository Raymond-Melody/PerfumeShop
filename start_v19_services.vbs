' V19 Services 自动启动器
' 此脚本会在用户登录时自动启动 V19 API 和 Admin 服务
Dim shell, fso, scriptDir, batchPath
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
batchPath = scriptDir & "\start_v19_services.bat"
' 静默运行，不显示命令窗口
shell.Run """" & batchPath & """", 7, False
