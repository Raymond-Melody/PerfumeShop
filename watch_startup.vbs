' V19 dotnet watch auto-start (via Registry Run)
' Runs watch_v19_services.bat with visible windows
Dim shell, fso, scriptDir, batchPath
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
batchPath = scriptDir & "\watch_v19_services.bat"
shell.Run """" & batchPath & """", 1, False
