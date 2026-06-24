<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' V10.4: 备份状态 API
' 返回 JSON 格式的备份系统状态
' ============================================
Response.Charset = "UTF-8"
Response.ContentType = "application/json"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<%
On Error Resume Next

Call OpenConnection()

Dim backupPath : backupPath = Server.MapPath("/database/backups/")
Dim fso : Set fso = Server.CreateObject("Scripting.FileSystemObject")

' 1. 获取最后备份信息
Dim lastBackupName, lastBackupSize, lastBackupTime, lastBackupSizeMB, fileVerified
lastBackupName = "" : lastBackupSize = 0 : lastBackupTime = "" : lastBackupSizeMB = 0 : fileVerified = False

If fso.FolderExists(backupPath) Then
    Dim latestFile, f
    Dim latestTime : latestTime = CDate("2000-01-01")
    For Each f In fso.GetFolder(backupPath).Files
        If LCase(fso.GetExtensionName(f.Name)) = "bak" Then
            If f.DateLastModified > latestTime Then
                latestTime = f.DateLastModified
                Set latestFile = f
            End If
        End If
    Next
    If Not latestFile Is Nothing Then
        lastBackupName = latestFile.Name
        lastBackupSize = latestFile.Size
        lastBackupTime = latestFile.DateLastModified
        lastBackupSizeMB = Round(latestFile.Size / 1048576, 2)
        fileVerified = (latestFile.Size > 512) ' 文件至少有512字节即视为基本有效
    End If
End If

' 2. 统计备份总数
Dim totalBackups : totalBackups = 0
If fso.FolderExists(backupPath) Then
    For Each f In fso.GetFolder(backupPath).Files
        If LCase(fso.GetExtensionName(f.Name)) = "bak" Then
            totalBackups = totalBackups + 1
        End If
    Next
End If

' 3. 最近30天备份数
Dim recentBackups : recentBackups = 0
If fso.FolderExists(backupPath) Then
    For Each f In fso.GetFolder(backupPath).Files
        If LCase(fso.GetExtensionName(f.Name)) = "bak" Then
            If DateDiff("d", f.DateLastModified, Now()) <= 30 Then
                recentBackups = recentBackups + 1
            End If
        End If
    Next
End If

' 4. 数据库大小
Dim dbSizeMB : dbSizeMB = 0
On Error Resume Next
Dim rs : Set rs = conn.Execute("SELECT SUM(size)*8/1024 FROM sys.database_files WHERE type=0")
If Not rs Is Nothing And Not rs.EOF Then
    If Not IsNull(rs(0)) Then dbSizeMB = Round(CDbl(rs(0)), 2)
    rs.Close
End If
Set rs = Nothing
On Error GoTo 0

' 5. 下次计划备份时间
Dim nextScheduled
nextScheduled = DateAdd("d", 1, Date()) & "T02:00:00"

' 6. 构建 JSON 响应
Dim json
json = "{"
json = json & """status"": ""ok"","
json = json & """version"": """ & SYS_VERSION & ""","
json = json & """lastBackup"": {"
json = json & """fileName"": """ & EscapeJSON(lastBackupName) & ""","
json = json & """sizeBytes"": " & lastBackupSize & ","
json = json & """sizeMB"": " & lastBackupSizeMB & ","
json = json & """time"": """ & EscapeJSON(lastBackupTime) & ""","
json = json & """verified"": " & LCase(fileVerified) & ""
json = json & "},"
json = json & """totals"": {"
json = json & """totalBackups"": " & totalBackups & ","
json = json & """recent30Days"": " & recentBackups & ","
json = json & """databaseSizeMB"": " & dbSizeMB & ""
json = json & "},"
json = json & """schedule"": {"
json = json & """frequency"": ""daily"","
json = json & """time"": ""02:00"","
json = json & """nextRun"": """ & EscapeJSON(nextScheduled) & """"
json = json & "},"
json = json & """generatedAt"": """ & EscapeJSON(Now()) & """"
json = json & "}"

Response.Write json

' 清理
Set fso = Nothing
Call CloseConnection()

' JSON 转义函数
Function EscapeJSON(str)
    If IsNull(str) Or IsEmpty(str) Then
        EscapeJSON = ""
        Exit Function
    End If
    str = CStr(str)
    str = Replace(str, "\", "\\")
    str = Replace(str, """", "\""")
    str = Replace(str, vbCrLf, "\n")
    str = Replace(str, vbCr, "\n")
    str = Replace(str, vbLf, "\n")
    str = Replace(str, vbTab, "\t")
    EscapeJSON = str
End Function
%>
