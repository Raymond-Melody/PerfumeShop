<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' V17.0 健康检查API
' P0: SQL Server稳定性监控
' 用法: GET /api/health_check.asp
' ============================================
Response.Charset = "UTF-8"
Response.ContentType = "application/json"
%>
<!--#include file="../includes/config.asp"-->
<%
Dim dbStatus, dbProvider, dbError, startTime
dbStatus = "unknown"
dbProvider = "unknown"
dbError = ""
startTime = Timer()

' 1. 数据库连接测试
On Error Resume Next
Dim diagConn
Set diagConn = Server.CreateObject("ADODB.Connection")
If Err.Number = 0 Then
    diagConn.Open "Provider=MSOLEDBSQL;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;TrustServerCertificate=yes;"
    If Err.Number = 0 Then
        dbStatus = "ok"
        dbProvider = "MSOLEDBSQL"
    Else
        dbError = Err.Description
        Err.Clear
        Set diagConn = Nothing
        Set diagConn = Server.CreateObject("ADODB.Connection")
        diagConn.Open "Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;"
        If Err.Number = 0 Then
            dbStatus = "ok_fallback"
            dbProvider = "SQLOLEDB"
        Else
            dbStatus = "fail"
            If dbError <> "" Then dbError = dbError & "; " & Err.Description
        End If
    End If
End If

' 2. 数据库元数据
Dim dbVersion, dbName, dbSizeMB, tableCount
dbVersion = ""
dbName = ""
dbSizeMB = 0
tableCount = 0

If dbStatus <> "fail" And Not diagConn Is Nothing Then
    Dim diagRs
    Set diagRs = diagConn.Execute("SELECT @@VERSION AS Ver")
    If Not diagRs Is Nothing And Not diagRs.EOF Then
        dbVersion = Left(diagRs("Ver") & "", 100)
        diagRs.Close
    End If
    Set diagRs = Nothing
    
    Set diagRs = diagConn.Execute("SELECT DB_NAME() AS DB")
    If Not diagRs Is Nothing And Not diagRs.EOF Then
        dbName = diagRs("DB") & ""
        diagRs.Close
    End If
    Set diagRs = Nothing
    
    Set diagRs = diagConn.Execute("SELECT CAST(SUM(size) * 8 / 1024.0 AS DECIMAL(10,2)) FROM sys.database_files")
    If Not diagRs Is Nothing And Not diagRs.EOF Then
        If Not IsNull(diagRs(0)) Then dbSizeMB = CDbl(diagRs(0))
        diagRs.Close
    End If
    Set diagRs = Nothing
    
    Set diagRs = diagConn.Execute("SELECT COUNT(*) FROM sys.tables")
    If Not diagRs Is Nothing And Not diagRs.EOF Then
        tableCount = CLng(diagRs(0))
        diagRs.Close
    End If
    Set diagRs = Nothing
    
    diagConn.Close
End If
Set diagConn = Nothing
Err.Clear

Dim responseTimeMs
responseTimeMs = Round((Timer() - startTime) * 1000, 1)

' 3. 磁盘空间检查
Dim diskFreeMB, diskStatus
diskFreeMB = 0
diskStatus = "ok"
Dim fso, drive
Set fso = CreateObject("Scripting.FileSystemObject")
If Err.Number = 0 Then
    Set drive = fso.GetDrive(fso.GetDriveName(Server.MapPath("/")))
    If Err.Number = 0 Then
        diskFreeMB = CLng(drive.FreeSpace / 1024 / 1024)
        If diskFreeMB > 0 And diskFreeMB < 500 Then diskStatus = "warn"
    End If
    Set drive = Nothing
End If
Set fso = Nothing
Err.Clear
On Error GoTo 0

' 4. 健康状态判定
Dim overallStatus
If dbStatus = "ok" Or dbStatus = "ok_fallback" Then
    overallStatus = "healthy"
Else
    overallStatus = "unhealthy"
End If

' 5. Feature flags (Avoid IIf with Const - VBScript bug)
Dim ms, dal, i18, pw3
If FEATURE_MSOLEDBSQL Then ms = "true" Else ms = "false"
If FEATURE_DAL_ENABLED Then dal = "true" Else dal = "false"
If FEATURE_I18N Then i18 = "true" Else i18 = "false"
If FEATURE_PASSWORD_V3 Then pw3 = "true" Else pw3 = "false"

' 6. 输出JSON
Response.Write "{"
Response.Write """status"": """ & overallStatus & ""","
Response.Write """version"": """ & SYS_VERSION & ""","
Response.Write """responseTimeMs"": " & responseTimeMs & ","
Response.Write """database"": {"
Response.Write """status"": """ & dbStatus & ""","
Response.Write """provider"": """ & dbProvider & ""","
Response.Write """version"": """ & Replace(dbVersion, """", "\""") & ""","
Response.Write """name"": """ & dbName & ""","
Response.Write """sizeMB"": " & dbSizeMB & ","
Response.Write """tableCount"": " & tableCount
If dbError <> "" Then Response.Write ", ""error"": """ & Replace(dbError, """", "\""") & """"
Response.Write "},"
Response.Write """disk"": {"
Response.Write """freeMB"": " & diskFreeMB & ","
Response.Write """status"": """ & diskStatus & """"
Response.Write "},"
Response.Write """asp"": {"
Response.Write """sessionTimeout"": " & Session.Timeout & ","
Response.Write """scriptTimeout"": " & Server.ScriptTimeout
Response.Write "},"
Response.Write """featureFlags"": {"
Response.Write """MSOLEDBSQL"": " & ms & ","
Response.Write """DAL_ENABLED"": " & dal & ","
Response.Write """I18N"": " & i18 & ","
Response.Write """PASSWORD_V3"": " & pw3
Response.Write "}"
Response.Write "}"
%>
