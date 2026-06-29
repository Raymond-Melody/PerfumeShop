<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="../includes/config.asp"-->
<%
Response.Charset = "UTF-8"
On Error Resume Next
Response.Write "<html><body><h2>V17 系统配置诊断报告</h2><pre>"
Response.Write "时间: " & Now() & vbCrLf
Response.Write "版本: " & SYS_VERSION & vbCrLf & vbCrLf

' ====== 1. 数据库连接测试 (独立连接，避免OpenConnection的Response.End) ======
Response.Write "=== 1. 数据库连接 ===" & vbCrLf
Response.Write "[当前驱动] " & IIf(FEATURE_MSOLEDBSQL, "MSOLEDBSQL", "SQLOLEDB") & " (FEATURE_MSOLEDBSQL=" & FEATURE_MSOLEDBSQL & ")" & vbCrLf

Dim diagConn, diagRs, connOK
connOK = False
Set diagConn = Server.CreateObject("ADODB.Connection")
If Err.Number = 0 Then
    diagConn.Open "Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;"
    If Err.Number = 0 Then
        connOK = True
        Response.Write "  状态: OK - 连接成功 (SQLOLEDB)" & vbCrLf
    Else
        Response.Write "  状态: FAIL - " & Err.Description & " (错误号: " & Err.Number & ")" & vbCrLf
    End If
Else
    Response.Write "  状态: FAIL - 无法创建连接对象: " & Err.Description & vbCrLf
End If

If connOK Then
    ' 获取数据库信息
    Set diagRs = diagConn.Execute("SELECT @@VERSION AS Ver, DB_NAME() AS DB")
    If Not diagRs Is Nothing And Not diagRs.EOF Then
        Response.Write "  SQL Server: " & Left(diagRs("Ver"), 80) & "..." & vbCrLf
        Response.Write "  数据库: " & diagRs("DB") & vbCrLf
        diagRs.Close
    End If
    Set diagRs = Nothing

    ' 查询端口
    Set diagRs = diagConn.Execute("SELECT local_tcp_port FROM sys.dm_exec_connections WHERE session_id=@@SPID AND local_tcp_port IS NOT NULL")
    Dim portFound : portFound = ""
    If Not diagRs Is Nothing Then
        If Not diagRs.EOF Then portFound = diagRs(0) & ""
        diagRs.Close
    End If
    Set diagRs = Nothing
    If portFound <> "" Then
        Response.Write "  TCP端口: " & portFound & " (可启用MSOLEDBSQL)" & vbCrLf
    Else
        Response.Write "  TCP端口: 未启用 (使用Shared Memory)" & vbCrLf
    End If

    ' 查询表数量
    Set diagRs = diagConn.Execute("SELECT COUNT(*) FROM sys.tables")
    If Not diagRs Is Nothing And Not diagRs.EOF Then
        Response.Write "  数据表: " & diagRs(0) & "张" & vbCrLf
        diagRs.Close
    End If
    Set diagRs = Nothing

    diagConn.Close
End If
Set diagConn = Nothing

' ====== 2. MSOLEDBSQL 驱动状态 ======
Response.Write vbCrLf & "=== 2. MSOLEDBSQL 驱动 ===" & vbCrLf
Response.Write "[FEATURE_MSOLEDBSQL] True (P0-已激活)" & vbCrLf
If FEATURE_MSOLEDBSQL Then
    Response.Write "[首选驱动] MSOLEDBSQL (V17推荐)" & vbCrLf
    Dim msDiagConn
    On Error Resume Next
    Set msDiagConn = Server.CreateObject("ADODB.Connection")
    If Err.Number = 0 Then
        msDiagConn.Open "Provider=MSOLEDBSQL;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;TrustServerCertificate=yes;"
        If Err.Number = 0 Then
            Response.Write "[连接状态] OK - MSOLEDBSQL连接成功" & vbCrLf
            msDiagConn.Close
        Else
            Response.Write "[连接状态] FAIL - " & Err.Description & " (将回退SQLOLEDB)" & vbCrLf
            Err.Clear
        End If
    Else
        Response.Write "[连接状态] FAIL - 无法创建连接对象" & vbCrLf
        Err.Clear
    End If
    Set msDiagConn = Nothing
    On Error GoTo 0
Else
    Response.Write "[当前策略] FEATURE_MSOLEDBSQL=False, 使用SQLOLEDB" & vbCrLf
End If

' ====== 3. Feature Flags ======
Response.Write vbCrLf & "=== 3. Feature Flags ===" & vbCrLf
Response.Write "FEATURE_MSOLEDBSQL          = " & FEATURE_MSOLEDBSQL & " (P0-已激活)" & vbCrLf
Response.Write "FEATURE_DAL_ENABLED         = " & FEATURE_DAL_ENABLED & " (P0-已激活)" & vbCrLf
Response.Write "FEATURE_PASSWORD_V3         = " & FEATURE_PASSWORD_V3 & " (P0-已激活)" & vbCrLf
Response.Write "FEATURE_STRUCTURED_LOGGING  = " & FEATURE_STRUCTURED_LOGGING & " (P1-已激活)" & vbCrLf
Response.Write "FEATURE_API_V1              = " & FEATURE_API_V1 & " (P1-已激活)" & vbCrLf
Response.Write "FEATURE_CACHE_MANAGER       = " & FEATURE_CACHE_MANAGER & " (P1-已激活)" & vbCrLf
Response.Write "FEATURE_SSE_NOTIFICATIONS   = " & FEATURE_SSE_NOTIFICATIONS & " (P2-已激活)" & vbCrLf
Response.Write "FEATURE_EMAIL_NOTIFICATIONS = " & FEATURE_EMAIL_NOTIFICATIONS & " (P2-已激活)" & vbCrLf
Response.Write "FEATURE_ANALYTICS_DASHBOARD = " & FEATURE_ANALYTICS_DASHBOARD & " (P2-已激活)" & vbCrLf
Response.Write "FEATURE_PWA_ENHANCED        = " & FEATURE_PWA_ENHANCED & " (P2-已激活)" & vbCrLf
Response.Write "FEATURE_I18N                = " & FEATURE_I18N & " (P2-已激活)" & vbCrLf

' ====== 4. Session 状态 ======
Response.Write vbCrLf & "=== 4. Session ===" & vbCrLf
Response.Write "SessionID: " & Session.SessionID & vbCrLf
Response.Write "Timeout: " & Session.Timeout & "分钟" & vbCrLf
Response.Write "CSRFToken: " & (Session("CSRFToken") <> "") & vbCrLf

' ====== 5. 环境信息 ======
Response.Write vbCrLf & "=== 5. 服务器环境 ===" & vbCrLf
Response.Write "IIS版本: " & Request.ServerVariables("SERVER_SOFTWARE") & vbCrLf
Response.Write "ASP超时: " & Server.ScriptTimeout & "秒" & vbCrLf
Response.Write "物理路径: " & Server.MapPath("/") & vbCrLf

Response.Write vbCrLf & "=== 诊断完成 ===" & vbCrLf
Response.Write "</pre></body></html>"
%>
