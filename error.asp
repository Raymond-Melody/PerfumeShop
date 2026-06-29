<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/dal.asp"-->
<%
' ============================================
' V17.0 统一错误处理页面
' 支持: 自定义错误码、错误消息、自动记录日志
' 用法: Response.Redirect "/error.asp?code=404&msg=页面未找到"
'       Server.Transfer "/error.asp?code=500&msg=服务器错误"
' ============================================
Response.Charset = "UTF-8"

Dim errCode, errMsg, errTitle, errIcon, showDetail, logError

' 获取错误参数
errCode = Trim(Request.QueryString("code"))
errMsg = Trim(Request.QueryString("msg"))
showDetail = Trim(Request.QueryString("detail"))
logError = Request.QueryString("log")

If errCode = "" Then errCode = "500"
If errMsg = "" Then
    If FEATURE_I18N Then
        errMsg = T("error_msg_" & errCode, Empty)
        ' Fallback: if key not found, use default
        If Left(errMsg, 1) = "[" Then
            Select Case errCode
                Case "403": errMsg = "抱歉，您没有权限访问此页面"
                Case "404": errMsg = "抱歉，您访问的页面不存在"
                Case "500": errMsg = "系统遇到了一个意外错误，请稍后重试"
                Case "503": errMsg = "系统正在维护中，请稍后访问"
                Case "400": errMsg = "请求参数错误"
                Case "429": errMsg = "请求过于频繁，请稍后再试"
                Case Else:  errMsg = T("error_msg_default", Empty)
            End Select
        End If
    Else
        Select Case errCode
            Case "403": errMsg = "抱歉，您没有权限访问此页面"
            Case "404": errMsg = "抱歉，您访问的页面不存在"
            Case "500": errMsg = "系统遇到了一个意外错误，请稍后重试"
            Case "503": errMsg = "系统正在维护中，请稍后访问"
            Case "400": errMsg = "请求参数错误"
            Case "429": errMsg = "请求过于频繁，请稍后再试"
            Case Else:  errMsg = "系统遇到了一个意外错误"
        End Select
    End If
End If

' 设置HTTP状态码
Select Case errCode
    Case "403": Response.Status = "403 Forbidden"
    Case "404": Response.Status = "404 Not Found"
    Case "500": Response.Status = "500 Internal Server Error"
    Case "503": Response.Status = "503 Service Unavailable"
    Case "400": Response.Status = "400 Bad Request"
    Case "429": Response.Status = "429 Too Many Requests"
End Select

' 设置标题和图标
Select Case errCode
    Case "403"
        If FEATURE_I18N Then errTitle = T("error_title_403", Empty) Else errTitle = "权限不足"
        errIcon = "&#x1F512;"
    Case "404"
        If FEATURE_I18N Then errTitle = T("error_title_404", Empty) Else errTitle = "页面未找到"
        errIcon = "&#x1F50D;"
    Case "500"
        If FEATURE_I18N Then errTitle = T("error_title_500", Empty) Else errTitle = "服务器错误"
        errIcon = "&#x26A0;&#xFE0F;"
    Case "503"
        If FEATURE_I18N Then errTitle = T("error_title_503", Empty) Else errTitle = "服务不可用"
        errIcon = "&#x1F527;"
    Case "400"
        If FEATURE_I18N Then errTitle = T("error_title_400", Empty) Else errTitle = "请求错误"
        errIcon = "&#x274C;"
    Case "429"
        If FEATURE_I18N Then errTitle = T("error_title_429", Empty) Else errTitle = "请求过多"
        errIcon = "&#x23F3;"
    Case Else
        If FEATURE_I18N Then errTitle = T("error_title_default", Empty) Else errTitle = "系统错误"
        errIcon = "&#x26A0;&#xFE0F;"
End Select

' 记录错误日志（如果启用了结构化日志）
If logError <> "0" Then
    On Error Resume Next
    ' V17: 使用DAL参数化查询记录错误日志，杜绝SQL注入
    Call OpenConnection()
    If conn Is Nothing Or conn.State <> 1 Then
        ' DB连接失败时静默跳过日志
        Err.Clear
    Else
        Dim logSql, logParams(4)
        logSql = "INSERT INTO AppLogs (LogLevel, LogMessage, LogSource, IPAddress, PageURL) " & _
                 "VALUES (@LogLevel, @LogMessage, @LogSource, @IPAddress, @PageURL)"
        logParams(0) = Array("@LogLevel", DAL_adVarChar, 10, "ERROR")
        logParams(1) = Array("@LogMessage", DAL_adVarChar, 2000, "错误页面(代码:" & errCode & "): " & errMsg)
        logParams(2) = Array("@LogSource", DAL_adVarChar, 100, "error.asp")
        logParams(3) = Array("@IPAddress", DAL_adVarChar, 50, Left(Request.ServerVariables("REMOTE_ADDR"), 50))
        logParams(4) = Array("@PageURL", DAL_adVarChar, 500, Left(Request.ServerVariables("SCRIPT_NAME"), 500))
        DAL_Execute logSql, logParams
    End If
    Call CloseConnection()
    Err.Clear
    On Error GoTo 0
End If
%>
<!DOCTYPE html>
<html lang="<%= I18N_HtmlLang() %>">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title><%= errTitle %> - <% If FEATURE_I18N Then %><%= T("site_name", Empty) %><% Else %>香氛定制<% End If %></title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%)}
.error-wrapper{width:100%;max-width:520px;padding:20px}
.error-container{background:#fff;border-radius:16px;box-shadow:0 20px 60px rgba(0,0,0,0.15);padding:48px 40px 40px;text-align:center}
.error-icon{font-size:64px;line-height:1;margin-bottom:8px}
.error-code{font-size:72px;font-weight:800;color:#2d3748;line-height:1;margin-bottom:4px;background:linear-gradient(135deg,#667eea,#764ba2);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}
.error-title{font-size:20px;color:#4a5568;margin-bottom:16px;font-weight:600}
.error-message{color:#718096;font-size:15px;line-height:1.6;margin-bottom:32px}
.error-detail{background:#f7fafc;border-radius:8px;padding:12px 16px;margin-bottom:24px;text-align:left}
.error-detail pre{font-size:12px;color:#a0aec0;white-space:pre-wrap;word-break:break-all;font-family:"SFMono-Regular",Consolas,"Liberation Mono",Menlo,monospace;margin:0}
.action-buttons{display:flex;gap:12px;justify-content:center;flex-wrap:wrap}
.btn{display:inline-flex;align-items:center;gap:6px;padding:12px 28px;border-radius:8px;font-size:14px;font-weight:500;text-decoration:none;transition:all 0.2s;cursor:pointer;border:none}
.btn-primary{background:#667eea;color:#fff}
.btn-primary:hover{background:#5a6fd6;transform:translateY(-1px);box-shadow:0 4px 12px rgba(102,126,234,0.4)}
.btn-secondary{background:#edf2f7;color:#4a5568}
.btn-secondary:hover{background:#e2e8f0;transform:translateY(-1px)}
@media(max-width:480px){.error-container{padding:32px 24px}.error-code{font-size:56px}}
</style>
</head>
<body>
<div class="error-wrapper">
<div class="error-container">
    <div class="error-icon"><%= errIcon %></div>
    <div class="error-code"><%= errCode %></div>
    <div class="error-title"><%= errTitle %></div>
    <div class="error-message"><%= Server.HTMLEncode(errMsg) %></div>
    <% If showDetail <> "" Then %>
    <div class="error-detail">
        <pre><%= Server.HTMLEncode(showDetail) %></pre>
    </div>
    <% End If %>
    <div class="action-buttons">
        <a href="javascript:history.back()" class="btn btn-secondary">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M19 12H5M12 19l-7-7 7-7"/></svg>
            <% If FEATURE_I18N Then %><%= T("error_btn_back", Empty) %><% Else %>返回上页<% End If %>
        </a>
        <a href="/" class="btn btn-primary">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/></svg>
            <% If FEATURE_I18N Then %><%= T("error_btn_home", Empty) %><% Else %>返回首页<% End If %>
        </a>
    </div>
</div>
</div>
</body>
</html>
