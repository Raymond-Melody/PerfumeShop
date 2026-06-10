<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
Call OpenConnection()

' SafeNum 已在 connection.asp 中定义，此处不再重复定义

' 扫描类型
Dim scanType : scanType = Request.QueryString("type")
Dim results, scanCount : scanCount = 0

If scanType <> "" Then
    ReDim results(0, 2)
    
    If scanType = "sqli" Or scanType = "all" Then
        ' SQL注入风险扫描 - 检查包含 Request + 拼接的 ASP 文件
        Dim fso, folder, file, fContent, lineNum, riskLine, riskCount
        riskCount = 0
        On Error Resume Next
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        If Not fso Is Nothing Then
            Set folder = fso.GetFolder(Server.MapPath("/admin"))
            ScanFolder folder, fso, riskCount
            Set folder = Nothing
            Set fso = Nothing
        End If
        On Error GoTo 0
    End If
    
    If scanType = "xss" Or scanType = "all" Then
        ' XSS检测 - 检查 Response.Write 是否使用 HTMLEncode
        Dim xssRisk : xssRisk = CheckXSSRisk()
        If xssRisk > 0 Then
            ReDim Preserve results(scanCount, 2)
            results(scanCount, 0) = "XSS风险"
            results(scanCount, 1) = "中"
            results(scanCount, 2) = "发现 " & xssRisk & " 处未转义的Response.Write输出，建议使用Server.HTMLEncode"
            scanCount = scanCount + 1
        End If
    End If
    
    If scanType = "auth" Or scanType = "all" Then
        ' 权限检查 - 验证所有admin子目录是否有auth.asp
        CheckAuthFiles
    End If
End If

Sub ScanFolder(folder, fso, ByRef riskCount)
    Dim f, ext
    For Each f In folder.Files
        ext = LCase(fso.GetExtensionName(f.Name))
        If ext = "asp" Then
            On Error Resume Next
            fContent = f.OpenAsTextStream(1).ReadAll
            If Err.Number = 0 Then
                ' 检测不安全的SQL拼接模式
                If InStr(1, fContent, "Request.", 1) > 0 And InStr(1, fContent, "SafeSQL(", 1) = 0 Then
                    If InStr(1, fContent, "conn.Execute", 1) > 0 Or InStr(1, fContent, "ExecuteQuery", 1) > 0 Then
                        riskCount = riskCount + 1
                        ReDim Preserve results(scanCount, 2)
                        results(scanCount, 0) = "SQL注入风险"
                        results(scanCount, 1) = "高"
                        results(scanCount, 2) = f.Path
                        scanCount = scanCount + 1
                    End If
                End If
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next
    For Each f In folder.SubFolders
        If LCase(f.Name) <> "includes" And LCase(f.Name) <> "css" Then
            ScanFolder f, fso, riskCount
        End If
    Next
End Sub

Function CheckXSSRisk()
    CheckXSSRisk = 0
    ' 抽样检查主要页面
    Dim checkFiles : checkFiles = Array("/admin/operation/orders.asp", "/admin/operation/products.asp")
    Dim cf, cfContent, i, fsoXss
    On Error Resume Next
    Set fsoXss = Server.CreateObject("Scripting.FileSystemObject")
    If fsoXss Is Nothing Then Exit Function
    For i = 0 To UBound(checkFiles)
        cfContent = ""
        Set cf = fsoXss.OpenTextFile(Server.MapPath(checkFiles(i)), 1)
        If Err.Number = 0 Then cfContent = cf.ReadAll : cf.Close
        Err.Clear
        If cfContent <> "" Then
            If InStr(cfContent, "Response.Write") > 0 And InStr(cfContent, "Server.HTMLEncode") = 0 Then
                CheckXSSRisk = CheckXSSRisk + 1
            End If
        End If
    Next
    Set fsoXss = Nothing
    On Error GoTo 0
End Function

Sub CheckAuthFiles()
    Dim adminPath, authPath, subFolders, sf, subPath
    adminPath = Server.MapPath("/admin")
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Dim missingAuth : missingAuth = ""
    For Each sf In fso.GetFolder(adminPath).SubFolders
        subPath = adminPath & "\" & sf.Name & "\includes\auth.asp"
        If Not fso.FileExists(subPath) Then
            If missingAuth <> "" Then missingAuth = missingAuth & ", "
            missingAuth = missingAuth & sf.Name
        End If
    Next
    Set fso = Nothing
    If missingAuth <> "" Then
        ReDim Preserve results(scanCount, 2)
        results(scanCount, 0) = "权限缺失"
        results(scanCount, 1) = "高"
        results(scanCount, 2) = "以下目录缺少auth.asp: " & missingAuth
        scanCount = scanCount + 1
    End If
End Sub

' 统计
Dim adminCount : adminCount = GetScalar("SELECT COUNT(*) FROM AdminUsers")
Dim logCount : logCount = GetScalar("SELECT COUNT(*) FROM AdminLogs")
Dim failedLogin : failedLogin = GetScalar("SELECT COUNT(*) FROM AdminLogs WHERE ActionType='登录失败' AND CreatedAt >= DATEADD(HOUR,-24,GETDATE())")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>安全审计 - 站点技术管理</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; font-family: 'Segoe UI',Arial,sans-serif; }
        .main-content { margin-left: 250px; padding: 30px; min-height: 100vh; }
        .page-header { margin-bottom: 25px; }
        .page-title { color: #fff; font-size: 24px; margin: 0 0 8px; }
        .breadcrumb { color: #888; font-size: 13px; }
        .breadcrumb a { color: #00bcd4; text-decoration: none; }
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); padding: 20px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); text-align: center; }
        .stat-card .num { font-size: 28px; font-weight: 700; color: #00bcd4; }
        .stat-card .label { font-size: 13px; color: #888; margin-top: 5px; }
        .stat-card.warn .num { color: #FF9800; }
        .stat-card.danger .num { color: #f44336; }
        .scan-actions { display: flex; gap: 10px; margin-bottom: 25px; flex-wrap: wrap; }
        .scan-btn { padding: 14px 28px; border: none; border-radius: 10px; font-size: 15px; cursor: pointer; transition: all 0.2s; text-decoration: none; display: inline-flex; align-items: center; gap: 10px; font-weight: 600; }
        .scan-btn.sql { background: linear-gradient(135deg, #f44336, #c62828); color: #fff; }
        .scan-btn.xss { background: linear-gradient(135deg, #FF9800, #E65100); color: #fff; }
        .scan-btn.auth { background: linear-gradient(135deg, #2196F3, #0D47A1); color: #fff; }
        .scan-btn.all { background: linear-gradient(135deg, #4CAF50, #1B5E20); color: #fff; }
        .scan-btn:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(0,0,0,0.4); }
        .card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); margin-bottom: 20px; overflow: hidden; }
        .card-header { padding: 15px 20px; border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 16px; }
        .card-body { padding: 20px; }
        table { width: 100%; border-collapse: collapse; }
        th { text-align: left; padding: 12px; background: rgba(0,188,212,0.06); border-bottom: 1px solid rgba(255,255,255,0.06); font-size: 13px; color: #999; }
        td { padding: 12px; border-bottom: 1px solid rgba(255,255,255,0.04); font-size: 14px; }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
        .badge-high { background: rgba(244,67,54,0.2); color: #EF9A9A; }
        .badge-medium { background: rgba(255,152,0,0.2); color: #FFB74D; }
        .badge-low { background: rgba(76,175,80,0.2); color: #A5D6A7; }
        .empty { text-align: center; padding: 40px; color: #666; }
        .empty i { font-size: 48px; display: block; margin-bottom: 12px; opacity: 0.3; }
        @media (max-width: 768px) { .main-content { margin-left: 0; } .stats-grid { grid-template-columns: 1fr 1fr; } }
    </style>
</head>
<body>
<!--#include file="includes/nav.asp"-->
<div class="main-content">
    <div class="page-header">
        <h2 class="page-title"><i class="fas fa-shield-alt"></i> 安全审计</h2>
        <div class="breadcrumb"><a href="index.asp">系统中心</a> / 安全审计</div>
    </div>

    <div class="stats-grid">
        <div class="stat-card"><div class="num"><%= adminCount %></div><div class="label">管理员总数</div></div>
        <div class="stat-card"><div class="num"><%= logCount %></div><div class="label">操作日志总数</div></div>
        <div class="stat-card warn"><div class="num"><%= failedLogin %></div><div class="label">24h登录失败</div></div>
        <div class="stat-card"><div class="num"><%= scanCount %></div><div class="label">当前发现风险</div></div>
    </div>

    <div class="scan-actions">
        <a href="?type=sqli" class="scan-btn sql"><i class="fas fa-database"></i> SQL注入扫描</a>
        <a href="?type=xss" class="scan-btn xss"><i class="fas fa-code"></i> XSS检测</a>
        <a href="?type=auth" class="scan-btn auth"><i class="fas fa-lock"></i> 权限审计</a>
        <a href="?type=all" class="scan-btn all"><i class="fas fa-shield-virus"></i> 全面扫描</a>
    </div>

    <% If scanType <> "" Then %>
    <div class="card">
        <div class="card-header"><i class="fas fa-clipboard-list"></i> 扫描结果 (<%= IIf(scanType="all","全面扫描", UCase(scanType)) %>)</div>
        <div class="card-body">
            <% If scanCount > 0 Then %>
            <table>
                <thead><tr><th>风险类型</th><th>严重级别</th><th>详情</th></tr></thead>
                <tbody>
                <% Dim ri
                For ri = 0 To scanCount - 1
                    Dim rBadge : rBadge = "badge-low"
                    If results(ri, 1) = "高" Then rBadge = "badge-high" Else If results(ri, 1) = "中" Then rBadge = "badge-medium"
                %>
                <tr>
                    <td><strong><%= results(ri, 0) %></strong></td>
                    <td><span class="badge <%= rBadge %>"><%= results(ri, 1) %></span></td>
                    <td style="font-size:12px;word-break:break-all;"><%= results(ri, 2) %></td>
                </tr>
                <% Next %>
                </tbody>
            </table>
            <% Else %>
            <div class="empty"><i class="fas fa-check-circle" style="color:#4CAF50;"></i>未发现安全风险</div>
            <% End If %>
        </div>
    </div>
    <% End If %>

    <!-- 安全基线清单 -->
    <div class="card">
        <div class="card-header"><i class="fas fa-clipboard-check"></i> 安全基线检查清单</div>
        <div class="card-body">
            <table>
                <thead><tr><th>检查项</th><th>状态</th><th>建议</th></tr></thead>
                <tbody>
                    <tr><td>HTTPS强制</td><td><span class="badge badge-low">待确认</span></td><td>配置IIS URL Rewrite强制HTTPS</td></tr>
                    <tr><td>SQL注入防护</td><td><span class="badge badge-medium">部分覆盖</span></td><td>所有Request输入应使用SafeSQL()</td></tr>
                    <tr><td>XSS防护</td><td><span class="badge badge-medium">部分覆盖</span></td><td>输出使用Server.HTMLEncode()</td></tr>
                    <tr><td>CSRF防护</td><td><span class="badge badge-low">已配置</span></td><td>已启用CSRF Token验证</td></tr>
                    <tr><td>会话安全</td><td><span class="badge badge-low">已配置</span></td><td>Session超时已配置</td></tr>
                    <tr><td>备份机制</td><td><span class="badge badge-medium">需完善</span></td><td>请前往备份中心配置自动备份</td></tr>
                    <tr><td>登录保护</td><td><span class="badge badge-medium">需完善</span></td><td>建议启用登录失败锁定机制</td></tr>
                </tbody>
            </table>
        </div>
    </div>
</div>
</body>
</html>
<%
Call CloseConnection()
%>
