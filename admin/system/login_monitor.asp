<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
Call OpenConnection()

Function GetScalar(sql)
    Dim rs, val : val = ""
    On Error Resume Next
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        If Not rs Is Nothing Then
            If Not rs.EOF Then
                val = rs(0)
                rs.Close
            End If
        End If
    Else
        Err.Clear
    End If
    Set rs = Nothing : GetScalar = val
End Function

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

' 自动创建 LoginAlerts 表
On Error Resume Next
conn.Execute "SELECT TOP 1 * FROM LoginAlerts WHERE 1=0"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE LoginAlerts (" & _
        "AlertID INT IDENTITY(1,1) PRIMARY KEY," & _
        "AlertType NVARCHAR(50)," & _
        "AlertLevel NVARCHAR(20) DEFAULT 'info'," & _
        "AlertMessage NVARCHAR(500)," & _
        "IPAddress NVARCHAR(50)," & _
        "AdminID INT NULL," & _
        "IsRead BIT DEFAULT 0," & _
        "CreatedAt DATETIME DEFAULT GETDATE()" & _
        ")"
End If
On Error GoTo 0

' 生成告警
Dim msg : msg = ""
If Request.QueryString("action") = "check" Then
    ' 检查24小时内连续失败登录
    Dim failedIPs
    Set failedIPs = conn.Execute("SELECT IPAddress, COUNT(*) AS Cnt FROM AdminLogs " & _
        "WHERE ActionType='登录失败' AND CreatedAt >= DATEADD(HOUR,-24,GETDATE()) " & _
        "GROUP BY IPAddress HAVING COUNT(*) >= 5")
    
    If Not failedIPs Is Nothing Then
        Do While Not failedIPs.EOF
            Dim fIP : fIP = failedIPs("IPAddress")
            Dim fCnt : fCnt = failedIPs("Cnt")
            ' 检查是否已在黑名单
            Dim inBL : inBL = SafeNum(GetScalar("SELECT COUNT(*) FROM IPBlacklist WHERE IPAddress='" & Replace(fIP,"'","''") & "' AND IsActive=1"))
            
            If inBL = 0 Then
                ' 加入黑名单
                On Error Resume Next
                conn.Execute "INSERT INTO IPBlacklist (IPAddress, Reason, BlockedBy, ExpiresAt) VALUES ('" & _
                    Replace(fIP,"'","''") & "', '自动封禁：24h内" & fCnt & "次登录失败', " & Session("AdminID") & ", DATEADD(DAY,7,GETDATE()))"
                
                ' 记录告警
                conn.Execute "INSERT INTO LoginAlerts (AlertType, AlertLevel, AlertMessage, IPAddress) VALUES (" & _
                    "'auto_block', 'high', '自动封禁IP: " & fIP & "（24h内" & fCnt & "次登录失败）', '" & fIP & "')"
                On Error GoTo 0
            Else
                ' 记录告警
                On Error Resume Next
                conn.Execute "INSERT INTO LoginAlerts (AlertType, AlertLevel, AlertMessage, IPAddress, IsRead) VALUES (" & _
                    "'repeat_attack', 'critical', '已封禁IP再次尝试: " & fIP & "（24h内" & fCnt & "次登录失败）', '" & fIP & "', 0)"
                On Error GoTo 0
            End If
            
            failedIPs.MoveNext
        Loop
        failedIPs.Close
    End If
    Set failedIPs = Nothing
    
    msg = "告警检查完成，已自动处理异常登录"
End If

' 标记已读
If Request.Form("mark_read") <> "" Then
    conn.Execute "UPDATE LoginAlerts SET IsRead=1 WHERE AlertID=" & CInt(Request.Form("mark_read"))
End If

' 清空告警
If Request.Form("clear_all") <> "" Then
    conn.Execute "DELETE FROM LoginAlerts WHERE IsRead=1"
    msg = "已清空已读告警"
End If

' 最近登录
Dim rsLogin
Set rsLogin = ExecuteQuery("SELECT TOP 50 al.*, " & _
    "CASE WHEN al.ActionType='登录成功' THEN 1 WHEN al.ActionType='登录失败' THEN 0 ELSE -1 END AS IsSuccess " & _
    "FROM AdminLogs al WHERE al.ActionType IN ('登录成功','登录失败') ORDER BY al.CreatedAt DESC")

' 告警列表
Dim rsAlerts
Set rsAlerts = ExecuteQuery("SELECT TOP 30 * FROM LoginAlerts ORDER BY IsRead ASC, CreatedAt DESC")

' 最近7天统计
Dim todaySuccess, todayFail, weekSuccess, weekFail
todaySuccess = SafeNum(GetScalar("SELECT COUNT(*) FROM AdminLogs WHERE ActionType='登录成功' AND CreatedAt>=DATEADD(HOUR,-24,GETDATE())"))
todayFail = SafeNum(GetScalar("SELECT COUNT(*) FROM AdminLogs WHERE ActionType='登录失败' AND CreatedAt>=DATEADD(HOUR,-24,GETDATE())"))
weekSuccess = SafeNum(GetScalar("SELECT COUNT(*) FROM AdminLogs WHERE ActionType='登录成功' AND CreatedAt>=DATEADD(DAY,-7,GETDATE())"))
weekFail = SafeNum(GetScalar("SELECT COUNT(*) FROM AdminLogs WHERE ActionType='登录失败' AND CreatedAt>=DATEADD(DAY,-7,GETDATE())"))

' 在线管理员数
Dim onlineAdmins : onlineAdmins = SafeNum(GetScalar("SELECT COUNT(*) FROM AdminUsers WHERE LastLogin >= DATEADD(HOUR,-1,GETDATE())"))

' 未读告警数
Dim unreadAlerts : unreadAlerts = SafeNum(GetScalar("SELECT COUNT(*) FROM LoginAlerts WHERE IsRead=0"))
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>登录监控 - 站点技术管理</title>
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
        .stat-card.success .num { color: #4CAF50; }
        .stat-card.danger .num { color: #f44336; }
        .stat-card.warn .num { color: #FF9800; }
        .msg { padding: 12px 20px; border-radius: 8px; margin-bottom: 20px; font-weight: 500; }
        .msg-success { background: rgba(76,175,80,0.15); color: #4CAF50; border: 1px solid rgba(76,175,80,0.3); }
        .card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); margin-bottom: 20px; overflow: hidden; }
        .card-header { padding: 15px 20px; border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 16px; display: flex; justify-content: space-between; align-items: center; }
        .card-body { padding: 20px; }
        table { width: 100%; border-collapse: collapse; }
        th { text-align: left; padding: 12px; background: rgba(0,188,212,0.06); border-bottom: 1px solid rgba(255,255,255,0.06); font-size: 13px; color: #999; }
        td { padding: 10px 12px; border-bottom: 1px solid rgba(255,255,255,0.04); font-size: 13px; }
        tr:hover { background: rgba(255,255,255,0.03); }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
        .badge-success { background: rgba(76,175,80,0.2); color: #A5D6A7; }
        .badge-danger { background: rgba(244,67,54,0.2); color: #EF9A9A; }
        .badge-info { background: rgba(0,188,212,0.2); color: #80DEEA; }
        .badge-warn { background: rgba(255,152,0,0.2); color: #FFB74D; }
        .badge-critical { background: rgba(183,28,28,0.3); color: #FF8A80; border: 1px solid rgba(183,28,28,0.5); }
        .badge-high { background: rgba(244,67,54,0.2); color: #EF9A9A; }
        .badge-low { background: rgba(76,175,80,0.2); color: #A5D6A7; }
        .alert-row { background: rgba(244,67,54,0.06); }
        .alert-row.unread { border-left: 3px solid #f44336; }
        .pulse { animation: pulse 2s infinite; }
        @keyframes pulse { 0%,100% { opacity: 1; } 50% { opacity: 0.5; } }
        .login-success { color: #4CAF50; }
        .login-fail { color: #f44336; }
        .ip-text { font-family: 'Consolas',monospace; font-size: 12px; color: #888; }
        .time-text { color: #888; font-size: 12px; }
        .week-summary { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }
        .summary-item { background: rgba(255,255,255,0.03); padding: 15px; border-radius: 8px; text-align: center; }
        .summary-item .val { font-size: 22px; font-weight: 700; }
        .summary-item .sub { font-size: 12px; color: #888; margin-top: 3px; }
        .no-data { text-align: center; padding: 30px; color: #666; }
        .no-data i { font-size: 36px; display: block; margin-bottom: 10px; opacity: 0.3; }
        @media (max-width: 768px) { .main-content { margin-left: 0; } .stats-grid { grid-template-columns: 1fr 1fr; } }
    </style>
</head>
<body>
<!--#include file="includes/nav.asp"-->
<div class="main-content">
    <div class="page-header">
        <h2 class="page-title"><i class="fas fa-user-shield"></i> 登录监控</h2>
        <div class="breadcrumb"><a href="index.asp">系统中心</a> / 登录监控</div>
    </div>

    <% If msg <> "" Then %>
    <div class="msg msg-<%= IIf(InStr(msg,"完成")>0,"success","success") %>"><i class="fas fa-info-circle"></i> <%= msg %></div>
    <% End If %>

    <!-- 统计卡片 -->
    <div class="stats-grid">
        <div class="stat-card success"><div class="num"><%= todaySuccess %></div><div class="label">24h登录成功</div></div>
        <div class="stat-card danger"><div class="num"><%= todayFail %></div><div class="label">24h登录失败</div></div>
        <div class="stat-card"><div class="num"><%= onlineAdmins %></div><div class="label">当前在线管理</div></div>
        <div class="stat-card <%= IIf(unreadAlerts>0,"warn","") %>">
            <div class="num"><%= IIf(unreadAlerts>0,"<i class='fas fa-exclamation-triangle pulse'></i> ","") %><%= unreadAlerts %></div>
            <div class="label">未读告警</div>
        </div>
    </div>

    <!-- 操作栏 -->
    <div style="display:flex;gap:10px;margin-bottom:20px;flex-wrap:wrap;">
        <a href="?action=check" class="btn btn-warn"><i class="fas fa-search"></i> 运行告警检查</a>
        <form method="post" style="display:inline;">
            <input type="hidden" name="clear_all" value="1">
            <button type="submit" class="btn btn-ghost"><i class="fas fa-trash-alt"></i> 清空已读告警</button>
        </form>
        <a href="ip_blacklist.asp" class="btn btn-ghost"><i class="fas fa-ban"></i> 黑名单管理</a>
    </div>

    <!-- 7天统计 -->
    <div class="card">
        <div class="card-header"><i class="fas fa-chart-bar"></i> 近7天登录统计</div>
        <div class="card-body">
            <div class="week-summary">
                <div class="summary-item">
                    <div class="val" style="color:#4CAF50;"><%= weekSuccess %></div>
                    <div class="sub">登录成功</div>
                </div>
                <div class="summary-item">
                    <div class="val" style="color:<%= IIf(weekFail>10,"#f44336","#FF9800") %>;"><%= weekFail %></div>
                    <div class="sub">登录失败</div>
                </div>
            </div>
            <div style="text-align:center;margin-top:12px;font-size:12px;color:#888;">
                <% If weekSuccess > 0 Or weekFail > 0 Then %>
                成功率: <%= IIf(weekSuccess+weekFail>0, FormatNumber(weekSuccess/(weekSuccess+weekFail)*100,1) & "%", "N/A") %>
                <% End If %>
            </div>
        </div>
    </div>

    <!-- 告警列表 -->
    <div class="card">
        <div class="card-header">
            <span><i class="fas fa-bell"></i> 安全告警 (<%= IIf(unreadAlerts>0,"<span style='color:#f44336;'>" & unreadAlerts & " 未读</span>","无未读") %>)</span>
        </div>
        <div class="card-body">
            <table>
                <thead><tr><th style="width:40px;"></th><th>类型</th><th>级别</th><th>告警信息</th><th>时间</th><th>操作</th></tr></thead>
                <tbody>
                    <% If Not rsAlerts Is Nothing Then
                        Dim alertHasRows : alertHasRows = False
                        Do While Not rsAlerts.EOF
                            alertHasRows = True
                            Dim aLevel : aLevel = rsAlerts("AlertLevel")
                            Dim levelBadge
                            Select Case aLevel
                                Case "critical": levelBadge = "badge-critical"
                                Case "high": levelBadge = "badge-high"
                                Case "medium": levelBadge = "badge-warn"
                                Case "low": levelBadge = "badge-low"
                                Case Else: levelBadge = "badge-info"
                            End Select
                    %>
                    <tr class="<%= IIf(CBool(rsAlerts("IsRead") And 1),"","alert-row unread") %>">
                        <td><%= IIf(CBool(rsAlerts("IsRead") And 1),"<i class='fas fa-envelope-open' style='color:#666;'></i>","<i class='fas fa-envelope pulse' style='color:#f44336;'></i>") %></td>
                        <td><span class="badge badge-info"><%= rsAlerts("AlertType") %></span></td>
                        <td><span class="badge <%= levelBadge %>"><%= UCase(aLevel) %></span></td>
                        <td style="max-width:350px;word-break:break-all;"><%= rsAlerts("AlertMessage") %></td>
                        <td class="time-text"><%= rsAlerts("CreatedAt") %></td>
                        <td>
                            <form method="post" style="display:inline;">
                                <input type="hidden" name="mark_read" value="<%= rsAlerts("AlertID") %>">
                                <button type="submit" class="btn btn-ghost btn-sm" <%= IIf(CBool(rsAlerts("IsRead") And 1),"disabled","") %>>
                                    <i class="fas fa-check"></i> 已读
                                </button>
                            </form>
                        </td>
                    </tr>
                    <% 
                            rsAlerts.MoveNext
                        Loop
                        rsAlerts.Close
                        If Not alertHasRows Then
                    %>
                    <tr><td colspan="6" class="no-data"><i class="fas fa-shield-alt" style="color:#4CAF50;"></i>暂无安全告警</td></tr>
                    <% End If
                    Else %>
                    <tr><td colspan="6" class="no-data"><i class="fas fa-shield-alt" style="color:#4CAF50;"></i>暂无安全告警</td></tr>
                    <% End If %>
                </tbody>
            </table>
        </div>
    </div>

    <!-- 最近登录记录 -->
    <div class="card">
        <div class="card-header"><i class="fas fa-history"></i> 最近50条登录记录</div>
        <div class="card-body">
            <table>
                <thead><tr><th>时间</th><th>管理员</th><th>结果</th><th>IP地址</th><th>备注</th></tr></thead>
                <tbody>
                    <% If Not rsLogin Is Nothing Then
                        Dim loginHasRows : loginHasRows = False
                        Do While Not rsLogin.EOF
                            loginHasRows = True
                            Dim isSuccess : isSuccess = CInt(rsLogin("IsSuccess"))
                    %>
                    <tr>
                        <td class="time-text"><%= rsLogin("CreatedAt") %></td>
                        <td><%= IIf(Not IsNull(rsLogin("AdminID")),rsLogin("AdminID"),"—") %></td>
                        <td>
                            <% If isSuccess = 1 Then %>
                            <span class="badge badge-success"><i class="fas fa-check"></i> 成功</span>
                            <% ElseIf isSuccess = 0 Then %>
                            <span class="badge badge-danger"><i class="fas fa-times"></i> 失败</span>
                            <% Else %>
                            <span class="badge badge-info">其他</span>
                            <% End If %>
                        </td>
                        <td class="ip-text"><%= IIf(Not IsNull(rsLogin("IPAddress")),rsLogin("IPAddress"),"—") %></td>
                        <td style="font-size:12px;color:#999;"><%= IIf(Not IsNull(rsLogin("Notes")),rsLogin("Notes"),"") %></td>
                    </tr>
                    <% 
                            rsLogin.MoveNext
                        Loop
                        rsLogin.Close
                        If Not loginHasRows Then
                    %>
                    <tr><td colspan="5" class="no-data"><i class="fas fa-inbox"></i>暂无登录记录</td></tr>
                    <% End If
                    Else %>
                    <tr><td colspan="5" class="no-data"><i class="fas fa-inbox"></i>暂无登录记录</td></tr>
                    <% End If %>
                </tbody>
            </table>
        </div>
    </div>

    <!-- 监控规则说明 -->
    <div class="card">
        <div class="card-header"><i class="fas fa-cog"></i> 监控规则</div>
        <div class="card-body" style="font-size:13px;color:#999;line-height:1.8;">
            <p><i class="fas fa-shield-alt" style="color:#f44336;"></i> <strong>暴力破解检测</strong>：同一IP 24小时内失败 ≥ 5次 → 自动加入黑名单（7天）</p>
            <p><i class="fas fa-shield-alt" style="color:#FF9800;"></i> <strong>已封禁IP重试</strong>：黑名单IP再次尝试登录 → 发送严重告警</p>
            <p><i class="fas fa-shield-alt" style="color:#00bcd4;"></i> <strong>异常登录检测</strong>：异地IP登录 → 触发中等告警</p>
            <p><i class="fas fa-info-circle" style="color:#888;"></i> 点击"运行告警检查"按钮可手动触发所有规则检查</p>
        </div>
    </div>
</div>
</body>
</html>
<%
Call CloseConnection()
%>
