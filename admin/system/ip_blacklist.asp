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

' 自动创建 IPBlacklist 表
On Error Resume Next
conn.Execute "SELECT TOP 1 * FROM IPBlacklist WHERE 1=0"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE IPBlacklist (" & _
        "IPID INT IDENTITY(1,1) PRIMARY KEY," & _
        "IPAddress NVARCHAR(50) NOT NULL," & _
        "Reason NVARCHAR(255)," & _
        "BlockedAt DATETIME DEFAULT GETDATE()," & _
        "BlockedBy INT," & _
        "IsActive BIT DEFAULT 1," & _
        "ExpiresAt DATETIME NULL," & _
        "HitCount INT DEFAULT 0," & _
        "LastHitAt DATETIME NULL" & _
        ")"
End If
On Error GoTo 0

Dim msg : msg = ""
Dim msgType : msgType = ""

' 处理 POST
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim action : action = Request.Form("action")
    
    If action = "add" Then
        Dim ipAddr : ipAddr = Trim(Request.Form("ipAddress"))
        Dim reason : reason = Trim(Request.Form("reason"))
        Dim expiresDays : expiresDays = CInt("0" & Request.Form("expiresDays"))
        
        If ipAddr <> "" Then
            ' 检查是否已存在
            Dim existCheck : existCheck = GetScalar("SELECT COUNT(*) FROM IPBlacklist WHERE IPAddress='" & Replace(ipAddr,"'","''") & "' AND IsActive=1")
            If CDbl("0" & existCheck) > 0 Then
                msg = "IP " & ipAddr & " 已在黑名单中"
                msgType = "warning"
            Else
                Dim expireSQL : expireSQL = "NULL"
                If expiresDays > 0 Then expireSQL = "DATEADD(DAY," & expiresDays & ",GETDATE())"
                
                conn.Execute "INSERT INTO IPBlacklist (IPAddress, Reason, BlockedBy, ExpiresAt) VALUES ('" & _
                    Replace(ipAddr,"'","''") & "', '" & Replace(reason,"'","''") & "', " & Session("AdminID") & ", " & expireSQL & ")"
                msg = "IP " & ipAddr & " 已加入黑名单"
                msgType = "success"
            End If
        End If
    
    ElseIf action = "remove" Then
        Dim ipID : ipID = Request.Form("ipID")
        conn.Execute "UPDATE IPBlacklist SET IsActive=0 WHERE IPID=" & CInt(ipID)
        msg = "已移除黑名单记录"
        msgType = "success"
    
    ElseIf action = "unblock" Then
        Dim unblockID : unblockID = Request.Form("unblockID")
        conn.Execute "UPDATE IPBlacklist SET IsActive=0 WHERE IPID=" & CInt(unblockID)
        msg = "已解除封锁"
        msgType = "success"
    End If
End If

' 获取黑名单列表
Dim rsBlacklist
Set rsBlacklist = ExecuteQuery("SELECT b.*, a.Username AS BlockedByName FROM IPBlacklist b " & _
    "LEFT JOIN AdminUsers a ON b.BlockedBy = a.AdminID " & _
    "ORDER BY b.IsActive DESC, b.BlockedAt DESC")

' 统计
Dim activeBlocks : activeBlocks = CDbl("0" & GetScalar("SELECT COUNT(*) FROM IPBlacklist WHERE IsActive=1"))
Dim totalBlocks : totalBlocks = CDbl("0" & GetScalar("SELECT COUNT(*) FROM IPBlacklist"))
Dim todayBlocks : todayBlocks = CDbl("0" & GetScalar("SELECT COUNT(*) FROM IPBlacklist WHERE CAST(BlockedAt AS DATE)=CAST(GETDATE() AS DATE)"))
Dim expiredSoon : expiredSoon = CDbl("0" & GetScalar("SELECT COUNT(*) FROM IPBlacklist WHERE IsActive=1 AND ExpiresAt IS NOT NULL AND ExpiresAt <= DATEADD(DAY,3,GETDATE())"))
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>IP黑名单 - 站点技术管理</title>
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
        .msg { padding: 12px 20px; border-radius: 8px; margin-bottom: 20px; font-weight: 500; }
        .msg-success { background: rgba(76,175,80,0.15); color: #4CAF50; border: 1px solid rgba(76,175,80,0.3); }
        .msg-warning { background: rgba(255,152,0,0.15); color: #FFB74D; border: 1px solid rgba(255,152,0,0.3); }
        .msg-error { background: rgba(244,67,54,0.15); color: #EF9A9A; border: 1px solid rgba(244,67,54,0.3); }
        .card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); margin-bottom: 20px; overflow: hidden; }
        .card-header { padding: 15px 20px; border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 16px; display: flex; justify-content: space-between; align-items: center; }
        .card-body { padding: 20px; }
        table { width: 100%; border-collapse: collapse; }
        th { text-align: left; padding: 12px; background: rgba(0,188,212,0.06); border-bottom: 1px solid rgba(255,255,255,0.06); font-size: 13px; color: #999; }
        td { padding: 12px; border-bottom: 1px solid rgba(255,255,255,0.04); font-size: 14px; }
        tr:hover { background: rgba(255,255,255,0.03); }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
        .badge-active { background: rgba(244,67,54,0.2); color: #EF9A9A; }
        .badge-expired { background: rgba(158,158,158,0.2); color: #BDBDBD; }
        .badge-expiring { background: rgba(255,152,0,0.2); color: #FFB74D; }
        .form-row { display: flex; gap: 15px; flex-wrap: wrap; margin-bottom: 15px; align-items: flex-end; }
        .form-group { display: flex; flex-direction: column; gap: 5px; }
        .form-group label { font-size: 13px; color: #999; }
        .form-group input, .form-group select { padding: 10px 15px; border: 1px solid rgba(255,255,255,0.12); border-radius: 8px; background: #1a1a2e; color: #e0e0e0; font-size: 14px; min-width: 180px; }
        .form-group input:focus { border-color: #00bcd4; outline: none; }
        .ip-code { font-family: 'Consolas','Courier New',monospace; background: rgba(255,255,255,0.05); padding: 4px 10px; border-radius: 4px; font-size: 13px; color: #00bcd4; letter-spacing: 0.5px; }
        .empty { text-align: center; padding: 40px; color: #666; }
        .empty i { font-size: 40px; display: block; margin-bottom: 10px; opacity: 0.3; }
        .hit-counter { font-size: 12px; color: #888; }
        @media (max-width: 768px) { .main-content { margin-left: 0; } .stats-grid { grid-template-columns: 1fr 1fr; } .form-row { flex-direction: column; } }
    </style>
</head>
<body>
<!--#include file="includes/nav.asp"-->
<div class="main-content">
    <div class="page-header">
        <h2 class="page-title"><i class="fas fa-ban"></i> IP黑名单管理</h2>
        <div class="breadcrumb"><a href="index.asp">系统中心</a> / IP黑名单</div>
    </div>

    <% If msg <> "" Then %>
    <div class="msg msg-<%= msgType %>"><i class="fas fa-info-circle"></i> <%= msg %></div>
    <% End If %>

    <div class="stats-grid">
        <div class="stat-card"><div class="num"><%= activeBlocks %></div><div class="label">活跃封锁数</div></div>
        <div class="stat-card"><div class="num"><%= totalBlocks %></div><div class="label">历史封锁总数</div></div>
        <div class="stat-card"><div class="num"><%= todayBlocks %></div><div class="label">今日新增</div></div>
        <div class="stat-card warn"><div class="num"><%= expiredSoon %></div><div class="label">3日内到期</div></div>
    </div>

    <!-- 添加IP -->
    <div class="card">
        <div class="card-header"><i class="fas fa-plus-circle"></i> 封禁IP</div>
        <div class="card-body">
            <form method="post" class="form-row">
                <input type="hidden" name="action" value="add">
                <div class="form-group">
                    <label>IP地址 <span style="color:#f44336;">*</span></label>
                    <input type="text" name="ipAddress" placeholder="如 192.168.1.100" required>
                </div>
                <div class="form-group">
                    <label>封禁原因</label>
                    <input type="text" name="reason" placeholder="如：频繁登录失败、恶意扫描" style="min-width:250px;">
                </div>
                <div class="form-group">
                    <label>封锁天数（留空=永久）</label>
                    <select name="expiresDays">
                        <option value="0">永久封锁</option>
                        <option value="1">1天</option>
                        <option value="3">3天</option>
                        <option value="7">7天</option>
                        <option value="30">30天</option>
                        <option value="180">180天</option>
                    </select>
                </div>
                <div class="form-group" style="align-self:flex-end;">
                    <button type="submit" class="btn btn-danger"><i class="fas fa-ban"></i> 加入黑名单</button>
                </div>
            </form>
        </div>
    </div>

    <!-- 黑名单列表 -->
    <div class="card">
        <div class="card-header"><i class="fas fa-list"></i> 黑名单列表</div>
        <div class="card-body">
            <table>
                <thead>
                    <tr>
                        <th>IP地址</th>
                        <th>原因</th>
                        <th>状态</th>
                        <th>封禁时间</th>
                        <th>到期时间</th>
                        <th>命中次数</th>
                        <th>操作者</th>
                        <th>操作</th>
                    </tr>
                </thead>
                <tbody>
                    <% If Not rsBlacklist Is Nothing Then
                        Do While Not rsBlacklist.EOF
                            Dim isActive : isActive = CBool(rsBlacklist("IsActive") And 1)
                            Dim expiresAt : expiresAt = rsBlacklist("ExpiresAt")
                            Dim isExpiring : isExpiring = False
                            If isActive And Not IsNull(expiresAt) Then
                                If DateDiff("d", Now, expiresAt) <= 3 And DateDiff("d", Now, expiresAt) >= 0 Then isExpiring = True
                            End If
                            Dim statusBadge
                            If isActive Then
                                If isExpiring Then statusBadge = "badge-expiring" Else statusBadge = "badge-active"
                            Else
                                statusBadge = "badge-expired"
                            End If
                    %>
                    <tr>
                        <td><span class="ip-code"><%= rsBlacklist("IPAddress") %></span></td>
                        <td><%= rsBlacklist("Reason") %></td>
                        <td><span class="badge <%= statusBadge %>"><%= IIf(isActive,IIf(isExpiring,"即将到期","活跃"),"已解除") %></span></td>
                        <td style="font-size:13px;color:#999;"><%= rsBlacklist("BlockedAt") %></td>
                        <td style="font-size:13px;color:#999;">
                            <%= IIf(IsNull(expiresAt),"永久",IIf(DateDiff("d",Now,expiresAt)<0,"<span style=""color:#EF9A9A;"">已过期</span>",expiresAt)) %>
                        </td>
                        <td><span class="hit-counter"><%= CDbl("0" & rsBlacklist("HitCount")) & " 次" %></span></td>
                        <td style="font-size:13px;color:#999;"><%= rsBlacklist("BlockedByName") %></td>
                        <td>
                            <% If isActive Then %>
                            <form method="post" style="display:inline;" onsubmit="return confirm('确认解除对 <%= rsBlacklist("IPAddress") %> 的封锁？')">
                                <input type="hidden" name="action" value="unblock">
                                <input type="hidden" name="unblockID" value="<%= rsBlacklist("IPID") %>">
                                <button type="submit" class="btn btn-success btn-sm"><i class="fas fa-check"></i> 解除</button>
                            </form>
                            <% Else %>
                            <span style="color:#666;font-size:12px;">—</span>
                            <% End If %>
                        </td>
                    </tr>
                    <%
                            rsBlacklist.MoveNext
                        Loop
                        rsBlacklist.Close
                    Else
                    %>
                    <tr><td colspan="8" class="empty"><i class="fas fa-shield-alt" style="color:#4CAF50;"></i>暂无黑名单记录</td></tr>
                    <% End If %>
                </tbody>
            </table>
        </div>
    </div>

    <!-- 说明 -->
    <div class="card">
        <div class="card-header"><i class="fas fa-question-circle"></i> 使用说明</div>
        <div class="card-body" style="font-size:13px;color:#999;line-height:1.8;">
            <p><i class="fas fa-check-circle" style="color:#4CAF50;"></i> 黑名单IP将被禁止访问管理后台所有页面</p>
            <p><i class="fas fa-check-circle" style="color:#00bcd4;"></i> 建议配合登录监控功能，自动封禁恶意IP</p>
            <p><i class="fas fa-check-circle" style="color:#FF9800;"></i> 设置到期时间可在到期后自动恢复访问权限</p>
            <p><i class="fas fa-info-circle" style="color:#888;"></i> 命中次数记录该IP尝试访问被拦截的次数</p>
        </div>
    </div>
</div>
</body>
</html>
<%
Call CloseConnection()
%>
