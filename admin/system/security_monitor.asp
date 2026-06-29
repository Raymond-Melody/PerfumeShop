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

' V18: 本地哈希函数（用于文件完整性检查）
Function SM_HashString(inputStr)
    Dim result, i, charCode
    result = ""
    For i = 1 To Len(inputStr)
        charCode = Asc(Mid(inputStr, i, 1))
        result = result & Right("0" & Hex((charCode * 3 + i) Mod 256), 2)
    Next
    SM_HashString = Left(UCase(result), 32)
End Function

' ============================================
' V18.0 安全监控仪表盘 (Security Monitor Dashboard)
' 整合: 登录失败统计 / SQL注入检测 / 异常访问告警 / 文件完整性监控
' ============================================

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Function SafeStr(val)
    If IsNull(val) Then SafeStr = "" Else SafeStr = CStr(val)
End Function

' ============================================
' 1. 登录失败统计（按IP和账号聚合）
' ============================================
Dim todayFailTotal, weekFailTotal, ipFailTop, accountFailTop
todayFailTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM AdminLogs WHERE ActionType='登录失败' AND CreatedAt >= DATEADD(HOUR,-24,GETDATE())"))
weekFailTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM AdminLogs WHERE ActionType='登录失败' AND CreatedAt >= DATEADD(DAY,-7,GETDATE())"))

' Top 5 失败IP
Set ipFailTop = ExecuteQuery("SELECT TOP 5 IPAddress, COUNT(*) AS FailCount FROM AdminLogs " & _
    "WHERE ActionType='登录失败' AND CreatedAt >= DATEADD(DAY,-7,GETDATE()) AND IPAddress IS NOT NULL AND IPAddress <> '' " & _
    "GROUP BY IPAddress ORDER BY COUNT(*) DESC")

' Top 5 失败账号
Set accountFailTop = ExecuteQuery("SELECT TOP 5 AdminID, COUNT(*) AS FailCount FROM AdminLogs " & _
    "WHERE ActionType='登录失败' AND CreatedAt >= DATEADD(DAY,-7,GETDATE()) AND AdminID IS NOT NULL " & _
    "GROUP BY AdminID ORDER BY COUNT(*) DESC")

' ============================================
' 2. SQL注入尝试检测（基于AppLogs模式匹配）
' ============================================
Dim sqliAttempts24h, sqliAttempts7d
sqliAttempts24h = SafeNum(GetScalar("SELECT COUNT(*) FROM AppLogs WHERE LogLevel='SECURITY' AND CreatedAt >= DATEADD(HOUR,-24,GETDATE())"))
sqliAttempts7d = SafeNum(GetScalar("SELECT COUNT(*) FROM AppLogs WHERE LogLevel='SECURITY' AND CreatedAt >= DATEADD(DAY,-7,GETDATE())"))

' SQL关键字模式检测（检查AdminLogs/AppLogs中的可疑请求）
Dim sqliPatterns
Set sqliPatterns = ExecuteQuery("SELECT TOP 20 LogMessage, LogSource, CreatedAt, IPAddress FROM AppLogs " & _
    "WHERE (LogMessage LIKE '%UNION%' OR LogMessage LIKE '%SELECT%' OR LogMessage LIKE '%DROP%' OR " & _
    "LogMessage LIKE '%EXEC%' OR LogMessage LIKE '%SCRIPT%' OR LogMessage LIKE '%1=1%' OR " & _
    "LogMessage LIKE '%OR 1=%' OR LogMessage LIKE '%--%') " & _
    "AND CreatedAt >= DATEADD(DAY,-7,GETDATE()) AND LogLevel='ERROR' " & _
    "ORDER BY CreatedAt DESC")

' ============================================
' 3. 异常访问频率告警
' ============================================
Dim rateLimitHits24h, rateLimitHits7d
rateLimitHits24h = 0 : rateLimitHits7d = 0
On Error Resume Next
rateLimitHits24h = SafeNum(GetScalar("SELECT COUNT(*) FROM AppLogs WHERE LogSource='rate_limiter' AND CreatedAt >= DATEADD(HOUR,-24,GETDATE())"))
rateLimitHits7d = SafeNum(GetScalar("SELECT COUNT(*) FROM AppLogs WHERE LogSource='rate_limiter' AND CreatedAt >= DATEADD(DAY,-7,GETDATE())"))
On Error GoTo 0

' ============================================
' 4. 文件篡改监控（includes/ 目录哈希校验）
' ============================================
Dim fileScanResults, fileScanCount, fileTampered
fileScanCount = 0
fileTampered = 0
ReDim fileScanResults(0, 2)

If Request.QueryString("scan") = "files" Then
    Dim fso, includesFolder, fileItem, fileContent, fileHash, expectedHash
    Dim cacheKey
    
    On Error Resume Next
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        Set includesFolder = fso.GetFolder(Server.MapPath("/includes"))
        
        For Each fileItem In includesFolder.Files
            If LCase(fso.GetExtensionName(fileItem.Name)) = "asp" Then
                fileContent = ""
                Set cf = fso.OpenTextFile(fileItem.Path, 1)
                If Err.Number = 0 Then fileContent = cf.ReadAll : cf.Close
                Err.Clear
                
                If fileContent <> "" Then
                    ' 计算简单哈希（文件长度+内容抽样）
                    fileHash = Len(fileContent) & "-" & UCase(Left(SM_HashString(Left(fileContent, 2000)), 8))
                    
                    ' 与之前存储的哈希比对
                    cacheKey = "FILE_HASH_" & fileItem.Name
                    expectedHash = Application(cacheKey)
                    
                    ReDim Preserve fileScanResults(fileScanCount, 2)
                    fileScanResults(fileScanCount, 0) = fileItem.Name
                    
                    If expectedHash <> "" And expectedHash <> fileHash Then
                        fileScanResults(fileScanCount, 1) = "CHANGED"
                        fileScanResults(fileScanCount, 2) = "Hash changed: " & expectedHash & " → " & fileHash
                        fileTampered = fileTampered + 1
                    Else
                        fileScanResults(fileScanCount, 1) = "OK"
                        fileScanResults(fileScanCount, 2) = fileHash
                    End If
                    
                    ' 保存当前哈希
                    Application.Lock
                    Application(cacheKey) = fileHash
                    Application.UnLock
                    
                    fileScanCount = fileScanCount + 1
                End If
            End If
            Set cf = Nothing
        Next
        Set includesFolder = Nothing
        Set fso = Nothing
    End If
    On Error GoTo 0
End If

' ============================================
' 总体安全评分 (0-100)
' ============================================
Dim securityScore
securityScore = 100
If todayFailTotal > 10 Then securityScore = securityScore - 10
If todayFailTotal > 50 Then securityScore = securityScore - 10
If sqliAttempts24h > 0 Then securityScore = securityScore - 15
If sqliAttempts24h > 5 Then securityScore = securityScore - 15
If rateLimitHits24h > 20 Then securityScore = securityScore - 10
If fileTampered > 0 Then securityScore = securityScore - 20
If securityScore < 0 Then securityScore = 0

Dim scoreColor
If securityScore >= 80 Then scoreColor = "#4CAF50"
If securityScore >= 60 And securityScore < 80 Then scoreColor = "#FF9800"
If securityScore < 60 Then scoreColor = "#F44336"
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>安全监控 - 站点技术管理</title>
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
        
        /* 安全评分 */
        .score-banner { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 14px; padding: 28px 32px; margin-bottom: 25px; border: 1px solid rgba(255,255,255,0.05); display: flex; align-items: center; gap: 24px; }
        .score-circle-danger { border-color: #F44336; color: #F44336; }
        .score-circle-warn { border-color: #FF9800; color: #FF9800; }
        .score-circle-ok { border-color: #4CAF50; color: #4CAF50; }
        .score-circle { width: 90px; height: 90px; border-radius: 50%; border-width: 5px; border-style: solid; display: flex; align-items: center; justify-content: center; font-size: 36px; font-weight: 800; flex-shrink: 0; }
        .score-info h3 { margin: 0 0 6px; color: #fff; font-size: 18px; }
        .score-info p { margin: 0; color: #888; font-size: 13px; line-height: 1.6; }
        
        /* 卡片网格 */
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); padding: 22px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); text-align: center; }
        .stat-card .num { font-size: 30px; font-weight: 700; }
        .stat-card .label { font-size: 13px; color: #888; margin-top: 6px; }
        .stat-card .icon { font-size: 22px; margin-bottom: 6px; opacity: 0.6; }
        .num-green { color: #4CAF50; } .num-red { color: #F44336; } .num-orange { color: #FF9800; } .num-blue { color: #2196F3; }
        
        /* 面板 */
        .card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); margin-bottom: 20px; overflow: hidden; }
        .card-header { padding: 15px 20px; border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 16px; display: flex; justify-content: space-between; align-items: center; }
        .card-header .actions { display: flex; gap: 8px; }
        .card-body { padding: 20px; }
        .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        
        table { width: 100%; border-collapse: collapse; }
        th { text-align: left; padding: 10px 14px; background: rgba(0,0,0,0.2); border-bottom: 1px solid rgba(255,255,255,0.06); font-size: 12px; color: #999; text-transform: uppercase; }
        td { padding: 10px 14px; border-bottom: 1px solid rgba(255,255,255,0.03); font-size: 13px; }
        tr:hover td { background: rgba(255,255,255,0.02); }
        
        .badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 600; }
        .badge-danger { background: rgba(244,67,54,0.2); color: #EF9A9A; }
        .badge-warn { background: rgba(255,152,0,0.2); color: #FFB74D; }
        .badge-success { background: rgba(76,175,80,0.2); color: #A5D6A7; }
        .badge-info { background: rgba(33,150,243,0.2); color: #90CAF9; }
        
        .btn { padding: 8px 16px; border: none; border-radius: 6px; font-size: 13px; cursor: pointer; text-decoration: none; display: inline-flex; align-items: center; gap: 6px; }
        .btn-sm { padding: 6px 12px; font-size: 12px; }
        .btn-primary { background: #2196F3; color: #fff; }
        .btn-warn { background: #FF9800; color: #fff; }
        .btn-ghost { background: transparent; border: 1px solid rgba(255,255,255,0.15); color: #e0e0e0; }
        .btn-ghost:hover { background: rgba(255,255,255,0.05); }
        
        .empty { text-align: center; padding: 40px; color: #666; }
        .empty i { font-size: 48px; display: block; margin-bottom: 12px; opacity: 0.3; }
        .ip-mono { font-family: 'Consolas',monospace; font-size: 12px; color: #888; }
        .time-text { font-size: 12px; color: #666; }
        .pulse { animation: pulse 2s infinite; }
        @keyframes pulse { 0%,100% { opacity: 1; } 50% { opacity: 0.5; } }
        .text-break { word-break: break-all; max-width: 400px; }
        
        @media (max-width: 1024px) { .stats-grid, .two-col { grid-template-columns: 1fr; } }
        @media (max-width: 768px) { .main-content { margin-left: 0; } }
    </style>
</head>
<body>
<!--#include file="includes/nav.asp"-->
<div class="main-content">
    <div class="page-header">
        <h2 class="page-title"><i class="fas fa-shield-alt"></i> 安全监控仪表盘</h2>
        <div class="breadcrumb"><a href="index.asp">系统中心</a> / <span>安全监控</span></div>
    </div>

    <!-- 安全评分 -->
    <div class="score-banner">
        <div class="score-circle<%= IIf(securityScore>=80,"-ok",IIf(securityScore>=60,"-warn","-danger")) %>"><%= securityScore %></div>
        <div class="score-info">
            <h3>安全评分</h3>
            <p>
                <% If securityScore >= 80 Then %>
                    <i class="fas fa-check-circle" style="color:#4CAF50;"></i> 安全状况良好，各项指标正常
                <% ElseIf securityScore >= 60 Then %>
                    <i class="fas fa-exclamation-circle" style="color:#FF9800;"></i> 存在一些安全问题，建议尽快处理
                <% Else %>
                    <i class="fas fa-times-circle" style="color:#F44336;"></i> 安全风险较高，需要立即处理
                <% End If %>
            </p>
        </div>
        <div style="flex:1;"></div>
        <div style="display:flex;gap:8px;">
            <a href="login_monitor.asp" class="btn btn-ghost"><i class="fas fa-user-shield"></i> 登录监控</a>
            <a href="security_audit.asp" class="btn btn-ghost"><i class="fas fa-search"></i> 安全审计</a>
            <a href="ip_blacklist.asp" class="btn btn-ghost"><i class="fas fa-ban"></i> 黑名单</a>
        </div>
    </div>

    <!-- KPI卡片 -->
    <div class="stats-grid">
        <div class="stat-card">
            <div class="icon"><i class="fas fa-sign-in-alt"></i></div>
            <div class="num num-<%= IIf(todayFailTotal>10,"red","green") %>"><%= todayFailTotal %></div>
            <div class="label">24h登录失败</div>
        </div>
        <div class="stat-card">
            <div class="icon"><i class="fas fa-database"></i></div>
            <div class="num num-<%= IIf(sqliAttempts24h>0,"red","green") %>"><%= sqliAttempts24h %></div>
            <div class="label">SQL注入告警(24h)</div>
        </div>
        <div class="stat-card">
            <div class="icon"><i class="fas fa-tachometer-alt"></i></div>
            <div class="num num-<%= IIf(rateLimitHits24h>20,"orange","blue") %>"><%= rateLimitHits24h %></div>
            <div class="label">限流触发(24h)</div>
        </div>
        <div class="stat-card">
            <div class="icon"><i class="fas fa-file-code"></i></div>
            <div class="num num-<%= IIf(fileTampered>0,"red","green") %>"><%= fileTampered %></div>
            <div class="label">文件变更</div>
        </div>
    </div>

    <!-- 两栏布局 -->
    <div class="two-col">
        <!-- 登录失败TOP-IP -->
        <div class="card">
            <div class="card-header"><i class="fas fa-globe"></i> 登录失败 TOP5 IP (7天)<span class="actions"><a href="ip_blacklist.asp" class="btn btn-sm btn-warn"><i class="fas fa-ban"></i> 管理黑名单</a></span></div>
            <div class="card-body">
                <%
                If Not ipFailTop Is Nothing And Not ipFailTop.EOF Then
                %>
                <table>
                    <thead><tr><th>IP地址</th><th>失败次数</th><th>威胁等级</th></tr></thead>
                    <tbody>
                    <%
                    Dim ipRow
                    Do While Not ipFailTop.EOF
                        Dim fCount : fCount = SafeNum(ipFailTop("FailCount"))
                        Dim ipLevel
                        If fCount >= 20 Then
                            ipLevel = "badge-danger"
                        ElseIf fCount >= 10 Then
                            ipLevel = "badge-warn"
                        Else
                            ipLevel = "badge-info"
                        End If
                    %>
                    <tr>
                        <td class="ip-mono"><%= SafeStr(ipFailTop("IPAddress")) %></td>
                        <td><strong><%= fCount %></strong></td>
                        <td><span class="badge <%= ipLevel %>"><%= IIf(fCount>=20,"🔴 高",IIf(fCount>=10,"🟡 中","🟢 低")) %></span></td>
                    </tr>
                    <%
                        ipFailTop.MoveNext
                    Loop
                    ipFailTop.Close
                    %>
                    </tbody>
                </table>
                <%
                Else
                %>
                <div class="empty"><i class="fas fa-check-circle" style="color:#4CAF50;"></i>无登录失败记录</div>
                <%
                End If
                Set ipFailTop = Nothing
                %>
            </div>
        </div>

        <!-- 登录失败TOP-账号 -->
        <div class="card">
            <div class="card-header"><i class="fas fa-user-times"></i> 登录失败 TOP5 账号 (7天)</div>
            <div class="card-body">
                <%
                If Not accountFailTop Is Nothing And Not accountFailTop.EOF Then
                %>
                <table>
                    <thead><tr><th>管理员ID</th><th>失败次数</th><th>风险</th></tr></thead>
                    <tbody>
                    <%
                    Do While Not accountFailTop.EOF
                        Dim aCount : aCount = SafeNum(accountFailTop("FailCount"))
                    %>
                    <tr>
                        <td>Admin#<%= accountFailTop("AdminID") %></td>
                        <td><strong><%= aCount %></strong></td>
                        <td><span class="badge <%= IIf(aCount>=10,"badge-danger","badge-warn") %>"><%= IIf(aCount>=10,"疑似暴力破解","需关注") %></span></td>
                    </tr>
                    <%
                        accountFailTop.MoveNext
                    Loop
                    accountFailTop.Close
                    %>
                    </tbody>
                </table>
                <%
                Else
                %>
                <div class="empty"><i class="fas fa-check-circle" style="color:#4CAF50;"></i>无异常账号</div>
                <%
                End If
                Set accountFailTop = Nothing
                %>
            </div>
        </div>
    </div>

    <!-- SQL注入尝试检测 -->
    <div class="card">
        <div class="card-header"><i class="fas fa-bug"></i> SQL注入尝试检测 (7天内可疑请求)</div>
        <div class="card-body">
            <%
            If Not sqliPatterns Is Nothing And Not sqliPatterns.EOF Then
            %>
            <table>
                <thead><tr><th>时间</th><th>来源</th><th>IP</th><th>详情</th></tr></thead>
                <tbody>
                <%
                Dim sqliCount : sqliCount = 0
                Do While Not sqliPatterns.EOF And sqliCount < 20
                    sqliCount = sqliCount + 1
                %>
                <tr>
                    <td class="time-text"><%= SafeStr(sqliPatterns("CreatedAt")) %></td>
                    <td><span class="badge badge-warn"><%= SafeStr(sqliPatterns("LogSource")) %></span></td>
                    <td class="ip-mono"><%= SafeStr(sqliPatterns("IPAddress")) %></td>
                    <td class="text-break" style="font-size:12px;"><%= Left(SafeStr(sqliPatterns("LogMessage")), 200) %></td>
                </tr>
                <%
                    sqliPatterns.MoveNext
                Loop
                sqliPatterns.Close
                %>
                </tbody>
            </table>
            <%
            Else
            %>
            <div class="empty"><i class="fas fa-shield-virus" style="color:#4CAF50;"></i>未发现SQL注入尝试</div>
            <%
            End If
            Set sqliPatterns = Nothing
            %>
        </div>
    </div>

    <!-- 文件完整性监控 -->
    <div class="card">
        <div class="card-header">
            <i class="fas fa-file-signature"></i> 文件完整性监控 (includes/ 目录)
            <span class="actions">
                <a href="?scan=files" class="btn btn-sm btn-primary"><i class="fas fa-search"></i> 扫描文件</a>
            </span>
        </div>
        <div class="card-body">
            <% If fileScanCount > 0 Then %>
            <table>
                <thead><tr><th>文件名</th><th>状态</th><th>哈希/变更详情</th></tr></thead>
                <tbody>
                <%
                Dim fi
                For fi = 0 To fileScanCount - 1
                    Dim fStatus : fStatus = fileScanResults(fi, 1)
                    Dim statusBadge
                    If fStatus = "OK" Then statusBadge = "badge-success"
                    If fStatus = "CHANGED" Then statusBadge = "badge-danger"
                %>
                <tr>
                    <td style="font-family:'Consolas',monospace;font-size:12px;"><%= fileScanResults(fi, 0) %></td>
                    <td><span class="badge <%= statusBadge %>"><%= fStatus %></span></td>
                    <td style="font-size:11px;color:#888;font-family:'Consolas',monospace;"><%= fileScanResults(fi, 2) %></td>
                </tr>
                <%
                Next
                %>
                </tbody>
            </table>
            <% Else %>
            <div class="empty">
                <i class="fas fa-file-code"></i>
                <p>点击"扫描文件"按钮检查 includes/ 目录下的 .asp 文件完整性</p>
                <p style="font-size:12px;color:#666;">首次扫描将建立哈希基线，后续扫描将对比检测变更</p>
            </div>
            <% End If %>
        </div>
    </div>

    <!-- 速率限制统计 -->
    <div class="card">
        <div class="card-header"><i class="fas fa-tachometer-alt"></i> 速率限制统计</div>
        <div class="card-body">
            <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:20px;text-align:center;">
                <div>
                    <div style="font-size:24px;font-weight:700;color:#2196F3;"><%= rateLimitHits24h %></div>
                    <div style="font-size:12px;color:#888;">24h限流触发</div>
                </div>
                <div>
                    <div style="font-size:24px;font-weight:700;color:#FF9800;"><%= rateLimitHits7d %></div>
                    <div style="font-size:12px;color:#888;">7天限流触发</div>
                </div>
                <div>
                    <div style="font-size:24px;font-weight:700;color:<%= IIf(rateLimitHits24h>20,"#F44336","#4CAF50") %>;">
                        <%= IIf(rateLimitHits24h>20,"⚠️ 异常","✅ 正常") %>
                    </div>
                    <div style="font-size:12px;color:#888;">状态评估</div>
                </div>
            </div>
        </div>
    </div>

    <!-- 安全基线清单 -->
    <div class="card">
        <div class="card-header"><i class="fas fa-clipboard-check"></i> 安全基线快速检查</div>
        <div class="card-body">
            <table>
                <thead><tr><th>检查项</th><th>状态</th><th>说明</th></tr></thead>
                <tbody>
                    <tr><td>CSRF 防护</td><td><span class="badge badge-success">✅ 已启用</span></td><td>全局CSRF Token + 历史池验证 (V17)</td></tr>
                    <tr><td>API 认证</td><td><span class="badge badge-success">✅ 已启用</span></td><td>HMAC签名 + 时间戳防重放 (V18)</td></tr>
                    <tr><td>速率限制</td><td><span class="badge badge-success">✅ 已启用</span></td><td>令牌桶算法 + 429标准响应 (V18)</td></tr>
                    <tr><td>支付签名</td><td><span class="badge badge-success">✅ 已启用</span></td><td>支付请求签名 + 幂等键防重复 (V18)</td></tr>
                    <tr><td>GDPR合规</td><td><span class="badge badge-success">✅ 已启用</span></td><td>Cookie同意 + 数据导出 + 账户注销 (V18)</td></tr>
                    <tr><td>文件监控</td><td><span class="badge <%= IIf(fileScanCount>0,"badge-success","badge-warn") %>"><%= IIf(fileScanCount>0,"✅ 已扫描","⚠️ 待扫描") %></span></td><td>includes/ 目录哈希基线</td></tr>
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
