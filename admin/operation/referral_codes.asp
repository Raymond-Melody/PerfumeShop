<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<!--#include file="../../includes/member_utils.asp"-->
<%
Call OpenConnection()

Dim actionMsg, actionError, activeTab
actionMsg = ""
actionError = ""
activeTab = Request.QueryString("tab")
If activeTab = "" Then activeTab = "generate"

' V14: 处理POST请求
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not ValidateCSRFToken() Then
        actionError = "安全验证失败，请刷新页面重试"
    Else
        Dim postAction
        postAction = Request.Form("action")
        
        If postAction = "generate_code" Then
            ' 生成管理员推荐码
            Dim maxUses, daysValid
            maxUses = Request.Form("max_uses")
            daysValid = Request.Form("days_valid")
            
            If maxUses = "" Or Not IsNumeric(maxUses) Then maxUses = 20
            If daysValid = "" Or Not IsNumeric(daysValid) Then daysValid = 5
            
            maxUses = CInt(maxUses)
            daysValid = CInt(daysValid)
            
            If maxUses < 1 Then maxUses = 1
            If maxUses > 1000 Then maxUses = 1000
            If daysValid < 1 Then daysValid = 1
            If daysValid > 30 Then daysValid = 30
            
            ' 使用管理员ID作为推荐人ID（负数表示管理员生成）
            Dim adminToken, adminStoreResult
            adminToken = MU_GenerateReferralToken(Session("AdminID"), daysValid, maxUses, "admin")
            If adminToken <> "" Then
                If MU_StoreReferralToken(adminToken, "admin") Then
                    actionMsg = "推荐码已生成！有效期" & daysValid & "天，可使用" & maxUses & "次"
                    activeTab = "list"
                Else
                    actionError = "存储推荐码失败: " & Session("LastDBError")
                End If
            Else
                actionError = "生成推荐码失败"
            End If
            
        ElseIf postAction = "deactivate" Then
            ' 失效推荐码
            Dim deactivateId
            deactivateId = Request.Form("token_id")
            If deactivateId <> "" And IsNumeric(deactivateId) Then
                If ExecuteNonQuery("UPDATE ReferralTokens SET IsActive = 0 WHERE TokenID = " & CLng(deactivateId)) Then
                    actionMsg = "推荐码已失效"
                    ' 记录操作日志
                    Call LogAdminAction("失效推荐码", "operation/referral_codes", "ReferralTokens", deactivateId, "管理员手动失效推荐码")
                Else
                    actionError = "操作失败"
                End If
            End If
        End If
    End If
End If

' 确保CSRF令牌存在
Call EnsureCSRFToken()

' 获取推荐码列表
Dim rsCodes, codeSQL
If activeTab = "list" Or activeTab = "" Then
    codeSQL = "SELECT rt.*, u.Username AS AdminName FROM ReferralTokens rt " & _
              "LEFT JOIN AdminUsers u ON rt.ReferrerUserID = u.AdminID AND rt.ReferrerType = 'admin' " & _
              "WHERE rt.ReferrerType = 'admin' ORDER BY rt.CreatedAt DESC"
Else
    codeSQL = "SELECT rt.*, u.Username AS AdminName FROM ReferralTokens rt " & _
              "LEFT JOIN AdminUsers u ON rt.ReferrerUserID = u.AdminID AND rt.ReferrerType = 'admin' " & _
              "WHERE rt.ReferrerType = 'admin' AND rt.IsActive = 1 AND rt.ExpiresAt > GETDATE() AND rt.UsedCount < rt.MaxUses ORDER BY rt.CreatedAt DESC"
End If
Set rsCodes = ExecuteQuery(codeSQL)

' 获取推荐关系概览统计
Dim totalRelations, todayInvitees
totalRelations = GetScalar("SELECT COUNT(*) FROM ReferralRelations WHERE Depth = 1")
todayInvitees = GetScalar("SELECT COUNT(*) FROM ReferralRelations WHERE Depth = 1 AND CAST(CreatedAt AS DATE) = CAST(GETDATE() AS DATE)")
If IsNull(totalRelations) Then totalRelations = 0
If IsNull(todayInvitees) Then todayInvitees = 0
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>推荐码管理 - 运营管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; }
        .main-content { padding: 25px; }
        .page-header {
            display: flex; justify-content: space-between; align-items: center;
            margin-bottom: 25px; padding-bottom: 15px;
            border-bottom: 1px solid rgba(255,255,255,0.08);
        }
        .page-title { font-size: 24px; color: #fff; margin: 0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #00bcd4; }
        .breadcrumb { font-size: 13px; color: #888; }
        .breadcrumb a { color: #00bcd4; text-decoration: none; }
        
        .tabs { display: flex; gap: 5px; margin-bottom: 25px; border-bottom: 2px solid rgba(255,255,255,0.08); }
        .tab { padding: 12px 25px; color: #888; text-decoration: none; font-size: 15px; border-bottom: 2px solid transparent; margin-bottom: -2px; transition: all 0.3s; }
        .tab:hover { color: #fff; }
        .tab.active { color: #00bcd4; border-bottom-color: #00bcd4; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            padding: 20px; border-radius: 10px; text-align: center;
            border: 1px solid rgba(255,255,255,0.05);
        }
        .stat-value { font-size: 28px; font-weight: bold; color: #fff; }
        .stat-label { color: #aaa; font-size: 13px; margin-top: 5px; }
        
        .card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px; padding: 25px; margin-bottom: 25px;
            border: 1px solid rgba(255,255,255,0.05);
        }
        .card h3 { color: #fff; font-size: 18px; margin-bottom: 20px; display: flex; align-items: center; gap: 10px; }
        .card h3 i { color: #00bcd4; }
        
        .form-group { margin-bottom: 15px; }
        .form-group label { display: block; color: #ccc; margin-bottom: 5px; font-size: 14px; }
        .form-group input, .form-group select {
            width: 100%; padding: 10px; border-radius: 6px;
            border: 1px solid rgba(255,255,255,0.1); background: rgba(0,0,0,0.3); color: #fff;
            font-size: 14px;
        }
        .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }
        
        .btn { display: inline-block; padding: 10px 20px; border-radius: 6px; border: none; cursor: pointer; font-size: 14px; text-decoration: none; }
        .btn-primary { background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); color: #fff; }
        .btn-primary:hover { opacity: 0.9; }
        .btn-danger { background: #dc3545; color: #fff; }
        .btn-danger:hover { background: #c82333; }
        .btn-sm { padding: 5px 12px; font-size: 12px; }
        
        .alert { padding: 12px 15px; border-radius: 8px; margin-bottom: 15px; }
        .alert-success { background: rgba(76,175,80,0.2); color: #4caf50; border: 1px solid rgba(76,175,80,0.3); }
        .alert-error { background: rgba(244,67,54,0.2); color: #f44336; border: 1px solid rgba(244,67,54,0.3); }
        
        table { width: 100%; border-collapse: collapse; }
        th { text-align: left; padding: 10px; background: rgba(0,0,0,0.2); color: #888; font-size: 12px; text-transform: uppercase; }
        td { padding: 10px; border-bottom: 1px solid rgba(255,255,255,0.05); }
        tr:hover td { background: rgba(255,255,255,0.03); }
        
        .badge { display: inline-block; padding: 3px 10px; border-radius: 10px; font-size: 11px; font-weight: 500; }
        .badge-active { background: rgba(76,175,80,0.2); color: #4caf50; }
        .badge-expired { background: rgba(244,67,54,0.2); color: #f44336; }
        .badge-exhausted { background: rgba(255,152,0,0.2); color: #ff9800; }
        .badge-inactive { background: rgba(158,158,158,0.2); color: #9e9e9e; }
        
        .referral-url {
            background: rgba(0,0,0,0.3); padding: 8px 12px; border-radius: 4px;
            font-family: monospace; font-size: 12px; word-break: break-all;
            color: #00bcd4; margin: 10px 0;
        }
        
        @media (max-width: 768px) {
            .stats-grid { grid-template-columns: repeat(2, 1fr); }
            .form-row { grid-template-columns: 1fr; }
        }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-user-friends"></i> 推荐码管理</h2>
            <div class="breadcrumb">
                <a href="index.asp">运营中心</a> / <span>推荐码管理</span>
            </div>
        </div>
        
        <!-- 统计概览 -->
        <div class="stats-grid">
            <% 
            Dim totalCodes, activeCodes
            totalCodes = GetScalar("SELECT COUNT(*) FROM ReferralTokens WHERE ReferrerType = 'admin'")
            activeCodes = GetScalar("SELECT COUNT(*) FROM ReferralTokens WHERE ReferrerType = 'admin' AND IsActive = 1 AND ExpiresAt > GETDATE() AND UsedCount < MaxUses")
            If IsNull(totalCodes) Then totalCodes = 0
            If IsNull(activeCodes) Then activeCodes = 0
            %>
            <div class="stat-card">
                <div class="stat-value"><%= totalCodes %></div>
                <div class="stat-label">累计推荐码</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%= activeCodes %></div>
                <div class="stat-label">有效推荐码</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%= totalRelations %></div>
                <div class="stat-label">总推荐关系</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%= todayInvitees %></div>
                <div class="stat-label">今日新增会员</div>
            </div>
        </div>
        
        <% If actionMsg <> "" Then %>
        <div class="alert alert-success"><i class="fas fa-check-circle"></i> <%= actionMsg %></div>
        <% End If %>
        <% If actionError <> "" Then %>
        <div class="alert alert-error"><i class="fas fa-exclamation-circle"></i> <%= actionError %></div>
        <% End If %>
        
        <!-- Tab 切换 -->
        <div class="tabs">
            <a href="?tab=generate" class="tab <%= IIf(activeTab = "generate", "active", "") %>">
                <i class="fas fa-plus-circle"></i> 生成推荐码
            </a>
            <a href="?tab=list" class="tab <%= IIf(activeTab = "list", "active", "") %>">
                <i class="fas fa-list"></i> 推荐码列表
            </a>
            <a href="?tab=relations" class="tab <%= IIf(activeTab = "relations", "active", "") %>">
                <i class="fas fa-sitemap"></i> 推荐关系链
            </a>
        </div>
        
        <% If activeTab = "generate" Then %>
        <!-- 生成推荐码 -->
        <div class="card">
            <h3><i class="fas fa-qrcode"></i> 生成管理员推荐码</h3>
            <p style="color:#888;margin-bottom:20px;">管理员生成的推荐码具有更高的使用次数和自定义有效期</p>
            
            <form method="post">
                <%= GetCSRFTokenField() %>
                <input type="hidden" name="action" value="generate_code">
                
                <div class="form-row">
                    <div class="form-group">
                        <label>有效期（天）</label>
                        <select name="days_valid">
                            <option value="1">1天</option>
                            <option value="3">3天</option>
                            <option value="5" selected>5天</option>
                            <option value="7">7天</option>
                            <option value="14">14天</option>
                            <option value="30">30天</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>最大使用次数</label>
                        <input type="number" name="max_uses" value="20" min="1" max="1000">
                    </div>
                </div>
                
                <button type="submit" class="btn btn-primary">
                    <i class="fas fa-magic"></i> 生成推荐码
                </button>
            </form>
            
            <% If actionMsg <> "" Then %>
            <div style="margin-top:20px;background:rgba(76,175,80,0.1);border:1px solid rgba(76,175,80,0.3);border-radius:10px;padding:20px;">
                <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;">
                    <i class="fas fa-check-circle" style="color:#4caf50;font-size:18px;"></i>
                    <span style="color:#4caf50;font-weight:bold;"><%= actionMsg %></span>
                </div>
                <% 
                ' 从最近生成的Token中获取URL
                Dim rsLatest
                Set rsLatest = ExecuteQuery("SELECT TOP 1 OriginalToken FROM ReferralTokens WHERE ReferrerType = 'admin' ORDER BY CreatedAt DESC")
                If Not rsLatest Is Nothing And Not rsLatest.EOF Then
                    Dim latestToken
                    latestToken = rsLatest("OriginalToken")
                    Dim fullRegUrl, regProtocol
                    If Request.ServerVariables("HTTPS") = "on" Then
                        regProtocol = "https://"
                    Else
                        regProtocol = "http://"
                    End If
                    fullRegUrl = regProtocol & Request.ServerVariables("HTTP_HOST") & "/user/register.asp?token=" & Server.URLEncode(latestToken)
                %>
                <div style="margin-bottom:10px;">
                    <label style="color:#aaa;font-size:12px;display:block;margin-bottom:5px;">注册链接（点击复制）</label>
                    <div style="display:flex;gap:8px;align-items:center;">
                        <input type="text" readonly value="<%= fullRegUrl %>" id="generatedUrl" style="flex:1;padding:10px 14px;border-radius:6px;border:1px solid rgba(255,255,255,0.1);background:rgba(0,0,0,0.4);color:#00bcd4;font-family:monospace;font-size:13px;" onclick="this.select()">
                        <button onclick="copyGenUrl()" class="btn btn-sm" style="background:#4caf50;color:#fff;padding:10px 18px;white-space:nowrap;">
                            <i class="fas fa-copy"></i> 复制链接
                        </button>
                    </div>
                </div>
                <div style="font-size:12px;color:#888;">
                    <i class="fas fa-info-circle"></i> 将此链接发送给新用户即可完成推荐注册
                </div>
                <%
                    rsLatest.Close
                End If
                Set rsLatest = Nothing
                %>
            </div>
            <script>
            function copyGenUrl() {
                var input = document.getElementById('generatedUrl');
                input.select();
                input.setSelectionRange(0, 99999);
                if (navigator.clipboard && navigator.clipboard.writeText) {
                    navigator.clipboard.writeText(input.value).then(function() {
                        var btn = event.target.closest('button');
                        btn.innerHTML = '<i class="fas fa-check"></i> 已复制';
                        setTimeout(function() { btn.innerHTML = '<i class="fas fa-copy"></i> 复制链接'; }, 2000);
                    });
                } else {
                    document.execCommand('copy');
                }
            }
            </script>
            <% End If %>
        </div>
        
        <% ElseIf activeTab = "list" Then %>
        <!-- 推荐码列表 -->
        <div class="card">
            <h3><i class="fas fa-list-ul"></i> 推荐码列表</h3>
            
            <table>
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>生成者</th>
                        <th>使用次数/上限</th>
                        <th>有效期至</th>
                        <th>状态</th>
                        <th>生成时间</th>
                        <th>操作</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    If Not rsCodes Is Nothing Then
                        Do While Not rsCodes.EOF
                            Dim codeId, codeAdminName, codeUsed, codeMax, codeExpires, codeActive, codeCreated
                            codeId = rsCodes("TokenID")
                            codeAdminName = rsCodes("AdminName")
                            If IsNull(codeAdminName) Then codeAdminName = "管理员#" & rsCodes("ReferrerUserID")
                            codeUsed = rsCodes("UsedCount")
                            codeMax = rsCodes("MaxUses")
                            codeExpires = rsCodes("ExpiresAt")
                            codeActive = rsCodes("IsActive")
                            codeCreated = rsCodes("CreatedAt")
                            
                            ' Normalize DATETIME2(7) for VBScript CDate compatibility
                            ' Must strip fractional seconds entirely (VBScript IsDate/CDate rejects .NNN)
                            If Not IsNull(codeExpires) Then
                                Dim expTmp : expTmp = CStr(codeExpires & "")
                                Dim expDot : expDot = InStr(expTmp, ".")
                                If expDot > 0 Then
                                    codeExpires = Left(expTmp, expDot - 1)
                                End If
                            End If
                            If Not IsNull(codeCreated) Then
                                Dim crtTmp : crtTmp = CStr(codeCreated & "")
                                Dim crtDot : crtDot = InStr(crtTmp, ".")
                                If crtDot > 0 Then
                                    codeCreated = Left(crtTmp, crtDot - 1)
                                End If
                            End If
                            
                            ' 判断状态
                            Dim codeStatus, codeStatusClass
                            If codeActive = 0 Then
                                codeStatus = "已失效"
                                codeStatusClass = "badge-inactive"
                            ElseIf IsNull(codeExpires) Then
                                codeStatus = "已过期"
                                codeStatusClass = "badge-expired"
                            ElseIf Not IsDate(codeExpires) Then
                                codeStatus = "已过期"
                                codeStatusClass = "badge-expired"
                            ElseIf CDate(codeExpires) < Now() Then
                                codeStatus = "已过期"
                                codeStatusClass = "badge-expired"
                            ElseIf CLng(codeUsed) >= CLng(codeMax) Then
                                codeStatus = "已用完"
                                codeStatusClass = "badge-exhausted"
                            Else
                                codeStatus = "有效"
                                codeStatusClass = "badge-active"
                            End If
                    %>
                    <tr>
                        <td><%= codeId %></td>
                        <td><%= codeAdminName %></td>
                        <td><%= codeUsed %> / <%= codeMax %></td>
                        <td><%= SafeFormatDateTime(codeExpires, 2) %></td>
                        <td><span class="badge <%= codeStatusClass %>"><%= codeStatus %></span></td>
                        <td><%= SafeFormatDateTime(codeCreated, 2) %></td>
                        <td>
                            <% 
                            Dim codeOriginalToken
                            codeOriginalToken = rsCodes("OriginalToken")
                            If codeStatus = "有效" And Not IsNull(codeOriginalToken) And codeOriginalToken <> "" Then 
                                Dim codeFullUrl, codeProtocol
                                If Request.ServerVariables("HTTPS") = "on" Then
                                    codeProtocol = "https://"
                                Else
                                    codeProtocol = "http://"
                                End If
                                codeFullUrl = codeProtocol & Request.ServerVariables("HTTP_HOST") & "/user/register.asp?token=" & Server.URLEncode(codeOriginalToken)
                            %>
                            <button type="button" class="btn btn-sm" onclick="copyCodeUrl(this, '<%= Server.HTMLEncode(codeFullUrl) %>')" style="background:#4caf50;color:#fff;padding:4px 10px;margin-right:4px;">
                                <i class="fas fa-copy"></i> 复制
                            </button>
                            <form method="post" style="display:inline;" onsubmit="return confirm('确定要使此推荐码失效吗？');">
                                <%= GetCSRFTokenField() %>
                                <input type="hidden" name="action" value="deactivate">
                                <input type="hidden" name="token_id" value="<%= codeId %>">
                                <button type="submit" class="btn btn-danger btn-sm">失效</button>
                            </form>
                            <% Else %>
                            <span style="color:#666;font-size:11px;">-</span>
                            <% End If %>
                        </td>
                    </tr>
                    <%
                            rsCodes.MoveNext
                        Loop
                        rsCodes.Close
                    End If
                    Set rsCodes = Nothing
                    %>
                </tbody>
            </table>
        </div>
        
        <% ElseIf activeTab = "relations" Then %>
        <!-- 推荐关系链 -->
        <div class="card">
            <h3><i class="fas fa-sitemap"></i> 推荐关系链</h3>
            
            <table>
                <thead>
                    <tr>
                        <th>上级会员</th>
                        <th>下级会员</th>
                        <th>层级</th>
                        <th>建立时间</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    Dim rsRelations
                    Set rsRelations = ExecuteQuery("SELECT TOP 100 rr.*, u1.Username AS AncestorName, u2.Username AS DescendantName " & _
                        "FROM ReferralRelations rr " & _
                        "LEFT JOIN Users u1 ON rr.AncestorUserID = u1.UserID " & _
                        "LEFT JOIN Users u2 ON rr.DescendantUserID = u2.UserID " & _
                        "ORDER BY rr.CreatedAt DESC")
                    
                    If Not rsRelations Is Nothing Then
                        Do While Not rsRelations.EOF
                            Dim relAncestor, relDescendant, relDepth, relCreated
                            relAncestor = rsRelations("AncestorName")
                            If IsNull(relAncestor) Then relAncestor = "ID:" & rsRelations("AncestorUserID")
                            relDescendant = rsRelations("DescendantName")
                            If IsNull(relDescendant) Then relDescendant = "ID:" & rsRelations("DescendantUserID")
                            relDepth = rsRelations("Depth")
                            relCreated = rsRelations("CreatedAt")
                    %>
                    <tr>
                        <td><%= relAncestor %></td>
                        <td><%= relDescendant %></td>
                        <td>第<%= relDepth %>级</td>
                        <td><%= SafeFormatDateTime(relCreated, 2) %></td>
                    </tr>
                    <%
                            rsRelations.MoveNext
                        Loop
                        rsRelations.Close
                    End If
                    Set rsRelations = Nothing
                    %>
                </tbody>
            </table>
        </div>
        <% End If %>
    </div>
    
    <script>
    function copyCodeUrl(btn, url) {
        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(url).then(function() {
                var orig = btn.innerHTML;
                btn.innerHTML = '<i class="fas fa-check"></i>';
                setTimeout(function() { btn.innerHTML = orig; }, 1500);
            });
        } else {
            var ta = document.createElement('textarea');
            ta.value = url;
            document.body.appendChild(ta);
            ta.select();
            document.execCommand('copy');
            document.body.removeChild(ta);
            btn.innerHTML = '<i class="fas fa-check"></i>';
            setTimeout(function() { btn.innerHTML = '<i class="fas fa-copy"></i> \u590d\u5236'; }, 1500);
        }
    }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
