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

Function SafeField(rs, fldName, defVal)
    On Error Resume Next
    Dim v : v = rs(fldName)
    If Err.Number <> 0 Or IsNull(v) Then
        Err.Clear
        SafeField = defVal
    Else
        SafeField = v
    End If
    On Error GoTo 0
End Function

' 尝试扩展Users表（非关键，失败不影响页面）
On Error Resume Next
conn.Execute "IF COL_LENGTH('Users','TotalSpent') IS NULL ALTER TABLE Users ADD TotalSpent DECIMAL(18,2) DEFAULT 0"
Err.Clear
conn.Execute "IF COL_LENGTH('Users','OrderCount') IS NULL ALTER TABLE Users ADD OrderCount INT DEFAULT 0"
Err.Clear
conn.Execute "IF COL_LENGTH('Users','LastOrderDate') IS NULL ALTER TABLE Users ADD LastOrderDate DATETIME NULL"
Err.Clear
conn.Execute "IF COL_LENGTH('Users','PreferredNote') IS NULL ALTER TABLE Users ADD PreferredNote NVARCHAR(50)"
Err.Clear
conn.Execute "IF COL_LENGTH('Users','CustomerTier') IS NULL ALTER TABLE Users ADD CustomerTier NVARCHAR(20) DEFAULT 'bronze'"
Err.Clear
conn.Execute "IF COL_LENGTH('Users','FavoriteCategory') IS NULL ALTER TABLE Users ADD FavoriteCategory NVARCHAR(100)"
Err.Clear
On Error GoTo 0

' 同步CRM数据（非关键，失败不阻塞）
On Error Resume Next
conn.Execute "UPDATE u SET u.TotalSpent = ISNULL((SELECT SUM(o.TotalAmount) FROM Orders o WHERE o.UserID=u.UserID AND o.Status IN ('paid','shipped','delivered')),0) FROM Users u"
Err.Clear
conn.Execute "UPDATE u SET u.OrderCount = ISNULL((SELECT COUNT(*) FROM Orders o WHERE o.UserID=u.UserID AND o.Status IN ('paid','shipped','delivered')),0) FROM Users u"
Err.Clear
conn.Execute "UPDATE u SET u.LastOrderDate = (SELECT MAX(o.CreatedAt) FROM Orders o WHERE o.UserID=u.UserID) FROM Users u"
Err.Clear
On Error GoTo 0

' 分页参数
Dim page, pageSize
page = CInt(IIf(Request.QueryString("page") = "", 1, Request.QueryString("page")))
pageSize = 20

' 筛选
Dim filterTier : filterTier = Request.QueryString("tier")
Dim sqlWhere : sqlWhere = ""
If filterTier <> "" Then sqlWhere = sqlWhere & " AND CustomerTier='" & Replace(filterTier,"'","''") & "'"

' 获取客户列表（使用子查询计算消费数据，不依赖TotalSpent列是否存在）
Dim rsCustomers
Set rsCustomers = Server.CreateObject("ADODB.Recordset")
rsCustomers.PageSize = pageSize
rsCustomers.CursorLocation = 3
Dim custSQL
custSQL = "SELECT u.*, " & _
    "ISNULL((SELECT SUM(o.TotalAmount) FROM Orders o WHERE o.UserID=u.UserID AND o.Status IN ('paid','shipped','delivered')),0) AS CalcSpent, " & _
    "ISNULL((SELECT COUNT(*) FROM Orders o WHERE o.UserID=u.UserID AND o.Status IN ('paid','shipped','delivered')),0) AS CalcOrders, " & _
    "(SELECT MAX(o.CreatedAt) FROM Orders o WHERE o.UserID=u.UserID) AS CalcLastOrder " & _
    "FROM Users u WHERE 1=1 " & sqlWhere & " ORDER BY CalcSpent DESC, u.CreatedAt DESC"
rsCustomers.Open custSQL, conn, 1, 1

If Not rsCustomers.EOF Then
    If page > rsCustomers.PageCount Then page = rsCustomers.PageCount
    If page < 1 Then page = 1
    rsCustomers.AbsolutePage = page
End If

' 获取统计（直接从Orders表计算，不依赖TotalSpent列）
Dim totalUsers, todayUsers, vipCount
Dim avgSpent, totalRevenue, highValueCount
totalUsers = SafeNum(GetScalar("SELECT COUNT(*) FROM Users"))
todayUsers = SafeNum(GetScalar("SELECT COUNT(*) FROM Users WHERE CAST(CreatedAt AS DATE)=CAST(GETDATE() AS DATE)"))
vipCount = SafeNum(GetScalar("SELECT COUNT(*) FROM Users WHERE IsVIP=1"))
totalRevenue = SafeNum(GetScalar("SELECT ISNULL(SUM(TotalAmount),0) FROM Orders WHERE Status IN ('paid','shipped','delivered')"))
avgSpent = IIf(totalUsers>0, SafeNum(GetScalar("SELECT AVG(sub.ts) FROM (SELECT ISNULL(SUM(TotalAmount),0) AS ts FROM Orders WHERE Status IN ('paid','shipped','delivered') GROUP BY UserID HAVING SUM(TotalAmount)>0) sub")), 0)
highValueCount = SafeNum(GetScalar("SELECT COUNT(*) FROM (SELECT UserID FROM Orders WHERE Status IN ('paid','shipped','delivered') GROUP BY UserID HAVING SUM(TotalAmount)>=1000) sub"))

Call LogAdminAction("查看客户列表", "operation", "Users", "", "")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>客户管理 - 运营管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; }
        .main-content { margin-left: 260px; padding: 30px; min-height: 100vh; }
        .page-header { margin-bottom: 25px; }
        .page-title { color: #fff; font-size: 24px; margin: 0 0 8px; }
        .breadcrumb { color: #888; font-size: 13px; }
        .breadcrumb a { color: #00bcd4; text-decoration: none; }
        .stats-cards { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }
        .stats-sub { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 20px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); padding: 20px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); text-align: center; cursor: pointer; transition: all 0.2s; }
        .stat-card:hover { transform: translateY(-2px); }
        .stat-card .icon { font-size: 28px; margin-bottom: 8px; }
        .stat-card h3 { font-size: 28px; margin: 8px 0; color: #00bcd4; font-weight: 700; }
        .stat-card p { color: #888; margin: 0; font-size: 13px; }
        .stat-card.gold h3 { color: #FFD700; }
        .stat-card.green h3 { color: #4CAF50; }
        .stat-card.purple h3 { color: #CE93D8; }
        .customers-table { width: 100%; background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; overflow: hidden; border: 1px solid rgba(255,255,255,0.05); }
        .customers-table th { background: linear-gradient(135deg, #00bcd4, #00838f); color: white; padding: 12px; text-align: left; font-size: 13px; }
        .customers-table td { padding: 12px; border-bottom: 1px solid rgba(255,255,255,0.04); font-size: 13px; color: #e0e0e0; }
        .customers-table tr:hover { background: rgba(255,255,255,0.03); }
        .user-avatar { width: 38px; height: 38px; border-radius: 50%; background: linear-gradient(135deg, #00bcd4, #00838f); color: white; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 14px; }
        .vip-badge { display: inline-block; padding: 3px 8px; background: #FFD700; color: #1a1a2e; border-radius: 10px; font-size: 11px; font-weight: bold; }
        .tier-badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 600; }
        .tier-gold { background: rgba(255,215,0,0.2); color: #FFD700; border: 1px solid rgba(255,215,0,0.3); }
        .tier-silver { background: rgba(192,192,192,0.2); color: #C0C0C0; border: 1px solid rgba(192,192,192,0.3); }
        .tier-bronze { background: rgba(205,127,50,0.2); color: #CD853F; border: 1px solid rgba(205,127,50,0.3); }
        .amount-text { color: #4CAF50; font-weight: 600; }
        .filter-tabs { display: flex; gap: 10px; margin-bottom: 20px; }
        .filter-tab { padding: 8px 16px; border-radius: 20px; font-size: 13px; text-decoration: none; color: #888; background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.08); transition: all 0.2s; }
        .filter-tab:hover, .filter-tab.active { background: rgba(0,188,212,0.15); color: #00bcd4; border-color: rgba(0,188,212,0.3); }
        .pagination { display: flex; justify-content: center; gap: 10px; margin-top: 20px; }
        .pagination a { padding: 8px 15px; background: #2d2d44; border-radius: 6px; text-decoration: none; color: #00bcd4; border: 1px solid rgba(255,255,255,0.06); }
        .pagination a.active { background: linear-gradient(135deg, #00bcd4, #00838f); color: white; }
        .note-tag { display: inline-block; padding: 2px 8px; background: rgba(0,188,212,0.12); color: #80DEEA; border-radius: 8px; font-size: 11px; }
        .cell-secondary { font-size: 12px; color: #888; }
        @media (max-width: 768px) { .main-content { margin-left: 0; } .stats-cards { grid-template-columns: 1fr 1fr; } .stats-sub { grid-template-columns: 1fr; } }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-users"></i> 客户CRM画像</h2>
            <div class="breadcrumb">
                <a href="index.asp">运营中心</a> / <span>客户管理</span>
            </div>
        </div>
        
        <div class="stats-cards">
            <div class="stat-card" onclick="location.href='?'">
                <i class="fas fa-users" style="color:#00bcd4;font-size:28px;"></i>
                <h3><%= totalUsers %></h3>
                <p>总客户数</p>
            </div>
            <div class="stat-card green" onclick="location.href='?'">
                <i class="fas fa-user-plus" style="color:#4CAF50;font-size:28px;"></i>
                <h3><%= todayUsers %></h3>
                <p>今日新增</p>
            </div>
            <div class="stat-card gold">
                <i class="fas fa-crown" style="color:#FFD700;font-size:28px;"></i>
                <h3><%= vipCount %></h3>
                <p>VIP会员</p>
            </div>
            <div class="stat-card purple">
                <i class="fas fa-gem" style="color:#CE93D8;font-size:28px;"></i>
                <h3><%= highValueCount %></h3>
                <p>高价值客户(≥¥1000)</p>
            </div>
        </div>
        
        <div class="stats-sub">
            <div class="stat-card">
                <p style="font-size:12px;color:#888;margin:0;">累计营收</p>
                <h3 style="font-size:22px;color:#4CAF50;">¥<%= FormatNumber(totalRevenue,0) %></h3>
            </div>
            <div class="stat-card">
                <p style="font-size:12px;color:#888;margin:0;">客均消费</p>
                <h3 style="font-size:22px;color:#00bcd4;">¥<%= FormatNumber(avgSpent,0) %></h3>
            </div>
            <div class="stat-card">
                <p style="font-size:12px;color:#888;margin:0;">转化率</p>
                <h3 style="font-size:22px;color:#FF9800;"><%= IIf(totalUsers>0, FormatNumber(SafeNum(GetScalar("SELECT COUNT(DISTINCT UserID) FROM Orders WHERE Status IN ('paid','shipped','delivered')"))/totalUsers*100,1), 0) %>%</h3>
            </div>
        </div>

        <div class="filter-tabs">
            <a href="?" class="filter-tab <%= IIf(filterTier="","active","") %>"><i class="fas fa-users"></i> 全部</a>
            <a href="?tier=gold" class="filter-tab <%= IIf(filterTier="gold","active","") %>">🥇 金牌</a>
            <a href="?tier=silver" class="filter-tab <%= IIf(filterTier="silver","active","") %>">🥈 银牌</a>
            <a href="?tier=bronze" class="filter-tab <%= IIf(filterTier="bronze","active","") %>">🥉 铜牌</a>
        </div>
        
        <table class="customers-table">
            <thead>
                <tr>
                    <th>客户</th>
                    <th>等级</th>
                    <th>累计消费</th>
                    <th>订单数</th>
                    <th>偏好香型</th>
                    <th>最后购买</th>
                    <th>注册时间</th>
                    <th>操作</th>
                </tr>
            </thead>
            <tbody>
                <% If Not rsCustomers Is Nothing Then
                Do While Not rsCustomers.EOF
                    Dim tier : tier = SafeField(rsCustomers, "CustomerTier", "bronze")
                    If tier = "" Then tier = "bronze"
                    Dim tierClass
                    Select Case tier
                        Case "gold": tierClass = "tier-gold"
                        Case "silver": tierClass = "tier-silver"
                        Case Else: tierClass = "tier-bronze"
                    End Select
                    Dim tierLabel
                    Select Case tier
                        Case "gold": tierLabel = "金牌"
                        Case "silver": tierLabel = "银牌"
                        Case Else: tierLabel = "铜牌"
                    End Select
                    Dim spent : spent = SafeNum(rsCustomers("CalcSpent"))
                    Dim ordCnt : ordCnt = SafeNum(rsCustomers("CalcOrders"))
                    ' 自动算等级
                    If spent >= 5000 Then tierLabel = "金牌" : tierClass = "tier-gold"
                    If spent >= 2000 And spent < 5000 Then tierLabel = "银牌" : tierClass = "tier-silver"
                    If spent < 2000 Then tierLabel = "铜牌" : tierClass = "tier-bronze"
                    If ordCnt = 0 Then tierLabel = "新客" : tierClass = ""
                %>
                <tr>
                    <td>
                        <div style="display: flex; align-items: center; gap: 12px;">
                            <div class="user-avatar"><%= UCase(Left(rsCustomers("Username"), 1)) %></div>
                            <div>
                                <div style="font-weight: 600;"><%= rsCustomers("Username") %></div>
                                <div class="cell-secondary"><%= rsCustomers("Email") %></div>
                            </div>
                            <% If (rsCustomers("IsVIP") & "") = "1" Then %>
                            <span class="vip-badge">VIP</span>
                            <% End If %>
                        </div>
                    </td>
                    <td>
                        <% If ordCnt > 0 Then %>
                        <span class="tier-badge <%= tierClass %>"><%= tierLabel %></span>
                        <% Else %>
                        <span style="color:#b0b0b0;font-size:12px;">新客</span>
                        <% End If %>
                    </td>
                    <td class="amount-text">¥<%= FormatNumber(spent,0) %></td>
                    <td><%= ordCnt %></td>
                    <td>
                        <% 
                        Dim prefNote : prefNote = SafeField(rsCustomers, "PreferredNote", "")
                        If prefNote <> "" Then Response.Write "<span class='note-tag'>" & prefNote & "</span>" Else Response.Write "<span style='color:#b0b0b0;'>—</span>"
                        %>
                    </td>
                    <td class="cell-secondary">
                        <% Dim lastOrd : lastOrd = SafeField(rsCustomers, "CalcLastOrder", "")
                        If lastOrd = "" Then Response.Write "—" Else Response.Write lastOrd
                        %>
                    </td>
                    <td class="cell-secondary"><%= SafeFormatDateTime(rsCustomers("CreatedAt"), 2) %></td>
                    <td>
                        <a href="customer_detail.asp?id=<%= rsCustomers("UserID") %>" class="btn btn--primary btn-sm"><i class="fas fa-eye"></i> 详情</a>
                    </td>
                </tr>
                <% rsCustomers.MoveNext
                Loop
                Dim totalPageCount : totalPageCount = rsCustomers.PageCount
                rsCustomers.Close
                End If %>
            </tbody>
        </table>
        
        <% If totalPageCount > 1 Then
            Dim tierParam : tierParam = IIf(filterTier<>"","&tier=" & filterTier,"")
            Dim totalPages : totalPages = totalPageCount
        %>
        <div class="pagination">
            <% If page > 1 Then %>
            <a href="customers.asp?page=<%= page-1 %><%= tierParam %>"><i class="fas fa-chevron-left"></i></a>
            <% End If %>
            <% Dim i
            For i = 1 To totalPages %>
            <a href="customers.asp?page=<%= i %><%= tierParam %>" class="<%= IIf(i=page, "active", "") %>"><%= i %></a>
            <% Next
            If page < totalPages Then %>
            <a href="customers.asp?page=<%= page+1 %><%= tierParam %>"><i class="fas fa-chevron-right"></i></a>
            <% End If %>
        </div>
        <% End If %>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
