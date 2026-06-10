<%@ Language="VBScript" CodePage="65001" %>
<% Response.Charset = "UTF-8" %>
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<!--#include file="../includes/role_auth.asp"-->
<%
Call OpenConnection()
' auth check handled by role_auth.asp
Call LogAdminAction("查看绩效仪表板", "operation", "performance_dashboard", "", "")

Function GC(sql)
    Dim rs, c: c = 0
    Set rs = conn.Execute(sql)
    If Not rs Is Nothing Then If Not rs.EOF Then c = rs(0) : rs.Close
    Set rs = Nothing: GC = c
End Function

' KPI
Dim totalRevenue, totalOrders, totalCustomers, avgRating, stockTurnover, returnRate
totalRevenue = GC("SELECT ISNULL(SUM(TotalAmount),0) FROM Orders WHERE Status='Paid'")
totalOrders = GC("SELECT COUNT(*) FROM Orders WHERE Status NOT IN ('Pending','Cancelled')")
totalCustomers = GC("SELECT COUNT(*) FROM Users WHERE IsActive=1")
avgRating = GC("SELECT ISNULL(AVG(CAST(Rating AS FLOAT)),0) FROM ProductReviews WHERE Status='Approved'")
stockTurnover = GC("SELECT ISNULL(SUM(StockQty),0) FROM RawMaterialInventory")
returnRate = GC("SELECT COUNT(*) FROM Orders WHERE Status='Returned'")

' 产品线业绩
Dim rsPL
Set rsPL = conn.Execute("SELECT ProductType, COUNT(*) AS PCount, ISNULL(SUM(o.TotalAmount),0) AS Sales FROM Products p LEFT JOIN Orders o ON 1=0 GROUP BY ProductType")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>绩效评估仪表板</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background:#1a1a2e; color:#e0e0e0; }
        .main-content { padding:30px; margin-left:260px; }
        .page-title { font-size:24px; margin-bottom:25px; display:flex; align-items:center; gap:12px; }
        .page-title i { color:#00bcd4; }
        .kpi-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:15px;margin-bottom:25px; }
        .kpi-card { background:linear-gradient(135deg,#2d2d44,#1e1e32); border-radius:12px; padding:20px; border:1px solid rgba(255,255,255,0.06); }
        .kpi-card .lbl { font-size:12px; color:#888; }
        .kpi-card .val { font-size:24px; font-weight:700; margin:8px 0; }
        .kpi-card .sub { font-size:11px; }
        .section-card { background:linear-gradient(135deg,#2d2d44,#1e1e32); border-radius:12px; padding:20px; margin-bottom:20px; border:1px solid rgba(255,255,255,0.06); }
        .section-title { color:#00bcd4; font-size:16px; margin:0 0 15px 0; display:flex; align-items:center; gap:10px; }
        .data-table { width:100%; border-collapse:collapse; font-size:13px; }
        .data-table th { background:#1a1a2e; color:#888; padding:10px; text-align:left; border-bottom:1px solid #3a3a4a; }
        .data-table td { padding:8px 10px; border-bottom:1px solid rgba(255,255,255,0.04); }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    <div class="main-content">
        <div class="breadcrumb" style="color:#888;font-size:14px;margin-bottom:20px;">
            <a href="index.asp" style="color:#00bcd4;text-decoration:none;">运营管理</a> / <span>绩效仪表板</span>
        </div>
        <h2 class="page-title"><i class="fas fa-chart-bar"></i> 绩效评估仪表板</h2>

        <div class="kpi-grid">
            <div class="kpi-card"><div class="lbl">总营收</div><div class="val" style="color:#4CAF50;">¥<%= FormatNumber(totalRevenue,0) %></div><div class="sub" style="color:#888;">已支付订单合计</div></div>
            <div class="kpi-card"><div class="lbl">订单总数</div><div class="val" style="color:#2196F3;"><%= totalOrders %></div><div class="sub" style="color:#888;">有效订单</div></div>
            <div class="kpi-card"><div class="lbl">客户总数</div><div class="val" style="color:#FF9800;"><%= totalCustomers %></div><div class="sub" style="color:#888;">活跃用户</div></div>
            <div class="kpi-card"><div class="lbl">平均评分</div><div class="val" style="color:#9C27B0;"><%= FormatNumber(avgRating,1) %></div><div class="sub" style="color:#888;">客户满意度</div></div>
        </div>

        <div class="section-card">
            <h3 class="section-title"><i class="fas fa-layer-group"></i> 三大产品线业绩对比</h3>
            <table class="data-table">
                <tr><th>产品线</th><th>产品数量</th><th>销售额</th></tr>
                <% If Not rsPL Is Nothing Then
                    Do While Not rsPL.EOF %>
                <tr><td><%= rsPL("ProductType") %></td><td><%= rsPL("PCount") %></td><td class="cost">¥<%= FormatNumber(rsPL("Sales"),2) %></td></tr>
                <%
                    rsPL.MoveNext
                    Loop
                    rsPL.Close
                End If %>
            </table>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>