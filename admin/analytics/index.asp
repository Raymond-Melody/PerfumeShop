<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="../includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
' ============================================
' V15.0 数据分析仪表盘 (Analytics Dashboard)
' ============================================
Call OpenConnection()

' --- 概览卡片 ---
Dim todayOrders, todayRevenue, weekOrders, weekRevenue, monthOrders, monthRevenue
todayOrders = CLng(GetScalar("SELECT COUNT(*) FROM Orders WHERE CAST(CreatedAt AS DATE)=CAST(GETDATE() AS DATE)"))
todayRevenue = CDbl(GetScalar("SELECT ISNULL(SUM(TotalAmount),0) FROM Orders WHERE CAST(CreatedAt AS DATE)=CAST(GETDATE() AS DATE) AND Status<>'Cancelled'"))
weekOrders = CLng(GetScalar("SELECT COUNT(*) FROM Orders WHERE CreatedAt>=DATEADD(DAY,-7,GETDATE())"))
weekRevenue = CDbl(GetScalar("SELECT ISNULL(SUM(TotalAmount),0) FROM Orders WHERE CreatedAt>=DATEADD(DAY,-7,GETDATE()) AND Status<>'Cancelled'"))
monthOrders = CLng(GetScalar("SELECT COUNT(*) FROM Orders WHERE CreatedAt>=DATEADD(DAY,-30,GETDATE())"))
monthRevenue = CDbl(GetScalar("SELECT ISNULL(SUM(TotalAmount),0) FROM Orders WHERE CreatedAt>=DATEADD(DAY,-30,GETDATE()) AND Status<>'Cancelled'"))

' --- 订单状态分布 ---
Dim rsStatus
Set rsStatus = conn.Execute("SELECT Status, COUNT(*) AS Cnt, ISNULL(SUM(TotalAmount),0) AS Amt FROM Orders GROUP BY Status ORDER BY Status")

' --- 热门产品Top10 ---
Dim rsTop
Set rsTop = conn.Execute("SELECT TOP 10 p.ProductName, COUNT(od.DetailID) AS SalesQty, ISNULL(SUM(od.Subtotal),0) AS SalesAmt FROM OrderDetails od INNER JOIN Products p ON od.ProductID=p.ProductID GROUP BY p.ProductName ORDER BY SalesQty DESC")

' --- 每日营收趋势（30天） ---
Dim rsDaily
Set rsDaily = conn.Execute("SELECT CAST(CreatedAt AS DATE) AS OrderDate, COUNT(*) AS Cnt, ISNULL(SUM(TotalAmount),0) AS Amt FROM Orders WHERE CreatedAt>=DATEADD(DAY,-30,GETDATE()) GROUP BY CAST(CreatedAt AS DATE) ORDER BY OrderDate")

' V16: 用户增长趋势（30天）
Dim rsUserGrowth
Set rsUserGrowth = conn.Execute("SELECT CAST(CreatedAt AS DATE) AS RegDate, COUNT(*) AS Cnt FROM Users WHERE CreatedAt>=DATEADD(DAY,-30,GETDATE()) GROUP BY CAST(CreatedAt AS DATE) ORDER BY RegDate")

' V16: 转化漏斗
Dim totalVisits, addToCartCount, checkoutCount, orderCount, conversionRate
totalVisits = CLng(GetScalar("SELECT COUNT(*) FROM TrackingEvents WHERE EventType='page_view' AND CreatedAt>=DATEADD(DAY,-30,GETDATE())"))
addToCartCount = CLng(GetScalar("SELECT COUNT(*) FROM Cart WHERE CreatedAt>=DATEADD(DAY,-30,GETDATE())"))
checkoutCount = CLng(GetScalar("SELECT COUNT(*) FROM TrackingEvents WHERE EventType='checkout_start' AND CreatedAt>=DATEADD(DAY,-30,GETDATE())"))
orderCount = CLng(GetScalar("SELECT COUNT(*) FROM Orders WHERE CreatedAt>=DATEADD(DAY,-30,GETDATE())"))
If totalVisits > 0 Then conversionRate = Round(orderCount / totalVisits * 100, 2) Else conversionRate = 0

' V16: 今日实时数据
Dim todayVisits, todayNewUsers, pendingOrders
todayVisits = CLng(GetScalar("SELECT COUNT(*) FROM TrackingEvents WHERE EventType='page_view' AND CAST(CreatedAt AS DATE)=CAST(GETDATE() AS DATE)"))
todayNewUsers = CLng(GetScalar("SELECT COUNT(*) FROM Users WHERE CAST(CreatedAt AS DATE)=CAST(GETDATE() AS DATE)"))
pendingOrders = CLng(GetScalar("SELECT COUNT(*) FROM Orders WHERE Status='Pending'"))

%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>数据分析仪表盘 - V18</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
<link rel="stylesheet" href="/css/admin.css">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
.dashboard-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px; padding: 20px; }
.stat-card { background: #2d2d44; border-radius: 10px; padding: 20px; color: #e0e0e0; }
.stat-card .stat-value { font-size: 28px; font-weight: 700; margin: 8px 0; }
.stat-card .stat-label { color: #888; font-size: 13px; }
.stat-card .stat-icon { float: right; font-size: 32px; opacity: 0.3; }
.stat-card.revenue { border-left: 4px solid #4caf50; }
.stat-card.orders { border-left: 4px solid #2196f3; }
.stat-card.users { border-left: 4px solid #ff9800; }
.stat-card.avg { border-left: 4px solid #9c27b0; }
.charts-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; padding: 0 20px 20px; }
@media (max-width: 768px) { .charts-grid { grid-template-columns: 1fr; } }
.chart-card { background: #2d2d44; border-radius: 10px; padding: 20px; }
.chart-card h3 { color: #e0e0e0; margin: 0 0 16px; font-size: 16px; }
.chart-container { position: relative; height: 300px; }
.table-card { background: #2d2d44; border-radius: 10px; padding: 20px; margin: 0 20px 20px; }
.table-card h3 { color: #e0e0e0; margin: 0 0 16px; }
.data-table { width: 100%; border-collapse: collapse; color: #ccc; }
.data-table th { text-align: left; padding: 10px; border-bottom: 1px solid #444; color: #888; font-size: 12px; }
.data-table td { padding: 10px; border-bottom: 1px solid #333; font-size: 13px; }
.data-table tr:hover { background: rgba(255,255,255,0.03); }
.badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; }
.badge-success { background: rgba(76,175,80,0.2); color: #4caf50; }
.badge-warning { background: rgba(255,152,0,0.2); color: #ff9800; }
.badge-info { background: rgba(33,150,243,0.2); color: #2196f3; }
.badge-danger { background: rgba(244,67,54,0.2); color: #f44336; }
.header { padding: 20px; display: flex; justify-content: space-between; align-items: center; }
.header h1 { color: #e0e0e0; margin: 0; }
.period-btn { background: #333; border: none; color: #ccc; padding: 6px 14px; border-radius: 6px; cursor: pointer; margin-left: 6px; }
.period-btn.active { background: #2196f3; color: #fff; }
</style>
</head>
<body>
<!--#include file="../includes/nav.asp"-->
<div class="main-content">
<div class="header">
    <h1><i class="fas fa-chart-bar"></i> 数据分析仪表盘 V18</h1>
    <div>
        <span style="color:#888;font-size:12px;margin-right:12px;">实时数据</span>
        <button class="period-btn active" onclick="changePeriod(7)">7天</button>
        <button class="period-btn" onclick="changePeriod(30)">30天</button>
        <button class="period-btn" onclick="changePeriod(90)">90天</button>
    </div>
</div>

<!-- V16: 概览卡片（扩展） -->
<div class="dashboard-grid">
    <div class="stat-card revenue">
        <div class="stat-icon"><i class="fas fa-dollar-sign"></i></div>
        <div class="stat-label">今日营收</div>
        <div class="stat-value">¥<%= FormatNumber(todayRevenue, 2) %></div>
    </div>
    <div class="stat-card orders">
        <div class="stat-icon"><i class="fas fa-shopping-cart"></i></div>
        <div class="stat-label">今日订单</div>
        <div class="stat-value"><%= todayOrders %></div>
    </div>
    <div class="stat-card users">
        <div class="stat-icon"><i class="fas fa-eye"></i></div>
        <div class="stat-label">今日访问</div>
        <div class="stat-value"><%= todayVisits %></div>
    </div>
    <div class="stat-card avg">
        <div class="stat-icon"><i class="fas fa-user-plus"></i></div>
        <div class="stat-label">今日新用户</div>
        <div class="stat-value"><%= todayNewUsers %></div>
    </div>
    <div class="stat-card" style="border-left:4px solid #e91e63;">
        <div class="stat-icon"><i class="fas fa-chart-pie"></i></div>
        <div class="stat-label">30天转化率</div>
        <div class="stat-value"><%= conversionRate %>%</div>
    </div>
    <div class="stat-card" style="border-left:4px solid #ff5722;">
        <div class="stat-icon"><i class="fas fa-clock"></i></div>
        <div class="stat-label">待处理订单</div>
        <div class="stat-value"><%= pendingOrders %></div>
    </div>
</div>

<!-- V16: 转化漏斗 -->
<div style="padding:0 20px 20px;">
    <div class="chart-card">
        <h3><i class="fas fa-filter"></i> 30天转化漏斗</h3>
        <div style="display:flex;align-items:center;justify-content:space-around;padding:20px 0;">
            <%
            Dim funnelSteps(3), funnelLabels(3), funnelWidths(3), funnelColors(3)
            funnelLabels(0) = "访问": funnelSteps(0) = totalVisits: funnelWidths(0) = 100: funnelColors(0) = "#2196f3"
            funnelLabels(1) = "加购": funnelSteps(1) = addToCartCount: If totalVisits>0 Then funnelWidths(1) = Round(addToCartCount/totalVisits*100,0) Else funnelWidths(1) = 0: funnelColors(1) = "#4caf50"
            funnelLabels(2) = "结账": funnelSteps(2) = checkoutCount: If totalVisits>0 Then funnelWidths(2) = Round(checkoutCount/totalVisits*100,0) Else funnelWidths(2) = 0: funnelColors(2) = "#ff9800"
            funnelLabels(3) = "下单": funnelSteps(3) = orderCount: If totalVisits>0 Then funnelWidths(3) = Round(orderCount/totalVisits*100,0) Else funnelWidths(3) = 0: funnelColors(3) = "#f44336"
            Dim fi
            For fi = 0 To 3
            %>
            <div style="text-align:center;flex:1;">
                <div style="background:<%= funnelColors(fi) %>;color:#fff;padding:12px;border-radius:8px;margin:4px;width:<%= funnelWidths(fi) %>%;min-width:80px;margin:0 auto;">
                    <div style="font-size:22px;font-weight:700;"><%= funnelSteps(fi) %></div>
                    <div style="font-size:11px;opacity:0.9;"><%= funnelLabels(fi) %></div>
                </div>
            </div>
            <% Next %>
        </div>
    </div>
</div>

<!-- 图表区 -->
<div class="charts-grid">
    <div class="chart-card">
        <h3><i class="fas fa-chart-line"></i> 每日营收趋势</h3>
        <div class="chart-container"><canvas id="revenueChart"></canvas></div>
    </div>
    <div class="chart-card">
        <h3><i class="fas fa-chart-pie"></i> 订单状态分布</h3>
        <div class="chart-container"><canvas id="statusChart"></canvas></div>
    </div>
</div>

<!-- 热门产品表 -->
<div class="table-card">
    <h3><i class="fas fa-fire"></i> 热门产品 Top 10</h3>
    <table class="data-table">
        <thead><tr><th>#</th><th>产品名称</th><th>销量</th><th>销售额</th></tr></thead>
        <tbody>
            <%
            Dim rank: rank = 0
            If Not rsTop Is Nothing Then
                Do While Not rsTop.EOF And rank < 10
                    rank = rank + 1
            %>
            <tr>
                <td><%= rank %></td>
                <td><%= rsTop("ProductName") %></td>
                <td><%= rsTop("SalesQty") %></td>
                <td>¥<%= FormatNumber(rsTop("SalesAmt"), 2) %></td>
            </tr>
            <%
                    rsTop.MoveNext
                Loop
                rsTop.Close
            End If
            Set rsTop = Nothing
            %>
        </tbody>
    </table>
</div>

<script>
// 营收趋势图
(function() {
    var dailyData = [];
    <% If Not rsDaily Is Nothing Then
        Do While Not rsDaily.EOF %>
            dailyData.push({date: '<%= rsDaily("OrderDate") %>', cnt: <%= rsDaily("Cnt") %>, amt: <%= rsDaily("Amt") %>});
        <% rsDaily.MoveNext
        Loop
        rsDaily.Close
    End If
    Set rsDaily = Nothing %>
    
    var ctx = document.getElementById('revenueChart').getContext('2d');
    new Chart(ctx, {
        type: 'line',
        data: {
            labels: dailyData.map(function(d) { return d.date; }),
            datasets: [{
                label: '营收 (¥)',
                data: dailyData.map(function(d) { return d.amt; }),
                borderColor: '#4caf50',
                backgroundColor: 'rgba(76,175,80,0.1)',
                fill: true,
                tension: 0.3
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { labels: { color: '#aaa' } } },
            scales: {
                x: { ticks: { color: '#888', maxTicksLimit: 10 } },
                y: { ticks: { color: '#888', callback: function(v) { return '¥' + v; } } }
            }
        }
    });
})();

// 状态分布图
(function() {
    var statusData = {};
    <% If Not rsStatus Is Nothing Then
        Do While Not rsStatus.EOF %>
            statusData['<%= rsStatus("Status") %>'] = <%= rsStatus("Cnt") %>;
        <% rsStatus.MoveNext
        Loop
        rsStatus.Close
    End If
    Set rsStatus = Nothing %>
    
    var labels = Object.keys(statusData);
    var values = Object.values(statusData);
    var colors = ['#4caf50','#2196f3','#ff9800','#f44336','#9c27b0','#00bcd4'];
    
    new Chart(document.getElementById('statusChart'), {
        type: 'doughnut',
        data: {
            labels: labels,
            datasets: [{ data: values, backgroundColor: colors.slice(0, labels.length) }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: 'bottom', labels: { color: '#aaa', padding: 15 } } }
        }
    });
})();

function changePeriod(days) {
    document.querySelectorAll('.period-btn').forEach(function(b) { b.classList.remove('active'); });
    event.target.classList.add('active');
    window.location = '?days=' + days;
}
</script>
</div>
<!-- .main-content -->
</body>
</html>
<%
Call CloseConnection()
%>