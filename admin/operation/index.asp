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

' 获取运营统计数据
Dim todayOrders, weekOrders, monthOrders, totalOrders
todayOrders = GetScalar("SELECT COUNT(*) FROM Orders WHERE CAST(CreatedAt AS DATE) = CAST(GETDATE() AS DATE)")
weekOrders = GetScalar("SELECT COUNT(*) FROM Orders WHERE CreatedAt >= DATEADD(day, -7, CAST(GETDATE() AS DATE))")
monthOrders = GetScalar("SELECT COUNT(*) FROM Orders WHERE CreatedAt >= DATEADD(month, -1, CAST(GETDATE() AS DATE))")
totalOrders = GetScalar("SELECT COUNT(*) FROM Orders")

Dim todayRevenue, weekRevenue, monthRevenue
todayRevenue = GetScalar("SELECT ISNULL(SUM(TotalAmount), 0) FROM Orders WHERE CAST(CreatedAt AS DATE) = CAST(GETDATE() AS DATE) AND Status = 'Paid'")
weekRevenue = GetScalar("SELECT ISNULL(SUM(TotalAmount), 0) FROM Orders WHERE CreatedAt >= DATEADD(day, -7, CAST(GETDATE() AS DATE)) AND Status = 'Paid'")
monthRevenue = GetScalar("SELECT ISNULL(SUM(TotalAmount), 0) FROM Orders WHERE CreatedAt >= DATEADD(month, -1, CAST(GETDATE() AS DATE)) AND Status = 'Paid'")

Dim totalCustomers, newCustomersToday
totalCustomers = GetScalar("SELECT COUNT(*) FROM Users")
newCustomersToday = GetScalar("SELECT COUNT(*) FROM Users WHERE CAST(CreatedAt AS DATE) = CAST(GETDATE() AS DATE)")

Dim pendingOrders, processingOrders, completedOrders
pendingOrders = GetScalar("SELECT COUNT(*) FROM Orders WHERE Status = 'Pending'")
processingOrders = GetScalar("SELECT COUNT(*) FROM Orders WHERE Status = 'Processing'")
completedOrders = GetScalar("SELECT COUNT(*) FROM Orders WHERE Status = 'Paid'")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>运营概览 - 运营管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        /* 深色主题 */
        body {
            background: #1a1a2e;
            color: #e0e0e0;
        }
        .main-content {
            color: #e0e0e0;
        }
        
        /* 页面标题区 */
        .page-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 25px;
            padding-bottom: 15px;
            border-bottom: 1px solid rgba(255,255,255,0.08);
        }
        .page-title {
            font-size: 24px;
            color: #fff;
            margin: 0;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .page-title i { color: #00bcd4; }
        .breadcrumb { font-size: 13px; color: #888; }
        .breadcrumb a { color: #00bcd4; text-decoration: none; }
        .breadcrumb a:hover { text-decoration: underline; }
        
        /* 统计卡片 */
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin-bottom: 30px; }
        .stat-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            padding: 25px;
            border-radius: 12px;
            border: 1px solid rgba(255,255,255,0.05);
            text-align: center;
            transition: transform 0.3s, box-shadow 0.3s;
        }
        .stat-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.3);
            border-color: rgba(0,188,212,0.2);
        }
        .stat-icon { width: 60px; height: 60px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 0 auto 15px; font-size: 24px; color: white; }
        .stat-card.orders .stat-icon { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); }
        .stat-card.revenue .stat-icon { background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); }
        .stat-card.customers .stat-icon { background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%); }
        .stat-card.products .stat-icon { background: linear-gradient(135deg, #43e97b 0%, #38f9d7 100%); }
        .stat-value { font-size: 32px; font-weight: bold; color: #fff; margin-bottom: 5px; }
        .stat-label { color: #aaa; font-size: 14px; }
        .stat-change { font-size: 12px; margin-top: 8px; }
        .stat-change.up { color: #4CAF50; }
        .stat-change.down { color: #f44336; }
        
        /* 中部面板 */
        .dashboard-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 25px; margin-bottom: 30px; }
        .dashboard-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 25px;
            border: 1px solid rgba(255,255,255,0.05);
        }
        .dashboard-card h3 { font-size: 18px; color: #fff; margin-bottom: 20px; display: flex; align-items: center; gap: 10px; }
        .dashboard-card h3 i { color: #00bcd4; }
        
        /* 订单状态 */
        .order-status-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; }
        .status-item {
            background: rgba(0,0,0,0.2);
            padding: 20px;
            border-radius: 10px;
            text-align: center;
            transition: transform 0.2s;
        }
        .status-item:hover { transform: translateY(-2px); }
        .status-item.pending { border-left: 4px solid #ff9800; }
        .status-item.processing { border-left: 4px solid #2196F3; }
        .status-item.completed { border-left: 4px solid #4CAF50; }
        .status-number { font-size: 28px; font-weight: bold; margin-bottom: 5px; }
        .status-item.pending .status-number { color: #ff9800; }
        .status-item.processing .status-number { color: #2196F3; }
        .status-item.completed .status-number { color: #4CAF50; }
        .status-label { color: #888; font-size: 13px; }
        
        /* 快捷操作 */
        .quick-actions { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; }
        .action-btn {
            display: flex;
            flex-direction: column;
            align-items: center;
            padding: 20px;
            background: rgba(0,188,212,0.08);
            border-radius: 10px;
            text-decoration: none;
            color: #ccc;
            transition: all 0.3s;
            border: 1px solid rgba(255,255,255,0.05);
        }
        .action-btn:hover {
            background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%);
            color: #fff;
            transform: translateY(-3px);
            border-color: #00bcd4;
        }
        .action-btn i { font-size: 24px; margin-bottom: 10px; }
        .action-btn span { font-size: 13px; }
        
        /* 最近订单表格 */
        .recent-orders { width: 100%; border-collapse: collapse; }
        .recent-orders th {
            text-align: left;
            padding: 12px;
            background: rgba(0,0,0,0.2);
            font-weight: 600;
            color: #888;
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .recent-orders td {
            padding: 12px;
            border-bottom: 1px solid rgba(255,255,255,0.05);
            color: #e0e0e0;
        }
        .recent-orders tr:hover td { background: rgba(255,255,255,0.03); }
        .recent-orders td a { color: #00bcd4; text-decoration: none; }
        .recent-orders td a:hover { text-decoration: underline; }
        
        .order-status { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 500; }
        .status-pending { background: rgba(255,152,0,0.2); color: #ff9800; }
        .status-paid { background: rgba(76,175,80,0.2); color: #4caf50; }
        .status-processing { background: rgba(33,150,243,0.2); color: #2196f3; }
        

        /* 响应式 */
        @media (max-width: 1200px) {
            .stats-grid { grid-template-columns: repeat(2, 1fr); }
            .dashboard-grid { grid-template-columns: 1fr; }
            .quick-actions { grid-template-columns: repeat(2, 1fr); }
        }
        @media (max-width: 768px) {
            .stats-grid { grid-template-columns: 1fr; }
            .quick-actions { grid-template-columns: 1fr; }
            .order-status-grid { grid-template-columns: 1fr; }
        }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-home"></i> 运营概览</h2>
            <div class="breadcrumb">
                <a href="index.asp">运营中心</a> / <span>概览</span>
            </div>
        </div>
        
        <!-- 核心指标卡片 -->
        <div class="stats-grid">
            <div class="stat-card orders">
                <div class="stat-icon"><i class="fas fa-shopping-cart"></i></div>
                <div class="stat-value"><%= todayOrders %></div>
                <div class="stat-label">今日订单</div>
                <div class="stat-change up"><i class="fas fa-arrow-up"></i> 较昨日</div>
            </div>
            <div class="stat-card revenue">
                <div class="stat-icon"><i class="fas fa-yen-sign"></i></div>
                <div class="stat-value">¥<%= FormatNumber(CDbl("0" & todayRevenue), 0) %></div>
                <div class="stat-label">今日营收</div>
                <div class="stat-change up"><i class="fas fa-arrow-up"></i> 较昨日</div>
            </div>
            <div class="stat-card customers">
                <div class="stat-icon"><i class="fas fa-users"></i></div>
                <div class="stat-value"><%= newCustomersToday %></div>
                <div class="stat-label">今日新客</div>
                <div class="stat-label" style="margin-top: 5px; color: #999;">累计 <%= totalCustomers %> 人</div>
            </div>
            <div class="stat-card products">
                <div class="stat-icon"><i class="fas fa-box"></i></div>
                <div class="stat-value"><%= GetScalar("SELECT COUNT(*) FROM Products WHERE IsActive <> 0") %></div>
                <div class="stat-label">在售商品</div>
                <div class="stat-label" style="margin-top: 5px; color: #999;">累计 <%= GetScalar("SELECT COUNT(*) FROM Products") %> 款</div>
            </div>
        </div>
        
        <!-- 中部面板 -->
        <div class="dashboard-grid">
            <!-- 订单状态统计 -->
            <div class="dashboard-card">
                <h3><i class="fas fa-chart-pie"></i> 订单状态分布</h3>
                <div class="order-status-grid">
                    <div class="status-item pending">
                        <div class="status-number"><%= pendingOrders %></div>
                        <div class="status-label">待付款</div>
                    </div>
                    <div class="status-item processing">
                        <div class="status-number"><%= processingOrders %></div>
                        <div class="status-label">处理中</div>
                    </div>
                    <div class="status-item completed">
                        <div class="status-number"><%= completedOrders %></div>
                        <div class="status-label">已完成</div>
                    </div>
                </div>
            </div>
            
            <!-- 快捷操作 -->
            <div class="dashboard-card">
                <h3><i class="fas fa-bolt"></i> 快捷操作</h3>
                <div class="quick-actions">
                    <a href="orders.asp" class="action-btn">
                        <i class="fas fa-shopping-cart"></i>
                        <span>订单管理</span>
                    </a>
                    <a href="customers.asp" class="action-btn">
                        <i class="fas fa-users"></i>
                        <span>客户管理</span>
                    </a>
                    <a href="tier_management.asp" class="action-btn">
                        <i class="fas fa-layer-group"></i>
                        <span>会员等级</span>
                    </a>
                    <a href="points.asp" class="action-btn">
                        <i class="fas fa-coins"></i>
                        <span>积分管理</span>
                    </a>
                    <a href="recipes.asp" class="action-btn">
                        <i class="fas fa-fire"></i>
                        <span>配方推荐</span>
                    </a>
                </div>
            </div>
        </div>
        
        <!-- 最近订单 -->
        <div class="dashboard-card">
            <h3><i class="fas fa-list"></i> 最近订单</h3>
            <table class="recent-orders">
                <thead>
                    <tr>
                        <th>订单号</th>
                        <th>客户</th>
                        <th>金额</th>
                        <th>状态</th>
                        <th>时间</th>
                    </tr>
                </thead>
                <tbody>
                    <% 
                    Dim rsRecent
                    Set rsRecent = ExecuteQuery("SELECT TOP 5 o.*, u.Username, u.FullName FROM Orders o LEFT JOIN Users u ON o.UserID = u.UserID ORDER BY o.CreatedAt DESC")
                    If Not rsRecent Is Nothing Then
                        Do While Not rsRecent.EOF
                    %>
                    <tr>
                        <td><a href="order_detail.asp?id=<%= rsRecent("OrderID") %>"><%= rsRecent("OrderNo") %></a></td>
                        <td><%= IIf(rsRecent("FullName")<>"", rsRecent("FullName"), rsRecent("Username")) %></td>
                        <td>¥<%= FormatNumber(CDbl("0" & rsRecent("TotalAmount")), 2) %></td>
                        <td>
                            <% Select Case rsRecent("Status")
                                Case "Pending" %>
                                <span class="order-status status-pending">待付款</span>
                            <% Case "Paid" %>
                                <span class="order-status status-paid">已支付</span>
                            <% Case "Processing" %>
                                <span class="order-status status-processing">处理中</span>
                            <% End Select %>
                        </td>
                        <td><%= SafeFormatDateTime(rsRecent("CreatedAt"), 2) %></td>
                    </tr>
                    <% 
                        rsRecent.MoveNext
                        Loop
                        rsRecent.Close
                    End If
                    %>
                </tbody>
            </table>
            <div style="text-align: center; margin-top: 15px;">
                <a href="orders.asp" class="admin-btn admin-btn-primary">查看全部订单</a>
            </div>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
