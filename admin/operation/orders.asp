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

' 获取筛选参数
Dim statusFilter, dateFilter, keyword
statusFilter = Request.QueryString("status")
dateFilter = Request.QueryString("date")
keyword = Request.QueryString("keyword")

' 构建查询条件
Dim whereClause
whereClause = "WHERE 1=1"

If statusFilter <> "" Then
    whereClause = whereClause & " AND o.Status = '" & SafeSQL(statusFilter) & "'"
End If

If dateFilter <> "" Then
    Select Case dateFilter
        Case "today"
            whereClause = whereClause & " AND CAST(o.CreatedAt AS DATE) = CAST(GETDATE() AS DATE)"
        Case "week"
            whereClause = whereClause & " AND o.CreatedAt >= DATEADD(day, -7, CAST(GETDATE() AS DATE))"
        Case "month"
            whereClause = whereClause & " AND o.CreatedAt >= DATEADD(month, -1, CAST(GETDATE() AS DATE))"
    End Select
End If

If keyword <> "" Then
    whereClause = whereClause & " AND (o.OrderNo LIKE '%" & SafeSQL(keyword) & "%' OR u.Username LIKE '%" & SafeSQL(keyword) & "%' OR u.FullName LIKE '%" & SafeSQL(keyword) & "%')"
End If

' 获取订单列表
Dim rsOrders
Set rsOrders = ExecuteQuery(_
    "SELECT o.*, u.Username, u.FullName, u.Email, u.Phone, " & _
    "(SELECT COUNT(*) FROM OrderDetails oi WHERE oi.OrderID = o.OrderID) AS ItemCount " & _
    "FROM Orders o " & _
    "LEFT JOIN Users u ON o.UserID = u.UserID " & _
    whereClause & " " & _
    "ORDER BY o.CreatedAt DESC")

' 记录访问日志
Call LogAdminAction("查看订单列表", "operation_orders", "", "", "")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>订单管理 - 运营管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        .filter-bar { background: linear-gradient(135deg, #2d2d44, #1e1e32); padding: 20px; border-radius: 10px; margin-bottom: 25px; box-shadow: 0 4px 15px rgba(0,0,0,0.2); }
        .filter-form { display: flex; gap: 15px; flex-wrap: wrap; align-items: flex-end; }
        .filter-group { display: flex; flex-direction: column; gap: 5px; }
        .filter-group label { font-size: 13px; color: #b0b0b0; font-weight: 500; }
        .filter-group select, .filter-group input { padding: 10px 15px; border: 1px solid rgba(255,255,255,0.15); border-radius: 8px; font-size: 14px; min-width: 150px; }
        .filter-group select:focus, .filter-group input:focus { border-color: #00bcd4; outline: none; }
        
        .orders-table { width: 100%; background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 10px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.2); }
        .orders-table th { background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); color: white; padding: 15px; text-align: left; font-weight: 500; }
        .orders-table td { padding: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); }
        .orders-table tr:hover { background: rgba(255,255,255,0.04); }
        .orders-table tr:last-child td { border-bottom: none; }
        
        .order-no { font-weight: 600; color: #00bcd4; }
        .customer-info { display: flex; flex-direction: column; }
        .customer-name { font-weight: 500; color: #e0e0e0; }
        .customer-contact { font-size: 12px; color: #888; margin-top: 3px; }
        
        .order-amount { font-size: 16px; font-weight: 600; color: #e0e0e0; }
        .order-items { font-size: 12px; color: #888; margin-top: 3px; }
        
        .status-badge { display: inline-block; padding: 6px 14px; border-radius: 20px; font-size: 12px; font-weight: 500; }
        .status-pending { background: rgba(255,152,0,0.12); color: #e65100; }
        .status-paid { background: rgba(46,125,50,0.2); color: #2e7d32; }
        .status-processing { background: rgba(25,118,210,0.2); color: #00bcd4; }
        .status-shipped { background: rgba(123,31,162,0.2); color: #7b1fa2; }
        .status-completed { background: rgba(0,105,92,0.2); color: #00695c; }
        .status-cancelled { background: rgba(198,40,40,0.2); color: #c62828; }
        
        .action-btns { display: flex; gap: 8px; }
        .action-btn { padding: 6px 12px; border-radius: 6px; font-size: 12px; text-decoration: none; transition: all 0.3s; }
        .action-btn.view { background: rgba(25,118,210,0.2); color: #00bcd4; }
        .action-btn.view:hover { background: #00bcd4; color: white; }
        .action-btn.edit { background: rgba(255,152,0,0.12); color: #e65100; }
        .action-btn.edit:hover { background: #e65100; color: white; }
        
        .pagination { display: flex; justify-content: center; gap: 10px; margin-top: 25px; }
        .pagination a { padding: 10px 15px; background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 8px; text-decoration: none; color: #b0b0b0; box-shadow: 0 4px 15px rgba(0,0,0,0.2); }
        .pagination a:hover, .pagination a.active { background: #00bcd4; color: white; }
        
        .stats-summary { display: grid; grid-template-columns: repeat(5, 1fr); gap: 15px; margin-bottom: 25px; }
        .summary-item { background: linear-gradient(135deg, #2d2d44, #1e1e32); padding: 20px; border-radius: 10px; text-align: center; box-shadow: 0 4px 15px rgba(0,0,0,0.2); }
        .summary-number { font-size: 24px; font-weight: bold; margin-bottom: 5px; }
        .summary-label { font-size: 13px; color: #b0b0b0; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-shopping-cart"></i> 订单管理</h2>
            <div class="breadcrumb">
                <a href="index.asp">运营中心</a> / <span>订单管理</span>
            </div>
        </div>
        
        <!-- 统计概览 -->
        <div class="stats-summary">
            <div class="summary-item">
                <div class="summary-number" style="color: #ff9800;"><%= GetScalar("SELECT COUNT(*) FROM Orders WHERE Status = 'Pending'") %></div>
                <div class="summary-label">待付款</div>
            </div>
            <div class="summary-item">
                <div class="summary-number" style="color: #4CAF50;"><%= GetScalar("SELECT COUNT(*) FROM Orders WHERE Status = 'Paid'") %></div>
                <div class="summary-label">已支付</div>
            </div>
            <div class="summary-item">
                <div class="summary-number" style="color: #2196F3;"><%= GetScalar("SELECT COUNT(*) FROM Orders WHERE Status = 'Processing'") %></div>
                <div class="summary-label">处理中</div>
            </div>
            <div class="summary-item">
                <div class="summary-number" style="color: #9C27B0;"><%= GetScalar("SELECT COUNT(*) FROM Orders WHERE Status = 'Shipped'") %></div>
                <div class="summary-label">已发货</div>
            </div>
            <div class="summary-item">
                <div class="summary-number" style="color: #e0e0e0;"><%= GetScalar("SELECT COUNT(*) FROM Orders") %></div>
                <div class="summary-label">全部订单</div>
            </div>
        </div>
        
        <!-- 筛选栏 -->
        <div class="filter-bar">
            <form class="filter-form" method="get" action="orders.asp">
                <div class="filter-group">
                    <label>订单状态</label>
                    <select name="status">
                        <option value="">全部状态</option>
                        <option value="Pending" <%= IIf(statusFilter="Pending", "selected", "") %>>待付款</option>
                        <option value="Paid" <%= IIf(statusFilter="Paid", "selected", "") %>>已支付</option>
                        <option value="Processing" <%= IIf(statusFilter="Processing", "selected", "") %>>处理中</option>
                        <option value="Shipped" <%= IIf(statusFilter="Shipped", "selected", "") %>>已发货</option>
                        <option value="Completed" <%= IIf(statusFilter="Completed", "selected", "") %>>已完成</option>
                        <option value="Cancelled" <%= IIf(statusFilter="Cancelled", "selected", "") %>>已取消</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label>时间范围</label>
                    <select name="date">
                        <option value="">全部时间</option>
                        <option value="today" <%= IIf(dateFilter="today", "selected", "") %>>今日</option>
                        <option value="week" <%= IIf(dateFilter="week", "selected", "") %>>本周</option>
                        <option value="month" <%= IIf(dateFilter="month", "selected", "") %>>本月</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label>关键词搜索</label>
                    <input type="text" name="keyword" value="<%= keyword %>" placeholder="订单号/客户姓名">
                </div>
                <div class="filter-group">
                    <button type="submit" class="admin-btn admin-btn-primary"><i class="fas fa-search"></i> 筛选</button>
                    <a href="orders.asp" class="admin-btn admin-btn-secondary"><i class="fas fa-undo"></i> 重置</a>
                </div>
            </form>
        </div>
        
        <!-- 订单列表 -->
        <table class="orders-table">
            <thead>
                <tr>
                    <th>订单号</th>
                    <th>客户信息</th>
                    <th>订单金额</th>
                    <th>下单时间</th>
                    <th>状态</th>
                    <th>操作</th>
                </tr>
            </thead>
            <tbody>
                <% If Not rsOrders Is Nothing Then %>
                <% Do While Not rsOrders.EOF %>
                <tr>
                    <td>
                        <div class="order-no">#<%= rsOrders("OrderNo") %></div>
                    </td>
                    <td>
                        <div class="customer-info">
                            <span class="customer-name"><%= IIf(rsOrders("FullName")<>"", rsOrders("FullName"), rsOrders("Username")) %></span>
                            <% If rsOrders("Phone") <> "" Then %>
                            <span class="customer-contact"><i class="fas fa-phone"></i> <%= rsOrders("Phone") %></span>
                            <% End If %>
                        </div>
                    </td>
                    <td>
                        <div class="order-amount">¥<%= FormatNumber(CDbl("0" & rsOrders("TotalAmount")), 2) %></div>
                        <div class="order-items"><%= rsOrders("ItemCount") %> 件商品</div>
                    </td>
                    <td><%= SafeFormatDateTime(rsOrders("CreatedAt"), 2) %></td>
                    <td>
                        <% Select Case rsOrders("Status")
                            Case "Pending" %>
                            <span class="status-badge status-pending">待付款</span>
                        <% Case "Paid" %>
                            <span class="status-badge status-paid">已支付</span>
                        <% Case "Processing" %>
                            <span class="status-badge status-processing">处理中</span>
                        <% Case "Shipped" %>
                            <span class="status-badge status-shipped">已发货</span>
                        <% Case "Completed" %>
                            <span class="status-badge status-completed">已完成</span>
                        <% Case "Cancelled" %>
                            <span class="status-badge status-cancelled">已取消</span>
                        <% End Select %>
                    </td>
                    <td>
                        <div class="action-btns">
                            <a href="order_detail.asp?order_id=<%= rsOrders("OrderID") %>" class="action-btn view"><i class="fas fa-eye"></i> 查看</a>
                            <a href="order_edit.asp?id=<%= rsOrders("OrderID") %>" class="action-btn edit"><i class="fas fa-edit"></i> 编辑</a>
                        </div>
                    </td>
                </tr>
                <% rsOrders.MoveNext %>
                <% Loop %>
                <% rsOrders.Close %>
                <% End If %>
            </tbody>
        </table>
        
        <!-- 分页 -->
        <div class="pagination">
            <a href="#" class="active">1</a>
            <a href="#">2</a>
            <a href="#">3</a>
            <a href="#"><i class="fas fa-chevron-right"></i></a>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
