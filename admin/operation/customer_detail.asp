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

' 设置超时防止无限等待
Server.ScriptTimeout = 30
conn.CommandTimeout = 15

' 获取用户ID
Dim userId
userId = Request.QueryString("id")

If userId = "" Or Not IsNumeric(userId) Then
    Response.Redirect "customers.asp"
    Response.End
End If

' 获取用户基本信息
Dim rsUser
Dim userNotFound : userNotFound = False
Dim username, email, fullName, phone, address, city, postalCode, createdAt, isActive, isVIP, points
Dim orderCount, totalSpent
Dim rsOrders
Set rsOrders = Nothing
username = "" : email = "" : fullName = "" : phone = "" : address = "" : city = "" : postalCode = ""
orderCount = 0 : totalSpent = 0 : isVIP = False : points = 0

Set rsUser = ExecuteQuery("SELECT * FROM Users WHERE UserID = " & CLng(userId))

If rsUser Is Nothing Then
    ' 查询执行失败
    userNotFound = True
ElseIf rsUser.EOF Then
    ' 未找到用户
    rsUser.Close
    Set rsUser = Nothing
    userNotFound = True
Else
    userNotFound = False
    ' 存储用户信息到变量
    username = rsUser("Username")
    email = rsUser("Email")
    fullName = rsUser("FullName") & ""
    phone = rsUser("Phone") & ""
    address = rsUser("Address") & ""
    city = rsUser("City") & ""
    postalCode = rsUser("PostalCode") & ""
    createdAt = rsUser("CreatedAt")
    isActive = rsUser("IsActive")
    ' 检查是否有VIP和Points字段（可能在后续版本中添加）
    On Error Resume Next
    isVIP = rsUser("IsVIP")
    If Err.Number <> 0 Then
        isVIP = False
        Err.Clear
    End If
    points = rsUser("Points")
    If Err.Number <> 0 Then
        points = 0
        Err.Clear
    End If
    On Error GoTo 0
    
    rsUser.Close
    Set rsUser = Nothing
    
    ' 获取订单统计
    orderCount = GetScalar("SELECT COUNT(*) FROM Orders WHERE UserID = " & CLng(userId))
    totalSpent = GetScalar("SELECT ISNULL(SUM(TotalAmount),0) FROM Orders WHERE UserID = " & CLng(userId) & " AND Status IN ('Paid','paid','shipped','delivered')")
    If IsNull(totalSpent) Or totalSpent = "" Then totalSpent = 0
    
    ' 获取订单历史（限制返回最近50条防止数据过多）
    Set rsOrders = ExecuteQuery("SELECT TOP 50 OrderID, OrderNo, TotalAmount, Status, CreatedAt FROM Orders WHERE UserID = " & CLng(userId) & " ORDER BY CreatedAt DESC")
End If

On Error Resume Next
Call LogAdminAction("查看客户详情", "operation", "Users", userId, "")
Err.Clear
On Error GoTo 0
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>客户详情 - <%= Server.HTMLEncode(username & "") %> - 运营管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        .customer-header { display: flex; align-items: center; gap: 20px; margin-bottom: 30px; padding: 25px; background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); border-radius: 12px; color: white; }
        .customer-avatar { width: 80px; height: 80px; border-radius: 50%; background: linear-gradient(135deg, #2d2d44, #1e1e32); color: #00bcd4; display: flex; align-items: center; justify-content: center; font-size: 32px; font-weight: bold; }
        .customer-info h2 { margin: 0 0 5px 0; font-size: 24px; }
        .customer-info p { margin: 0; opacity: 0.9; }
        .vip-badge-large { display: inline-block; padding: 5px 15px; background: #ffd700; color: #1a1a2e; border-radius: 20px; font-size: 14px; font-weight: bold; margin-left: 10px; }
        .status-badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; }
        .status-active { background: #4CAF50; color: white; }
        .status-inactive { background: #f44336; color: white; }
        
        .info-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; margin-bottom: 30px; }
        .info-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); padding: 20px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.2); }
        .info-card h3 { margin-top: 0; border-bottom: 2px solid rgba(255,255,255,0.08); padding-bottom: 10px; margin-bottom: 15px; font-size: 16px; color: #e0e0e0; }
        .info-row { display: flex; margin-bottom: 12px; font-size: 14px; }
        .info-row .label { width: 100px; color: #b0b0b0; font-weight: bold; }
        .info-row .value { flex: 1; color: #e0e0e0; }
        
        .stats-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 30px; }
        .stat-box { background: linear-gradient(135deg, #2d2d44, #1e1e32); padding: 20px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.2); text-align: center; }
        .stat-box i { font-size: 28px; color: #00bcd4; margin-bottom: 10px; }
        .stat-box .number { font-size: 24px; font-weight: bold; color: #e0e0e0; margin-bottom: 5px; }
        .stat-box .label { color: #b0b0b0; font-size: 13px; }
        
        .orders-section { background: linear-gradient(135deg, #2d2d44, #1e1e32); padding: 20px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.2); }
        .orders-section h3 { margin-top: 0; border-bottom: 2px solid rgba(255,255,255,0.08); padding-bottom: 10px; margin-bottom: 15px; font-size: 16px; color: #e0e0e0; }
        .orders-table { width: 100%; border-collapse: collapse; }
        .orders-table th { background: linear-gradient(135deg, #00bcd4, #00838f); padding: 12px; text-align: left; font-weight: 600; color: #fff; border-bottom: 2px solid rgba(255,255,255,0.08); }
        .orders-table td { padding: 12px; border-bottom: 1px solid rgba(255,255,255,0.08); }
        .orders-table tr:hover { background: rgba(255,255,255,0.04); }
        .order-status { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: bold; }
        .status-Pending { background: rgba(255,152,0,0.12); color: #ff9800; }
        .status-Paid { background: rgba(46,125,50,0.2); color: #66bb6a; }
        .status-Failed { background: rgba(198,40,40,0.2); color: #ef5350; }
        .status-Refunded { background: rgba(25,118,210,0.2); color: #42a5f5; }
        .btn-view { color: #00bcd4; text-decoration: none; }
        .btn-view:hover { text-decoration: underline; }
        
        .not-found { text-align: center; padding: 60px 20px; background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.2); }
        .not-found i { font-size: 64px; color: rgba(255,255,255,0.1); margin-bottom: 20px; }
        .not-found h3 { color: #b0b0b0; margin-bottom: 10px; }
        .not-found p { color: #888; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-user"></i> 客户详情</h2>
            <div class="breadcrumb">
                <a href="index.asp">运营中心</a> / <a href="customers.asp">客户管理</a> / <span>客户详情</span>
            </div>
        </div>
        
        <% If userNotFound Then %>
        <div class="not-found">
            <i class="fas fa-user-slash"></i>
            <h3>未找到该客户</h3>
            <p>该用户ID不存在或已被删除</p>
            <br>
            <a href="customers.asp" class="admin-btn admin-btn-primary"><i class="fas fa-arrow-left"></i> 返回客户列表</a>
        </div>
        <% Else %>
        
        <!-- 客户头部信息 -->
        <div class="customer-header">
            <div class="customer-avatar"><%= Left(username, 1) %></div>
            <div class="customer-info">
                <h2>
                    <%= Server.HTMLEncode(username) %>
                    <% If isVIP Then %>
                    <span class="vip-badge-large"><i class="fas fa-crown"></i> VIP</span>
                    <% End If %>
                </h2>
                <p>
                    <i class="fas fa-envelope"></i> <%= Server.HTMLEncode(email) %>
                    <% If phone <> "" Then %>
                    &nbsp;|&nbsp; <i class="fas fa-phone"></i> <%= Server.HTMLEncode(phone) %>
                    <% End If %>
                    &nbsp;|&nbsp; 
                    <% If isActive Then %>
                    <span class="status-badge status-active"><i class="fas fa-check-circle"></i> 正常</span>
                    <% Else %>
                    <span class="status-badge status-inactive"><i class="fas fa-ban"></i> 禁用</span>
                    <% End If %>
                </p>
            </div>
        </div>
        
        <!-- 统计卡片 -->
        <div class="stats-row">
            <div class="stat-box">
                <i class="fas fa-shopping-bag"></i>
                <div class="number"><%= orderCount %></div>
                <div class="label">订单总数</div>
            </div>
            <div class="stat-box">
                <i class="fas fa-yen-sign"></i>
                <div class="number">¥<%= FormatNumber(CDbl("0" & totalSpent), 2) %></div>
                <div class="label">累计消费</div>
            </div>
            <div class="stat-box">
                <i class="fas fa-coins"></i>
                <div class="number"><%= points %></div>
                <div class="label">当前积分</div>
            </div>
            <div class="stat-box">
                <i class="fas fa-calendar-alt"></i>
                <div class="number"><%= SafeFormatDateTime(createdAt, 2) %></div>
                <div class="label">注册日期</div>
            </div>
        </div>
        
        <!-- 详细信息 -->
        <div class="info-grid">
            <div class="info-card">
                <h3><i class="fas fa-id-card"></i> 基本信息</h3>
                <div class="info-row">
                    <span class="label">用户ID:</span>
                    <span class="value">#<%= userId %></span>
                </div>
                <div class="info-row">
                    <span class="label">用户名:</span>
                    <span class="value"><%= Server.HTMLEncode(username) %></span>
                </div>
                <div class="info-row">
                    <span class="label">真实姓名:</span>
                    <span class="value"><%= Server.HTMLEncode(fullName) %></span>
                </div>
                <div class="info-row">
                    <span class="label">电子邮箱:</span>
                    <span class="value"><%= Server.HTMLEncode(email) %></span>
                </div>
                <div class="info-row">
                    <span class="label">联系电话:</span>
                    <span class="value"><%= Server.HTMLEncode(phone) %></span>
                </div>
            </div>
            
            <div class="info-card">
                <h3><i class="fas fa-map-marker-alt"></i> 地址信息</h3>
                <div class="info-row">
                    <span class="label">所在城市:</span>
                    <span class="value"><%= Server.HTMLEncode(city) %></span>
                </div>
                <div class="info-row">
                    <span class="label">详细地址:</span>
                    <span class="value"><%= Server.HTMLEncode(address) %></span>
                </div>
                <div class="info-row">
                    <span class="label">邮政编码:</span>
                    <span class="value"><%= Server.HTMLEncode(postalCode) %></span>
                </div>
            </div>
        </div>
        
        <!-- 订单历史 -->
        <div class="orders-section">
            <h3><i class="fas fa-history"></i> 订单历史</h3>
            <% If Not rsOrders Is Nothing Then %>
            <% If Not rsOrders.EOF Then %>
            <table class="orders-table">
                <thead>
                    <tr>
                        <th>订单号</th>
                        <th>下单时间</th>
                        <th>订单金额</th>
                        <th>状态</th>
                        <th>操作</th>
                    </tr>
                </thead>
                <tbody>
                    <% Do While Not rsOrders.EOF %>
                    <tr>
                        <td><strong>#<%= rsOrders("OrderNo") %></strong></td>
                        <td><%= SafeFormatDateTime(rsOrders("CreatedAt"), 2) %></td>
                        <td>¥<%= FormatNumber(CDbl("0" & rsOrders("TotalAmount")), 2) %></td>
                        <td>
                            <% 
                            Dim orderStatus
                            orderStatus = rsOrders("Status") & ""
                            Select Case orderStatus
                                Case "Pending"
                            %>
                            <span class="order-status status-Pending">待支付</span>
                            <%      Case "Paid" %>
                            <span class="order-status status-Paid">已支付</span>
                            <%      Case "Failed" %>
                            <span class="order-status status-Failed">支付失败</span>
                            <%      Case "Refunded" %>
                            <span class="order-status status-Refunded">已退款</span>
                            <%      Case Else %>
                            <span class="order-status"><%= Server.HTMLEncode(orderStatus) %></span>
                            <% End Select %>
                        </td>
                        <td>
                            <a href="order_detail.asp?order_id=<%= rsOrders("OrderID") %>" class="btn-view"><i class="fas fa-eye"></i> 查看</a>
                        </td>
                    </tr>
                    <% rsOrders.MoveNext %>
                    <% Loop %>
                </tbody>
            </table>
            <% Else %>
            <div style="text-align: center; padding: 40px; color: #999;">
                <i class="fas fa-inbox" style="font-size: 48px; margin-bottom: 15px; display: block;"></i>
                该客户暂无订单记录
            </div>
            <% End If %>
            <% rsOrders.Close %>
            <% Set rsOrders = Nothing %>
            <% Else %>
            <div style="text-align: center; padding: 40px; color: #999;">
                <i class="fas fa-inbox" style="font-size: 48px; margin-bottom: 15px; display: block;"></i>
                无法加载订单数据
            </div>
            <% End If %>
        </div>
        
        <% End If %>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
