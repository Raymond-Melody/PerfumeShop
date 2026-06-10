<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"

' 检查是否登录
If Session("UserID") = "" Then
    Response.Redirect "/user/login.asp"
End If
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/member_utils.asp"-->
<%
Call OpenConnection()

Dim userId, rsUser, userLevel, userPoints
userId = Session("UserID")

' 获取会员等级积分信息
userLevel = MU_CalcUserLevel(userId)
userPoints = MU_GetUserPoints(userId)

' 获取优惠券数量
Dim couponCount
couponCount = 0
Dim rsCoupon
Set rsCoupon = conn.Execute("SELECT COUNT(*) FROM SiteSettings WHERE SettingKey LIKE 'Coupon_%' AND SettingKey LIKE '%_User" & userId & "'")
If Not rsCoupon Is Nothing Then
    If Not rsCoupon.EOF Then couponCount = rsCoupon(0)
    rsCoupon.Close
End If
Set rsCoupon = Nothing

' 获取用户信息
Set rsUser = ExecuteQuery("SELECT * FROM Users WHERE UserID = " & userId)

' 获取订单统计
Dim orderCount, totalSpent
orderCount = GetScalar("SELECT COUNT(*) FROM Orders WHERE UserID = " & userId)
totalSpent = GetScalar("SELECT ISNULL(SUM(TotalAmount), 0) FROM Orders WHERE UserID = " & userId & " AND Status NOT IN ('Cancelled')")
If IsNull(orderCount) Then orderCount = 0
If IsNull(totalSpent) Then totalSpent = 0

' 获取最近订单
Dim rsOrders
Set rsOrders = ExecuteQuery("SELECT TOP 5 * FROM Orders WHERE UserID = " & userId & " ORDER BY CreatedAt DESC")
%>
<!--#include file="../includes/header.asp"-->

<div class="breadcrumb">
    <div class="container">
        <a href="/index.asp">首页</a>
        <span class="separator">/</span>
        <span>个人中心</span>
    </div>
</div>

<div class="container">
    <div class="user-center">
        <!-- 侧边栏 -->
        <aside class="user-sidebar">
            <div class="user-profile">
                <h3><%= HTMLEncode(Session("Username")) %></h3>
                <p><%= HTMLEncode(Session("Email")) %></p>
                <div style="margin-top:8px;">
                    <% Call MU_RenderLevelBadge(userId) %>
                </div>
                <div style="margin-top:5px;font-size:12px;color:#888;">
                    积分: <strong><%= userPoints %></strong> | 
                    等级折扣: <strong><%= FormatPercent(1 - MU_GetLevelDiscount(userLevel), 0) %></strong>
                </div>
            </div>
            
            <nav class="user-nav">
                <a href="/user/index.asp" class="active"><i class="fas fa-home"></i> 个人中心</a>
                <% If Not rsUser Is Nothing Then %>
                    <% If rsUser("UserRole") = "KOL" Then %>
                        <a href="/user/kol_products.asp" style="background: #fff0f6; color: #eb2f96; font-weight: bold;"><i class="fas fa-star"></i> KOL推荐管理</a>
                    <% End If %>
                <% End If %>
                <a href="/user/orders.asp"><i class="fas fa-list"></i> 我的订单</a>
                <a href="/user/settings.asp"><i class="fas fa-user-edit"></i> 账户设置</a>
                <a href="/user/addresses.asp"><i class="fas fa-map-marker-alt"></i> 收货地址</a>
                <a href="/user/favorites.asp"><i class="fas fa-heart"></i> 我的收藏</a>
                <a href="/user/logout.asp"><i class="fas fa-sign-out-alt"></i> 退出登录</a>
            </nav>
        </aside>

        <!-- 主内容 -->
        <div class="user-main">
            <div class="welcome-section">
                <h1>欢迎回来，<%= HTMLEncode(Session("Username")) %>！</h1>
                <p>在这里管理您的订单和账户信息</p>
            </div>
            
            <!-- 统计卡片 -->
            <div class="stats-grid">
                <div class="stat-card">
                    <div class="stat-icon"><i class="fas fa-shopping-bag"></i></div>
                    <div class="stat-info">
                        <span class="stat-value"><%= orderCount %></span>
                        <span class="stat-label">订单数量</span>
                    </div>
                </div>
                <div class="stat-card">
                    <div class="stat-icon"><i class="fas fa-yen-sign"></i></div>
                    <div class="stat-info">
                        <span class="stat-value"><%= FormatMoney(totalSpent) %></span>
                        <span class="stat-label">累计消费</span>
                    </div>
                </div>
                <div class="stat-card">
                    <div class="stat-icon"><i class="fas fa-coins"></i></div>
                    <div class="stat-info">
                        <span class="stat-value"><%= userPoints %></span>
                        <span class="stat-label">积分余额</span>
                    </div>
                </div>
                <div class="stat-card">
                    <div class="stat-icon"><i class="fas fa-ticket-alt"></i></div>
                    <div class="stat-info">
                        <span class="stat-value"><%= couponCount %></span>
                        <span class="stat-label">优惠券</span>
                    </div>
                </div>
            </div>
            
            <!-- 快捷入口 -->
            <div class="quick-actions">
                <h2>快捷操作</h2>
                <div class="action-grid">
                    <a href="/products.asp" class="action-item">
                        <i class="fas fa-spray-can"></i>
                        <span>选购香水</span>
                    </a>
                    <a href="/customize.asp" class="action-item">
                        <i class="fas fa-magic"></i>
                        <span>开始定制</span>
                    </a>
                    <a href="/cart.asp" class="action-item">
                        <i class="fas fa-shopping-cart"></i>
                        <span>我的购物车</span>
                    </a>
                    <a href="/user/orders.asp" class="action-item">
                        <i class="fas fa-truck"></i>
                        <span>查看物流</span>
                    </a>
                </div>
            </div>
            
            <!-- 最近订单 -->
            <div class="recent-orders">
                <div class="section-header">
                    <h2>最近订单</h2>
                    <a href="/user/orders.asp">查看全部 <i class="fas fa-arrow-right"></i></a>
                </div>
                
                <% If Not rsOrders Is Nothing And Not rsOrders.EOF Then %>
                <div class="orders-list">
                    <% Do While Not rsOrders.EOF %>
                    <div class="order-item">
                        <div class="order-info">
                            <span class="order-no">订单号: <%= rsOrders("OrderNo") %></span>
                            <span class="order-date"><%= SafeFormatDateTime(rsOrders("CreatedAt"), 2) %></span>
                        </div>
                        <div class="order-amount">
                            <%= FormatMoney(rsOrders("TotalAmount")) %>
                        </div>
                        <div class="order-status">
                            <%
                            Dim statusClass, statusText
                            Select Case rsOrders("Status")
                                Case "Pending"
                                    statusClass = "pending"
                                    statusText = "待付款"
                                Case "Paid"
                                    statusClass = "paid"
                                    statusText = "已付款"
                                Case "Processing"
                                    statusClass = "processing"
                                    statusText = "制作中"
                                Case "Shipped"
                                    statusClass = "shipped"
                                    statusText = "已发货"
                                Case "Delivered"
                                    statusClass = "delivered"
                                    statusText = "已完成"
                                Case "Cancelled"
                                    statusClass = "cancelled"
                                    statusText = "已取消"
                                Case Else
                                    statusClass = ""
                                    statusText = rsOrders("Status")
                            End Select
                            %>
                            <span class="status-badge <%= statusClass %>"><%= statusText %></span>
                        </div>
                        <div class="order-action">
                            <a href="/user/order_detail.asp?order_id=<%= rsOrders("OrderID") %>" class="btn btn-sm btn-outline">查看详情</a>
                        </div>
                    </div>
                    <%
                        rsOrders.MoveNext
                    Loop
                    %>
                </div>
                <%
                Else
                %>
                <div class="empty-orders">
                    <i class="fas fa-inbox"></i>
                    <p>暂无订单记录</p>
                    <a href="/products.asp" class="btn btn-primary">去选购</a>
                </div>
                <%
                End If
                %>
            </div>
        </div>
    </div>
</div>

<%
If Not rsOrders Is Nothing Then
    rsOrders.Close
    Set rsOrders = Nothing
End If
If Not rsUser Is Nothing Then
    rsUser.Close
    Set rsUser = Nothing
End If
%>

<!--#include file="../includes/footer.asp"-->
<%
Call CloseConnection()
%>
