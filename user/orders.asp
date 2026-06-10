<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/payment_config.asp"-->
<%
Call OpenConnection()

' 检查用户是否登录
If Session("UserID") = "" Then
    Response.Redirect "/user/login.asp?return=" & Server.URLEncode(Request.ServerVariables("URL"))
    Response.End
End If

Dim userId
userId = Session("UserID")

' 状态数字到字符串的映射
Function GetStatusString(statusNum)
    If Not IsNumeric(statusNum) Then
        GetStatusString = "Pending"
        Exit Function
    End If
    Select Case CInt(statusNum)
        Case 0: GetStatusString = "Pending"
        Case 1: GetStatusString = "Paid"
        Case 2: GetStatusString = "Failed"
        Case 3: GetStatusString = "Refunded"
        Case Else: GetStatusString = "Pending"
    End Select
End Function

' 获取筛选参数
Dim orderStatusFilter, statusString, searchKeyword, timeRange, startDate, endDate, paymentMethod
orderStatusFilter = Request.QueryString("status")
searchKeyword = Trim(Request.QueryString("search"))
timeRange = Request.QueryString("time_range")
startDate = Request.QueryString("start_date")
endDate = Request.QueryString("end_date")
paymentMethod = Request.QueryString("payment_method")

' 构建查询条件
Dim whereClause
whereClause = " WHERE o.UserID = " & userId & " AND o.Status <> 'Deleted'"

' 添加订单状态筛选
If orderStatusFilter <> "" And IsNumeric(orderStatusFilter) Then
    statusString = GetStatusString(orderStatusFilter)
    whereClause = whereClause & " AND o.Status = '" & statusString & "'"
End If

' 添加关键词搜索
If searchKeyword <> "" Then
    whereClause = whereClause & " AND (o.OrderNo LIKE '%" & SafeSQL(searchKeyword) & "%' OR o.Notes LIKE '%" & SafeSQL(searchKeyword) & "%' OR o.ShippingName LIKE '%" & SafeSQL(searchKeyword) & "%' OR o.PaymentMethod LIKE '%" & SafeSQL(searchKeyword) & "%')"
End If

' 添加时间范围筛选
If timeRange <> "" Then
    Dim timeCondition
    Select Case timeRange
        Case "last_month"
            timeCondition = "DATEADD(month, -1, GETDATE())"
        Case "last_quarter"
            timeCondition = "DATEADD(month, -3, GETDATE())"
        Case "last_half_year"
            timeCondition = "DATEADD(month, -6, GETDATE())"
        Case "last_year"
            timeCondition = "DATEADD(year, -1, GETDATE())"
        Case Else
            timeCondition = ""
    End Select
    
    If timeCondition <> "" Then
        whereClause = whereClause & " AND o.CreatedAt >= " & timeCondition
    End If
End If

' 添加具体日期范围筛选
If startDate <> "" Then
    whereClause = whereClause & " AND o.CreatedAt >= '" & startDate & "'"
End If

If endDate <> "" Then
    whereClause = whereClause & " AND o.CreatedAt <= '" & endDate & "'"
End If

' 添加支付方式筛选
If paymentMethod <> "" And IsNumeric(paymentMethod) Then
    whereClause = whereClause & " AND o.PaymentMethod = " & paymentMethod
End If

' 使用子查询替代LEFT JOIN避免Access兼容性问题
Dim finalSql
finalSql = "SELECT o.*, o.ShippingName AS FullName, " & _
    "(SELECT TOP 1 ReviewID FROM ProductReviews WHERE OrderID = o.OrderID AND UserID = " & userId & ") AS ReviewID, " & _
    "(SELECT TOP 1 [Status] FROM ProductReviews WHERE OrderID = o.OrderID AND UserID = " & userId & ") AS ReviewStatus " & _
    "FROM Orders o " & whereClause & " ORDER BY o.CreatedAt DESC"
' 调试：将SQL保存到Session以便查看
Session("LastOrdersSQL") = finalSql
Set rsOrders = ExecuteQuery(finalSql)
%>
<!--#include file="../includes/header.asp"-->
<style>
.search-filters {
    display: flex;
    flex-wrap: wrap;
    gap: 15px;
    margin-bottom: 20px;
    padding: 15px;
    background-color: #f8f9fa;
    border-radius: 5px;
    align-items: center;
}

.search-box {
    display: flex;
    align-items: center;
    gap: 10px;
}

.search-box input {
    padding: 8px 12px;
    border: 1px solid #ddd;
    border-radius: 4px;
    width: 300px;
}

.date-filter {
    display: flex;
    align-items: center;
    gap: 10px;
}

.date-filter select, .date-filter input {
    padding: 8px;
    border: 1px solid #ddd;
    border-radius: 4px;
}

.advanced-search-panel {
    margin-top: 15px;
    padding: 15px;
    background-color: #e9ecef;
    border-radius: 5px;
}

.advanced-search-fields {
    display: flex;
    flex-wrap: wrap;
    gap: 20px;
}

.field-group {
    display: flex;
    align-items: center;
    gap: 8px;
}

.field-group label {
    font-weight: bold;
    min-width: 80px;
}

.btn-sm {
    padding: 6px 12px;
    font-size: 14px;
}

.btn-secondary {
    background-color: #6c757d;
    color: white;
    border: none;
}

.btn-danger {
    background-color: #dc3545;
    color: white;
    border: none;
}

/* 评价状态样式 */
.review-badge {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    padding: 6px 12px;
    border-radius: 4px;
    font-size: 14px;
    font-weight: 500;
}

.review-badge.reviewed {
    background-color: #d4edda;
    color: #155724;
    border: 1px solid #c3e6cb;
}

.btn-review {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    padding: 6px 12px;
    background-color: #ffc107;
    color: #212529;
    border: none;
    border-radius: 4px;
    text-decoration: none;
    font-size: 14px;
    font-weight: 500;
    transition: background-color 0.2s;
}

.btn-review:hover {
    background-color: #e0a800;
    color: #212529;
}

.btn-review i {
    color: #fff;
}
</style>

<!-- 面包屑导航 -->
<div class="breadcrumb">
    <div class="container">
        <a href="/index.asp">首页</a>
        <span class="separator">/</span>
        <a href="/user/index.asp">个人中心</a>
        <span class="separator">/</span>
        <span>我的订单</span>
    </div>
</div>

<div class="container">
    <div class="user-center">
        <!-- 侧边栏 -->
        <aside class="user-sidebar">
            <div class="user-profile">
                <h3><%= HTMLEncode(Session("Username")) %></h3>
                <p><%= HTMLEncode(Session("Email")) %></p>
            </div>
            
            <nav class="user-nav">
                <a href="/user/index.asp"><i class="fas fa-home"></i> 个人中心</a>
                <a href="/user/orders.asp" class="active"><i class="fas fa-list"></i> 我的订单</a>
                <a href="/user/settings.asp"><i class="fas fa-user-edit"></i> 账户设置</a>
                <a href="/user/addresses.asp"><i class="fas fa-map-marker-alt"></i> 收货地址</a>
                <a href="/user/favorites.asp"><i class="fas fa-heart"></i> 我的收藏</a>
                <a href="/user/logout.asp"><i class="fas fa-sign-out-alt"></i> 退出登录</a>
            </nav>
        </aside>
        
        <!-- 主内容 -->
        <div class="user-main">
            <div class="user-card">
                <h2 class="card-title"><i class="fas fa-list"></i> 我的订单</h2>
                
                <!-- 搜索和筛选工具栏 -->
                <div class="search-filters">
                    <!-- 搜索框 -->
                    <div class="search-box">
                        <input type="text" id="searchInput" placeholder="搜索订单号、商品名称、收货人..." value="<%= HTMLEncode(Request.QueryString("search")) %>" />
                        <button id="searchBtn" class="btn btn-sm">搜索</button>
                    </div>
                    
                    <!-- 时间范围筛选 -->
                    <div class="date-filter">
                        <select id="timeRange">
                            <option value="">全部时间</option>
                            <option value="last_month" <% If Request.QueryString("time_range") = "last_month" Then Response.Write "selected" %>>最近一个月</option>
                            <option value="last_quarter" <% If Request.QueryString("time_range") = "last_quarter" Then Response.Write "selected" %>>最近三个月</option>
                            <option value="last_half_year" <% If Request.QueryString("time_range") = "last_half_year" Then Response.Write "selected" %>>最近半年</option>
                            <option value="last_year" <% If Request.QueryString("time_range") = "last_year" Then Response.Write "selected" %>>最近一年</option>
                        </select>
                        
                        <input type="date" id="startDate" value="<%= Server.HTMLEncode(Request.QueryString("start_date")) %>" />
                        <span>至</span>
                        <input type="date" id="endDate" value="<%= Server.HTMLEncode(Request.QueryString("end_date")) %>" />
                        <button id="applyDateBtn" class="btn btn-sm">应用</button>
                    </div>
                    
                    <!-- 高级搜索按钮 -->
                    <button id="advancedSearchToggle" class="btn btn-sm">高级搜索</button>
                </div>
                
                <!-- 高级搜索面板 -->
                <div id="advancedSearchPanel" class="advanced-search-panel" style="display:none;">
                    <div class="advanced-search-fields">
                        <div class="field-group">
                            <label>订单状态:</label>
                            <select name="status">
                                <option value="">全部状态</option>
                                <option value="Pending" <% If Request.QueryString("status") = "Pending" Then Response.Write "selected" %>>待支付</option>
                                <option value="Paid" <% If Request.QueryString("status") = "Paid" Then Response.Write "selected" %>>已支付</option>
                                <option value="Failed" <% If Request.QueryString("status") = "Failed" Then Response.Write "selected" %>>支付失败</option>
                                <option value="Refunded" <% If Request.QueryString("status") = "Refunded" Then Response.Write "selected" %>>已退款</option>
                            </select>
                        </div>
                        
                        <div class="field-group">
                            <label>支付方式:</label>
                            <select name="payment_method">
                                <option value="">全部方式</option>
                                <option value="1" <% If Request.QueryString("payment_method") = "1" Then Response.Write "selected" %>>微信支付</option>
                                <option value="2" <% If Request.QueryString("payment_method") = "2" Then Response.Write "selected" %>>支付宝</option>
                                <option value="3" <% If Request.QueryString("payment_method") = "3" Then Response.Write "selected" %>>PayPal</option>
                                <option value="4" <% If Request.QueryString("payment_method") = "4" Then Response.Write "selected" %>>货到付款</option>
                            </select>
                        </div>
                        
                        <div class="field-group">
                            <button id="advancedSearchBtn" class="btn btn-sm">高级搜索</button>
                            <button id="clearFiltersBtn" class="btn btn-sm btn-secondary">清空筛选</button>
                        </div>
                    </div>
                </div>
                
                <!-- 订单筛选 -->
                <div class="order-filters">
                    <a href="/user/orders.asp?<% If Request.QueryString("search") <> "" Then Response.Write "search=" & Server.URLEncode(Request.QueryString("search")) & "&" End If %><% If Request.QueryString("time_range") <> "" Then Response.Write "time_range=" & Request.QueryString("time_range") & "&" End If %><% If Request.QueryString("start_date") <> "" Then Response.Write "start_date=" & Request.QueryString("start_date") & "&" End If %><% If Request.QueryString("end_date") <> "" Then Response.Write "end_date=" & Request.QueryString("end_date") & "&" End If %><% If Request.QueryString("payment_method") <> "" Then Response.Write "payment_method=" & Request.QueryString("payment_method") & "&" End If %><% If Request.QueryString("status") <> "" Then Response.Write "status=" & Request.QueryString("status") & "&" End If %>" class="<% If orderStatusFilter = "" Then Response.Write "active" End If %>">全部订单</a>
                    <a href="/user/orders.asp?status=<%= PAYMENT_STATUS_PENDING %>&<% If Request.QueryString("search") <> "" Then Response.Write "search=" & Server.URLEncode(Request.QueryString("search")) & "&" End If %><% If Request.QueryString("time_range") <> "" Then Response.Write "time_range=" & Request.QueryString("time_range") & "&" End If %><% If Request.QueryString("start_date") <> "" Then Response.Write "start_date=" & Request.QueryString("start_date") & "&" End If %><% If Request.QueryString("end_date") <> "" Then Response.Write "end_date=" & Request.QueryString("end_date") & "&" End If %><% If Request.QueryString("payment_method") <> "" Then Response.Write "payment_method=" & Request.QueryString("payment_method") & "&" End If %>" class="<% If orderStatusFilter = PAYMENT_STATUS_PENDING Then Response.Write "active" End If %>">待支付</a>
                    <a href="/user/orders.asp?status=<%= PAYMENT_STATUS_PAID %>&<% If Request.QueryString("search") <> "" Then Response.Write "search=" & Server.URLEncode(Request.QueryString("search")) & "&" End If %><% If Request.QueryString("time_range") <> "" Then Response.Write "time_range=" & Request.QueryString("time_range") & "&" End If %><% If Request.QueryString("start_date") <> "" Then Response.Write "start_date=" & Request.QueryString("start_date") & "&" End If %><% If Request.QueryString("end_date") <> "" Then Response.Write "end_date=" & Request.QueryString("end_date") & "&" End If %><% If Request.QueryString("payment_method") <> "" Then Response.Write "payment_method=" & Request.QueryString("payment_method") & "&" End If %>" class="<% If orderStatusFilter = PAYMENT_STATUS_PAID Then Response.Write "active" End If %>">已支付</a>
                    <a href="/user/orders.asp?status=<%= PAYMENT_STATUS_FAILED %>&<% If Request.QueryString("search") <> "" Then Response.Write "search=" & Server.URLEncode(Request.QueryString("search")) & "&" End If %><% If Request.QueryString("time_range") <> "" Then Response.Write "time_range=" & Request.QueryString("time_range") & "&" End If %><% If Request.QueryString("start_date") <> "" Then Response.Write "start_date=" & Request.QueryString("start_date") & "&" End If %><% If Request.QueryString("end_date") <> "" Then Response.Write "end_date=" & Request.QueryString("end_date") & "&" End If %><% If Request.QueryString("payment_method") <> "" Then Response.Write "payment_method=" & Request.QueryString("payment_method") & "&" End If %>" class="<% If orderStatusFilter = PAYMENT_STATUS_FAILED Then Response.Write "active" End If %>">支付失败</a>
                </div>
                
                <!-- 调试信息 -->
                <% If Request.QueryString("debug") = "1" Then %>
                <div style="background:#f0f0f0;padding:10px;margin:10px 0;border:1px solid #ccc;">
                    <p><strong>调试信息:</strong></p>
                    <p>UserID: <%= userId %></p>
                    <p>SQL: <%= HTMLEncode(Session("LastOrdersSQL")) %></p>
                    <p>rsOrders Is Nothing: <%= (rsOrders Is Nothing) %></p>
                    <% If Not rsOrders Is Nothing Then %>
                    <p>rsOrders.EOF: <%= rsOrders.EOF %></p>
                    <% End If %>
                    <p>LastDBError: <%= HTMLEncode(Session("LastDBError")) %></p>
                </div>
                <% End If %>
                
                <%
                If Not rsOrders Is Nothing Then
                    If Not rsOrders.EOF Then
                %>
                <div class="orders-list">
                    <%
                    Dim orderStatusForReview, hasReviewId, reviewStatusVal
                    Do While Not rsOrders.EOF
                    %>
                    <div class="order-item">
                        <div class="order-header">
                            <span class="order-no">订单号: <%= rsOrders("OrderNo") %></span>
                            <span class="order-date"><%= rsOrders("CreatedAt") %></span>
                            <span class="order-status status-<%= rsOrders("Status") %>">
                                <% Select Case rsOrders("Status")
                                    Case "Pending"
                                        Response.Write "待支付"
                                    Case "Paid"
                                        Response.Write "已支付"
                                    Case "Failed"
                                        Response.Write "支付失败"
                                    Case "Refunded"
                                        Response.Write "已退款"
                                    Case Else
                                        Response.Write "未知状态"
                                End Select %>
                            </span>
                        </div>
                        
                        <div class="order-body">
                            <div class="order-summary">
                                <div class="order-amount">
                                    <span>金额:</span>
                                    <strong><%= FormatMoney(rsOrders("TotalAmount")) %></strong>
                                </div>
                                <div class="order-products">
                                    <span>商品:</span>
                                    <span>
                                        <%
                                        ' 解析Notes字段获取商品信息
                                        Dim orderNotes, productSummary
                                        orderNotes = rsOrders("Notes")
                                        
                                        If InStr(orderNotes, "商品详情: ") > 0 Then
                                            ' 提取商品详情部分
                                            Dim notesParts
                                            notesParts = Split(orderNotes, " | 支付流水号:")
                                            productSummary = Trim(notesParts(0))
                                            If InStr(productSummary, "商品详情: ") = 1 Then
                                                productSummary = Mid(productSummary, 6)  ' 移除"商品详情: "前缀
                                            End If
                                            
                                            ' 如果有多个商品，只显示第一个商品的概要
                                            Dim productItems
                                            productItems = Split(productSummary, "|")
                                            If UBound(productItems) >= 0 Then
                                                Dim firstItem
                                                firstItem = Trim(productItems(0))
                                                
                                                ' 提取商品名称部分（在括号前）
                                                Dim namePart
                                                Dim openParenPos
                                                openParenPos = InStr(firstItem, "(")
                                                If openParenPos > 0 Then
                                                    namePart = Trim(Left(firstItem, openParenPos - 1))
                                                Else
                                                    namePart = firstItem
                                                End If
                                                
                                                ' 如果有定制信息，也显示一部分
                                                Dim customPart
                                                Dim bracketStart
                                                bracketStart = InStr(namePart, " [")
                                                If bracketStart > 0 Then
                                                    ' 显示商品名称和部分定制信息
                                                    Dim shortCustom
                                                    shortCustom = Mid(namePart, bracketStart + 2)  ' 跳过" ["
                                                    shortCustom = Left(shortCustom, Len(shortCustom) - 1)  ' 去掉"]"
                                                    
                                                    ' 只显示关键定制信息的概要（前调、容量等）
                                                    Dim customParts, customPreview
                                                    customParts = Split(shortCustom, ", ")
                                                    customPreview = ""
                                                    Dim j
                                                    For j = 0 To UBound(customParts)
                                                        If InStr(customParts(j), "前调:") > 0 Or _
                                                           InStr(customParts(j), "ml") > 0 Or _
                                                           InStr(customParts(j), "瓶身:") > 0 Then
                                                            If customPreview <> "" Then customPreview = customPreview & ", "
                                                            customPreview = customPreview & customParts(j)
                                                        End If
                                                        ' 只取前几个关键信息，避免过长
                                                        If j >= 2 Then Exit For
                                                    Next
                                                    
                                                    If customPreview <> "" Then
                                                        Response.Write HTMLEncode(Trim(Left(namePart, bracketStart - 1)) & " (" & customPreview & ")")
                                                    Else
                                                        Response.Write HTMLEncode(Left(firstItem, 50))
                                                    End If
                                                Else
                                                    Response.Write HTMLEncode(Left(firstItem, 50))
                                                End If
                                            Else
                                                Response.Write HTMLEncode(Left(productSummary, 30))
                                            End If
                                        Else
                                            ' 如果没有商品详情，显示原始内容或提示
                                            If Len(orderNotes) > 0 Then
                                                Response.Write HTMLEncode(Left(orderNotes, 30))
                                            Else
                                                Response.Write "-"
                                            End If
                                        End If
                                        %>
                                    </span>
                                </div>
                            </div>
                            
                            <div class="order-payment">
                                <span>支付方式:</span>
                                <span>
                                    <%
                                    Dim pmValue
                                    If IsNumeric(rsOrders("PaymentMethod")) Then
                                        pmValue = CInt(rsOrders("PaymentMethod"))
                                    Else
                                        pmValue = 0  ' 未知支付方式
                                    End If
                                    Select Case pmValue
                                        Case PAYMENT_METHOD_WECHAT
                                            Response.Write "微信支付"
                                        Case PAYMENT_METHOD_ALIPAY
                                            Response.Write "支付宝"
                                        Case PAYMENT_METHOD_PAYPAL
                                            Response.Write "PayPal"
                                        Case PAYMENT_METHOD_COD
                                            Response.Write "货到付款"
                                        Case Else
                                            Response.Write "未知"
                                    End Select %>
                                </span>
                            </div>
                            
                            <div class="order-contact">
                                <span>收货人:</span>
                                <span><%= HTMLEncode(rsOrders("FullName")) %></span>
                            </div>
                        </div>
                        
                        <div class="order-actions">
                            <% If rsOrders("Status") = "Pending" Then %>
                            <a href="/checkout.asp?order_id=<%= rsOrders("OrderID") %>" class="btn btn-primary">立即支付</a>
                            <% End If %> 
                            <a href="/user/order_detail.asp?order_id=<%= rsOrders("OrderID") %>" class="btn btn-outline">查看详情</a>
                            
                            ' 评价状态显示
                            <% 
                            orderStatusForReview = rsOrders("Status")
                            hasReviewId = Not IsNull(rsOrders("ReviewID"))
                            reviewStatusVal = ""
                            If hasReviewId Then
                                reviewStatusVal = rsOrders("ReviewStatus")
                            End If
                            
                            ' 已评价显示标签
                            If hasReviewId Then
                            %>
                            <span class="review-badge reviewed">
                                <i class="fas fa-check-circle"></i> 已评价
                            </span>
                            <%
                            ' 未评价且订单状态允许评价（Paid或Completed）
                            ElseIf orderStatusForReview = "Paid" Or orderStatusForReview = "Completed" Then
                            %>
                            <a href="/user/order_detail.asp?order_id=<%= rsOrders("OrderID") %>#review-section" class="btn btn-review">
                                <i class="fas fa-star"></i> 去评价
                            </a>
                            <%
                            End If
                            %>
                            
                            <a href="javascript:void(0)" onclick="deleteOrder(this.getAttribute('data-order-id'))" data-order-id="<%= rsOrders("OrderID") %>" class="btn btn-danger btn-sm">删除</a>
                        </div>
                    </div>
                    <%
                    rsOrders.MoveNext
                    Loop
                    rsOrders.Close
                    Set rsOrders = Nothing
                    %>
                </div>
                <%
                    Else
                %>
                <div class="empty-orders">
                    <i class="fas fa-inbox"></i>
                    <h3>暂无订单</h3>
                    <p>您还没有任何订单记录</p>
                    <a href="/products.asp" class="btn btn-primary">去购物</a>
                </div>
                <%
                    End If
                End If
                %>
            </div>
        </div>
    </div>
</div>

<!--#include file="../includes/footer.asp"-->
<script>
// 搜索和筛选功能
$(document).ready(function() {
    // 搜索按钮点击事件
    $('#searchBtn').click(function() {
        var keyword = $('#searchInput').val();
        var url = window.location.pathname;
        var params = new URLSearchParams(window.location.search);
        
        if (keyword.trim() !== '') {
            params.set('search', keyword);
        } else {
            params.delete('search');
        }
        
        // 清除分页参数
        params.delete('page');
        
        window.location.href = url + '?' + params.toString();
    });
    
    // 回车键搜索
    $('#searchInput').keypress(function(e) {
        if (e.which === 13) {
            $('#searchBtn').click();
        }
    });
    
    // 高级搜索切换
    $('#advancedSearchToggle').click(function() {
        $('#advancedSearchPanel').slideToggle();
    });
    
    // 应用日期筛选
    $('#applyDateBtn').click(function() {
        var startDate = $('#startDate').val();
        var endDate = $('#endDate').val();
        var timeRange = $('#timeRange').val();
        
        var url = window.location.pathname;
        var params = new URLSearchParams(window.location.search);
        
        if (timeRange) {
            params.set('time_range', timeRange);
        } else {
            params.delete('time_range');
        }
        
        if (startDate) {
            params.set('start_date', startDate);
        } else {
            params.delete('start_date');
        }
        
        if (endDate) {
            params.set('end_date', endDate);
        } else {
            params.delete('end_date');
        }
        
        // 清除分页参数
        params.delete('page');
        
        window.location.href = url + '?' + params.toString();
    });
    
    // 高级搜索
    $('#advancedSearchBtn').click(function() {
        var status = $('select[name="status"]').val();
        var paymentMethod = $('select[name="payment_method"]').val();
        
        var url = window.location.pathname;
        var params = new URLSearchParams(window.location.search);
        
        if (status) {
            params.set('status', status);
        } else {
            params.delete('status');
        }
        
        if (paymentMethod) {
            params.set('payment_method', paymentMethod);
        } else {
            params.delete('payment_method');
        }
        
        // 清除分页参数
        params.delete('page');
        
        window.location.href = url + '?' + params.toString();
    });
    
    // 清空筛选
    $('#clearFiltersBtn').click(function() {
        window.location.href = window.location.pathname;
    });
});

// 删除订单函数
function deleteOrder(orderId) {
    if (confirm('确定要删除这个订单吗？此操作不可撤销。')) {
        // 发送AJAX请求删除订单
        $.ajax({
            url: '/user/delete_order.asp',
            type: 'POST',
            data: {
                orderId: orderId,
                csrf_token: csrfToken
            },
            dataType: 'json',
            success: function(response) {
                if (response.success) {
                    // 删除成功，刷新页面
                    alert('订单删除成功');
                    location.reload();
                } else {
                    alert('删除失败：' + response.message);
                }
            },
            error: function() {
                alert('删除请求失败，请稍后重试');
            }
        });
    }
}
</script>
<%
Call CloseConnection()
%>