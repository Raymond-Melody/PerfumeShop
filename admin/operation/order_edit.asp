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

' 获取订单ID
Dim orderId
orderId = Request.QueryString("id")

If orderId = "" Or Not IsNumeric(orderId) Then
    Response.Redirect "orders.asp"
    Response.End
End If

' 初始化变量
Dim orderNo, totalAmount, paymentMethod, orderStatus, createdAt, notes
Dim shippingName, shippingPhone, shippingAddress, customerName, customerEmail, customerPhone
Dim errorMsg, successMsg
errorMsg = ""
successMsg = ""

' 处理POST请求 - 更新订单
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' 验证CSRF令牌
    If Not ValidateCSRFToken() Then
        errorMsg = "安全验证失败，请刷新页面重试"
    Else
        ' 获取表单数据
        orderStatus = Trim(Request.Form("status"))
        shippingName = Trim(Request.Form("shipping_name"))
        shippingPhone = Trim(Request.Form("shipping_phone"))
        shippingAddress = Trim(Request.Form("shipping_address"))
        notes = Trim(Request.Form("notes"))
        
        ' 验证必填字段
        If orderStatus = "" Then
            errorMsg = "请选择订单状态"
        ElseIf shippingName = "" Then
            errorMsg = "请输入收货人姓名"
        ElseIf shippingPhone = "" Then
            errorMsg = "请输入收货电话"
        ElseIf shippingAddress = "" Then
            errorMsg = "请输入收货地址"
        Else
            ' 安全处理SQL字符串
            Dim safeStatus, safeShippingName, safeShippingPhone, safeShippingAddress, safeNotes
            safeStatus = SafeSQL(orderStatus)
            safeShippingName = SafeSQL(shippingName)
            safeShippingPhone = SafeSQL(shippingPhone)
            safeShippingAddress = SafeSQL(shippingAddress)
            safeNotes = SafeSQL(notes)
            
            ' 更新订单
            Dim sql, result
            sql = "UPDATE Orders SET " & _
                  "[Status] = '" & safeStatus & "', " & _
                  "ShippingName = '" & safeShippingName & "', " & _
                  "ShippingPhone = '" & safeShippingPhone & "', " & _
                  "ShippingAddress = '" & safeShippingAddress & "', " & _
                  "Notes = '" & safeNotes & "', " & _
                  "UpdatedAt = GETDATE() " & _
                  "WHERE OrderID = " & CLng(orderId)
            
            result = ExecuteNonQuery(sql)
            
            If result Then
                Call LogAdminAction("编辑订单", "operation", "Orders", orderId, "订单号: " & orderNo)
                Response.Redirect "order_detail.asp?order_id=" & orderId & "&msg=updated"
                Response.End
            Else
                errorMsg = "更新失败: " & Session("LastDBError")
            End If
        End If
    End If
End If

' 获取订单基本信息
Dim rsOrder
Set rsOrder = ExecuteQuery("SELECT o.*, u.Username, u.Email, u.Phone FROM Orders o LEFT JOIN Users u ON o.UserID = u.UserID WHERE o.OrderID = " & CLng(orderId))

If rsOrder Is Nothing Or rsOrder.EOF Then
    errorMsg = "订单不存在"
Else
    ' 存储订单信息到变量
    orderNo = rsOrder("OrderNo")
    totalAmount = rsOrder("TotalAmount")
    paymentMethod = rsOrder("PaymentMethod")
    orderStatus = rsOrder("Status")
    createdAt = rsOrder("CreatedAt")
    notes = rsOrder("Notes")
    shippingName = rsOrder("ShippingName")
    shippingPhone = rsOrder("ShippingPhone")
    shippingAddress = rsOrder("ShippingAddress")
    customerName = rsOrder("Username")
    customerEmail = rsOrder("Email")
    customerPhone = rsOrder("Phone")
    
    rsOrder.Close
End If
Set rsOrder = Nothing

' 获取订单商品列表
Dim rsDetails
Set rsDetails = ExecuteQuery("SELECT od.*, p.ProductType FROM OrderDetails od LEFT JOIN Products p ON od.ProductID = p.ProductID WHERE od.OrderID = " & CLng(orderId))

Call LogAdminAction("访问订单编辑页面", "operation", "Orders", orderId, "")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>编辑订单 #<%= SafeOutput(orderNo) %> - 管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        .form-container { max-width: 900px; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.08); }
        .form-group { margin-bottom: 20px; }
        .form-group label { display: block; margin-bottom: 8px; font-weight: 500; color: #333; }
        .form-group label .required { color: #e74c3c; margin-left: 4px; }
        .form-control { width: 100%; padding: 12px 15px; border: 1px solid #ddd; border-radius: 8px; font-size: 14px; box-sizing: border-box; }
        .form-control:focus { outline: none; border-color: #667eea; }
        textarea.form-control { min-height: 80px; resize: vertical; }
        select.form-control { height: 42px; }
        .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        .form-actions { display: flex; gap: 15px; margin-top: 30px; padding-top: 20px; border-top: 1px solid #f0f0f0; }
        .alert { padding: 15px; border-radius: 8px; margin-bottom: 20px; }
        .alert-error { background: #ffebee; color: #c62828; border: 1px solid #ffcdd2; }
        .alert-success { background: #e8f5e9; color: #2e7d32; border: 1px solid #c8e6c9; }
        
        .info-section { background: #f9f9f9; padding: 20px; border-radius: 8px; margin-bottom: 25px; border: 1px solid #eee; }
        .info-section h3 { margin-top: 0; margin-bottom: 15px; font-size: 16px; color: #333; border-bottom: 1px solid #ddd; padding-bottom: 10px; }
        .info-row { display: flex; margin-bottom: 10px; font-size: 14px; }
        .info-row .label { width: 100px; color: #666; font-weight: bold; }
        .info-row .value { flex: 1; color: #333; }
        
        .status-Pending { color: #f39c12; }
        .status-Paid { color: #27ae60; font-weight: bold; }
        .status-Processing { color: #3498db; }
        .status-Shipped { color: #9b59b6; }
        .status-Delivered { color: #2ecc71; }
        .status-Cancelled { color: #e74c3c; }
        .status-Refunded { color: #95a5a6; }
        
        .product-card { border: 1px solid #eee; border-radius: 8px; padding: 15px; margin-bottom: 15px; background: #fff; }
        .product-header { display: flex; justify-content: space-between; align-items: center; border-bottom: 1px dashed #eee; padding-bottom: 10px; margin-bottom: 10px; }
        .product-name { font-weight: bold; font-size: 16px; color: #007bff; }
        .product-info { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; font-size: 13px; color: #666; }
        
        .readonly-field { background: #f5f5f5; color: #666; cursor: not-allowed; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-edit"></i> 编辑订单 #<%= SafeOutput(orderNo) %></h2>
            <div class="breadcrumb">
                <a href="index.asp">运营中心</a> / <a href="orders.asp">订单管理</a> / <span>编辑订单</span>
            </div>
        </div>
        
        <% If errorMsg <> "" Then %>
        <div class="alert alert-error"><i class="fas fa-exclamation-circle"></i> <%= errorMsg %></div>
        <% End If %>
        
        <% If successMsg <> "" Then %>
        <div class="alert alert-success"><i class="fas fa-check-circle"></i> <%= successMsg %></div>
        <% End If %>
        
        <% If errorMsg = "订单不存在" Then %>
        <div class="form-container">
            <div class="alert alert-error">
                <i class="fas fa-exclamation-triangle"></i> 订单不存在或已被删除
            </div>
            <div class="form-actions">
                <a href="orders.asp" class="admin-btn admin-btn-secondary">
                    <i class="fas fa-arrow-left"></i> 返回订单列表
                </a>
            </div>
        </div>
        <% Else %>
        <div class="form-container">
            <form method="post" action="">
                <%= GetCSRFTokenField() %>
                
                <!-- 只读信息区域 -->
                <div class="info-section">
                    <h3><i class="fas fa-info-circle"></i> 订单基本信息（只读）</h3>
                    <div class="form-row">
                        <div class="info-row">
                            <span class="label">订单号:</span>
                            <span class="value"><%= SafeOutput(orderNo) %></span>
                        </div>
                        <div class="info-row">
                            <span class="label">下单时间:</span>
                            <span class="value"><%= createdAt %></span>
                        </div>
                    </div>
                    <div class="form-row">
                        <div class="info-row">
                            <span class="label">客户账号:</span>
                            <span class="value"><%= SafeOutput(customerName & "") %> (<%= SafeOutput(customerEmail & "") %>)</span>
                        </div>
                        <div class="info-row">
                            <span class="label">订单金额:</span>
                            <span class="value"><%= FormatMoney(CDbl("0" & totalAmount)) %></span>
                        </div>
                    </div>
                </div>
                
                <!-- 可编辑区域 -->
                <div class="form-row">
                    <div class="form-group">
                        <label>订单状态 <span class="required">*</span></label>
                        <select name="status" class="form-control" required>
                            <option value="Pending" <%= IIF(orderStatus="Pending", "selected", "") %>>待支付</option>
                            <option value="Paid" <%= IIF(orderStatus="Paid", "selected", "") %>>已支付</option>
                            <option value="Processing" <%= IIF(orderStatus="Processing", "selected", "") %>>处理中</option>
                            <option value="Shipped" <%= IIF(orderStatus="Shipped", "selected", "") %>>已发货</option>
                            <option value="Delivered" <%= IIF(orderStatus="Delivered", "selected", "") %>>已送达</option>
                            <option value="Cancelled" <%= IIF(orderStatus="Cancelled", "selected", "") %>>已取消</option>
                            <option value="Refunded" <%= IIF(orderStatus="Refunded", "selected", "") %>>已退款</option>
                        </select>
                    </div>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>收货人姓名 <span class="required">*</span></label>
                        <input type="text" name="shipping_name" class="form-control" value="<%= SafeOutput(shippingName) %>" required>
                    </div>
                    <div class="form-group">
                        <label>收货电话 <span class="required">*</span></label>
                        <input type="text" name="shipping_phone" class="form-control" value="<%= SafeOutput(shippingPhone) %>" required>
                    </div>
                </div>
                
                <div class="form-group">
                    <label>收货地址 <span class="required">*</span></label>
                    <textarea name="shipping_address" class="form-control" required><%= SafeOutput(shippingAddress) %></textarea>
                </div>
                
                <div class="form-group">
                    <label>订单备注</label>
                    <textarea name="notes" class="form-control" placeholder="请输入订单备注（可选）"><%= SafeOutput(notes & "") %></textarea>
                </div>
                
                <!-- 订单商品列表（只读） -->
                <div class="info-section" style="margin-top: 30px;">
                    <h3><i class="fas fa-shopping-bag"></i> 订单商品（只读）</h3>
                    <% 
                    If Not rsDetails Is Nothing Then
                        If Not rsDetails.EOF Then
                            Do While Not rsDetails.EOF
                    %>
                    <div class="product-card">
                        <div class="product-header">
                            <span class="product-name"><%= SafeOutput(rsDetails("ProductName") & "") %></span>
                            <span>数量: <strong><%= rsDetails("Quantity") %></strong> | 单价: <%= FormatMoney(CDbl("0" & rsDetails("UnitPrice"))) %></span>
                        </div>
                        <div class="product-info">
                            <div>前调: <%= SafeOutput(rsDetails("TopNoteName") & "") %></div>
                            <div>中调: <%= SafeOutput(rsDetails("MiddleNoteName") & "") %></div>
                            <div>后调: <%= SafeOutput(rsDetails("BaseNoteName") & "") %></div>
                            <div>容量: <%= SafeOutput(rsDetails("VolumeName") & "") %></div>
                            <div>瓶身: <%= SafeOutput(rsDetails("BottleName") & "") %></div>
                            <div>刻字: <%= SafeOutput(rsDetails("CustomLabel") & "") %></div>
                        </div>
                    </div>
                    <%
                                rsDetails.MoveNext
                            Loop
                            rsDetails.Close
                        Else
                            Response.Write "<div class='text-muted'>此订单无详细商品数据</div>"
                        End If
                        Set rsDetails = Nothing
                    End If
                    %>
                </div>
                
                <div class="form-actions">
                    <button type="submit" class="admin-btn admin-btn-primary">
                        <i class="fas fa-save"></i> 保存修改
                    </button>
                    <a href="order_detail.asp?order_id=<%= orderId %>" class="admin-btn admin-btn-secondary">
                        <i class="fas fa-times"></i> 取消
                    </a>
                </div>
            </form>
        </div>
        <% End If %>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
