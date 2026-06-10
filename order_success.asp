<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/payment_config.asp"-->
<%
Call OpenConnection()

Dim orderId
orderId = Request.QueryString("order_id")

If orderId = "" Or Not IsNumeric(orderId) Then
    Response.Redirect "/cart.asp"
    Response.End
End If

' 获取订单信息
Dim rsOrder, orderInfo
Set rsOrder = ExecuteQuery("SELECT o.*, u.Username FROM Orders o LEFT JOIN Users u ON o.UserID = u.UserID WHERE o.OrderID = " & orderId)

If rsOrder Is Nothing Or rsOrder.EOF Then
    Response.Redirect "/cart.asp"
    Response.End
End If

orderInfo = rsOrder

' 复制订单信息到变量，因为RecordSet即将关闭
Dim orderNo, orderAmount, paymentMethod, orderStatus, createdAt
orderNo = orderInfo("OrderNo")
orderAmount = orderInfo("TotalAmount")  ' 注意：应该是TotalAmount，不是OrderAmount'
paymentMethod = orderInfo("PaymentMethod")
orderStatus = orderInfo("Status")  ' 注意：应该是Status，不是OrderStatus'
createdAt = orderInfo("CreatedAt")

rsOrder.Close
Set rsOrder = Nothing
%>
<!--#include file="includes/header.asp"-->

<!-- 面包屑导航 -->
<div class="breadcrumb">
    <div class="container">
        <a href="/index.asp">首页</a>
        <span class="separator">/</span>
        <a href="/cart.asp">购物车</a>
        <span class="separator">/</span>
        <span>订单成功</span>
    </div>
</div>

<div class="container">
    <div class="order-success-page">
        <div class="success-card">
            <div class="success-icon">
                <i class="fas fa-check-circle"></i>
            </div>
            <h2>订单提交成功！</h2>
            <p>您的订单已成功提交，我们会尽快为您处理</p>
            
            <div class="order-details">
                <div class="detail-row">
                    <span>订单号:</span>
                    <span><%= orderNo %></span>
                </div>
                
                <div class="detail-row">
                    <span>订单金额:</span>
                    <span class="amount"><%= FormatMoney(orderAmount) %></span>
                </div>
                
                <div class="detail-row">
                    <span>支付方式:</span>
                    <span>
                        <%
                        Dim pmVal
                        If IsNumeric(paymentMethod) Then pmVal = CInt(paymentMethod) Else pmVal = -1
                        Select Case pmVal
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
                
                <div class="detail-row">
                    <span>订单状态:</span>
                    <span class="status-<%= orderStatus %>">
                        <%
                        Select Case Trim(orderStatus & "")
                            Case "0", "Pending"
                                Response.Write "待支付"
                            Case "1", "Paid"
                                Response.Write "已支付"
                            Case "2", "Failed"
                                Response.Write "支付失败"
                            Case "3", "Refunded"
                                Response.Write "已退款"
                            Case Else
                                Response.Write "未知状态"
                        End Select %>
                    </span>
                </div>
                
                <div class="detail-row">
                    <span>下单时间:</span>
                    <span><%= createdAt %></span>
                </div>
            </div>
            
            <div class="success-actions">
                <a href="/user/orders.asp" class="btn btn-primary">查看订单</a>
                <a href="/products.asp" class="btn btn-outline">继续购物</a>
            </div>
        </div>
        
        <div class="order-tips">
            <h3>温馨提示</h3>
            <ul>
                <li><i class="fas fa-check"></i> 订单提交后，我们会在24小时内处理</li>
                <li><i class="fas fa-check"></i> 如需发票，请联系客服</li>
                <li><i class="fas fa-check"></i> 如有任何问题，请拨打客服电话 <%= SITE_PHONE %></li>
            </ul>
        </div>
    </div>
</div>

<!--#include file="includes/footer.asp"-->
<%
Call CloseConnection()
%>