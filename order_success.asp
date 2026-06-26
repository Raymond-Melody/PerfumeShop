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

' V14: 会员登录检查
If Session("UserID") = "" Or IsNull(Session("UserID")) Then
    Response.Redirect "/user/login.asp?return=" & Server.URLEncode(Request.ServerVariables("SCRIPT_NAME") & "?" & Request.ServerVariables("QUERY_STRING"))
    Response.End
End If

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
        <a href="/index.asp"><% If FEATURE_I18N Then %><%= T("breadcrumb_home", Empty) %><% Else %>首页<% End If %></a>
        <span class="separator">/</span>
        <a href="/cart.asp"><% If FEATURE_I18N Then %><%= T("cart_title", Empty) %><% Else %>购物车<% End If %></a>
        <span class="separator">/</span>
        <span><% If FEATURE_I18N Then %><%= T("order_breadcrumb_success", Empty) %><% Else %>订单成功<% End If %></span>
    </div>
</div>

<div class="container">
    <div class="order-success-page">
        <div class="success-card">
            <div class="success-icon">
                <i class="fas fa-check-circle"></i>
            </div>
            <h2><% If FEATURE_I18N Then %><%= T("order_success_title", Empty) %><% Else %>订单提交成功！<% End If %></h2>
            <p><% If FEATURE_I18N Then %><%= T("order_success_msg", Empty) %><% Else %>您的订单已成功提交，我们会尽快为您处理<% End If %></p>
            
            <div class="order-details">
                <div class="detail-row">
                    <span><% If FEATURE_I18N Then %><%= T("order_no", Empty) %><% Else %>订单号<% End If %>:</span>
                    <span><%= orderNo %></span>
                </div>
                
                <div class="detail-row">
                    <span><% If FEATURE_I18N Then %><%= T("order_amount", Empty) %><% Else %>订单金额<% End If %>:</span>
                    <span class="amount"><%= FormatMoney(orderAmount) %></span>
                </div>
                
                <div class="detail-row">
                    <span><% If FEATURE_I18N Then %><%= T("order_payment_method", Empty) %><% Else %>支付方式<% End If %>:</span>
                    <span>
                        <%
                        Dim pmVal
                        If IsNumeric(paymentMethod) Then pmVal = CInt(paymentMethod) Else pmVal = -1
                        Select Case pmVal
                            Case PAYMENT_METHOD_WECHAT
                                Response.Write T("payment_wechat", Empty)
                            Case PAYMENT_METHOD_ALIPAY
                                Response.Write T("payment_alipay", Empty)
                            Case PAYMENT_METHOD_PAYPAL
                                Response.Write "PayPal"
                            Case PAYMENT_METHOD_COD
                                Response.Write T("payment_cod", Empty)
                            Case Else
                                Response.Write T("payment_unknown", Empty)
                        End Select %>
                    </span>
                </div>
                
                <div class="detail-row">
                    <span><% If FEATURE_I18N Then %><%= T("order_status", Empty) %><% Else %>订单状态<% End If %>:</span>
                    <span class="status-<%= orderStatus %>">
                        <%
                        Select Case Trim(orderStatus & "")
                            Case "0", "Pending"
                                Response.Write T("status_pending", Empty)
                            Case "1", "Paid"
                                Response.Write T("status_paid", Empty)
                            Case "2", "Failed"
                                Response.Write T("status_failed", Empty)
                            Case "3", "Refunded"
                                Response.Write T("status_refunded", Empty)
                            Case Else
                                Response.Write T("status_unknown", Empty)
                        End Select %>
                    </span>
                </div>
                
                <div class="detail-row">
                    <span><% If FEATURE_I18N Then %><%= T("order_time", Empty) %><% Else %>下单时间<% End If %>:</span>
                    <span><%= createdAt %></span>
                </div>
            </div>
            
            <div class="success-actions">
                <a href="/user/orders.asp" class="btn btn-primary"><% If FEATURE_I18N Then %><%= T("order_btn_view", Empty) %><% Else %>查看订单<% End If %></a>
                <a href="/products.asp" class="btn btn-outline"><% If FEATURE_I18N Then %><%= T("order_btn_continue", Empty) %><% Else %>继续购物<% End If %></a>
            </div>
        </div>
        
        <div class="order-tips">
            <h3><% If FEATURE_I18N Then %><%= T("order_tips_title", Empty) %><% Else %>温馨提示<% End If %></h3>
            <ul>
                <li><i class="fas fa-check"></i> <% If FEATURE_I18N Then %><%= T("order_tips_1", Empty) %><% Else %>订单提交后，我们会在24小时内处理<% End If %></li>
                <li><i class="fas fa-check"></i> <% If FEATURE_I18N Then %><%= T("order_tips_2", Empty) %><% Else %>如需发票，请联系客服<% End If %></li>
                <li><i class="fas fa-check"></i> <% If FEATURE_I18N Then %><%= T("order_tips_3", Empty) %><% Else %>如有任何问题，请拨打客服电话 <% End If %><%= SITE_PHONE %></li>
            </ul>
        </div>
    </div>
</div>

<!--#include file="includes/footer.asp"-->
<%
Call CloseConnection()
%>