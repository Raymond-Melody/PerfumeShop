<%@ Language="VBScript" CodePage="65001" %>
<%
' 支付回调处理页面
' 处理微信支付、支付宝、PayPal等支付平台的异步通知

Response.ContentType = "text/plain"
Response.Charset = "UTF-8"

' 读取POST数据
Dim callbackData, callbackString
callbackData = Request.Form
callbackString = ""

Dim formKey
For Each formKey In Request.Form
    callbackString = callbackString & formKey & "=" & Request.Form(formKey) & "&"
Next

' 移除最后的&符号
If Len(callbackString) > 0 Then
    callbackString = Left(callbackString, Len(callbackString) - 1)
End If

%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/payment_handler.asp"-->
<!--#include file="includes/cost_engine.asp"-->
<%

Call OpenConnection()

' 简单的日志记录（实际应用中应记录到文件或数据库）
Response.Write("支付回调接收到数据: " & callbackString & vbCrLf)

' 检测支付平台
Dim paymentMethod, orderNo, transactionId, verifyResult

' 这里需要根据实际的回调数据格式来检测支付平台
' 微信支付通常有: transaction_id, out_trade_no 等字段
If InStr(LCase(callbackString), "transaction_id") > 0 And InStr(LCase(callbackString), "out_trade_no") > 0 Then
    paymentMethod = PAYMENT_METHOD_WECHAT
    orderNo = Request.Form("out_trade_no")  ' 微信支付商户订单号
    transactionId = Request.Form("transaction_id")  ' 微信支付交易号
ElseIf InStr(LCase(callbackString), "trade_no") > 0 And InStr(LCase(callbackString), "out_trade_no") > 0 Then
    paymentMethod = PAYMENT_METHOD_ALIPAY
    orderNo = Request.Form("out_trade_no")  ' 支付宝商户订单号
    transactionId = Request.Form("trade_no")  ' 支付宝交易号
ElseIf InStr(LCase(callbackString), "payment_id") > 0 And InStr(LCase(callbackString), "invoice_number") > 0 Then
    paymentMethod = PAYMENT_METHOD_PAYPAL
    orderNo = Request.Form("invoice_number")  ' PayPal发票号（对应我们的订单号）
    transactionId = Request.Form("payment_id")  ' PayPal交易ID
Else
    ' 如果无法识别支付平台，尝试从自定义参数中获取
    orderNo = Request.Form("order_no")
    paymentMethod = Request.Form("payment_method")
    transactionId = Request.Form("transaction_id")
End If

' 验证支付回调
verifyResult = VerifyPaymentCallback(paymentMethod, callbackString)

If verifyResult And orderNo <> "" And transactionId <> "" Then
    ' 获取订单信息
    Dim rsOrder, orderId
    Set rsOrder = ExecuteQuery("SELECT OrderID FROM Orders WHERE OrderNo = '" & SafeSQL(orderNo) & "'")
    
    If Not rsOrder Is Nothing Then
        If Not rsOrder.EOF Then
            orderId = rsOrder("OrderID")
            
            ' 更新订单状态为已支付
            Call UpdateOrderPaymentStatus(orderId, PAYMENT_STATUS_PAID, transactionId)
            
            ' 成本自动传导：支付成功后重新计算订单成本和利润
            Call CE_UpdateOrderCosts(orderId)
            
            ' 返回成功响应（根据支付平台要求）
            ' 微信支付需要返回<xml><return_code><![CDATA[SUCCESS]]></return_code><return_msg><![CDATA[OK]]></return_msg></xml>
            If paymentMethod = PAYMENT_METHOD_WECHAT Then
                Response.ContentType = "application/xml"
                Response.Write("<xml><return_code><![CDATA[SUCCESS]]></return_code><return_msg><![CDATA[OK]]></return_msg></xml>")
            Else
                Response.Write("success")
            End If
            Response.End
        Else
            Response.Write("订单不存在: " & orderNo)
        End If
        rsOrder.Close
        Set rsOrder = Nothing
    Else
        Response.Write("查询订单失败")
    End If
Else
    Response.Write("支付验证失败")
End If

Call CloseConnection()
%>