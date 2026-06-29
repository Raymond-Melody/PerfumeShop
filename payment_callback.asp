<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' V18.0 支付回调处理页面 (增强版)
' 处理微信支付、支付宝、PayPal等支付平台的异步通知
' 新增: IP白名单校验 + 回调签名二次验证
' ============================================

Response.ContentType = "text/plain"
Response.Charset = "UTF-8"

' V18: 记录回调请求日志（签名验证前）
Dim callbackIP
callbackIP = Request.ServerVariables("REMOTE_ADDR")

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
<!--#include file="includes/audit_utils.asp"-->
<!--#include file="includes/cost_engine.asp"-->
<%

Call OpenConnection()

' ============================================
' V18: IP 白名单校验
' ============================================
If Not IsCallbackIPAllowed(callbackIP) Then
    ' 记录可疑回调到审计日志
    If FEATURE_GDPR_COMPLIANCE Then
        Call LogPrivacyAction(AUDIT_ACTION_PRIVACY_ACCESS, 0, "SYSTEM", _
            "PaymentCallback blocked: unauthorized IP " & callbackIP & " | Data: " & Left(callbackString, 200))
    End If
    Response.Write("IP not allowed: " & callbackIP)
    Call CloseConnection()
    Response.End
End If

' ============================================
' V18: 回调签名验证（二次校验）
' ============================================
' 简单的日志记录（实际应用中应记录到文件或数据库）
Response.Write("支付回调接收到数据: " & callbackString & vbCrLf)

' 检测支付平台
Dim paymentMethod, orderNo, transactionId, verifyResult, signatureValid

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

' 验证支付回调签名（V18增强）
Dim callbackVerifyResult
callbackVerifyResult = VerifyCallbackSignature(paymentMethod, callbackString)

' 原有验证（兼容层）
verifyResult = VerifyPaymentCallback(paymentMethod, callbackString)

' V18: 双重验证 - 签名验证 + 原有验证
signatureValid = callbackVerifyResult And verifyResult

If signatureValid And orderNo <> "" And transactionId <> "" Then
    ' 获取订单信息
    Dim rsOrder, orderId
    Set rsOrder = ExecuteQuery("SELECT OrderID FROM Orders WHERE OrderNo = '" & SafeSQL(orderNo) & "'")
    
    If Not rsOrder Is Nothing Then
        If Not rsOrder.EOF Then
            orderId = rsOrder("OrderID")
            
            ' V18: 幂等检查 - 防止回调重复处理
            If CheckPaymentIdempotency(orderId) Then
                ' 已经处理过，但仍然返回成功响应
                If CLng(paymentMethod) = PAYMENT_METHOD_WECHAT Then
                    Response.ContentType = "application/xml"
                    Response.Write("<xml><return_code><![CDATA[SUCCESS]]></return_code><return_msg><![CDATA[OK]]></return_msg></xml>")
                Else
                    Response.Write("success (idempotent)")
                End If
                rsOrder.Close
                Set rsOrder = Nothing
                Call CloseConnection()
                Response.End
            End If
            
            ' 更新订单状态为已支付
            Call UpdateOrderPaymentStatus(orderId, PAYMENT_STATUS_PAID, transactionId)
            
            ' 成本自动传导：支付成功后重新计算订单成本和利润
            Call CE_UpdateOrderCosts(orderId)
            
            ' V18: 记录支付回调审计
            If FEATURE_GDPR_COMPLIANCE Then
                Call LogPrivacyAction(AUDIT_ACTION_PRIVACY_ACCESS, 0, "SYSTEM", _
                    "PaymentCallback success: OrderID=" & orderId & " TransID=" & transactionId & " Method=" & paymentMethod & " IP=" & callbackIP)
            End If
            
            ' 返回成功响应（根据支付平台要求）
            ' 微信支付需要返回<xml><return_code><![CDATA[SUCCESS]]></return_code><return_msg><![CDATA[OK]]></return_msg></xml>
            If CLng(paymentMethod) = PAYMENT_METHOD_WECHAT Then
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
    ' V18: 记录验证失败到审计日志
    If FEATURE_GDPR_COMPLIANCE Then
        Call LogPrivacyAction(AUDIT_ACTION_PRIVACY_ACCESS, 0, "SYSTEM", _
            "PaymentCallback FAILED: signatureValid=" & signatureValid & " Method=" & paymentMethod & " IP=" & callbackIP & " Data=" & Left(callbackString, 200))
    End If
    Response.Write("支付验证失败")
End If

Call CloseConnection()
%>