<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "application/json"

If Session("UserID") = "" Then
    Response.Write "{""success"":false,""message"":""请先登录""}"
    Response.End
End If
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<%
Call OpenConnection()

' CSRF验证
If Not ValidateCSRFToken() Then
    Response.Status = "403 Forbidden"
    Response.Write "{""success"":false,""message"":""安全验证失败，请刷新页面重试""}"
    Response.End
End If

Dim orderId, userId
orderId = Request.Form("orderId")
userId = Session("UserID")

If orderId = "" Or Not IsNumeric(orderId) Then
    Response.Write "{""success"":false,""message"":""无效的订单""}"
    Response.End
End If

' 检查订单归属和状态
Dim currentStatus
currentStatus = GetScalar("SELECT Status FROM Orders WHERE OrderID = " & CInt(orderId) & " AND UserID = " & userId)

If IsNull(currentStatus) Or currentStatus = "" Then
    Response.Write "{""success"":false,""message"":""订单不存在""}"
    Response.End
End If

If currentStatus <> "Shipped" Then
    Response.Write "{""success"":false,""message"":""只能确认已发货订单""}"
    Response.End
End If

' 确认收货
Dim sql
sql = "UPDATE Orders SET Status = 'Delivered', UpdatedAt = GETDATE() WHERE OrderID = " & CInt(orderId) & " AND UserID = " & userId

If ExecuteNonQuery(sql) Then
    Response.Write "{""success"":true}"
Else
    Response.Write "{""success"":false,""message"":""操作失败""}"
End If

Call CloseConnection()
%>
