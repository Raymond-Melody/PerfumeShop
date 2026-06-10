<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "application/json"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<%
Call OpenConnection()

' 返回JSON响应的函数
Function JsonResponse(success, message)
    JsonResponse = "{""success"": " & success & ", ""message"": """ & Replace(message, """", """"") & """}"
End Function

' CSRF验证
If Not ValidateCSRFToken() Then
    Response.Status = "403 Forbidden"
    Response.Write JsonResponse(False, "安全验证失败，请刷新页面重试")
    Response.End
End If

' 检查用户是否登录
If Session("UserID") = "" Then
    Response.Write JsonResponse(False, "请先登录")
    Response.End
End If

Dim userId, orderId
userId = Session("UserID")
orderId = Trim(Request.Form("orderId"))

' 验证参数
If orderId = "" Or Not IsNumeric(orderId) Then
    Response.Write JsonResponse(False, "无效的订单ID")
    Response.End
End If

' 验证订单归属（只能删除自己的订单）
Dim checkSql, rsCheck
checkSql = "SELECT OrderID FROM Orders WHERE OrderID = " & orderId & " AND UserID = " & userId
Set rsCheck = ExecuteQuery(checkSql)

If rsCheck Is Nothing Or rsCheck.EOF Then
    Response.Write JsonResponse(False, "订单不存在或无权删除")
    Response.End
End If

rsCheck.Close
Set rsCheck = Nothing

' 执行删除操作（这里使用软删除，将订单状态改为"Deleted"）
Dim deleteSql
deleteSql = "UPDATE Orders SET Status = 'Deleted', UpdatedAt = GETDATE() WHERE OrderID = " & orderId & " AND UserID = " & userId

On Error Resume Next
ExecuteNonQuery(deleteSql)
If Err.Number <> 0 Then
    Response.Write JsonResponse(False, "删除失败：" & Err.Description)
    Response.End
End If
On Error Goto 0

Response.Write JsonResponse(True, "订单删除成功")

Call CloseConnection()
%>