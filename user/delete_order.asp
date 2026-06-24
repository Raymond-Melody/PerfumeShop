<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "application/json"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<%
Call OpenConnection()

Function JsonResponse(success, message)
    JsonResponse = "{""success"": " & success & ", ""message"": """ & Replace(message, """", """""") & """}"
End Function

If Not ValidateCSRFToken() Then
    Response.Write JsonResponse(False, "安全验证失败，请刷新页面重试")
    Call CloseConnection()
    Response.End
End If

If Session("UserID") = "" Then
    Response.Write JsonResponse(False, "请先登录")
    Call CloseConnection()
    Response.End
End If

Dim userId, orderId
userId = Session("UserID")
orderId = Trim(Request.Form("orderId"))

If orderId = "" Or Not IsNumeric(orderId) Then
    Response.Write JsonResponse(False, "无效的订单ID")
    Call CloseConnection()
    Response.End
End If

Dim checkSql, rsCheck
checkSql = "SELECT OrderID FROM Orders WHERE OrderID = " & orderId & " AND UserID = " & userId
Set rsCheck = ExecuteQuery(checkSql)

If rsCheck Is Nothing Or rsCheck.EOF Then
    Response.Write JsonResponse(False, "订单不存在或无权删除")
    Call CloseConnection()
    Response.End
End If

rsCheck.Close
Set rsCheck = Nothing

Dim deleteSql
deleteSql = "UPDATE Orders SET Status = 'Deleted', UpdatedAt = GETDATE() WHERE OrderID = " & orderId & " AND UserID = " & userId

On Error Resume Next
conn.Execute deleteSql
If Err.Number <> 0 Then
    Dim errMsg
    errMsg = Err.Description
    Err.Clear
    On Error GoTo 0
    Response.Write JsonResponse(False, "删除失败：" & errMsg)
    Call CloseConnection()
    Response.End
End If
On Error GoTo 0

Response.Write JsonResponse(True, "订单删除成功")
Call CloseConnection()
%>
