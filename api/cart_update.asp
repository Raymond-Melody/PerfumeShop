<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "application/json"
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

Dim cartId, delta, newQty
Dim sessionId, userId
Dim whereClause

cartId = Request.Form("cartId")
delta = Request.Form("delta")
sessionId = Session.SessionID
userId = Session("UserID")

' 验证参数
If cartId = "" Or Not IsNumeric(cartId) Then
    Response.Write "{""success"":false,""message"":""无效的购物车项""}"
    Response.End
End If

If delta = "" Or Not IsNumeric(delta) Then
    delta = 0
End If
delta = CInt(delta)

' 构建权限检查条件
If userId <> "" Then
    whereClause = "CartID = " & CInt(cartId) & " AND UserID = " & userId
Else
    whereClause = "CartID = " & CInt(cartId) & " AND SessionID = '" & SafeSQL(sessionId) & "'"
End If

' 获取当前数量
Dim currentQty
currentQty = GetScalar("SELECT Quantity FROM Cart WHERE " & whereClause)
If IsNull(currentQty) Or currentQty = "" Then
    Response.Write "{""success"":false,""message"":""购物车项不存在""}"
    Response.End
End If

newQty = CInt(currentQty) + delta
If newQty < 1 Then newQty = 1
If newQty > 99 Then newQty = 99

' 更新数量
Dim sql
sql = "UPDATE Cart SET Quantity = " & newQty & " WHERE " & whereClause

If ExecuteNonQuery(sql) Then
    Response.Write "{""success"":true,""quantity"":" & newQty & "}"
Else
    Response.Write "{""success"":false,""message"":""更新失败""}"
End If

Call CloseConnection()
%>
