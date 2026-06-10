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

Dim cartId
Dim sessionId, userId
Dim whereClause

cartId = Request.Form("cartId")
sessionId = Session.SessionID
userId = Session("UserID")

' 验证参数
If cartId = "" Or Not IsNumeric(cartId) Then
    Response.Write "{""success"":false,""message"":""无效的购物车项""}"
    Response.End
End If

' 构建权限检查条件
If userId <> "" Then
    whereClause = "CartID = " & CInt(cartId) & " AND UserID = " & userId
Else
    whereClause = "CartID = " & CInt(cartId) & " AND SessionID = '" & SafeSQL(sessionId) & "'"
End If

' 删除购物车项
Dim sql
sql = "DELETE FROM Cart WHERE " & whereClause

If ExecuteNonQuery(sql) Then
    Response.Write "{""success"":true}"
Else
    Response.Write "{""success"":false,""message"":""删除失败""}"
End If

Call CloseConnection()
%>
