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

Dim sessionId, userId
Dim whereClause

sessionId = Session.SessionID
userId = Session("UserID")

' 构建权限检查条件
If userId <> "" Then
    whereClause = "UserID = " & userId
Else
    whereClause = "SessionID = '" & SafeSQL(sessionId) & "'"
End If

' 清空购物车
Dim sql
sql = "DELETE FROM Cart WHERE " & whereClause

If ExecuteNonQuery(sql) Then
    Response.Write "{""success"":true}"
Else
    Response.Write "{""success"":false,""message"":""操作失败""}"
End If

Call CloseConnection()
%>
