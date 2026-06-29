<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/api_response.asp"-->
<!--#include file="../includes/api_guard.asp"-->
<%
Call OpenConnection()

' V18: API 守卫（速率限制）
If Not API_Guard("api", False) Then Response.End

' CSRF验证
If Not API_CheckCSRF() Then
    Call API_Error(API_ERR_CSRF_INVALID, API_GetErrorMessage(API_ERR_CSRF_INVALID))
    Response.End
End If

Dim cartId, sessionId, userId, whereClause
cartId = Request.Form("cartId")
sessionId = Session.SessionID
userId = Session("UserID")

If cartId = "" Or Not IsNumeric(cartId) Then
    Call API_Error(API_ERR_PARAM_INVALID, "无效的购物车项")
    Response.End
End If

If userId <> "" Then
    whereClause = "CartID = " & CInt(cartId) & " AND UserID = " & userId
Else
    whereClause = "CartID = " & CInt(cartId) & " AND SessionID = '" & Replace(sessionId, "'", "''") & "'"
End If

Dim sql
sql = "DELETE FROM Cart WHERE " & whereClause
conn.Execute sql

If Err.Number <> 0 Then
    Call API_Error(API_ERR_DB_ERROR, "删除失败")
    Err.Clear
Else
    Call API_Success(Null, "已移除")
End If

Call CloseConnection()
%>
