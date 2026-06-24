<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/api_response.asp"-->
<%
Call OpenConnection()

' CSRF验证
If Not API_CheckCSRF() Then
    Call API_Error(API_ERR_CSRF_INVALID, API_GetErrorMessage(API_ERR_CSRF_INVALID))
    Response.End
End If

Dim sessionId, userId, whereClause
sessionId = Session.SessionID
userId = Session("UserID")

If userId <> "" Then
    whereClause = "UserID = " & userId
Else
    whereClause = "SessionID = '" & Replace(sessionId, "'", "''") & "'"
End If

Dim sql
sql = "DELETE FROM Cart WHERE " & whereClause
conn.Execute sql

If Err.Number <> 0 Then
    Call API_Error(API_ERR_DB_ERROR, "操作失败")
    Err.Clear
Else
    Call API_Success(Null, "购物车已清空")
End If

Call CloseConnection()
%>
