<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/plain"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<%
Call OpenConnection()

Dim sessionId, userId
Dim whereClause, count

sessionId = Session.SessionID
userId = Session("UserID")

' 构建条件
If userId <> "" Then
    whereClause = "UserID = " & userId
Else
    whereClause = "SessionID = '" & SafeSQL(sessionId) & "'"
End If

' 获取购物车数量 (Access兼容)
Dim rs, countVal
countVal = 0
Set rs = ExecuteQuery("SELECT SUM(Quantity) AS TotalQty FROM Cart WHERE " & whereClause)
If Not rs Is Nothing Then
    If Not rs.EOF Then
        If Not IsNull(rs("TotalQty")) Then
            countVal = rs("TotalQty")
        End If
    End If
    rs.Close
    Set rs = Nothing
End If

Response.Write countVal

Call CloseConnection()
%>
