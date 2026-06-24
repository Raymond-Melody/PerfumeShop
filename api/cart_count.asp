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

' V10.3: ETag 支持 — 基于 SYS_VERSION + 用户标识 + 购物车数量
Dim eTag, clientETag
eTag = """" & SafeSHA256Hash(SYS_VERSION & "|cart|" & userId & "|" & sessionId & "|" & countVal) & """"
clientETag = Request.ServerVariables("HTTP_IF_NONE_MATCH")

If clientETag <> "" And clientETag = eTag Then
    Response.Status = "304 Not Modified"
    Call CloseConnection()
    Response.End
End If

Response.AddHeader "ETag", eTag
Response.AddHeader "Cache-Control", "no-cache"
Response.Write countVal

Call CloseConnection()
%>
