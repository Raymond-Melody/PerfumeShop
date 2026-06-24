<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' V15.0 购物车数量API - 统一响应格式
' ============================================
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/api_response.asp"-->
<%
Call OpenConnection()

Dim sessionId, userId, whereClause, count
sessionId = Session.SessionID
userId = Session("UserID")

' 构建条件
If userId <> "" Then
    whereClause = "UserID = " & userId
Else
    whereClause = "SessionID = '" & Replace(sessionId, "'", "''") & "'"
End If

' 获取购物车数量
Dim rs, countVal, rsCart
countVal = 0
Set rsCart = conn.Execute("SELECT SUM(Quantity) AS TotalQty FROM Cart WHERE " & whereClause)
If Not rsCart Is Nothing Then
    If Not rsCart.EOF Then
        If Not IsNull(rsCart("TotalQty")) Then
            countVal = CLng(rsCart("TotalQty"))
        End If
    End If
    rsCart.Close
End If
Set rsCart = Nothing

' ETag 支持（向后兼容）
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

' V15: 返回标准化JSON格式
Dim result
Set result = Server.CreateObject("Scripting.Dictionary")
result.Add "count", countVal
Call API_Success(result, "success")
Set result = Nothing

Call CloseConnection()
%>
