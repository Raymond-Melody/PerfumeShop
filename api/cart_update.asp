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

Dim cartId, delta, newQty, sessionId, userId, whereClause
cartId = Request.Form("cartId")
delta = Request.Form("delta")
sessionId = Session.SessionID
userId = Session("UserID")

If cartId = "" Or Not IsNumeric(cartId) Then
    Call API_Error(API_ERR_PARAM_INVALID, "无效的购物车项")
    Response.End
End If

If delta = "" Or Not IsNumeric(delta) Then delta = 0
delta = CInt(delta)

If userId <> "" Then
    whereClause = "CartID = " & CInt(cartId) & " AND UserID = " & userId
Else
    whereClause = "CartID = " & CInt(cartId) & " AND SessionID = '" & Replace(sessionId, "'", "''") & "'"
End If

Dim rs, currentQty
currentQty = Null
Set rs = conn.Execute("SELECT Quantity FROM Cart WHERE " & whereClause)
If Not rs Is Nothing Then
    If Not rs.EOF Then currentQty = rs("Quantity")
    rs.Close
End If
Set rs = Nothing

If IsNull(currentQty) Then
    Call API_Error(API_ERR_NOT_FOUND, "购物车项不存在")
    Response.End
End If

newQty = CInt(currentQty) + delta
If newQty < 1 Then newQty = 1
If newQty > 99 Then newQty = 99

' V16: 实时库存检查
Dim productStock, productID
productID = GetScalar("SELECT ProductID FROM Cart WHERE " & whereClause)
If Not IsNull(productID) And productID <> "" Then
    productStock = GetScalar("SELECT Stock FROM Products WHERE ProductID = " & CLng(productID))
    If Not IsNull(productStock) And CLng(productStock) > 0 And newQty > CLng(productStock) Then
        Dim productName
        productName = GetScalar("SELECT ProductName FROM Products WHERE ProductID = " & CLng(productID))
        Call API_Error(API_ERR_BUSINESS_RULE, "库存不足！" & productName & " 当前仅剩 " & productStock & " 件")
        Response.End
    End If
End If

conn.Execute "UPDATE Cart SET Quantity = " & newQty & " WHERE " & whereClause

If Err.Number <> 0 Then
    Call API_Error(API_ERR_DB_ERROR, "更新失败")
    Err.Clear
Else
    Dim result
    Set result = Server.CreateObject("Scripting.Dictionary")
    result.Add "quantity", newQty
    If Not IsNull(productStock) And productStock <> "" Then
        result.Add "stock", CLng(productStock)
    End If
    Call API_Success(result, "success")
    Set result = Nothing
End If

Call CloseConnection()
%>
