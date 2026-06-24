<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/api_response.asp"-->
<%
If Not API_RequireLogin() Then Response.End

Call OpenConnection()

' CSRF验证
If Not API_CheckCSRF() Then
    Call API_Error(API_ERR_CSRF_INVALID, API_GetErrorMessage(API_ERR_CSRF_INVALID))
    Response.End
End If

Dim orderId, userId
orderId = Request.Form("orderId")
userId = Session("UserID")

If orderId = "" Or Not IsNumeric(orderId) Then
    Call API_Error(API_ERR_PARAM_INVALID, "无效的订单")
    Response.End
End If

orderId = CLng(orderId)

' 检查订单归属和状态
Dim rs, currentStatus
currentStatus = ""
Set rs = conn.Execute("SELECT Status FROM Orders WHERE OrderID = " & orderId & " AND UserID = " & userId)
If Not rs Is Nothing Then
    If Not rs.EOF Then currentStatus = rs("Status")
    rs.Close
End If
Set rs = Nothing

If currentStatus = "" Then
    Call API_Error(API_ERR_NOT_FOUND, "订单不存在")
    Response.End
End If

If currentStatus <> "Pending" Then
    Call API_Error(API_ERR_BUSINESS_RULE, "只能取消待付款订单")
    Response.End
End If

conn.Execute "UPDATE Orders SET Status = 'Cancelled', UpdatedAt = GETDATE() WHERE OrderID = " & orderId & " AND UserID = " & userId

If Err.Number <> 0 Then
    Call API_Error(API_ERR_DB_ERROR, "操作失败")
    Err.Clear
Else
    Call API_Success(Null, "订单已取消")
End If

Call CloseConnection()
%>
