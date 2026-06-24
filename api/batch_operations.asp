<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' V16.0 批量操作工具 (Batch Operations)
' 支持: 批量发货、批量取消、批量上下架
' ============================================
Response.Charset = "UTF-8"
Response.ContentType = "application/json"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/api_response.asp"-->
<!--#include file="../includes/audit_utils.asp"-->
<%
Call OpenConnection()

' V16: 确保审计日志表存在
Call EnsureAuditLogTable()

' 权限检查
If Session("AdminID") = "" Then
    Call API_Error(API_ERR_AUTH_REQUIRED, "请先登录管理后台")
    Response.End
End If

If Not API_CheckCSRF() Then
    Call API_Error(API_ERR_CSRF_INVALID, "安全验证失败")
    Response.End
End If

Dim action, ids, idArray, i, id
action = Trim(Request.Form("action"))
ids = Trim(Request.Form("ids"))

If ids = "" Then
    Call API_Error(API_ERR_PARAM_MISSING, "请选择要操作的记录")
    Response.End
End If

idArray = Split(ids, ",")
Dim successCount, failCount
successCount = 0
failCount = 0

Select Case action
    Case "batch_ship"
        ' 批量发货
        For i = 0 To UBound(idArray)
            id = Trim(idArray(i))
            If IsNumeric(id) And id <> "" Then
                Dim trackingNo
                trackingNo = Trim(Request.Form("tracking_no"))
                If ExecuteNonQuery("UPDATE Orders SET Status='Shipped', TrackingNo='" & SafeSQL(trackingNo) & "', ShippedAt=GETDATE() WHERE OrderID=" & CLng(id) & " AND Status='Paid'") Then
                    successCount = successCount + 1
                Else
                    failCount = failCount + 1
                End If
            End If
        Next
        
    Case "batch_cancel"
        ' 批量取消订单
        For i = 0 To UBound(idArray)
            id = Trim(idArray(i))
            If IsNumeric(id) And id <> "" Then
                If ExecuteNonQuery("UPDATE Orders SET Status='Cancelled', CancelledAt=GETDATE() WHERE OrderID=" & CLng(id) & " AND Status IN ('Pending','Paid')") Then
                    successCount = successCount + 1
                Else
                    failCount = failCount + 1
                End If
            End If
        Next
        
    Case "batch_list"
        ' 批量上架产品
        For i = 0 To UBound(idArray)
            id = Trim(idArray(i))
            If IsNumeric(id) And id <> "" Then
                If ExecuteNonQuery("UPDATE Products SET IsActive=1 WHERE ProductID=" & CLng(id)) Then
                    successCount = successCount + 1
                Else
                    failCount = failCount + 1
                End If
            End If
        Next
        
    Case "batch_unlist"
        ' 批量下架产品
        For i = 0 To UBound(idArray)
            id = Trim(idArray(i))
            If IsNumeric(id) And id <> "" Then
                If ExecuteNonQuery("UPDATE Products SET IsActive=0 WHERE ProductID=" & CLng(id)) Then
                    successCount = successCount + 1
                Else
                    failCount = failCount + 1
                End If
            End If
        Next
        
    Case "batch_delete_cart"
        ' 批量清理过期购物车
        For i = 0 To UBound(idArray)
            id = Trim(idArray(i))
            If IsNumeric(id) And id <> "" Then
                If ExecuteNonQuery("DELETE FROM Cart WHERE CartID=" & CLng(id)) Then
                    successCount = successCount + 1
                Else
                    failCount = failCount + 1
                End If
            End If
        Next
        
    Case Else
        Call API_Error(API_ERR_PARAM_INVALID, "未知操作类型: " & action)
        Response.End
End Select

' V16: 记录审计日志
Call AuditBatch(action, UBound(idArray) + 1, successCount, failCount, "IDs: " & ids)

' 构建响应
Dim responseData
Set responseData = Server.CreateObject("Scripting.Dictionary")
responseData.Add "action", action
responseData.Add "total", UBound(idArray) + 1
responseData.Add "successCount", successCount
responseData.Add "failCount", failCount
Call API_Success(responseData, "批量操作完成：成功" & successCount & "条，失败" & failCount & "条")

Call CloseConnection()
%>
