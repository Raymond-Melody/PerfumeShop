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
<!--#include file="../includes/dal.asp"-->
<!--#include file="../includes/api_response.asp"-->
<!--#include file="../includes/audit_utils.asp"-->
<!--#include file="../includes/api_guard.asp"-->
<%
Call OpenConnection()

' V17: 确保审计日志表存在
Call EnsureAuditLogTable()

' 权限检查
If Session("AdminID") = "" Then
    Call API_Error(API_ERR_AUTH_REQUIRED, "请先登录管理后台")
    Response.End
End If

' V18: API 守卫（速率限制）
If Not API_Guard("api", False) Then Response.End

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
        ' V17: 使用参数化DAL查询
        For i = 0 To UBound(idArray)
            id = Trim(idArray(i))
            If IsNumeric(id) And id <> "" Then
                Dim trackingNo
                trackingNo = Trim(Request.Form("tracking_no"))
                Dim shipParams(1)
                shipParams(0) = Array("@TrackingNo", DAL_adVarChar, 100, Left(trackingNo, 100))
                shipParams(1) = Array("@OrderID", DAL_adInteger, 0, CLng(id))
                If DAL_Execute("UPDATE Orders SET Status='Shipped', TrackingNo=@TrackingNo, ShippedAt=GETDATE() WHERE OrderID=@OrderID AND Status='Paid'", shipParams) >= 0 Then
                    successCount = successCount + 1
                Else
                    failCount = failCount + 1
                End If
            End If
        Next
        
    Case "batch_cancel"
        ' V17: 使用参数化DAL查询
        For i = 0 To UBound(idArray)
            id = Trim(idArray(i))
            If IsNumeric(id) And id <> "" Then
                Dim cancelParams(0)
                cancelParams(0) = Array("@OrderID", DAL_adInteger, 0, CLng(id))
                If DAL_Execute("UPDATE Orders SET Status='Cancelled', CancelledAt=GETDATE() WHERE OrderID=@OrderID AND Status IN ('Pending','Paid')", cancelParams) >= 0 Then
                    successCount = successCount + 1
                Else
                    failCount = failCount + 1
                End If
            End If
        Next
        
    Case "batch_list"
        ' V17: 使用参数化DAL查询
        For i = 0 To UBound(idArray)
            id = Trim(idArray(i))
            If IsNumeric(id) And id <> "" Then
                Dim listParams(0)
                listParams(0) = Array("@ProductID", DAL_adInteger, 0, CLng(id))
                If DAL_Execute("UPDATE Products SET IsActive=1 WHERE ProductID=@ProductID", listParams) >= 0 Then
                    successCount = successCount + 1
                Else
                    failCount = failCount + 1
                End If
            End If
        Next
        
    Case "batch_unlist"
        ' V17: 使用参数化DAL查询
        For i = 0 To UBound(idArray)
            id = Trim(idArray(i))
            If IsNumeric(id) And id <> "" Then
                Dim unlistParams(0)
                unlistParams(0) = Array("@ProductID", DAL_adInteger, 0, CLng(id))
                If DAL_Execute("UPDATE Products SET IsActive=0 WHERE ProductID=@ProductID", unlistParams) >= 0 Then
                    successCount = successCount + 1
                Else
                    failCount = failCount + 1
                End If
            End If
        Next
        
    Case "batch_delete_cart"
        ' V17: 使用参数化DAL查询
        For i = 0 To UBound(idArray)
            id = Trim(idArray(i))
            If IsNumeric(id) And id <> "" Then
                Dim cartDelParams(0)
                cartDelParams(0) = Array("@CartID", DAL_adInteger, 0, CLng(id))
                If DAL_Execute("DELETE FROM Cart WHERE CartID=@CartID", cartDelParams) >= 0 Then
                    successCount = successCount + 1
                Else
                    failCount = failCount + 1
                End If
            End If
        Next
        
    Case "batch_update_status"
        ' V17: 订单批量状态更新
        Dim newStatus : newStatus = Trim(Request.Form("new_status"))
        Dim validStatuses : validStatuses = Array("Pending", "Paid", "Processing", "Shipped", "Delivered", "Cancelled")
        Dim statusValid : statusValid = False
        For i = 0 To UBound(validStatuses)
            If newStatus = validStatuses(i) Then statusValid = True : Exit For
        Next
        If Not statusValid Then
            Call API_Error(API_ERR_PARAM_INVALID, "无效的订单状态: " & newStatus)
            Response.End
        End If
        
        For i = 0 To UBound(idArray)
            id = Trim(idArray(i))
            If IsNumeric(id) And id <> "" Then
                Dim statusParams(1)
                statusParams(0) = Array("@Status", DAL_adVarChar, 50, newStatus)
                statusParams(1) = Array("@OrderID", DAL_adInteger, 0, CLng(id))
                If DAL_Execute("UPDATE Orders SET Status=@Status, UpdatedAt=GETDATE() WHERE OrderID=@OrderID AND Status<>'Delivered' AND Status<>'Cancelled'", statusParams) >= 0 Then
                    successCount = successCount + 1
                Else
                    failCount = failCount + 1
                End If
            End If
        Next
        
    Case "batch_category"
        ' V17: 产品分类批量调整
        Dim newTypeCode : newTypeCode = Trim(Request.Form("type_code"))
        If newTypeCode = "" Then
            Call API_Error(API_ERR_PARAM_MISSING, "请选择目标分类")
            Response.End
        End If
        
        For i = 0 To UBound(idArray)
            id = Trim(idArray(i))
            If IsNumeric(id) And id <> "" Then
                Dim catParams(1)
                catParams(0) = Array("@TypeCode", DAL_adVarChar, 50, newTypeCode)
                catParams(1) = Array("@ProductID", DAL_adInteger, 0, CLng(id))
                If DAL_Execute("UPDATE Products SET TypeCode=@TypeCode WHERE ProductID=@ProductID", catParams) >= 0 Then
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
