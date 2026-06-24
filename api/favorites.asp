<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' V15.0 收藏功能API - 统一响应格式
' 处理添加、删除和检查收藏状态
' ============================================
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/api_response.asp"-->
<%
On Error Resume Next

' 检查用户是否登录
If Not API_RequireLogin() Then Response.End

' 打开数据库连接
Call OpenConnection()

' CSRF验证 - 对于修改操作（add/remove）需要验证
Dim action, productId, userId, needsCSRF
action = Request.Form("action")
If action = "" Then action = Request.QueryString("action")
productId = Request.Form("productId")
If productId = "" Then productId = Request.QueryString("productId")
userId = Session("UserID")

needsCSRF = (action = "add" Or action = "remove")
If needsCSRF And Not API_CheckCSRF() Then
    Call API_Error(API_ERR_CSRF_INVALID, API_GetErrorMessage(API_ERR_CSRF_INVALID))
    Call CloseConnection()
    Response.End
End If

' 验证参数
If action = "" Then
    Call API_Error(API_ERR_PARAM_MISSING, "缺少操作类型参数")
    Call CloseConnection()
    Response.End
End If

If Not IsNumeric(productId) Then
    Call API_Error(API_ERR_PARAM_INVALID, "无效的产品ID")
    Call CloseConnection()
    Response.End
End If

productId = CLng(productId)

' 处理不同的操作
Select Case action
    Case "add"
        Dim isAlreadyFav, rsCheck
        isAlreadyFav = False
        Set rsCheck = conn.Execute("SELECT FavoriteID FROM UserFavorites WHERE UserID=" & userId & " AND ProductID=" & productId)
        If Not rsCheck Is Nothing Then
            If Not rsCheck.EOF Then isAlreadyFav = True
            rsCheck.Close
        End If
        Set rsCheck = Nothing
        
        If isAlreadyFav Then
            Dim favResult
            Set favResult = Server.CreateObject("Scripting.Dictionary")
            favResult.Add "action", "added"
            favResult.Add "isFavorite", True
            Call API_Success(favResult, "已收藏")
            Set favResult = Nothing
        Else
            conn.Execute "INSERT INTO UserFavorites (UserID, ProductID, CreatedTime) VALUES (" & userId & ", " & productId & ", GETDATE())"
            If Err.Number <> 0 Then
                Call API_Error(API_ERR_DB_ERROR, "收藏失败")
                Err.Clear
            Else
                Dim addResult
                Set addResult = Server.CreateObject("Scripting.Dictionary")
                addResult.Add "action", "added"
                addResult.Add "isFavorite", True
                Call API_Success(addResult, "收藏成功")
                Set addResult = Nothing
            End If
        End If
        
    Case "remove"
        conn.Execute "DELETE FROM UserFavorites WHERE UserID=" & userId & " AND ProductID=" & productId
        If Err.Number <> 0 Then
            Call API_Error(API_ERR_DB_ERROR, "取消收藏失败")
            Err.Clear
        Else
            Dim removeResult
            Set removeResult = Server.CreateObject("Scripting.Dictionary")
            removeResult.Add "action", "removed"
            removeResult.Add "isFavorite", False
            Call API_Success(removeResult, "取消收藏成功")
            Set removeResult = Nothing
        End If
        
    Case "check"
        Dim isFav, rsCheck2
        isFav = False
        Set rsCheck2 = conn.Execute("SELECT FavoriteID FROM UserFavorites WHERE UserID=" & userId & " AND ProductID=" & productId)
        If Not rsCheck2 Is Nothing Then
            If Not rsCheck2.EOF Then isFav = True
            rsCheck2.Close
        End If
        Set rsCheck2 = Nothing
        
        ' ETag 支持（向后兼容）
        Dim checkETag, clientCheckETag
        checkETag = """" & SafeSHA256Hash(SYS_VERSION & "|fav|" & userId & "|" & productId & "|" & isFav) & """"
        clientCheckETag = Request.ServerVariables("HTTP_IF_NONE_MATCH")
        
        If clientCheckETag <> "" And clientCheckETag = checkETag Then
            Response.Status = "304 Not Modified"
            Call CloseConnection()
            Response.End
        End If
        
        Response.AddHeader "ETag", checkETag
        Response.AddHeader "Cache-Control", "no-cache"
        
        Dim checkResult
        Set checkResult = Server.CreateObject("Scripting.Dictionary")
        checkResult.Add "isFavorite", isFav
        Call API_Success(checkResult, "success")
        Set checkResult = Nothing
        
    Case Else
        Call API_Error(API_ERR_PARAM_INVALID, "无效的操作类型")
End Select

Call CloseConnection()
On Error GoTo 0
%>