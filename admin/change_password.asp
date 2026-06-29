<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/password_utils.asp"-->
<%
Call OpenConnection()

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim currentPassword, newPassword, confirmPassword
    currentPassword = Trim(Request.Form("current_password"))
    newPassword = Trim(Request.Form("new_password"))
    confirmPassword = Trim(Request.Form("confirm_password"))
    
    ' 验证输入
    If newPassword <> confirmPassword Then
        Response.Write "<script>alert('" & T("admin_chpwd_mismatch", Empty) & "'); history.go(-1);</script>"
        Response.End
    End If
    
    If Len(newPassword) < 6 Then
        Response.Write "<script>alert('" & T("admin_chpwd_too_short", Empty) & "'); history.go(-1);</script>"
        Response.End
    End If
    
    ' 验证当前密码
    Dim rsAdmin, adminId, currentHash
    adminId = Session("AdminID")
    
    Set rsAdmin = ExecuteQuery("SELECT PasswordHash FROM AdminUsers WHERE AdminID = " & adminId)
    If rsAdmin Is Nothing Or rsAdmin.EOF Then
        Response.Write "<script>alert('" & T("admin_chpwd_account_not_found", Empty) & "'); history.go(-1);</script>"
        Response.End
    End If
    
    currentHash = rsAdmin("PasswordHash")
    If IsNull(currentHash) Then currentHash = ""
    
    ' V17.2: 使用VerifyPassword支持V1/V2/V3全格式 + 全角/半角双向兼容
    If Not VerifyPassword(currentPassword, currentHash) Then
        Response.Write "<script>alert('" & T("admin_chpwd_wrong_current", Empty) & "'); history.go(-1);</script>"
        Response.End
    End If
    
    ' 更新密码 - 使用当前推荐哈希算法(V3/V2)
    Dim newHash
    newHash = HashPassword(newPassword)
    
    Dim updateSql
    updateSql = "UPDATE AdminUsers SET PasswordHash = '" & newHash & "' WHERE AdminID = " & adminId
    
    If ExecuteNonQuery(updateSql) Then
        Response.Write "<script>alert('" & T("admin_chpwd_success", Empty) & "'); location.href='settings.asp';</script>"
    Else
        Response.Write "<script>alert('" & T("admin_chpwd_failed", Empty) & "'); history.go(-1);</script>"
    End If
End If

Call CloseConnection()
%>