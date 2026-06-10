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
        Response.Write "<script>alert('新密码与确认密码不匹配'); history.go(-1);</script>"
        Response.End
    End If
    
    If Len(newPassword) < 6 Then
        Response.Write "<script>alert('新密码长度至少为6位'); history.go(-1);</script>"
        Response.End
    End If
    
    ' 验证当前密码
    Dim rsAdmin, adminId, currentHash
    adminId = Session("AdminID")
    
    Set rsAdmin = ExecuteQuery("SELECT PasswordHash FROM AdminUsers WHERE AdminID = " & adminId)
    If rsAdmin Is Nothing Or rsAdmin.EOF Then
        Response.Write "<script>alert('管理员账户不存在'); history.go(-1);</script>"
        Response.End
    End If
    
    currentHash = rsAdmin("PasswordHash")
    If IsNull(currentHash) Then currentHash = ""
    
    Dim inputHash
    inputHash = GenerateSimpleHash(currentPassword)
    
    If inputHash <> currentHash Then
        Response.Write "<script>alert('当前密码不正确'); history.go(-1);</script>"
        Response.End
    End If
    
    ' 更新密码
    Dim newHash
    newHash = GenerateSimpleHash(newPassword)
    
    Dim updateSql
    updateSql = "UPDATE AdminUsers SET PasswordHash = '" & newHash & "' WHERE AdminID = " & adminId
    
    If ExecuteNonQuery(updateSql) Then
        Response.Write "<script>alert('密码修改成功'); location.href='settings.asp';</script>"
    Else
        Response.Write "<script>alert('密码修改失败'); history.go(-1);</script>"
    End If
End If

Call CloseConnection()
%>