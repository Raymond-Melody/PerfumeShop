<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"

' 如果已登录，跳转到个人中心
If Session("UserID") <> "" Then
    Response.Redirect "/user/index.asp"
End If
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/password_utils.asp"-->
<!--#include file="../includes/email_utils.asp"-->
<%
Call OpenConnection()

Dim usernameOrEmail, errorMessage, successMessage
errorMessage = ""
successMessage = ""
usernameOrEmail = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    usernameOrEmail = Trim(Request.Form("username_or_email"))
    
    If usernameOrEmail = "" Then
        errorMessage = "请输入用户名或邮箱地址"
    Else
        ' 查找用户
        Dim rsUser
        Set rsUser = ExecuteQuery("SELECT UserID, Username, Email, FullName FROM Users WHERE (Username = '" & SafeSQL(usernameOrEmail) & "') OR (Email = '" & SafeSQL(usernameOrEmail) & "')")
        
        If Not rsUser Is Nothing And Not rsUser.EOF Then
            ' 生成重置令牌
            Dim resetToken, expiryTime
            resetToken = GenerateResetToken()
            expiryTime = DateAdd("h", 1, Now())
            
            ' 更新数据库中的重置令牌
            Dim updateSql
            updateSql = "UPDATE Users SET ResetToken = '" & SafeSQL(resetToken) & "', ResetTokenExpiry = '" & expiryTime & "' WHERE UserID = " & rsUser("UserID")
            Dim updateResult
            updateResult = ExecuteNonQuery(updateSql)
            
            If updateResult Then
                ' 发送密码重置邮件
                On Error Resume Next
                Call SendUserPasswordResetEmail(rsUser("Email"), rsUser("FullName"), resetToken)
                On Error GoTo 0
                
                successMessage = "密码重置链接已发送到您的邮箱。请检查您的收件箱。"
            Else
                errorMessage = "发生错误，请稍后重试"
            End If
        Else
            errorMessage = "找不到与输入匹配的用户账户"
        End If
        
        If Not rsUser Is Nothing Then
            rsUser.Close
            Set rsUser = Nothing
        End If
    End If
End If
%>
<!--#include file="../includes/header.asp"-->

<div class="auth-page">
    <div class="auth-container">
        <div class="auth-card">
            <div class="auth-header">
                <h1>找回密码</h1>
                <p>输入您的用户名或邮箱，我们将发送密码重置链接</p>
            </div>
            
            <% If errorMessage <> "" Then %>
            <div class="alert alert-error">
                <i class="fas fa-exclamation-circle"></i> <%= HTMLEncode(errorMessage) %>
            </div>
            <% End If %>
            
            <% If successMessage <> "" Then %>
            <div class="alert alert-success">
                <i class="fas fa-check-circle"></i> <%= HTMLEncode(successMessage) %>
            </div>
            <% Else %>
            <form method="post" class="auth-form">
                <div class="form-group">
                    <label for="username_or_email"><i class="fas fa-user"></i> 用户名或邮箱</label>
                    <input type="text" id="username_or_email" name="username_or_email" value="<%= HTMLEncode(usernameOrEmail) %>" placeholder="请输入注册时使用的用户名或邮箱" required>
                </div>
                
                <button type="submit" class="btn btn-primary btn-lg btn-block">
                    <i class="fas fa-paper-plane"></i> 发送重置链接
                </button>
            </form>
            <% End If %>
            
            <div class="auth-footer">
                <p>想起密码了？ <a href="/user/login.asp">返回登录</a></p>
            </div>
        </div>
        
        <div class="auth-side">
            <div class="auth-promo">
                <i class="fas fa-lock-open"></i>
                <h2>安全提示</h2>
                <p>密码重置链接在1小时内有效，请尽快完成操作</p>
                <ul class="promo-features">
                    <li><i class="fas fa-check"></i> 链接将发送到注册邮箱</li>
                    <li><i class="fas fa-check"></i> 请在安全环境下操作</li>
                    <li><i class="fas fa-check"></i> 设置强密码保护账户</li>
                </ul>
            </div>
        </div>
    </div>
</div>

<!--#include file="../includes/footer.asp"-->
<%
Call CloseConnection()
%>
