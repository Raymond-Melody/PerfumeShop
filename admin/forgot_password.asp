<%@ Language="VBScript" CodePage="65001" %>

<%
' 包含必要的文件
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/password_utils.asp"-->
<!--#include file="../includes/email_utils.asp"-->
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"

' 打开数据库连接
On Error Resume Next
OpenConnection
If Err.Number <> 0 Then
    Response.Write "<div class='error'>数据库连接错误: " & Err.Description & " 错误号: " & Err.Number & "</div>"
    Response.End
End If
On Error GoTo 0

Dim usernameOrEmail, errorMessage, successMessage
errorMessage = ""
successMessage = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    usernameOrEmail = Trim(Request.Form("username_or_email"))
    
    If usernameOrEmail = "" Then
        errorMessage = "请输入用户名或邮箱地址"
    Else
        ' 查找用户
        Dim rsUser
        Set rsUser = ExecuteQuery("SELECT AdminID, Username, Email, FullName FROM AdminUsers WHERE (Username = '" & SafeSQL(usernameOrEmail) & "') OR (Email = '" & SafeSQL(usernameOrEmail) & "')")
        
        If Not rsUser Is Nothing And Not rsUser.EOF Then
            ' 生成重置令牌
            Dim resetToken, expiryTime
            resetToken = GenerateResetToken()
            expiryTime = DateAdd("h", 1, Now()) ' 1小时后过期
            
            ' 更新数据库中的重置令牌
            Dim updateSql
            updateSql = "UPDATE AdminUsers SET ResetToken = '" & SafeSQL(resetToken) & "', ResetTokenExpiry = '" & expiryTime & "' WHERE AdminID = " & rsUser("AdminID")
            Dim updateResult
            updateResult = ExecuteNonQuery(updateSql)
            
            If updateResult Then
                ' 发送密码重置邮件
                Call SendPasswordResetEmail(rsUser("Email"), rsUser("FullName"), resetToken)
                
                successMessage = "密码重置链接已发送到您的邮箱。请检查您的收件箱。"
            Else
                errorMessage = "发生错误，请稍后重试"
            End If
        Else
            errorMessage = "找不到与输入匹配的管理员账户"
        End If
        
        If Not rsUser Is Nothing Then
            rsUser.Close
            Set rsUser = Nothing
        End If
    End If
End If
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>忘记密码 - 管理员登录</title>
    <link rel="stylesheet" href="/css/style.css">
    <style>
        .admin-login-container {
            max-width: 400px;
            margin: 100px auto;
            padding: 30px;
            background: #fff;
            border-radius: 8px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        }
        
        .admin-login-container h2 {
            text-align: center;
            margin-bottom: 20px;
            color: var(--primary-color);
        }
        
        .form-group {
            margin-bottom: 20px;
        }
        
        .form-group label {
            display: block;
            margin-bottom: 5px;
            font-weight: 500;
        }
        
        .form-group input {
            width: 100%;
            padding: 12px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 16px;
        }
        

        .error-message {
            color: #dc3545;
            text-align: center;
            margin-bottom: 15px;
        }
        
        .success-message {
            color: #28a745;
            text-align: center;
            margin-bottom: 15px;
        }
        
        .back-to-login {
            text-align: center;
            margin-top: 15px;
        }
        
        .back-to-login a {
            color: var(--primary-color);
            text-decoration: none;
        }
        
        .back-to-login a:hover {
            text-decoration: underline;
        }
    </style>
</head>
<body>
    <div class="admin-login-container">
        <h2>重置密码</h2>
        
        <% If errorMessage <> "" Then %>
        <div class="error-message"><%= errorMessage %></div>
        <% End If %>
        
        <% If successMessage <> "" Then %>
        <div class="success-message"><%= successMessage %></div>
        <% End If %>
        
        <form method="post">
            <div class="form-group">
                <label for="username_or_email">用户名或邮箱地址</label>
                <input type="text" id="username_or_email" name="username_or_email" value="<%= usernameOrEmail %>" required>
            </div>
            
            <button type="submit" class="btn">发送重置链接</button>
        </form>
        
        <div class="back-to-login">
            <a href="login.asp">&larr; 返回登录页面</a>
        </div>
    </div>
</body>
</html>
<%
CloseConnection
%>