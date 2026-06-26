<%@ Language="VBScript" CodePage="65001" %>

<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"

' 检查是否已登录
If Not IsEmpty(Session("AdminID")) And Session("AdminID") <> "" Then
    Response.Redirect "portal.asp"
End If

' 包含必要的文件
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/dal.asp"-->
<!--#include file="../includes/dal_users.asp"-->
<!--#include file="../includes/password_utils.asp"-->
<%

' === 打开数据库连接 ===
' 注意：直接调用，不带括号
On Error Resume Next
OpenConnection
If Err.Number <> 0 Then
    Response.Write "<div class='error'>" & T("admin_login_db_conn_error", Empty) & Err.Description & " (Error: " & Err.Number & ")</div>"
    
    Response.End
End If
On Error GoTo 0

Dim username, password, rememberMe, errorMessage
errorMessage = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' CSRF验证
    If Not ValidateCSRFToken() Then
        ' V17.2: 验证失败时强制刷新令牌，允许用户立即重试
        Call GenerateCSRFToken()
        errorMessage = T("admin_login_csrf_error", Empty)
    Else
        username = Trim(Request.Form("username"))
        password = Request.Form("password")
        rememberMe = Request.Form("remember_me")
        
        ' 速率限制检查
        If IsLoginLocked("Admin") Then
            errorMessage = T("admin_login_locked", Empty)
        ElseIf username = "" Or password = "" Then
            errorMessage = T("admin_login_empty_fields", Empty)
        Else
        ' 检查AdminUsers表是否存在（简化检测）
        Dim tableExists
        tableExists = False
        
        On Error Resume Next
        Dim rsCheck
        Set rsCheck = ExecuteQuery("SELECT COUNT(*) AS cnt FROM AdminUsers")
        If Err.Number = 0 And Not rsCheck Is Nothing Then
            tableExists = True
            rsCheck.Close
            Set rsCheck = Nothing
        Else
            Err.Clear
        End If
        On Error GoTo 0
        
        If Not tableExists Then
            errorMessage = T("admin_login_table_missing", Empty)
        Else
            ' V17: 使用参数化查询防止SQL注入
            Dim rsUser, sqlParams(0)
            sqlParams(0) = Array("@Username", DAL_adVarChar, 50, username)
            Set rsUser = DAL_GetList("SELECT AdminID, Username, PasswordHash, IsActive FROM AdminUsers WHERE Username=@Username AND ISNULL(IsActive, 0)<>0", sqlParams)
                        
            If Not rsUser Is Nothing And IsObject(rsUser) Then
                If Not rsUser.EOF And Not rsUser.BOF Then
                    ' 验证密码
                    Dim storedHash, adminId, adminUsername
                    ' 安全地获取字段值
                    If IsNull(rsUser.Fields("PasswordHash").Value) Then
                        storedHash = ""
                    Else
                        storedHash = rsUser.Fields("PasswordHash").Value
                    End If
                                
                    If IsNull(rsUser.Fields("AdminID").Value) Then
                        adminId = ""
                    Else
                        adminId = rsUser.Fields("AdminID").Value
                    End If
                                
                    If IsNull(rsUser.Fields("Username").Value) Then
                        adminUsername = ""
                    Else
                        adminUsername = rsUser.Fields("Username").Value
                    End If
                    
                    ' V17: 使用参数化密码验证 (支持V1/V2/V3自动升级)
                    If AdminVerifyAndUpgrade(password, storedHash, adminId) Then
                        ' 登录成功
                        Session("AdminID") = adminId
                        Session("AdminUsername") = adminUsername
                        Session("AdminName") = adminUsername
                        Session("AdminRealName") = adminUsername
                        
                        ' 重置登录失败计数
                        Call ResetLoginFailure("Admin")
                        
                        ' 如果选择了记住我，设置加密Cookie
                        If rememberMe <> "" Then
                            Response.Cookies("AdminRememberMe") = GenerateSecureToken(adminId)
                            Response.Cookies("AdminRememberMe").Expires = DateAdd("d", 30, Now())
                            Response.Cookies("AdminRememberMe").Path = "/"
                            If LCase(Request.ServerVariables("HTTPS")) = "on" Then
                                Response.Cookies("AdminRememberMe").Secure = True
                            End If
                        End If
                        
                        Response.Redirect "portal.asp"
                    Else
                        Call RecordLoginFailure("Admin")
                        errorMessage = T("admin_login_wrong_credentials", Empty)
                    End If
                Else
                    Call RecordLoginFailure("Admin")
                    errorMessage = T("admin_login_wrong_credentials", Empty)
                End If
                                
                If Not rsUser Is Nothing Then
                    rsUser.Close
                    Set rsUser = Nothing
                End If
            Else
                errorMessage = T("admin_login_db_error", Empty) & DAL_GetLastError()
            End If
        End If
        End If
    End If
End If

Call EnsureCSRFToken()
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title><% If FEATURE_I18N Then %><%= T("admin_login_page_title", Empty) %><% Else %>管理员登录 - 香氛定制电商网站<% End If %></title>
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
            margin-bottom: 10px;
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
        
        .remember-forgot-container {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
        }
        
        .remember-me {
            display: flex;
            align-items: center;
        }
        
        .remember-me input {
            width: auto;
            margin-right: 5px;
        }
        
        .forgot-password {
            text-align: right;
        }
        
        .forgot-password a {
            color: var(--primary-color);
            text-decoration: none;
            font-size: 14px;
        }
        
        .forgot-password a:hover {
            text-decoration: underline;
        }
        

        .error-message {
            color: #dc3545;
            text-align: center;
            margin-bottom: 15px;
            padding: 10px;
            background: #f8d7da;
            border: 1px solid #f5c6cb;
            border-radius: 4px;
        }
        
        .login-info {
            color: #007bff;
            text-align: center;
            margin-bottom: 15px;
            font-size: 14px;
        }
    </style>
</head>
<body>
    <div class="admin-login-container">
        <h2><% If FEATURE_I18N Then %><%= T("admin_login_heading", Empty) %><% Else %>管理员登录<% End If %></h2>
        
        
        <div class="login-info">
            <% If FEATURE_I18N Then %><%= T("admin_login_info", Empty) %><% Else %>请输入管理员账户信息登录<% End If %><br>
        </div>
        
        <% If errorMessage <> "" Then %>
        <div class="error-message"><%= errorMessage %></div>
        <% End If %>
        
        <form method="post">
            <%= GetCSRFTokenField() %>
            <div class="form-group">
                <label for="username"><% If FEATURE_I18N Then %><%= T("admin_login_username", Empty) %><% Else %>用户名<% End If %></label>
                <input type="text" id="username" name="username" value="<%= username %>" required>
            </div>
            
            <div class="form-group">
                <label for="password"><% If FEATURE_I18N Then %><%= T("admin_login_password", Empty) %><% Else %>密码<% End If %></label>
                <input type="password" id="password" name="password" required>
            </div>
            
            <div class="remember-forgot-container">
                <div class="remember-me">
                    <input type="checkbox" id="remember_me" name="remember_me" value="1">
                    <label for="remember_me"><% If FEATURE_I18N Then %><%= T("admin_login_remember", Empty) %><% Else %>记住我<% End If %></label>
                </div>
                
                <div class="forgot-password">
                    <a href="forgot_password.asp"><% If FEATURE_I18N Then %><%= T("admin_login_forgot", Empty) %><% Else %>忘记密码？<% End If %></a>
                </div>
            </div>
            
            <button type="submit" class="btn"><% If FEATURE_I18N Then %><%= T("admin_login_btn", Empty) %><% Else %>登录<% End If %></button>
        </form>
    </div>
</body>
</html>
<%
' === 关闭数据库连接 ===
On Error Resume Next
CloseConnection
On Error GoTo 0
%>