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
<!--#include file="../includes/dal.asp"-->
<!--#include file="../includes/password_utils.asp"-->
<%
Call OpenConnection()

Dim errorMsg, successMsg
errorMsg = ""
successMsg = ""

' 检查URL参数消息（如注册成功后跳转）
If Request.QueryString("msg") = "registered" Then
    successMsg = T("user_login_success_reg", Empty)
ElseIf Request.QueryString("msg") = "pwd_changed" Then
    successMsg = T("user_login_success_pwd", Empty)
End If

' 处理登录
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' CSRF验证
    If Not ValidateCSRFToken() Then
        ' V17.2: 验证失败时强制刷新令牌，允许用户立即重试
        Call GenerateCSRFToken()
        errorMsg = T("user_login_csrf_fail", Empty)
    ' 速率限制检查
    ElseIf IsLoginLocked("User") Then
        errorMsg = T("user_login_locked_15min", Empty)
    Else
        Dim username, password
        username = Trim(Request.Form("username"))
        password = Trim(Request.Form("password"))
    
        If username = "" Or password = "" Then
            errorMsg = T("user_login_empty", Empty)
        Else
            ' V15: 使用安全的密码哈希验证
            Dim rsUser, storedHash, verifiedPwd
            verifiedPwd = False
            
            ' V17: 参数化查询防止SQL注入
            Dim loginParams(0)
            loginParams(0) = Array("@Username", DAL_adVarChar, 50, username)
            Set rsUser = DAL_GetList("SELECT UserID, Username, Email, FullName, [Password] FROM Users WHERE (Username=@Username OR Email=@Username) AND IsActive<>0", loginParams)
            
            If Not rsUser Is Nothing Then
                            If Not rsUser.EOF Then
                ' 获取存储的密码哈希（兼容明文旧数据）
                If IsNull(rsUser("Password")) Then
                    storedHash = ""
                Else
                    storedHash = rsUser("Password")
                End If
                
                ' V15: 使用VerifyPassword验证（支持V1/V2/V3和向后兼容明文）
                If storedHash = "" Then
                    verifiedPwd = False
                ElseIf Left(storedHash, 3) = "V3$" Or Left(storedHash, 3) = "V2_" Then
                    ' 使用V15密码验证
                    verifiedPwd = VerifyPassword(password, storedHash)
                ElseIf Len(storedHash) <= 40 And InStr(storedHash, "$") = 0 Then
                    ' 兼容旧版明文密码（逐步淘汰）
                    verifiedPwd = (storedHash = password)
                Else
                    ' V1格式或其他哈希
                    verifiedPwd = VerifyPassword(password, storedHash)
                End If
                
                If verifiedPwd Then
                    ' 登录成功 - 重置失败计数
                    Call ResetLoginFailure("User")
                    
                    Session("UserID") = rsUser("UserID")
                    Session("Username") = rsUser("Username")
                    Session("Email") = rsUser("Email")
                    Session("FullName") = rsUser("FullName")
                    
                    ' V17: 使用DAL参数化查询合并匿名购物车
                    Dim sessionId
                    sessionId = Session.SessionID
                    Dim cartParams(1)
                    cartParams(0) = Array("@UserID", DAL_adInteger, 0, CLng(rsUser("UserID")))
                    cartParams(1) = Array("@SessionID", DAL_adVarChar, 50, sessionId)
                    DAL_Execute "UPDATE Cart SET UserID = @UserID, SessionID = NULL WHERE SessionID = @SessionID AND UserID IS NULL", cartParams
                    
                    ' V17.2: 密码含全角字符时自动升级为标准化哈希 (DAL参数化)
                    Dim normalizedPwd
                    normalizedPwd = NormalizePasswordInput(password)
                    If normalizedPwd <> password Then
                        Dim newNormalizedHash, normPwdParams(1)
                        newNormalizedHash = HashPassword(password)
                        normPwdParams(0) = Array("@Password", DAL_adVarChar, 255, newNormalizedHash)
                        normPwdParams(1) = Array("@UserID", DAL_adInteger, 0, CLng(rsUser("UserID")))
                        DAL_Execute "UPDATE Users SET [Password] = @Password WHERE UserID = @UserID", normPwdParams
                    End If
                    
                    ' V15: 检查是否需要升级密码哈希 (DAL参数化)
                    If NeedsPasswordUpgrade(storedHash) Then
                        Dim newHash, upgPwdParams(1)
                        newHash = HashPassword(password)
                        upgPwdParams(0) = Array("@Password", DAL_adVarChar, 255, newHash)
                        upgPwdParams(1) = Array("@UserID", DAL_adInteger, 0, CLng(rsUser("UserID")))
                        DAL_Execute "UPDATE Users SET [Password] = @Password WHERE UserID = @UserID", upgPwdParams
                    End If
                    
                    rsUser.Close
                    Set rsUser = Nothing
                    
                    ' 跳转
                    Dim returnUrl
                    returnUrl = Request.QueryString("return")
                    If returnUrl = "" Then returnUrl = "/user/index.asp"
                    Response.Redirect returnUrl
                End If
                End If
            End If
            ' 登录失败 - 记录失败次数
            If Not verifiedPwd Then
                Call RecordLoginFailure("User")
                errorMsg = T("user_login_invalid", Empty)
            End If
            If Not rsUser Is Nothing Then
                rsUser.Close
                Set rsUser = Nothing
            End If
        End If
    End If
End If

' 确保CSRF令牌存在
Call EnsureCSRFToken()
%>
<!--#include file="../includes/header.asp"-->

<div class="auth-page">
    <div class="auth-container">
        <div class="auth-card">
            <div class="auth-header">
                <h1><% If FEATURE_I18N Then %><%= T("user_login_welcome", Empty) %><% Else %>欢迎回来<% End If %></h1>
                <p><% If FEATURE_I18N Then %><%= T("user_login_subtitle", Empty) %><% Else %>登录您的账户，继续定制香氛之旅<% End If %></p>
            </div>
            
            <% If errorMsg <> "" Then %>
            <div class="alert alert-error">
                <i class="fas fa-exclamation-circle"></i> <%= HTMLEncode(errorMsg) %>
            </div>
            <% End If %>
            
            <% If successMsg <> "" Then %>
            <div class="alert alert-success">
                <i class="fas fa-check-circle"></i> <%= HTMLEncode(successMsg) %>
            </div>
            <% End If %>
            
            <form method="post" class="auth-form" id="loginForm">
                <%= GetCSRFTokenField() %>
                <div class="form-group">
                    <label for="username"><i class="fas fa-user"></i> <% If FEATURE_I18N Then %><%= T("user_login_username", Empty) %><% Else %>用户名/邮箱<% End If %></label>
                    <input type="text" id="username" name="username" placeholder="<% If FEATURE_I18N Then %><%= T("user_login_username_placeholder", Empty) %><% Else %>请输入用户名或邮箱<% End If %>" required>
                </div>
                
                <div class="form-group">
                    <label for="password"><i class="fas fa-lock"></i> <% If FEATURE_I18N Then %><%= T("user_login_password", Empty) %><% Else %>密码<% End If %></label>
                    <div class="password-input">
                        <input type="password" id="password" name="password" placeholder="<% If FEATURE_I18N Then %><%= T("user_login_password_placeholder", Empty) %><% Else %>请输入密码<% End If %>" required>
                        <button type="button" class="toggle-password" onclick="togglePassword()">
                            <i class="fas fa-eye"></i>
                        </button>
                    </div>
                </div>
                
                <div class="form-options">
                    <label class="checkbox-label">
                        <input type="checkbox" name="remember"> <% If FEATURE_I18N Then %><%= T("user_login_remember", Empty) %><% Else %>记住我<% End If %>
                    </label>
                    <a href="/user/forgot.asp" class="forgot-link"><% If FEATURE_I18N Then %><%= T("user_login_forgot", Empty) %><% Else %>忘记密码？<% End If %></a>
                </div>
                
                <button type="submit" class="btn btn-primary btn-lg btn-block">
                    <i class="fas fa-sign-in-alt"></i> <% If FEATURE_I18N Then %><%= T("user_login_btn", Empty) %><% Else %>登录<% End If %>
                </button>
            </form>
            
            <div class="auth-divider">
                <span><% If FEATURE_I18N Then %><%= T("user_login_or", Empty) %><% Else %>或<% End If %></span>
            </div>
            
            <div class="social-login">
                <button type="button" class="btn btn-social btn-wechat">
                    <i class="fab fa-weixin"></i> <% If FEATURE_I18N Then %><%= T("user_login_wechat", Empty) %><% Else %>微信登录<% End If %>
                </button>
            </div>
            
            <div class="auth-footer">
                <p><% If FEATURE_I18N Then %><%= T("user_login_no_account", Empty) %><% Else %>还没有账户？<% End If %> <a href="/user/register.asp"><% If FEATURE_I18N Then %><%= T("user_login_register_now", Empty) %><% Else %>立即注册<% End If %></a></p>
                <p class="auth-notice"><i class="fas fa-info-circle"></i> <% If FEATURE_I18N Then %><%= T("user_register_referral_required", Empty) %><% Else %>本平台采用会员推荐制，需通过现有会员推荐链接注册<% End If %></p>
            </div>
        </div>
        
        <div class="auth-side">
            <div class="auth-promo">
                <i class="fas fa-spray-can"></i>
                <h2>香氛定制</h2>
                <p>定制你的专属香水，让气味成为你独特的名片</p>
                <ul class="promo-features">
                    <li><i class="fas fa-check"></i> 个性化香调搭配</li>
                    <li><i class="fas fa-check"></i> 优质天然原料</li>
                    <li><i class="fas fa-check"></i> 专属瓶身定制</li>
                    <li><i class="fas fa-check"></i> 会员专属优惠</li>
                </ul>
            </div>
        </div>
    </div>
</div>

<script>
function togglePassword() {
    var pwd = document.getElementById('password');
    var icon = document.querySelector('.toggle-password i');
    if (pwd.type === 'password') {
        pwd.type = 'text';
        icon.className = 'fas fa-eye-slash';
    } else {
        pwd.type = 'password';
        icon.className = 'fas fa-eye';
    }
}
</script>

<!--#include file="../includes/footer.asp"-->
<%
Call CloseConnection()
%>
