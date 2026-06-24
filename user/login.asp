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
<%
Call OpenConnection()

Dim errorMsg, successMsg
errorMsg = ""
successMsg = ""

' 检查URL参数消息（如注册成功后跳转）
If Request.QueryString("msg") = "registered" Then
    successMsg = "注册成功！请使用您的账号和密码登录"
End If

' 处理登录
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' CSRF验证
    If Not ValidateCSRFToken() Then
        errorMsg = "安全验证失败，请刷新页面重试"
    ' 速率限制检查
    ElseIf IsLoginLocked("User") Then
        errorMsg = "登录失败次数过多，请15分钟后再试"
    Else
        Dim username, password
        username = Trim(Request.Form("username"))
        password = Trim(Request.Form("password"))
    
        If username = "" Or password = "" Then
            errorMsg = "请输入用户名和密码"
        Else
            ' 查询用户
            Dim rsUser, hashedPwd
            hashedPwd = password ' 实际项目中应该使用加密
            
            Set rsUser = ExecuteQuery("SELECT * FROM Users WHERE (Username = '" & SafeSQL(username) & "' OR Email = '" & SafeSQL(username) & "') AND [Password] = '" & SafeSQL(hashedPwd) & "' AND IsActive <> 0")
            
            If Not rsUser Is Nothing And Not rsUser.EOF Then
                ' 登录成功 - 重置失败计数
                Call ResetLoginFailure("User")
                
                Session("UserID") = rsUser("UserID")
                Session("Username") = rsUser("Username")
                Session("Email") = rsUser("Email")
                Session("FullName") = rsUser("FullName")
                
                ' 合并匿名购物车到用户购物车
                Dim sessionId
                sessionId = Session.SessionID
                Call ExecuteNonQuery("UPDATE Cart SET UserID = " & rsUser("UserID") & ", SessionID = NULL WHERE SessionID = '" & SafeSQL(sessionId) & "' AND UserID IS NULL")
                
                rsUser.Close
                Set rsUser = Nothing
                
                ' 跳转
                Dim returnUrl
                returnUrl = Request.QueryString("return")
                If returnUrl = "" Then returnUrl = "/user/index.asp"
                Response.Redirect returnUrl
            Else
                ' 登录失败 - 记录失败次数
                Call RecordLoginFailure("User")
                errorMsg = "用户名或密码错误"
                If Not rsUser Is Nothing Then
                    rsUser.Close
                    Set rsUser = Nothing
                End If
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
                <h1>欢迎回来</h1>
                <p>登录您的账户，继续定制香氛之旅</p>
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
                    <label for="username"><i class="fas fa-user"></i> 用户名/邮箱</label>
                    <input type="text" id="username" name="username" placeholder="请输入用户名或邮箱" required>
                </div>
                
                <div class="form-group">
                    <label for="password"><i class="fas fa-lock"></i> 密码</label>
                    <div class="password-input">
                        <input type="password" id="password" name="password" placeholder="请输入密码" required>
                        <button type="button" class="toggle-password" onclick="togglePassword()">
                            <i class="fas fa-eye"></i>
                        </button>
                    </div>
                </div>
                
                <div class="form-options">
                    <label class="checkbox-label">
                        <input type="checkbox" name="remember"> 记住我
                    </label>
                    <a href="/user/forgot.asp" class="forgot-link">忘记密码？</a>
                </div>
                
                <button type="submit" class="btn btn-primary btn-lg btn-block">
                    <i class="fas fa-sign-in-alt"></i> 登录
                </button>
            </form>
            
            <div class="auth-divider">
                <span>或</span>
            </div>
            
            <div class="social-login">
                <button type="button" class="btn btn-social btn-wechat">
                    <i class="fab fa-weixin"></i> 微信登录
                </button>
            </div>
            
            <div class="auth-footer">
                <p>还没有账户？ <a href="/user/register.asp">立即注册</a></p>
                <p class="auth-notice"><i class="fas fa-info-circle"></i> 本平台采用会员推荐制，需通过现有会员推荐链接注册</p>
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
