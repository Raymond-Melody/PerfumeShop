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

' 处理注册
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' CSRF验证
    If Not ValidateCSRFToken() Then
        errorMsg = "安全验证失败，请刷新页面重试"
    Else
        Dim username, email, password, confirmPwd, fullName, phone
        username = Trim(Request.Form("username"))
        email = Trim(Request.Form("email"))
        password = Trim(Request.Form("password"))
        confirmPwd = Trim(Request.Form("confirmPassword"))
        fullName = Trim(Request.Form("fullName"))
        phone = Trim(Request.Form("phone"))
        
        ' 验证
        If username = "" Or email = "" Or password = "" Then
            errorMsg = "请填写所有必填项"
        ElseIf Len(username) < 3 Or Len(username) > 20 Then
            errorMsg = "用户名长度应为3-20个字符"
        ElseIf Len(password) < 6 Then
            errorMsg = "密码长度至少6个字符"
        ElseIf password <> confirmPwd Then
            errorMsg = "两次输入的密码不一致"
        ElseIf InStr(email, "@") = 0 Then
            errorMsg = "请输入有效的邮箱地址"
        Else
            ' 检查用户名是否已存在
            Dim existCount
            existCount = GetScalar("SELECT COUNT(*) FROM Users WHERE Username = '" & SafeSQL(username) & "'")
            If existCount > 0 Then
                errorMsg = "用户名已被使用"
            Else
                existCount = GetScalar("SELECT COUNT(*) FROM Users WHERE Email = '" & SafeSQL(email) & "'")
                If existCount > 0 Then
                    errorMsg = "邮箱已被注册"
                Else
                    ' 插入新用户
                    Dim sql
                    sql = "INSERT INTO Users (Username, [Password], Email, FullName, Phone, CreatedAt, IsActive) VALUES (" & _
                        "'" & username & "', '" & SafeSQL(password) & "', '" & email & "', " & _
                        "'" & fullName & "', '" & phone & "', GETDATE(), 1)"
                    
                    If ExecuteNonQuery(sql) Then
                        successMsg = "注册成功！请登录"
                        ' 可以选择自动登录
                        Response.Redirect "/user/login.asp?msg=registered"
                    Else
                        errorMsg = "注册失败: " & Session("LastDBError")
                    End If
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
        <div class="auth-card register-card">
            <div class="auth-header">
                <h1>创建账户</h1>
                <p>加入我们，开启专属香氛定制之旅</p>
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
            
            <form method="post" class="auth-form" id="registerForm">
                <%= GetCSRFTokenField() %>
                <div class="form-row">
                    <div class="form-group">
                        <label for="username"><i class="fas fa-user"></i> 用户名 *</label>
                        <input type="text" id="username" name="username" placeholder="3-20个字符" required minlength="3" maxlength="20">
                    </div>
                    <div class="form-group">
                        <label for="email"><i class="fas fa-envelope"></i> 邮箱 *</label>
                        <input type="email" id="email" name="email" placeholder="用于登录和接收通知" required>
                    </div>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label for="password"><i class="fas fa-lock"></i> 密码 *</label>
                        <input type="password" id="password" name="password" placeholder="至少6个字符" required minlength="6">
                    </div>
                    <div class="form-group">
                        <label for="confirmPassword"><i class="fas fa-lock"></i> 确认密码 *</label>
                        <input type="password" id="confirmPassword" name="confirmPassword" placeholder="再次输入密码" required>
                    </div>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label for="fullName"><i class="fas fa-id-card"></i> 姓名</label>
                        <input type="text" id="fullName" name="fullName" placeholder="您的真实姓名（选填）">
                    </div>
                    <div class="form-group">
                        <label for="phone"><i class="fas fa-phone"></i> 手机号</label>
                        <input type="tel" id="phone" name="phone" placeholder="用于订单通知（选填）">
                    </div>
                </div>
                
                <div class="form-group">
                    <label class="checkbox-label">
                        <input type="checkbox" name="agree" required>
                        我已阅读并同意 <a href="/terms.asp" target="_blank">服务条款</a> 和 <a href="/privacy.asp" target="_blank">隐私政策</a>
                    </label>
                </div>
                
                <button type="submit" class="btn btn-primary btn-lg btn-block">
                    <i class="fas fa-user-plus"></i> 立即注册
                </button>
            </form>
            
            <div class="auth-divider">
                <span>或</span>
            </div>
            
            <div class="social-login">
                <button type="button" class="btn btn-social btn-wechat">
                    <i class="fab fa-weixin"></i> 微信快捷注册
                </button>
            </div>
            
            <div class="auth-footer">
                <p>已有账户？ <a href="/user/login.asp">立即登录</a></p>
            </div>
        </div>
        
        <div class="auth-side">
            <div class="auth-promo">
                <i class="fas fa-gift"></i>
                <h2>新用户专享</h2>
                <p>注册即享多重好礼</p>
                <ul class="promo-features">
                    <li><i class="fas fa-check"></i> 首单立减50元</li>
                    <li><i class="fas fa-check"></i> 免费香水小样</li>
                    <li><i class="fas fa-check"></i> 生日专属优惠</li>
                    <li><i class="fas fa-check"></i> 积分兑换好礼</li>
                </ul>
            </div>
        </div>
    </div>
</div>

<script>
$('#registerForm').submit(function(e) {
    var pwd = $('#password').val();
    var confirmPwd = $('#confirmPassword').val();
    if (pwd !== confirmPwd) {
        alert('两次输入的密码不一致');
        e.preventDefault();
        return false;
    }
});
</script>

<!--#include file="../includes/footer.asp"-->
<%
Call CloseConnection()
%>
