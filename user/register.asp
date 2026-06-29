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
<!--#include file="../includes/member_utils.asp"-->
<%
Call OpenConnection()

Dim errorMsg, successMsg
errorMsg = ""
successMsg = ""

' V14: 获取推荐Token参数
Dim referralToken, tokenValidation, hasValidToken, referrerName, referrerUserId, expiryDate, tokenHash
referralToken = Trim(Request.QueryString("token"))
hasValidToken = False
referrerName = ""
referrerUserId = 0
expiryDate = ""
tokenHash = ""

If referralToken <> "" Then
    Set tokenValidation = MU_ValidateReferralToken(referralToken)
    If tokenValidation("success") Then
        hasValidToken = True
        referrerName = tokenValidation("referrerName")
        referrerUserId = tokenValidation("referrerUserId")
        expiryDate = tokenValidation("expiryDate")
        tokenHash = tokenValidation("tokenHash")
    Else
        errorMsg = tokenValidation("message")
    End If
    Set tokenValidation = Nothing
End If

' 处理注册
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' CSRF验证
    If Not ValidateCSRFToken() Then
        errorMsg = T("user_register_csrf_fail", Empty)
    Else
        ' V14: 重新验证Token（防止篡改）
        Dim postToken, postTokenValidation
        postToken = Trim(Request.Form("referral_token"))
        Dim postReferrerUserId, postTokenHash
        postReferrerUserId = 0
        postTokenHash = ""
        
        If postToken <> "" Then
            Set postTokenValidation = MU_ValidateReferralToken(postToken)
            If postTokenValidation("success") Then
                postReferrerUserId = postTokenValidation("referrerUserId")
                postTokenHash = postTokenValidation("tokenHash")
            Else
                errorMsg = T("user_register_token_expired", Empty) & ": " & postTokenValidation("message")
                Set postTokenValidation = Nothing
            End If
        Else
            errorMsg = T("user_register_referral_required", Empty)
        End If
        
        If errorMsg = "" Then
            ' V14: 设备指纹
            Dim deviceFingerprint
            deviceFingerprint = Trim(Request.Form("device_fp"))
            
            ' V14: 速率限制检查
            Dim clientIP, rateCheck
            clientIP = Request.ServerVariables("REMOTE_ADDR")
            Set rateCheck = MU_CheckRegistrationRateLimit(clientIP, deviceFingerprint)
            If Not rateCheck("allowed") Then
                errorMsg = rateCheck("message")
            End If
            Set rateCheck = Nothing
        End If
        
        If errorMsg = "" Then
            Dim username, email, password, confirmPwd, fullName, phone
            username = Trim(Request.Form("username"))
            email = Trim(Request.Form("email"))
            password = Trim(Request.Form("password"))
            confirmPwd = Trim(Request.Form("confirmPassword"))
            fullName = Trim(Request.Form("fullName"))
            phone = Trim(Request.Form("phone"))
            
            ' 验证
            If username = "" Or email = "" Or password = "" Then
                errorMsg = T("user_register_fill_all", Empty)
            ElseIf Len(username) < 3 Or Len(username) > 20 Then
                errorMsg = T("user_register_username_len", Empty)
            ElseIf Len(password) < 6 Then
                errorMsg = T("user_register_pwd_len", Empty)
            ElseIf password <> confirmPwd Then
                errorMsg = T("user_register_pwd_mismatch", Empty)
            ElseIf InStr(email, "@") = 0 Then
                errorMsg = T("user_register_invalid_email", Empty)
            Else
                ' V17: 使用DAL参数化查询检查用户名/邮箱是否已存在
            Dim existCount, existParams(0)
            existParams(0) = Array("@Username", DAL_adVarChar, 50, username)
            existCount = CLng(DAL_GetScalar("SELECT COUNT(*) FROM Users WHERE Username = @Username", existParams, 0))
            If existCount > 0 Then
                errorMsg = T("user_register_username_taken", Empty)
            Else
                existParams(0) = Array("@Email", DAL_adVarChar, 100, email)
                existCount = CLng(DAL_GetScalar("SELECT COUNT(*) FROM Users WHERE Email = @Email", existParams, 0))
                If existCount > 0 Then
                    errorMsg = T("user_register_email_taken", Empty)
                Else
                    ' V17: 使用DAL参数化INSERT，杜绝SQL注入
                    Dim hashedPassword, insertSql, insertParams(7)
                    hashedPassword = HashPassword(password)
                    insertSql = "INSERT INTO Users (Username, [Password], Email, FullName, Phone, ReferrerUserID, DeviceFingerprint, CreatedAt, IsActive) " & _
                        "VALUES (@Username, @Password, @Email, @FullName, @Phone, @ReferrerUserID, @DeviceFingerprint, GETDATE(), 1); SELECT SCOPE_IDENTITY();"
                    insertParams(0) = Array("@Username", DAL_adVarChar, 50, username)
                    insertParams(1) = Array("@Password", DAL_adVarChar, 255, hashedPassword)
                    insertParams(2) = Array("@Email", DAL_adVarChar, 100, email)
                    insertParams(3) = Array("@FullName", DAL_adVarChar, 100, fullName)
                    insertParams(4) = Array("@Phone", DAL_adVarChar, 20, phone)
                    insertParams(5) = Array("@ReferrerUserID", DAL_adInteger, 0, postReferrerUserId)
                    insertParams(6) = Array("@DeviceFingerprint", DAL_adVarChar, 200, deviceFingerprint)
                    
                    newUserId = CLng(DAL_GetScalar(insertSql, insertParams, 0))
                            
                    If newUserId > 0 Then
                                Call MU_RecordReferralChain(newUserId, postReferrerUserId)
                            
                            ' V14: 标记Token已使用
                            If postTokenHash <> "" Then
                                Call MU_MarkTokenUsed(postTokenHash)
                            End If
                            
                            ' V14: 记录成功注册（速率限制）
                            Call MU_RecordRegistrationAttempt(clientIP, deviceFingerprint, True, postTokenHash)
                            
                            successMsg = T("user_register_success", Empty)
                            Response.Redirect "/user/login.asp?msg=registered"
                    Else
                            ' 记录失败尝试
                            Call MU_RecordRegistrationAttempt(clientIP, deviceFingerprint, False, postTokenHash)
                            errorMsg = T("user_register_failed", Empty) & ": " & DAL_GetLastError()
                    End If
                    End If
                End If
            End If
        End If
        
        ' V14: 保留验证通过的信息用于重新显示
        If postReferrerUserId > 0 And errorMsg <> "" Then
            hasValidToken = True
            referrerUserId = postReferrerUserId
            tokenHash = postTokenHash
        End If
        
        Set postTokenValidation = Nothing
    End If
End If

' 确保CSRF令牌存在
Call EnsureCSRFToken()
%>
<!--#include file="../includes/header.asp"-->

<div class="auth-page">
    <div class="auth-container">
        <div class="auth-card register-card">
            <% If hasValidToken Then %>
            <!-- V14: 有效Token，显示注册表单 -->
            <div class="auth-header">
                <h1><% If FEATURE_I18N Then %><%= T("user_register_title_referral", Empty) %><% Else %>会员推荐注册<% End If %></h1>
                <div style="background:linear-gradient(135deg,#e8f5e9,#f1f8e9);border-radius:10px;padding:14px 18px;margin-top:12px;border:1px solid #c8e6c9;">
                    <div style="display:flex;align-items:center;gap:10px;">
                        <div style="width:40px;height:40px;background:#4CAF50;border-radius:50%;display:flex;align-items:center;justify-content:center;">
                            <i class="fas fa-user-check" style="color:#fff;font-size:18px;"></i>
                        </div>
                        <div>
                            <div style="font-size:13px;color:#666;"><% If FEATURE_I18N Then %><%= T("user_register_referrer_label", Empty) %><% Else %>推荐人<% End If %></div>
                            <div style="font-size:16px;font-weight:bold;color:#2e7d32;"><%= HTMLEncode(referrerName) %></div>
                        </div>
                    </div>
                    <% If expiryDate <> "" Then %>
                    <div style="margin-top:8px;font-size:12px;color:#689f38;">
                        <i class="fas fa-clock"></i> <% If FEATURE_I18N Then %><%= T("user_register_referral_valid_until", Empty) %><% Else %>推荐链接有效期至<% End If %> <%= expiryDate %>
                    </div>
                    <% End If %>
                </div>
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
                <input type="hidden" name="referral_token" value="<%= HTMLEncode(referralToken) %>">
                <input type="hidden" name="device_fp" id="deviceFingerprint" value="">
                <div class="form-row">
                    <div class="form-group">
                        <label for="username"><i class="fas fa-user"></i> <% If FEATURE_I18N Then %><%= T("user_register_username", Empty) %><% Else %>用户名<% End If %> *</label>
                        <input type="text" id="username" name="username" placeholder="<% If FEATURE_I18N Then %><%= T("user_register_username_placeholder", Empty) %><% Else %>3-20个字符<% End If %>" required minlength="3" maxlength="20">
                    </div>
                    <div class="form-group">
                        <label for="email"><i class="fas fa-envelope"></i> <% If FEATURE_I18N Then %><%= T("user_register_email", Empty) %><% Else %>邮箱<% End If %> *</label>
                        <input type="email" id="email" name="email" placeholder="<% If FEATURE_I18N Then %><%= T("user_register_email_placeholder", Empty) %><% Else %>用于登录和接收通知<% End If %>" required>
                    </div>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label for="password"><i class="fas fa-lock"></i> <% If FEATURE_I18N Then %><%= T("user_register_password", Empty) %><% Else %>密码<% End If %> *</label>
                        <input type="password" id="password" name="password" placeholder="<% If FEATURE_I18N Then %><%= T("user_register_password_placeholder", Empty) %><% Else %>至少6个字符<% End If %>" required minlength="6">
                    </div>
                    <div class="form-group">
                        <label for="confirmPassword"><i class="fas fa-lock"></i> <% If FEATURE_I18N Then %><%= T("user_register_confirm", Empty) %><% Else %>确认密码<% End If %> *</label>
                        <input type="password" id="confirmPassword" name="confirmPassword" placeholder="<% If FEATURE_I18N Then %><%= T("user_register_confirm_placeholder", Empty) %><% Else %>再次输入密码<% End If %>" required>
                    </div>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label for="fullName"><i class="fas fa-id-card"></i> <% If FEATURE_I18N Then %><%= T("user_register_fullname", Empty) %><% Else %>姓名<% End If %></label>
                        <input type="text" id="fullName" name="fullName" placeholder="<% If FEATURE_I18N Then %><%= T("user_register_fullname_placeholder", Empty) %><% Else %>您的真实姓名（选填）<% End If %>">
                    </div>
                    <div class="form-group">
                        <label for="phone"><i class="fas fa-phone"></i> <% If FEATURE_I18N Then %><%= T("user_register_phone", Empty) %><% Else %>手机号<% End If %></label>
                        <input type="tel" id="phone" name="phone" placeholder="<% If FEATURE_I18N Then %><%= T("user_register_phone_placeholder", Empty) %><% Else %>用于订单通知（选填）<% End If %>">
                    </div>
                </div>
                
                <div class="form-group">
                    <label class="checkbox-label">
                        <input type="checkbox" name="agree" required>
                        <% If FEATURE_I18N Then %><%= T("user_register_agree_prefix", Empty) %><% Else %>我已阅读并同意<% End If %> <a href="/terms.asp" target="_blank"><% If FEATURE_I18N Then %><%= T("user_register_agree_terms", Empty) %><% Else %>服务条款<% End If %></a> 和 <a href="/privacy.asp" target="_blank"><% If FEATURE_I18N Then %><%= T("user_register_agree_privacy", Empty) %><% Else %>隐私政策<% End If %></a>
                    </label>
                </div>
                
                <button type="submit" class="btn btn-primary btn-lg btn-block">
                    <i class="fas fa-user-plus"></i> <% If FEATURE_I18N Then %><%= T("user_register_btn", Empty) %><% Else %>立即注册<% End If %>
                </button>
            </form>
            
            <% Else %>
            <!-- V14: 无有效Token，显示推荐制说明 -->
            <div class="auth-header">
                <h1><% If FEATURE_I18N Then %><%= T("user_register_title_no_token", Empty) %><% Else %>会员推荐制<% End If %></h1>
                <p style="color:#888;"><% If FEATURE_I18N Then %><%= T("user_register_subtitle_no_token", Empty) %><% Else %>本平台采用会员推荐制注册<% End If %></p>
            </div>
            
            <% If errorMsg <> "" Then %>
            <div class="alert alert-error">
                <i class="fas fa-exclamation-circle"></i> <%= HTMLEncode(errorMsg) %>
            </div>
            <% End If %>
            
            <div class="referral-info-box" style="text-align:center;padding:30px 20px;">
                <div style="width:80px;height:80px;background:linear-gradient(135deg,#fff8e1,#fff3e0);border-radius:50%;display:flex;align-items:center;justify-content:center;margin:0 auto 20px;">
                    <i class="fas fa-user-shield" style="font-size:36px;color:#e0a800;"></i>
                </div>
                <h3 style="color:#333;margin-bottom:10px;"><% If FEATURE_I18N Then %><%= T("user_register_only_invited", Empty) %><% Else %>仅限受邀注册<% End If %></h3>
                <p style="color:#666;line-height:1.8;"><% If FEATURE_I18N Then %><%= T("user_register_thanks", Empty) %><% Else %>感谢您对香氛定制的关注！<% End If %></p>
                <p style="color:#666;line-height:1.8;"><% If FEATURE_I18N Then %><%= T("user_register_community_quality", Empty) %><% Else %>为确保社区品质，本平台采用<strong>会员邀请制</strong>，<% End If %></p>
                <p style="color:#666;line-height:1.8;"><% If FEATURE_I18N Then %><%= T("user_register_new_user_need_link", Empty) %><% Else %>新用户需通过现有会员的专属推荐链接完成注册。<% End If %></p>
                <div style="margin-top:20px;padding:15px;background:#f8f9fa;border-radius:8px;text-align:left;">
                    <p style="color:#333;font-weight:bold;margin-bottom:8px;"><i class="fas fa-info-circle"></i> <% If FEATURE_I18N Then %><%= T("user_register_how_to_get", Empty) %><% Else %>如何获得注册资格？<% End If %></p>
                    <ul style="color:#666;font-size:14px;line-height:2;padding-left:20px;">
                        <li><% If FEATURE_I18N Then %><%= T("user_register_how_step1", Empty) %><% Else %>请已经是我们会员的朋友发送推荐链接给您<% End If %></li>
                        <li><% If FEATURE_I18N Then %><%= T("user_register_how_step2", Empty) %><% Else %>推荐链接将引导您进入专属注册页面<% End If %></li>
                        <li><% If FEATURE_I18N Then %><%= T("user_register_how_step3", Empty) %><% Else %>链接有效期为30天，请在有效期内完成注册<% End If %></li>
                    </ul>
                </div>
            </div>
            <% End If %>
            
            <div class="auth-footer">
                <p><% If FEATURE_I18N Then %><%= T("user_register_has_account", Empty) %><% Else %>已有账户？<% End If %> <a href="/user/login.asp"><% If FEATURE_I18N Then %><%= T("user_register_login_now", Empty) %><% Else %>立即登录<% End If %></a></p>
            </div>
        </div>
        
        <div class="auth-side">
            <div class="auth-promo">
                <i class="fas fa-gift"></i>
                <h2><% If FEATURE_I18N Then %><%= T("user_register_member_benefits", Empty) %><% Else %>会员专属权益<% End If %></h2>
                <p><% If FEATURE_I18N Then %><%= T("user_register_member_subtitle", Empty) %><% Else %>成为会员即享多重特权<% End If %></p>
                <ul class="promo-features">
                    <li><i class="fas fa-check"></i> <% If FEATURE_I18N Then %><%= T("user_register_benefit1", Empty) %><% Else %>首单立减50元<% End If %></li>
                    <li><i class="fas fa-check"></i> <% If FEATURE_I18N Then %><%= T("user_register_benefit2", Empty) %><% Else %>免费香水小样<% End If %></li>
                    <li><i class="fas fa-check"></i> <% If FEATURE_I18N Then %><%= T("user_register_benefit3", Empty) %><% Else %>生日专属优惠<% End If %></li>
                    <li><i class="fas fa-check"></i> <% If FEATURE_I18N Then %><%= T("user_register_benefit4", Empty) %><% Else %>推荐新会员获积分<% End If %></li>
                </ul>
            </div>
        </div>
    </div>
</div>

<script>
// V14: 生成设备指纹
(function() {
    var fp = '';
    fp += navigator.userAgent.substring(0, 50) + '|';
    fp += screen.width + 'x' + screen.height + '|';
    fp += navigator.language + '|';
    fp += new Date().getTimezoneOffset();
    document.getElementById('deviceFingerprint').value = fp;
})();

$('#registerForm').submit(function(e) {
    var pwd = $('#password').val();
    var confirmPwd = $('#confirmPassword').val();
    if (pwd !== confirmPwd) {
        alert('<% If FEATURE_I18N Then %><%= T("user_register_pwd_mismatch", Empty) %><% Else %>两次输入的密码不一致<% End If %>');
        e.preventDefault();
        return false;
    }
});
</script>

<!--#include file="../includes/footer.asp"-->
<%
Call CloseConnection()
%>
