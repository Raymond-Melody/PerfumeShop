<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/password_utils.asp"-->
<!--#include file="../includes/member_utils.asp"-->
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"

' 检查用户是否登录
If Session("UserID") = "" Then
    Response.Redirect "/user/login.asp?return=" & Server.URLEncode(Request.ServerVariables("URL"))
End If

Call OpenConnection()

' 处理表单提交
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' CSRF验证
    If Not ValidateCSRFToken() Then
        Response.Write "<script>alert('" & T("user_settings_csrf_fail", Empty) & "'); history.back();</script>"
        Response.End
    End If
    
    ' === 密码修改 ===
    If Request.Form("form_action") = "change_password" Then
        Dim currentPassword, newPassword, confirmPassword, pwdError
        currentPassword = Trim(Request.Form("current_password"))
        newPassword = Trim(Request.Form("new_password"))
        confirmPassword = Trim(Request.Form("confirm_password"))
        pwdError = ""
        
        If currentPassword = "" Or newPassword = "" Or confirmPassword = "" Then
            pwdError = T("user_settings_pwd_fill_all", Empty)
        ElseIf Len(newPassword) < 6 Then
            pwdError = T("user_settings_pwd_len", Empty)
        ElseIf newPassword <> confirmPassword Then
            pwdError = T("user_settings_pwd_mismatch", Empty)
        Else
            ' 验证当前密码
            Dim rsPwd, storedHash, pwdVerified
            pwdVerified = False
            Set rsPwd = ExecuteQuery("SELECT [Password] FROM Users WHERE UserID = " & Session("UserID"))
            If Not rsPwd Is Nothing And Not rsPwd.EOF Then
                storedHash = rsPwd("Password")
                If IsNull(storedHash) Then storedHash = ""
                
                If storedHash = "" Then
                    pwdVerified = False
                ElseIf Left(storedHash, 3) = "V3$" Or Left(storedHash, 3) = "V2_" Then
                    pwdVerified = VerifyPassword(currentPassword, storedHash)
                ElseIf Len(storedHash) <= 40 And InStr(storedHash, "$") = 0 Then
                    ' 兼容旧版明文密码
                    pwdVerified = (storedHash = currentPassword)
                Else
                    pwdVerified = VerifyPassword(currentPassword, storedHash)
                End If
                
                If pwdVerified Then
                    ' 生成新密码哈希并更新
                    Dim newHash
                    newHash = HashPassword(newPassword)
                    If ExecuteNonQuery("UPDATE Users SET [Password] = '" & SafeSQL(newHash) & "' WHERE UserID = " & Session("UserID")) Then
                        Session("pwd_success") = T("user_settings_pwd_changed", Empty)
                        rsPwd.Close : Set rsPwd = Nothing
                        Response.Redirect "login.asp?msg=pwd_changed"
                    Else
                        pwdError = T("user_settings_pwd_fail", Empty)
                    End If
                Else
                    pwdError = T("user_settings_pwd_wrong", Empty)
                End If
                rsPwd.Close
            End If
            Set rsPwd = Nothing
        End If
        
        If pwdError <> "" Then
            Session("pwd_error") = pwdError
            Response.Redirect "settings.asp#change-password"
        End If
    Else
        ' === 个人信息修改 ===
        Dim fullName, email, phone, address, city, postalCode
        fullName = SafeSQL(Trim(Request.Form("fullName")))
        email = SafeSQL(Trim(Request.Form("email")))
        phone = SafeSQL(Trim(Request.Form("phone")))
        address = SafeSQL(Trim(Request.Form("address")))
        city = SafeSQL(Trim(Request.Form("city")))
        postalCode = SafeSQL(Trim(Request.Form("postalCode")))
        
        ' 更新用户信息
        Dim sql
        sql = "UPDATE Users SET FullName = '" & fullName & "', Email = '" & email & "', Phone = '" & phone & "', Address = '" & address & "', City = '" & city & "', PostalCode = '" & postalCode & "' WHERE UserID = " & Session("UserID")
        
        If ExecuteNonQuery(sql) Then
            ' 刷新Session信息
            Session("FullName") = fullName
            Response.Write "<script>alert('" & T("user_settings_saved", Empty) & "'); location.href='settings.asp';</script>"
        Else
            Response.Write "<script>alert('" & T("user_settings_save_fail", Empty) & "'); location.href='settings.asp';</script>"
        End If
    End If
End If

Dim userId, rsUser
userId = Session("UserID")

' 获取用户信息
Set rsUser = ExecuteQuery("SELECT * FROM Users WHERE UserID = " & userId & " AND IsActive <> 0")
If rsUser Is Nothing Or rsUser.EOF Then
    Response.Redirect "/user/login.asp"
End If

' 确保CSRF令牌存在
Call EnsureCSRFToken()

' 获取密码修改反馈消息
Dim pwdErrorMsg, pwdSuccessMsg
pwdErrorMsg = Session("pwd_error")
pwdSuccessMsg = Session("pwd_success")
Session("pwd_error") = ""
Session("pwd_success") = ""
%>
<!--#include file="../includes/header.asp"-->

<!-- 面包屑导航 -->
<div class="breadcrumb">
    <div class="container">
        <a href="/index.asp"><% If FEATURE_I18N Then %><%= T("breadcrumb_home", Empty) %><% Else %>首页<% End If %></a>
        <span class="separator">/</span>
        <a href="/user/index.asp"><% If FEATURE_I18N Then %><%= T("user_nav_center", Empty) %><% Else %>个人中心<% End If %></a>
        <span class="separator">/</span>
        <span><% If FEATURE_I18N Then %><%= T("user_settings_title", Empty) %><% Else %>账户设置<% End If %></span>
    </div>
</div>

<div class="container">
    <div class="user-center">
        <!--#include file="nav.asp"-->
        
        <!-- 主内容 -->
        <div class="user-main">
            <%
            ' V18: 会员等级卡片
            If FEATURE_MEMBER_TIERS Then
                Dim tierInfo, nextTier, rsBenefits, tierTotalSpent, tierDiscount
                Set tierInfo = MU_V18_CalcUserTier(userId)
                Set nextTier = MU_V18_GetNextTierInfo(userId)
                tierTotalSpent = tierInfo("totalSpent")
                tierDiscount = tierInfo("discountRate")
            %>
            <div class="user-card tier-card" style="background: linear-gradient(135deg, <%= tierInfo("badgeBg") %>, #fff); border-left: 4px solid <%= tierInfo("color") %>;">
                <div class="tier-header">
                    <div class="tier-badge" style="background: <%= tierInfo("color") %>;">
                        <i class="fas <%= tierInfo("iconClass") %>"></i>
                    </div>
                    <div class="tier-info">
                        <h3 style="color: <%= tierInfo("color") %>;"><%= tierInfo("tierName") %></h3>
                        <p class="tier-discount">
                            <% If tierDiscount < 1 Then %>
                            <span class="discount-tag"><%= FormatPercent(1 - tierDiscount, 0) %> OFF</span>
                            <% End If %>
                            <% If tierInfo("freeShipping") Then %>
                            <span class="benefit-tag"><i class="fas fa-truck"></i> 免运费</span>
                            <% End If %>
                            <% If tierInfo("priorityShipping") Then %>
                            <span class="benefit-tag"><i class="fas fa-rocket"></i> 优先发货</span>
                            <% End If %>
                            <% If tierInfo("dedicatedSupport") Then %>
                            <span class="benefit-tag"><i class="fas fa-headset"></i> 专属客服</span>
                            <% End If %>
                        </p>
                    </div>
                </div>
                
                <%
                ' 升级进度条
                If Not nextTier Is Nothing Then
                    If nextTier.Exists("isMax") And nextTier("isMax") Then
                %>
                <div class="tier-max-badge">
                    <i class="fas fa-crown"></i> 已达最高等级，尊享全部权益
                </div>
                <%
                    Else
                        Dim progressPct, needSpent, nextName, nextColor
                        progressPct = nextTier("progress")
                        needSpent = nextTier("needSpent")
                        nextName = nextTier("nextTierName")
                        nextColor = nextTier("color")
                %>
                <div class="tier-progress-section">
                    <div class="tier-progress-header">
                        <span>距升级 <strong style="color:<%= nextColor %>;"><%= nextName %></strong> 还需消费 <strong><%= MU_V18_FormatSpent(needSpent) %> 元</strong></span>
                        <span class="tier-progress-pct"><%= progressPct %>%</span>
                    </div>
                    <div class="tier-progress-bar">
                        <div class="tier-progress-fill" style="width: <%= progressPct %>%; background: <%= nextColor %>;"></div>
                    </div>
                </div>
                <%
                    End If
                End If
                %>
                
                <%
                ' 权益列表
                Set rsBenefits = MU_V18_GetTierBenefits(tierInfo("tierCode"))
                If Not rsBenefits Is Nothing And Not rsBenefits.EOF Then
                %>
                <div class="tier-benefits">
                    <h4><i class="fas fa-check-circle"></i> 当前权益</h4>
                    <div class="benefits-grid">
                        <%
                        Dim bCount: bCount = 0
                        Do While Not rsBenefits.EOF And bCount < 8
                            bCount = bCount + 1
                        %>
                        <div class="benefit-item">
                            <i class="fas <%= rsBenefits("BenefitIcon") %>" style="color: <%= tierInfo("color") %>;"></i>
                            <span class="benefit-name"><%= rsBenefits("BenefitName") %></span>
                            <% If rsBenefits("BenefitDesc") <> "" Then %>
                            <span class="benefit-desc"><%= rsBenefits("BenefitDesc") %></span>
                            <% End If %>
                        </div>
                        <%
                            rsBenefits.MoveNext
                        Loop
                        rsBenefits.Close
                        %>
                    </div>
                </div>
                <%
                    End If
                    Set rsBenefits = Nothing
                %>
            </div>
            <%
                Set tierInfo = Nothing
                Set nextTier = Nothing
            End If
            %>
            
            <div class="user-card">
                <h2 class="card-title"><i class="fas fa-cog"></i> <% If FEATURE_I18N Then %><%= T("user_settings_title", Empty) %><% Else %>账户设置<% End If %></h2>
                
                <form method="post" class="form-horizontal" action="settings.asp">
                    <%= GetCSRFTokenField() %>
                    <div class="form-group">
                        <label for="fullName"><% If FEATURE_I18N Then %><%= T("user_settings_name", Empty) %><% Else %>姓名<% End If %></label>
                        <input type="text" id="fullName" name="fullName" value="<%= HTMLEncode(rsUser("FullName")) %>" placeholder="<% If FEATURE_I18N Then %><%= T("user_settings_name_placeholder", Empty) %><% Else %>请输入您的姓名<% End If %>">
                    </div>
                    
                    <div class="form-group">
                        <label for="email"><% If FEATURE_I18N Then %><%= T("user_settings_email", Empty) %><% Else %>邮箱<% End If %></label>
                        <input type="email" id="email" name="email" value="<%= HTMLEncode(rsUser("Email")) %>" placeholder="<% If FEATURE_I18N Then %><%= T("user_settings_email_placeholder", Empty) %><% Else %>请输入邮箱地址<% End If %>">
                    </div>
                    
                    <div class="form-group">
                        <label for="phone"><% If FEATURE_I18N Then %><%= T("user_settings_phone", Empty) %><% Else %>手机号<% End If %></label>
                        <input type="tel" id="phone" name="phone" value="<%= HTMLEncode(rsUser("Phone")) %>" placeholder="<% If FEATURE_I18N Then %><%= T("user_settings_phone_placeholder", Empty) %><% Else %>请输入手机号<% End If %>">
                    </div>
                    
                    <div class="form-group">
                        <label for="address"><% If FEATURE_I18N Then %><%= T("user_settings_address", Empty) %><% Else %>地址<% End If %></label>
                        <textarea id="address" name="address" placeholder="<% If FEATURE_I18N Then %><%= T("user_settings_address_placeholder", Empty) %><% Else %>请输入收货地址<% End If %>"><%= HTMLEncode(rsUser("Address")) %></textarea>
                    </div>
                    
                    <div class="form-group">
                        <label for="city"><% If FEATURE_I18N Then %><%= T("user_settings_city", Empty) %><% Else %>城市<% End If %></label>
                        <input type="text" id="city" name="city" value="<%= HTMLEncode(rsUser("City")) %>" placeholder="<% If FEATURE_I18N Then %><%= T("user_settings_city_placeholder", Empty) %><% Else %>请输入所在城市<% End If %>">
                    </div>
                    
                    <div class="form-group">
                        <label for="postalCode"><% If FEATURE_I18N Then %><%= T("user_settings_postal", Empty) %><% Else %>邮政编码<% End If %></label>
                        <input type="text" id="postalCode" name="postalCode" value="<%= HTMLEncode(rsUser("PostalCode")) %>" placeholder="<% If FEATURE_I18N Then %><%= T("user_settings_postal_placeholder", Empty) %><% Else %>请输入邮政编码<% End If %>">
                    </div>
                    
                    <div class="form-actions">
                        <button type="submit" class="btn btn-primary"><% If FEATURE_I18N Then %><%= T("user_settings_save_btn", Empty) %><% Else %>保存设置<% End If %></button>
                        <button type="button" class="btn btn-text" onclick="history.back()"><% If FEATURE_I18N Then %><%= T("user_settings_cancel", Empty) %><% Else %>取消<% End If %></button>
                    </div>
                </form>
            </div>

            <!-- 修改密码卡片 -->
            <div class="user-card" id="change-password">
                <h2 class="card-title"><i class="fas fa-lock"></i> <% If FEATURE_I18N Then %><%= T("user_settings_change_pwd", Empty) %><% Else %>修改密码<% End If %></h2>
                
                <% If pwdErrorMsg <> "" Then %>
                <div class="alert alert-error" style="margin-bottom:16px;">
                    <i class="fas fa-exclamation-circle"></i> <%= HTMLEncode(pwdErrorMsg) %>
                </div>
                <% End If %>
                <% If pwdSuccessMsg <> "" Then %>
                <div class="alert alert-success" style="margin-bottom:16px;">
                    <i class="fas fa-check-circle"></i> <%= HTMLEncode(pwdSuccessMsg) %>
                </div>
                <% End If %>
                
                <form method="post" class="form-horizontal" action="settings.asp#change-password">
                    <%= GetCSRFTokenField() %>
                    <input type="hidden" name="form_action" value="change_password">
                    
                    <div class="form-group">
                        <label for="current_password"><% If FEATURE_I18N Then %><%= T("user_settings_current_pwd", Empty) %><% Else %>当前密码<% End If %></label>
                        <input type="password" id="current_password" name="current_password" placeholder="<% If FEATURE_I18N Then %><%= T("user_settings_current_pwd_placeholder", Empty) %><% Else %>请输入当前密码<% End If %>" required>
                    </div>
                    
                    <div class="form-group">
                        <label for="new_password"><% If FEATURE_I18N Then %><%= T("user_settings_new_pwd", Empty) %><% Else %>新密码<% End If %></label>
                        <input type="password" id="new_password" name="new_password" placeholder="<% If FEATURE_I18N Then %><%= T("user_settings_new_pwd_placeholder", Empty) %><% Else %>至少6个字符<% End If %>" required minlength="6">
                    </div>
                    
                    <div class="form-group">
                        <label for="confirm_password"><% If FEATURE_I18N Then %><%= T("user_settings_confirm_new_pwd", Empty) %><% Else %>确认新密码<% End If %></label>
                        <input type="password" id="confirm_password" name="confirm_password" placeholder="<% If FEATURE_I18N Then %><%= T("user_settings_confirm_new_pwd_placeholder", Empty) %><% Else %>再次输入新密码<% End If %>" required minlength="6">
                    </div>
                    
                    <div class="form-actions">
                        <button type="submit" class="btn btn-primary">
                            <i class="fas fa-key"></i> <% If FEATURE_I18N Then %><%= T("user_settings_update_pwd_btn", Empty) %><% Else %>更新密码<% End If %>
                        </button>
                    </div>
                </form>
            </div>
        </div>
    </div>
</div>

<script>
function removeAvatar() {
    alert('<% If FEATURE_I18N Then %><%= T("user_settings_avatar_deleted", Empty) %><% Else %>头像已删除<% End If %>');
}
</script>

<% If FEATURE_MEMBER_TIERS Then %>
<style>
/* V18 会员等级卡片样式 */
.tier-card {
    padding: 20px 24px;
    border-radius: 12px;
    margin-bottom: 20px;
    animation: tierFadeIn 0.5s ease;
}
@keyframes tierFadeIn {
    from { opacity: 0; transform: translateY(-10px); }
    to   { opacity: 1; transform: translateY(0); }
}
.tier-header {
    display: flex;
    align-items: center;
    gap: 16px;
    margin-bottom: 16px;
}
.tier-badge {
    width: 56px;
    height: 56px;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 24px;
    color: #fff;
    flex-shrink: 0;
    box-shadow: 0 3px 10px rgba(0,0,0,0.15);
}
.tier-info h3 {
    margin: 0 0 6px 0;
    font-size: 20px;
    font-weight: 700;
}
.tier-discount {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    margin: 0;
}
.discount-tag {
    display: inline-flex;
    align-items: center;
    padding: 3px 10px;
    border-radius: 12px;
    font-size: 13px;
    font-weight: 700;
    background: #e74c3c;
    color: #fff;
}
.benefit-tag {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    padding: 3px 10px;
    border-radius: 12px;
    font-size: 12px;
    background: #f0f0f0;
    color: #555;
}
.tier-max-badge {
    padding: 12px 18px;
    background: linear-gradient(135deg, #ffd700, #ffaa00);
    color: #5d4037;
    border-radius: 8px;
    font-weight: 600;
    text-align: center;
    margin-bottom: 16px;
}
.tier-max-badge i {
    margin-right: 6px;
}
.tier-progress-section {
    margin-bottom: 16px;
}
.tier-progress-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    font-size: 13px;
    color: #666;
    margin-bottom: 8px;
}
.tier-progress-pct {
    font-weight: 700;
    font-size: 15px;
    color: #333;
}
.tier-progress-bar {
    height: 10px;
    background: #e0e0e0;
    border-radius: 5px;
    overflow: hidden;
}
.tier-progress-fill {
    height: 100%;
    border-radius: 5px;
    transition: width 0.6s ease;
    min-width: 2px;
}
.tier-benefits {
    border-top: 1px solid #e8e3dc;
    padding-top: 14px;
}
.tier-benefits h4 {
    font-size: 14px;
    font-weight: 600;
    color: #555;
    margin: 0 0 10px 0;
}
.tier-benefits h4 i {
    margin-right: 6px;
    color: #4CAF50;
}
.benefits-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));
    gap: 10px;
}
.benefit-item {
    display: flex;
    flex-wrap: wrap;
    align-items: baseline;
    gap: 6px;
    padding: 8px 12px;
    background: rgba(255,255,255,0.7);
    border-radius: 8px;
}
.benefit-item i {
    font-size: 14px;
    flex-shrink: 0;
}
.benefit-name {
    font-size: 13px;
    font-weight: 600;
    color: #333;
}
.benefit-desc {
    font-size: 11px;
    color: #999;
    width: 100%;
    margin-top: -2px;
    padding-left: 20px;
}
@media (max-width: 480px) {
    .tier-header {
        flex-direction: column;
        text-align: center;
    }
    .tier-badge {
        width: 48px;
        height: 48px;
        font-size: 20px;
    }
    .tier-discount {
        justify-content: center;
    }
    .benefits-grid {
        grid-template-columns: 1fr;
    }
}
</style>
<% End If %>

<!--#include file="../includes/footer.asp"-->
<%
Call CloseConnection()
%>