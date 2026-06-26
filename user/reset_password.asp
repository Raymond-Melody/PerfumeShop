<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/i18n.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/password_utils.asp"-->
<%
Call OpenConnection()

Dim token, errorMsg, successMsg, step
token = Trim(Request.QueryString("token"))
errorMsg = ""
successMsg = ""
step = "verify"  ' verify | reset | done

If token = "" Then
    errorMsg = T("user_reset_missing_token", Empty)
    step = "error"
Else
    ' 验证令牌有效性
    Dim rsToken
    Set rsToken = ExecuteQuery("SELECT UserID, Username, Email, FullName, ResetTokenExpiry FROM Users WHERE ResetToken = '" & SafeSQL(token) & "' AND ResetTokenExpiry IS NOT NULL")
    
    If rsToken Is Nothing Or rsToken.EOF Then
        errorMsg = T("user_reset_token_invalid", Empty)
        step = "error"
    Else
        ' 检查是否过期
        Dim expiryDate
        expiryDate = rsToken("ResetTokenExpiry")
        If Not IsNull(expiryDate) Then
            Dim expiryStr, dotPos
            expiryStr = CStr(expiryDate & "")
            dotPos = InStr(expiryStr, ".")
            If dotPos > 0 Then expiryStr = Left(expiryStr, dotPos - 1)
            If IsDate(expiryStr) Then
                If CDate(expiryStr) < Now() Then
                    errorMsg = T("user_reset_token_expired", Empty)
                    step = "error"
                End If
            End If
        End If
    End If
End If

' 处理密码重置提交
If step = "verify" And Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim newPassword, confirmPassword
    newPassword = Trim(Request.Form("new_password"))
    confirmPassword = Trim(Request.Form("confirm_password"))
    
    If newPassword = "" Or confirmPassword = "" Then
        errorMsg = T("user_settings_pwd_fill_all", Empty)
    ElseIf Len(newPassword) < 6 Then
        errorMsg = T("user_settings_pwd_len", Empty)
    ElseIf newPassword <> confirmPassword Then
        errorMsg = T("user_settings_pwd_mismatch", Empty)
    Else
        ' 生成新密码哈希并更新
        Dim newHash
        newHash = HashPassword(newPassword)
        Dim userId
        userId = rsToken("UserID")
        
        If ExecuteNonQuery("UPDATE Users SET [Password] = '" & SafeSQL(newHash) & "', ResetToken = NULL, ResetTokenExpiry = NULL WHERE UserID = " & userId) Then
            successMsg = T("user_reset_success", Empty)
            step = "done"
        Else
            errorMsg = T("user_settings_pwd_fail", Empty)
        End If
    End If
End If

' 关闭记录集
If Not rsToken Is Nothing Then
    If step <> "done" Then rsToken.Close
    If step = "done" Then rsToken.Close
End If
Set rsToken = Nothing
%>
<!--#include file="../includes/header.asp"-->

<div class="auth-page">
    <div class="auth-container">
        <div class="auth-card">
            <% If step = "error" Then %>
            <div class="auth-header">
                <h1><% If FEATURE_I18N Then Response.Write T("user_reset_link_invalid", Empty) Else %>链接无效<% End If %></h1>
                <p><%= HTMLEncode(errorMsg) %></p>
            </div>
            <div class="auth-footer">
                <p><a href="/user/forgot.asp" class="btn btn-primary"><% If FEATURE_I18N Then Response.Write T("user_reset_reapply", Empty) Else %>重新申请重置<% End If %></a></p>
                <p><a href="/user/login.asp"><% If FEATURE_I18N Then Response.Write T("user_forgot_back_link", Empty) Else %>返回登录<% End If %></a></p>
            </div>
            
            <% ElseIf step = "done" Then %>
            <div class="auth-header">
                <h1><% If FEATURE_I18N Then Response.Write T("user_reset_success_title", Empty) Else %>重置成功<% End If %></h1>
                <p><%= HTMLEncode(successMsg) %></p>
            </div>
            <div class="auth-footer">
                <p><a href="/user/login.asp" class="btn btn-primary"><% If FEATURE_I18N Then Response.Write T("user_reset_goto_login", Empty) Else %>前往登录<% End If %></a></p>
            </div>
            
            <% Else %>
            <div class="auth-header">
                <h1><% If FEATURE_I18N Then Response.Write T("user_reset_set_new", Empty) Else %>设置新密码<% End If %></h1>
                <p><% If FEATURE_I18N Then Response.Write T("user_reset_set_for_account", Array(rsToken("Username"))) Else %>为账户 <strong><%= HTMLEncode(rsToken("Username")) %></strong> 设置新密码<% End If %></p>
            </div>
            
            <% If errorMsg <> "" Then %>
            <div class="alert alert-error">
                <i class="fas fa-exclamation-circle"></i> <%= HTMLEncode(errorMsg) %>
            </div>
            <% End If %>
            
            <form method="post" class="auth-form">
                <div class="form-group">
                    <label for="new_password"><i class="fas fa-lock"></i> <% If FEATURE_I18N Then Response.Write T("user_reset_new", Empty) Else %>新密码<% End If %></label>
                    <input type="password" id="new_password" name="new_password" placeholder="<% If FEATURE_I18N Then Response.Write T("user_reset_new_placeholder", Empty) Else %>至少6个字符<% End If %>" required minlength="6">
                </div>
                
                <div class="form-group">
                    <label for="confirm_password"><i class="fas fa-lock"></i> <% If FEATURE_I18N Then Response.Write T("user_reset_confirm", Empty) Else %>确认新密码<% End If %></label>
                    <input type="password" id="confirm_password" name="confirm_password" placeholder="<% If FEATURE_I18N Then Response.Write T("user_reset_confirm_placeholder", Empty) Else %>再次输入新密码<% End If %>" required minlength="6">
                </div>
                
                <button type="submit" class="btn btn-primary btn-lg btn-block">
                    <i class="fas fa-key"></i> <% If FEATURE_I18N Then Response.Write T("user_reset_btn", Empty) Else %>重置密码<% End If %>
                </button>
            </form>
            
            <div class="auth-footer">
                <p><a href="/user/login.asp"><% If FEATURE_I18N Then Response.Write T("user_forgot_back_link", Empty) Else %>返回登录<% End If %></a></p>
            </div>
            <% End If %>
        </div>
        
        <div class="auth-side">
            <div class="auth-promo">
                <i class="fas fa-shield-alt"></i>
                <h2>安全提示</h2>
                <p>请设置一个您能记住但他人难以猜测的密码</p>
                <ul class="promo-features">
                    <li><i class="fas fa-check"></i> 至少6个字符</li>
                    <li><i class="fas fa-check"></i> 建议包含字母和数字</li>
                    <li><i class="fas fa-check"></i> 避免使用生日、手机号等</li>
                </ul>
            </div>
        </div>
    </div>
</div>

<!--#include file="../includes/footer.asp"-->
<%
Call CloseConnection()
%>
