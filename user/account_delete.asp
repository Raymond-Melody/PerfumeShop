<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"

' 必须登录
If Session("UserID") = "" Or IsEmpty(Session("UserID")) Then
    Response.Redirect "login.asp?return=" & Server.URLEncode("account_delete.asp")
End If
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/dal.asp"-->
<!--#include file="../includes/password_utils.asp"-->
<%
Call OpenConnection()

Dim userId, action, errorMsg, successMsg
userId = CLng(Session("UserID"))
action = Trim(Request.Form("action"))
errorMsg = ""
successMsg = ""

' 检查是否已经有挂起的注销请求
Dim existingDeletion, deletionDate, coolingEnd
existingDeletion = False
Dim delParams(0)
delParams(0) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
Dim rsDel
Set rsDel = DAL_GetRow("SELECT DeletionRequestedAt, DATEADD(day, 30, DeletionRequestedAt) AS CoolingEnd, IsDeleted FROM Users WHERE UserID=@UserID", delParams)
If Not rsDel Is Nothing And Not rsDel.EOF Then
    If Not IsNull(rsDel("IsDeleted")) And rsDel("IsDeleted") = True Then
        existingDeletion = True
        ' 已删除 - 显示完成状态
    ElseIf Not IsNull(rsDel("DeletionRequestedAt")) Then
        existingDeletion = True
        deletionDate = rsDel("DeletionRequestedAt")
        coolingEnd = rsDel("CoolingEnd")
    End If
    rsDel.Close
End If
Set rsDel = Nothing

' 处理操作
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' CSRF 验证
    If Not ValidateCSRFToken() Then
        Call GenerateCSRFToken()
        errorMsg = T("privacy_delete_csrf_error", Empty)
        If errorMsg = "" Then errorMsg = "安全验证失败，请刷新页面重试"
    Else
        action = Trim(Request.Form("action"))
        
        If action = "request_delete" Then
            ' 验证密码
            Dim password
            password = Trim(Request.Form("password"))
            If password = "" Then
                errorMsg = T("privacy_delete_password_required", Empty)
                If errorMsg = "" Then errorMsg = "请输入密码以确认操作"
            Else
                ' 验证密码
                Dim userRow, storedHash
                Set userRow = DAL_GetRow("SELECT [Password] FROM Users WHERE UserID=@UserID", delParams)
                If Not userRow Is Nothing And Not userRow.EOF Then
                    storedHash = userRow("Password") & ""
                    userRow.Close
                    If VerifyPassword(password, storedHash) Then
                        ' 发起注销请求
                        Dim updateParams(1)
                        updateParams(0) = Array("@DeletionRequestedAt", DAL_adDateTime, 0, Now())
                        updateParams(1) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
                        DAL_Execute "UPDATE Users SET DeletionRequestedAt=@DeletionRequestedAt, IsActive=0 WHERE UserID=@UserID", updateParams
                        
                        ' 记录隐私操作日志
                        If FEATURE_GDPR_COMPLIANCE Then
                            On Error Resume Next
                            Dim logParams(4)
                            ReDim logParams(0)
                            logParams(0) = Array("@LogLevel", DAL_adVarChar, 10, "PRIVACY")
                            ReDim Preserve logParams(1)
                            logParams(1) = Array("@LogMessage", DAL_adVarChar, 2000, "账户注销请求 - UserID:" & userId & " 用户名:" & Session("Username"))
                            ReDim Preserve logParams(2)
                            logParams(2) = Array("@LogSource", DAL_adVarChar, 100, "user/account_delete.asp")
                            ReDim Preserve logParams(3)
                            logParams(3) = Array("@IPAddress", DAL_adVarChar, 50, Left(Request.ServerVariables("REMOTE_ADDR"), 50))
                            ReDim Preserve logParams(4)
                            logParams(4) = Array("@PageURL", DAL_adVarChar, 500, Left(Request.ServerVariables("SCRIPT_NAME"), 500))
                            DAL_Execute "INSERT INTO AppLogs (LogLevel, LogMessage, LogSource, IPAddress, PageURL) VALUES (@LogLevel, @LogMessage, @LogSource, @IPAddress, @PageURL)", logParams
                            Err.Clear
                            On Error GoTo 0
                        End If
                        
                        ' 清除 Session
                        Session.Abandon()
                        
                        successMsg = T("privacy_delete_success", Empty)
                        If successMsg = "" Then successMsg = "注销请求已提交。您的账户将在30天冷静期后被永久删除。在此期间重新登录即可取消注销。"
                        existingDeletion = True
                    Else
                        errorMsg = T("privacy_delete_wrong_password", Empty)
                        If errorMsg = "" Then errorMsg = "密码错误，请重试"
                    End If
                Else
                    errorMsg = T("privacy_delete_error", Empty)
                    If errorMsg = "" Then errorMsg = "操作失败，请稍后重试"
                End If
                Set userRow = Nothing
            End If
            
        ElseIf action = "cancel_delete" Then
            ' 取消注销请求
            Dim cancelParams(0)
            cancelParams(0) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
            DAL_Execute "UPDATE Users SET DeletionRequestedAt=NULL, IsActive=1 WHERE UserID=@UserID", cancelParams
            
            successMsg = T("privacy_cancel_success", Empty)
            If successMsg = "" Then successMsg = "注销请求已取消，您的账户已恢复正常状态。"
            existingDeletion = False
        End If
    End If
End If

Call EnsureCSRFToken()
%>
<!--#include file="../includes/header.asp"-->

<div class="delete-page">
    <div class="container">
        <div class="delete-header">
            <h1><% If FEATURE_I18N Then %><%= T("privacy_delete_title", Empty) %><% Else %>账户注销<% End If %></h1>
            <p><% If FEATURE_I18N Then %><%= T("privacy_delete_desc", Empty) %><% Else %>根据数据保护法规（GDPR），您有权请求删除您的账户和个人数据。<% End If %></p>
        </div>

        <% If errorMsg <> "" Then %>
        <div class="alert alert-error"><i class="fas fa-exclamation-circle"></i> <%= HTMLEncode(errorMsg) %></div>
        <% End If %>

        <% If successMsg <> "" Then %>
        <div class="alert alert-success"><i class="fas fa-check-circle"></i> <%= HTMLEncode(successMsg) %></div>
        <% End If %>

        <% If existingDeletion And Session("UserID") = "" Then %>
        <!-- 注销后状态（已清除Session） -->
        <div class="delete-status-card">
            <div class="status-icon"><i class="fas fa-clock"></i></div>
            <h2><% If FEATURE_I18N Then %><%= T("privacy_delete_processing", Empty) %><% Else %>注销处理中<% End If %></h2>
            <p><% If FEATURE_I18N Then %><%= T("privacy_delete_processing_text", Empty) %><% Else %>您的账户注销请求已提交。账户将在30天冷静期后被永久删除。在此期间您可以重新登录来取消注销。<% End If %></p>
            <a href="login.asp" class="btn btn-primary"><% If FEATURE_I18N Then %><%= T("privacy_delete_re_login", Empty) %><% Else %>重新登录<% End If %></a>
        </div>

        <% ElseIf existingDeletion Then %>
        <!-- 挂起的注销请求（已登录状态） -->
        <div class="delete-status-card">
            <div class="status-icon"><i class="fas fa-hourglass-half"></i></div>
            <h2><% If FEATURE_I18N Then %><%= T("privacy_delete_pending", Empty) %><% Else %>注销请求已提交<% End If %></h2>
            <div class="cooling-timeline">
                <div class="cooling-info">
                    <p><% If FEATURE_I18N Then %><%= T("privacy_delete_requested", Empty) %><% Else %>请求时间：<% End If %><%= deletionDate %></p>
                    <p><% If FEATURE_I18N Then %><%= T("privacy_delete_cooling_end", Empty) %><% Else %>冷静期截止：<% End If %><%= coolingEnd %></p>
                </div>
            </div>
            <p class="cooling-note"><% If FEATURE_I18N Then %><%= T("privacy_delete_cooling_note", Empty) %><% Else %>在冷静期内，重新登录即可自动取消注销。如需立即取消，请点击下方按钮。<% End If %></p>
            <form method="post" class="cancel-form">
                <%= GetCSRFTokenField() %>
                <input type="hidden" name="action" value="cancel_delete">
                <button type="submit" class="btn btn-success btn-lg">
                    <i class="fas fa-undo"></i> <% If FEATURE_I18N Then %><%= T("privacy_delete_cancel_btn", Empty) %><% Else %>取消注销，恢复账户<% End If %>
                </button>
            </form>
        </div>

        <% Else %>
        <!-- 注销确认表单 -->
        <div class="delete-warning">
            <div class="warning-icon"><i class="fas fa-exclamation-triangle"></i></div>
            <h3><% If FEATURE_I18N Then %><%= T("privacy_delete_warning_title", Empty) %><% Else %>请仔细阅读以下信息<% End If %></h3>
            <ul class="warning-list">
                <li><% If FEATURE_I18N Then %><%= T("privacy_delete_warning_1", Empty) %><% Else %>账户注销后，您将无法使用本站的任何会员服务<% End If %></li>
                <li><% If FEATURE_I18N Then %><%= T("privacy_delete_warning_2", Empty) %><% Else %>您的个人资料、收货地址将被永久删除<% End If %></li>
                <li><% If FEATURE_I18N Then %><%= T("privacy_delete_warning_3", Empty) %><% Else %>根据税务法规，订单记录将保留7年（已脱敏处理）<% End If %></li>
                <li><% If FEATURE_I18N Then %><%= T("privacy_delete_warning_4", Empty) %><% Else %>注销后有30天冷静期，期间登录即可恢复账户<% End If %></li>
                <li><% If FEATURE_I18N Then %><%= T("privacy_delete_warning_5", Empty) %><% Else %>会员积分、优惠券等资产将不予退还<% End If %></li>
            </ul>
        </div>

        <div class="delete-form-card">
            <h3><% If FEATURE_I18N Then %><%= T("privacy_delete_confirm_title", Empty) %><% Else %>确认账户注销<% End If %></h3>
            <form method="post" class="delete-form" id="deleteForm">
                <%= GetCSRFTokenField() %>
                <input type="hidden" name="action" value="request_delete">
                
                <div class="form-group">
                    <label for="password"><% If FEATURE_I18N Then %><%= T("privacy_delete_password_label", Empty) %><% Else %>请输入密码以确认<% End If %></label>
                    <input type="password" id="password" name="password" required 
                           placeholder="<% If FEATURE_I18N Then %><%= T("privacy_delete_password_placeholder", Empty) %><% Else %>输入当前密码<% End If %>">
                </div>

                <div class="form-actions">
                    <button type="submit" class="btn btn-danger btn-lg" id="btnDeleteSubmit" data-confirm="<% If FEATURE_I18N Then %><%= Server.HTMLEncode(T("privacy_delete_confirm_dialog", Empty)) %><% Else %>确认提交注销请求？此操作有30天冷静期。<% End If %>">
                        <i class="fas fa-trash-alt"></i> <% If FEATURE_I18N Then %><%= T("privacy_delete_submit", Empty) %><% Else %>提交注销请求<% End If %>
                    </button>
                    <a href="settings.asp" class="btn btn-secondary">
                        <i class="fas fa-arrow-left"></i> <% If FEATURE_I18N Then %><%= T("privacy_delete_back", Empty) %><% Else %>返回设置<% End If %>
                    </a>
                </div>
            </form>
        </div>
        <% End If %>

        <div class="delete-footer">
            <p><i class="fas fa-question-circle"></i> <% If FEATURE_I18N Then %><%= T("privacy_delete_questions", Empty) %><% Else %>有疑问？请查看我们的 <% End If %><a href="privacy.asp"><% If FEATURE_I18N Then %><%= T("privacy_delete_privacy_link", Empty) %><% Else %>隐私政策<% End If %></a></p>
        </div>
    </div>
</div>

<style>
.delete-page { padding: 40px 0; max-width: 700px; margin: 0 auto; }
.delete-header { text-align: center; margin-bottom: 32px; }
.delete-header h1 { font-size: 2rem; color: #c53030; margin-bottom: 8px; }
.delete-header p { color: #718096; }
.delete-warning { background: #fffff0; border: 1px solid #faf089; border-radius: 12px; padding: 24px; margin-bottom: 24px; }
.warning-icon { text-align: center; font-size: 2.5rem; color: #d69e2e; margin-bottom: 12px; }
.delete-warning h3 { color: #744210; margin-bottom: 12px; }
.warning-list { padding-left: 20px; }
.warning-list li { color: #718096; line-height: 1.8; margin-bottom: 6px; }
.delete-form-card { background: #fff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 24px; }
.delete-form-card h3 { color: #2d3748; margin-bottom: 20px; }
.form-group { margin-bottom: 20px; }
.form-group label { display: block; font-weight: 500; color: #4a5568; margin-bottom: 8px; }
.form-group input { width: 100%; padding: 12px; border: 1px solid #e2e8f0; border-radius: 8px; font-size: 16px; }
.form-actions { display: flex; gap: 12px; flex-wrap: wrap; }
.btn-danger { background: #e53e3e; color: #fff; border: none; }
.btn-danger:hover { background: #c53030; }
.btn-success { background: #38a169; color: #fff; border: none; }
.btn-success:hover { background: #2f855a; }
.delete-status-card { text-align: center; padding: 40px 24px; background: #fff; border: 1px solid #e2e8f0; border-radius: 12px; }
.status-icon { font-size: 4rem; color: #d69e2e; margin-bottom: 16px; }
.delete-status-card h2 { color: #2d3748; margin-bottom: 12px; }
.delete-status-card p { color: #718096; margin-bottom: 24px; line-height: 1.8; }
.cooling-timeline { margin: 20px 0; }
.cooling-info { display: inline-block; text-align: left; background: #f7fafc; padding: 16px 24px; border-radius: 8px; }
.cooling-info p { margin: 4px 0; color: #4a5568; font-size: 0.95rem; }
.cooling-note { color: #718096; font-size: 0.9rem; }
.cancel-form { margin-top: 20px; }
.delete-footer { text-align: center; margin-top: 32px; padding-top: 20px; border-top: 1px solid #e2e8f0; }
.delete-footer p { color: #a0aec0; }
.delete-footer a { color: #667eea; }
.alert-error { background: #fff5f5; border: 1px solid #fed7d7; color: #c53030; padding: 12px 16px; border-radius: 8px; margin-bottom: 20px; }
.alert-success { background: #f0fff4; border: 1px solid #c6f6d5; color: #276749; padding: 12px 16px; border-radius: 8px; margin-bottom: 20px; }
@media (max-width: 640px) {
    .form-actions { flex-direction: column; }
    .form-actions .btn { width: 100%; text-align: center; }
}
</style>

<script nonce="<%= Session("csp_nonce") %>">
(function() {
    var btn = document.getElementById('btnDeleteSubmit');
    if (btn) {
        btn.addEventListener('click', function(e) {
            var msg = this.getAttribute('data-confirm') || '确认提交注销请求？';
            if (!confirm(msg)) { e.preventDefault(); return false; }
        });
    }
})();
</script>

<!--#include file="../includes/footer.asp"-->
<%
Call CloseConnection()
%>
