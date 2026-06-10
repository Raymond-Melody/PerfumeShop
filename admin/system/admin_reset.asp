<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<!--#include file="../../includes/password_utils.asp"-->
<%
Call OpenConnection()

' 获取管理员ID
Dim resetAdminId
resetAdminId = Request.QueryString("id")

' 验证ID
If resetAdminId = "" Or Not IsNumeric(resetAdminId) Then
    Response.Write "<script>alert('无效的管理员ID'); location.href='admins.asp';</script>"
    Response.End
End If

' 获取要重置密码的管理员信息
Dim rsAdmin
Set rsAdmin = ExecuteQuery("SELECT AdminID, Username, FullName, Email FROM AdminUsers WHERE AdminID = " & CLng(resetAdminId))

If rsAdmin Is Nothing Or rsAdmin.EOF Then
    Response.Write "<script>alert('管理员不存在'); location.href='admins.asp';</script>"
    Response.End
End If

' 不能重置自己的密码（为了安全，需要通过修改密码功能）
If CInt(resetAdminId) = CInt(Session("AdminID")) Then
    Response.Write "<script>alert('不能通过此功能重置自己的密码，请使用修改密码功能'); location.href='admins.asp';</script>"
    Response.End
End If

' 处理表单提交
Dim successMsg, errorMsg
successMsg = ""
errorMsg = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' 验证CSRF令牌
    If Not ValidateCSRFToken() Then
        errorMsg = "安全验证失败，请刷新页面后重试"
    Else
        ' 获取表单数据
        Dim newPassword, confirmPassword
        newPassword = Trim(Request.Form("new_password"))
        confirmPassword = Trim(Request.Form("confirm_password"))
        
        ' 验证密码
        If newPassword = "" Then
            errorMsg = "请输入新密码"
        ElseIf Len(newPassword) < 6 Then
            errorMsg = "密码长度至少为6位"
        ElseIf newPassword <> confirmPassword Then
            errorMsg = "两次输入的密码不一致"
        Else
            ' 生成密码哈希
            Dim hashedPassword
            hashedPassword = GenerateSimpleHash(newPassword)
            
            ' 构建更新SQL
            Dim updateSql
            updateSql = "UPDATE AdminUsers SET PasswordHash = '" & SafeSQL(hashedPassword) & "', ResetToken = NULL, ResetTokenExpiry = NULL WHERE AdminID = " & CLng(resetAdminId)
            
            If ExecuteNonQuery(updateSql) Then
                Call LogAdminAction("重置管理员密码", "system", "AdminUsers", CStr(resetAdminId), rsAdmin("FullName"))
                successMsg = "密码重置成功"
            Else
                errorMsg = "密码重置失败：" & Session("LastDBError")
            End If
        End If
    End If
End If

' 生成随机密码函数（用于显示建议密码）
Function GenerateRandomPassword()
    Dim chars, password, i, charIndex
    chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*"
    password = ""
    Randomize
    For i = 1 To 10
        charIndex = Int(Rnd * Len(chars)) + 1
        password = password & Mid(chars, charIndex, 1)
    Next
    GenerateRandomPassword = password
End Function
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>重置密码 - 系统管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; }
        .form-container { max-width: 500px; margin: 0 auto; background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; padding: 30px; box-shadow: 0 2px 10px rgba(0,0,0,0.08); }
        .admin-info { background: #2d2d44; border-radius: 8px; padding: 20px; margin-bottom: 25px; border: 1px solid rgba(255,255,255,0.06); }
        .admin-info-row { display: flex; padding: 8px 0; }
        .admin-info-label { width: 80px; color: #888; font-weight: 500; }
        .admin-info-value { flex: 1; color: #e0e0e0; font-weight: 600; }
        .form-group { margin-bottom: 20px; }
        .form-label { display: block; margin-bottom: 8px; font-weight: 500; color: #e0e0e0; }
        .form-control { width: 100%; padding: 12px 15px; border: 1px solid rgba(255,255,255,0.15); border-radius: 8px; font-size: 14px; box-sizing: border-box; background: #2d2d44; color: #e0e0e0; }
        .form-control:focus { outline: none; border-color: #00bcd4; }
        .password-strength { margin-top: 8px; height: 4px; background: #3a3a3a; border-radius: 2px; overflow: hidden; }
        .password-strength-bar { height: 100%; width: 0; transition: all 0.3s; }
        .strength-weak { background: #f44336; width: 33%; }
        .strength-medium { background: #ff9800; width: 66%; }
        .strength-strong { background: #4caf50; width: 100%; }
        .form-actions { display: flex; gap: 15px; justify-content: center; margin-top: 30px; }
        .alert { padding: 12px 20px; border-radius: 8px; margin-bottom: 20px; }
        .alert-success { background: rgba(46, 125, 50, 0.2); color: #81c784; border: 1px solid rgba(46, 125, 50, 0.3); }
        .alert-error { background: rgba(198, 40, 40, 0.2); color: #ef9a9a; border: 1px solid rgba(198, 40, 40, 0.3); }
        .suggested-password { background: rgba(33, 150, 243, 0.1); border: 1px dashed rgba(33, 150, 243, 0.4); border-radius: 8px; padding: 15px; margin-bottom: 20px; text-align: center; }
        .suggested-password-label { color: #888; font-size: 12px; margin-bottom: 5px; }
        .suggested-password-value { font-family: monospace; font-size: 16px; color: #64b5f6; font-weight: 600; letter-spacing: 1px; }
        .warning-box { background: rgba(230, 81, 0, 0.15); border: 1px solid rgba(255, 152, 0, 0.3); border-radius: 8px; padding: 15px; margin-bottom: 20px; color: #ffb74d; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-key"></i> 重置管理员密码</h2>
            <div class="breadcrumb">
                <a href="index.asp">系统中心</a> / <a href="admins.asp">管理员管理</a> / <span>重置密码</span>
            </div>
        </div>
        
        <div class="form-container">
            <% If successMsg <> "" Then %>
            <div class="alert alert-success">
                <i class="fas fa-check-circle"></i> <%= successMsg %>
            </div>
            <div class="form-actions">
                <a href="admins.asp" class="admin-btn admin-btn-primary">
                    <i class="fas fa-arrow-left"></i> 返回列表
                </a>
            </div>
            <% Else %>
            
            <div class="warning-box">
                <i class="fas fa-exclamation-triangle"></i> 您正在为管理员 <strong><%= HTMLEncode(rsAdmin("FullName")) %></strong> 重置密码。此操作不可撤销。
            </div>
            
            <div class="admin-info">
                <div class="admin-info-row">
                    <div class="admin-info-label">用户名</div>
                    <div class="admin-info-value"><%= HTMLEncode(rsAdmin("Username")) %></div>
                </div>
                <div class="admin-info-row">
                    <div class="admin-info-label">姓名</div>
                    <div class="admin-info-value"><%= HTMLEncode(rsAdmin("FullName")) %></div>
                </div>
                <div class="admin-info-row">
                    <div class="admin-info-label">邮箱</div>
                    <div class="admin-info-value"><%= HTMLEncode(rsAdmin("Email")) %></div>
                </div>
            </div>
            
            <% If errorMsg <> "" Then %>
            <div class="alert alert-error">
                <i class="fas fa-exclamation-circle"></i> <%= errorMsg %>
            </div>
            <% End If %>
            
            <div class="suggested-password">
                <div class="suggested-password-label">建议密码（点击复制）</div>
                <div class="suggested-password-value">
                    <span id="suggestedPwd"><%= GenerateRandomPassword() %></span>
                    <button type="button" class="btn btn--info btn--sm" onclick="copyPassword()">
                        <i class="fas fa-copy"></i> 复制
                    </button>
                </div>
            </div>
            
            <form method="post" action="admin_reset.asp?id=<%= resetAdminId %>">
                <%= GetCSRFTokenField() %>
                
                <div class="form-group">
                    <label for="new_password" class="form-label">新密码 <span style="color: #f44336;">*</span></label>
                    <input type="password" id="new_password" name="new_password" class="form-control" required minlength="6" oninput="checkStrength(this.value)">
                    <div class="password-strength">
                        <div class="password-strength-bar" id="strengthBar"></div>
                    </div>
                </div>
                
                <div class="form-group">
                    <label for="confirm_password" class="form-label">确认密码 <span style="color: #f44336;">*</span></label>
                    <input type="password" id="confirm_password" name="confirm_password" class="form-control" required>
                </div>
                
                <div class="form-actions">
                    <button type="submit" class="admin-btn admin-btn-primary" onclick="return confirm('确定要重置该管理员的密码吗？')">
                        <i class="fas fa-key"></i> 确认重置
                    </button>
                    <a href="admins.asp" class="admin-btn admin-btn-outline">
                        <i class="fas fa-times"></i> 取消
                    </a>
                </div>
            </form>
            <% End If %>
        </div>
    </div>
    
    <script>
    function copyPassword() {
        var pwd = document.getElementById('suggestedPwd').innerText;
        navigator.clipboard.writeText(pwd).then(function() {
            document.getElementById('new_password').value = pwd;
            document.getElementById('confirm_password').value = pwd;
            checkStrength(pwd);
            alert('密码已复制并填入表单');
        });
    }
    
    function checkStrength(password) {
        var bar = document.getElementById('strengthBar');
        var strength = 0;
        
        if (password.length >= 6) strength++;
        if (password.length >= 10) strength++;
        if (/[a-z]/.test(password) && /[A-Z]/.test(password)) strength++;
        if (/[0-9]/.test(password)) strength++;
        if (/[^a-zA-Z0-9]/.test(password)) strength++;
        
        bar.className = 'password-strength-bar';
        if (strength <= 2) {
            bar.classList.add('strength-weak');
        } else if (strength <= 4) {
            bar.classList.add('strength-medium');
        } else {
            bar.classList.add('strength-strong');
        }
    }
    </script>
</body>
</html>
<%
If Not rsAdmin Is Nothing Then
    rsAdmin.Close
    Set rsAdmin = Nothing
End If
Call CloseConnection()
%>
