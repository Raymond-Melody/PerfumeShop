<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
Call OpenConnection()

Function GetScalar(sql)
    Dim rs, val : val = ""
    On Error Resume Next
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        If Not rs Is Nothing Then
            If Not rs.EOF Then
                val = rs(0)
                rs.Close
            End If
        End If
    Else
        Err.Clear
    End If
    Set rs = Nothing : GetScalar = val
End Function

' 自动创建 SiteSettings 表（如不存在）
On Error Resume Next
conn.Execute "IF OBJECT_ID('SiteSettings','U') IS NULL BEGIN CREATE TABLE SiteSettings (SettingID INT IDENTITY(1,1) PRIMARY KEY, SiteName NVARCHAR(100) DEFAULT '香氛定制', SiteDescription NVARCHAR(500) DEFAULT '', CreatedAt DATETIME DEFAULT GETDATE()) END"
If Err.Number <> 0 Then Err.Clear
' 确保至少有一条记录
conn.Execute "IF NOT EXISTS (SELECT 1 FROM SiteSettings) INSERT INTO SiteSettings DEFAULT VALUES"
If Err.Number <> 0 Then Err.Clear

' 自动创建 SiteSettings 安全策略配置字段
On Error Resume Next
conn.Execute "SELECT Security_PasswordMinLength FROM SiteSettings WHERE 1=0"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "ALTER TABLE SiteSettings ADD Security_PasswordMinLength INT DEFAULT 8"
End If
conn.Execute "SELECT Security_SessionTimeout FROM SiteSettings WHERE 1=0"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "ALTER TABLE SiteSettings ADD Security_SessionTimeout INT DEFAULT 30"
End If
conn.Execute "SELECT Security_LoginMaxAttempts FROM SiteSettings WHERE 1=0"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "ALTER TABLE SiteSettings ADD Security_LoginMaxAttempts INT DEFAULT 5"
End If
conn.Execute "SELECT Security_MFAEnabled FROM SiteSettings WHERE 1=0"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "ALTER TABLE SiteSettings ADD Security_MFAEnabled BIT DEFAULT 0"
End If
conn.Execute "SELECT Security_LockoutMinutes FROM SiteSettings WHERE 1=0"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "ALTER TABLE SiteSettings ADD Security_LockoutMinutes INT DEFAULT 30"
End If
On Error GoTo 0

' 获取当前设置
Dim siteSettingsRs
Set siteSettingsRs = ExecuteQuery("SELECT TOP 1 * FROM SiteSettings")

Dim pwMinLen, sessionTimeout, loginMaxAttempts, mfaEnabled, lockoutMinutes
pwMinLen = 8 : sessionTimeout = 30 : loginMaxAttempts = 5 : mfaEnabled = 0 : lockoutMinutes = 30

If Not siteSettingsRs Is Nothing Then
    If Not siteSettingsRs.EOF Then
        On Error Resume Next
        Dim tmpVal
        tmpVal = siteSettingsRs("Security_PasswordMinLength")
        If Err.Number = 0 And Not IsNull(tmpVal) Then pwMinLen = CInt(tmpVal)
        Err.Clear
        tmpVal = siteSettingsRs("Security_SessionTimeout")
        If Err.Number = 0 And Not IsNull(tmpVal) Then sessionTimeout = CInt(tmpVal)
        Err.Clear
        tmpVal = siteSettingsRs("Security_LoginMaxAttempts")
        If Err.Number = 0 And Not IsNull(tmpVal) Then loginMaxAttempts = CInt(tmpVal)
        Err.Clear
        tmpVal = siteSettingsRs("Security_MFAEnabled")
        If Err.Number = 0 And Not IsNull(tmpVal) Then mfaEnabled = CInt(tmpVal)
        Err.Clear
        tmpVal = siteSettingsRs("Security_LockoutMinutes")
        If Err.Number = 0 And Not IsNull(tmpVal) Then lockoutMinutes = CInt(tmpVal)
        Err.Clear
        On Error GoTo 0
    End If
End If

' 获取当前管理员信息
Dim rsAdmin, adminId
adminId = Session("AdminID")

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim formAction : formAction = Request.Form("form_action")
    
    If formAction = "profile" Then
        ' 处理个人信息修改
        Dim username, email
        username = SafeSQL(Trim(Request.Form("username")))
        email = SafeSQL(Trim(Request.Form("email")))
        
        Dim updateSql
        updateSql = "UPDATE AdminUsers SET Username = '" & username & "', Email = '" & email & "' WHERE AdminID = " & adminId
        
        If ExecuteNonQuery(updateSql) Then
            Session("AdminUsername") = username
            Response.Write "<script>alert('设置已保存'); location.href='settings.asp';</script>"
        Else
            Dim errorMsg
            errorMsg = Replace(Session("LastDBError"), "'", "'")
            Response.Write "<script>alert('保存失败：" & errorMsg & "');</script>"
        End If
    
    ElseIf formAction = "security" Then
        ' 保存安全策略
        pwMinLen = CInt(Request.Form("pwMinLen"))
        sessionTimeout = CInt(Request.Form("sessionTimeout"))
        loginMaxAttempts = CInt(Request.Form("loginMaxAttempts"))
        mfaEnabled = IIf(Request.Form("mfaEnabled")="1", 1, 0)
        lockoutMinutes = CInt(Request.Form("lockoutMinutes"))
        
        conn.Execute "UPDATE SiteSettings SET Security_PasswordMinLength=" & pwMinLen & _
            ", Security_SessionTimeout=" & sessionTimeout & _
            ", Security_LoginMaxAttempts=" & loginMaxAttempts & _
            ", Security_MFAEnabled=" & mfaEnabled & _
            ", Security_LockoutMinutes=" & lockoutMinutes
        
        Response.Write "<script>alert('安全策略已保存'); location.href='settings.asp';</script>"
    End If
End If

' 获取管理员信息
Set rsAdmin = ExecuteQuery("SELECT * FROM AdminUsers WHERE AdminID = " & adminId & " AND IsActive = 1")
If rsAdmin Is Nothing Or rsAdmin.EOF Then
    Response.Redirect "../login.asp"
    Response.End
End If
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>管理员设置 - 香氛定制电商网站</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; }
        .admin-form-control[style*="background: #f8f9fa"] { background: #2d2d44 !important; color: #e0e0e0 !important; }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="admin-card">
            <div class="admin-card-header">
                <h2 class="admin-card-title"><i class="fas fa-cog"></i> 管理员设置</h2>
            </div>
            
            <div class="admin-card-body">
                <form method="post" action="settings.asp" class="form-horizontal">
                    <input type="hidden" name="form_action" value="profile">
                    <div class="admin-form-group">
                        <label for="username" class="admin-form-label">管理员用户名</label>
                        <input type="text" id="username" name="username" value="<%= HTMLEncode(rsAdmin("Username")) %>" class="admin-form-control" required>
                    </div>
                    
                    <div class="admin-form-group">
                        <label for="email" class="admin-form-label">邮箱地址</label>
                        <input type="email" id="email" name="email" value="<%= HTMLEncode(rsAdmin("Email")) %>" class="admin-form-control">
                    </div>
                    
                    <div class="admin-form-group">
                        <label class="admin-form-label">注册时间</label>
                        <div class="admin-form-control" style="background: #f8f9fa; border: none; padding: 12px 15px;">
                            <%= FormatDateField(rsAdmin("CreatedAt")) %>
                        </div>
                    </div>
                    
                    <div class="admin-form-group">
                        <label class="admin-form-label">最后登录</label>
                        <div class="admin-form-control" style="background: #f8f9fa; border: none; padding: 12px 15px;">
                            <%= FormatDateField(rsAdmin("LastLogin")) %>
                        </div>
                    </div>
                    
                    <div class="admin-form-group">
                        <label class="admin-form-label">账户状态</label>
                        <div class="admin-form-control" style="background: #f8f9fa; border: none; padding: 12px 15px;">
                            <%= IIf(rsAdmin("IsActive") <> 0, "<span class='status-badge status-paid'>启用</span>", "<span class='status-badge status-cancelled'>禁用</span>") %>
                        </div>
                    </div>
                    
                    <div class="form-actions" style="text-align: center; margin-top: 30px; padding: 0;">
                        <button type="submit" class="admin-btn admin-btn-primary admin-btn-lg">
                            <i class="fas fa-save"></i> 保存设置
                        </button>
                        <a href="index.asp" class="admin-btn admin-btn-outline admin-btn-lg">
                            <i class="fas fa-arrow-left"></i> 返回管理
                        </a>
                    </div>
                </form>
            </div>
        </div>
        
        <!-- 修改密码部分 -->
        <div class="admin-card">
            <div class="admin-card-header">
                <h2 class="admin-card-title"><i class="fas fa-key"></i> 修改密码</h2>
            </div>
            
            <div class="admin-card-body">
                <form method="post" action="../change_password.asp" class="form-horizontal">
                    <div class="admin-form-row">
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label for="current_password" class="admin-form-label">当前密码</label>
                                <input type="password" id="current_password" name="current_password" class="admin-form-control" required>
                            </div>
                        </div>
                    </div>
                    
                    <div class="admin-form-row">
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label for="new_password" class="admin-form-label">新密码</label>
                                <input type="password" id="new_password" name="new_password" class="admin-form-control" required>
                            </div>
                        </div>
                        
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label for="confirm_password" class="admin-form-label">确认新密码</label>
                                <input type="password" id="confirm_password" name="confirm_password" class="admin-form-control" required>
                            </div>
                        </div>
                    </div>
                    
                    <div class="form-actions" style="text-align: center; margin-top: 20px; padding: 0;">
                        <button type="submit" class="admin-btn admin-btn-primary">
                            <i class="fas fa-key"></i> 修改密码
                        </button>
                    </div>
                </form>
            </div>
        </div>
        
        <!-- 安全策略配置 -->
        <div class="admin-card">
            <div class="admin-card-header">
                <h2 class="admin-card-title"><i class="fas fa-shield-alt"></i> 安全策略配置</h2>
            </div>
            <div class="admin-card-body">
                <form method="post" action="settings.asp" class="form-horizontal">
                    <input type="hidden" name="form_action" value="security">
                    
                    <div class="admin-form-group">
                        <label for="pwMinLen" class="admin-form-label">密码最小长度</label>
                        <input type="number" id="pwMinLen" name="pwMinLen" value="<%= pwMinLen %>" class="admin-form-control" min="6" max="32" style="max-width:150px;">
                        <small style="color:#888;">建议至少8位，包含大小写字母+数字+特殊字符</small>
                    </div>
                    
                    <div class="admin-form-group">
                        <label for="sessionTimeout" class="admin-form-label">会话超时（分钟）</label>
                        <input type="number" id="sessionTimeout" name="sessionTimeout" value="<%= sessionTimeout %>" class="admin-form-control" min="5" max="480" style="max-width:150px;">
                        <small style="color:#888;">超过此时间无操作将自动登出</small>
                    </div>
                    
                    <div class="admin-form-group">
                        <label for="loginMaxAttempts" class="admin-form-label">登录失败上限</label>
                        <input type="number" id="loginMaxAttempts" name="loginMaxAttempts" value="<%= loginMaxAttempts %>" class="admin-form-control" min="1" max="20" style="max-width:150px;">
                        <small style="color:#888;">达到上限后IP将被临时封锁</small>
                    </div>
                    
                    <div class="admin-form-group">
                        <label for="lockoutMinutes" class="admin-form-label">封锁时长（分钟）</label>
                        <input type="number" id="lockoutMinutes" name="lockoutMinutes" value="<%= lockoutMinutes %>" class="admin-form-control" min="5" max="1440" style="max-width:150px;">
                        <small style="color:#888;">登录失败超限后的IP封锁时间</small>
                    </div>
                    
                    <div class="admin-form-group">
                        <label class="admin-form-label">MFA双因素认证</label>
                        <div style="padding: 10px 0;">
                            <label style="display:inline-flex;align-items:center;gap:8px;color:#e0e0e0;cursor:pointer;">
                                <input type="checkbox" name="mfaEnabled" value="1" <%= IIf(mfaEnabled=1,"checked","") %>>
                                启用双因素认证（当前为测试功能）
                            </label>
                        </div>
                        <small style="color:#888;">启用后管理员登录需要额外验证码</small>
                    </div>
                    
                    <div class="form-actions" style="text-align: center; margin-top: 30px; padding: 0;">
                        <button type="submit" class="admin-btn admin-btn-primary admin-btn-lg">
                            <i class="fas fa-shield-alt"></i> 保存安全策略
                        </button>
                    </div>
                </form>
            </div>
        </div>
    </div>
</body>
</html>
<%
If Not rsAdmin Is Nothing Then
    rsAdmin.Close
    Set rsAdmin = Nothing
End If
If Not siteSettingsRs Is Nothing Then
    siteSettingsRs.Close
    Set siteSettingsRs = Nothing
End If
Call CloseConnection()
%>