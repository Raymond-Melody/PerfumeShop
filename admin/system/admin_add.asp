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
        Dim username, password, confirmPassword, fullName, email, department, roleId, isActive
        username = SafeSQL(Trim(Request.Form("username")))
        password = Trim(Request.Form("password"))
        confirmPassword = Trim(Request.Form("confirm_password"))
        fullName = SafeSQL(Trim(Request.Form("full_name")))
        email = SafeSQL(Trim(Request.Form("email")))
        department = SafeSQL(Trim(Request.Form("department")))
        roleId = Request.Form("role_id")
        isActive = Request.Form("is_active")
        
        ' 验证必填字段
        If username = "" Then
            errorMsg = "请输入用户名"
        ElseIf password = "" Then
            errorMsg = "请输入密码"
        ElseIf Len(password) < 6 Then
            errorMsg = "密码长度至少为6位"
        ElseIf password <> confirmPassword Then
            errorMsg = "两次输入的密码不一致"
        ElseIf fullName = "" Then
            errorMsg = "请输入姓名"
        ElseIf Not IsNumeric(roleId) Then
            errorMsg = "请选择角色"
        Else
            ' 检查用户名是否已存在
            Dim rsCheck
            Set rsCheck = ExecuteQuery("SELECT AdminID FROM AdminUsers WHERE Username = '" & username & "'")
            
            If Not rsCheck Is Nothing And Not rsCheck.EOF Then
                errorMsg = "用户名已存在，请使用其他用户名"
                rsCheck.Close
                Set rsCheck = Nothing
            Else
                If Not rsCheck Is Nothing Then
                    rsCheck.Close
                    Set rsCheck = Nothing
                End If
                
                ' 生成密码哈希
                Dim hashedPassword
                hashedPassword = GenerateSimpleHash(password)
                
                ' 构建INSERT SQL
                Dim insertSql
                insertSql = "INSERT INTO AdminUsers (Username, PasswordHash, FullName, Email, Department, RoleID, IsActive, IsLocked, CreatedAt) VALUES (" & _
                    "'" & username & "', " & _
                    "'" & SafeSQL(hashedPassword) & "', " & _
                    "'" & fullName & "', " & _
                    "'" & email & "', " & _
                    "'" & department & "', " & _
                    CLng(roleId) & ", " & _
                    IIF(isActive = "1", "True", "False") & ", " & _
                    "False, " & _
                    "GETDATE())"
                
                If ExecuteNonQuery(insertSql) Then
                    Dim newAdminId
                    newAdminId = GetLastInsertID("AdminUsers")
                    Call LogAdminAction("添加管理员", "system", "AdminUsers", CStr(newAdminId), fullName)
                    successMsg = "管理员添加成功"
                Else
                    errorMsg = "添加失败：" & Session("LastDBError")
                End If
            End If
        End If
    End If
End If

' 获取角色列表
Dim rsRoles
Set rsRoles = ExecuteQuery("SELECT * FROM AdminRoles ORDER BY RoleID")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>添加管理员 - 系统管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; }
        .form-container { max-width: 600px; margin: 0 auto; background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; padding: 30px; box-shadow: 0 2px 10px rgba(0,0,0,0.08); }
        .form-group { margin-bottom: 20px; }
        .form-label { display: block; margin-bottom: 8px; font-weight: 500; color: #e0e0e0; }
        .form-control { width: 100%; padding: 12px 15px; border: 1px solid rgba(255,255,255,0.15); border-radius: 8px; font-size: 14px; box-sizing: border-box; background: #2d2d44; color: #e0e0e0; }
        .form-control:focus { outline: none; border-color: #00bcd4; }
        .form-select { width: 100%; padding: 12px 15px; border: 1px solid rgba(255,255,255,0.15); border-radius: 8px; font-size: 14px; background: #2d2d44; color: #e0e0e0; }
        .checkbox-group { display: flex; align-items: center; gap: 10px; }
        .checkbox-group input[type="checkbox"] { width: 18px; height: 18px; }
        .form-actions { display: flex; gap: 15px; justify-content: center; margin-top: 30px; }
        .alert { padding: 12px 20px; border-radius: 8px; margin-bottom: 20px; }
        .alert-success { background: rgba(46, 125, 50, 0.2); color: #81c784; border: 1px solid rgba(46, 125, 50, 0.3); }
        .alert-error { background: rgba(198, 40, 40, 0.2); color: #ef9a9a; border: 1px solid rgba(198, 40, 40, 0.3); }
        .password-strength { margin-top: 8px; height: 4px; background: #3a3a3a; border-radius: 2px; overflow: hidden; }
        .password-strength-bar { height: 100%; width: 0; transition: all 0.3s; }
        .strength-weak { background: #f44336; width: 33%; }
        .strength-medium { background: #ff9800; width: 66%; }
        .strength-strong { background: #4caf50; width: 100%; }
        .required { color: #f44336; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-user-plus"></i> 添加管理员</h2>
            <div class="breadcrumb">
                <a href="index.asp">系统中心</a> / <a href="admins.asp">管理员管理</a> / <span>添加管理员</span>
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
                <a href="admin_add.asp" class="admin-btn admin-btn-outline">
                    <i class="fas fa-plus"></i> 继续添加
                </a>
            </div>
            <% Else %>
            
            <% If errorMsg <> "" Then %>
            <div class="alert alert-error">
                <i class="fas fa-exclamation-circle"></i> <%= errorMsg %>
            </div>
            <% End If %>
            
            <form method="post" action="admin_add.asp">
                <%= GetCSRFTokenField() %>
                
                <div class="form-group">
                    <label for="username" class="form-label">用户名 <span class="required">*</span></label>
                    <input type="text" id="username" name="username" class="form-control" value="<%= HTMLEncode(Request.Form("username")) %>" required maxlength="50">
                </div>
                
                <div class="form-group">
                    <label for="password" class="form-label">密码 <span class="required">*</span></label>
                    <input type="password" id="password" name="password" class="form-control" required minlength="6" oninput="checkStrength(this.value)">
                    <div class="password-strength">
                        <div class="password-strength-bar" id="strengthBar"></div>
                    </div>
                </div>
                
                <div class="form-group">
                    <label for="confirm_password" class="form-label">确认密码 <span class="required">*</span></label>
                    <input type="password" id="confirm_password" name="confirm_password" class="form-control" required minlength="6">
                </div>
                
                <div class="form-group">
                    <label for="full_name" class="form-label">姓名 <span class="required">*</span></label>
                    <input type="text" id="full_name" name="full_name" class="form-control" value="<%= HTMLEncode(Request.Form("full_name")) %>" required maxlength="50">
                </div>
                
                <div class="form-group">
                    <label for="email" class="form-label">邮箱</label>
                    <input type="email" id="email" name="email" class="form-control" value="<%= HTMLEncode(Request.Form("email")) %>" maxlength="100">
                </div>
                
                <div class="form-group">
                    <label for="department" class="form-label">部门</label>
                    <input type="text" id="department" name="department" class="form-control" value="<%= HTMLEncode(Request.Form("department")) %>" maxlength="50">
                </div>
                
                <div class="form-group">
                    <label for="role_id" class="form-label">角色 <span class="required">*</span></label>
                    <select id="role_id" name="role_id" class="form-select" required>
                        <option value="">请选择角色</option>
                        <% If Not rsRoles Is Nothing Then %>
                        <% Do While Not rsRoles.EOF %>
                        <option value="<%= rsRoles("RoleID") %>" <%= IIF(CStr(rsRoles("RoleID")) = CStr(Request.Form("role_id")), "selected", "") %>><%= HTMLEncode(rsRoles("RoleName")) %></option>
                        <% rsRoles.MoveNext %>
                        <% Loop %>
                        <% rsRoles.Close %>
                        <% End If %>
                    </select>
                </div>
                
                <div class="form-group">
                    <div class="checkbox-group">
                        <input type="checkbox" id="is_active" name="is_active" value="1" checked>
                        <label for="is_active">账户启用</label>
                    </div>
                </div>
                
                <div class="form-actions">
                    <button type="submit" class="admin-btn admin-btn-primary">
                        <i class="fas fa-save"></i> 保存
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
If Not rsRoles Is Nothing Then
    Set rsRoles = Nothing
End If
Call CloseConnection()
%>
