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

' 获取管理员ID
Dim editAdminId
editAdminId = Request.QueryString("id")

' 验证ID
If editAdminId = "" Or Not IsNumeric(editAdminId) Then
    Response.Write "<script>alert('无效的管理员ID'); location.href='admins.asp';</script>"
    Response.End
End If

' 获取要编辑的管理员信息
Dim rsAdmin
Set rsAdmin = ExecuteQuery("SELECT a.*, r.RoleName FROM AdminUsers a LEFT JOIN AdminRoles r ON a.RoleID = r.RoleID WHERE a.AdminID = " & CLng(editAdminId))

If rsAdmin Is Nothing Or rsAdmin.EOF Then
    Response.Write "<script>alert('管理员不存在'); location.href='admins.asp';</script>"
    Response.End
End If

' 获取角色列表
Dim rsRoles
Set rsRoles = ExecuteQuery("SELECT * FROM AdminRoles ORDER BY RoleID")

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
        Dim fullName, email, department, roleId, isLocked, isActive
        fullName = SafeSQL(Trim(Request.Form("full_name")))
        email = SafeSQL(Trim(Request.Form("email")))
        department = SafeSQL(Trim(Request.Form("department")))
        roleId = Request.Form("role_id")
        isLocked = Request.Form("is_locked")
        isActive = Request.Form("is_active")
        
        ' 验证必填字段
        If fullName = "" Then
            errorMsg = "请输入姓名"
        ElseIf Not IsNumeric(roleId) Then
            errorMsg = "请选择角色"
        Else
            ' 构建更新SQL
            Dim updateSql
            updateSql = "UPDATE AdminUsers SET " & _
                "FullName = '" & fullName & "', " & _
                "Email = '" & email & "', " & _
                "Department = '" & department & "', " & _
                "RoleID = " & CLng(roleId) & ", " & _
                "IsLocked = " & IIF(isLocked = "1", "1", "0") & ", " & _
                "IsActive = " & IIF(isActive = "1", "1", "0") & _
                " WHERE AdminID = " & CLng(editAdminId)
            
            If ExecuteNonQuery(updateSql) Then
                Call LogAdminAction("编辑管理员", "system", "AdminUsers", CStr(editAdminId), fullName)
                successMsg = "管理员信息已更新"
                ' 刷新数据
                Set rsAdmin = ExecuteQuery("SELECT a.*, r.RoleName FROM AdminUsers a LEFT JOIN AdminRoles r ON a.RoleID = r.RoleID WHERE a.AdminID = " & CLng(editAdminId))
            Else
                errorMsg = "更新失败：" & Session("LastDBError")
            End If
        End If
    End If
End If
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>编辑管理员 - 系统管理中心</title>
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
        .readonly-field { background: #3a3a3a; color: #888; }
        .info-row { display: flex; padding: 12px 0; border-bottom: 1px solid rgba(255,255,255,0.06); }
        .info-label { width: 120px; color: #888; font-weight: 500; }
        .info-value { flex: 1; color: #e0e0e0; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-user-edit"></i> 编辑管理员</h2>
            <div class="breadcrumb">
                <a href="index.asp">系统中心</a> / <a href="admins.asp">管理员管理</a> / <span>编辑管理员</span>
            </div>
        </div>
        
        <div class="form-container">
            <% If successMsg <> "" Then %>
            <div class="alert alert-success">
                <i class="fas fa-check-circle"></i> <%= successMsg %>
            </div>
            <% End If %>
            
            <% If errorMsg <> "" Then %>
            <div class="alert alert-error">
                <i class="fas fa-exclamation-circle"></i> <%= errorMsg %>
            </div>
            <% End If %>
            
            <div class="info-row">
                <div class="info-label">用户名</div>
                <div class="info-value"><%= HTMLEncode(rsAdmin("Username")) %></div>
            </div>
            
            <form method="post" action="admin_edit.asp?id=<%= editAdminId %>">
                <%= GetCSRFTokenField() %>
                
                <div class="form-group">
                    <label for="full_name" class="form-label">姓名 <span style="color: #f44336;">*</span></label>
                    <input type="text" id="full_name" name="full_name" class="form-control" value="<%= HTMLEncode(rsAdmin("FullName")) %>" required>
                </div>
                
                <div class="form-group">
                    <label for="email" class="form-label">邮箱</label>
                    <input type="email" id="email" name="email" class="form-control" value="<%= HTMLEncode(rsAdmin("Email")) %>">
                </div>
                
                <div class="form-group">
                    <label for="department" class="form-label">部门</label>
                    <input type="text" id="department" name="department" class="form-control" value="<%= HTMLEncode(rsAdmin("Department")) %>">
                </div>
                
                <div class="form-group">
                    <label for="role_id" class="form-label">角色 <span style="color: #f44336;">*</span></label>
                    <select id="role_id" name="role_id" class="form-select" required>
                        <%
                        Dim currentRoleId
                        If IsNull(rsAdmin("RoleID")) Or rsAdmin("RoleID") = "" Then
                            currentRoleId = 0
                        Else
                            currentRoleId = CInt(rsAdmin("RoleID"))
                        End If
                        %>
                        <% If Not rsRoles Is Nothing Then %>
                        <% Do While Not rsRoles.EOF %>
                        <option value="<%= rsRoles("RoleID") %>" <%= IIF(CInt(rsRoles("RoleID")) = currentRoleId, "selected", "") %>><%= HTMLEncode(rsRoles("RoleName")) %></option>
                        <% rsRoles.MoveNext %>
                        <% Loop %>
                        <% rsRoles.Close %>
                        <% End If %>
                    </select>
                </div>
                
                <div class="form-group">
                    <div class="checkbox-group">
                        <input type="checkbox" id="is_active" name="is_active" value="1" <%= IIF(rsAdmin("IsActive") = True, "checked", "") %>>
                        <label for="is_active">账户启用</label>
                    </div>
                </div>
                
                <div class="form-group">
                    <div class="checkbox-group">
                        <input type="checkbox" id="is_locked" name="is_locked" value="1" <%= IIF(rsAdmin("IsLocked") = True, "checked", "") %>>
                        <label for="is_locked">账户锁定</label>
                    </div>
                </div>
                
                <div class="form-actions">
                    <button type="submit" class="admin-btn admin-btn-primary">
                        <i class="fas fa-save"></i> 保存修改
                    </button>
                    <a href="admins.asp" class="admin-btn admin-btn-outline">
                        <i class="fas fa-arrow-left"></i> 返回列表
                    </a>
                </div>
            </form>
        </div>
    </div>
</body>
</html>
<%
If Not rsAdmin Is Nothing Then
    rsAdmin.Close
    Set rsAdmin = Nothing
End If
Call CloseConnection()
%>
