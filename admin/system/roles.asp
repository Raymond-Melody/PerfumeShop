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

' 自动创建 RolePermissions 表（操作级权限）
On Error Resume Next
conn.Execute "SELECT TOP 1 * FROM RolePermissions WHERE 1=0"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE RolePermissions (" & _
        "PermID INT IDENTITY(1,1) PRIMARY KEY," & _
        "RoleID INT NOT NULL," & _
        "ModuleCode NVARCHAR(50) NOT NULL," & _
        "CanView BIT DEFAULT 0," & _
        "CanCreate BIT DEFAULT 0," & _
        "CanEdit BIT DEFAULT 0," & _
        "CanDelete BIT DEFAULT 0," & _
        "CanExport BIT DEFAULT 0," & _
        "CanApprove BIT DEFAULT 0" & _
        ")"
End If

' 扩展 AdminRoles 表增加 ModuleAccess 字段
conn.Execute "SELECT ModuleAccess FROM AdminRoles WHERE 1=0"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "ALTER TABLE AdminRoles ADD ModuleAccess NVARCHAR(500)"
    ' 初始化现有角色
    conn.Execute "UPDATE AdminRoles SET ModuleAccess='operation,production,finance,content,semifinished,prodcenter,logistics' WHERE RoleCode='superadmin'"
    conn.Execute "UPDATE AdminRoles SET ModuleAccess='operation,content' WHERE RoleCode='operator'"
    conn.Execute "UPDATE AdminRoles SET ModuleAccess='production,finance' WHERE RoleCode='manager'"
End If
On Error GoTo 0

Dim msg : msg = ""

' 处理 POST - 更新操作级权限
If Request.ServerVariables("REQUEST_METHOD") = "POST" And Request.Form("action") = "save_perm" Then
    Dim permRoleID, permModule
    permRoleID = CInt(Request.Form("roleID"))
    permModule = Request.Form("moduleCode")
    
    ' 删除旧权限
    conn.Execute "DELETE FROM RolePermissions WHERE RoleID=" & permRoleID & " AND ModuleCode='" & Replace(permModule,"'","''") & "'"
    
    ' 插入新权限
    conn.Execute "INSERT INTO RolePermissions (RoleID, ModuleCode, CanView, CanCreate, CanEdit, CanDelete, CanExport, CanApprove) VALUES (" & _
        permRoleID & ", '" & permModule & "', " & _
        IIf(Request.Form("CanView")="1",1,0) & ", " & _
        IIf(Request.Form("CanCreate")="1",1,0) & ", " & _
        IIf(Request.Form("CanEdit")="1",1,0) & ", " & _
        IIf(Request.Form("CanDelete")="1",1,0) & ", " & _
        IIf(Request.Form("CanExport")="1",1,0) & ", " & _
        IIf(Request.Form("CanApprove")="1",1,0) & ")"
    msg = "权限已保存"
End If

' 模块列表
Dim modules : modules = Array("operation", "production", "finance", "content", "system", "purchase", "semifinished", "prodcenter", "logistics")
Dim moduleNames : moduleNames = Array("运营管理", "生产管理", "财务管理", "内容管理", "系统管理", "采购管理", "半成品生产", "产品生产", "物流管理")

' 获取角色列表
Dim rsRoles
Set rsRoles = ExecuteQuery("SELECT * FROM AdminRoles ORDER BY RoleID")

Call LogAdminAction("查看角色列表", "system", "AdminRoles", "", "")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>角色管理 - 系统管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; }
        .main-content { margin-left: 250px; padding: 30px; min-height: 100vh; }
        .page-header { margin-bottom: 25px; }
        .page-title { color: #fff; font-size: 24px; margin: 0 0 8px; }
        .breadcrumb { color: #888; font-size: 13px; }
        .breadcrumb a { color: #00bcd4; text-decoration: none; }
        .msg { padding: 12px 20px; border-radius: 8px; margin-bottom: 20px; }
        .msg-success { background: rgba(76,175,80,0.15); color: #4CAF50; border: 1px solid rgba(76,175,80,0.3); }
        .card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); margin-bottom: 20px; overflow: hidden; }
        .card-header { padding: 15px 20px; border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 16px; display: flex; justify-content: space-between; align-items: center; }
        .card-body { padding: 20px; }
        .roles-table { width: 100%; border-collapse: collapse; }
        .roles-table th { background: linear-gradient(135deg, #00bcd4, #00838f); color: white; padding: 12px; text-align: left; font-size: 13px; }
        .roles-table td { padding: 12px; border-bottom: 1px solid rgba(255,255,255,0.04); color: #e0e0e0; font-size: 13px; }
        .roles-table tr:hover { background: rgba(255,255,255,0.03); }
        .role-code { font-family: 'Consolas',monospace; background: rgba(255,255,255,0.08); padding: 4px 8px; border-radius: 4px; font-size: 12px; color: #00bcd4; }
        .permission-tag { display: inline-block; padding: 4px 10px; background: rgba(0,188,212,0.15); color: #00bcd4; border-radius: 12px; font-size: 11px; margin: 2px; border: 1px solid rgba(0,188,212,0.3); }
        .perm-check { display: inline-flex; align-items: center; gap: 4px; margin: 3px 8px 3px 0; font-size: 12px; }
        .perm-check input { accent-color: #00bcd4; }
        .modal-overlay { display: none; position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.7); z-index: 1000; justify-content: center; align-items: center; }
        .modal-overlay.show { display: flex; }
        .modal { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 16px; border: 1px solid rgba(255,255,255,0.1); padding: 30px; width: 90%; max-width: 600px; max-height: 80vh; overflow-y: auto; }
        .modal h3 { margin: 0 0 5px; font-size: 18px; color: #fff; }
        .modal .subtitle { font-size: 13px; color: #888; margin-bottom: 20px; }
        .perm-section { margin-bottom: 18px; padding: 12px; background: rgba(255,255,255,0.03); border-radius: 8px; }
        .perm-section strong { display: block; margin-bottom: 8px; color: #00bcd4; font-size: 14px; }
        .modal-actions { display: flex; gap: 10px; justify-content: flex-end; margin-top: 20px; }
        .module-tags { display: flex; flex-wrap: wrap; gap: 5px; }
        .module-tag { display: inline-block; padding: 3px 8px; border-radius: 10px; font-size: 11px; background: rgba(0,188,212,0.12); color: #80DEEA; }
        @media (max-width: 768px) { .main-content { margin-left: 0; } }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-user-tag"></i> 角色权限管理</h2>
            <div class="breadcrumb">
                <a href="index.asp">系统中心</a> / <span>角色管理</span>
            </div>
        </div>

        <% If msg <> "" Then %>
        <div class="msg msg-success"><i class="fas fa-check-circle"></i> <%= msg %></div>
        <% End If %>
        
        <div class="card">
            <div class="card-header"><i class="fas fa-list"></i> 角色列表 <span style="font-size:12px;color:#888;font-weight:normal;">点击"操作权限"可细化每个模块的CRUD权限</span></div>
            <div class="card-body">
                <table class="roles-table">
                    <thead>
                        <tr>
                            <th>ID</th>
                            <th>角色名称</th>
                            <th>角色代码</th>
                            <th>说明</th>
                            <th>可访问模块</th>
                            <th>操作</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% 
                        On Error Resume Next
                        If Not rsRoles Is Nothing Then
                            Do While Not rsRoles.EOF
                                Err.Clear
                                Dim moduleAccess
                                moduleAccess = rsRoles("ModuleAccess")
                                If IsNull(moduleAccess) Or VarType(moduleAccess) = 10 Then
                                    moduleAccess = rsRoles("Permissions")
                                End If
                                If IsNull(moduleAccess) Or VarType(moduleAccess) = 10 Then moduleAccess = ""
                                Dim modArr : modArr = Split(moduleAccess, ",")
                        %>
                        <tr>
                            <td><%= rsRoles("RoleID") %></td>
                            <td><strong><%= rsRoles("RoleName") %></strong></td>
                            <td><span class="role-code"><%= rsRoles("RoleCode") %></span></td>
                            <td><%= rsRoles("Description") %></td>
                            <td>
                                <div class="module-tags">
                                    <% Dim ma, mn
                                    For Each ma In modArr
                                        ma = Trim(ma)
                                        mn = ma
                                        Dim mi
                                        For mi = 0 To UBound(modules)
                                            If modules(mi) = ma Then mn = moduleNames(mi) : Exit For
                                        Next
                                        If ma <> "" Then Response.Write "<span class='module-tag'>" & mn & "</span>"
                                    Next %>
                                </div>
                            </td>
                            <td>
                                <button class="btn btn-primary" onclick="openPermModal(<%= rsRoles("RoleID") %>, '<%= rsRoles("RoleCode") %>', '<%= rsRoles("RoleName") %>', '<%= moduleAccess %>')"><i class="fas fa-lock"></i> 操作权限</button>
                            </td>
                        </tr>
                        <% 
                                rsRoles.MoveNext
                            Loop
                            rsRoles.Close
                        End If 
                        On Error GoTo 0
                        %>
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- 操作级权限编辑弹窗 -->
    <div class="modal-overlay" id="permModal">
        <div class="modal">
            <h3><i class="fas fa-lock"></i> 操作级权限设置</h3>
            <div class="subtitle">角色: <strong id="permRoleName"></strong> | 模块: <strong id="permModuleName"></strong></div>
            <form method="post" id="permForm">
                <input type="hidden" name="action" value="save_perm">
                <input type="hidden" name="roleID" id="permRoleID">
                <input type="hidden" name="moduleCode" id="permModuleCode">
                <div class="perm-section">
                    <strong>操作权限</strong>
                    <label class="perm-check"><input type="checkbox" name="CanView" value="1"> 查看</label>
                    <label class="perm-check"><input type="checkbox" name="CanCreate" value="1"> 新增</label>
                    <label class="perm-check"><input type="checkbox" name="CanEdit" value="1"> 编辑</label>
                    <label class="perm-check"><input type="checkbox" name="CanDelete" value="1"> 删除</label>
                    <label class="perm-check"><input type="checkbox" name="CanExport" value="1"> 导出</label>
                    <label class="perm-check"><input type="checkbox" name="CanApprove" value="1"> 审批</label>
                </div>
                <div class="modal-actions">
                    <button type="button" class="btn btn-ghost" onclick="closePermModal()">取消</button>
                    <button type="submit" class="btn btn-success"><i class="fas fa-save"></i> 保存权限</button>
                </div>
            </form>
        </div>
    </div>

    <script>
    var allModules = ['operation','production','finance','content','system','purchase','semifinished','prodcenter','logistics'];
    var allModuleNames = ['运营管理','生产管理','财务管理','内容管理','系统管理','采购管理','半成品生产','产品生产','物流管理'];
    var currentRoleID, currentRoleCode;
    
    function openPermModal(roleID, roleCode, roleName, moduleAccess) {
        currentRoleID = roleID;
        currentRoleCode = roleCode;
        document.getElementById('permRoleID').value = roleID;
        document.getElementById('permRoleName').innerText = roleName;
        
        var modal = document.getElementById('permModal');
        var html = '<div class="page-header" style="margin-bottom:15px;"><h3 style="margin:0;font-size:18px;"><i class="fas fa-lock"></i> 操作级权限设置</h3>';
        html += '<div style="font-size:13px;color:#888;margin-top:5px;">角色: <strong>' + roleName + '</strong></div></div>';
        html += '<div style="max-height:50vh;overflow-y:auto;">';
        
        for (var i = 0; i < allModules.length; i++) {
            var mod = allModules[i];
            var modName = allModuleNames[i];
            html += '<div class="perm-section">';
            html += '<strong>' + modName + ' (' + mod + ')</strong>';
            html += '<form method="post" style="margin-top:8px;">';
            html += '<input type="hidden" name="action" value="save_perm">';
            html += '<input type="hidden" name="roleID" value="' + roleID + '">';
            html += '<input type="hidden" name="moduleCode" value="' + mod + '">';
            html += '<label class="perm-check"><input type="checkbox" name="CanView" value="1" checked> 查看</label>';
            html += '<label class="perm-check"><input type="checkbox" name="CanCreate" value="1"> 新增</label>';
            html += '<label class="perm-check"><input type="checkbox" name="CanEdit" value="1"> 编辑</label>';
            html += '<label class="perm-check"><input type="checkbox" name="CanDelete" value="1"> 删除</label>';
            html += '<label class="perm-check"><input type="checkbox" name="CanExport" value="1"> 导出</label>';
            html += '<label class="perm-check"><input type="checkbox" name="CanApprove" value="1"> 审批</label>';
            html += '<button type="submit" class="btn btn-success" style="margin-left:10px;"><i class="fas fa-save"></i> 保存</button>';
            html += '</form></div>';
        }
        html += '</div>';
        html += '<div class="modal-actions"><button type="button" class="btn btn-ghost" onclick="closePermModal()">关闭</button></div>';
        
        modal.innerHTML = '<div class="modal" style="max-width:700px;">' + html + '</div>';
        modal.classList.add('show');
    }
    
    function closePermModal() {
        document.getElementById('permModal').classList.remove('show');
        location.reload();
    }
    
    document.getElementById('permModal').addEventListener('click', function(e) {
        if (e.target === this) closePermModal();
    });
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
