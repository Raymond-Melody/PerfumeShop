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

' 获取管理员列表
Dim rsAdmins
Set rsAdmins = ExecuteQuery(_
    "SELECT a.*, r.RoleName, r.RoleCode " & _
    "FROM AdminUsers a " & _
    "LEFT JOIN AdminRoles r ON a.RoleID = r.RoleID " & _
    "ORDER BY a.AdminID")

Call LogAdminAction("查看管理员列表", "system", "AdminUsers", "", "")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>管理员管理 - 系统管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; }
        .admins-table { width: 100%; background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; overflow: hidden; box-shadow: 0 2px 10px rgba(0,0,0,0.08); }
        .admins-table th { background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); color: white; padding: 15px; text-align: left; }
        .admins-table td { padding: 15px; border-bottom: 1px solid rgba(255,255,255,0.06); color: #e0e0e0; }
        .admins-table tr:hover { background: rgba(255,255,255,0.05); }
        .admin-avatar { width: 40px; height: 40px; border-radius: 50%; background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); color: white; display: flex; align-items: center; justify-content: center; font-weight: bold; }
        .role-badge { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 500; }
        .role-super { background: rgba(255, 107, 107, 0.2); color: #ff6b6b; border: 1px solid rgba(255, 107, 107, 0.3); }
        .role-manager { background: rgba(0, 188, 212, 0.15); color: #00bcd4; border: 1px solid rgba(0, 188, 212, 0.3); }
        .role-staff { background: rgba(76, 175, 80, 0.15); color: #4CAF50; border: 1px solid rgba(76, 175, 80, 0.3); }
        .status-online { color: #4CAF50; }
        .status-offline { color: #666; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-users-cog"></i> 管理员管理</h2>
            <div class="breadcrumb">
                <a href="index.asp">系统中心</a> / <span>管理员管理</span>
            </div>
        </div>
        
        <div style="margin-bottom: 20px;">
            <a href="admin_add.asp" class="admin-btn admin-btn-primary"><i class="fas fa-plus"></i> 添加管理员</a>
        </div>
        
        <table class="admins-table">
            <thead>
                <tr>
                    <th>管理员</th>
                    <th>角色</th>
                    <th>部门</th>
                    <th>状态</th>
                    <th>最后登录</th>
                    <th>操作</th>
                </tr>
            </thead>
            <tbody>
                <% If Not rsAdmins Is Nothing Then %>
                <% Do While Not rsAdmins.EOF %>
                <tr>
                    <td>
                        <div style="display: flex; align-items: center; gap: 15px;">
                            <div class="admin-avatar"><%= Left(rsAdmins("Username"), 1) %></div>
                            <div>
                                <div style="font-weight: 600;"><%= rsAdmins("FullName") %></div>
                                <div style="font-size: 12px; color: #999;"><%= rsAdmins("Username") %></div>
                            </div>
                        </div>
                    </td>
                    <td>
                        <% 
                        Dim roleClass
                        Select Case rsAdmins("RoleCode")
                            Case "SUPER_ADMIN": roleClass = "role-super"
                            Case "OP_MANAGER", "PROD_MANAGER", "FIN_MANAGER": roleClass = "role-manager"
                            Case Else: roleClass = "role-staff"
                        End Select
                        %>
                        <span class="role-badge <%= roleClass %>"><%= rsAdmins("RoleName") %></span>
                    </td>
                    <td><%= rsAdmins("Department") %></td>
                    <td>
                        <% If rsAdmins("IsLocked") = True Then %>
                        <span style="color: #f44336;"><i class="fas fa-lock"></i> 已锁定</span>
                        <% Else %>
                        <span class="status-online"><i class="fas fa-circle" style="font-size: 8px;"></i> 正常</span>
                        <% End If %>
                    </td>
                    <td>
                        <% If Not IsNull(rsAdmins("LastLogin")) Then %>
                        <%= FormatDateTime(rsAdmins("LastLogin"), 2) %>
                        <% Else %>
                        <span style="color: #999;">未登录</span>
                        <% End If %>
                    </td>
                    <td>
                        <a href="admin_edit.asp?id=<%= rsAdmins("AdminID") %>" class="btn btn--primary btn--sm"><i class="fas fa-edit"></i> 编辑</a>
                        <a href="admin_reset.asp?id=<%= rsAdmins("AdminID") %>" class="btn btn--warning btn--sm"><i class="fas fa-key"></i> 重置密码</a>
                    </td>
                </tr>
                <% rsAdmins.MoveNext %>
                <% Loop %>
                <% rsAdmins.Close %>
                <% End If %>
            </tbody>
        </table>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
