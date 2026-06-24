<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<!--#include file="../../includes/audit_utils.asp"-->
<%
Call OpenConnection()
Call EnsureAuditLogTable()

' 分页参数
Dim page, pageSize
On Error Resume Next
page = CInt(Request.QueryString("page"))
If Err.Number <> 0 Or page < 1 Then page = 1
Err.Clear
On Error GoTo 0
pageSize = 25

' 筛选条件
Dim filterAction, filterDateFrom, filterDateTo
filterAction = Request.QueryString("action")
filterDateFrom = Request.QueryString("dateFrom")
filterDateTo = Request.QueryString("dateTo")

If filterDateFrom = "" Then filterDateFrom = DateAdd("d", -7, Date())
If filterDateTo = "" Then filterDateTo = Date()

' 获取日志数据
Dim rsLogs, totalCount
totalCount = GetAuditLogCount(filterAction, filterDateFrom, filterDateTo)
Set rsLogs = GetAuditLogs(page, pageSize, filterAction, filterDateFrom, filterDateTo)

' 分页计算
Dim totalPages
totalPages = Int((totalCount + pageSize - 1) / pageSize)
If totalPages < 1 Then totalPages = 1
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>V16 审计日志 - <%= SITE_NAME %></title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <link rel="stylesheet" href="/css/admin.css">
    <style>
        .audit-container { padding: 24px; max-width: 1400px; }
        .audit-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; flex-wrap: wrap; gap: 12px; }
        .audit-header h1 { font-size: 1.5rem; color: #2c1810; }
        .filter-bar { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; margin-bottom: 20px; }
        .filter-bar select, .filter-bar input { padding: 8px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.9rem; }
        .filter-bar .btn { padding: 8px 16px; }
        .audit-stats { display: flex; gap: 16px; margin-bottom: 20px; flex-wrap: wrap; }
        .audit-stat { background: #fff; border-radius: 8px; padding: 16px 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); min-width: 120px; }
        .audit-stat .stat-value { font-size: 1.5rem; font-weight: 700; color: #8B4513; }
        .audit-stat .stat-label { font-size: 0.8rem; color: #888; margin-top: 4px; }
        .audit-table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
        .audit-table th { background: #f8f6f3; padding: 12px 14px; text-align: left; font-size: 0.85rem; font-weight: 600; color: #555; border-bottom: 2px solid #e8e0d8; }
        .audit-table td { padding: 10px 14px; border-bottom: 1px solid #f0f0f0; font-size: 0.85rem; }
        .audit-table tr:hover { background: #fdfcf9; }
        .badge-action { padding: 3px 10px; border-radius: 12px; font-size: 0.75rem; font-weight: 600; display: inline-block; }
        .badge-create { background: #e8f5e9; color: #2d8a4e; }
        .badge-update { background: #e3f2fd; color: #1565c0; }
        .badge-delete { background: #ffebee; color: #c62828; }
        .badge-export { background: #fff3e0; color: #e65100; }
        .badge-batch { background: #f3e5f5; color: #7b1fa2; }
        .badge-login { background: #e0f7fa; color: #00695c; }
        .badge-view { background: #f5f5f5; color: #616161; }
        .pagination { display: flex; gap: 4px; justify-content: center; margin-top: 20px; }
        .pagination a, .pagination span { padding: 8px 14px; border: 1px solid #ddd; border-radius: 6px; text-decoration: none; color: #333; font-size: 0.9rem; }
        .pagination a:hover { background: #8B4513; color: #fff; border-color: #8B4513; }
        .pagination .current { background: #8B4513; color: #fff; border-color: #8B4513; }
        .pagination .disabled { color: #ccc; }
        .no-data { text-align: center; padding: 40px; color: #888; }
    </style>
</head>
<body>
<!--#include file="includes/nav.asp"-->

<main class="admin-content">
<div class="audit-container">
    <div class="audit-header">
        <h1><i class="fas fa-shield-alt"></i> V16 管理员操作审计日志</h1>
        <span style="color:#888; font-size:0.85rem;">自动记录所有管理后台操作</span>
    </div>

    <!-- 统计卡片 -->
    <div class="audit-stats">
        <div class="audit-stat">
            <div class="stat-value"><%= totalCount %></div>
            <div class="stat-label">总操作记录</div>
        </div>
        <div class="audit-stat">
            <div class="stat-value"><%= GetAuditLogCount(AUDIT_ACTION_BATCH, filterDateFrom, filterDateTo) %></div>
            <div class="stat-label">批量操作</div>
        </div>
        <div class="audit-stat">
            <div class="stat-value"><%= GetAuditLogCount(AUDIT_ACTION_EXPORT, filterDateFrom, filterDateTo) %></div>
            <div class="stat-label">数据导出</div>
        </div>
        <div class="audit-stat">
            <div class="stat-value"><%= GetAuditLogCount(AUDIT_ACTION_DELETE, filterDateFrom, filterDateTo) %></div>
            <div class="stat-label">删除操作</div>
        </div>
    </div>

    <!-- 筛选栏 -->
    <form method="get" class="filter-bar">
        <select name="action">
            <option value="">全部操作类型</option>
            <option value="<%= AUDIT_ACTION_CREATE %>"<% If filterAction = AUDIT_ACTION_CREATE Then %> selected<% End If %>>创建</option>
            <option value="<%= AUDIT_ACTION_UPDATE %>"<% If filterAction = AUDIT_ACTION_UPDATE Then %> selected<% End If %>>更新</option>
            <option value="<%= AUDIT_ACTION_DELETE %>"<% If filterAction = AUDIT_ACTION_DELETE Then %> selected<% End If %>>删除</option>
            <option value="<%= AUDIT_ACTION_EXPORT %>"<% If filterAction = AUDIT_ACTION_EXPORT Then %> selected<% End If %>>导出</option>
            <option value="<%= AUDIT_ACTION_BATCH %>"<% If filterAction = AUDIT_ACTION_BATCH Then %> selected<% End If %>>批量操作</option>
            <option value="<%= AUDIT_ACTION_LOGIN %>"<% If filterAction = AUDIT_ACTION_LOGIN Then %> selected<% End If %>>登录</option>
        </select>
        <input type="date" name="dateFrom" value="<%= filterDateFrom %>" title="开始日期">
        <input type="date" name="dateTo" value="<%= filterDateTo %>" title="结束日期">
        <button type="submit" class="btn btn-primary"><i class="fas fa-filter"></i> 筛选</button>
        <a href="audit_logs.asp" class="btn btn-outline"><i class="fas fa-redo"></i> 重置</a>
    </form>

    <!-- 日志表格 -->
    <% If Not rsLogs Is Nothing Then %>
    <div class="table-responsive">
    <table class="audit-table">
        <thead>
            <tr>
                <th>ID</th>
                <th>时间</th>
                <th>管理员</th>
                <th>操作类型</th>
                <th>目标类型</th>
                <th>目标</th>
                <th>详情</th>
                <th>IP地址</th>
            </tr>
        </thead>
        <tbody>
        <%
        Dim logCount
        logCount = 0
        Do While Not rsLogs.EOF
            logCount = logCount + 1
            Dim actionBadgeClass
            Select Case rsLogs("ActionType") & ""
                Case AUDIT_ACTION_CREATE: actionBadgeClass = "badge-create"
                Case AUDIT_ACTION_UPDATE: actionBadgeClass = "badge-update"
                Case AUDIT_ACTION_DELETE: actionBadgeClass = "badge-delete"
                Case AUDIT_ACTION_EXPORT: actionBadgeClass = "badge-export"
                Case AUDIT_ACTION_BATCH: actionBadgeClass = "badge-batch"
                Case AUDIT_ACTION_LOGIN: actionBadgeClass = "badge-login"
                Case Else: actionBadgeClass = "badge-view"
            End Select
        %>
            <tr>
                <td><%= rsLogs("LogID") %></td>
                <td><%= rsLogs("CreatedAt") %></td>
                <td><%= HTMLEncode(rsLogs("AdminName") & "") %></td>
                <td><span class="badge-action <%= actionBadgeClass %>"><%= rsLogs("ActionType") %></span></td>
                <td><%= rsLogs("TargetType") %></td>
                <td><%= HTMLEncode(Left(rsLogs("TargetName") & "", 50)) %></td>
                <td><%= HTMLEncode(Left(rsLogs("Details") & "", 100)) %></td>
                <td><%= rsLogs("IPAddress") %></td>
            </tr>
        <%
            rsLogs.MoveNext
        Loop
        %>
        </tbody>
    </table>
    </div>
    <%
    rsLogs.Close
    Set rsLogs = Nothing
    %>
    <% If logCount = 0 Then %>
    <div class="no-data"><i class="fas fa-inbox"></i> 暂无操作记录</div>
    <% End If %>

    <!-- 分页 -->
    <% If totalPages > 1 Then %>
    <div class="pagination">
        <% If page > 1 Then %>
            <a href="?page=<%= page - 1 %>&action=<%= Server.URLEncode(filterAction) %>&dateFrom=<%= Server.URLEncode(filterDateFrom) %>&dateTo=<%= Server.URLEncode(filterDateTo) %>">上一页</a>
        <% Else %>
            <span class="disabled">上一页</span>
        <% End If %>

        <%
        Dim p, startP, endP
        startP = page - 2
        If startP < 1 Then startP = 1
        endP = page + 2
        If endP > totalPages Then endP = totalPages
        For p = startP To endP
            If p = page Then
        %>
            <span class="current"><%= p %></span>
        <%
            Else
        %>
            <a href="?page=<%= p %>&action=<%= Server.URLEncode(filterAction) %>&dateFrom=<%= Server.URLEncode(filterDateFrom) %>&dateTo=<%= Server.URLEncode(filterDateTo) %>"><%= p %></a>
        <%
            End If
        Next
        %>

        <% If page < totalPages Then %>
            <a href="?page=<%= page + 1 %>&action=<%= Server.URLEncode(filterAction) %>&dateFrom=<%= Server.URLEncode(filterDateFrom) %>&dateTo=<%= Server.URLEncode(filterDateTo) %>">下一页</a>
        <% Else %>
            <span class="disabled">下一页</span>
        <% End If %>
    </div>
    <% End If %>
    <% End If %>
</div>
</main>

<%
Call CloseConnection()
%>
</body>
</html>
