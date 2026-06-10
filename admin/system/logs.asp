<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
Call OpenConnection()

' 分页参数
Dim page, pageSize
On Error Resume Next
page = CInt(Request.QueryString("page"))
If Err.Number <> 0 Or page < 1 Then page = 1
Err.Clear
On Error GoTo 0
pageSize = 20

' 筛选条件
Dim filterModule, filterAdmin, filterSensitive, filterAction
filterModule = Request.QueryString("module")
filterAdmin = Request.QueryString("admin")
filterSensitive = Request.QueryString("sensitive")
filterAction = Request.QueryString("action_type")

' 敏感操作关键词
Dim sensitiveActions : sensitiveActions = Array("删除", "delete", "DROP", "TRUNCATE", "重置", "封禁", "解锁", "修改权限", "更改角色", "修改密码", "备份", "恢复")

' 构建查询
Dim sqlWhere, sqlOrder
sqlWhere = "WHERE 1=1"
If filterModule <> "" Then sqlWhere = sqlWhere & " AND ModuleCode = '" & Replace(filterModule,"'","''") & "'"
If filterAdmin <> "" Then sqlWhere = sqlWhere & " AND AdminID = " & filterAdmin
If filterAction <> "" Then sqlWhere = sqlWhere & " AND ActionType = '" & Replace(filterAction,"'","''") & "'"

If filterSensitive = "1" Then
    ' 筛选敏感操作
    Dim sensWhere : sensWhere = ""
    Dim sa
    For Each sa In sensitiveActions
        If sensWhere <> "" Then sensWhere = sensWhere & " OR "
        sensWhere = sensWhere & "ActionType LIKE '%" & sa & "%' OR Notes LIKE '%" & sa & "%'"
    Next
    sqlWhere = sqlWhere & " AND (" & sensWhere & ")"
End If

sqlOrder = "ORDER BY CreatedAt DESC"

' 获取总数
Dim totalCount
totalCount = CLng("0" & GetScalar("SELECT COUNT(*) FROM AdminLogs " & sqlWhere))

' 获取日志列表
Dim rsLogs
If totalCount > 0 Then
    Set rsLogs = ExecuteQuery(_
        "SELECT TOP " & pageSize & " * FROM AdminLogs " & sqlWhere & " AND LogID NOT IN (" & _
        "SELECT TOP " & (page-1)*pageSize & " LogID FROM AdminLogs " & sqlWhere & " " & sqlOrder & _
        ") " & sqlOrder)
Else
    Set rsLogs = Nothing
End If

' 获取模块列表和操作类型
Dim rsModules
Set rsModules = ExecuteQuery("SELECT DISTINCT ModuleCode FROM AdminLogs WHERE ModuleCode IS NOT NULL ORDER BY ModuleCode")
Dim rsActionTypes
Set rsActionTypes = ExecuteQuery("SELECT DISTINCT ActionType FROM AdminLogs WHERE ActionType IS NOT NULL ORDER BY ActionType")

On Error Resume Next
Call LogAdminAction("查看操作日志", "system", "AdminLogs", "", "")
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>操作日志 - 系统管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; }
        .filter-bar { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 20px; border-radius: 12px; margin-bottom: 20px; box-shadow: 0 2px 10px rgba(0,0,0,0.08); }
        .filter-bar form { display: flex; gap: 15px; align-items: center; flex-wrap: wrap; }
        .filter-bar select, .filter-bar input { padding: 10px 15px; border: 1px solid rgba(255,255,255,0.15); border-radius: 8px; background: #2d2d44; color: #e0e0e0; }
        .filter-bar .btn-sensi { background: rgba(244,67,54,0.15); color: #EF9A9A; border: 1px solid rgba(244,67,54,0.3); padding: 8px 16px; border-radius: 6px; cursor: pointer; font-size: 13px; transition: all 0.2s; text-decoration: none; display: inline-flex; align-items: center; gap: 5px; }
        .filter-bar .btn-sensi.active { background: rgba(244,67,54,0.3); color: #ff5252; }
        .logs-table { width: 100%; background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; overflow: hidden; box-shadow: 0 2px 10px rgba(0,0,0,0.08); }
        .logs-table th { background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); color: white; padding: 15px; text-align: left; }
        .logs-table td { padding: 12px 15px; border-bottom: 1px solid rgba(255,255,255,0.06); font-size: 13px; color: #e0e0e0; }
        .logs-table tr:hover { background: rgba(255,255,255,0.05); }
        .module-tag { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 11px; }
        .module-operation { background: rgba(25, 118, 210, 0.2); color: #64b5f6; border: 1px solid rgba(25, 118, 210, 0.3); }
        .module-production { background: rgba(46, 125, 50, 0.2); color: #81c784; border: 1px solid rgba(46, 125, 50, 0.3); }
        .module-finance { background: rgba(245, 124, 0, 0.2); color: #ffb74d; border: 1px solid rgba(245, 124, 0, 0.3); }
        .module-content { background: rgba(194, 24, 91, 0.2); color: #f48fb1; border: 1px solid rgba(194, 24, 91, 0.3); }
        .module-system { background: rgba(123, 31, 162, 0.2); color: #ce93d8; border: 1px solid rgba(123, 31, 162, 0.3); }
        .action-text { font-weight: 500; color: #fff; }
        .time-text { color: #888; font-size: 12px; }
        .pagination { display: flex; justify-content: center; gap: 10px; margin-top: 20px; }
        .pagination a { padding: 8px 15px; background: #2d2d44; border-radius: 6px; text-decoration: none; color: #00bcd4; border: 1px solid rgba(255,255,255,0.06); }
        .pagination a.active { background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); color: white; border-color: transparent; }
        .pagination a:hover:not(.active) { background: rgba(255,255,255,0.1); }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-history"></i> 操作日志</h2>
            <div class="breadcrumb">
                <a href="index.asp">系统中心</a> / <span>操作日志</span>
            </div>
        </div>
        
        <div class="filter-bar">
            <form method="get" action="logs.asp">
                <select name="module">
                    <option value="">所有模块</option>
                    <% If Not rsModules Is Nothing Then
                    Do While Not rsModules.EOF %>
                    <option value="<%= rsModules("ModuleCode") %>" <%= IIf(filterModule=rsModules("ModuleCode"), "selected", "") %>><%= rsModules("ModuleCode") %></option>
                    <% rsModules.MoveNext
                    Loop
                    rsModules.Close
                    End If %>
                </select>
                <select name="action_type" style="min-width:130px;">
                    <option value="">所有操作</option>
                    <% If Not rsActionTypes Is Nothing Then
                    Do While Not rsActionTypes.EOF %>
                    <option value="<%= rsActionTypes("ActionType") %>" <%= IIf(filterAction=rsActionTypes("ActionType"), "selected", "") %>><%= rsActionTypes("ActionType") %></option>
                    <% rsActionTypes.MoveNext
                    Loop
                    rsActionTypes.Close
                    End If %>
                </select>
                <a href="logs.asp?sensitive=<%= IIf(filterSensitive="1","0","1") %>" class="btn-sensi <%= IIf(filterSensitive="1","active","") %>">
                    <i class="fas fa-exclamation-triangle"></i> 敏感操作
                </a>
                <button type="submit" class="admin-btn admin-btn-primary"><i class="fas fa-filter"></i> 筛选</button>
                <a href="logs.asp" class="admin-btn admin-btn-secondary"><i class="fas fa-undo"></i> 重置</a>
            </form>
        </div>
        
        <table class="logs-table">
            <thead>
                <tr>
                    <th>时间</th>
                    <th>模块</th>
                    <th>操作</th>
                    <th>表</th>
                    <th>记录ID</th>
                    <th>备注</th>
                </tr>
            </thead>
            <tbody>
                <% If Not rsLogs Is Nothing Then %>
                <% Do While Not rsLogs.EOF %>
                <tr>
                    <td class="time-text"><%= rsLogs("CreatedAt") %></td>
                    <td>
                        <% 
                        Dim modClass
                        Select Case rsLogs("ModuleCode")
                            Case "operation": modClass = "module-operation"
                            Case "production": modClass = "module-production"
                            Case "finance": modClass = "module-finance"
                            Case "content": modClass = "module-content"
                            Case "system": modClass = "module-system"
                            Case Else: modClass = ""
                        End Select
                        %>
                        <span class="module-tag <%= modClass %>"><%= rsLogs("ModuleCode") %></span>
                    </td>
                    <td class="action-text"><%= rsLogs("ActionType") %></td>
                    <td><%= rsLogs("TableName") %></td>
                    <td><%= rsLogs("RecordID") %></td>
                    <td><%= rsLogs("Notes") %></td>
                </tr>
                <% rsLogs.MoveNext %>
                <% Loop %>
                <% rsLogs.Close %>
                <% End If %>
            </tbody>
        </table>
        
        <% If totalCount > pageSize Then %>
        <div class="pagination">
            <% 
            Dim extraParams : extraParams = "&module=" & filterModule & "&action_type=" & filterAction & "&sensitive=" & filterSensitive
            If page > 1 Then %>
            <a href="logs.asp?page=<%= page-1 %><%= extraParams %>"><i class="fas fa-chevron-left"></i></a>
            <% End If
            
            Dim totalPages, startPage, endPage
            totalPages = Int((totalCount + pageSize - 1) / pageSize)
            startPage = IIf(page - 2 < 1, 1, page - 2)
            endPage = IIf(startPage + 4 > totalPages, totalPages, startPage + 4)
            
            Dim i
            For i = startPage To endPage %>
            <a href="logs.asp?page=<%= i %><%= extraParams %>" class="<%= IIf(i=page, "active", "") %>"><%= i %></a>
            <% Next
            If page < totalPages Then %>
            <a href="logs.asp?page=<%= page+1 %><%= extraParams %>"><i class="fas fa-chevron-right"></i></a>
            <% End If %>
        </div>
        <% End If %>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
