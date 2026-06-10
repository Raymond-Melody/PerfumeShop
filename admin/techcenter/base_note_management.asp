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

' V9: 自动添加 UnitPrice 字段
On Error Resume Next
conn.Execute "SELECT UnitPrice FROM BaseNotes WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE BaseNotes ADD UnitPrice DECIMAL(19,4) NULL"
On Error GoTo 0

' ========== 安全工具函数 ==========
Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then
        SafeNum = 0
    Else
        On Error Resume Next
        SafeNum = CDbl(val)
        If Err.Number <> 0 Then
            SafeNum = 0
            Err.Clear
        End If
        On Error GoTo 0
    End If
End Function

' 安全计数函数
Function SafeCount(val)
    If IsNull(val) Or val = "" Then
        SafeCount = 0
    ElseIf IsNumeric(val) Then
        SafeCount = CLng(val)
    Else
        SafeCount = 0
    End If
End Function

' ========== 处理表单提交 ==========
Dim action, msg
action = Request.Form("action")
msg = ""

If action = "add" Then
    ' 添加基香
    Dim addName, addDesc, addIngredients, addActive, addUnitPrice
    addName = Trim(Request.Form("baseNoteName"))
    addDesc = Trim(Request.Form("description"))
    addIngredients = Trim(Request.Form("ingredients"))
    addActive = Request.Form("isActive")
    If addActive = "" Then addActive = 1
    addUnitPrice = Request.Form("unitPrice")
    If addUnitPrice = "" Or Not IsNumeric(addUnitPrice) Then addUnitPrice = 0
    
    ' 处理换行符分隔的成分
    If InStr(addIngredients, vbCrLf) > 0 Or InStr(addIngredients, vbLf) > 0 Or InStr(addIngredients, vbCr) > 0 Then
        Dim ingLines, ingLine, ingResult
        addIngredients = Replace(addIngredients, vbCrLf, vbLf)
        addIngredients = Replace(addIngredients, vbCr, vbLf)
        ingLines = Split(addIngredients, vbLf)
        ingResult = ""
        For Each ingLine In ingLines
            ingLine = Trim(ingLine)
            If ingLine <> "" Then
                If ingResult <> "" Then ingResult = ingResult & ","
                ingResult = ingResult & ingLine
            End If
        Next
        addIngredients = ingResult
    End If
    
    If addName = "" Then
        msg = "基香名称不能为空"
    ElseIf addIngredients = "" Then
        msg = "成分列表不能为空"
    Else
        Dim addSql
        addSql = "INSERT INTO BaseNotes (BaseNoteName, Description, Ingredients, IsActive, UnitPrice) VALUES (" & _
                 "'" & SafeSQL(addName) & "', " & _
                 "'" & SafeSQL(addDesc) & "', " & _
                 "'" & SafeSQL(addIngredients) & "', " & _
                 addActive & ", " & SafeNum(addUnitPrice) & ")"
        
        If ExecuteNonQuery(addSql) Then
            Response.Redirect "base_note_management.asp?msg=" & Server.URLEncode("基香添加成功")
        Else
            msg = "添加失败：" & Session("LastDBError")
        End If
    End If
    
ElseIf action = "edit" Then
    ' 编辑基香
    Dim editId, editName, editDesc, editIngredients, editActive, editUnitPrice
    editId = Request.Form("baseNoteId")
    editName = Trim(Request.Form("baseNoteName"))
    editDesc = Trim(Request.Form("description"))
    editIngredients = Trim(Request.Form("ingredients"))
    editActive = Request.Form("isActive")
    If editActive = "" Then editActive = 1
    editUnitPrice = Request.Form("unitPrice")
    If editUnitPrice = "" Or Not IsNumeric(editUnitPrice) Then editUnitPrice = 0
    
    ' 处理换行符分隔的成分
    If InStr(editIngredients, vbCrLf) > 0 Or InStr(editIngredients, vbLf) > 0 Or InStr(editIngredients, vbCr) > 0 Then
        Dim editLines, editLine, editResult
        editIngredients = Replace(editIngredients, vbCrLf, vbLf)
        editIngredients = Replace(editIngredients, vbCr, vbLf)
        editLines = Split(editIngredients, vbLf)
        editResult = ""
        For Each editLine In editLines
            editLine = Trim(editLine)
            If editLine <> "" Then
                If editResult <> "" Then editResult = editResult & ","
                editResult = editResult & editLine
            End If
        Next
        editIngredients = editResult
    End If
    
    If editName = "" Then
        msg = "基香名称不能为空"
    ElseIf editIngredients = "" Then
        msg = "成分列表不能为空"
    ElseIf IsNumeric(editId) Then
        Dim editSql
        editSql = "UPDATE BaseNotes SET " & _
                  "BaseNoteName = '" & SafeSQL(editName) & "', " & _
                  "Description = '" & SafeSQL(editDesc) & "', " & _
                  "Ingredients = '" & SafeSQL(editIngredients) & "', " & _
                  "IsActive = " & editActive & ", " & _
                  "UnitPrice = " & SafeNum(editUnitPrice) & " " & _
                  "WHERE BaseNoteID = " & CLng(editId)
        
        If ExecuteNonQuery(editSql) Then
            Response.Redirect "base_note_management.asp?msg=" & Server.URLEncode("基香更新成功")
        Else
            msg = "更新失败：" & Session("LastDBError")
        End If
    End If
    
ElseIf action = "toggle_status" Then
    ' 切换状态（软删除/恢复）
    Dim toggleId, toggleActive
    toggleId = Request.Form("baseNoteId")
    toggleActive = Request.Form("isActive")
    
    If IsNumeric(toggleId) Then
        If ExecuteNonQuery("UPDATE BaseNotes SET IsActive = " & toggleActive & " WHERE BaseNoteID = " & CLng(toggleId)) Then
            Response.Redirect "base_note_management.asp?msg=" & Server.URLEncode("状态更新成功")
        Else
            msg = "状态更新失败"
        End If
    End If
End If

' ========== 获取搜索参数 ==========
Dim searchKeyword
searchKeyword = Request.QueryString("search")

' ========== 获取基香列表 ==========
Dim sql, rsBaseNotes
If searchKeyword <> "" Then
    sql = "SELECT * FROM BaseNotes WHERE BaseNoteName LIKE '%" & SafeSQL(searchKeyword) & "%' ORDER BY BaseNoteID DESC"
Else
    sql = "SELECT * FROM BaseNotes ORDER BY BaseNoteID DESC"
End If
Set rsBaseNotes = ExecuteQuery(sql)

' ========== 获取统计 ==========
Dim totalCount, activeCount, inactiveCount
totalCount = SafeCount(GetScalar("SELECT COUNT(*) FROM BaseNotes"))
activeCount = SafeCount(GetScalar("SELECT COUNT(*) FROM BaseNotes WHERE IsActive <> 0"))
inactiveCount = totalCount - activeCount
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>基香生成管理 - 产品技术管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        /* 暗色主题 */
        body {
            background: #1a1a2e;
            color: #e0e0e0;
        }
        
        /* 统计卡片 */
        .stats-bar {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 15px;
            margin-bottom: 25px;
        }
        .stat-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 10px;
            padding: 20px;
            text-align: center;
            border: 1px solid rgba(255,255,255,0.05);
            transition: all 0.3s ease;
        }
        .stat-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 15px rgba(0,188,212,0.2);
        }
        .stat-value {
            font-size: 28px;
            font-weight: 700;
            color: #fff;
        }
        .stat-label {
            font-size: 12px;
            color: #888;
            margin-top: 5px;
            text-transform: uppercase;
        }
        .stat-card.active .stat-value { color: #4caf50; }
        .stat-card.inactive .stat-value { color: #f44336; }
        
        /* 筛选栏 */
        .filter-bar {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 10px;
            padding: 15px 20px;
            margin-bottom: 20px;
            display: flex;
            gap: 15px;
            align-items: center;
            flex-wrap: wrap;
        }
        .filter-group {
            display: flex;
            align-items: center;
            gap: 8px;
            flex: 1;
        }
        .filter-label {
            color: #888;
            font-size: 13px;
        }
        .filter-input {
            background: rgba(0,0,0,0.3);
            border: 1px solid #3a3a5a;
            border-radius: 6px;
            padding: 8px 12px;
            color: #fff;
            font-size: 13px;
            width: 250px;
        }
        /* filter-btn 和 action-btn 样式已迁移至 /css/buttons.css 统一管理系统 */
        
        /* 基香卡片网格 */
        .base-note-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(350px, 1fr));
            gap: 20px;
        }
        .base-note-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 20px;
            border: 1px solid rgba(255,255,255,0.05);
            transition: all 0.3s ease;
        }
        .base-note-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.3);
            border-color: rgba(0,188,212,0.2);
        }
        .base-note-header {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            margin-bottom: 15px;
        }
        .base-note-title {
            font-size: 16px;
            font-weight: 600;
            color: #fff;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .base-note-title i {
            color: #00bcd4;
        }
        .base-note-id {
            font-size: 11px;
            color: #999;
            background: rgba(0,0,0,0.3);
            padding: 2px 8px;
            border-radius: 4px;
        }
        .base-note-desc {
            color: #888;
            font-size: 13px;
            margin-bottom: 15px;
            line-height: 1.5;
            min-height: 40px;
        }
        .base-note-desc.empty {
            color: #999;
            font-style: italic;
        }
        
        /* 成分标签 */
        .ingredients-section {
            margin-bottom: 15px;
        }
        .ingredients-label {
            font-size: 11px;
            color: #999;
            text-transform: uppercase;
            margin-bottom: 8px;
        }
        .ingredients-list {
            display: flex;
            flex-wrap: wrap;
            gap: 6px;
        }
        .ingredient-tag {
            background: rgba(0,188,212,0.15);
            color: #00bcd4;
            padding: 4px 10px;
            border-radius: 4px;
            font-size: 12px;
            border: 1px solid rgba(0,188,212,0.3);
        }
        .ingredient-tag.more {
            background: rgba(255,255,255,0.1);
            color: #888;
            border-color: rgba(255,255,255,0.1);
        }
        .ingredients-empty {
            color: #999;
            font-size: 12px;
            font-style: italic;
        }
        
        /* 卡片底部 */
        .base-note-footer {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding-top: 15px;
            border-top: 1px solid rgba(255,255,255,0.05);
        }
        .status-badge {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 4px;
            font-size: 11px;
        }
        .status-badge.active {
            background: rgba(76, 175, 80, 0.2);
            color: #4caf50;
        }
        .status-badge.inactive {
            background: rgba(244, 67, 54, 0.2);
            color: #f44336;
        }
        .action-btns {
            display: flex;
            gap: 8px;
        }
        /* action-btn 样式由 /css/buttons.css 统一管理 */
        
        /* 模态框 */
        .admin-modal {
            display: none;
            position: fixed;
            z-index: 1000;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.7);
        }
        .admin-modal-content {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            margin: 5% auto;
            border-radius: 12px;
            width: 90%;
            max-width: 500px;
            border: 1px solid rgba(255,255,255,0.1);
            box-shadow: 0 20px 60px rgba(0,0,0,0.5);
        }
        .admin-modal-header {
            padding: 20px;
            border-bottom: 1px solid #3a3a5a;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .admin-modal-title {
            font-size: 18px;
            font-weight: 600;
            color: #fff;
            margin: 0;
        }
        .admin-modal-close {
            background: none;
            border: none;
            color: #888;
            font-size: 24px;
            cursor: pointer;
        }
        .admin-modal-close:hover {
            color: #fff;
        }
        .admin-modal-body {
            padding: 20px;
        }
        .admin-modal-footer {
            padding: 15px 20px;
            border-top: 1px solid #3a3a5a;
            display: flex;
            justify-content: flex-end;
            gap: 10px;
        }
        
        /* 表单样式 */
        .form-group {
            margin-bottom: 15px;
        }
        .form-label {
            display: block;
            margin-bottom: 8px;
            color: #b0b0b0;
            font-size: 13px;
        }
        .form-label .required {
            color: #f44336;
        }
        .form-control {
            width: 100%;
            background: rgba(0,0,0,0.3);
            border: 1px solid #3a3a5a;
            border-radius: 6px;
            padding: 10px 12px;
            color: #fff;
            font-size: 14px;
            box-sizing: border-box;
        }
        .form-control:focus {
            outline: none;
            border-color: #00bcd4;
        }
        textarea.form-control {
            resize: vertical;
            min-height: 80px;
        }
        .form-hint {
            font-size: 12px;
            color: #666;
            margin-top: 6px;
        }
        
        /* 空状态 */
        .empty-state {
            text-align: center;
            padding: 60px 20px;
            color: #666;
        }
        .empty-state i {
            font-size: 64px;
            margin-bottom: 20px;
            color: #444;
        }
        .empty-state h3 {
            font-size: 18px;
            margin-bottom: 10px;
            color: #888;
        }
        
        /* 响应式 */
        @media (max-width: 1200px) {
            .stats-bar { grid-template-columns: repeat(3, 1fr); }
        }
        @media (max-width: 768px) {
            .stats-bar { grid-template-columns: 1fr; }
            .filter-bar { flex-direction: column; align-items: stretch; }
            .base-note-grid { grid-template-columns: 1fr; }
        }

        /* ====== 弹窗颜色覆盖（防 admin.css 污染）====== */
        .admin-modal-content .admin-modal-title,
        .admin-modal-content h1,
        .admin-modal-content h2,
        .admin-modal-content h3,
        .admin-modal-content h4,
        .admin-modal-content h5,
        .admin-modal-content h6 { color: #ffffff !important; }
        .admin-modal-content .admin-modal-close { color: #bbb !important; }
        .admin-modal-content .admin-modal-close:hover { color: #fff !important; }
        .admin-modal-content .form-label,
        .admin-modal-content .admin-form-label { color: #e0e0e0 !important; }
        .admin-modal-content .form-control,
        .admin-modal-content .admin-form-control { color: #fff !important; background: rgba(0,0,0,0.3) !important; border-color: #3a3a5a !important; }
        .admin-modal-content .form-control::placeholder,
        .admin-modal-content .admin-form-control::placeholder { color: #999 !important; }
        .admin-modal-content small,
        .admin-modal-content .form-text,
        .admin-modal-content .text-muted { color: #aaa !important; }
        .admin-modal-footer { background: rgba(0,0,0,0.2) !important; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <!-- 页面标题 -->
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-wine-bottle"></i> 基香生成管理</h2>
            <div class="breadcrumb">
                <a href="index.asp">技术中心</a> / <span>基香管理</span>
            </div>
        </div>
        
        <% If Request.QueryString("msg") <> "" Then %>
        <div class="alert alert-success" style="background: rgba(76,175,80,0.1); border: 1px solid rgba(76,175,80,0.3); color: #4caf50; padding: 12px 15px; border-radius: 6px; margin-bottom: 20px;">
            <i class="fas fa-check-circle"></i> <%= Server.HTMLEncode(Request.QueryString("msg")) %>
        </div>
        <% End If %>
        
        <% If msg <> "" Then %>
        <div class="alert alert-error" style="background: rgba(244,67,54,0.1); border: 1px solid rgba(244,67,54,0.3); color: #f44336; padding: 12px 15px; border-radius: 6px; margin-bottom: 20px;">
            <i class="fas fa-exclamation-circle"></i> <%= Server.HTMLEncode(msg) %>
        </div>
        <% End If %>
        
        <!-- 统计区域 -->
        <div style="display: grid; grid-template-columns: 1fr 200px; gap: 20px; margin-bottom: 25px;">
            <!-- 统计卡片 -->
            <div class="stats-bar" style="margin-bottom: 0;">
                <div class="stat-card">
                    <div class="stat-value"><%= totalCount %></div>
                    <div class="stat-label">基香总数</div>
                </div>
                <div class="stat-card active">
                    <div class="stat-value"><%= activeCount %></div>
                    <div class="stat-label">已启用</div>
                </div>
                <div class="stat-card inactive">
                    <div class="stat-value"><%= inactiveCount %></div>
                    <div class="stat-label">已禁用</div>
                </div>
            </div>
            
            <!-- 统计图表 -->
            <div style="background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 10px; padding: 15px; border: 1px solid rgba(255,255,255,0.05); display: flex; flex-direction: column; align-items: center; justify-content: center;">
                <div style="position: relative; width: 120px; height: 120px;">
                    <canvas id="statusChart" width="120" height="120"></canvas>
                    <div style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); text-align: center;">
                        <div style="font-size: 20px; font-weight: 700; color: #fff;"><%= totalCount %></div>
                        <div style="font-size: 10px; color: #888;">总数</div>
                    </div>
                </div>
                <div style="display: flex; gap: 15px; margin-top: 10px; font-size: 11px;">
                    <span style="color: #888;"><span style="display: inline-block; width: 8px; height: 8px; background: #4CAF50; border-radius: 50%; margin-right: 4px;"></span>启用 <%= activeCount %></span>
                    <span style="color: #888;"><span style="display: inline-block; width: 8px; height: 8px; background: #9E9E9E; border-radius: 50%; margin-right: 4px;"></span>禁用 <%= inactiveCount %></span>
                </div>
            </div>
        </div>
        
        <!-- 筛选栏 -->
        <div class="filter-bar">
            <div class="filter-group">
                <span class="filter-label"><i class="fas fa-search"></i> 搜索:</span>
                <input type="text" class="filter-input" id="searchInput" placeholder="输入基香名称..." value="<%= Server.HTMLEncode(searchKeyword) %>">
            </div>
            <button class="filter-btn" onclick="doSearch()">
                <i class="fas fa-search"></i> 搜索
            </button>
            <% If searchKeyword <> "" Then %>
            <button class="filter-btn" onclick="location.href='base_note_management.asp'">
                <i class="fas fa-times"></i> 清除
            </button>
            <% End If %>
            <% If isManager Then %>
            <button class="filter-btn secondary" onclick="showAddModal()">
                <i class="fas fa-plus"></i> 新建基香
            </button>
            <% End If %>
        </div>
        
        <!-- 基香卡片列表 -->
        <div class="base-note-grid">
            <% 
            If Not rsBaseNotes Is Nothing Then 
                Do While Not rsBaseNotes.EOF
                    Dim baseNoteIngredients, ingArray, ingDisplay, ingCount
                    baseNoteIngredients = Trim(rsBaseNotes("Ingredients") & "")
                    ingCount = 0
                    
                    If baseNoteIngredients <> "" Then
                        ingArray = Split(baseNoteIngredients, ",")
                        ingCount = UBound(ingArray) + 1
                    End If
            %>
            <div class="base-note-card">
                <div class="base-note-header">
                    <div class="base-note-title">
                        <i class="fas fa-flask"></i>
                        <%= HTMLEncode(rsBaseNotes("BaseNoteName")) %>
                    </div>
                    <span class="base-note-id">#<%= rsBaseNotes("BaseNoteID") %></span>
                </div>
                
                <% 
                Dim bnUnitPrice
                bnUnitPrice = 0
                On Error Resume Next
                bnUnitPrice = CDbl(rsBaseNotes("UnitPrice"))
                If Err.Number <> 0 Then bnUnitPrice = 0 : Err.Clear
                On Error GoTo 0
                %>
                <% If bnUnitPrice > 0 Then %>
                <div style="margin-bottom: 10px; color: #ff9800; font-size: 13px; font-weight: 600;">
                    <i class="fas fa-tag"></i> ¥<%= FormatNumber(bnUnitPrice, 4) %>/ml
                </div>
                <% End If %>
                <% 
                Dim baseDesc
                baseDesc = Trim(rsBaseNotes("Description") & "")
                If baseDesc <> "" Then
                %>
                <div class="base-note-desc"><%= HTMLEncode(Left(baseDesc, 100)) %><% If Len(baseDesc) > 100 Then Response.Write "..." %></div>
                <% Else %>
                <div class="base-note-desc empty">暂无描述</div>
                <% End If %>
                
                <div class="ingredients-section">
                    <div class="ingredients-label"><i class="fas fa-leaf"></i> 成分列表 (<%= ingCount %>)</div>
                    <% If ingCount > 0 Then %>
                    <div class="ingredients-list">
                        <% 
                        Dim displayCount
                        displayCount = 0
                        For Each ingDisplay In ingArray
                            If Trim(ingDisplay) <> "" Then
                                displayCount = displayCount + 1
                                If displayCount <= 5 Then
                        %>
                        <span class="ingredient-tag"><%= HTMLEncode(Trim(ingDisplay)) %></span>
                        <% 
                                End If
                            End If
                        Next
                        If ingCount > 5 Then
                        %>
                        <span class="ingredient-tag more">+<%= ingCount - 5 %> 更多</span>
                        <% End If %>
                    </div>
                    <% Else %>
                    <div class="ingredients-empty">未设置成分</div>
                    <% End If %>
                </div>
                
                <div class="base-note-footer">
                    <span class="status-badge <%= IIf(rsBaseNotes("IsActive"), "active", "inactive") %>">
                        <%= IIf(rsBaseNotes("IsActive"), "启用", "禁用") %>
                    </span>
                    <div class="action-btns">
                        <button class="action-btn edit" onclick="showEditModal(<%= rsBaseNotes("BaseNoteID") %>, '<%= SafeOutput(rsBaseNotes("BaseNoteName")) %>', '<%= SafeOutput(rsBaseNotes("Description")) %>', '<%= SafeOutput(rsBaseNotes("Ingredients")) %>', <%= IIf(rsBaseNotes("IsActive"), 1, 0) %>, <%= bnUnitPrice %>)">
                            <i class="fas fa-edit"></i> 编辑
                        </button>
                        <% If isManager Then %>
                        <form method="post" style="display:inline;" onsubmit="return confirm('<%= IIf(rsBaseNotes("IsActive"), "确定要禁用此基香吗？", "确定要启用此基香吗？") %>')">
                            <input type="hidden" name="action" value="toggle_status">
                            <input type="hidden" name="baseNoteId" value="<%= rsBaseNotes("BaseNoteID") %>">
                            <input type="hidden" name="isActive" value="<%= IIf(rsBaseNotes("IsActive"), 0, 1) %>">
                            <button type="submit" class="action-btn <%= IIf(rsBaseNotes("IsActive"), "reject", "edit") %>">
                                <i class="fas fa-<%= IIf(rsBaseNotes("IsActive"), "ban", "check") %>"></i> <%= IIf(rsBaseNotes("IsActive"), "禁用", "启用") %>
                            </button>
                        </form>
                        <% End If %>
                    </div>
                </div>
            </div>
            <% 
                    rsBaseNotes.MoveNext
                Loop
                rsBaseNotes.Close
                Set rsBaseNotes = Nothing
            Else
            %>
            <div class="empty-state" style="grid-column: 1 / -1;">
                <i class="fas fa-flask"></i>
                <h3>暂无基香数据</h3>
                <p>点击"新建基香"按钮添加第一个基香</p>
            </div>
            <% End If %>
        </div>
    </div>
    
    <!-- 添加/编辑基香模态框 -->
    <div id="baseNoteModal" class="admin-modal">
        <div class="admin-modal-content">
            <div class="admin-modal-header">
                <h3 class="admin-modal-title" id="modalTitle"><i class="fas fa-plus"></i> 新建基香</h3>
                <button class="admin-modal-close" onclick="closeModal()">&times;</button>
            </div>
            <form method="post" id="baseNoteForm">
                <div class="admin-modal-body">
                    <input type="hidden" name="action" id="formAction" value="add">
                    <input type="hidden" name="baseNoteId" id="baseNoteId" value="">
                    
                    <div class="form-group">
                        <label class="form-label">基香名称 <span class="required">*</span></label>
                        <input type="text" name="baseNoteName" id="baseNoteName" class="form-control" required placeholder="输入基香名称">
                    </div>
                    
                    <div class="form-group">
                        <label class="form-label">描述</label>
                        <textarea name="description" id="description" class="form-control" rows="3" placeholder="输入基香描述信息"></textarea>
                    </div>
                    
                    <div class="form-group">
                        <label class="form-label">成分列表 <span class="required">*</span></label>
                        <textarea name="ingredients" id="ingredients" class="form-control" rows="5" required placeholder="例如：&#10;乙醇&#10;蒸馏水&#10;香精油"></textarea>
                        <div class="form-hint">
                            <i class="fas fa-info-circle"></i> 支持逗号分隔或每行一个成分，用于过敏源追溯
                        </div>
                    </div>
                    
                    <div class="form-group">
                        <label class="form-label">单价 (¥/ml)</label>
                        <input type="number" name="unitPrice" id="unitPrice" class="form-control" step="0.0001" min="0" placeholder="输入基香单价，如 0.0250">
                        <div class="form-hint">
                            <i class="fas fa-info-circle"></i> 基香采购单价，用于成本传导
                        </div>
                    </div>
                    
                    <div class="form-group">
                        <label class="form-label">状态</label>
                        <select name="isActive" id="isActive" class="form-control">
                            <option value="1">启用</option>
                            <option value="0">禁用</option>
                        </select>
                    </div>
                </div>
                <div class="admin-modal-footer">
                    <button type="button" class="action-btn" onclick="closeModal()">取消</button>
                    <button type="submit" class="action-btn edit">保存</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
        // 搜索功能
        function doSearch() {
            var keyword = document.getElementById('searchInput').value;
            location.href = 'base_note_management.asp?search=' + encodeURIComponent(keyword);
        }
        
        // 回车搜索
        document.getElementById('searchInput') && document.getElementById('searchInput').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                doSearch();
            }
        });
        
        // 显示添加模态框
        function showAddModal() {
            document.getElementById('modalTitle').innerHTML = '<i class="fas fa-plus"></i> 新建基香';
            document.getElementById('formAction').value = 'add';
            document.getElementById('baseNoteId').value = '';
            document.getElementById('baseNoteName').value = '';
            document.getElementById('description').value = '';
            document.getElementById('ingredients').value = '';
            document.getElementById('unitPrice').value = '';
            document.getElementById('isActive').value = '1';
            document.getElementById('baseNoteModal').style.display = 'block';
        }
        
        // 显示编辑模态框
        function showEditModal(id, name, desc, ingredients, isActive, unitPrice) {
            document.getElementById('modalTitle').innerHTML = '<i class="fas fa-edit"></i> 编辑基香';
            document.getElementById('formAction').value = 'edit';
            document.getElementById('baseNoteId').value = id;
            document.getElementById('baseNoteName').value = name;
            document.getElementById('description').value = desc;
            document.getElementById('unitPrice').value = unitPrice || '';
            
            // 将逗号分隔的成分转换为换行显示
            if (ingredients && ingredients.indexOf(',') !== -1) {
                ingredients = ingredients.split(',').join('\n');
            }
            document.getElementById('ingredients').value = ingredients;
            document.getElementById('isActive').value = isActive;
            document.getElementById('baseNoteModal').style.display = 'block';
        }
        
        // 关闭模态框
        function closeModal() {
            document.getElementById('baseNoteModal').style.display = 'none';
        }
        
        // 点击模态框外部关闭
        window.onclick = function(event) {
            var modal = document.getElementById('baseNoteModal');
            if (event.target == modal) {
                modal.style.display = 'none';
            }
        }
        
        // 初始化 Chart.js 统计图表
        document.addEventListener('DOMContentLoaded', function() {
            var ctx = document.getElementById('statusChart');
            if (ctx) {
                var totalCount = <%= totalCount %>;
                var activeCount = <%= activeCount %>;
                var inactiveCount = <%= inactiveCount %>;
                
                // 如果没有数据，显示空状态
                if (totalCount === 0) {
                    activeCount = 0;
                    inactiveCount = 1; // 占位，显示灰色圆环
                }
                
                new Chart(ctx, {
                    type: 'doughnut',
                    data: {
                        labels: ['启用', '禁用'],
                        datasets: [{
                            data: [activeCount, inactiveCount],
                            backgroundColor: ['#4CAF50', '#9E9E9E'],
                            borderWidth: 0,
                            hoverOffset: 4
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: true,
                        cutout: '70%',
                        plugins: {
                            legend: {
                                display: false
                            },
                            tooltip: {
                                callbacks: {
                                    label: function(context) {
                                        var label = context.label || '';
                                        var value = context.parsed || 0;
                                        var total = context.dataset.data.reduce(function(a, b) { return a + b; }, 0);
                                        var percentage = total > 0 ? Math.round((value / total) * 100) + '%' : '0%';
                                        return label + ': ' + value + ' (' + percentage + ')';
                                    }
                                }
                            }
                        }
                    }
                });
            }
        });
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
