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
    ' 添加瓶型
    Dim addName, addDesc, addPrice, addImage, addActive
    addName = Trim(Request.Form("bottleName"))
    addDesc = Trim(Request.Form("description"))
    addPrice = Request.Form("priceAddition")
    If addPrice = "" Then addPrice = 0
    addImage = Trim(Request.Form("imageURL"))
    addActive = Request.Form("isActive")
    If addActive = "" Then addActive = 1
    
    If addName = "" Then
        msg = "瓶型名称不能为空"
    Else
        Dim addSql
        addSql = "INSERT INTO BottleStyles (BottleName, Description, PriceAddition, ImageURL, IsActive) VALUES (" & _
                 "'" & SafeSQL(addName) & "', " & _
                 "'" & SafeSQL(addDesc) & "', " & _
                 addPrice & ", " & _
                 "'" & SafeSQL(addImage) & "', " & _
                 addActive & ")"
        
        If ExecuteNonQuery(addSql) Then
            Response.Redirect "bottle_management.asp?msg=" & Server.URLEncode("瓶型添加成功")
        Else
            msg = "添加失败：" & Session("LastDBError")
        End If
    End If
    
ElseIf action = "edit" Then
    ' 编辑瓶型
    Dim editId, editName, editDesc, editPrice, editImage, editActive
    editId = Request.Form("bottleId")
    editName = Trim(Request.Form("bottleName"))
    editDesc = Trim(Request.Form("description"))
    editPrice = Request.Form("priceAddition")
    If editPrice = "" Then editPrice = 0
    editImage = Trim(Request.Form("imageURL"))
    editActive = Request.Form("isActive")
    If editActive = "" Then editActive = 1
    
    If editName = "" Then
        msg = "瓶型名称不能为空"
    ElseIf IsNumeric(editId) Then
        Dim editSql
        editSql = "UPDATE BottleStyles SET " & _
                  "BottleName = '" & SafeSQL(editName) & "', " & _
                  "Description = '" & SafeSQL(editDesc) & "', " & _
                  "PriceAddition = " & editPrice & ", " & _
                  "ImageURL = '" & SafeSQL(editImage) & "', " & _
                  "IsActive = " & editActive & " " & _
                  "WHERE BottleID = " & CLng(editId)
        
        If ExecuteNonQuery(editSql) Then
            Response.Redirect "bottle_management.asp?msg=" & Server.URLEncode("瓶型更新成功")
        Else
            msg = "更新失败：" & Session("LastDBError")
        End If
    End If
    
ElseIf action = "toggle_status" Then
    ' 切换状态（软删除/恢复）
    Dim toggleId, toggleActive
    toggleId = Request.Form("bottleId")
    toggleActive = Request.Form("isActive")
    
    If IsNumeric(toggleId) Then
        If ExecuteNonQuery("UPDATE BottleStyles SET IsActive = " & toggleActive & " WHERE BottleID = " & CLng(toggleId)) Then
            Response.Redirect "bottle_management.asp?msg=" & Server.URLEncode("状态更新成功")
        Else
            msg = "状态更新失败"
        End If
    End If
End If

' ========== 获取搜索参数 ==========
Dim searchKeyword
searchKeyword = Request.QueryString("search")

' ========== 获取瓶型列表 ==========
Dim sql, rsBottles
If searchKeyword <> "" Then
    sql = "SELECT * FROM BottleStyles WHERE BottleName LIKE '%" & SafeSQL(searchKeyword) & "%' ORDER BY BottleID DESC"
Else
    sql = "SELECT * FROM BottleStyles ORDER BY BottleID DESC"
End If
Set rsBottles = ExecuteQuery(sql)

' ========== 获取统计 ==========
Dim totalCount, activeCount, inactiveCount
totalCount = SafeCount(GetScalar("SELECT COUNT(*) FROM BottleStyles"))
activeCount = SafeCount(GetScalar("SELECT COUNT(*) FROM BottleStyles WHERE IsActive <> 0"))
inactiveCount = totalCount - activeCount
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>瓶型管理 - 产品技术管理中心</title>
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
        /* .filter-btn 样式已由 /css/buttons.css Section 5 统一管理 */
        
        /* 瓶型卡片网格 */
        .bottle-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(350px, 1fr));
            gap: 20px;
        }
        .bottle-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 20px;
            border: 1px solid rgba(255,255,255,0.05);
            transition: all 0.3s ease;
        }
        .bottle-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.3);
            border-color: rgba(0,188,212,0.2);
        }
        .bottle-header {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            margin-bottom: 15px;
        }
        .bottle-title {
            font-size: 16px;
            font-weight: 600;
            color: #fff;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .bottle-title i {
            color: #00bcd4;
        }
        .bottle-id {
            font-size: 11px;
            color: #999;
            background: rgba(0,0,0,0.3);
            padding: 2px 8px;
            border-radius: 4px;
        }
        .bottle-desc {
            color: #888;
            font-size: 13px;
            margin-bottom: 15px;
            line-height: 1.5;
            min-height: 40px;
        }
        .bottle-desc.empty {
            color: #999;
            font-style: italic;
        }
        
        /* 瓶型图片 */
        .bottle-image {
            width: 100%;
            height: 150px;
            background: rgba(0,0,0,0.3);
            border-radius: 8px;
            margin-bottom: 15px;
            display: flex;
            align-items: center;
            justify-content: center;
            overflow: hidden;
        }
        .bottle-image img {
            max-width: 100%;
            max-height: 100%;
            object-fit: contain;
        }
        .bottle-image .no-image {
            color: #888;
            font-size: 48px;
        }
        
        /* 附加费用 */
        .price-section {
            margin-bottom: 15px;
        }
        .price-label {
            font-size: 11px;
            color: #999;
            text-transform: uppercase;
            margin-bottom: 8px;
        }
        .price-value {
            font-size: 18px;
            font-weight: 600;
            color: #ff9800;
        }
        
        /* 卡片底部 */
        .bottle-footer {
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
        /* .action-btn 样式已由 /css/buttons.css Section 6 统一管理 */
        
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
            .bottle-grid { grid-template-columns: 1fr; }
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
        /* 图片上传组件 */
        .image-upload-wrapper {
            border: 2px dashed rgba(255,255,255,0.15);
            border-radius: 12px;
            padding: 16px;
            text-align: center;
            transition: border-color 0.3s;
        }
        .image-upload-wrapper:hover,
        .image-upload-wrapper.dragover {
            border-color: rgba(0,188,212,0.5);
        }
        .image-preview {
            width: 100%;
            max-width: 280px;
            height: 200px;
            margin: 0 auto 12px;
            border-radius: 8px;
            overflow: hidden;
            background: rgba(0,0,0,0.2);
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
        }
        .image-preview img {
            max-width: 100%;
            max-height: 100%;
            object-fit: contain;
        }
        .image-placeholder {
            color: #666;
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 8px;
        }
        .image-placeholder i {
            font-size: 36px;
            color: #555;
        }
        .image-placeholder span {
            font-size: 13px;
        }
        .image-upload-actions {
            display: flex;
            gap: 8px;
            justify-content: center;
            margin-top: 8px;
        }
        .image-upload-actions .btn-sm {
            padding: 6px 14px;
            font-size: 13px;
            border-radius: 6px;
            border: none;
            cursor: pointer;
        }
        .upload-progress {
            margin-top: 10px;
            background: rgba(0,0,0,0.3);
            border-radius: 6px;
            height: 24px;
            position: relative;
            overflow: hidden;
        }
        .progress-bar {
            height: 100%;
            background: linear-gradient(90deg, #00bcd4, #00e5ff);
            border-radius: 6px;
            width: 0%;
            transition: width 0.3s;
        }
        .progress-text {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            font-size: 12px;
            color: #fff;
        }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <!-- 页面标题 -->
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-wine-bottle"></i> 瓶型管理</h2>
            <div class="breadcrumb">
                <a href="index.asp">技术中心</a> / <span>瓶型管理</span>
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
                    <div class="stat-label">瓶型总数</div>
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
                <input type="text" class="filter-input" id="searchInput" placeholder="输入瓶型名称..." value="<%= Server.HTMLEncode(searchKeyword) %>">
            </div>
            <button class="filter-btn" onclick="doSearch()">
                <i class="fas fa-search"></i> 搜索
            </button>
            <% If searchKeyword <> "" Then %>
            <button class="filter-btn" onclick="location.href='bottle_management.asp'">
                <i class="fas fa-times"></i> 清除
            </button>
            <% End If %>
            <% If isManager Then %>
            <button class="filter-btn secondary" onclick="showAddModal()">
                <i class="fas fa-plus"></i> 新建瓶型
            </button>
            <% End If %>
        </div>
        
        <!-- 瓶型卡片列表 -->
        <div class="bottle-grid">
            <% 
            If Not rsBottles Is Nothing Then 
                Do While Not rsBottles.EOF
            %>
            <div class="bottle-card">
                <div class="bottle-header">
                    <div class="bottle-title">
                        <i class="fas fa-wine-bottle"></i>
                        <%= HTMLEncode(rsBottles("BottleName")) %>
                    </div>
                    <span class="bottle-id">#<%= rsBottles("BottleID") %></span>
                </div>
                
                <div class="bottle-image">
                    <% 
                    Dim imgUrl
                    imgUrl = Trim(rsBottles("ImageURL") & "")
                    If imgUrl <> "" Then
                    %>
                    <img src="<%= HTMLEncode(imgUrl) %>" alt="<%= HTMLEncode(rsBottles("BottleName")) %>">
                    <% Else %>
                    <i class="fas fa-image no-image"></i>
                    <% End If %>
                </div>
                
                <% 
                Dim bottleDesc
                bottleDesc = Trim(rsBottles("Description") & "")
                If bottleDesc <> "" Then
                %>
                <div class="bottle-desc"><%= HTMLEncode(Left(bottleDesc, 100)) %><% If Len(bottleDesc) > 100 Then Response.Write "..." %></div>
                <% Else %>
                <div class="bottle-desc empty">暂无描述</div>
                <% End If %>
                
                <div class="price-section">
                    <div class="price-label"><i class="fas fa-tag"></i> 附加费用</div>
                    <div class="price-value">¥<%= FormatNumber(SafeNum(rsBottles("PriceAddition")), 2) %></div>
                </div>
                
                <div class="bottle-footer">
                    <span class="status-badge <%= IIf(rsBottles("IsActive"), "active", "inactive") %>">
                        <%= IIf(rsBottles("IsActive"), "启用", "禁用") %>
                    </span>
                    <div class="action-btns">
                        <button class="action-btn edit" onclick="showEditModal(<%= rsBottles("BottleID") %>, '<%= SafeOutput(rsBottles("BottleName")) %>', '<%= SafeOutput(rsBottles("Description")) %>', <%= rsBottles("PriceAddition") %>, '<%= SafeOutput(rsBottles("ImageURL")) %>', <%= IIf(rsBottles("IsActive"), 1, 0) %>)">
                            <i class="fas fa-edit"></i> 编辑
                        </button>
                        <% If isManager Then %>
                        <form method="post" style="display:inline;" onsubmit="return confirm('<%= IIf(rsBottles("IsActive"), "确定要禁用此瓶型吗？", "确定要启用此瓶型吗？") %>')">
                            <input type="hidden" name="action" value="toggle_status">
                            <input type="hidden" name="bottleId" value="<%= rsBottles("BottleID") %>">
                            <input type="hidden" name="isActive" value="<%= IIf(rsBottles("IsActive"), 0, 1) %>">
                            <button type="submit" class="action-btn <%= IIf(rsBottles("IsActive"), "reject", "approve") %>">
                                <i class="fas fa-<%= IIf(rsBottles("IsActive"), "ban", "check") %>"></i> <%= IIf(rsBottles("IsActive"), "禁用", "启用") %>
                            </button>
                        </form>
                        <% End If %>
                    </div>
                </div>
            </div>
            <% 
                    rsBottles.MoveNext
                Loop
                rsBottles.Close
                Set rsBottles = Nothing
            Else
            %>
            <div class="empty-state" style="grid-column: 1 / -1;">
                <i class="fas fa-wine-bottle"></i>
                <h3>暂无瓶型数据</h3>
                <p>点击"新建瓶型"按钮添加第一个瓶型</p>
            </div>
            <% End If %>
        </div>
    </div>
    
    <!-- 添加/编辑瓶型模态框 -->
    <div id="bottleModal" class="admin-modal">
        <div class="admin-modal-content">
            <div class="admin-modal-header">
                <h3 class="admin-modal-title" id="modalTitle"><i class="fas fa-plus"></i> 新建瓶型</h3>
                <button class="admin-modal-close" onclick="closeModal()">&times;</button>
            </div>
            <form method="post" id="bottleForm">
                <div class="admin-modal-body">
                    <input type="hidden" name="action" id="formAction" value="add">
                    <input type="hidden" name="bottleId" id="bottleId" value="">
                    <input type="hidden" name="csrf_token" value="<%= Session("CSRFToken") %>">
                    
                    <div class="form-group">
                        <label class="form-label">瓶型名称 <span class="required">*</span></label>
                        <input type="text" name="bottleName" id="bottleName" class="form-control" required placeholder="输入瓶型名称">
                    </div>
                    
                    <div class="form-group">
                        <label class="form-label">描述</label>
                        <textarea name="description" id="description" class="form-control" rows="3" placeholder="输入瓶型描述信息"></textarea>
                    </div>
                    
                    <div class="form-group">
                        <label class="form-label">附加费用 (¥)</label>
                        <input type="number" name="priceAddition" id="priceAddition" class="form-control" step="0.01" min="0" value="0" placeholder="0.00">
                        <div class="form-hint">
                            <i class="fas fa-info-circle"></i> 选择此瓶型时的额外费用
                        </div>
                    </div>
                    
                    <div class="form-group">
                        <label class="form-label">图片</label>
                        <div class="image-upload-wrapper">
                            <div class="image-preview" id="imagePreview_bottle">
                                <img id="previewImg_bottle" src="" alt="预览" style="display:none;">
                                <div class="image-placeholder" id="placeholder_bottle">
                                    <i class="fas fa-cloud-upload-alt"></i>
                                    <span>点击上传或拖拽图片</span>
                                </div>
                            </div>
                            <input type="file" id="fileInput_bottle" accept="image/jpeg,image/png,image/gif,image/webp,image/svg+xml" style="display:none;">
                            <div class="image-upload-actions">
                                <button type="button" class="admin-btn admin-btn-info btn-sm" onclick="document.getElementById('fileInput_bottle').click();">
                                    <i class="fas fa-upload"></i> 选择图片
                                </button>
                                <button type="button" class="admin-btn admin-btn-secondary btn-sm" onclick="toggleUrlInput_bottle()">
                                    <i class="fas fa-link"></i> 输入URL
                                </button>
                            </div>
                            <div id="urlInputWrapper_bottle" style="display:none; margin-top:8px;">
                                <input type="text" id="manualUrl_bottle" class="admin-form-control" placeholder="输入图片URL地址" style="font-size:13px;">
                                <button type="button" class="admin-btn admin-btn-secondary btn-sm" onclick="applyManualUrl_bottle()" style="margin-top:4px;">确认</button>
                            </div>
                            <div class="upload-progress" id="uploadProgress_bottle" style="display:none;">
                                <div class="progress-bar" id="progressBar_bottle"></div>
                                <span class="progress-text" id="progressText_bottle">上传中...</span>
                            </div>
                            <div style="font-size:11px;color:#888;margin-top:6px;">如果原图超过 180KB，将自动压缩后再上传</div>
                            <input type="hidden" name="imageURL" id="imageURL_bottle" value="">
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
            location.href = 'bottle_management.asp?search=' + encodeURIComponent(keyword);
        }
        
        // 回车搜索
        document.getElementById('searchInput') && document.getElementById('searchInput').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                doSearch();
            }
        });
        
        // 显示添加模态框
        function showAddModal() {
            document.getElementById('modalTitle').innerHTML = '<i class="fas fa-plus"></i> 新建瓶型';
            document.getElementById('formAction').value = 'add';
            document.getElementById('bottleId').value = '';
            document.getElementById('bottleName').value = '';
            document.getElementById('description').value = '';
            document.getElementById('priceAddition').value = '0';
            document.getElementById('imageURL_bottle').value = '';
            document.getElementById('previewImg_bottle').style.display = 'none';
            document.getElementById('placeholder_bottle').style.display = 'flex';
            document.getElementById('fileInput_bottle').value = '';
            document.getElementById('isActive').value = '1';
            document.getElementById('bottleModal').style.display = 'block';
        }
        
        // 显示编辑模态框
        function showEditModal(id, name, desc, price, imageUrl, isActive) {
            document.getElementById('modalTitle').innerHTML = '<i class="fas fa-edit"></i> 编辑瓶型';
            document.getElementById('formAction').value = 'edit';
            document.getElementById('bottleId').value = id;
            document.getElementById('bottleName').value = name;
            document.getElementById('description').value = desc;
            document.getElementById('priceAddition').value = price;
            document.getElementById('imageURL_bottle').value = imageUrl;
            if (imageUrl) {
                document.getElementById('previewImg_bottle').src = imageUrl;
                document.getElementById('previewImg_bottle').style.display = 'block';
                document.getElementById('placeholder_bottle').style.display = 'none';
            } else {
                document.getElementById('previewImg_bottle').style.display = 'none';
                document.getElementById('placeholder_bottle').style.display = 'flex';
            }
            document.getElementById('isActive').value = isActive;
            document.getElementById('bottleModal').style.display = 'block';
        }
        
        // 关闭模态框
        function closeModal() {
            document.getElementById('bottleModal').style.display = 'none';
        }
        
        // 点击模态框外部关闭
        window.onclick = function(event) {
            var modal = document.getElementById('bottleModal');
            if (event.target == modal) {
                modal.style.display = 'none';
            }
        }

        // 图片压缩函数 - bottle
        function compressImage_bottle(file, maxSizeKB, callback) {
            if (file.type === 'image/svg+xml') {
                callback(file, false);
                return;
            }
            var maxSize = maxSizeKB * 1024;
            if (file.size <= maxSize) {
                callback(file, false);
                return;
            }
            var reader = new FileReader();
            reader.onload = function(e) {
                var img = new Image();
                img.onload = function() {
                    var canvas = document.createElement('canvas');
                    var ctx = canvas.getContext('2d');
                    var maxDim = 1200;
                    var width = img.width;
                    var height = img.height;
                    if (width > maxDim || height > maxDim) {
                        if (width > height) {
                            height = Math.round(height * maxDim / width);
                            width = maxDim;
                        } else {
                            width = Math.round(width * maxDim / height);
                            height = maxDim;
                        }
                    }
                    canvas.width = width;
                    canvas.height = height;
                    ctx.drawImage(img, 0, 0, width, height);
                    var quality = 0.8;
                    var tryCompress = function() {
                        canvas.toBlob(function(blob) {
                            if (blob.size > maxSize && quality > 0.1) {
                                quality -= 0.1;
                                tryCompress();
                            } else {
                                var compressedFile = new File([blob], file.name.replace(/\.[^.]+$/, '.jpg'), {
                                    type: 'image/jpeg',
                                    lastModified: Date.now()
                                });
                                callback(compressedFile, true);
                            }
                        }, 'image/jpeg', quality);
                    };
                    tryCompress();
                };
                img.src = e.target.result;
            };
            reader.readAsDataURL(file);
        }

        // 图片上传 - bottle
        document.getElementById('fileInput_bottle').addEventListener('change', function(e) {
            var file = e.target.files[0];
            if (!file) return;
            var maxSize = 5 * 1024 * 1024;
            if (file.size > maxSize) { alert('文件大小不能超过5MB'); return; }
            var allowedTypes = ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'image/svg+xml'];
            if (allowedTypes.indexOf(file.type) === -1) { alert('仅支持 JPG/PNG/GIF/WebP/SVG 格式'); return; }
            var reader = new FileReader();
            reader.onload = function(ev) {
                document.getElementById('previewImg_bottle').src = ev.target.result;
                document.getElementById('previewImg_bottle').style.display = 'block';
                document.getElementById('placeholder_bottle').style.display = 'none';
            };
            reader.readAsDataURL(file);
            compressImage_bottle(file, 180, function(fileToUpload, wasCompressed) {
                uploadImage_bottle(fileToUpload, 'bottle', wasCompressed);
            });
        });

        function uploadImage_bottle(file, uploadType, wasCompressed) {
            var formData = new FormData();
            formData.append('file', file);
            formData.append('type', uploadType);
            var csrfInput = document.querySelector('input[name="csrf_token"]');
            if (csrfInput) formData.append('csrf_token', csrfInput.value);
            var progressDiv = document.getElementById('uploadProgress_bottle');
            var progressBar = document.getElementById('progressBar_bottle');
            var progressText = document.getElementById('progressText_bottle');
            progressDiv.style.display = 'block';
            progressBar.style.width = '0%';
            progressText.textContent = '上传中...';
            var xhr = new XMLHttpRequest();
            xhr.upload.addEventListener('progress', function(e) {
                if (e.lengthComputable) {
                    var pct = Math.round(e.loaded / e.total * 100);
                    progressBar.style.width = pct + '%';
                    progressText.textContent = pct + '%';
                }
            });
            xhr.addEventListener('load', function() {
                try {
                    var resp = JSON.parse(xhr.responseText);
                    if (resp.success) {
                        document.getElementById('imageURL_bottle').value = resp.url;
                        progressBar.style.width = '100%';
                        progressText.textContent = wasCompressed ? '上传成功（图片已自动压缩）' : '上传成功';
                        setTimeout(function() { progressDiv.style.display = 'none'; }, 2000);
                    } else {
                        alert('上传失败: ' + (resp.error || '未知错误'));
                        progressDiv.style.display = 'none';
                    }
                } catch(ex) {
                    alert('上传响应解析失败');
                    progressDiv.style.display = 'none';
                }
            });
            xhr.addEventListener('error', function() {
                alert('上传请求失败，请检查网络');
                progressDiv.style.display = 'none';
            });
            xhr.open('POST', '/api/upload.asp', true);
            xhr.send(formData);
        }

        (function() {
            var wrapper = document.getElementById('imagePreview_bottle').parentElement;
            wrapper.addEventListener('dragover', function(e) { e.preventDefault(); wrapper.classList.add('dragover'); });
            wrapper.addEventListener('dragleave', function() { wrapper.classList.remove('dragover'); });
            wrapper.addEventListener('drop', function(e) {
                e.preventDefault();
                wrapper.classList.remove('dragover');
                var file = e.dataTransfer.files[0];
                if (file) {
                    document.getElementById('fileInput_bottle').files = e.dataTransfer.files;
                    document.getElementById('fileInput_bottle').dispatchEvent(new Event('change'));
                }
            });
            document.getElementById('imagePreview_bottle').addEventListener('click', function() {
                document.getElementById('fileInput_bottle').click();
            });
        })();

        function toggleUrlInput_bottle() {
            var el = document.getElementById('urlInputWrapper_bottle');
            el.style.display = el.style.display === 'none' ? 'block' : 'none';
        }

        function applyManualUrl_bottle() {
            var url = document.getElementById('manualUrl_bottle').value.trim();
            if (url) {
                document.getElementById('imageURL_bottle').value = url;
                document.getElementById('previewImg_bottle').src = url;
                document.getElementById('previewImg_bottle').style.display = 'block';
                document.getElementById('placeholder_bottle').style.display = 'none';
                document.getElementById('urlInputWrapper_bottle').style.display = 'none';
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
