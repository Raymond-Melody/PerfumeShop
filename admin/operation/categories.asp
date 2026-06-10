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

' 处理表单提交
Dim action, catId, catName, sortOrder, isActive, message
action = Request.Form("action")
message = ""

If action = "add" Or action = "edit" Then
    catId = Request.Form("catId")
    catName = Trim(Request.Form("catName"))
    sortOrder = Request.Form("sortOrder")
    isActive = IIF(Request.Form("isActive") = "on", 1, 0)
    
    If catName = "" Then
        message = "分类名称不能为空"
    Else
        Dim sql
        If action = "add" Then
            sql = "INSERT INTO Categories (CategoryName, SortOrder, IsActive) VALUES (" & _
                "'" & SafeSQL(catName) & "', " & _
                CInt(IIF(IsNumeric(sortOrder), sortOrder, 0)) & ", " & _
                isActive & ")"
            
            If ExecuteNonQuery(sql) Then
                message = "✓ 分类添加成功！"
            Else
                message = "✗ 添加失败：" & Session("LastDBError")
            End If
        ElseIf action = "edit" And IsNumeric(catId) Then
            sql = "UPDATE Categories SET " & _
                "CategoryName = '" & SafeSQL(catName) & "', " & _
                "SortOrder = " & CInt(IIF(IsNumeric(sortOrder), sortOrder, 0)) & ", " & _
                "IsActive = " & isActive & " " & _
                "WHERE CategoryID = " & CLng(catId)
            
            If ExecuteNonQuery(sql) Then
                message = "✓ 分类更新成功！"
            Else
                message = "✗ 更新失败：" & Session("LastDBError")
            End If
        End If
    End If
ElseIf action = "delete" Then
    catId = Request.Form("catId")
    If IsNumeric(catId) Then
        If ExecuteNonQuery("DELETE FROM Categories WHERE CategoryID = " & CLng(catId)) Then
            message = "✓ 分类删除成功！"
        Else
            message = "✗ 删除失败：" & Session("LastDBError")
        End If
    End If
End If

' 获取所有分类
Dim rsCats
Set rsCats = ExecuteQuery("SELECT * FROM Categories ORDER BY SortOrder, CategoryName")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>商品分类管理 - 营运管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        /* 深色主题 */
        body {
            background: #1a1a2e;
            color: #e0e0e0;
        }
        .main-content {
            color: #e0e0e0;
        }
        
        /* 页面标题区 */
        .page-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 25px;
            padding-bottom: 15px;
            border-bottom: 1px solid rgba(255,255,255,0.08);
        }
        .page-title {
            font-size: 24px;
            color: #fff;
            margin: 0;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .page-title i { color: #00bcd4; }
        .breadcrumb { font-size: 13px; color: #888; }
        .breadcrumb a { color: #00bcd4; text-decoration: none; }
        .breadcrumb a:hover { text-decoration: underline; }
        
        /* 提示/成功消息 */
        .alert-success {
            background: rgba(76,175,80,0.15);
            border-left: 4px solid #4caf50;
            color: #81c784;
            padding: 12px 18px;
            border-radius: 6px;
            margin-bottom: 20px;
        }
        .alert-error {
            background: rgba(244,67,54,0.15);
            border-left: 4px solid #f44336;
            color: #ef5350;
            padding: 12px 18px;
            border-radius: 6px;
            margin-bottom: 20px;
        }
        
        /* 分类卡片网格 */
        .category-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
        }
        .category-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 24px;
            border: 1px solid rgba(255,255,255,0.05);
            transition: all 0.3s ease;
            display: flex;
            flex-direction: column;
            position: relative;
            overflow: hidden;
        }
        .category-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 4px;
            height: 100%;
            background: linear-gradient(180deg, #00bcd4 0%, #00838f 100%);
            opacity: 0;
            transition: opacity 0.3s ease;
            border-radius: 12px 0 0 12px;
        }
        .category-card:hover::before {
            opacity: 1;
        }
        .category-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.3);
            border-color: rgba(0,188,212,0.2);
        }
        .category-card .cat-icon {
            width: 48px;
            height: 48px;
            border-radius: 10px;
            background: linear-gradient(135deg, rgba(0,188,212,0.15) 0%, rgba(0,131,143,0.1) 100%);
            display: flex;
            align-items: center;
            justify-content: center;
            margin-bottom: 16px;
            font-size: 20px;
            color: #00bcd4;
        }
        .category-card .cat-name {
            font-size: 17px;
            font-weight: 600;
            color: #fff;
            margin-bottom: 6px;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .category-card .cat-id {
            font-size: 11px;
            color: #999;
            background: rgba(0,0,0,0.3);
            padding: 2px 8px;
            border-radius: 4px;
            display: inline-block;
        }
        .category-card .cat-meta {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-bottom: 16px;
            align-items: center;
        }
        .status-badge {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            padding: 3px 8px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 500;
        }
        .status-active { background: rgba(76,175,80,0.2); color: #4caf50; }
        .status-inactive { background: rgba(244,67,54,0.2); color: #ef5350; }
        
        .category-card .cat-footer {
            display: flex;
            gap: 8px;
            padding-top: 16px;
            border-top: 1px solid rgba(255,255,255,0.05);
            margin-top: auto;
        }
        
        /* 空状态 */
        .empty-state {
            grid-column: 1 / -1;
            text-align: center;
            padding: 60px 20px;
            color: #666;
        }
        .empty-state i {
            font-size: 48px;
            margin-bottom: 15px;
            color: #555;
        }
        .empty-state h3 { color: #888; margin-bottom: 10px; }
        .empty-state p { color: #666; font-size: 14px; }
        
        /* 模态框深色 */
        .modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.7); z-index: 1000; }
        .modal.active { display: flex; align-items: center; justify-content: center; }
        .modal-content {
            background: #1e1e32;
            padding: 30px;
            border-radius: 12px;
            max-width: 500px;
            width: 90%;
            border: 1px solid rgba(255,255,255,0.08);
            box-shadow: 0 20px 60px rgba(0,0,0,0.5);
        }
        .modal-content h2 {
            color: #fff;
            margin-bottom: 20px;
            font-size: 20px;
        }
        .modal-content .form-group {
            margin-bottom: 18px;
        }
        .modal-content .form-label {
            display: block;
            color: #aaa;
            font-size: 13px;
            margin-bottom: 6px;
            font-weight: 500;
        }
        .modal-content .form-control {
            width: 100%;
            padding: 10px 12px;
            background: rgba(255,255,255,0.05);
            border: 1px solid rgba(255,255,255,0.12);
            border-radius: 6px;
            color: #e0e0e0;
            font-size: 14px;
            box-sizing: border-box;
            transition: border-color 0.2s;
        }
        .modal-content .form-control:focus {
            border-color: #00bcd4;
            outline: none;
        }
        .modal-content .checkbox-label {
            color: #ccc;
            font-size: 14px;
            display: flex;
            align-items: center;
            gap: 8px;
            cursor: pointer;
        }
        .modal-content .checkbox-label input[type="checkbox"] {
            accent-color: #00bcd4;
        }
        .modal-footer-btns {
            display: flex;
            justify-content: flex-end;
            gap: 10px;
            margin-top: 24px;
            padding-top: 20px;
            border-top: 1px solid rgba(255,255,255,0.05);
        }
        
        /* 响应式 */
        @media (max-width: 1200px) {
            .category-grid { grid-template-columns: repeat(2, 1fr); }
        }
        @media (max-width: 768px) {
            .category-grid { grid-template-columns: 1fr; }
            .page-header { flex-direction: column; align-items: flex-start; gap: 12px; }
        }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <div>
                <h2 class="page-title"><i class="fas fa-tags"></i> 商品分类管理</h2>
                <div class="breadcrumb">
                    <a href="index.asp">运营中心</a> / <span>商品分类</span>
                </div>
            </div>
            <button class="btn btn-primary" onclick="openAddModal()">
                <i class="fas fa-plus"></i> 添加新分类
            </button>
        </div>
        
        <% If message <> "" Then %>
        <div class="<%= IIF(InStr(message, "✓") > 0, "alert-success", "alert-error") %>">
            <i class="fas <%= IIF(InStr(message, "✓") > 0, "fa-check-circle", "fa-exclamation-circle") %>"></i>
            <%= message %>
        </div>
        <% End If %>
        
        <!-- 分类卡片网格 -->
        <div class="category-grid">
            <%
            Dim hasCats
            hasCats = False
            If Not rsCats Is Nothing Then
                If Not rsCats.EOF Then hasCats = True
            End If
            
            If hasCats Then
                Do While Not rsCats.EOF
                    Dim catIdVal, catNameVal, catSortVal, catActiveVal, catActiveBool
                    catIdVal = rsCats("CategoryID")
                    catNameVal = HTMLEncode(rsCats("CategoryName") & "")
                    catSortVal = rsCats("SortOrder")
                    catActiveBool = (rsCats("IsActive") <> 0)
                    catActiveVal = IIF(catActiveBool, "true", "false")
            %>
            <div class="category-card">
                <div class="cat-icon">
                    <i class="fas fa-folder"></i>
                </div>
                <div class="cat-name">
                    <span><%= catNameVal %></span>
                </div>
                <span class="cat-id">#<%= catIdVal %></span>
                
                <div class="cat-meta" style="margin-top: 10px;">
                    <span style="font-size: 12px; color: #888;">
                        <i class="fas fa-sort-numeric-down" style="color: #00bcd4;"></i> 排序: <%= catSortVal %>
                    </span>
                    <span class="status-badge <%= IIF(catActiveBool, "status-active", "status-inactive") %>">
                        <i class="fas <%= IIF(catActiveBool, "fa-check-circle", "fa-ban") %>"></i>
                        <%= IIF(catActiveBool, "启用", "禁用") %>
                    </span>
                </div>
                
                <div class="cat-footer">
                    <button class="btn btn-outline btn-sm" onclick="openEditModal(<%= catIdVal %>, '<%= Replace(rsCats("CategoryName") & "", "'", "\'") %>', <%= catSortVal %>, <%= catActiveVal %>)">
                        <i class="fas fa-edit"></i> 编辑
                    </button>
                    <button class="btn btn-danger btn-sm" onclick="deleteCat(<%= catIdVal %>)">
                        <i class="fas fa-trash"></i> 删除
                    </button>
                </div>
            </div>
            <%
                    rsCats.MoveNext
                Loop
                rsCats.Close
                Set rsCats = Nothing
            Else
            %>
            <div class="empty-state">
                <i class="fas fa-tags"></i>
                <h3>暂无分类</h3>
                <p>点击上方按钮添加第一个商品分类</p>
            </div>
            <% End If %>
        </div>
    </div>
    
<%
' 确保记录集被关闭
If Not rsCats Is Nothing Then
    On Error Resume Next
    rsCats.Close
    Set rsCats = Nothing
    On Error GoTo 0
End If
%>

    <!-- 模态框 -->
    <div id="catModal" class="modal">
        <div class="modal-content">
            <h2 id="modalTitle">添加分类</h2>
            <form method="post">
                <input type="hidden" id="formAction" name="action" value="add">
                <input type="hidden" id="catId" name="catId" value="">
                
                <div class="form-group">
                    <label class="form-label">分类名称</label>
                    <input type="text" id="catName" name="catName" class="form-control" required placeholder="输入分类名称">
                </div>
                
                <div class="form-group">
                    <label class="form-label">排序 (越小越靠前)</label>
                    <input type="number" id="sortOrder" name="sortOrder" class="form-control" value="0">
                </div>
                
                <div class="form-group">
                    <label class="checkbox-label">
                        <input type="checkbox" id="isActive" name="isActive" checked>
                        <span>启用该分类</span>
                    </label>
                </div>
                
                <div class="modal-footer-btns">
                    <button type="button" class="btn btn-outline" onclick="closeModal()">取消</button>
                    <button type="submit" class="btn btn-primary">
                        <i class="fas fa-save"></i> 保存
                    </button>
                </div>
            </form>
        </div>
    </div>

    <script>
        function openAddModal() {
            document.getElementById('modalTitle').textContent = '添加分类';
            document.getElementById('formAction').value = 'add';
            document.getElementById('catId').value = '';
            document.getElementById('catName').value = '';
            document.getElementById('sortOrder').value = '0';
            document.getElementById('isActive').checked = true;
            document.getElementById('catModal').classList.add('active');
        }
        
        function openEditModal(id, name, sort, active) {
            document.getElementById('modalTitle').textContent = '编辑分类';
            document.getElementById('formAction').value = 'edit';
            document.getElementById('catId').value = id;
            document.getElementById('catName').value = name;
            document.getElementById('sortOrder').value = sort;
            document.getElementById('isActive').checked = active;
            document.getElementById('catModal').classList.add('active');
        }
        
        function closeModal() {
            document.getElementById('catModal').classList.remove('active');
        }
        
        // 点击模态框外部关闭
        document.getElementById('catModal').addEventListener('click', function(e) {
            if (e.target === this) closeModal();
        });
        
        function deleteCat(id) {
            if (confirm('确定要删除此分类吗？关联商品的分类信息将失效。')) {
                var form = document.createElement('form');
                form.method = 'POST';
                form.innerHTML = '<input type="hidden" name="action" value="delete"><input type="hidden" name="catId" value="' + id + '">';
                document.body.appendChild(form);
                form.submit();
            }
        }
    </script>
</body>
</html>
<% Call CloseConnection() %>
