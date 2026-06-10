<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->

<%
' 已迁移到产品技术管理中心
Response.Redirect "../techcenter/product_settings.asp"
Response.End
%>

<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
Call OpenConnection()

' 处理表单提交
Dim action, baseNoteId, baseNoteName, description, ingredients, isActive, message
action = Request.Form("action")
message = ""

If action = "add" Or action = "edit" Then
    baseNoteId = Request.Form("baseNoteId")
    baseNoteName = Trim(Request.Form("baseNoteName"))
    description = Trim(Request.Form("description"))
    ingredients = Trim(Request.Form("ingredients"))
    
    ' 将换行符分隔的成分转换为逗号分隔，确保后续去重能正常工作
    If InStr(ingredients, vbCrLf) > 0 Or InStr(ingredients, vbLf) > 0 Or InStr(ingredients, vbCr) > 0 Then
        Dim ingLines, ingLine, ingResult
        ' 先统一转换为vbLf
        ingredients = Replace(ingredients, vbCrLf, vbLf)
        ingredients = Replace(ingredients, vbCr, vbLf)
        ingLines = Split(ingredients, vbLf)
        ingResult = ""
        For Each ingLine In ingLines
            ingLine = Trim(ingLine)
            If ingLine <> "" Then
                If ingResult <> "" Then ingResult = ingResult & ","
                ingResult = ingResult & ingLine
            End If
        Next
        ingredients = ingResult
    End If
    
    isActive = Request.Form("isActive")
    If isActive = "" Then isActive = 0
    
    If baseNoteName = "" Then
        message = "基香名称不能为空"
    Else
        Dim sql
        If action = "add" Then
            sql = "INSERT INTO BaseNotes (BaseNoteName, Description, Ingredients, IsActive) VALUES (" & _
                "'" & SafeSQL(baseNoteName) & "', " & _
                "'" & SafeSQL(description) & "', " & _
                "'" & SafeSQL(ingredients) & "', " & _
                isActive & ")"
            
            If ExecuteNonQuery(sql) Then
                message = "✓ 基香添加成功！"
            Else
                message = "✗ 添加失败：" & Session("LastDBError")
            End If
        ElseIf action = "edit" And IsNumeric(baseNoteId) Then
            sql = "UPDATE BaseNotes SET " & _
                "BaseNoteName = '" & SafeSQL(baseNoteName) & "', " & _
                "Description = '" & SafeSQL(description) & "', " & _
                "Ingredients = '" & SafeSQL(ingredients) & "', " & _
                "IsActive = " & isActive & " " & _
                "WHERE BaseNoteID = " & CLng(baseNoteId)
            
            If ExecuteNonQuery(sql) Then
                message = "✓ 基香更新成功！"
            Else
                message = "✗ 更新失败：" & Session("LastDBError")
            End If
        End If
    End If
ElseIf action = "delete" Then
    baseNoteId = Request.Form("baseNoteId")
    If IsNumeric(baseNoteId) Then
        If ExecuteNonQuery("DELETE FROM BaseNotes WHERE BaseNoteID = " & CLng(baseNoteId)) Then
            message = "✓ 基香删除成功！"
        Else
            message = "✗ 删除失败：" & Session("LastDBError")
        End If
    End If
End If

' 获取所有基香列表
Dim rsBaseNotes
Set rsBaseNotes = ExecuteQuery("SELECT * FROM BaseNotes ORDER BY BaseNoteID DESC")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>基香管理 - 香氛定制电商网站</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        .base-note-badge {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 4px;
            font-size: 12px;
            font-weight: bold;
            background: #e7f3ff;
            color: #0066cc;
            border: 1px solid #b3d9ff;
        }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <% If message <> "" Then %>
        <div class="alert <%= IIF(InStr(message, "✓") > 0, "alert-success", "alert-error") %>">
            <i class="fas fa-<%= IIF(InStr(message, "✓") > 0, "check-circle", "exclamation-circle") %>"></i>
            <%= message %>
        </div>
        <% End If %>
        
        <div class="admin-card">
            <div class="admin-card-header">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <h2 class="admin-card-title"><i class="fas fa-flask"></i> 基香成分管理</h2>
                    <button class="admin-btn admin-btn-primary" onclick="showAddForm()">
                        <i class="fas fa-plus"></i> 添加基香
                    </button>
                </div>
            </div>
            
            <div class="admin-card-body">
                <div class="alert alert-info" style="margin-bottom: 20px;">
                    <i class="fas fa-info-circle"></i>
                    <strong>说明：</strong>基香是香水的基础成分，用于品牌定香类商品的成分信息录入。添加基香后，可在商品管理中为品牌定香商品关联相应基香，以便后续过敏源追溯。
                </div>
                
                <div class="table-responsive">
                    <table class="admin-table">
                        <thead>
                            <tr>
                                <th>ID</th>
                                <th>基香名称</th>
                                <th>描述</th>
                                <th>成分列表</th>
                                <th>状态</th>
                                <th>操作</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% 
                            If Not rsBaseNotes Is Nothing Then 
                                Do While Not rsBaseNotes.EOF
                            %>
                            <tr>
                                <td><%= rsBaseNotes("BaseNoteID") %></td>
                                <td><strong><%= HTMLEncode(rsBaseNotes("BaseNoteName")) %></strong></td>
                                <td>
                                    <% 
                                    Dim baseDesc
                                    baseDesc = Trim(rsBaseNotes("Description") & "")
                                    If baseDesc <> "" Then
                                        Response.Write Left(HTMLEncode(baseDesc), 50)
                                        If Len(baseDesc) > 50 Then Response.Write "..."
                                    Else
                                        Response.Write "<span style='color:#999;'>-</span>"
                                    End If
                                    %>
                                </td>
                                <td>
                                    <% 
                                    Dim baseIngredients
                                    baseIngredients = Trim(rsBaseNotes("Ingredients") & "")
                                    If baseIngredients <> "" Then
                                        Response.Write "<span class='base-note-badge'><i class='fas fa-leaf'></i> 有成分</span>"
                                    Else
                                        Response.Write "<span style='color:#999;'>未设置</span>"
                                    End If
                                    %>
                                </td>
                                <td>
                                    <span class="status-badge <%= IIF(rsBaseNotes("IsActive"), "status-paid", "status-cancelled") %>">
                                        <%= IIF(rsBaseNotes("IsActive"), "启用", "禁用") %>
                                    </span>
                                </td>
                                <td>
                                    <div class="admin-table-actions">
                                        <button class="admin-btn admin-btn-sm admin-btn-outline" onclick="showEditForm(this)" 
                                            data-id="<%= rsBaseNotes("BaseNoteID") %>" 
                                            data-name="<%= Replace(rsBaseNotes("BaseNoteName") & "", """", "&quot;") %>" 
                                            data-desc="<%= Replace(rsBaseNotes("Description") & "", """", "&quot;") %>" 
                                            data-ingredients="<%= Replace(rsBaseNotes("Ingredients") & "", """", "&quot;") %>"
                                            data-active="<%= rsBaseNotes("IsActive") %>">
                                            <i class="fas fa-edit"></i> 编辑
                                        </button>
                                        <form method="post" style="display:inline;" onsubmit="return confirm('确定要删除此基香吗？')">
                                            <input type="hidden" name="action" value="delete">
                                            <input type="hidden" name="baseNoteId" value="<%= rsBaseNotes("BaseNoteID") %>">
                                            <button type="submit" class="admin-btn admin-btn-sm admin-btn-danger">
                                                <i class="fas fa-trash"></i> 删除
                                            </button>
                                        </form>
                                    </div>
                                </td>
                            </tr>
                            <% 
                                    rsBaseNotes.MoveNext
                                Loop
                                rsBaseNotes.Close
                                Set rsBaseNotes = Nothing
                            Else
                                Response.Write "<tr><td colspan='6' style='text-align:center;color:#999;'>暂无数据</td></tr>"
                            End If
                            %>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    
    <!-- 添加/编辑基香模态框 -->
    <div id="baseNoteModal" class="admin-modal">
        <div class="admin-modal-content">
            <div class="admin-modal-header">
                <h3 id="modalTitle" class="admin-modal-title">添加基香</h3>
                <button class="admin-modal-close" onclick="closeModal()">&times;</button>
            </div>
            <form id="baseNoteForm" method="post">
                <div class="admin-modal-body">
                    <input type="hidden" id="formAction" name="action" value="add">
                    <input type="hidden" id="baseNoteId" name="baseNoteId" value="">
                    
                    <div class="admin-form-group">
                        <label for="baseNoteName" class="admin-form-label">基香名称 *</label>
                        <input type="text" id="baseNoteName" name="baseNoteName" class="admin-form-control" required>
                    </div>
                    
                    <div class="admin-form-group">
                        <label for="description" class="admin-form-label">描述</label>
                        <textarea id="description" name="description" class="admin-form-control" rows="3" placeholder="输入基香的描述信息"></textarea>
                    </div>
                    
                    <div class="admin-form-group">
                        <label for="ingredients" class="admin-form-label">成分列表 *（逗号或换行分隔）</label>
                        <textarea id="ingredients" name="ingredients" class="admin-form-control" rows="5" placeholder="例如：乙醇, 蒸馏水, 香精油&#10;或每行一个成分" required></textarea>
                        <div style="margin-top: 8px; padding: 10px; background: #f8f9fa; border-radius: 4px; font-size: 12px; color: #666;">
                            <i class="fas fa-info-circle"></i> 重要提示：
                            <ul style="margin: 5px 0 0 0; padding-left: 20px;">
                                <li>支持逗号分隔（中文或英文逗号均可）</li>
                                <li>也支持每行一个成分</li>
                                <li>此信息用于后续过敏源追溯</li>
                            </ul>
                        </div>
                    </div>
                    
                    <div class="admin-form-group">
                        <label for="isActive" class="admin-form-label">状态</label>
                        <select id="isActive" name="isActive" class="admin-form-control">
                            <option value="1">启用</option>
                            <option value="0">禁用</option>
                        </select>
                    </div>
                </div>
                <div class="admin-modal-footer">
                    <button type="button" class="admin-btn admin-btn-outline" onclick="closeModal()">取消</button>
                    <button type="submit" class="admin-btn admin-btn-primary">保存</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
        function showAddForm() {
            document.getElementById('modalTitle').textContent = '添加基香';
            document.getElementById('formAction').value = 'add';
            document.getElementById('baseNoteId').value = '';
            document.getElementById('baseNoteName').value = '';
            document.getElementById('description').value = '';
            document.getElementById('ingredients').value = '';
            document.getElementById('isActive').value = '1';
            document.getElementById('baseNoteModal').style.display = 'block';
        }
        
        function showEditForm(button) {
            var id = button.getAttribute('data-id');
            var name = button.getAttribute('data-name');
            var desc = button.getAttribute('data-desc');
            var ingredients = button.getAttribute('data-ingredients') || '';
            var active = button.getAttribute('data-active');
            
            // 将逗号分隔的成分转换为换行符显示，方便编辑
            if (ingredients && ingredients.indexOf(',') !== -1) {
                ingredients = ingredients.split(',').join('\n');
            }
            
            document.getElementById('modalTitle').textContent = '编辑基香';
            document.getElementById('formAction').value = 'edit';
            document.getElementById('baseNoteId').value = id;
            document.getElementById('baseNoteName').value = name;
            document.getElementById('description').value = desc || '';
            document.getElementById('ingredients').value = ingredients;
            document.getElementById('isActive').value = active;
            document.getElementById('baseNoteModal').style.display = 'block';
        }
        
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
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
