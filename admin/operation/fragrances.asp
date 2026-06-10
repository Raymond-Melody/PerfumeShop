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
Call OpenConnection

' 处理表单提交
Dim action, noteId, noteName, noteType, description, priceAddition, isActive, image, recommendedPercentage, baseNoteIDs
action = Request.Form("action")

If action = "add" Or action = "edit" Then
    noteName = SafeSQL(Request.Form("noteName"))
    noteType = SafeSQL(Request.Form("noteType"))
    description = SafeSQL(Request.Form("description"))
    priceAddition = Request.Form("priceAddition")
    If priceAddition = "" Or Not IsNumeric(priceAddition) Then priceAddition = 0
    recommendedPercentage = Request.Form("recommendedPercentage")
    If recommendedPercentage = "" Or Not IsNumeric(recommendedPercentage) Then recommendedPercentage = 0
    isActive = Request.Form("isActive")
    If isActive = "" Then isActive = 1
    image = SafeSQL(Request.Form("image"))
    If image = "" Then image = "/images/default-note.jpg"
    ' 获取多选的基香ID数组
    baseNoteIDs = Request.Form("baseNoteSelect")
    
    If action = "add" Then
        Dim addSql, newNoteId
        addSql = "INSERT INTO FragranceNotes (NoteName, NoteType, Description, PriceAddition, RecommendedPercentage, IsActive, ImageURL) VALUES ('" & _
                 noteName & "', '" & noteType & "', '" & description & "', " & CDbl(priceAddition) & ", " & CInt(recommendedPercentage) & ", " & CInt(isActive) & ", '" & image & "')"
        If ExecuteNonQuery(addSql) Then
            ' 获取新插入的香调ID
            Dim rsNewNoteId
            Set rsNewNoteId = ExecuteQuery("SELECT SCOPE_IDENTITY()")
            If Not rsNewNoteId Is Nothing Then
                newNoteId = rsNewNoteId(0)
                rsNewNoteId.Close
                Set rsNewNoteId = Nothing
                
                ' 保存基香关联
                If baseNoteIDs <> "" Then
                    Dim baseNoteArr, bId
                    baseNoteArr = Split(baseNoteIDs, ",")
                    For Each bId In baseNoteArr
                        If IsNumeric(bId) Then
                            ExecuteNonQuery "INSERT INTO NoteIngredients (NoteID, BaseNoteID, Percentage) VALUES (" & newNoteId & ", " & CLng(bId) & ", 0)"
                        End If
                    Next
                End If
            End If
            Response.Redirect "fragrances.asp?msg=添加成功"
        Else
            Response.Write "<script>alert('添加失败：" & Replace(Session("LastDBError"), "'", "\'") & "');</script>"
        End If
    ElseIf action = "edit" Then
        noteId = Request.Form("noteId")
        Dim editSql
        editSql = "UPDATE FragranceNotes SET NoteName = '" & noteName & "', NoteType = '" & noteType & "', " & _
                  "Description = '" & description & "', PriceAddition = " & CDbl(priceAddition) & ", " & _
                  "RecommendedPercentage = " & CInt(recommendedPercentage) & ", " & _
                  "IsActive = " & CInt(isActive) & ", ImageURL = '" & image & "' WHERE NoteID = " & CInt(noteId)
        If ExecuteNonQuery(editSql) Then
            ' 删除旧的基香关联
            ExecuteNonQuery "DELETE FROM NoteIngredients WHERE NoteID = " & CInt(noteId)
            
            ' 保存新的基香关联
            If baseNoteIDs <> "" Then
                Dim baseNoteArrEdit, bIdEdit
                baseNoteArrEdit = Split(baseNoteIDs, ",")
                For Each bIdEdit In baseNoteArrEdit
                    If IsNumeric(bIdEdit) Then
                        ExecuteNonQuery "INSERT INTO NoteIngredients (NoteID, BaseNoteID, Percentage) VALUES (" & CInt(noteId) & ", " & CLng(bIdEdit) & ", 0)"
                    End If
                Next
            End If
            Response.Redirect "fragrances.asp?msg=更新成功"
        Else
            Response.Write "<script>alert('更新失败');</script>"
        End If
    End If
ElseIf action = "delete" Then
    noteId = Request.Form("noteId")
    ' 先删除关联的基香记录
    ExecuteNonQuery "DELETE FROM NoteIngredients WHERE NoteID = " & CInt(noteId)
    ' 再删除香调
    Dim deleteSql
    deleteSql = "DELETE FROM FragranceNotes WHERE NoteID = " & CInt(noteId)
    If ExecuteNonQuery(deleteSql) Then
        Response.Redirect "fragrances.asp?msg=删除成功"
    Else
        Response.Write "<script>alert('删除失败');</script>"
    End If
End If

' 获取香调列表
Dim rsNotes, sql
sql = "SELECT * FROM FragranceNotes ORDER BY NoteType, NoteID DESC"
Set rsNotes = ExecuteQuery(sql)
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>香调管理 - 香氛定制电商网站</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <% If Request.QueryString("msg") <> "" Then %>
        <div class="alert alert-success">
            <i class="fas fa-check-circle"></i>
            <%= Server.HTMLEncode(Request.QueryString("msg")) %>
        </div>
        <% End If %>
        
        <div class="admin-card">
            <div class="admin-card-header">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <h2 class="admin-card-title">香调管理</h2>
                    <button class="admin-btn admin-btn-primary" onclick="showAddForm()">
                        <i class="fas fa-plus"></i> 添加香调
                    </button>
                </div>
            </div>
            
            <div class="admin-card-body">
                <div class="table-responsive">
                    <table class="admin-table">
                        <thead>
                            <tr>
                                <th>ID</th>
                                <th>香调名称</th>
                                <th>类型</th>
                                <th>描述</th>
                                <th>附加单价</th>
                                <th>默认百分比</th>
                                <th>状态</th>
                                <th>操作</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% If Not rsNotes Is Nothing Then %>
                            <% Do While Not rsNotes.EOF %>
                            <tr>
                                <td><%= rsNotes("NoteID") %></td>
                                <td><%= HTMLEncode(rsNotes("NoteName")) %></td>
                                <td><%= HTMLEncode(rsNotes("NoteType")) %></td>
                                <td><%= Left(HTMLEncode(rsNotes("Description") & ""), 50) & "..." %></td>
                                <td><%= FormatMoney(rsNotes("PriceAddition")) %></td>
                                <td><% On Error Resume Next: Response.Write rsNotes("RecommendedPercentage"): If Err.Number <> 0 Then Response.Write "0": Err.Clear: On Error Goto 0 %>%</td>
                                <td>
                                    <span class="status-badge <%= IIf(rsNotes("IsActive") <> 0, "status-paid", "status-cancelled") %>">
                                        <%= IIf(rsNotes("IsActive") <> 0, "启用", "禁用") %>
                                    </span>
                                </td>
                                <td>
                                    <div class="admin-table-actions">
                                        <button class="admin-btn admin-btn-sm admin-btn-outline" onclick="showEditForm(this)" 
                                            data-id="<%= rsNotes("NoteID") %>" 
                                            data-name="<%= HTMLEncode(rsNotes("NoteName") & "") %>" 
                                            data-type="<%= rsNotes("NoteType") %>"
                                            data-desc="<%= HTMLEncode(rsNotes("Description") & "") %>" 
                                            data-price="<%= rsNotes("PriceAddition") %>" 
                                            data-percent="<% On Error Resume Next: Response.Write rsNotes("RecommendedPercentage"): If Err.Number <> 0 Then Response.Write "0": Err.Clear: On Error Goto 0 %>"
                                            data-active="<%= rsNotes("IsActive") %>" 
                                            data-image="<%= HTMLEncode(rsNotes("ImageURL") & "") %>"
                                            data-basenoteids="<%
                                                ' 查询该香调关联的所有基香ID
                                                Dim rsNoteBaseIds
                                                Set rsNoteBaseIds = ExecuteQuery("SELECT BaseNoteID FROM NoteIngredients WHERE NoteID = " & rsNotes("NoteID"))
                                                If Not rsNoteBaseIds Is Nothing Then
                                                    Dim baseIdList, firstId
                                                    baseIdList = ""
                                                    firstId = True
                                                    Do While Not rsNoteBaseIds.EOF
                                                        If Not firstId Then baseIdList = baseIdList & ","
                                                        baseIdList = baseIdList & rsNoteBaseIds("BaseNoteID")
                                                        firstId = False
                                                        rsNoteBaseIds.MoveNext
                                                    Loop
                                                    rsNoteBaseIds.Close
                                                    Set rsNoteBaseIds = Nothing
                                                    Response.Write baseIdList
                                                End If
                                            %>">
                                            <i class="fas fa-edit"></i> 编辑
                                        </button>
                                        <form method="post" style="display:inline;" onsubmit="return confirm('确定要删除吗？')">
                                            <input type="hidden" name="action" value="delete">
                                            <input type="hidden" name="noteId" value="<%= rsNotes("NoteID") %>">
                                            <button type="submit" class="admin-btn admin-btn-sm admin-btn-danger">
                                                <i class="fas fa-trash"></i> 删除
                                            </button>
                                        </form>
                                    </div>
                                </td>
                            </tr>
                            <% rsNotes.MoveNext %>
                            <% Loop %>
                            <% End If %>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    
    <!-- 添加/编辑香调模态框 -->
    <div id="noteModal" class="admin-modal">
        <div class="admin-modal-content">
            <div class="admin-modal-header">
                <h3 id="modalTitle" class="admin-modal-title">添加香调</h3>
                <button class="admin-modal-close" onclick="closeModal()">&times;</button>
            </div>
            <form id="noteForm" method="post">
                <div class="admin-modal-body">
                    <input type="hidden" id="formAction" name="action" value="add">
                    <input type="hidden" id="noteId" name="noteId" value="">
                    
                    <div class="admin-form-group">
                        <label for="noteName" class="admin-form-label">香调名称 *</label>
                        <input type="text" id="noteName" name="noteName" class="admin-form-control" required>
                    </div>
                    
                    <div class="admin-form-row">
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label for="noteType" class="admin-form-label">类型 *</label>
                                <select id="noteType" name="noteType" class="admin-form-control" required>
                                    <option value="前调">前调</option>
                                    <option value="中调">中调</option>
                                    <option value="后调">后调</option>
                                </select>
                            </div>
                        </div>
                        <div class="admin-form-col">
                            <div class="admin-form-group" id="baseNoteField">
                                <label class="admin-form-label">关联基香 <small style="color:#666;">(可多选)</small></label>
                                <div style="max-height: 150px; overflow-y: auto; border: 1px solid #ddd; border-radius: 4px; padding: 10px; background: #fff;">
                                    <%
                                    Dim rsBaseNotesForNote
                                    Set rsBaseNotesForNote = ExecuteQuery("SELECT BaseNoteID, BaseNoteName FROM BaseNotes WHERE IsActive <> 0 ORDER BY BaseNoteID")
                                    If Not rsBaseNotesForNote Is Nothing Then
                                        Do While Not rsBaseNotesForNote.EOF
                                    %>
                                    <label style="display: block; margin-bottom: 5px; cursor: pointer;">
                                        <input type="checkbox" name="baseNoteSelect" value="<%= rsBaseNotesForNote("BaseNoteID") %>" style="margin-right: 8px;">
                                        <%= HTMLEncode(rsBaseNotesForNote("BaseNoteName")) %>
                                    </label>
                                    <%
                                            rsBaseNotesForNote.MoveNext
                                        Loop
                                        rsBaseNotesForNote.Close
                                        Set rsBaseNotesForNote = Nothing
                                    End If
                                    %>
                                </div>
                                <small style="color: #666;">选择该香调包含的基香成分</small>
                            </div>
                        </div>
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label for="priceAddition" class="admin-form-label">附加单价 *</label>
                                <input type="number" id="priceAddition" name="priceAddition" step="0.01" min="0" class="admin-form-control" required>
                            </div>
                        </div>
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label for="recommendedPercentage" class="admin-form-label">默认百分比 (%) *</label>
                                <input type="number" id="recommendedPercentage" name="recommendedPercentage" min="0" max="100" class="admin-form-control" required>
                            </div>
                        </div>
                    </div>
                    
                    <div class="admin-form-group">
                        <label for="description" class="admin-form-label">描述</label>
                        <textarea id="description" name="description" class="admin-form-control" rows="3"></textarea>
                    </div>
                    
                    <div class="admin-form-row">
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label for="image" class="admin-form-label">图片URL</label>
                                <input type="text" id="image" name="image" class="admin-form-control" value="">
                            </div>
                        </div>
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label for="isActive" class="admin-form-label">状态</label>
                                <select id="isActive" name="isActive" class="admin-form-control">
                                    <option value="1">启用</option>
                                    <option value="0">禁用</option>
                                </select>
                            </div>
                        </div>
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
            document.getElementById('modalTitle').textContent = '添加香调';
            document.getElementById('formAction').value = 'add';
            document.getElementById('noteId').value = '';
            document.getElementById('noteName').value = '';
            document.getElementById('noteType').value = '前调';
            document.getElementById('description').value = '';
            document.getElementById('priceAddition').value = '0';
            document.getElementById('recommendedPercentage').value = '0';
            document.getElementById('image').value = '';
            document.getElementById('isActive').value = '1';
            // 清空所有基香复选框
            var checkboxes = document.querySelectorAll('input[name="baseNoteSelect"]');
            checkboxes.forEach(function(cb) {
                cb.checked = false;
            });
            document.getElementById('noteModal').style.display = 'block';
        }
        
        function showEditForm(button) {
            var id = button.getAttribute('data-id');
            var name = button.getAttribute('data-name');
            var type = button.getAttribute('data-type');
            var desc = button.getAttribute('data-desc');
            var price = button.getAttribute('data-price');
            var percent = button.getAttribute('data-percent');
            var active = button.getAttribute('data-active');
            var image = button.getAttribute('data-image');
            var baseNoteIds = button.getAttribute('data-basenoteids') || '';
            
            document.getElementById('modalTitle').textContent = '编辑香调';
            document.getElementById('formAction').value = 'edit';
            document.getElementById('noteId').value = id;
            document.getElementById('noteName').value = name;
            document.getElementById('noteType').value = type;
            document.getElementById('description').value = desc;
            document.getElementById('priceAddition').value = price;
            document.getElementById('recommendedPercentage').value = percent;
            document.getElementById('image').value = image;
            document.getElementById('isActive').value = active;
            
            // 设置基香复选框
            var checkboxes = document.querySelectorAll('input[name="baseNoteSelect"]');
            checkboxes.forEach(function(cb) {
                cb.checked = false;
            });
            
            if (baseNoteIds) {
                var ids = baseNoteIds.split(',');
                ids.forEach(function(id) {
                    var cb = document.querySelector('input[name="baseNoteSelect"][value="' + id.trim() + '"]');
                    if (cb) cb.checked = true;
                });
            }
            
            document.getElementById('noteModal').style.display = 'block';
        }
        
        function closeModal() {
            document.getElementById('noteModal').style.display = 'none';
        }
        
        window.onclick = function(event) {
            var modal = document.getElementById('noteModal');
            if (event.target == modal) {
                modal.style.display = 'none';
            }
        }
    </script>
</body>
</html>
<%
If Not rsNotes Is Nothing Then
    rsNotes.Close
    Set rsNotes = Nothing
End If
Call CloseConnection
%>
