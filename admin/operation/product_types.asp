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
<!--#include file="../../includes/product_type_utils.asp"-->
<%
Call OpenConnection()

' 检查权限 - 只有 OP_MANAGER 和 SUPER_ADMIN 可以修改
Dim canEdit
canEdit = (Session("AdminRoleCode") = "OP_MANAGER" OR Session("AdminRoleCode") = "SUPER_ADMIN")

' 处理POST请求（添加、编辑、删除）
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' CSRF验证
    If Not ValidateCSRFToken() Then
        Session("ErrorMessage") = "安全验证失败，请重新操作"
        Response.Redirect "product_types.asp"
        Response.End
    End If
    
    Dim postAction, configId
    postAction = Request.Form("action")
    
    If postAction = "delete" Then
        ' 删除操作
        configId = Request.Form("configId")
        If configId <> "" And IsNumeric(configId) Then
            ' 检查该类型下是否有商品
            Dim productCountVal, productCount
            productCountVal = GetScalar("SELECT COUNT(*) FROM Products WHERE ProductType = (SELECT TypeCode FROM ProductTypeConfig WHERE ConfigID = " & CLng(configId) & ")")
            productCount = CLng("0" & productCountVal)
            
            If productCount > 0 Then
                ' 获取该类型下的上架商品数量
                Dim activeProductCountVal, activeProductCount
                activeProductCountVal = GetScalar("SELECT COUNT(*) FROM Products WHERE ProductType = (SELECT TypeCode FROM ProductTypeConfig WHERE ConfigID = " & CLng(configId) & ") AND IsActive<>0")
                activeProductCount = CLng("0" & activeProductCountVal)
                Session("ErrorMessage") = "该类型下有 " & productCount & " 件商品（其中 " & activeProductCount & " 件上架中），无法删除。请先将商品移至其他类型。"
            Else
                Dim deleteSql
                deleteSql = "DELETE FROM ProductTypeConfig WHERE ConfigID = " & CLng(configId)
                If ExecuteNonQuery(deleteSql) Then
                    Session("SuccessMessage") = "商品类型删除成功"
                    Call LogAdminAction("删除商品类型", "product_type", "ProductTypeConfig", configId, "")
                Else
                    Session("ErrorMessage") = "删除失败: " & Session("LastDBError")
                End If
            End If
        End If
    ElseIf postAction = "add" OR postAction = "edit" Then
        ' 添加或编辑操作
        Dim typeCode, displayName, navName, description, icon, requiresReview, requiresRatio, displayOrder, isActive
        
        typeCode = SafeSQL(Request.Form("typeCode"))
        displayName = SafeSQL(Request.Form("displayName"))
        navName = SafeSQL(Request.Form("navName"))
        description = SafeSQL(Request.Form("description"))
        icon = SafeSQL(Request.Form("icon"))
        displayOrder = Request.Form("displayOrder")
        
        ' 复选框处理 - 未勾选时不会提交
        If Len(Request.Form("requiresReview")) > 0 Then
            requiresReview = 1
        Else
            requiresReview = 0
        End If
        
        If Len(Request.Form("requiresRatio")) > 0 Then
            requiresRatio = 1
        Else
            requiresRatio = 0
        End If
        
        If Len(Request.Form("isActive")) > 0 Then
            isActive = 1
        Else
            isActive = 0
        End If
        
        ' 验证必填字段
        If typeCode = "" OR displayName = "" Then
            Session("ErrorMessage") = "类型代码和显示名称不能为空"
        Else
            If postAction = "add" Then
                ' 检查类型代码是否已存在
                Dim existCountVal, existCount
                existCountVal = GetScalar("SELECT COUNT(*) FROM ProductTypeConfig WHERE TypeCode = '" & typeCode & "'")
                existCount = CLng("0" & existCountVal)
                
                If existCount > 0 Then
                    Session("ErrorMessage") = "类型代码 '" & typeCode & "' 已存在"
                Else
                    Dim addSql
                    addSql = "INSERT INTO ProductTypeConfig (TypeCode, DisplayName, NavName, Description, Icon, RequiresReview, RequiresRatio, DisplayOrder, IsActive) VALUES ('" & _
                             typeCode & "', '" & displayName & "', '" & navName & "', '" & description & "', '" & icon & "', " & requiresReview & ", " & requiresRatio & ", " & CLng("0" & displayOrder) & ", " & isActive & ")"
                    If ExecuteNonQuery(addSql) Then
                        Session("SuccessMessage") = "商品类型添加成功"
                        Call LogAdminAction("添加商品类型", "product_type", "ProductTypeConfig", "", typeCode)
                    Else
                        Session("ErrorMessage") = "添加失败: " & Session("LastDBError")
                    End If
                End If
            ElseIf postAction = "edit" Then
                configId = Request.Form("configId")
                If configId <> "" And IsNumeric(configId) Then
                    Dim editSql
                    editSql = "UPDATE ProductTypeConfig SET DisplayName = '" & displayName & "', NavName = '" & navName & "', Description = '" & description & _
                              "', Icon = '" & icon & "', RequiresReview = " & requiresReview & ", RequiresRatio = " & requiresRatio & _
                              ", DisplayOrder = " & CLng("0" & displayOrder) & ", IsActive = " & isActive & " WHERE ConfigID = " & CLng(configId)
                    If ExecuteNonQuery(editSql) Then
                        Session("SuccessMessage") = "商品类型更新成功"
                        Call LogAdminAction("编辑商品类型", "product_type", "ProductTypeConfig", configId, typeCode)
                    Else
                        Session("ErrorMessage") = "更新失败: " & Session("LastDBError")
                    End If
                End If
            End If
        End If
    End If
    
    Response.Redirect "product_types.asp"
    Response.End
End If

' 获取查询参数
Dim qsAction, editId
qsAction = Request.QueryString("action")
editId = Request.QueryString("id")

' 预加载所有商品类型数据（必须在打开主Recordset前执行，避免Access MARS问题）
Dim allTypes, typeCount
typeCount = 0

Dim rsTypes
Set rsTypes = ExecuteQuery("SELECT ConfigID, TypeCode, DisplayName, NavName, Description, Icon, RequiresReview, RequiresRatio, DisplayOrder, IsActive FROM ProductTypeConfig ORDER BY DisplayOrder ASC, ConfigID ASC")

If Not rsTypes Is Nothing Then
    ' 先统计记录数
    Do While Not rsTypes.EOF
        typeCount = typeCount + 1
        rsTypes.MoveNext
    Loop
    
    ' 如果有数据，重新定位到开头并加载到数组
    If typeCount > 0 Then
        rsTypes.MoveFirst
        ReDim allTypes(typeCount - 1, 9)
        
        Dim idx
        idx = 0
        Do While Not rsTypes.EOF
            allTypes(idx, 0) = rsTypes("ConfigID").Value
            allTypes(idx, 1) = rsTypes("TypeCode").Value
            allTypes(idx, 2) = rsTypes("DisplayName").Value
            allTypes(idx, 3) = rsTypes("NavName").Value
            allTypes(idx, 4) = rsTypes("Description").Value
            allTypes(idx, 5) = rsTypes("Icon").Value
            allTypes(idx, 6) = rsTypes("RequiresReview").Value
            allTypes(idx, 7) = rsTypes("RequiresRatio").Value
            allTypes(idx, 8) = rsTypes("DisplayOrder").Value
            allTypes(idx, 9) = rsTypes("IsActive").Value
            idx = idx + 1
            rsTypes.MoveNext
        Loop
    End If
    
    rsTypes.Close
End If
Set rsTypes = Nothing

' 统计每个类型的商品数量（在打开其他Recordset前执行）
' 使用Dictionary存储每种类型的统计数据：key=TypeCode, value=数组[总数, 上架数]
Dim productStats
Set productStats = CreateObject("Scripting.Dictionary")

Dim rsStats
Set rsStats = ExecuteQuery("SELECT ProductType, COUNT(*) AS Total, SUM(IIF(IsActive<>0, 1, 0)) AS ActiveCount FROM Products GROUP BY ProductType")
Dim statKey, statArray
If Not rsStats Is Nothing Then
    Do While Not rsStats.EOF
        statKey = CStr(rsStats("ProductType").Value)
        ReDim statArray(1)
        statArray(0) = CLng("0" & rsStats("Total").Value)      ' 总数
        statArray(1) = CLng("0" & rsStats("ActiveCount").Value) ' 上架数
        productStats.Add statKey, statArray
        rsStats.MoveNext
    Loop
    rsStats.Close
End If
Set rsStats = Nothing

' 如果是编辑模式，获取要编辑的数据
Dim editData(9)
If qsAction = "edit" And editId <> "" And IsNumeric(editId) Then
    Dim rsEdit
    Set rsEdit = ExecuteQuery("SELECT * FROM ProductTypeConfig WHERE ConfigID = " & CLng(editId))
    If Not rsEdit Is Nothing Then
        If Not rsEdit.EOF Then
            editData(0) = rsEdit("ConfigID").Value
            editData(1) = rsEdit("TypeCode").Value
            editData(2) = rsEdit("DisplayName").Value
            editData(3) = rsEdit("NavName").Value
            editData(4) = rsEdit("Description").Value
            editData(5) = rsEdit("Icon").Value
            editData(6) = rsEdit("RequiresReview").Value
            editData(7) = rsEdit("RequiresRatio").Value
            editData(8) = rsEdit("DisplayOrder").Value
            editData(9) = rsEdit("IsActive").Value
        End If
        rsEdit.Close
    End If
    Set rsEdit = Nothing
End If

' 获取提示消息
Dim successMsg, errorMsg
successMsg = Session("SuccessMessage")
errorMsg = Session("ErrorMessage")
Session("SuccessMessage") = ""
Session("ErrorMessage") = ""

' 记录访问日志
Call LogAdminAction("查看商品类型列表", "product_type", "", "", "")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>商品类型管理 - 运营管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        .page-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 25px; padding-bottom: 15px; border-bottom: 1px solid #e0e0e0; }
        .page-title { font-size: 24px; color: #333; margin: 0; }
        .page-title i { color: #667eea; margin-right: 10px; }
        .breadcrumb { font-size: 14px; color: #666; }
        .breadcrumb a { color: #667eea; text-decoration: none; }
        .breadcrumb a:hover { text-decoration: underline; }
        
        .alert { padding: 15px 20px; border-radius: 8px; margin-bottom: 20px; }
        .alert-success { background: #e8f5e9; color: #2e7d32; border-left: 4px solid #4CAF50; }
        .alert-error { background: #ffebee; color: #c62828; border-left: 4px solid #f44336; }
        
        .admin-card { background: white; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); margin-bottom: 25px; }
        .admin-card-header { padding: 20px; border-bottom: 1px solid #f0f0f0; display: flex; justify-content: space-between; align-items: center; }
        .admin-card-title { font-size: 18px; color: #333; margin: 0; }
        .admin-card-body { padding: 20px; }
        
        .admin-btn { padding: 10px 20px; border-radius: 6px; font-size: 14px; cursor: pointer; border: none; text-decoration: none; display: inline-block; transition: all 0.3s; }
        .admin-btn-primary { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; }
        .admin-btn-primary:hover { opacity: 0.9; }
        .admin-btn-primary:disabled { background: #ccc; cursor: not-allowed; }
        .admin-btn-secondary { background: #f0f0f0; color: #666; }
        .admin-btn-secondary:hover { background: #e0e0e0; }
        .admin-btn-danger { background: #ffebee; color: #c62828; }
        .admin-btn-danger:hover { background: #c62828; color: white; }
        .admin-btn-sm { padding: 6px 12px; font-size: 12px; }
        
        .admin-table { width: 100%; border-collapse: collapse; }
        .admin-table th { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 15px; text-align: left; font-weight: 500; }
        .admin-table td { padding: 15px; border-bottom: 1px solid #f0f0f0; }
        .admin-table tr:hover { background: #f8f9fa; }
        .admin-table tr:last-child td { border-bottom: none; }
        
        .status-badge { display: inline-block; padding: 6px 14px; border-radius: 20px; font-size: 12px; font-weight: 500; }
        .status-active { background: #e8f5e9; color: #2e7d32; }
        .status-inactive { background: #ffebee; color: #c62828; }
        
        .icon-preview { width: 40px; height: 40px; border-radius: 8px; background: #f0f0f0; display: flex; align-items: center; justify-content: center; color: #667eea; font-size: 18px; }
        
        .action-btns { display: flex; gap: 8px; }
        
        .empty-state { text-align: center; padding: 60px 20px; background: white; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
        .empty-state i { font-size: 64px; color: #ddd; margin-bottom: 20px; }
        .empty-state h3 { color: #666; margin-bottom: 10px; }
        .empty-state p { color: #999; }
        
        /* 表单样式 */
        .form-container { max-width: 600px; }
        .admin-form-group { margin-bottom: 20px; }
        .admin-form-label { display: block; margin-bottom: 8px; font-weight: 500; color: #333; }
        .admin-form-label .required { color: #f44336; }
        .admin-form-control { width: 100%; padding: 12px 15px; border: 2px solid #e0e0e0; border-radius: 8px; font-size: 14px; box-sizing: border-box; }
        .admin-form-control:focus { border-color: #667eea; outline: none; }
        .admin-form-control:read-only { background: #f5f5f5; cursor: not-allowed; }
        textarea.admin-form-control { resize: vertical; min-height: 100px; }
        
        .checkbox-group { display: flex; align-items: center; gap: 10px; }
        .checkbox-group input[type="checkbox"] { width: 18px; height: 18px; cursor: pointer; }
        .checkbox-group label { cursor: pointer; margin: 0; }
        
        .form-actions { display: flex; gap: 15px; margin-top: 30px; padding-top: 20px; border-top: 1px solid #f0f0f0; }
        
        .readonly-notice { background: #e3f2fd; padding: 10px 15px; border-radius: 6px; color: #1976d2; font-size: 13px; margin-bottom: 20px; }
        
        .help-text { font-size: 12px; color: #999; margin-top: 5px; }
    </style>
    <script>
        // 商品统计数据，由服务器端生成
        var productStatsData = {};
        <% 
        ' 将统计数据输出到JavaScript
        Dim jsKey
        For Each jsKey In productStats.Keys
            Response.Write "productStatsData['" & jsKey & "'] = {total: " & productStats(jsKey)(0) & ", active: " & productStats(jsKey)(1) & "};" & vbCrLf
        Next
        %>
        
        // 删除确认函数
        function confirmDelete(typeCode) {
            var stats = productStatsData[typeCode];
            if (stats && stats.total > 0) {
                var msg = "该类型下有 " + stats.total + " 件商品（其中 " + stats.active + " 件上架中），无法删除。请先将商品移至其他类型。";
                alert(msg);
                return false;
            }
            return confirm("确定要删除此商品类型吗？删除后不可恢复。");
        }
        
        // 禁用确认函数（用于编辑表单中的启用复选框）
        function confirmDisable(typeCode) {
            var stats = productStatsData[typeCode];
            if (stats && stats.active > 0) {
                return confirm("该类型下有 " + stats.active + " 件上架商品，禁用后这些商品将不在前台显示，确定要禁用吗？");
            }
            return true;
        }
    </script>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <div>
                <h2 class="page-title"><i class="fas fa-tags"></i> 商品类型管理</h2>
                <div class="breadcrumb">
                    <a href="index.asp">运营中心</a> / <span>商品类型</span>
                </div>
            </div>
            <% If qsAction = "" Then %>
            <a href="product_types.asp?action=add" class="admin-btn admin-btn-primary <%= IIf(canEdit, "", "disabled") %>" <%= IIf(canEdit, "", "onclick='return false;'") %>>
                <i class="fas fa-plus"></i> 添加类型
            </a>
            <% End If %>
        </div>
        
        <% If successMsg <> "" Then %>
        <div class="alert alert-success"><i class="fas fa-check-circle"></i> <%= SafeOutput(successMsg) %></div>
        <% End If %>
        
        <% If errorMsg <> "" Then %>
        <div class="alert alert-error"><i class="fas fa-exclamation-circle"></i> <%= SafeOutput(errorMsg) %></div>
        <% End If %>
        
        <% If Not canEdit Then %>
        <div class="readonly-notice">
            <i class="fas fa-info-circle"></i> 您当前为运营专员，仅可查看不可修改
        </div>
        <% End If %>
        
        <% If qsAction = "add" OR qsAction = "edit" Then %>
        <!-- 添加/编辑表单 -->
        <div class="admin-card">
            <div class="admin-card-header">
                <h3 class="admin-card-title"><%= IIf(qsAction = "add", "添加商品类型", "编辑商品类型") %></h3>
            </div>
            <div class="admin-card-body">
                <% If canEdit Then %>
                <form method="post" action="product_types.asp" class="form-container">
                    <%= GetCSRFTokenField() %>
                    <input type="hidden" name="action" value="<%= qsAction %>">
                    <% If qsAction = "edit" Then %>
                    <input type="hidden" name="configId" value="<%= editData(0) %>">
                    <% End If %>
                    
                    <div class="admin-form-group">
                        <label class="admin-form-label">类型代码 <span class="required">*</span></label>
                        <% If qsAction = "add" Then %>
                        <input type="text" name="typeCode" class="admin-form-control" required maxlength="50" placeholder="如：Fixed(品牌定香), Custom(用户定制), KOL">
                        <div class="help-text">唯一标识，添加后不可修改，建议使用英文</div>
                        <% Else %>
                        <input type="text" class="admin-form-control" value="<%= SafeOutput(editData(1)) %>" readonly>
                        <input type="hidden" name="typeCode" value="<%= SafeOutput(editData(1)) %>">
                        <div class="help-text">类型代码不可修改</div>
                        <% End If %>
                    </div>
                    
                    <div class="admin-form-group">
                        <label class="admin-form-label">显示名称 <span class="required">*</span></label>
                        <input type="text" name="displayName" class="admin-form-control" required maxlength="100" 
                               value="<%= IIf(qsAction = "edit", SafeOutput(editData(2)), "") %>" placeholder="如：品牌定香">
                    </div>
                    
                    <div class="admin-form-group">
                        <label class="admin-form-label">栏目名称</label>
                        <input type="text" name="navName" class="admin-form-control" maxlength="100" 
                               value="<%= IIf(qsAction = "edit", SafeOutput(editData(3)), "") %>" placeholder="如：品牌定香（为空则不在前台导航显示）">
                        <div class="help-text">为空则不在前台导航显示</div>
                    </div>
                    
                    <div class="admin-form-group">
                        <label class="admin-form-label">描述</label>
                        <textarea name="description" class="admin-form-control" rows="3" placeholder="类型描述说明"><%= IIf(qsAction = "edit", SafeOutput(editData(4)), "") %></textarea>
                    </div>
                    
                    <div class="admin-form-group">
                        <label class="admin-form-label">图标</label>
                        <input type="text" name="icon" class="admin-form-control" maxlength="100" 
                               value="<%= IIf(qsAction = "edit", SafeOutput(editData(5)), "fas fa-box") %>" placeholder="如：fas fa-gem">
                        <div class="help-text">Font Awesome 图标类名，如：fas fa-gem, fas fa-wind</div>
                    </div>
                    
                    <div class="admin-form-row" style="display: flex; gap: 30px;">
                        <div class="admin-form-group">
                            <div class="checkbox-group">
                                <input type="checkbox" name="requiresReview" id="requiresReview" value="1" 
                                       <%= IIf(qsAction = "edit" AND editData(6) = True, "checked", "") %>>
                                <label for="requiresReview">需要审核</label>
                            </div>
                            <div class="help-text">该类型商品需要运营审核后才能上架</div>
                        </div>
                        
                        <div class="admin-form-group">
                            <div class="checkbox-group">
                                <input type="checkbox" name="requiresRatio" id="requiresRatio" value="1" 
                                       <%= IIf(qsAction = "edit" AND editData(7) = True, "checked", "") %>>
                                <label for="requiresRatio">需要配比</label>
                            </div>
                            <div class="help-text">该类型商品需要设置香调配比</div>
                        </div>
                    </div>
                    
                    <div class="admin-form-group">
                        <label class="admin-form-label">排序号</label>
                        <input type="number" name="displayOrder" class="admin-form-control" min="0" 
                               value="<%= IIf(qsAction = "edit", editData(8), "0") %>">
                        <div class="help-text">数字越小排序越靠前</div>
                    </div>
                    
                    <div class="admin-form-group">
                        <div class="checkbox-group">
                            <input type="checkbox" name="isActive" id="isActive" value="1" 
                                   <%= IIf(qsAction = "edit" AND editData(9) = True, "checked", "checked") %>
                                   <% If qsAction = "edit" Then %>onchange="if(!this.checked) { if(!confirmDisable('<%= SafeOutput(editData(1)) %>')) { this.checked = true; } }"<% End If %>>
                            <label for="isActive">启用</label>
                        </div>
                        <div class="help-text">禁用后该类型不会在前台显示</div>
                    </div>
                    
                    <div class="form-actions">
                        <button type="submit" class="admin-btn admin-btn-primary">
                            <i class="fas fa-save"></i> 保存
                        </button>
                        <a href="product_types.asp" class="admin-btn admin-btn-secondary">取消</a>
                    </div>
                </form>
                <% Else %>
                <div class="empty-state">
                    <i class="fas fa-lock"></i>
                    <h3>权限不足</h3>
                    <p>您没有权限添加或编辑商品类型</p>
                    <a href="product_types.asp" class="admin-btn admin-btn-secondary" style="margin-top: 20px;">返回列表</a>
                </div>
                <% End If %>
            </div>
        </div>
        <% Else %>
        <!-- 列表视图 -->
        <% If typeCount = 0 Then %>
        <div class="empty-state">
            <i class="fas fa-tags"></i>
            <h3>暂无商品类型</h3>
            <p>还没有配置任何商品类型</p>
        </div>
        <% Else %>
        <div class="admin-card">
            <div class="admin-card-body">
                <table class="admin-table">
                    <thead>
                        <tr>
                            <th>排序</th>
                            <th>类型代码</th>
                            <th>显示名称</th>
                            <th>栏目名称</th>
                            <th>图标</th>
                            <th>需要审核</th>
                            <th>需要配比</th>
                            <th>状态</th>
                            <th>商品数量</th>
                            <th>操作</th>
                        </tr>
                    </thead>
                    <tbody>
                        <%
                        Dim typeCodeKey, totalCount, activeCount
                        For idx = 0 To typeCount - 1
                        %>
                        <tr>
                            <td><%= allTypes(idx, 8) %></td>
                            <td><code><%= SafeOutput(allTypes(idx, 1)) %></code></td>
                            <td><strong><%= SafeOutput(allTypes(idx, 2)) %></strong></td>
                            <td>
                                <% If allTypes(idx, 3) <> "" Then %>
                                <%= SafeOutput(allTypes(idx, 3)) %>
                                <% Else %>
                                <span style="color: #999;">-</span>
                                <% End If %>
                            </td>
                            <td>
                                <div class="icon-preview">
                                    <% If allTypes(idx, 5) <> "" Then %>
                                    <i class="<%= SafeOutput(allTypes(idx, 5)) %>"></i>
                                    <% Else %>
                                    <i class="fas fa-box"></i>
                                    <% End If %>
                                </div>
                            </td>
                            <td>
                                <% If allTypes(idx, 6) = True Then %>
                                <span class="status-badge status-active">是</span>
                                <% Else %>
                                <span style="color: #999;">否</span>
                                <% End If %>
                            </td>
                            <td>
                                <% If allTypes(idx, 7) = True Then %>
                                <span class="status-badge status-active">是</span>
                                <% Else %>
                                <span style="color: #999;">否</span>
                                <% End If %>
                            </td>
                            <td>
                                <% If allTypes(idx, 9) = True Then %>
                                <span class="status-badge status-active">启用</span>
                                <% Else %>
                                <span class="status-badge status-inactive">禁用</span>
                                <% End If %>
                            </td>
                            <td>
                                <%
                                typeCodeKey = CStr(allTypes(idx, 1))
                                If productStats.Exists(typeCodeKey) Then
                                    totalCount = productStats(typeCodeKey)(0)
                                    activeCount = productStats(typeCodeKey)(1)
                                Else
                                    totalCount = 0
                                    activeCount = 0
                                End If
                                %>
                                <a href="products.asp?type=<%= Server.URLEncode(allTypes(idx, 1)) %>" class="status-badge" style="background: #e3f2fd; color: #1976d2; text-decoration: none;" title="上架商品数 / 总商品数">
                                    <%= activeCount %> / <%= totalCount %>
                                </a>
                            </td>
                            <td>
                                <div class="action-btns">
                                    <a href="product_types.asp?action=edit&id=<%= allTypes(idx, 0) %>" class="admin-btn admin-btn-sm admin-btn-secondary" title="编辑">
                                        <i class="fas fa-edit"></i>
                                    </a>
                                    <% If canEdit Then %>
                                    <form method="post" style="display: inline;" onsubmit="return confirmDelete('<%= allTypes(idx, 1) %>');">
                                        <%= GetCSRFTokenField() %>
                                        <input type="hidden" name="action" value="delete">
                                        <input type="hidden" name="configId" value="<%= allTypes(idx, 0) %>">
                                        <button type="submit" class="admin-btn admin-btn-sm admin-btn-danger" title="删除">
                                            <i class="fas fa-trash"></i>
                                        </button>
                                    </form>
                                    <% End If %>
                                </div>
                            </td>
                        </tr>
                        <% Next %>
                    </tbody>
                </table>
            </div>
        </div>
        <% End If %>
        <% End If %>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
