<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
' V9.0架构：配方管理已迁移至产品技术管理中心
Response.Redirect "../techcenter/formula_management.asp"
Response.End

Call OpenConnection()

' 获取配方ID（编辑模式）
Dim recipeId, isEditMode
recipeId = Request.QueryString("id")
isEditMode = False

If recipeId <> "" And IsNumeric(recipeId) Then
    isEditMode = True
End If

' 初始化变量
Dim recipeName, description, productId, isActive, sortOrder
recipeName = ""
description = ""
productId = ""
isActive = True
sortOrder = 0

Dim errorMsg, successMsg
errorMsg = ""
successMsg = ""

' 处理POST请求
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' 验证CSRF令牌
    If Not ValidateCSRFToken() Then
        errorMsg = "安全验证失败，请刷新页面重试"
    Else
        ' 获取表单数据
        recipeName = Trim(Request.Form("recipe_name"))
        description = Trim(Request.Form("description"))
        productId = Request.Form("product_id")
        isActive = (Request.Form("is_active") = "1")
        sortOrder = Request.Form("sort_order")
        
        ' 验证必填字段
        If recipeName = "" Then
            errorMsg = "请输入配方名称"
        ElseIf productId = "" Then
            errorMsg = "请选择关联产品"
        Else
            ' 安全处理SQL字符串
            Dim safeRecipeName, safeDescription
            safeRecipeName = SafeSQL(recipeName)
            safeDescription = SafeSQL(description)
            
            Dim sql, result
            
            If isEditMode Then
                ' 更新现有配方
                sql = "UPDATE RecommendedRecipes SET " & _
                      "RecipeName = '" & safeRecipeName & "', " & _
                      "Description = '" & safeDescription & "', " & _
                      "ProductID = " & productId & ", " & _
                      "IsActive = " & IIf(isActive, "True", "False") & ", " & _
                      "SortOrder = " & IIf(sortOrder <> "", sortOrder, "0") & " " & _
                      "WHERE RecipeID = " & recipeId
                
                result = ExecuteNonQuery(sql)
                
                If result Then
                    Call LogAdminAction("编辑配方", "operation", "RecommendedRecipes", recipeId, safeRecipeName)
                    Response.Redirect "recipes.asp?msg=updated"
                    Response.End
                Else
                    errorMsg = "更新失败: " & Session("LastDBError")
                End If
            Else
                ' 创建新配方
                sql = "INSERT INTO RecommendedRecipes (RecipeName, Description, ProductID, IsActive, SortOrder) " & _
                      "VALUES ('" & safeRecipeName & "', '" & safeDescription & "', " & productId & ", " & _
                      IIf(isActive, "True", "False") & ", " & IIf(sortOrder <> "", sortOrder, "0") & ")"
                
                result = ExecuteNonQuery(sql)
                
                If result Then
                    Dim newId
                    newId = GetLastInsertID("RecommendedRecipes")
                    Call LogAdminAction("创建配方", "operation", "RecommendedRecipes", newId, safeRecipeName)
                    Response.Redirect "recipes.asp?msg=created"
                    Response.End
                Else
                    errorMsg = "创建失败: " & Session("LastDBError")
                End If
            End If
        End If
    End If
ElseIf isEditMode Then
    ' 编辑模式：加载现有数据
    Dim rsRecipe
    Set rsRecipe = ExecuteQuery("SELECT * FROM RecommendedRecipes WHERE RecipeID = " & recipeId)
    
    If Not rsRecipe Is Nothing And Not rsRecipe.EOF Then
        recipeName = rsRecipe("RecipeName")
        description = rsRecipe("Description")
        productId = rsRecipe("ProductID")
        isActive = rsRecipe("IsActive")
        If Not IsNull(rsRecipe("SortOrder")) Then
            sortOrder = rsRecipe("SortOrder")
        End If
        rsRecipe.Close
    Else
        errorMsg = "配方不存在"
        isEditMode = False
    End If
    Set rsRecipe = Nothing
End If

' 获取产品列表用于下拉选择
Dim rsProducts
Set rsProducts = ExecuteQuery("SELECT ProductID, ProductName FROM Products WHERE IsActive = 1 ORDER BY ProductName")

Call LogAdminAction(IIf(isEditMode, "编辑配方页面", "创建配方页面"), "operation", "RecommendedRecipes", recipeId, "")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title><%= IIf(isEditMode, "编辑", "创建") %>配方 - 运营管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        .form-container { max-width: 800px; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.08); }
        .form-group { margin-bottom: 20px; }
        .form-group label { display: block; margin-bottom: 8px; font-weight: 500; color: #333; }
        .form-group label .required { color: #e74c3c; margin-left: 4px; }
        .form-control { width: 100%; padding: 12px 15px; border: 1px solid #ddd; border-radius: 8px; font-size: 14px; box-sizing: border-box; }
        .form-control:focus { outline: none; border-color: #667eea; }
        textarea.form-control { min-height: 100px; resize: vertical; }
        .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        .form-actions { display: flex; gap: 15px; margin-top: 30px; padding-top: 20px; border-top: 1px solid #f0f0f0; }
        .alert { padding: 15px; border-radius: 8px; margin-bottom: 20px; }
        .alert-error { background: #ffebee; color: #c62828; border: 1px solid #ffcdd2; }
        .alert-success { background: #e8f5e9; color: #2e7d32; border: 1px solid #c8e6c9; }
        .checkbox-group { display: flex; align-items: center; gap: 10px; }
        .checkbox-group input[type="checkbox"] { width: 20px; height: 20px; }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-fire"></i> <%= IIf(isEditMode, "编辑", "创建") %>配方</h2>
            <div class="breadcrumb">
                <a href="index.asp">运营中心</a> / <a href="recipes.asp">配方推荐</a> / <span><%= IIf(isEditMode, "编辑配方", "创建配方") %></span>
            </div>
        </div>
        
        <% If errorMsg <> "" Then %>
        <div class="alert alert-error"><i class="fas fa-exclamation-circle"></i> <%= errorMsg %></div>
        <% End If %>
        
        <div class="form-container">
            <form method="post" action="">
                <%= GetCSRFTokenField() %>
                
                <div class="form-group">
                    <label>配方名称 <span class="required">*</span></label>
                    <input type="text" name="recipe_name" class="form-control" value="<%= SafeOutput(recipeName) %>" required>
                </div>
                
                <div class="form-group">
                    <label>配方描述</label>
                    <textarea name="description" class="form-control" placeholder="请输入配方描述（可选）"><%= SafeOutput(description) %></textarea>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>关联产品 <span class="required">*</span></label>
                        <select name="product_id" class="form-control" required>
                            <option value="">请选择产品</option>
                            <% If Not rsProducts Is Nothing Then %>
                            <% Do While Not rsProducts.EOF %>
                            <option value="<%= rsProducts("ProductID") %>" <%= IIf(CStr(productId)=CStr(rsProducts("ProductID")), "selected", "") %>>
                                <%= SafeOutput(rsProducts("ProductName")) %>
                            </option>
                            <% rsProducts.MoveNext %>
                            <% Loop %>
                            <% rsProducts.Close %>
                            <% End If %>
                        </select>
                    </div>
                    
                    <div class="form-group">
                        <label>排序号</label>
                        <input type="number" name="sort_order" class="form-control" value="<%= sortOrder %>" placeholder="数字越小排序越靠前">
                    </div>
                </div>
                
                <div class="form-group">
                    <label>状态</label>
                    <div class="checkbox-group">
                        <input type="checkbox" name="is_active" value="1" <%= IIf(isActive, "checked", "") %>>
                        <span>启用配方</span>
                    </div>
                </div>
                
                <div class="form-actions">
                    <button type="submit" class="admin-btn admin-btn-primary">
                        <i class="fas fa-save"></i> <%= IIf(isEditMode, "保存修改", "创建配方") %>
                    </button>
                    <a href="recipes.asp" class="admin-btn admin-btn-secondary">
                        <i class="fas fa-times"></i> 取消
                    </a>
                </div>
            </form>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
