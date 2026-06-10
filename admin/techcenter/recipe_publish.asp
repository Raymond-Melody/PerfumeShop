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
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Function SafeSQL(str)
    If IsNull(str) Or str = "" Then SafeSQL = "" Else SafeSQL = Replace(str, "'", "''")
End Function

Function IIF(cond, tVal, fVal)
    If cond Then IIF = tVal Else IIF = fVal
End Function

Dim action, msg, msgType, recipeId
action = Trim(Request.Form("action"))
msg = Trim(Request.QueryString("msg"))
msgType = "success"
If InStr(msg, "失败") > 0 Or InStr(msg, "错误") > 0 Then msgType = "error"
recipeId = SafeNum(Request.QueryString("recipe_id"))

' ========== POST 处理 ==========
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If action = "publish_accord" Then
        ' 发布香调配方（原材料→香调）
        Dim pubRecipeID, pubNoteID, pubBatchSize, pubMaterialCount
        pubRecipeID = SafeNum(Request.Form("recipe_id"))
        pubNoteID = SafeNum(Request.Form("note_id"))
        pubBatchSize = SafeNum(Request.Form("batch_size"))
        If pubBatchSize <= 0 Then pubBatchSize = 100
        pubMaterialCount = SafeNum(Request.Form("material_count"))
        
        If pubRecipeID > 0 And pubNoteID > 0 And pubMaterialCount > 0 Then
            ' 数据完整性校验：检查RecipeIngredients是否有数据
            Dim checkIngCount
            checkIngCount = SafeNum(GetScalar("SELECT COUNT(*) FROM RecipeIngredients WHERE RecipeID=" & pubRecipeID & " AND NoteID=" & pubNoteID))
            If checkIngCount = 0 Then
                msg = "发布失败：该配方+香调在RecipeIngredients中无数据，请先在配方管理中生成成分数据"
                msgType = "error"
            Else
            On Error Resume Next
            Err.Clear
            Call BeginTransaction()
            
            ' 获取配方和香调名称
            Dim pubRecipeName, pubNoteName
            pubRecipeName = ""
            pubNoteName = ""
            Dim rsInfo
            Set rsInfo = conn.Execute("SELECT RecipeName FROM Recipes WHERE RecipeID=" & pubRecipeID)
            If Not rsInfo Is Nothing Then
                If Not rsInfo.EOF Then pubRecipeName = CStr(rsInfo("RecipeName") & "")
                rsInfo.Close
            End If
            Set rsInfo = Nothing
            
            Set rsInfo = conn.Execute("SELECT NoteName FROM FragranceNotes WHERE NoteID=" & pubNoteID)
            If Not rsInfo Is Nothing Then
                If Not rsInfo.EOF Then pubNoteName = CStr(rsInfo("NoteName") & "")
                rsInfo.Close
            End If
            Set rsInfo = Nothing
            
            ' 创建 RecipeAccords 记录
            Dim newAccordID
            conn.Execute "INSERT INTO RecipeAccords (RecipeID, NoteID, BatchSize, Status, PublishedBy, PublishedAt, CreatedAt, RecipeName) VALUES (" & _
                pubRecipeID & ", " & pubNoteID & ", " & pubBatchSize & ", 'Published', '" & SafeSQL(Session("AdminUsername")) & "', GETDATE(), GETDATE(), '" & SafeSQL(pubRecipeName & "→" & pubNoteName) & "')"
            
            If Err.Number = 0 Then
                Set rsInfo = conn.Execute("SELECT @@Identity")
                newAccordID = 0
                If Not rsInfo Is Nothing Then
                    If Not rsInfo.EOF Then newAccordID = CLng(rsInfo(0))
                    rsInfo.Close
                End If
                Set rsInfo = Nothing
                
                If newAccordID > 0 Then
                    Dim anyAccordError, mi
                    anyAccordError = False
                    
                    For mi = 1 To pubMaterialCount
                        Dim matID, matName, matPct, matQty
                        matID = SafeNum(Request.Form("material_id_" & mi))
                        matName = Trim(Request.Form("material_name_" & mi))
                        matPct = SafeNum(Request.Form("material_pct_" & mi))
                        matQty = SafeNum(Request.Form("material_qty_" & mi))
                        
                        If matID >= 0 And matName <> "" And matPct > 0 Then
                            conn.Execute "INSERT INTO RecipeAccordMaterials (AccordRecipeID, MaterialID, MaterialName, Percentage, PlannedQty) VALUES (" & _
                                newAccordID & ", " & matID & ", '" & SafeSQL(matName) & "', " & matPct & ", " & matQty & ")"
                            If Err.Number <> 0 Then
                                anyAccordError = True
                                Err.Clear
                            End If
                        End If
                    Next
                    
                    If Not anyAccordError Then
                        ' 审计日志
                        conn.Execute "INSERT INTO RecipePublishLog (RecipeID, PublishType, TargetRecipeID, PublishedBy, PublishedAt, IPAddress) VALUES (" & _
                            pubRecipeID & ", 'Accord', " & newAccordID & ", '" & SafeSQL(Session("AdminUsername")) & "', GETDATE(), '" & SafeSQL(Request.ServerVariables("REMOTE_ADDR")) & "')"
                        
                        ' 标记旧版本为 Deprecated（同一 RecipeID+NoteID 的旧发布版本）
                        conn.Execute "UPDATE RecipeAccords SET Status='Deprecated' WHERE AccordRecipeID<>" & newAccordID & " AND RecipeID=" & pubRecipeID & " AND NoteID=" & pubNoteID & " AND Status='Published'"
                        
                        Call CommitTransaction()
                        Response.Redirect "recipe_publish.asp?msg=香调配方发布成功！" & pubRecipeName & "→" & pubNoteName
                        Response.End
                    Else
                        Call RollbackTransaction()
                        msg = "香调配方发布失败：原材料明细写入出错，数据已回滚"
                        msgType = "error"
                    End If
                Else
                    Call RollbackTransaction()
                    msg = "香调配方发布失败：无法获取发布ID"
                    msgType = "error"
                End If
            Else
                Call RollbackTransaction()
                msg = "香调配方发布失败: " & Err.Description
                msgType = "error"
                Err.Clear
            End If
            On Error GoTo 0
        End If ' End checkIngCount
        Else
            msg = "请选择配方、香调和原材料明细"
            msgType = "error"
        End If
    
    ElseIf action = "publish_product" Then
        ' 发布产品配方（香调→产品）
        Dim ppRecipeID, ppProductID, ppBatchSize, ppNoteCount
        ppRecipeID = SafeNum(Request.Form("recipe_id"))
        ppProductID = SafeNum(Request.Form("product_id"))
        ppBatchSize = SafeNum(Request.Form("batch_size"))
        If ppBatchSize <= 0 Then ppBatchSize = 100
        ppNoteCount = SafeNum(Request.Form("note_count"))
        
        If ppRecipeID > 0 And ppNoteCount > 0 Then
            ' 数据完整性校验：检查RecipeNotes百分比总和是否接近100%
            Dim checkPctSum
            checkPctSum = SafeNum(GetScalar("SELECT SUM(Percentage) FROM RecipeNotes WHERE RecipeID=" & ppRecipeID))
            If Abs(checkPctSum - 100) > 1 Then
                ' 发出警告但不阻断
                msg = "警告：配方香调百分比总和为" & FormatNumber(checkPctSum,1) & "%，偏离100%超过1%。发布已继续。"
                msgType = "error"
            End If
            On Error Resume Next
            Err.Clear
            Call BeginTransaction()
            
            Dim ppRecipeName, ppProductName
            ppRecipeName = ""
            ppProductName = ""
            Set rsInfo = conn.Execute("SELECT RecipeName FROM Recipes WHERE RecipeID=" & ppRecipeID)
            If Not rsInfo Is Nothing Then
                If Not rsInfo.EOF Then ppRecipeName = CStr(rsInfo("RecipeName") & "")
                rsInfo.Close
            End If
            Set rsInfo = Nothing
            
            If ppProductID > 0 Then
                Set rsInfo = conn.Execute("SELECT ProductName FROM Products WHERE ProductID=" & ppProductID)
                If Not rsInfo Is Nothing Then
                    If Not rsInfo.EOF Then ppProductName = CStr(rsInfo("ProductName") & "")
                    rsInfo.Close
                End If
                Set rsInfo = Nothing
            End If
            
            ' 创建 RecipeProducts 记录
            Dim newProductRID
            conn.Execute "INSERT INTO RecipeProducts (RecipeID, ProductID, BatchSize, Status, PublishedBy, PublishedAt, CreatedAt) VALUES (" & _
                ppRecipeID & ", " & ppProductID & ", " & ppBatchSize & ", 'Published', '" & SafeSQL(Session("AdminUsername")) & "', GETDATE(), GETDATE())"
            
            If Err.Number = 0 Then
                Set rsInfo = conn.Execute("SELECT @@Identity")
                newProductRID = 0
                If Not rsInfo Is Nothing Then
                    If Not rsInfo.EOF Then newProductRID = CLng(rsInfo(0))
                    rsInfo.Close
                End If
                Set rsInfo = Nothing
                
                If newProductRID > 0 Then
                    Dim anyProdError, ni
                    anyProdError = False
                    
                    For ni = 1 To ppNoteCount
                        Dim pnNoteID, pnNoteName, pnPct, pnQty
                        pnNoteID = SafeNum(Request.Form("pnote_id_" & ni))
                        pnNoteName = Trim(Request.Form("pnote_name_" & ni))
                        pnPct = SafeNum(Request.Form("pnote_pct_" & ni))
                        pnQty = SafeNum(Request.Form("pnote_qty_" & ni))
                        
                        If pnNoteID > 0 And pnPct > 0 Then
                            conn.Execute "INSERT INTO RecipeProductNotes (ProductRecipeID, NoteID, NoteName, Percentage, PlannedQty) VALUES (" & _
                                newProductRID & ", " & pnNoteID & ", '" & SafeSQL(pnNoteName) & "', " & pnPct & ", " & pnQty & ")"
                            If Err.Number <> 0 Then
                                anyProdError = True
                                Err.Clear
                            End If
                        End If
                    Next
                    
                    If Not anyProdError Then
                        conn.Execute "INSERT INTO RecipePublishLog (RecipeID, PublishType, TargetRecipeID, PublishedBy, PublishedAt, IPAddress) VALUES (" & _
                            ppRecipeID & ", 'Product', " & newProductRID & ", '" & SafeSQL(Session("AdminUsername")) & "', GETDATE(), '" & SafeSQL(Request.ServerVariables("REMOTE_ADDR")) & "')"
                        
                        conn.Execute "UPDATE RecipeProducts SET Status='Deprecated' WHERE ProductRecipeID<>" & newProductRID & " AND RecipeID=" & ppRecipeID & " AND Status='Published'"
                        
                        Call CommitTransaction()
                        Response.Redirect "recipe_publish.asp?msg=产品配方发布成功！" & ppRecipeName
                        Response.End
                    Else
                        Call RollbackTransaction()
                        msg = "产品配方发布失败：香调明细写入出错，数据已回滚"
                        msgType = "error"
                    End If
                Else
                    Call RollbackTransaction()
                    msg = "产品配方发布失败：无法获取发布ID"
                    msgType = "error"
                End If
            Else
                Call RollbackTransaction()
                msg = "产品配方发布失败: " & Err.Description
                msgType = "error"
                Err.Clear
            End If
            On Error GoTo 0
        Else
            msg = "请选择配方和香调明细"
            msgType = "error"
        End If
    
    ElseIf action = "deprecate_accord" Then
        Dim depID
        depID = SafeNum(Request.Form("accord_recipe_id"))
        If depID > 0 Then
            conn.Execute "UPDATE RecipeAccords SET Status='Deprecated' WHERE AccordRecipeID=" & depID
            conn.Execute "INSERT INTO RecipePublishLog (RecipeID, PublishType, TargetRecipeID, PublishedBy, PublishedAt, IPAddress) VALUES (0, 'Accord', " & depID & ", '" & SafeSQL(Session("AdminUsername")) & "', GETDATE(), '" & SafeSQL(Request.ServerVariables("REMOTE_ADDR")) & "')"
            Response.Redirect "recipe_publish.asp?msg=香调配方已废弃"
            Response.End
        End If
    
    ElseIf action = "deprecate_product" Then
        Dim depPID
        depPID = SafeNum(Request.Form("product_recipe_id"))
        If depPID > 0 Then
            conn.Execute "UPDATE RecipeProducts SET Status='Deprecated' WHERE ProductRecipeID=" & depPID
            conn.Execute "INSERT INTO RecipePublishLog (RecipeID, PublishType, TargetRecipeID, PublishedBy, PublishedAt, IPAddress) VALUES (0, 'Product', " & depPID & ", '" & SafeSQL(Session("AdminUsername")) & "', GETDATE(), '" & SafeSQL(Request.ServerVariables("REMOTE_ADDR")) & "')"
            Response.Redirect "recipe_publish.asp?msg=产品配方已废弃"
            Response.End
        End If
    End If
End If

' ========== 统计数据 ==========
Dim statTotalRecipes, statAccordPublished, statProductPublished, statPublishLogs
statTotalRecipes = 0 : statAccordPublished = 0 : statProductPublished = 0 : statPublishLogs = 0
On Error Resume Next
statTotalRecipes = SafeNum(GetScalar("SELECT COUNT(*) FROM Recipes WHERE IsActive=1"))
statAccordPublished = SafeNum(GetScalar("SELECT COUNT(*) FROM RecipeAccords WHERE Status='Published'"))
statProductPublished = SafeNum(GetScalar("SELECT COUNT(*) FROM RecipeProducts WHERE Status='Published'"))
statPublishLogs = SafeNum(GetScalar("SELECT COUNT(*) FROM RecipePublishLog"))
On Error GoTo 0
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>配方拆分发布 - 产品技术中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <style>
        /* 页面特有补充样式 - 使用 design-tokens 变量 */
        h2 { font-size: 18px; margin: 20px 0 10px; color: var(--dt-color-primary); border-bottom: 1px solid var(--dt-color-border); padding-bottom: 8px; }
        h3 { font-size: 15px; color: #ccc; margin: 12px 0 8px; }
        .message { padding: 12px 20px; border-radius: 6px; margin-bottom: 16px; font-weight: 500; }
        .message.success { background: rgba(76,175,80,0.15); color: var(--dt-color-success); border: 1px solid rgba(76,175,80,0.3); }
        .message.error { background: rgba(244,67,54,0.15); color: var(--dt-color-danger); border: 1px solid rgba(244,67,54,0.3); }
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin-bottom: 24px; }
        .stats-card { background: linear-gradient(135deg, var(--dt-color-bg-light) 0%, #1e1e32 100%); border-radius: 12px; padding: 25px; box-shadow: 0 4px 20px rgba(0,0,0,0.3); border: 1px solid rgba(255,255,255,0.05); text-align: center; transition: transform 0.2s ease; }
        .stats-card:hover { transform: translateY(-3px); }
        .stats-card .num { font-size: 28px; font-weight: 700; color: var(--dt-color-primary); display: block; }
        .stats-card .label { font-size: 13px; color: #888; margin-top: 6px; }
        .card { background: var(--dt-color-bg-light); border-radius: 8px; border: 1px solid var(--dt-color-border); margin-bottom: 20px; overflow: hidden; }
        .card-header { padding: 14px 20px; background: rgba(0,188,212,0.08); border-bottom: 1px solid var(--dt-color-border); font-weight: 600; font-size: 15px; display: flex; justify-content: space-between; align-items: center; }
        .card-body { padding: 16px 20px; }
        table { width: 100%; border-collapse: collapse; }
        th { text-align: left; padding: 10px 12px; background: rgba(0,188,212,0.06); border-bottom: 1px solid var(--dt-color-border); font-size: 13px; color: #888; font-weight: 600; }
        td { padding: 10px 12px; border-bottom: 1px solid var(--dt-color-border); font-size: 14px; color: var(--dt-color-text); }
        tr:hover { background: rgba(255,255,255,0.02); }
        input[type="text"], input[type="number"], textarea, select { width: 100%; padding: 9px 12px; background: var(--dt-color-bg-light); border: 1px solid var(--dt-color-border); border-radius: 5px; color: var(--dt-color-text); font-size: 14px; }
        input:focus, select:focus { border-color: var(--dt-color-primary); outline: none; }
        .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
        .grid-3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 12px; }
        .text-right { text-align: right; }
        .text-center { text-align: center; }
        .text-muted { color: #999; font-size: 13px; }
        .mb-2 { margin-bottom: 16px; }
        .mt-2 { margin-top: 16px; }
        .status-badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
        .status-published { background: rgba(76,175,80,0.2); color: var(--dt-color-success); }
        .status-draft { background: rgba(158,158,158,0.2); color: #9E9E9E; }
        .status-deprecated { background: rgba(244,67,54,0.2); color: var(--dt-color-danger); }
        .security-badge { display: inline-block; padding: 2px 8px; background: rgba(255,152,0,0.15); color: var(--dt-color-warning); border-radius: 4px; font-size: 11px; margin-left: 6px; }
        .split-panel { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        .accord-panel { border-left: 3px solid #FF9800; }
        .product-panel { border-left: 3px solid #2196F3; }
        .material-row, .note-row { display: flex; gap: 8px; align-items: center; padding: 4px 0; flex-wrap: wrap; }
        .material-row input, .note-row input { width: auto; min-width: 80px; }
        .material-row .mat-name { flex: 1; min-width: 120px; }
        .modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.7); z-index: 1000; justify-content: center; align-items: center; }
        .modal-overlay.active { display: flex; }
        .modal-content { background: var(--dt-color-bg-light); border-radius: 10px; padding: 24px; max-width: 800px; width: 90%; max-height: 80vh; overflow-y: auto; border: 1px solid var(--dt-color-border); }
        .modal-content h2 { margin-top: 0; border-bottom: 1px solid var(--dt-color-border); padding-bottom: 10px; }
        .close-btn { float: right; background: none; border: none; color: #999; font-size: 24px; cursor: pointer; }
        .close-btn:hover { color: #fff; }
        .ingredient-list { max-height: 300px; overflow-y: auto; margin: 10px 0; }
        .ingredient-item { display: flex; align-items: center; gap: 8px; padding: 6px 10px; border-radius: 4px; margin: 2px 0; }
        .ingredient-item:hover { background: rgba(255,255,255,0.03); }
        .ingredient-item label { flex: 1; cursor: pointer; }
        .ingredient-item input[type="checkbox"] { width: 16px; height: 16px; accent-color: var(--dt-color-primary); }
        .security-note { background: rgba(255,152,0,0.08); border: 1px solid rgba(255,152,0,0.3); border-radius: 6px; padding: 12px; margin-top: 16px; font-size: 13px; }
        .security-note i { color: #FF9800; }
        @media (max-width: 1200px) { .stats-grid { grid-template-columns: repeat(2, 1fr); } }
        @media (max-width: 768px) { .stats-grid { grid-template-columns: 1fr; } }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-lock"></i> 配方拆分发布</h2>
            <div class="breadcrumb">
                <a href="index.asp">技术中心</a> / <span>配方拆分发布</span>
            </div>
        </div>
        
        <% If msg <> "" Then %>
        <div class="message <%=msgType%>"><%=Server.HTMLEncode(msg)%></div>
        <% End If %>
        
        <div class="stats-grid">
            <div class="stats-card"><span class="num"><%=statTotalRecipes%></span><span class="label">活跃配方</span></div>
            <div class="stats-card"><span class="num" style="color:#FF9800;"><%=statAccordPublished%></span><span class="label">已发布香调配方</span></div>
            <div class="stats-card"><span class="num" style="color:#2196F3;"><%=statProductPublished%></span><span class="label">已发布产品配方</span></div>
            <div class="stats-card"><span class="num"><%=statPublishLogs%></span><span class="label">审计日志</span></div>
        </div>
        
        <div class="security-note">
            <i class="fas fa-shield-alt"></i> <strong>安全隔离说明：</strong>
            配方拆分发布后，香调生产车间仅能看到「原材料→香调」片段（RecipeAccords），
            产品制造车间仅能看到「香调→产品」片段（RecipeProducts）。
            技术中心是唯一持有完整配方的角色。所有发布操作均记录审计日志。
        </div>
        
        <!-- 配方列表 -->
        <div class="card">
            <div class="card-header">活跃配方列表 — 选择配方进行拆分发布</div>
            <div class="card-body" style="overflow-x:auto;">
                <table>
                    <thead>
                        <tr><th>配方编号</th><th>配方名称</th><th>类型</th><th>香调数</th><th>香调配方</th><th>产品配方</th><th>操作</th></tr>
                    </thead>
                    <tbody>
                    <%
                    Dim rsRecipes
                    Set rsRecipes = conn.Execute("SELECT RecipeID, RecipeName, RecipeCode, ProductType FROM Recipes WHERE IsActive=1 ORDER BY RecipeCode")
                    If Not rsRecipes Is Nothing Then
                        Do While Not rsRecipes.EOF
                            Dim rID, rName, rCode, rType
                            rID = rsRecipes("RecipeID")
                            rName = CStr(rsRecipes("RecipeName") & "")
                            rCode = CStr(rsRecipes("RecipeCode") & "")
                            rType = CStr(rsRecipes("ProductType") & "")
                            
                            ' 统计香调数量
                            Dim noteCount
                            noteCount = SafeNum(GetScalar("SELECT COUNT(*) FROM RecipeNotes WHERE RecipeID=" & rID))
                            
                            ' 检查是否已有已发布的香调/产品配方
                            Dim hasAccord, hasProduct
                            hasAccord = SafeNum(GetScalar("SELECT COUNT(*) FROM RecipeAccords WHERE RecipeID=" & rID & " AND Status='Published'"))
                            hasProduct = SafeNum(GetScalar("SELECT COUNT(*) FROM RecipeProducts WHERE RecipeID=" & rID & " AND Status='Published'"))
                    %>
                        <tr>
                            <td><strong><%=rCode%></strong></td>
                            <td><%=Server.HTMLEncode(rName)%></td>
                            <td><%=rType%></td>
                            <td><%=noteCount%></td>
                            <td>
                                <% If hasAccord > 0 Then %>
                                    <span class="status-badge status-published">已发布</span>
                                <% Else %>
                                    <span class="status-badge status-draft">未发布</span>
                                <% End If %>
                            </td>
                            <td>
                                <% If hasProduct > 0 Then %>
                                    <span class="status-badge status-published">已发布</span>
                                <% Else %>
                                    <span class="status-badge status-draft">未发布</span>
                                <% End If %>
                            </td>
                            <td>
                                <button class="btn btn-outline btn-sm" onclick="openPreviewModal(<%=rID%>,'<%=Server.HTMLEncode(rName)%>')" style="margin-right:4px;" title="预览配方"><i class="fas fa-eye"></i></button>
                                <button class="btn btn-primary btn-sm" onclick="openAccordModal(<%=rID%>,'<%=Server.HTMLEncode(rName)%>')">发布香调配方</button>
                                <button class="btn btn-outline btn-sm" onclick="openProductModal(<%=rID%>,'<%=Server.HTMLEncode(rName)%>')" style="margin-left:4px;">发布产品配方</button>
                            </td>
                        </tr>
                    <%
                            rsRecipes.MoveNext
                        Loop
                        rsRecipes.Close
                    End If
                    Set rsRecipes = Nothing
                    %>
                    </tbody>
                </table>
            </div>
        </div>
        
        <!-- 已发布的香调配方列表 -->
        <div class="card">
            <div class="card-header" style="border-left: 3px solid #FF9800;">已发布香调配方（RecipeAccords）— 香调生产车间可见</div>
            <div class="card-body" style="overflow-x:auto;">
                <table>
                    <thead>
                        <tr><th>配方名称</th><th>产出香调</th><th>批量</th><th>状态</th><th>发布时间</th><th>发布人</th><th>操作</th></tr>
                    </thead>
                    <tbody>
                    <%
                    Dim rsAccords
                    Set rsAccords = conn.Execute("SELECT ra.AccordRecipeID, ra.BatchSize, ra.Status, ra.PublishedBy, ra.PublishedAt, ra.RecipeName, r.RecipeName AS FullRecipeName, fn.NoteName FROM RecipeAccords ra LEFT JOIN Recipes r ON ra.RecipeID=r.RecipeID LEFT JOIN FragranceNotes fn ON ra.NoteID=fn.NoteID ORDER BY ra.PublishedAt DESC")
                    If Not rsAccords Is Nothing Then
                        Dim accordRowCount : accordRowCount = 0
                        Do While Not rsAccords.EOF
                            accordRowCount = accordRowCount + 1
                            Dim aStatus : aStatus = CStr(rsAccords("Status") & "")
                            Dim aStatusClass
                            If aStatus = "Published" Then
                                aStatusClass = "status-published"
                            ElseIf aStatus = "Draft" Then
                                aStatusClass = "status-draft"
                            Else
                                aStatusClass = "status-deprecated"
                            End If
                    %>
                        <tr>
                            <td><strong><%=rsAccords("RecipeName") & ""%></strong></td>
                            <td><%=rsAccords("NoteName") & ""%></td>
                            <td><%=rsAccords("BatchSize")%></td>
                            <td><span class="status-badge <%=aStatusClass%>"><%=aStatus%></span></td>
                            <td><%=rsAccords("PublishedAt") & ""%></td>
                            <td><%=rsAccords("PublishedBy") & ""%></td>
                            <td>
                                <% If aStatus = "Published" Then %>
                                <form method="post" style="display:inline;" onsubmit="return confirm('确定废弃此香调配方？废弃后香调车间将不可见。')">
                                    <input type="hidden" name="action" value="deprecate_accord">
                                    <input type="hidden" name="accord_recipe_id" value="<%=rsAccords("AccordRecipeID")%>">
                                    <button class="btn btn-outline-danger btn-xs">废弃</button>
                                </form>
                                <% End If %>
                            </td>
                        </tr>
                    <%
                            rsAccords.MoveNext
                        Loop
                        rsAccords.Close
                        If accordRowCount = 0 Then
                    %>
                        <tr><td colspan="7" class="text-center text-muted" style="padding:30px;">暂无已发布的香调配方</td></tr>
                    <% End If
                    End If
                    Set rsAccords = Nothing %>
                    </tbody>
                </table>
            </div>
        </div>
        
        <!-- 已发布的产品配方列表 -->
        <div class="card">
            <div class="card-header" style="border-left: 3px solid #2196F3;">已发布产品配方（RecipeProducts）— 产品制造车间可见</div>
            <div class="card-body" style="overflow-x:auto;">
                <table>
                    <thead>
                        <tr><th>配方</th><th>产品</th><th>批量</th><th>状态</th><th>发布时间</th><th>发布人</th><th>操作</th></tr>
                    </thead>
                    <tbody>
                    <%
                    Dim rsProducts
                    Set rsProducts = conn.Execute("SELECT rp.ProductRecipeID, rp.BatchSize, rp.Status, rp.PublishedBy, rp.PublishedAt, r.RecipeName, p.ProductName FROM RecipeProducts rp LEFT JOIN Recipes r ON rp.RecipeID=r.RecipeID LEFT JOIN Products p ON rp.ProductID=p.ProductID ORDER BY rp.PublishedAt DESC")
                    If Not rsProducts Is Nothing Then
                        Dim prodRowCount : prodRowCount = 0
                        Do While Not rsProducts.EOF
                            prodRowCount = prodRowCount + 1
                            Dim pStatus : pStatus = CStr(rsProducts("Status") & "")
                            Dim pStatusClass
                            If pStatus = "Published" Then
                                pStatusClass = "status-published"
                            ElseIf pStatus = "Draft" Then
                                pStatusClass = "status-draft"
                            Else
                                pStatusClass = "status-deprecated"
                            End If
                    %>
                        <tr>
                            <td><strong><%=rsProducts("RecipeName") & ""%></strong></td>
                            <td><%=rsProducts("ProductName") & ""%></td>
                            <td><%=rsProducts("BatchSize")%></td>
                            <td><span class="status-badge <%=pStatusClass%>"><%=pStatus%></span></td>
                            <td><%=rsProducts("PublishedAt") & ""%></td>
                            <td><%=rsProducts("PublishedBy") & ""%></td>
                            <td>
                                <% If pStatus = "Published" Then %>
                                <form method="post" style="display:inline;" onsubmit="return confirm('确定废弃此产品配方？废弃后制造车间将不可见。')">
                                    <input type="hidden" name="action" value="deprecate_product">
                                    <input type="hidden" name="product_recipe_id" value="<%=rsProducts("ProductRecipeID")%>">
                                    <button class="btn btn-outline-danger btn-xs">废弃</button>
                                </form>
                                <% End If %>
                            </td>
                        </tr>
                    <%
                            rsProducts.MoveNext
                        Loop
                        rsProducts.Close
                        If prodRowCount = 0 Then
                    %>
                        <tr><td colspan="7" class="text-center text-muted" style="padding:30px;">暂无已发布的产品配方</td></tr>
                    <% End If
                    End If
                    Set rsProducts = Nothing %>
                    </tbody>
                </table>
            </div>
        </div>
        
        <!-- 审计日志 -->
        <div class="card">
            <div class="card-header">配方发布审计日志（RecipePublishLog）</div>
            <div class="card-body" style="overflow-x:auto;">
                <table>
                    <thead>
                        <tr><th>时间</th><th>类型</th><th>目标配方ID</th><th>发布人</th><th>IP地址</th></tr>
                    </thead>
                    <tbody>
                    <%
                    Dim rsLogs
                    Set rsLogs = conn.Execute("SELECT TOP 20 PublishedAt, PublishType, TargetRecipeID, PublishedBy, IPAddress FROM RecipePublishLog ORDER BY PublishedAt DESC")
                    If Not rsLogs Is Nothing Then
                        Dim logCount : logCount = 0
                        Do While Not rsLogs.EOF
                            logCount = logCount + 1
                            Dim logType : logType = CStr(rsLogs("PublishType"))
                    %>
                        <tr>
                            <td><%=rsLogs("PublishedAt") & ""%></td>
                            <td><span class="status-badge <%=IIF(logType="Accord","status-published","status-published")%>" style="<%=IIF(logType="Accord","background:rgba(255,152,0,0.2);color:#FF9800;","background:rgba(33,150,243,0.2);color:#2196F3;")%>"><%=logType%></span></td>
                            <td><%=rsLogs("TargetRecipeID")%></td>
                            <td><%=rsLogs("PublishedBy") & ""%></td>
                            <td><%=rsLogs("IPAddress") & ""%></td>
                        </tr>
                    <%
                            rsLogs.MoveNext
                        Loop
                        rsLogs.Close
                        If logCount = 0 Then
                    %>
                        <tr><td colspan="5" class="text-center text-muted" style="padding:30px;">暂无审计记录</td></tr>
                    <% End If
                    End If
                    Set rsLogs = Nothing %>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    
    <!-- 配方预览弹窗 -->
    <div id="previewModal" class="modal-overlay">
        <div class="modal-content">
            <button class="close-btn" onclick="closePreviewModal()">&times;</button>
            <h2 style="color:#00bcd4;"><i class="fas fa-eye" style="color:#00bcd4;"></i> 配方预览</h2>
            <div id="previewContent"><p class="text-muted">加载中...</p></div>
        </div>
    </div>
    
    <!-- 发布香调配方弹窗 -->
    <div id="accordModal" class="modal-overlay">
        <div class="modal-content">
            <button class="close-btn" onclick="closeAccordModal()">&times;</button>
            <h2 style="color:#FF9800;"><i class="fas fa-flask" style="color:#FF9800;"></i> 发布香调配方（原材料→香调）</h2>
            <p class="text-muted">此配方将下发至香调生产车间，仅包含原材料级别信息，不含最终产品信息。</p>
            <form method="post" id="accordForm">
                <input type="hidden" name="action" value="publish_accord">
                <input type="hidden" name="recipe_id" id="accordRecipeID">
                <input type="hidden" name="material_count" id="accordMatCount" value="0">
                
                <div class="grid-2 mb-2">
                    <div>
                        <label class="text-muted">选择产出香调</label>
                        <select name="note_id" id="accordNoteSelect" required onchange="loadAccordIngredients()">
                            <option value="">-- 选择香调 --</option>
                        </select>
                    </div>
                    <div>
                        <label class="text-muted">批量大小 (ml)</label>
                        <input type="number" name="batch_size" value="100" min="1" required>
                    </div>
                </div>
                
                <h3>原材料明细（自动从完整配方提取）</h3>
                <div id="accordIngredients" class="ingredient-list">
                    <p class="text-muted">请先选择香调</p>
                </div>
                
                <div class="security-note mt-2">
                    <i class="fas fa-shield-alt"></i> 发布后，香调生产车间仅能看到这些原材料信息，无法获知最终产品。
                </div>
                
                <div class="mt-2 text-right">
                    <button type="button" class="btn btn-outline" onclick="closeAccordModal()">取消</button>
                    <button type="submit" class="btn btn-primary">发布香调配方</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- 发布产品配方弹窗 -->
    <div id="productModal" class="modal-overlay">
        <div class="modal-content">
            <button class="close-btn" onclick="closeProductModal()">&times;</button>
            <h2 style="color:#2196F3;"><i class="fas fa-industry" style="color:#2196F3;"></i> 发布产品配方（香调→产品）</h2>
            <p class="text-muted">此配方将下发至产品制造车间，仅包含香调级别信息，不含原材料成分。</p>
            <form method="post" id="productForm">
                <input type="hidden" name="action" value="publish_product">
                <input type="hidden" name="recipe_id" id="productRecipeID">
                <input type="hidden" name="note_count" id="productNoteCount" value="0">
                
                <div class="grid-2 mb-2">
                    <div>
                        <label class="text-muted">关联产品（可选）</label>
                        <select name="product_id" id="productSelect">
                            <option value="0">-- 不关联具体产品 --</option>
                        </select>
                    </div>
                    <div>
                        <label class="text-muted">批量大小</label>
                        <input type="number" name="batch_size" value="100" min="1" required>
                    </div>
                </div>
                
                <h3>香调组成（自动从完整配方提取）</h3>
                <div id="productNotes" class="ingredient-list">
                    <p class="text-muted">加载中...</p>
                </div>
                
                <div class="security-note mt-2">
                    <i class="fas fa-shield-alt"></i> 发布后，产品制造车间仅能看到香调名称和比例，无法获知香调由哪些原材料构成。
                </div>
                
                <div class="mt-2 text-right">
                    <button type="button" class="btn btn-outline" onclick="closeProductModal()">取消</button>
                    <button type="submit" class="btn btn-primary">发布产品配方</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- Loading 指示器 -->
    <div id="globalLoading" style="display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.5);z-index:2000;justify-content:center;align-items:center;">
        <div style="background:var(--dt-color-bg-light);border-radius:8px;padding:20px 40px;text-align:center;border:1px solid var(--dt-color-border);">
            <i class="fas fa-spinner fa-spin" style="font-size:24px;color:var(--dt-color-primary);"></i>
            <p style="margin-top:10px;color:#ccc;">加载中...</p>
        </div>
    </div>
    
    <script>
    // ========== Loading 指示器 ==========
    function showLoading() { document.getElementById('globalLoading').style.display = 'flex'; }
    function hideLoading() { document.getElementById('globalLoading').style.display = 'none'; }
    function showError(msg) { alert('操作失败：' + msg); }
    
    // ========== 配方预览 ==========
    function openPreviewModal(recipeId, recipeName) {
        document.getElementById('previewModal').classList.add('active');
        document.getElementById('previewContent').innerHTML = '<p class="text-muted">加载中...</p>';
        var xhr = new XMLHttpRequest();
        xhr.open('GET', 'recipe_publish.asp?ajax=recipe_preview&recipe_id=' + recipeId, true);
        xhr.onload = function() {
            if (xhr.status === 200) {
                document.getElementById('previewContent').innerHTML = xhr.responseText;
            } else {
                document.getElementById('previewContent').innerHTML = '<p style="color:#f44336;">加载失败</p>';
            }
        };
        xhr.onerror = function() {
            document.getElementById('previewContent').innerHTML = '<p style="color:#f44336;">网络错误，请重试</p>';
        };
        xhr.send();
    }
    function closePreviewModal() {
        document.getElementById('previewModal').classList.remove('active');
    }
    document.getElementById('previewModal').addEventListener('click', function(e) {
        if (e.target === this) closePreviewModal();
    });
    
    // ========== 香调配方弹窗 ==========
    function openAccordModal(recipeId, recipeName) {
        // 发布前校验：检查是否有RecipeIngredients数据
        showLoading();
        var xhr = new XMLHttpRequest();
        xhr.open('GET', 'recipe_publish.asp?ajax=check_ingredients&recipe_id=' + recipeId, true);
        xhr.onload = function() {
            hideLoading();
            if (xhr.status === 200) {
                try {
                    var data = JSON.parse(xhr.responseText);
                    if (!data.hasData) {
                        alert('请先在配方管理中生成成分数据（RecipeIngredients无数据）');
                        return;
                    }
                } catch(e) { /* 解析失败继续打开 */ }
            }
            document.getElementById('accordRecipeID').value = recipeId;
            document.getElementById('accordModal').classList.add('active');
            loadAccordNotes(recipeId);
        };
        xhr.onerror = function() {
            hideLoading();
            showError('网络请求失败');
        };
        xhr.send();
    }
    function closeAccordModal() {
        document.getElementById('accordModal').classList.remove('active');
    }
    
    function loadAccordNotes(recipeId) {
        var sel = document.getElementById('accordNoteSelect');
        sel.innerHTML = '<option value="">加载中...</option>';
        var xhr = new XMLHttpRequest();
        xhr.open('GET', 'recipe_publish.asp?ajax=accord_notes&recipe_id=' + recipeId, true);
        xhr.onload = function() {
            if (xhr.status === 200) {
                sel.innerHTML = '<option value="">-- 选择香调 --</option>' + xhr.responseText;
            } else {
                sel.innerHTML = '<option value="">加载失败</option>';
                showError('加载香调列表失败');
            }
        };
        xhr.onerror = function() {
            sel.innerHTML = '<option value="">网络错误</option>';
            showError('网络请求失败');
        };
        xhr.send();
    }
    
    function loadAccordIngredients() {
        var recipeId = document.getElementById('accordRecipeID').value;
        var noteId = document.getElementById('accordNoteSelect').value;
        var container = document.getElementById('accordIngredients');
        if (!noteId) { container.innerHTML = '<p class="text-muted">请先选择香调</p>'; return; }
        container.innerHTML = '<p class="text-muted"><i class="fas fa-spinner fa-spin"></i> 加载中...</p>';
        var xhr = new XMLHttpRequest();
        xhr.open('GET', 'recipe_publish.asp?ajax=accord_ingredients&recipe_id=' + recipeId + '&note_id=' + noteId, true);
        xhr.onload = function() {
            if (xhr.status === 200) {
                container.innerHTML = xhr.responseText;
                var count = container.querySelectorAll('.material-row').length;
                document.getElementById('accordMatCount').value = count;
            } else {
                container.innerHTML = '<p style="color:#f44336;">加载失败</p>';
                showError('加载原材料数据失败');
            }
        };
        xhr.onerror = function() {
            container.innerHTML = '<p style="color:#f44336;">网络错误</p>';
            showError('网络请求失败');
        };
        xhr.send();
    }
    
    // ========== 产品配方弹窗 ==========
    function openProductModal(recipeId, recipeName) {
        // 发布前校验：检查是否有RecipeIngredients数据
        showLoading();
        var xhr = new XMLHttpRequest();
        xhr.open('GET', 'recipe_publish.asp?ajax=check_ingredients&recipe_id=' + recipeId, true);
        xhr.onload = function() {
            hideLoading();
            if (xhr.status === 200) {
                try {
                    var data = JSON.parse(xhr.responseText);
                    if (!data.hasData) {
                        alert('请先在配方管理中生成成分数据（RecipeIngredients无数据）');
                        return;
                    }
                } catch(e) { /* 解析失败继续打开 */ }
            }
            document.getElementById('productRecipeID').value = recipeId;
            document.getElementById('productModal').classList.add('active');
            loadProductNotes(recipeId);
            loadProducts();
        };
        xhr.onerror = function() {
            hideLoading();
            showError('网络请求失败');
        };
        xhr.send();
    }
    function closeProductModal() {
        document.getElementById('productModal').classList.remove('active');
    }
    
    function loadProducts() {
        var sel = document.getElementById('productSelect');
        var xhr = new XMLHttpRequest();
        xhr.open('GET', 'recipe_publish.asp?ajax=product_list', true);
        xhr.onload = function() {
            if (xhr.status === 200) {
                sel.innerHTML = '<option value="0">-- 不关联具体产品 --</option>' + xhr.responseText;
            } else {
                showError('加载产品列表失败');
            }
        };
        xhr.onerror = function() { showError('网络请求失败'); };
        xhr.send();
    }
    
    function loadProductNotes(recipeId) {
        var container = document.getElementById('productNotes');
        container.innerHTML = '<p class="text-muted"><i class="fas fa-spinner fa-spin"></i> 加载中...</p>';
        var xhr = new XMLHttpRequest();
        xhr.open('GET', 'recipe_publish.asp?ajax=product_notes&recipe_id=' + recipeId, true);
        xhr.onload = function() {
            if (xhr.status === 200) {
                container.innerHTML = xhr.responseText;
                var count = container.querySelectorAll('.note-row').length;
                document.getElementById('productNoteCount').value = count;
            } else {
                container.innerHTML = '<p style="color:#f44336;">加载失败</p>';
                showError('加载香调组成失败');
            }
        };
        xhr.onerror = function() {
            container.innerHTML = '<p style="color:#f44336;">网络错误</p>';
            showError('网络请求失败');
        };
        xhr.send();
    }
    
    // 点击弹窗外部关闭
    document.getElementById('accordModal').addEventListener('click', function(e) {
        if (e.target === this) closeAccordModal();
    });
    document.getElementById('productModal').addEventListener('click', function(e) {
        if (e.target === this) closeProductModal();
    });
    </script>
</body>
</html>
<%
' ========== AJAX 端点 ==========
Dim ajaxAction
ajaxAction = Trim(Request.QueryString("ajax"))

If ajaxAction <> "" Then
    Response.Clear
    Response.ContentType = "text/html"
End If

If ajaxAction = "accord_notes" Then
    ' 返回配方中的香调列表（用于香调配方发布选择）
    Dim ajRecipeID
    ajRecipeID = SafeNum(Request.QueryString("recipe_id"))
    If ajRecipeID > 0 Then
        Dim rsAN
        Set rsAN = conn.Execute("SELECT DISTINCT rn.NoteID, fn.NoteName, fn.NoteType FROM RecipeNotes rn INNER JOIN FragranceNotes fn ON rn.NoteID=fn.NoteID WHERE rn.RecipeID=" & ajRecipeID & " ORDER BY fn.NoteType, fn.NoteName")
        If Not rsAN Is Nothing Then
            Do While Not rsAN.EOF
                Response.Write "<option value='" & rsAN("NoteID") & "'>" & Server.HTMLEncode(rsAN("NoteName") & " (" & rsAN("NoteType") & ")") & "</option>"
                rsAN.MoveNext
            Loop
            rsAN.Close
        End If
        Set rsAN = Nothing
    End If
    Response.End

ElseIf ajaxAction = "accord_ingredients" Then
    ' 返回配方中某个香调的原材料成分（用于映射到 RawMaterialInventory）
    Dim aiRecipeID, aiNoteID
    aiRecipeID = SafeNum(Request.QueryString("recipe_id"))
    aiNoteID = SafeNum(Request.QueryString("note_id"))
    If aiRecipeID > 0 And aiNoteID > 0 Then
        Dim rsAI, aiIdx
        aiIdx = 0
        Set rsAI = conn.Execute("SELECT IngredientName, Percentage FROM RecipeIngredients WHERE RecipeID=" & aiRecipeID & " AND NoteID=" & aiNoteID & " ORDER BY ID")
        If Not rsAI Is Nothing Then
            Do While Not rsAI.EOF
                aiIdx = aiIdx + 1
                Dim aiName, aiPct, aiMatID
                aiName = CStr(rsAI("IngredientName") & "")
                aiPct = SafeNum(rsAI("Percentage"))
                
                ' 尝试在 RawMaterialInventory 中匹配
                aiMatID = 0
                Dim rsMat
                Set rsMat = conn.Execute("SELECT TOP 1 MaterialID FROM RawMaterialInventory WHERE ItemName='" & SafeSQL(aiName) & "'")
                If Not rsMat Is Nothing Then
                    If Not rsMat.EOF Then aiMatID = SafeNum(rsMat("MaterialID"))
                    rsMat.Close
                End If
                Set rsMat = Nothing
                
                Response.Write "<div class='material-row'>" & _
                    "<input type='hidden' name='material_id_" & aiIdx & "' value='" & aiMatID & "'>" & _
                    "<span class='mat-name'><strong>" & Server.HTMLEncode(aiName) & "</strong>" & _
                    IIF(aiMatID>0, " <small style='color:#27ae60;'>(已匹配库存)</small>", " <small style='color:#f39c12;'>(未匹配库存)</small>") & "</span>" & _
                    "<input type='hidden' name='material_name_" & aiIdx & "' value='" & Server.HTMLEncode(aiName) & "'>" & _
                    "<label>比例(%)</label><input type='number' name='material_pct_" & aiIdx & "' value='" & aiPct & "' min='0' max='100' step='0.01' style='width:80px;'>" & _
                    "<label>计划量(g)</label><input type='number' name='material_qty_" & aiIdx & "' value='" & FormatNumber(aiPct, 1) & "' min='0' step='0.1' style='width:80px;'>" & _
                    "</div>"
                rsAI.MoveNext
            Loop
            rsAI.Close
        End If
        Set rsAI = Nothing
        If aiIdx = 0 Then
            Response.Write "<p class='text-muted'>该香调暂无原材料成分数据（需先在配方管理中配置 RecipeIngredients）</p>"
        End If
    End If
    Response.End

ElseIf ajaxAction = "product_notes" Then
    ' 返回配方中的香调组成（用于产品配方发布）
    Dim pnRecipeID
    pnRecipeID = SafeNum(Request.QueryString("recipe_id"))
    If pnRecipeID > 0 Then
        Dim rsPN, pnIdx
        pnIdx = 0
        Set rsPN = conn.Execute("SELECT rn.NoteID, rn.Percentage, fn.NoteName, fn.NoteType FROM RecipeNotes rn INNER JOIN FragranceNotes fn ON rn.NoteID=fn.NoteID WHERE rn.RecipeID=" & pnRecipeID & " ORDER BY rn.ID")
        If Not rsPN Is Nothing Then
            Do While Not rsPN.EOF
                pnIdx = pnIdx + 1
                Response.Write "<div class='note-row'>" & _
                    "<input type='hidden' name='pnote_id_" & pnIdx & "' value='" & rsPN("NoteID") & "'>" & _
                    "<span class='mat-name'><strong>" & Server.HTMLEncode(rsPN("NoteName") & "") & "</strong> <small>(" & rsPN("NoteType") & ")</small></span>" & _
                    "<input type='hidden' name='pnote_name_" & pnIdx & "' value='" & Server.HTMLEncode(rsPN("NoteName") & "") & "'>" & _
                    "<label>比例(%)</label><input type='number' name='pnote_pct_" & pnIdx & "' value='" & rsPN("Percentage") & "' min='0' max='100' step='0.01' style='width:80px;'>" & _
                    "<label>计划量(ml)</label><input type='number' name='pnote_qty_" & pnIdx & "' value='" & FormatNumber(SafeNum(rsPN("Percentage")), 1) & "' min='0' step='0.1' style='width:80px;'>" & _
                    "</div>"
                rsPN.MoveNext
            Loop
            rsPN.Close
        End If
        Set rsPN = Nothing
        If pnIdx = 0 Then
            Response.Write "<p class='text-muted'>该配方暂无香调组成数据</p>"
        End If
    End If
    Response.End

ElseIf ajaxAction = "product_list" Then
    ' 返回产品列表
    Dim rsPL
    Set rsPL = conn.Execute("SELECT ProductID, ProductName FROM Products ORDER BY ProductName")
    If Not rsPL Is Nothing Then
        Do While Not rsPL.EOF
            Response.Write "<option value='" & rsPL("ProductID") & "'>" & Server.HTMLEncode(rsPL("ProductName") & "") & "</option>"
            rsPL.MoveNext
        Loop
        rsPL.Close
    End If
    Set rsPL = Nothing
    Response.End

ElseIf ajaxAction = "recipe_preview" Then
    ' 返回配方完整信息HTML（配方名、类型、各香调及其成分列表）
    Dim prvRecipeID
    prvRecipeID = SafeNum(Request.QueryString("recipe_id"))
    If prvRecipeID > 0 Then
        Dim rsPrv, prvName, prvType, prvCode
        Set rsPrv = conn.Execute("SELECT RecipeName, RecipeCode, ProductType FROM Recipes WHERE RecipeID=" & prvRecipeID)
        If Not rsPrv Is Nothing Then
            If Not rsPrv.EOF Then
                prvName = CStr(rsPrv("RecipeName") & "")
                prvCode = CStr(rsPrv("RecipeCode") & "")
                prvType = CStr(rsPrv("ProductType") & "")
            End If
            rsPrv.Close
        End If
        Set rsPrv = Nothing
        
        Response.Write "<div style='margin-bottom:12px;'>"
        Response.Write "<h3 style='color:#fff;margin:0 0 4px;'>配方：" & Server.HTMLEncode(prvName) & "</h3>"
        Response.Write "<p class='text-muted' style='margin:0;'>编号：" & Server.HTMLEncode(prvCode) & " | 类型：" & Server.HTMLEncode(prvType) & "</p>"
        Response.Write "</div>"
        
        ' 获取香调组成
        Dim rsPrvNotes
        Set rsPrvNotes = conn.Execute("SELECT rn.NoteID, rn.Percentage, fn.NoteName, fn.NoteType FROM RecipeNotes rn INNER JOIN FragranceNotes fn ON rn.NoteID=fn.NoteID WHERE rn.RecipeID=" & prvRecipeID & " ORDER BY fn.NoteType, fn.NoteName")
        If Not rsPrvNotes Is Nothing Then
            Dim prvNoteIdx
            prvNoteIdx = 0
            Do While Not rsPrvNotes.EOF
                prvNoteIdx = prvNoteIdx + 1
                Dim prvNoteID, prvNoteName, prvNoteType, prvNotePct
                prvNoteID = rsPrvNotes("NoteID")
                prvNoteName = CStr(rsPrvNotes("NoteName") & "")
                prvNoteType = CStr(rsPrvNotes("NoteType") & "")
                prvNotePct = SafeNum(rsPrvNotes("Percentage"))
                
                Response.Write "<div style='background:rgba(255,255,255,0.03);border:1px solid rgba(255,255,255,0.08);border-radius:6px;padding:12px;margin-bottom:10px;'>"
                Response.Write "<div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;'>"
                Response.Write "<strong style='color:#00bcd4;'>" & Server.HTMLEncode(prvNoteName) & "</strong>"
                Response.Write "<span class='text-muted'>" & prvNoteType & " | " & FormatNumber(prvNotePct,1) & "%</span>"
                Response.Write "</div>"
                
                ' 获取该香调的成分
                Dim rsPrvIng
                Set rsPrvIng = conn.Execute("SELECT IngredientName, Percentage FROM RecipeIngredients WHERE RecipeID=" & prvRecipeID & " AND NoteID=" & prvNoteID & " ORDER BY ID")
                If Not rsPrvIng Is Nothing Then
                    Dim hasIng
                    hasIng = False
                    Do While Not rsPrvIng.EOF
                        hasIng = True
                        Response.Write "<div style='padding:3px 0 3px 16px;font-size:13px;color:#b0b0b0;'>" & _
                            "<i class='fas fa-circle' style='font-size:6px;margin-right:8px;color:#666;'></i>" & _
                            Server.HTMLEncode(rsPrvIng("IngredientName") & "") & _
                            " <span style='color:#888;'>(" & FormatNumber(SafeNum(rsPrvIng("Percentage")),2) & "%)</span></div>"
                        rsPrvIng.MoveNext
                    Loop
                    rsPrvIng.Close
                    If Not hasIng Then
                        Response.Write "<p class='text-muted' style='margin:0;padding-left:16px;font-size:12px;'>暂无成分数据</p>"
                    End If
                End If
                Set rsPrvIng = Nothing
                Response.Write "</div>"
                
                rsPrvNotes.MoveNext
            Loop
            rsPrvNotes.Close
            If prvNoteIdx = 0 Then
                Response.Write "<p class='text-muted'>该配方暂无香调组成数据</p>"
            End If
        End If
        Set rsPrvNotes = Nothing
    Else
        Response.Write "<p style='color:#f44336;'>无效的配方ID</p>"
    End If
    Response.End

ElseIf ajaxAction = "check_ingredients" Then
    ' 检查配方是否已有RecipeIngredients数据，返回JSON
    Response.ContentType = "application/json"
    Dim chkRecipeID, chkCount
    chkRecipeID = SafeNum(Request.QueryString("recipe_id"))
    chkCount = 0
    If chkRecipeID > 0 Then
        chkCount = SafeNum(GetScalar("SELECT COUNT(*) FROM RecipeIngredients WHERE RecipeID=" & chkRecipeID))
    End If
    If chkCount > 0 Then
        Response.Write "{""hasData"": true, ""count"": " & chkCount & "}"
    Else
        Response.Write "{""hasData"": false, ""count"": 0}"
    End If
    Response.End

End If

Call CloseConnection()
%>