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

Function SafeCount(val)
    If IsNull(val) Or val = "" Then
        SafeCount = 0
    ElseIf IsNumeric(val) Then
        SafeCount = CLng(val)
    Else
        SafeCount = 0
    End If
End Function

' ========== 获取SiteSettings中的最小比例配置 ==========
Dim minTopPercent, minMiddlePercent, minBasePercent
minTopPercent = 10
minMiddlePercent = 10
minBasePercent = 10

Dim rsSettings
Set rsSettings = ExecuteQuery("SELECT SettingKey, SettingValue FROM SiteSettings WHERE SettingKey IN ('MinTopPercent', 'MinMiddlePercent', 'MinBasePercent')")
If Not rsSettings Is Nothing Then
    Do While Not rsSettings.EOF
        If rsSettings("SettingKey") = "MinTopPercent" Then
            minTopPercent = SafeNum(rsSettings("SettingValue"))
        ElseIf rsSettings("SettingKey") = "MinMiddlePercent" Then
            minMiddlePercent = SafeNum(rsSettings("SettingValue"))
        ElseIf rsSettings("SettingKey") = "MinBasePercent" Then
            minBasePercent = SafeNum(rsSettings("SettingValue"))
        End If
        rsSettings.MoveNext
    Loop
    rsSettings.Close
End If
Set rsSettings = Nothing

If minTopPercent = 0 Then minTopPercent = 10
If minMiddlePercent = 0 Then minMiddlePercent = 10
If minBasePercent = 0 Then minBasePercent = 10

' ========== 处理表单提交 ==========
Dim action, msg, msgType
action = Request.Form("action")
msg = ""
msgType = "error"

' ========== 辅助函数 ==========
Function GetNoteTypeFromDB(noteId)
    Dim rsNT
    Set rsNT = ExecuteQuery("SELECT NoteType FROM FragranceNotes WHERE NoteID = " & CLng(noteId))
    If Not rsNT Is Nothing Then
        If Not rsNT.EOF Then
            GetNoteTypeFromDB = rsNT("NoteType") & ""
        Else
            GetNoteTypeFromDB = ""
        End If
        rsNT.Close
    Else
        GetNoteTypeFromDB = ""
    End If
    Set rsNT = Nothing
End Function

Function GenerateRecipeIngredients(recipeId)
    Dim rsRN, usedIngredients
    Dim currNoteId, rsNI
    Dim ingName, ingKey
    Dim ingPct
    Set usedIngredients = Server.CreateObject("Scripting.Dictionary")
    Set rsRN = ExecuteQuery("SELECT NoteID FROM RecipeNotes WHERE RecipeID = " & CLng(recipeId))
    If Not rsRN Is Nothing Then
        Do While Not rsRN.EOF
            currNoteId = rsRN("NoteID")
            Set rsNI = ExecuteQuery("SELECT ni.NoteID, bn.BaseNoteName, ni.Percentage AS IngPct FROM NoteIngredients ni INNER JOIN BaseNotes bn ON ni.BaseNoteID = bn.BaseNoteID WHERE ni.NoteID = " & currNoteId)
            If Not rsNI Is Nothing Then
                Do While Not rsNI.EOF
                    ingName = CStr(rsNI("BaseNoteName"))
                    ingKey = ingName
                    If Not usedIngredients.Exists(ingKey) Then
                        usedIngredients.Add ingKey, True
                        ingPct = 0
                        If Not IsNull(rsNI("IngPct")) And rsNI("IngPct") & "" <> "" Then
                            ingPct = CDbl(rsNI("IngPct"))
                        End If
                        ExecuteNonQuery "INSERT INTO RecipeIngredients (RecipeID, NoteID, IngredientName, Percentage) VALUES (" & recipeId & ", " & currNoteId & ", '" & SafeSQL(ingName) & "', " & ingPct & ")"
                    End If
                    rsNI.MoveNext
                Loop
                rsNI.Close
                Set rsNI = Nothing
            End If
            rsRN.MoveNext
        Loop
        rsRN.Close
    End If
    Set rsRN = Nothing
    Set usedIngredients = Nothing
End Function

If action = "add_formula" Then
    ' 新增配方
    Dim newRecipeName, newRecipeCode, newProductType, newDescription, newIsActive
    newRecipeName = Trim(Request.Form("formulaName"))
    newRecipeCode = Trim(Request.Form("recipeCode"))
    newProductType = Trim(Request.Form("productType"))
    newDescription = Trim(Request.Form("description"))
    newIsActive = Request.Form("isActive")
    If newIsActive = "" Then newIsActive = 1
    
    ' 获取选中的香调和百分比
    Dim selectedNotes, notePercent
    selectedNotes = Request.Form("selectedNotes")
    
    ' 校验配比
    Dim ratioValid, ratioErrorMsg
    ratioValid = True
    ratioErrorMsg = ""
    
    If selectedNotes = "" Then
        ratioValid = False
        ratioErrorMsg = "请至少选择一个香调"
    Else
        Dim noteArr, noteItem, totalPercent, totalTopPercent, totalMiddlePercent, totalBasePercent
        totalPercent = 0
        totalTopPercent = 0
        totalMiddlePercent = 0
        totalBasePercent = 0
        
        noteArr = Split(selectedNotes, ",")
        For Each noteItem In noteArr
            noteItem = Trim(noteItem)
            If noteItem <> "" And IsNumeric(noteItem) Then
                notePercent = Request.Form("notePercent_" & noteItem)
                If notePercent = "" Or Not IsNumeric(notePercent) Then notePercent = 0
                notePercent = SafeNum(notePercent)
                totalPercent = totalPercent + notePercent
                
                ' 获取香调类型
                Dim rsNoteTypeCheck
                Set rsNoteTypeCheck = ExecuteQuery("SELECT NoteType FROM FragranceNotes WHERE NoteID = " & CLng(noteItem))
                If Not rsNoteTypeCheck Is Nothing Then
                    If Not rsNoteTypeCheck.EOF Then
                        If rsNoteTypeCheck("NoteType") = "前调" Then
                            totalTopPercent = totalTopPercent + notePercent
                        ElseIf rsNoteTypeCheck("NoteType") = "中调" Then
                            totalMiddlePercent = totalMiddlePercent + notePercent
                        ElseIf rsNoteTypeCheck("NoteType") = "后调" Then
                            totalBasePercent = totalBasePercent + notePercent
                        End If
                    End If
                    rsNoteTypeCheck.Close
                End If
                Set rsNoteTypeCheck = Nothing
            End If
        Next
        
        ' 校验规则
        If totalTopPercent < minTopPercent Then
            ratioValid = False
            ratioErrorMsg = "前调比例不能低于" & minTopPercent & "%，当前为" & totalTopPercent & "%"
        ElseIf totalMiddlePercent < minMiddlePercent Then
            ratioValid = False
            ratioErrorMsg = "中调比例不能低于" & minMiddlePercent & "%，当前为" & totalMiddlePercent & "%"
        ElseIf totalBasePercent < minBasePercent Then
            ratioValid = False
            ratioErrorMsg = "后调比例不能低于" & minBasePercent & "%，当前为" & totalBasePercent & "%"
        ElseIf totalPercent <> 100 Then
            ratioValid = False
            ratioErrorMsg = "香调配比总和必须等于100%，当前为" & totalPercent & "%"
        End If
    End If
    
    If newRecipeName = "" Then
        msg = "配方名称不能为空"
    ElseIf newRecipeCode = "" Then
        msg = "配方编号不能为空"
    ElseIf newProductType = "" Then
        msg = "请选择产品类型"
    ElseIf Not ratioValid Then
        msg = ratioErrorMsg
    Else
        ' 确定审核状态
        Dim reviewStatusVal
        If newProductType = "KOL" Then
            reviewStatusVal = "Pending"
        Else
            reviewStatusVal = "Approved"
        End If
        
        Dim adminUsername
        adminUsername = Session("AdminUsername")
        If adminUsername = "" Then adminUsername = "System"
        
        BeginTransaction
        On Error Resume Next
        
        ' 插入配方主表
        Dim addSql, newRecipeId
        addSql = "INSERT INTO Recipes (RecipeName, RecipeCode, Description, ProductType, IsActive, ReviewStatus, CreatedBy, CreatedAt, UpdatedAt) VALUES (" & _
                 "'" & SafeSQL(newRecipeName) & "', " & _
                 "'" & SafeSQL(newRecipeCode) & "', " & _
                 "'" & SafeSQL(newDescription) & "', " & _
                 "'" & SafeSQL(newProductType) & "', " & _
                 newIsActive & ", " & _
                 "'" & reviewStatusVal & "', " & _
                 "'" & SafeSQL(adminUsername) & "', " & _
                 "GETDATE(), GETDATE())"
        
        ExecuteNonQuery addSql
        
        If Err.Number <> 0 Then
            RollbackTransaction
            msg = "添加失败：" & Err.Description
        Else
            newRecipeId = CLng(GetScalar("SELECT MAX(RecipeID) FROM Recipes"))
            
            ' 插入配方-香调关联
            Dim noteTypeVal
            If selectedNotes <> "" Then
                noteArr = Split(selectedNotes, ",")
                For Each noteItem In noteArr
                    noteItem = Trim(noteItem)
                    If noteItem <> "" And IsNumeric(noteItem) Then
                        notePercent = Request.Form("notePercent_" & noteItem)
                        If notePercent = "" Or Not IsNumeric(notePercent) Then notePercent = 0
                        notePercent = SafeNum(notePercent)
                        If notePercent > 0 Then
                            noteTypeVal = GetNoteTypeFromDB(CLng(noteItem))
                            ExecuteNonQuery "INSERT INTO RecipeNotes (RecipeID, NoteID, NoteType, Percentage) VALUES (" & newRecipeId & ", " & CLng(noteItem) & ", '" & SafeSQL(noteTypeVal) & "', " & notePercent & ")"
                        End If
                    End If
                Next
            End If
            
            ' 自动生成成分明细
            GenerateRecipeIngredients newRecipeId
            
            If Err.Number <> 0 Then
                RollbackTransaction
                msg = "保存香调配比失败：" & Err.Description
            Else
                CommitTransaction
                Response.Redirect "formula_management.asp?msg=" & Server.URLEncode("配方添加成功") & "&type=success"
            End If
        End If
        On Error GoTo 0
    End If
    
ElseIf action = "edit_formula" Then
    ' 编辑配方
    Dim editRecipeId, editRecipeName, editRecipeCode, editProductType, editDescription, editIsActive
    editRecipeId = Request.Form("formulaId")
    editRecipeName = Trim(Request.Form("formulaName"))
    editRecipeCode = Trim(Request.Form("recipeCode"))
    editProductType = Trim(Request.Form("productType"))
    editDescription = Trim(Request.Form("description"))
    editIsActive = Request.Form("isActive")
    If editIsActive = "" Then editIsActive = 1
    
    ' 获取选中的香调和百分比
    selectedNotes = Request.Form("selectedNotes")
    
    ' 校验配比
    ratioValid = True
    ratioErrorMsg = ""
    
    If selectedNotes = "" Then
        ratioValid = False
        ratioErrorMsg = "请至少选择一个香调"
    Else
        totalPercent = 0
        totalTopPercent = 0
        totalMiddlePercent = 0
        totalBasePercent = 0
        
        noteArr = Split(selectedNotes, ",")
        For Each noteItem In noteArr
            noteItem = Trim(noteItem)
            If noteItem <> "" And IsNumeric(noteItem) Then
                notePercent = Request.Form("notePercent_" & noteItem)
                If notePercent = "" Or Not IsNumeric(notePercent) Then notePercent = 0
                notePercent = SafeNum(notePercent)
                totalPercent = totalPercent + notePercent
                
                ' 获取香调类型
                Set rsNoteTypeCheck = ExecuteQuery("SELECT NoteType FROM FragranceNotes WHERE NoteID = " & CLng(noteItem))
                If Not rsNoteTypeCheck Is Nothing Then
                    If Not rsNoteTypeCheck.EOF Then
                        If rsNoteTypeCheck("NoteType") = "前调" Then
                            totalTopPercent = totalTopPercent + notePercent
                        ElseIf rsNoteTypeCheck("NoteType") = "中调" Then
                            totalMiddlePercent = totalMiddlePercent + notePercent
                        ElseIf rsNoteTypeCheck("NoteType") = "后调" Then
                            totalBasePercent = totalBasePercent + notePercent
                        End If
                    End If
                    rsNoteTypeCheck.Close
                End If
                Set rsNoteTypeCheck = Nothing
            End If
        Next
        
        ' 校验规则
        If totalTopPercent < minTopPercent Then
            ratioValid = False
            ratioErrorMsg = "前调比例不能低于" & minTopPercent & "%，当前为" & totalTopPercent & "%"
        ElseIf totalMiddlePercent < minMiddlePercent Then
            ratioValid = False
            ratioErrorMsg = "中调比例不能低于" & minMiddlePercent & "%，当前为" & totalMiddlePercent & "%"
        ElseIf totalBasePercent < minBasePercent Then
            ratioValid = False
            ratioErrorMsg = "后调比例不能低于" & minBasePercent & "%，当前为" & totalBasePercent & "%"
        ElseIf totalPercent <> 100 Then
            ratioValid = False
            ratioErrorMsg = "香调配比总和必须等于100%，当前为" & totalPercent & "%"
        End If
    End If
    
    If editRecipeName = "" Then
        msg = "配方名称不能为空"
    ElseIf editRecipeCode = "" Then
        msg = "配方编号不能为空"
    ElseIf editProductType = "" Then
        msg = "请选择产品类型"
    ElseIf Not IsNumeric(editRecipeId) Then
        msg = "无效的配方ID"
    ElseIf Not ratioValid Then
        msg = ratioErrorMsg
    Else
        BeginTransaction
        On Error Resume Next
        
        ' 更新配方主表
        Dim editSql
        editSql = "UPDATE Recipes SET " & _
                  "RecipeName = '" & SafeSQL(editRecipeName) & "', " & _
                  "RecipeCode = '" & SafeSQL(editRecipeCode) & "', " & _
                  "ProductType = '" & SafeSQL(editProductType) & "', " & _
                  "Description = '" & SafeSQL(editDescription) & "', " & _
                  "IsActive = " & editIsActive & ", " & _
                  "UpdatedAt = GETDATE() " & _
                  "WHERE RecipeID = " & CLng(editRecipeId)
        
        ExecuteNonQuery editSql
        
        If Err.Number <> 0 Then
            RollbackTransaction
            msg = "更新失败：" & Err.Description
        Else
            ' 删除旧的关联
            ExecuteNonQuery "DELETE FROM RecipeNotes WHERE RecipeID = " & CLng(editRecipeId)
            ExecuteNonQuery "DELETE FROM RecipeIngredients WHERE RecipeID = " & CLng(editRecipeId)
            
            ' 插入新的关联
            If selectedNotes <> "" Then
                noteArr = Split(selectedNotes, ",")
                For Each noteItem In noteArr
                    noteItem = Trim(noteItem)
                    If noteItem <> "" And IsNumeric(noteItem) Then
                        notePercent = Request.Form("notePercent_" & noteItem)
                        If notePercent = "" Or Not IsNumeric(notePercent) Then notePercent = 0
                        notePercent = SafeNum(notePercent)
                        If notePercent > 0 Then
                            noteTypeVal = GetNoteTypeFromDB(CLng(noteItem))
                            ExecuteNonQuery "INSERT INTO RecipeNotes (RecipeID, NoteID, NoteType, Percentage) VALUES (" & CLng(editRecipeId) & ", " & CLng(noteItem) & ", '" & SafeSQL(noteTypeVal) & "', " & notePercent & ")"
                        End If
                    End If
                Next
            End If
            
            ' 重新生成成分明细
            GenerateRecipeIngredients CLng(editRecipeId)
            
            If Err.Number <> 0 Then
                RollbackTransaction
                msg = "保存香调配比失败：" & Err.Description
            Else
                CommitTransaction
                Response.Redirect "formula_management.asp?msg=" & Server.URLEncode("配方更新成功") & "&type=success"
            End If
        End If
        On Error GoTo 0
    End If
    
ElseIf action = "toggle_status" Then
    ' 切换状态（软删除/恢复）
    Dim toggleId, toggleActive
    toggleId = Request.Form("formulaId")
    toggleActive = Request.Form("isActive")
    
    If IsNumeric(toggleId) Then
        If ExecuteNonQuery("UPDATE Recipes SET IsActive = " & toggleActive & ", UpdatedAt = GETDATE() WHERE RecipeID = " & CLng(toggleId)) Then
            Response.Redirect "formula_management.asp?msg=" & Server.URLEncode("状态更新成功") & "&type=success"
        Else
            msg = "状态更新失败"
        End If
    End If
    
ElseIf action = "approve" Then
    ' 审核通过
    Dim approveId
    approveId = Request.Form("formulaId")
    If IsNumeric(approveId) Then
        If ExecuteNonQuery("UPDATE Recipes SET ReviewStatus = 'Approved', UpdatedAt = GETDATE() WHERE RecipeID = " & CLng(approveId)) Then
            Response.Redirect "formula_management.asp?msg=" & Server.URLEncode("配方已通过审核") & "&type=success"
        Else
            msg = "审核操作失败"
        End If
    End If
    
ElseIf action = "reject" Then
    ' 审核拒绝
    Dim rejectId
    rejectId = Request.Form("formulaId")
    If IsNumeric(rejectId) Then
        If ExecuteNonQuery("UPDATE Recipes SET ReviewStatus = 'Rejected', UpdatedAt = GETDATE() WHERE RecipeID = " & CLng(rejectId)) Then
            Response.Redirect "formula_management.asp?msg=" & Server.URLEncode("配方已拒绝审核") & "&type=success"
        Else
            msg = "审核操作失败"
        End If
    End If
End If

' ========== 获取筛选参数 ==========
Dim searchKeyword, filterStatus, filterType, filterReviewStatus
searchKeyword = Request.QueryString("search")
filterStatus = Request.QueryString("status")
filterType = Request.QueryString("type")
filterReviewStatus = Request.QueryString("review")

' ========== 获取统计数据 ==========
Dim totalCount, activeCount, inactiveCount, customCount, kolCount, pendingCount
totalCount = SafeCount(GetScalar("SELECT COUNT(*) FROM Recipes WHERE ProductType <> 'Fixed'"))
activeCount = SafeCount(GetScalar("SELECT COUNT(*) FROM Recipes WHERE IsActive <> 0 AND ProductType <> 'Fixed'"))
inactiveCount = totalCount - activeCount
customCount = SafeCount(GetScalar("SELECT COUNT(*) FROM Recipes WHERE ProductType = 'Custom'"))
kolCount = SafeCount(GetScalar("SELECT COUNT(*) FROM Recipes WHERE ProductType = 'KOL'"))
pendingCount = SafeCount(GetScalar("SELECT COUNT(*) FROM Recipes WHERE ReviewStatus = 'Pending'"))

' ========== 构建查询SQL ==========
Dim sql, whereClause
' 默认排除品牌定香(Fixed)类型，品牌定香不使用配方管理
whereClause = "WHERE ProductType <> 'Fixed'"

If searchKeyword <> "" Then
    whereClause = "WHERE (RecipeName LIKE '%" & SafeSQL(searchKeyword) & "%' OR RecipeCode LIKE '%" & SafeSQL(searchKeyword) & "%')"
End If

If filterStatus <> "" Then
    If whereClause <> "" Then
        whereClause = whereClause & " AND IsActive = " & filterStatus
    Else
        whereClause = "WHERE IsActive = " & filterStatus
    End If
End If

If filterType <> "" Then
    If whereClause <> "" Then
        whereClause = whereClause & " AND ProductType = '" & SafeSQL(filterType) & "'"
    Else
        whereClause = "WHERE ProductType = '" & SafeSQL(filterType) & "'"
    End If
End If

If filterReviewStatus <> "" Then
    If whereClause <> "" Then
        whereClause = whereClause & " AND ReviewStatus = '" & SafeSQL(filterReviewStatus) & "'"
    Else
        whereClause = "WHERE ReviewStatus = '" & SafeSQL(filterReviewStatus) & "'"
    End If
End If

sql = "SELECT * FROM Recipes " & whereClause & " ORDER BY RecipeID DESC"

Dim rsRecipes
Set rsRecipes = ExecuteQuery(sql)

' ========== 获取所有启用的香调（用于新增/编辑） ==========
Dim rsTopNotes, rsMiddleNotes, rsBaseNotes
Set rsTopNotes = ExecuteQuery("SELECT NoteID, NoteName FROM FragranceNotes WHERE NoteType = '前调' AND IsActive <> 0 ORDER BY NoteName")
Set rsMiddleNotes = ExecuteQuery("SELECT NoteID, NoteName FROM FragranceNotes WHERE NoteType = '中调' AND IsActive <> 0 ORDER BY NoteName")
Set rsBaseNotes = ExecuteQuery("SELECT NoteID, NoteName FROM FragranceNotes WHERE NoteType = '后调' AND IsActive <> 0 ORDER BY NoteName")

' ========== 获取所有基香信息（用于成分溯源） ==========
Dim rsAllBaseNotes
Set rsAllBaseNotes = ExecuteQuery("SELECT BaseNoteID, BaseNoteName, Ingredients FROM BaseNotes WHERE IsActive <> 0")

' ========== 函数：获取配方的香调信息 ==========
Function GetRecipeNotes(recipeId)
    Dim result, rsFN
    result = ""
    Set rsFN = ExecuteQuery("SELECT rn.NoteID, rn.Percentage, n.NoteName, rn.NoteType FROM RecipeNotes rn INNER JOIN FragranceNotes n ON rn.NoteID = n.NoteID WHERE rn.RecipeID = " & recipeId & " ORDER BY rn.NoteType, n.NoteName")
    If Not rsFN Is Nothing Then
        Do While Not rsFN.EOF
            If result <> "" Then result = result & "|"
            result = result & rsFN("NoteID") & ":" & rsFN("NoteName") & ":" & rsFN("NoteType") & ":" & rsFN("Percentage")
            rsFN.MoveNext
        Loop
        rsFN.Close
    End If
    Set rsFN = Nothing
    GetRecipeNotes = result
End Function

' ========== 函数：获取配方成分明细（从RecipeIngredients） ==========
Function GetRecipeIngredientsList(recipeId)
    Dim result, rsRI
    result = ""
    Set rsRI = ExecuteQuery("SELECT DISTINCT IngredientName FROM RecipeIngredients WHERE RecipeID = " & recipeId & " ORDER BY IngredientName")
    If Not rsRI Is Nothing Then
        Do While Not rsRI.EOF
            If result <> "" Then result = result & ", "
            result = result & rsRI("IngredientName")
            rsRI.MoveNext
        Loop
        rsRI.Close
    End If
    Set rsRI = Nothing
    GetRecipeIngredientsList = result
End Function

' ========== 函数：获取香调的基香成分 ==========
Function GetNoteBaseNotes(noteId)
    Dim result, rsNB
    result = ""
    Set rsNB = ExecuteQuery("SELECT ni.BaseNoteID, bn.BaseNoteName FROM NoteIngredients ni INNER JOIN BaseNotes bn ON ni.BaseNoteID = bn.BaseNoteID WHERE ni.NoteID = " & noteId)
    If Not rsNB Is Nothing Then
        Do While Not rsNB.EOF
            If result <> "" Then result = result & ","
            result = result & rsNB("BaseNoteID") & ":" & rsNB("BaseNoteName")
            rsNB.MoveNext
        Loop
        rsNB.Close
    End If
    Set rsNB = Nothing
    GetNoteBaseNotes = result
End Function

' ========== 函数：获取基香的成分列表 ==========
Function GetBaseNoteIngredients(baseNoteId)
    Dim result, rsBI
    result = ""
    Set rsBI = ExecuteQuery("SELECT Ingredients FROM BaseNotes WHERE BaseNoteID = " & baseNoteId)
    If Not rsBI Is Nothing Then
        If Not rsBI.EOF Then
            result = Trim(rsBI("Ingredients") & "")
        End If
        rsBI.Close
    End If
    Set rsBI = Nothing
    GetBaseNoteIngredients = result
End Function
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>配方管理 - 产品技术管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body { background: #1a1a2e; color: #e0e0e0; }
        
        /* 统计卡片和图表区域 */
        .stats-section {
            display: grid;
            grid-template-columns: 1fr 300px;
            gap: 20px;
            margin-bottom: 25px;
        }
        .stats-cards {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 15px;
        }
        .stat-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 10px;
            padding: 20px;
            text-align: center;
            border: 1px solid rgba(255,255,255,0.05);
            transition: all 0.3s ease;
            cursor: pointer;
        }
        .stat-card:hover { transform: translateY(-2px); box-shadow: 0 4px 15px rgba(0,188,212,0.2); }
        .stat-card.active { border-color: #00bcd4; box-shadow: 0 0 15px rgba(0,188,212,0.3); }
        .stat-value { font-size: 32px; font-weight: 700; color: #fff; }
        .stat-label { font-size: 12px; color: #888; margin-top: 5px; text-transform: uppercase; }
        .stat-card.total .stat-value { color: #00bcd4; }
        .stat-card.enabled .stat-value { color: #4caf50; }
        .stat-card.disabled .stat-value { color: #f44336; }
        
        /* 图表容器 */
        .chart-container {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 10px;
            padding: 15px;
            border: 1px solid rgba(255,255,255,0.05);
            display: flex;
            align-items: center;
            justify-content: center;
            position: relative;
        }
        .chart-wrapper {
            position: relative;
            width: 180px;
            height: 180px;
        }
        .chart-center-text {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            text-align: center;
        }
        .chart-center-value {
            font-size: 28px;
            font-weight: 700;
            color: #fff;
        }
        .chart-center-label {
            font-size: 11px;
            color: #888;
        }
        
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
        .filter-group { display: flex; align-items: center; gap: 8px; }
        .filter-label { color: #888; font-size: 13px; }
        .filter-input {
            background: rgba(0,0,0,0.3);
            border: 1px solid #3a3a5a;
            border-radius: 6px;
            padding: 8px 12px;
            color: #fff;
            font-size: 13px;
            width: 200px;
        }
        .filter-input::placeholder { color: #999; }
        /* .filter-btn 样式已由 /css/buttons.css Section 5 统一管理 */
        
        /* 配方卡片网格 */
        .formula-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(380px, 1fr));
            gap: 20px;
        }
        .formula-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 20px;
            border: 1px solid rgba(255,255,255,0.05);
            transition: all 0.3s ease;
        }
        .formula-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.3);
            border-color: rgba(0,188,212,0.2);
        }
        .formula-header {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            margin-bottom: 12px;
        }
        .formula-title {
            font-size: 16px;
            font-weight: 600;
            color: #fff;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .formula-title i { color: #00bcd4; }
        .formula-id {
            font-size: 11px;
            color: #999;
            background: rgba(0,0,0,0.3);
            padding: 2px 8px;
            border-radius: 4px;
        }
        .formula-desc {
            color: #888;
            font-size: 13px;
            margin-bottom: 15px;
            line-height: 1.5;
            min-height: 20px;
        }
        .formula-desc.empty { color: #999; font-style: italic; }
        
        /* 香调配比展示 */
        .notes-section { margin-bottom: 15px; }
        .notes-label {
            font-size: 11px;
            color: #999;
            text-transform: uppercase;
            margin-bottom: 8px;
        }
        .notes-list { display: flex; flex-wrap: wrap; gap: 6px; }
        .note-tag {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            padding: 4px 10px;
            border-radius: 4px;
            font-size: 12px;
            border: 1px solid;
        }
        .note-tag.top { background: rgba(255,193,7,0.15); color: #ffc107; border-color: rgba(255,193,7,0.3); }
        .note-tag.middle { background: rgba(156,39,176,0.15); color: #ce93d8; border-color: rgba(156,39,176,0.3); }
        .note-tag.base { background: rgba(0,150,136,0.15); color: #80cbc4; border-color: rgba(0,150,136,0.3); }
        .note-percent { font-weight: 600; }
        
        /* 成分溯源 */
        .ingredients-section {
            margin-bottom: 15px;
            padding: 10px;
            background: rgba(0,0,0,0.2);
            border-radius: 6px;
        }
        .ingredients-label {
            font-size: 11px;
            color: #999;
            margin-bottom: 6px;
            display: flex;
            align-items: center;
            gap: 5px;
        }
        .ingredients-list {
            display: flex;
            flex-wrap: wrap;
            gap: 5px;
        }
        .ingredient-tag {
            background: rgba(0,188,212,0.1);
            color: #00bcd4;
            padding: 3px 8px;
            border-radius: 3px;
            font-size: 11px;
            border: 1px solid rgba(0,188,212,0.2);
        }
        .ingredients-empty { color: #999; font-size: 11px; font-style: italic; }
        
        /* 卡片底部 */
        .formula-footer {
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
        .status-badge.active { background: rgba(76, 175, 80, 0.2); color: #4caf50; }
        .status-badge.inactive { background: rgba(244, 67, 54, 0.2); color: #f44336; }
        .product-type-badge {
            display: inline-block;
            padding: 3px 8px;
            border-radius: 4px;
            font-size: 10px;
            font-weight: 600;
            margin-right: 6px;
            text-transform: uppercase;
        }
        .product-type-badge.fixed { background: rgba(33, 150, 243, 0.2); color: #2196f3; }
        .product-type-badge.custom { background: rgba(76, 175, 80, 0.2); color: #4caf50; }
        .product-type-badge.kol { background: rgba(255, 152, 0, 0.2); color: #ff9800; }
        .review-badge {
            display: inline-block;
            padding: 3px 8px;
            border-radius: 4px;
            font-size: 10px;
            font-weight: 600;
        }
        .review-badge.pending { background: rgba(255, 193, 7, 0.2); color: #ffc107; }
        .review-badge.approved { background: rgba(76, 175, 80, 0.2); color: #4caf50; }
        .review-badge.rejected { background: rgba(244, 67, 54, 0.2); color: #f44336; }
        .action-btns { display: flex; gap: 8px; flex-wrap: wrap; }
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
            overflow-y: auto;
        }
        .admin-modal-content {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%) !important;
            margin: 3% auto;
            border-radius: 12px;
            width: 90%;
            max-width: 700px;
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
        .admin-modal-title { font-size: 18px; font-weight: 600; color: #fff; margin: 0; }
        .admin-modal-close { background: none; border: none; color: #bbb; font-size: 24px; cursor: pointer; }
        .admin-modal-close:hover { color: #fff; }
        .admin-modal-body { padding: 20px; max-height: 70vh; overflow-y: auto; }
        .admin-modal-footer {
            padding: 15px 20px;
            border-top: 1px solid #3a3a5a;
            display: flex;
            justify-content: flex-end;
            gap: 10px;
        }
        
        /* 表单样式 */
        .form-group { margin-bottom: 15px; }
        .form-label { display: block; margin-bottom: 8px; color: #e0e0e0; font-size: 13px; }
        .admin-modal-content .form-label {
            color: #e0e0e0 !important;
        }
        .form-label .required { color: #f44336; }
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
        .form-control:focus { outline: none; border-color: #00bcd4; }
        textarea.form-control { resize: vertical; min-height: 60px; }
        .form-control::placeholder { color: #999; }
        .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }
        
        /* 香调选择区域 */
        .notes-selection {
            background: rgba(0,0,0,0.2);
            border-radius: 8px;
            padding: 15px;
            margin-bottom: 15px;
        }
        .notes-category { margin-bottom: 15px; }
        .notes-category:last-child { margin-bottom: 0; }
        .category-title {
            font-size: 13px;
            color: #c0c0c0;
            margin-bottom: 10px;
            padding-bottom: 5px;
            border-bottom: 1px solid rgba(255,255,255,0.05);
            display: flex;
            align-items: center;
            gap: 6px;
        }
        .category-title.top { color: #ffc107; }
        .category-title.middle { color: #ce93d8; }
        .category-title.base { color: #80cbc4; }
        .note-checkbox-list { display: flex; flex-wrap: wrap; gap: 10px; }
        .note-checkbox-item {
            display: flex;
            align-items: center;
            gap: 6px;
            background: rgba(255,255,255,0.03);
            padding: 8px 12px;
            border-radius: 6px;
            border: 1px solid rgba(255,255,255,0.05);
        }
        .note-checkbox-item:hover { background: rgba(255,255,255,0.05); }
        .note-checkbox-item input[type="checkbox"] { accent-color: #00bcd4; }
        .note-checkbox-item label { color: #e0e0e0; font-size: 13px; cursor: pointer; }
        .note-percent-input {
            width: 60px;
            background: rgba(0,0,0,0.3);
            border: 1px solid #3a3a5a;
            border-radius: 4px;
            padding: 4px 8px;
            color: #fff;
            font-size: 13px;
            text-align: center;
        }
        .note-percent-input::placeholder { color: #999; }
        .note-percent-input:focus { outline: none; border-color: #00bcd4; }
        .note-percent-input:disabled { opacity: 0.3; cursor: not-allowed; }
        
        /* 配比统计 */
        .ratio-summary {
            background: rgba(0,188,212,0.1);
            border: 1px solid rgba(0,188,212,0.2);
            border-radius: 8px;
            padding: 15px;
            margin-top: 15px;
        }
        .ratio-summary-title {
            font-size: 13px;
            color: #00bcd4;
            margin-bottom: 10px;
            display: flex;
            align-items: center;
            gap: 6px;
        }
        .ratio-bars { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 10px; }
        .ratio-bar-item { text-align: center; }
        .ratio-bar-label { font-size: 11px; color: #b0b0b0; margin-bottom: 4px; }
        .ratio-bar-value {
            font-size: 18px;
            font-weight: 600;
            margin-bottom: 4px;
        }
        .ratio-bar-value.valid { color: #4caf50; }
        .ratio-bar-value.invalid { color: #f44336; }
        .ratio-bar-min { font-size: 10px; color: #999; }
        .ratio-total {
            text-align: center;
            padding-top: 10px;
            border-top: 1px solid rgba(255,255,255,0.05);
            font-size: 14px;
        }
        .ratio-total-value { font-weight: 600; }
        .ratio-total-value.valid { color: #4caf50; }
        .ratio-total-value.invalid { color: #f44336; }
        
        /* 空状态 */
        .empty-state {
            text-align: center;
            padding: 60px 20px;
            color: #999;
            grid-column: 1 / -1;
        }
        .empty-state i { font-size: 64px; margin-bottom: 20px; color: #888; }
        .empty-state h3 { font-size: 18px; margin-bottom: 10px; color: #888; }
        
        /* 提示消息 */
        .alert {
            padding: 12px 15px;
            border-radius: 6px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .alert-success { background: rgba(76,175,80,0.1); border: 1px solid rgba(76,175,80,0.3); color: #4caf50; }
        .alert-error { background: rgba(244,67,54,0.1); border: 1px solid rgba(244,67,54,0.3); color: #f44336; }
        
        /* 响应式 */
        @media (max-width: 1200px) {
            .stats-section { grid-template-columns: 1fr; }
            .chart-container { display: none; }
        }
        @media (max-width: 768px) {
            .stats-cards { grid-template-columns: 1fr; }
            .formula-grid { grid-template-columns: 1fr; }
            .filter-bar { flex-direction: column; align-items: stretch; }
            .form-row { grid-template-columns: 1fr; }
            .ratio-bars { grid-template-columns: 1fr; }
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
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-flask"></i> 配方管理</h2>
            <div class="breadcrumb">
                <a href="index.asp">技术中心</a> / <span>配方管理</span>
            </div>
        </div>
        
        <% If Request.QueryString("msg") <> "" Then %>
        <div class="alert alert-<%= IIf(Request.QueryString("type")="success", "success", "error") %>">
            <i class="fas fa-<%= IIf(Request.QueryString("type")="success", "check-circle", "exclamation-circle") %>"></i>
            <%= Server.HTMLEncode(Request.QueryString("msg")) %>
        </div>
        <% End If %>
        
        <% If msg <> "" Then %>
        <div class="alert alert-error">
            <i class="fas fa-exclamation-circle"></i> <%= Server.HTMLEncode(msg) %>
        </div>
        <% End If %>
        
        <!-- 统计区域 -->
        <div class="stats-section">
            <div class="stats-cards">
                <div class="stat-card total <%= IIf(filterStatus="" And filterType="" And filterReviewStatus="", "active", "") %>" onclick="location.href='formula_management.asp'">
                    <div class="stat-value"><%= totalCount %></div>
                    <div class="stat-label">总配方数</div>
                </div>
                <div class="stat-card enabled <%= IIf(filterStatus="1", "active", "") %>" onclick="location.href='formula_management.asp?status=1'">
                    <div class="stat-value"><%= activeCount %></div>
                    <div class="stat-label">已启用</div>
                </div>
                <div class="stat-card disabled <%= IIf(filterStatus="0", "active", "") %>" onclick="location.href='formula_management.asp?status=0'">
                    <div class="stat-value"><%= inactiveCount %></div>
                    <div class="stat-label">已禁用</div>
                </div>
                <div class="stat-card" style="border-top: 3px solid #4caf50; <%= IIf(filterType="Custom", "border-color: #4caf50; box-shadow: 0 0 15px rgba(76,175,80,0.3);", "") %>" onclick="location.href='formula_management.asp?type=Custom'">
                    <div class="stat-value" style="color: #4caf50;"><%= customCount %></div>
                    <div class="stat-label">用户定制</div>
                </div>
                <div class="stat-card" style="border-top: 3px solid #9c27b0; <%= IIf(filterType="KOL", "border-color: #9c27b0; box-shadow: 0 0 15px rgba(156,39,176,0.3);", "") %>" onclick="location.href='formula_management.asp?type=KOL'">
                    <div class="stat-value" style="color: #9c27b0;"><%= kolCount %></div>
                    <div class="stat-label">KOL推荐</div>
                </div>
                <div class="stat-card" style="border-top: 3px solid #f44336; <%= IIf(filterReviewStatus="Pending", "border-color: #f44336; box-shadow: 0 0 15px rgba(244,67,54,0.3);", "") %>" onclick="location.href='formula_management.asp?review=Pending'">
                    <div class="stat-value" style="color: #f44336;"><%= pendingCount %></div>
                    <div class="stat-label">待审核</div>
                </div>
            </div>
            <div class="chart-container">
                <div class="chart-wrapper">
                    <canvas id="formulaChart"></canvas>
                    <div class="chart-center-text">
                        <div class="chart-center-value"><%= totalCount %></div>
                        <div class="chart-center-label">配方总数</div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- 筛选栏 -->
        <div class="filter-bar">
            <div class="filter-group" style="flex: 1;">
                <span class="filter-label"><i class="fas fa-search"></i> 搜索:</span>
                <input type="text" class="filter-input" id="searchInput" placeholder="输入配方名称或编号..." value="<%= Server.HTMLEncode(searchKeyword) %>">
            </div>
            <div class="filter-group">
                <span class="filter-label">类型:</span>
                <select class="filter-input" id="filterType" style="width: 120px;">
                    <option value="">全部</option>
                    <option value="Custom" <%= IIf(filterType="Custom","selected","") %>>用户定制</option>
                    <option value="KOL" <%= IIf(filterType="KOL","selected","") %>>KOL推荐</option>
                </select>
            </div>
            <div class="filter-group">
                <span class="filter-label">审核:</span>
                <select class="filter-input" id="filterReview" style="width: 120px;">
                    <option value="">全部</option>
                    <option value="Pending" <%= IIf(filterReviewStatus="Pending","selected","") %>>待审核</option>
                    <option value="Approved" <%= IIf(filterReviewStatus="Approved","selected","") %>>已通过</option>
                    <option value="Rejected" <%= IIf(filterReviewStatus="Rejected","selected","") %>>已拒绝</option>
                </select>
            </div>
            <button class="filter-btn" onclick="doSearch()">
                <i class="fas fa-search"></i> 搜索
            </button>
            <% If searchKeyword <> "" Or filterType <> "" Or filterReviewStatus <> "" Or filterStatus <> "" Then %>
            <button class="filter-btn" onclick="location.href='formula_management.asp'">
                <i class="fas fa-times"></i> 清除
            </button>
            <% End If %>
            <% If isManager Then %>
            <button class="filter-btn secondary" onclick="showAddModal()">
                <i class="fas fa-plus"></i> 新增配方
            </button>
            <% End If %>
        </div>
        
        <!-- 配方卡片列表 -->
        <div class="formula-grid">
            <% 
            Dim hasRecords
            hasRecords = False
            If Not rsRecipes Is Nothing Then 
                If Not rsRecipes.EOF Then
                    hasRecords = True
                End If
            End If
            
            If hasRecords Then 
                Do While Not rsRecipes.EOF
                    Dim recipeNotes, noteParts, noteInfo, noteTypeClass
                    Dim allIngredients, ingredientArr, ingredientDict
                    Dim notesArr, n
                    Dim baseNotes, baseParts, b
                    Dim baseInfo, baseId
                    Dim baseIngs, ingList, ing
                    Dim displayIngCount
                    recipeNotes = GetRecipeNotes(rsRecipes("RecipeID"))
                    
                    ' 获取成分明细（优先从RecipeIngredients读取）
                    allIngredients = GetRecipeIngredientsList(rsRecipes("RecipeID"))
                    
                    ' 如果RecipeIngredients没有数据，回退到NoteIngredients查询
                    If allIngredients = "" Then
                        allIngredients = ""
                        Set ingredientDict = CreateObject("Scripting.Dictionary")
                        
                        If recipeNotes <> "" Then
                            notesArr = Split(recipeNotes, "|")
                            For Each n In notesArr
                                If n <> "" Then
                                    noteParts = Split(n, ":")
                                    If UBound(noteParts) >= 0 Then
                                        baseNotes = GetNoteBaseNotes(noteParts(0))
                                        If baseNotes <> "" Then
                                            baseParts = Split(baseNotes, ",")
                                            For Each b In baseParts
                                                If b <> "" Then
                                                    baseInfo = Split(b, ":")
                                                    If UBound(baseInfo) >= 0 Then
                                                        baseId = baseInfo(0)
                                                        If Not ingredientDict.Exists(baseId) Then
                                                            ingredientDict.Add baseId, True
                                                            baseIngs = GetBaseNoteIngredients(baseId)
                                                            If baseIngs <> "" Then
                                                                ingList = Split(baseIngs, ",")
                                                                For Each ing In ingList
                                                                    ing = Trim(ing)
                                                                    If ing <> "" And Not ingredientDict.Exists(ing) Then
                                                                        ingredientDict.Add ing, True
                                                                        If allIngredients <> "" Then allIngredients = allIngredients & ", "
                                                                        allIngredients = allIngredients & ing
                                                                    End If
                                                                Next
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Next
                                        End If
                                    End If
                                End If
                            Next
                        End If
                        Set ingredientDict = Nothing
                    End If
            %>
            <div class="formula-card">
                <div class="formula-header">
                    <div class="formula-title">
                        <i class="fas fa-flask"></i>
                        <%= HTMLEncode(rsRecipes("RecipeName")) %>
                    </div>
                    <span class="formula-id">#<%= HTMLEncode(rsRecipes("RecipeCode")) %></span>
                </div>
                
                <div style="margin-bottom: 10px;">
                    <% Select Case rsRecipes("ProductType")
                        Case "Fixed" %>
                        <span class="product-type-badge fixed"><i class="fas fa-box"></i> 品牌定香</span>
                    <%  Case "Custom" %>
                        <span class="product-type-badge custom"><i class="fas fa-paint-brush"></i> 用户定制</span>
                    <%  Case "KOL" %>
                        <span class="product-type-badge kol"><i class="fas fa-star"></i> KOL推荐</span>
                    <%  Case Else %>
                        <span class="product-type-badge"><%= HTMLEncode(rsRecipes("ProductType")) %></span>
                    <% End Select %>
                    
                    <% Select Case rsRecipes("ReviewStatus")
                        Case "Pending" %>
                        <span class="review-badge pending"><i class="fas fa-clock"></i> 待审核</span>
                    <%  Case "Approved" %>
                        <span class="review-badge approved"><i class="fas fa-check-circle"></i> 已通过</span>
                    <%  Case "Rejected" %>
                        <span class="review-badge rejected"><i class="fas fa-times-circle"></i> 已拒绝</span>
                    <%  Case Else %>
                        <span class="review-badge approved"><i class="fas fa-check-circle"></i> 已通过</span>
                    <% End Select %>
                </div>
                
                <% 
                Dim recipeDesc
                recipeDesc = Trim(rsRecipes("Description") & "")
                If recipeDesc <> "" Then
                %>
                <div class="formula-desc"><%= HTMLEncode(Left(recipeDesc, 100)) %><% If Len(recipeDesc) > 100 Then Response.Write "..." %></div>
                <% Else %>
                <div class="formula-desc empty">暂无描述</div>
                <% End If %>
                
                <!-- 香调配比 -->
                <div class="notes-section">
                    <div class="notes-label"><i class="fas fa-percentage"></i> 香调配比</div>
                    <div class="notes-list">
                        <% If recipeNotes <> "" Then
                            notesArr = Split(recipeNotes, "|")
                            For Each n In notesArr
                                If n <> "" Then
                                    noteParts = Split(n, ":")
                                    If UBound(noteParts) >= 3 Then
                                        Select Case noteParts(2)
                                            Case "前调": noteTypeClass = "top"
                                            Case "中调": noteTypeClass = "middle"
                                            Case "后调": noteTypeClass = "base"
                                            Case Else: noteTypeClass = "top"
                                        End Select
                        %>
                        <span class="note-tag <%= noteTypeClass %>">
                            <%= HTMLEncode(noteParts(1)) %>
                            <span class="note-percent"><%= noteParts(3) %>%</span>
                        </span>
                        <% 
                                    End If
                                End If
                            Next
                        End If %>
                    </div>
                </div>
                
                <!-- 成分溯源 -->
                <div class="ingredients-section">
                    <div class="ingredients-label"><i class="fas fa-leaf"></i> 成分溯源</div>
                    <div class="ingredients-list">
                        <% 
                        If allIngredients <> "" Then
                            ingredientArr = Split(allIngredients, ", ")
                            displayIngCount = 0
                            For Each ing In ingredientArr
                                If Trim(ing) <> "" Then
                                    displayIngCount = displayIngCount + 1
                                    If displayIngCount <= 8 Then
                        %>
                        <span class="ingredient-tag"><%= HTMLEncode(Trim(ing)) %></span>
                        <% 
                                    End If
                                End If
                            Next
                            If displayIngCount > 8 Then
                        %>
                        <span class="ingredient-tag" style="background: rgba(255,255,255,0.1); color: #b0b0b0;">+<%= displayIngCount - 8 %> 更多</span>
                        <% 
                            End If
                        Else
                        %>
                        <span class="ingredients-empty">暂无成分信息</span>
                        <% End If %>
                    </div>
                </div>
                
                <div class="formula-footer">
                    <span class="status-badge <%= IIf(rsRecipes("IsActive"), "active", "inactive") %>">
                        <%= IIf(rsRecipes("IsActive"), "启用", "禁用") %>
                    </span>
                    <div class="action-btns">
                        <button class="action-btn edit" onclick="showEditModal(<%= rsRecipes("RecipeID") %>, '<%= SafeOutput(rsRecipes("RecipeName")) %>', '<%= SafeOutput(rsRecipes("RecipeCode")) %>', '<%= SafeOutput(rsRecipes("ProductType")) %>', '<%= SafeOutput(rsRecipes("Description")) %>', <%= IIf(rsRecipes("IsActive"), 1, 0) %>, '<%= recipeNotes %>')">
                            <i class="fas fa-edit"></i> 编辑
                        </button>
                        <% If isManager Then %>
                        <form method="post" style="display:inline;" onsubmit="return confirm('<%= IIf(rsRecipes("IsActive"), "确定要禁用此配方吗？", "确定要启用此配方吗？") %>')">
                            <input type="hidden" name="action" value="toggle_status">
                            <input type="hidden" name="formulaId" value="<%= rsRecipes("RecipeID") %>">
                            <input type="hidden" name="isActive" value="<%= IIf(rsRecipes("IsActive"), 0, 1) %>">
                            <button type="submit" class="action-btn <%= IIf(rsRecipes("IsActive"), "reject", "approve") %>">
                                <i class="fas fa-<%= IIf(rsRecipes("IsActive"), "ban", "check") %>"></i> <%= IIf(rsRecipes("IsActive"), "禁用", "启用") %>
                            </button>
                        </form>
                        <% End If %>
                        <% If rsRecipes("ReviewStatus") = "Pending" And isManager Then %>
                        <form method="post" style="display:inline;" onsubmit="return confirm('确定要通过该配方审核吗？')">
                            <input type="hidden" name="action" value="approve">
                            <input type="hidden" name="formulaId" value="<%= rsRecipes("RecipeID") %>">
                            <button type="submit" class="action-btn approve">
                                <i class="fas fa-check"></i> 通过
                            </button>
                        </form>
                        <form method="post" style="display:inline;" onsubmit="return confirm('确定要拒绝该配方审核吗？')">
                            <input type="hidden" name="action" value="reject">
                            <input type="hidden" name="formulaId" value="<%= rsRecipes("RecipeID") %>">
                            <button type="submit" class="action-btn reject">
                                <i class="fas fa-times"></i> 拒绝
                            </button>
                        </form>
                        <% End If %>
                    </div>
                </div>
            </div>
            <% 
                    rsRecipes.MoveNext
                Loop
                rsRecipes.Close
                Set rsRecipes = Nothing
            Else
            %>
            <div class="empty-state">
                <i class="fas fa-flask"></i>
                <h3>暂无配方数据</h3>
                <p>点击"新增配方"按钮创建第一个配方</p>
            </div>
            <% End If %>
        </div>
    </div>
    
    <!-- 新增/编辑配方模态框 -->
    <div id="formulaModal" class="admin-modal">
        <div class="admin-modal-content">
            <div class="admin-modal-header">
                <h3 class="admin-modal-title" id="modalTitle"><i class="fas fa-plus"></i> 新增配方</h3>
                <button class="admin-modal-close" onclick="closeModal()">&times;</button>
            </div>
            <form method="post" id="formulaForm" onsubmit="return validateForm()">
                <div class="admin-modal-body">
                    <input type="hidden" name="action" id="formAction" value="add_formula">
                    <input type="hidden" name="formulaId" id="formulaId" value="">
                    <input type="hidden" name="selectedNotes" id="selectedNotes" value="">
                    
                    <div class="form-group">
                        <label class="form-label">配方名称 <span class="required">*</span></label>
                        <input type="text" name="formulaName" id="formulaName" class="form-control" required placeholder="输入配方名称">
                    </div>
                    
                    <div class="form-row">
                        <div class="form-group">
                            <label class="form-label">配方编号 <span class="required">*</span></label>
                            <input type="text" name="recipeCode" id="recipeCode" class="form-control" required placeholder="输入配方编号">
                        </div>
                        <div class="form-group">
                            <label class="form-label">产品类型 <span class="required">*</span></label>
                            <select name="productType" id="productType" class="form-control" required>
                                <option value="">请选择</option>
                                <option value="Custom">用户定制</option>
                                <option value="KOL">KOL推荐</option>
                            </select>
                        </div>
                    </div>
                    
                    <div class="form-group">
                        <label class="form-label">配方描述</label>
                        <textarea name="description" id="description" class="form-control" placeholder="输入配方描述信息"></textarea>
                    </div>
                    
                    <div class="form-group">
                        <label class="form-label">状态</label>
                        <select name="isActive" id="isActive" class="form-control">
                            <option value="1">启用</option>
                            <option value="0">禁用</option>
                        </select>
                    </div>
                    
                    <!-- 香调选择区域 -->
                    <div class="notes-selection">
                        <div class="form-label"><i class="fas fa-list"></i> 选择香调并设置配比 <span class="required">*</span></div>
                        
                        <!-- 前调 -->
                        <div class="notes-category">
                            <div class="category-title top"><i class="fas fa-sun"></i> 前调</div>
                            <div class="note-checkbox-list">
                                <% If Not rsTopNotes Is Nothing Then
                                    Do While Not rsTopNotes.EOF
                                %>
                                <div class="note-checkbox-item">
                                    <input type="checkbox" id="note_<%= rsTopNotes("NoteID") %>" value="<%= rsTopNotes("NoteID") %>" onchange="togglePercentInput(this)">
                                    <label for="note_<%= rsTopNotes("NoteID") %>"><%= HTMLEncode(rsTopNotes("NoteName")) %></label>
                                    <input type="number" class="note-percent-input" id="percent_<%= rsTopNotes("NoteID") %>" name="notePercent_<%= rsTopNotes("NoteID") %>" min="0" max="100" placeholder="%" disabled onchange="calculateRatio()">
                                </div>
                                <% 
                                        rsTopNotes.MoveNext
                                    Loop
                                    rsTopNotes.Close
                                    Set rsTopNotes = Nothing
                                End If %>
                            </div>
                        </div>
                        
                        <!-- 中调 -->
                        <div class="notes-category">
                            <div class="category-title middle"><i class="fas fa-cloud"></i> 中调</div>
                            <div class="note-checkbox-list">
                                <% If Not rsMiddleNotes Is Nothing Then
                                    Do While Not rsMiddleNotes.EOF
                                %>
                                <div class="note-checkbox-item">
                                    <input type="checkbox" id="note_<%= rsMiddleNotes("NoteID") %>" value="<%= rsMiddleNotes("NoteID") %>" onchange="togglePercentInput(this)">
                                    <label for="note_<%= rsMiddleNotes("NoteID") %>"><%= HTMLEncode(rsMiddleNotes("NoteName")) %></label>
                                    <input type="number" class="note-percent-input" id="percent_<%= rsMiddleNotes("NoteID") %>" name="notePercent_<%= rsMiddleNotes("NoteID") %>" min="0" max="100" placeholder="%" disabled onchange="calculateRatio()">
                                </div>
                                <% 
                                        rsMiddleNotes.MoveNext
                                    Loop
                                    rsMiddleNotes.Close
                                    Set rsMiddleNotes = Nothing
                                End If %>
                            </div>
                        </div>
                        
                        <!-- 后调 -->
                        <div class="notes-category">
                            <div class="category-title base"><i class="fas fa-moon"></i> 后调</div>
                            <div class="note-checkbox-list">
                                <% If Not rsBaseNotes Is Nothing Then
                                    Do While Not rsBaseNotes.EOF
                                %>
                                <div class="note-checkbox-item">
                                    <input type="checkbox" id="note_<%= rsBaseNotes("NoteID") %>" value="<%= rsBaseNotes("NoteID") %>" onchange="togglePercentInput(this)">
                                    <label for="note_<%= rsBaseNotes("NoteID") %>"><%= HTMLEncode(rsBaseNotes("NoteName")) %></label>
                                    <input type="number" class="note-percent-input" id="percent_<%= rsBaseNotes("NoteID") %>" name="notePercent_<%= rsBaseNotes("NoteID") %>" min="0" max="100" placeholder="%" disabled onchange="calculateRatio()">
                                </div>
                                <% 
                                        rsBaseNotes.MoveNext
                                    Loop
                                    rsBaseNotes.Close
                                    Set rsBaseNotes = Nothing
                                End If %>
                            </div>
                        </div>
                    </div>
                    
                    <!-- 配比统计 -->
                    <div class="ratio-summary">
                        <div class="ratio-summary-title"><i class="fas fa-chart-pie"></i> 配比统计</div>
                        <div class="ratio-bars">
                            <div class="ratio-bar-item">
                                <div class="ratio-bar-label">前调</div>
                                <div class="ratio-bar-value" id="topPercentDisplay">0%</div>
                                <div class="ratio-bar-min">最低 <%= minTopPercent %>%</div>
                            </div>
                            <div class="ratio-bar-item">
                                <div class="ratio-bar-label">中调</div>
                                <div class="ratio-bar-value" id="middlePercentDisplay">0%</div>
                                <div class="ratio-bar-min">最低 <%= minMiddlePercent %>%</div>
                            </div>
                            <div class="ratio-bar-item">
                                <div class="ratio-bar-label">后调</div>
                                <div class="ratio-bar-value" id="basePercentDisplay">0%</div>
                                <div class="ratio-bar-min">最低 <%= minBasePercent %>%</div>
                            </div>
                        </div>
                        <div class="ratio-total">
                            总配比: <span class="ratio-total-value" id="totalPercentDisplay">0%</span> / 100%
                        </div>
                    </div>
                </div>
                <div class="admin-modal-footer">
                    <button type="button" class="action-btn" onclick="closeModal()">取消</button>
                    <button type="submit" class="action-btn edit"><i class="fas fa-save"></i> 保存</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
        // 香调类型映射
        var noteTypeMap = {};
        <% 
        Dim rsAllNotes
        Set rsAllNotes = ExecuteQuery("SELECT NoteID, NoteType FROM FragranceNotes WHERE IsActive <> 0")
        If Not rsAllNotes Is Nothing Then
            Do While Not rsAllNotes.EOF
        %>
        noteTypeMap['<%= rsAllNotes("NoteID") %>'] = '<%= rsAllNotes("NoteType") %>';
        <% 
                rsAllNotes.MoveNext
            Loop
            rsAllNotes.Close
        End If
        Set rsAllNotes = Nothing
        %>
        
        // 最小比例配置
        var minTopPercent = <%= minTopPercent %>;
        var minMiddlePercent = <%= minMiddlePercent %>;
        var minBasePercent = <%= minBasePercent %>;
        
        // 初始化图表
        var ctx = document.getElementById('formulaChart');
        if (ctx) {
            new Chart(ctx, {
                type: 'doughnut',
                data: {
                    labels: ['启用', '禁用'],
                    datasets: [{
                        data: [<%= activeCount %>, <%= inactiveCount %>],
                        backgroundColor: ['#4caf50', '#f44336'],
                        borderWidth: 0
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: true,
                    cutout: '70%',
                    plugins: {
                        legend: { display: false }
                    }
                }
            });
        }
        
        // 搜索功能
        function doSearch() {
            var keyword = document.getElementById('searchInput').value;
            var type = document.getElementById('filterType').value;
            var review = document.getElementById('filterReview').value;
            var url = 'formula_management.asp?search=' + encodeURIComponent(keyword);
            if (type) url += '&type=' + encodeURIComponent(type);
            if (review) url += '&review=' + encodeURIComponent(review);
            location.href = url;
        }
        
        document.getElementById('searchInput') && document.getElementById('searchInput').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                doSearch();
            }
        });
        
        // 显示新增模态框
        function showAddModal() {
            document.getElementById('modalTitle').innerHTML = '<i class="fas fa-plus"></i> 新增配方';
            document.getElementById('formAction').value = 'add_formula';
            document.getElementById('formulaId').value = '';
            document.getElementById('formulaName').value = '';
            document.getElementById('recipeCode').value = '';
            document.getElementById('productType').value = '';
            document.getElementById('description').value = '';
            document.getElementById('isActive').value = '1';
            document.getElementById('selectedNotes').value = '';
            
            // 重置所有复选框和输入框
            document.querySelectorAll('.note-checkbox-item input[type="checkbox"]').forEach(function(cb) {
                cb.checked = false;
            });
            document.querySelectorAll('.note-percent-input').forEach(function(input) {
                input.value = '';
                input.disabled = true;
            });
            
            calculateRatio();
            document.getElementById('formulaModal').style.display = 'block';
        }
        
        // 显示编辑模态框
        function showEditModal(id, name, recipeCode, productType, description, isActive, recipeNotes) {
            document.getElementById('modalTitle').innerHTML = '<i class="fas fa-edit"></i> 编辑配方';
            document.getElementById('formAction').value = 'edit_formula';
            document.getElementById('formulaId').value = id;
            document.getElementById('formulaName').value = name;
            document.getElementById('recipeCode').value = recipeCode;
            document.getElementById('productType').value = productType;
            document.getElementById('description').value = description;
            document.getElementById('isActive').value = isActive;
            
            // 重置所有复选框和输入框
            document.querySelectorAll('.note-checkbox-item input[type="checkbox"]').forEach(function(cb) {
                cb.checked = false;
            });
            document.querySelectorAll('.note-percent-input').forEach(function(input) {
                input.value = '';
                input.disabled = true;
            });
            
            // 设置已选中的香调
            if (recipeNotes) {
                var notes = recipeNotes.split('|');
                var selectedIds = [];
                
                notes.forEach(function(note) {
                    if (note) {
                        var parts = note.split(':');
                        if (parts.length >= 4) {
                            var noteId = parts[0];
                            var percent = parts[3];
                            selectedIds.push(noteId);
                            
                            var cb = document.getElementById('note_' + noteId);
                            var pctInput = document.getElementById('percent_' + noteId);
                            
                            if (cb) {
                                cb.checked = true;
                            }
                            if (pctInput) {
                                pctInput.disabled = false;
                                pctInput.value = percent;
                            }
                        }
                    }
                });
                
                document.getElementById('selectedNotes').value = selectedIds.join(',');
            }
            
            calculateRatio();
            document.getElementById('formulaModal').style.display = 'block';
        }
        
        // 切换百分比输入框
        function togglePercentInput(checkbox) {
            var noteId = checkbox.value;
            var percentInput = document.getElementById('percent_' + noteId);
            
            if (percentInput) {
                percentInput.disabled = !checkbox.checked;
                if (!checkbox.checked) {
                    percentInput.value = '';
                } else {
                    percentInput.focus();
                }
            }
            
            updateSelectedNotes();
            calculateRatio();
        }
        
        // 更新选中的香调列表
        function updateSelectedNotes() {
            var selected = [];
            document.querySelectorAll('.note-checkbox-item input[type="checkbox"]:checked').forEach(function(cb) {
                selected.push(cb.value);
            });
            document.getElementById('selectedNotes').value = selected.join(',');
        }
        
        // 计算配比
        function calculateRatio() {
            var totalTop = 0, totalMiddle = 0, totalBase = 0, total = 0;
            
            document.querySelectorAll('.note-checkbox-item input[type="checkbox"]:checked').forEach(function(cb) {
                var noteId = cb.value;
                var percentInput = document.getElementById('percent_' + noteId);
                var percent = parseFloat(percentInput.value) || 0;
                var noteType = noteTypeMap[noteId];
                
                total += percent;
                
                if (noteType === '前调') {
                    totalTop += percent;
                } else if (noteType === '中调') {
                    totalMiddle += percent;
                } else if (noteType === '后调') {
                    totalBase += percent;
                }
            });
            
            // 更新显示
            var topDisplay = document.getElementById('topPercentDisplay');
            var middleDisplay = document.getElementById('middlePercentDisplay');
            var baseDisplay = document.getElementById('basePercentDisplay');
            var totalDisplay = document.getElementById('totalPercentDisplay');
            
            topDisplay.textContent = totalTop + '%';
            topDisplay.className = 'ratio-bar-value ' + (totalTop >= minTopPercent ? 'valid' : 'invalid');
            
            middleDisplay.textContent = totalMiddle + '%';
            middleDisplay.className = 'ratio-bar-value ' + (totalMiddle >= minMiddlePercent ? 'valid' : 'invalid');
            
            baseDisplay.textContent = totalBase + '%';
            baseDisplay.className = 'ratio-bar-value ' + (totalBase >= minBasePercent ? 'valid' : 'invalid');
            
            totalDisplay.textContent = total + '%';
            totalDisplay.className = 'ratio-total-value ' + (total === 100 ? 'valid' : 'invalid');
        }
        
        // 表单验证
        function validateForm() {
            updateSelectedNotes();
            
            var selectedNotes = document.getElementById('selectedNotes').value;
            if (!selectedNotes) {
                alert('请至少选择一个香调');
                return false;
            }
            
            var totalTop = 0, totalMiddle = 0, totalBase = 0, total = 0;
            
            document.querySelectorAll('.note-checkbox-item input[type="checkbox"]:checked').forEach(function(cb) {
                var noteId = cb.value;
                var percentInput = document.getElementById('percent_' + noteId);
                var percent = parseFloat(percentInput.value) || 0;
                var noteType = noteTypeMap[noteId];
                
                total += percent;
                
                if (noteType === '前调') {
                    totalTop += percent;
                } else if (noteType === '中调') {
                    totalMiddle += percent;
                } else if (noteType === '后调') {
                    totalBase += percent;
                }
            });
            
            if (totalTop < minTopPercent) {
                alert('前调比例不能低于' + minTopPercent + '%，当前为' + totalTop + '%');
                return false;
            }
            if (totalMiddle < minMiddlePercent) {
                alert('中调比例不能低于' + minMiddlePercent + '%，当前为' + totalMiddle + '%');
                return false;
            }
            if (totalBase < minBasePercent) {
                alert('后调比例不能低于' + minBasePercent + '%，当前为' + totalBase + '%');
                return false;
            }
            if (total !== 100) {
                alert('香调配比总和必须等于100%，当前为' + total + '%');
                return false;
            }
            
            return true;
        }
        
        // 关闭模态框
        function closeModal() {
            document.getElementById('formulaModal').style.display = 'none';
        }
        
        // 点击模态框外部关闭
        window.onclick = function(event) {
            var modal = document.getElementById('formulaModal');
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
