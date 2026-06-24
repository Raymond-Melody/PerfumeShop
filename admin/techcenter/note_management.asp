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

' ========== SafeNum/SafeDiv 函数 ==========
Function SafeNum(val)
    On Error Resume Next
    If IsNull(val) Or val = "" Then
        SafeNum = 0
    ElseIf IsNumeric(val) Then
        SafeNum = CDbl(val)
    Else
        SafeNum = 0
    End If
    On Error GoTo 0
End Function

Function SafeDiv(numerator, denominator)
    On Error Resume Next
    If IsNull(denominator) Or denominator = "" Then
        SafeDiv = 0
    ElseIf Not IsNumeric(denominator) Then
        SafeDiv = 0
    ElseIf CDbl(denominator) = 0 Then
        SafeDiv = 0
    Else
        SafeDiv = CDbl(numerator) / CDbl(denominator)
    End If
    On Error GoTo 0
End Function

' ========== IIF 函数 ==========
Function IIF(cond, tVal, fVal)
    If cond Then IIF = tVal Else IIF = fVal
End Function

' ========== 处理表单提交 ==========
Dim action, noteId, noteName, noteType, priceAddition, recommendedPercentage, isActive, imageURL, isBaseNote, baseNoteIDs
action = Request.Form("action")

If action = "add" Or action = "edit" Then
    noteName = SafeSQL(Request.Form("noteName"))
    noteType = SafeSQL(Request.Form("noteType"))
    priceAddition = Request.Form("priceAddition")
    If priceAddition = "" Or Not IsNumeric(priceAddition) Then priceAddition = 0
    recommendedPercentage = Request.Form("recommendedPercentage")
    If recommendedPercentage = "" Or Not IsNumeric(recommendedPercentage) Then recommendedPercentage = 0
    isActive = Request.Form("isActive")
    If isActive = "" Then isActive = 1
    imageURL = SafeSQL(Request.Form("imageURL"))
    If imageURL = "" Then imageURL = "/images/default-note.jpg"
    isBaseNote = Request.Form("isBaseNote")
    If isBaseNote = "" Then isBaseNote = 0
    ' 获取多选的基香ID数组
    baseNoteIDs = Request.Form("baseNoteSelect")
    
    ' ===== 基香关联必选验证 + 配比100%验证 =====
    If baseNoteIDs = "" Then
        Response.Redirect "note_management.asp?msg=" & Server.URLEncode("错误：必须关联至少一个基香") & "&msgtype=error"
        Response.End
    End If
    
    Dim valBaseArr, valTotalPct, valBId
    valTotalPct = 0
    valBaseArr = Split(baseNoteIDs, ",")
    For Each valBId In valBaseArr
        If IsNumeric(Trim(valBId)) Then
            valTotalPct = valTotalPct + SafeNum(Request.Form("baseNotePct_" & Trim(valBId)))
        End If
    Next
    If Abs(valTotalPct - 100) > 0.01 Then
        Response.Redirect "note_management.asp?msg=" & Server.URLEncode("错误：基香配比总和必须为100%，当前为" & FormatNumber(valTotalPct, 2) & "%") & "&msgtype=error"
        Response.End
    End If
    ' ===== 验证结束 =====
    
    If action = "add" Then
        Dim addSql, newNoteId
        addSql = "INSERT INTO FragranceNotes (NoteName, NoteType, PriceAddition, RecommendedPercentage, IsActive, ImageURL, IsBaseNote) VALUES ('" & _
                 noteName & "', '" & noteType & "', " & SafeNum(priceAddition) & ", " & CInt(recommendedPercentage) & ", " & CInt(isActive) & ", '" & imageURL & "', " & CInt(isBaseNote) & ")"
        If ExecuteNonQuery(addSql) Then
            ' 获取新插入的香调ID
            Dim rsNewNoteId
            Set rsNewNoteId = ExecuteQuery("SELECT SCOPE_IDENTITY()")
            If Not rsNewNoteId Is Nothing Then
                newNoteId = rsNewNoteId(0)
                rsNewNoteId.Close
                Set rsNewNoteId = Nothing
                
                ' 保存基香关联（V9：带百分比）
                If baseNoteIDs <> "" Then
                    Dim baseNoteArr, bId, bPct
                    baseNoteArr = Split(baseNoteIDs, ",")
                    For Each bId In baseNoteArr
                        If IsNumeric(bId) Then
                            bPct = SafeNum(Request.Form("baseNotePct_" & Trim(bId)))
                            ExecuteNonQuery "INSERT INTO NoteIngredients (NoteID, BaseNoteID, Percentage) VALUES (" & newNoteId & ", " & CLng(bId) & ", " & bPct & ")"
                        End If
                    Next
                End If
            End If
            Response.Redirect "note_management.asp?msg=" & Server.URLEncode("添加成功")
        Else
            Response.Write "<script>alert('添加失败：" & Replace(Session("LastDBError"), "'", "\'") & "');</script>"
        End If
    ElseIf action = "edit" Then
        noteId = Request.Form("noteId")
        Dim editSql
        editSql = "UPDATE FragranceNotes SET NoteName = '" & noteName & "', NoteType = '" & noteType & "', " & _
                  "PriceAddition = " & SafeNum(priceAddition) & ", RecommendedPercentage = " & CInt(recommendedPercentage) & ", " & _
                  "IsActive = " & CInt(isActive) & ", ImageURL = '" & imageURL & "', IsBaseNote = " & CInt(isBaseNote) & " WHERE NoteID = " & CInt(noteId)
        If ExecuteNonQuery(editSql) Then
            ' 删除旧的基香关联
            ExecuteNonQuery "DELETE FROM NoteIngredients WHERE NoteID = " & CInt(noteId)
            
            ' 保存新的基香关联（V9：带百分比）
            If baseNoteIDs <> "" Then
                Dim baseNoteArrEdit, bIdEdit, bPctEdit
                baseNoteArrEdit = Split(baseNoteIDs, ",")
                For Each bIdEdit In baseNoteArrEdit
                    If IsNumeric(bIdEdit) Then
                        bPctEdit = SafeNum(Request.Form("baseNotePct_" & Trim(bIdEdit)))
                        ExecuteNonQuery "INSERT INTO NoteIngredients (NoteID, BaseNoteID, Percentage) VALUES (" & CInt(noteId) & ", " & CLng(bIdEdit) & ", " & bPctEdit & ")"
                    End If
                Next
            End If
            Response.Redirect "note_management.asp?msg=" & Server.URLEncode("更新成功")
        Else
            Response.Write "<script>alert('更新失败');</script>"
        End If
    End If
ElseIf action = "delete" Then
    noteId = Request.Form("noteId")
    ' 软删除 - 仅TECH_MANAGER可操作
    If isManager Then
        ' 先删除关联的基香记录
        ExecuteNonQuery "DELETE FROM NoteIngredients WHERE NoteID = " & CInt(noteId)
        Dim deleteSql
        deleteSql = "UPDATE FragranceNotes SET IsActive = 0 WHERE NoteID = " & CInt(noteId)
        If ExecuteNonQuery(deleteSql) Then
            Response.Redirect "note_management.asp?msg=" & Server.URLEncode("已禁用")
        Else
            Response.Write "<script>alert('操作失败');</script>"
        End If
    Else
        Response.Write "<script>alert('权限不足');</script>"
    End If
ElseIf action = "restore" Then
    noteId = Request.Form("noteId")
    ' 恢复 - 仅TECH_MANAGER可操作
    If isManager Then
        Dim restoreSql
        restoreSql = "UPDATE FragranceNotes SET IsActive = 1 WHERE NoteID = " & CInt(noteId)
        If ExecuteNonQuery(restoreSql) Then
            Response.Redirect "note_management.asp?msg=" & Server.URLEncode("已恢复")
        Else
            Response.Write "<script>alert('操作失败');</script>"
        End If
    Else
        Response.Write "<script>alert('权限不足');</script>"
    End If
End If

' ========== 获取筛选参数 ==========
Dim filterType, searchKeyword
filterType = Request.QueryString("type")
searchKeyword = SafeSQL(Request.QueryString("search"))

' ========== 构建查询条件 ==========
Dim whereClause
whereClause = ""

If filterType <> "" And filterType <> "all" Then
    If whereClause <> "" Then whereClause = whereClause & " AND "
    whereClause = whereClause & "NoteType = '" & filterType & "'"
End If

If searchKeyword <> "" Then
    If whereClause <> "" Then whereClause = whereClause & " AND "
    whereClause = whereClause & "NoteName LIKE '%" & searchKeyword & "%'"
End If

If whereClause <> "" Then
    whereClause = "WHERE " & whereClause
End If

' ========== 获取香调列表 ==========
Dim rsNotes, sql
sql = "SELECT * FROM FragranceNotes " & whereClause & " ORDER BY NoteType, NoteID DESC"
Set rsNotes = ExecuteQuery(sql)

' ========== 获取基香列表（用于关联选择） ==========
Dim rsBaseNotes
Set rsBaseNotes = ExecuteQuery("SELECT BaseNoteID, BaseNoteName FROM BaseNotes WHERE IsActive <> 0 ORDER BY BaseNoteName")

' ========== 获取统计数据 ==========
Dim totalCount, topCount, middleCount, baseCount
totalCount = 0
topCount = 0
middleCount = 0
baseCount = 0

Dim rsStats
Set rsStats = ExecuteQuery("SELECT NoteType, COUNT(*) AS CountNum FROM FragranceNotes GROUP BY NoteType")
If Not rsStats Is Nothing Then
    Do While Not rsStats.EOF
        If rsStats("NoteType") = "前调" Then topCount = rsStats("CountNum")
        If rsStats("NoteType") = "中调" Then middleCount = rsStats("CountNum")
        If rsStats("NoteType") = "后调" Then baseCount = rsStats("CountNum")
        rsStats.MoveNext
    Loop
    rsStats.Close
    Set rsStats = Nothing
End If
totalCount = topCount + middleCount + baseCount
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>香调管理 - 产品技术管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        /* 暗色主题基础 */
        body {
            background: #1a1a2e;
            color: #e0e0e0;
        }
        
        /* 页面头部 */
        .page-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 25px;
            padding-bottom: 15px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
        }
        .page-title {
            font-size: 24px;
            color: #fff;
            margin: 0;
        }
        .page-title i {
            color: #00bcd4;
            margin-right: 10px;
        }
        .breadcrumb {
            font-size: 14px;
            color: #888;
        }
        .breadcrumb a {
            color: #00bcd4;
            text-decoration: none;
        }
        .breadcrumb a:hover {
            text-decoration: underline;
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
            border: 1px solid rgba(255,255,255,0.05);
        }
        .filter-group {
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .filter-group.right {
            margin-left: auto;
        }
        .filter-label {
            color: #888;
            font-size: 13px;
        }
        .filter-tabs {
            display: flex;
            gap: 5px;
        }
        .filter-tab {
            padding: 8px 16px;
            border-radius: 6px;
            background: rgba(255,255,255,0.05);
            color: #b0b0b0;
            text-decoration: none;
            font-size: 13px;
            transition: all 0.2s ease;
            border: 1px solid transparent;
        }
        .filter-tab:hover {
            background: rgba(255,255,255,0.1);
            color: #fff;
        }
        .filter-tab.active {
            background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%);
            color: #fff;
            border-color: rgba(0,188,212,0.3);
        }
        .search-box {
            display: flex;
            gap: 10px;
        }
        .filter-input, .search-input {
            background: rgba(0,0,0,0.3);
            border: 1px solid #3a3a5a;
            border-radius: 6px;
            padding: 8px 12px;
            color: #fff;
            font-size: 13px;
            width: 200px;
            box-sizing: border-box;
        }
        .filter-input:focus, .search-input:focus {
            outline: none;
            border-color: #00bcd4;
        }
        .filter-input::placeholder, .search-input::placeholder {
            color: #999;
        }
        /* filter-btn/admin-btn/action-btn 样式已迁移至 /css/buttons.css 统一管理系统 */
        
        /* 卡片样式 */
        .admin-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            border: 1px solid rgba(255,255,255,0.05);
            margin-bottom: 25px;
        }
        .admin-card-header {
            padding: 20px;
            border-bottom: 1px solid rgba(255,255,255,0.05);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .admin-card-title {
            font-size: 18px;
            color: #fff;
            margin: 0;
        }
        .admin-card-body {
            padding: 20px;
        }
        /* admin-btn 样式由 /css/buttons.css 兼容映射统一管理 */
        
        /* 卡片网格布局 */
        .note-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(340px, 1fr));
            gap: 20px;
        }
        .note-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 20px;
            border: 1px solid rgba(255,255,255,0.05);
            transition: all 0.3s ease;
        }
        .note-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.3);
            border-color: rgba(0,188,212,0.2);
        }
        .note-header {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            margin-bottom: 12px;
        }
        .note-title {
            font-size: 16px;
            font-weight: 600;
            color: #fff;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .note-title i { color: #00bcd4; }
        .note-id {
            font-size: 11px;
            color: #999;
            background: rgba(0,0,0,0.3);
            padding: 2px 8px;
            border-radius: 4px;
        }
        .note-desc {
            color: #888;
            font-size: 13px;
            margin-bottom: 15px;
            line-height: 1.5;
            min-height: 20px;
        }
        .note-desc.empty { color: #999; font-style: italic; }
        .note-info-row {
            display: flex;
            gap: 15px;
            margin-bottom: 15px;
            flex-wrap: wrap;
        }
        .note-info-item {
            display: flex;
            align-items: center;
            gap: 6px;
            font-size: 13px;
            color: #b0b0b0;
        }
        .note-info-item i {
            color: #999;
            font-size: 12px;
        }
        .note-ingredients {
            margin-bottom: 15px;
            padding: 10px;
            background: rgba(0,0,0,0.2);
            border-radius: 6px;
        }
        .note-ingredients-label {
            font-size: 11px;
            color: #999;
            margin-bottom: 6px;
            display: flex;
            align-items: center;
            gap: 5px;
        }
        .note-ingredients-list {
            display: flex;
            flex-wrap: wrap;
            gap: 5px;
        }
        .note-ingredient-tag {
            background: rgba(0,188,212,0.1);
            color: #00bcd4;
            padding: 3px 8px;
            border-radius: 3px;
            font-size: 11px;
            border: 1px solid rgba(0,188,212,0.2);
        }
        .note-ingredients-empty { color: #999; font-size: 11px; font-style: italic; }
        .note-footer {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding-top: 15px;
            border-top: 1px solid rgba(255,255,255,0.05);
        }
        
        /* 表格样式（保留用于兼容性） */
        .admin-table {
            width: 100%;
            border-collapse: collapse;
        }
        .admin-table th {
            background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%);
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 500;
            font-size: 13px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .admin-table td {
            padding: 15px;
            border-bottom: 1px solid rgba(255,255,255,0.05);
            color: #e0e0e0;
        }
        .admin-table tr:hover td {
            background: rgba(255,255,255,0.02);
        }
        .admin-table tr:last-child td {
            border-bottom: none;
        }
        
        /* 状态标签 */
        .status-badge {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 4px;
            font-size: 11px;
        }
        .status-badge.active { background: rgba(76, 175, 80, 0.2); color: #4caf50; }
        .status-badge.inactive { background: rgba(244, 67, 54, 0.2); color: #f44336; }
        .status-badge.top { background: rgba(255, 193, 7, 0.15); color: #ffc107; border: 1px solid rgba(255,193,7,0.3); }
        .status-badge.middle { background: rgba(156, 39, 176, 0.15); color: #ce93d8; border: 1px solid rgba(156,39,176,0.3); }
        .status-badge.base { background: rgba(0, 150, 136, 0.15); color: #80cbc4; border: 1px solid rgba(0,150,136,0.3); }
        
        /* 操作按钮组 */
        .action-btns {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
        }
        /* action-btn 样式由 /css/buttons.css 统一管理 */
        
        /* 提示消息 */
        .alert {
            padding: 15px 20px;
            border-radius: 8px;
            margin-bottom: 20px;
        }
        .alert-success {
            background: rgba(76,175,80,0.1);
            color: #4caf50;
            border-left: 4px solid #4caf50;
        }
        .alert-error {
            background: rgba(244,67,54,0.1);
            color: #f44336;
            border-left: 4px solid #f44336;
        }
        
        /* 模态框样式 */
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
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
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
        .admin-modal-title {
            font-size: 18px;
            font-weight: 600;
            color: #fff;
            margin: 0;
        }
        .admin-modal-close {
            background: none;
            border: none;
            color: #bbb;
            font-size: 24px;
            cursor: pointer;
        }
        .admin-modal-close:hover {
            color: #fff;
        }
        .admin-modal-body {
            padding: 20px;
            max-height: 70vh;
            overflow-y: auto;
        }
        .admin-modal-footer {
            padding: 15px 20px;
            border-top: 1px solid #3a3a5a;
            display: flex;
            justify-content: flex-end;
            gap: 10px;
        }
        
        /* 表单样式 */
        .form-group, .admin-form-group {
            margin-bottom: 15px;
        }
        .form-label, .admin-form-label {
            display: block;
            margin-bottom: 8px;
            color: #e0e0e0;
            font-size: 13px;
        }
        .admin-modal-content .admin-form-label {
            color: #e0e0e0 !important;
        }
        .form-label .required { color: #f44336; }
        .form-control, .admin-form-control {
            width: 100%;
            background: rgba(0,0,0,0.3);
            border: 1px solid #3a3a5a;
            border-radius: 6px;
            padding: 10px 12px;
            color: #fff;
            font-size: 14px;
            box-sizing: border-box;
        }
        .form-control:focus, .admin-form-control:focus {
            outline: none;
            border-color: #00bcd4;
        }
        textarea.form-control, textarea.admin-form-control { resize: vertical; min-height: 60px; }
        .form-row, .admin-form-row {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 15px;
        }
        .form-col, .admin-form-col {
            flex: 1;
        }
        select.form-control option, select.admin-form-control option {
            background: #2d2d44;
            color: #fff;
        }
        .admin-form-control::placeholder { color: #999; }
        
        /* 基香复选框组样式 */
        .base-note-checkbox-group {
            max-height: 150px;
            overflow-y: auto;
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 8px;
            padding: 12px;
            background: rgba(255,255,255,0.03);
        }
        .base-note-checkbox-item {
            display: flex;
            align-items: center;
            padding: 6px 8px;
            margin-bottom: 4px;
            border-radius: 4px;
            cursor: pointer;
            transition: background 0.2s;
        }
        .base-note-checkbox-item:hover {
            background: rgba(255,255,255,0.05);
        }
        .base-note-checkbox-item input[type="checkbox"] {
            margin-right: 10px;
            width: 16px;
            height: 16px;
            cursor: pointer;
            accent-color: #00bcd4;
        }
        .base-note-checkbox-item span {
            color: #e0e0e0;
            font-size: 14px;
        }
        
        /* 空状态 */
        .empty-state {
            text-align: center;
            padding: 60px 20px;
            color: #999;
            grid-column: 1 / -1;
        }
        .empty-state i {
            font-size: 64px;
            margin-bottom: 20px;
            color: #888;
        }
        
        /* 统计区域样式 */
        .stats-section {
            display: grid;
            grid-template-columns: 2fr 1fr;
            gap: 20px;
            margin-bottom: 20px;
        }
        .stats-cards {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
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
        .stat-card.top .stat-value { color: #FF9800; }
        .stat-card.middle .stat-value { color: #4CAF50; }
        .stat-card.base .stat-value { color: #2196F3; }
        
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
        
        /* 响应式 */
        @media (max-width: 1200px) {
            .stats-section { grid-template-columns: 1fr; }
            .chart-container { display: none; }
        }
        @media (max-width: 992px) {
            .stats-cards { grid-template-columns: repeat(2, 1fr); }
        }
        @media (max-width: 768px) {
            .stats-cards { grid-template-columns: 1fr; }
            .note-grid { grid-template-columns: 1fr; }
            .filter-bar { flex-direction: column; align-items: stretch; }
            .filter-group.right { margin-left: 0; }
            .filter-tabs { flex-wrap: wrap; }
            .search-box { width: 100%; }
            .search-input { flex: 1; width: auto; }
            .form-row, .admin-form-row { grid-template-columns: 1fr; }
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
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <div>
                <h2 class="page-title"><i class="fas fa-leaf"></i> 香调管理</h2>
                <div class="breadcrumb">
                    <a href="index.asp">技术中心</a> / <span>香调管理</span>
                </div>
            </div>
        </div>
        
        <% If Request.QueryString("msg") <> "" Then 
            Dim alertType, alertIcon
            If Request.QueryString("msgtype") = "error" Then
                alertType = "alert-error"
                alertIcon = "fa-exclamation-circle"
            Else
                alertType = "alert-success"
                alertIcon = "fa-check-circle"
            End If
        %>
        <div class="alert <%= alertType %>">
            <i class="fas <%= alertIcon %>"></i>
            <%= Server.HTMLEncode(Request.QueryString("msg")) %>
        </div>
        <% End If %>
        
        <!-- 统计区域 -->
        <div class="stats-section">
            <div class="stats-cards">
                <div class="stat-card total <%= IIf(filterType = "" Or filterType = "all", "active", "") %>" onclick="location.href='note_management.asp'">
                    <div class="stat-value"><%= totalCount %></div>
                    <div class="stat-label">香调总数</div>
                </div>
                <div class="stat-card top <%= IIf(filterType = "前调", "active", "") %>" onclick="location.href='note_management.asp?type=前调'">
                    <div class="stat-value"><%= topCount %></div>
                    <div class="stat-label">前调</div>
                </div>
                <div class="stat-card middle <%= IIf(filterType = "中调", "active", "") %>" onclick="location.href='note_management.asp?type=中调'">
                    <div class="stat-value"><%= middleCount %></div>
                    <div class="stat-label">中调</div>
                </div>
                <div class="stat-card base <%= IIf(filterType = "后调", "active", "") %>" onclick="location.href='note_management.asp?type=后调'">
                    <div class="stat-value"><%= baseCount %></div>
                    <div class="stat-label">后调</div>
                </div>
            </div>
            <div class="chart-container">
                <div class="chart-wrapper">
                    <canvas id="noteTypeChart"></canvas>
                    <div class="chart-center-text">
                        <div class="chart-center-value"><%= totalCount %></div>
                        <div class="chart-center-label">香调总数</div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- 筛选栏 -->
        <div class="filter-bar">
            <div class="filter-group">
                <span class="filter-label"><i class="fas fa-filter"></i> 类型筛选：</span>
                <div class="filter-tabs">
                    <a href="note_management.asp<%= IIf(searchKeyword <> "", "?search=" & Server.URLEncode(searchKeyword), "") %>" class="filter-tab <%= IIf(filterType = "" Or filterType = "all", "active", "") %>">全部</a>
                    <a href="note_management.asp?type=前调<%= IIf(searchKeyword <> "", "&search=" & Server.URLEncode(searchKeyword), "") %>" class="filter-tab <%= IIf(filterType = "前调", "active", "") %>">前调</a>
                    <a href="note_management.asp?type=中调<%= IIf(searchKeyword <> "", "&search=" & Server.URLEncode(searchKeyword), "") %>" class="filter-tab <%= IIf(filterType = "中调", "active", "") %>">中调</a>
                    <a href="note_management.asp?type=后调<%= IIf(searchKeyword <> "", "&search=" & Server.URLEncode(searchKeyword), "") %>" class="filter-tab <%= IIf(filterType = "后调", "active", "") %>">后调</a>
                </div>
            </div>
            <div class="filter-group right">
                <form method="get" class="search-box">
                    <% If filterType <> "" And filterType <> "all" Then %>
                    <input type="hidden" name="type" value="<%= filterType %>">
                    <% End If %>
                    <input type="text" name="search" class="search-input" placeholder="搜索香调名称..." value="<%= HTMLEncode(searchKeyword) %>">
                    <button type="submit" class="filter-btn">
                        <i class="fas fa-search"></i> 搜索
                    </button>
                    <% If searchKeyword <> "" Then %>
                    <a href="note_management.asp<%= IIf(filterType <> "" And filterType <> "all", "?type=" & filterType, "") %>" class="filter-btn" style="background: #666;">
                        <i class="fas fa-times"></i> 清除
                    </a>
                    <% End If %>
                </form>
            </div>
            <button class="filter-btn secondary" onclick="showAddForm()">
                <i class="fas fa-plus"></i> 新增香调
            </button>
        </div>
        
        <!-- 香调卡片列表 -->
        <div class="note-grid">
            <% 
            Dim hasRecords
            hasRecords = False
            If Not rsNotes Is Nothing Then 
                If Not rsNotes.EOF Then
                    hasRecords = True
                End If
            End If
            
            If hasRecords Then 
                Do While Not rsNotes.EOF 
            %>
            <div class="note-card">
                <div class="note-header">
                    <div class="note-title">
                        <i class="fas fa-leaf"></i>
                        <%= HTMLEncode(rsNotes("NoteName")) %>
                    </div>
                    <span class="note-id">#<%= rsNotes("NoteID") %></span>
                </div>
                
                <div class="note-info-row">
                    <div class="note-info-item">
                        <% 
                        Dim noteTypeClass
                        noteTypeClass = ""
                        If rsNotes("NoteType") = "前调" Then
                            noteTypeClass = "top"
                        ElseIf rsNotes("NoteType") = "中调" Then
                            noteTypeClass = "middle"
                        ElseIf rsNotes("NoteType") = "后调" Then
                            noteTypeClass = "base"
                        End If
                        %>
                        <span class="status-badge <%= noteTypeClass %>">
                            <%= HTMLEncode(rsNotes("NoteType")) %>
                        </span>
                    </div>
                    <div class="note-info-item">
                        <i class="fas fa-percentage"></i>
                        <% 
                        Dim recPercent
                        recPercent = 0
                        On Error Resume Next
                        recPercent = rsNotes("RecommendedPercentage")
                        If Err.Number <> 0 Then recPercent = 0
                        On Error GoTo 0
                        Response.Write recPercent & "%"
                        %>
                    </div>
                    <div class="note-info-item">
                        <i class="fas fa-yen-sign"></i>
                        ¥<%= FormatNumber(SafeNum(rsNotes("PriceAddition")), 2) %>
                    </div>
                    <div class="note-info-item">
                        <i class="fas fa-layer-group"></i>
                        <% 
                        Dim isBaseNoteVal
                        isBaseNoteVal = 0
                        On Error Resume Next
                        isBaseNoteVal = rsNotes("IsBaseNote")
                        If Err.Number <> 0 Then isBaseNoteVal = 0
                        On Error GoTo 0
                        
                        If isBaseNoteVal <> 0 Then
                            Response.Write "基香"
                        Else
                            Response.Write "非基香"
                        End If
                        %>
                    </div>
                </div>
                
                <!-- 关联基香 -->
                <div class="note-ingredients">
                    <div class="note-ingredients-label"><i class="fas fa-vial"></i> 关联基香</div>
                    <div class="note-ingredients-list">
                        <% 
                        Dim rsNoteBaseNames
                        Set rsNoteBaseNames = ExecuteQuery("SELECT bn.BaseNoteName, ni.Percentage FROM NoteIngredients ni INNER JOIN BaseNotes bn ON ni.BaseNoteID = bn.BaseNoteID WHERE ni.NoteID = " & rsNotes("NoteID"))
                        If Not rsNoteBaseNames Is Nothing Then
                            Dim hasIngredients
                            hasIngredients = False
                            Do While Not rsNoteBaseNames.EOF
                                hasIngredients = True
                        %>
                        <span class="note-ingredient-tag"><%= HTMLEncode(rsNoteBaseNames("BaseNoteName")) %> (<%= SafeNum(rsNoteBaseNames("Percentage")) %>%)</span>
                        <% 
                                rsNoteBaseNames.MoveNext
                            Loop
                            rsNoteBaseNames.Close
                            Set rsNoteBaseNames = Nothing
                            If Not hasIngredients Then
                        %>
                        <span class="note-ingredients-empty">暂无关联基香</span>
                        <% 
                            End If
                        Else
                        %>
                        <span class="note-ingredients-empty">暂无关联基香</span>
                        <% End If %>
                    </div>
                </div>
                
                <div class="note-footer">
                    <% 
                    Dim isActiveVal
                    isActiveVal = 1
                    On Error Resume Next
                    isActiveVal = rsNotes("IsActive")
                    If Err.Number <> 0 Then isActiveVal = 1
                    On Error GoTo 0
                    
                    If isActiveVal <> 0 Then
                    %>
                    <span class="status-badge active">启用</span>
                    <% Else %>
                    <span class="status-badge inactive">禁用</span>
                    <% End If %>
                    <div class="action-btns">
                        <button class="action-btn edit" onclick="showEditForm(this)"
                            data-id="<%= rsNotes("NoteID") %>" 
                            data-name="<%= SafeOutput(rsNotes("NoteName")) %>" 
                            data-type="<%= rsNotes("NoteType") %>"
                            data-price="<%= SafeNum(rsNotes("PriceAddition")) %>" 
                            data-percent="<% On Error Resume Next: Response.Write rsNotes("RecommendedPercentage"): If Err.Number <> 0 Then Response.Write "0": Err.Clear: On Error GoTo 0 %>"
                            data-active="<% On Error Resume Next: Response.Write rsNotes("IsActive"): If Err.Number <> 0 Then Response.Write "1": Err.Clear: On Error GoTo 0 %>"
                            data-image="<%= SafeOutput(rsNotes("ImageURL") & "") %>"
                            data-isbasenote="<% On Error Resume Next: Response.Write rsNotes("IsBaseNote"): If Err.Number <> 0 Then Response.Write "0": Err.Clear: On Error GoTo 0 %>"
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
                            %>"
                            data-basenotepcts="<%
                                ' V9: 查询关联基香的百分比
                                Dim rsNoteBasePcts
                                Set rsNoteBasePcts = ExecuteQuery("SELECT BaseNoteID, Percentage FROM NoteIngredients WHERE NoteID = " & rsNotes("NoteID"))
                                If Not rsNoteBasePcts Is Nothing Then
                                    Dim pctList, firstPct
                                    pctList = ""
                                    firstPct = True
                                    Do While Not rsNoteBasePcts.EOF
                                        If Not firstPct Then pctList = pctList & ","
                                        pctList = pctList & rsNoteBasePcts("BaseNoteID") & ":" & SafeNum(rsNoteBasePcts("Percentage"))
                                        firstPct = False
                                        rsNoteBasePcts.MoveNext
                                    Loop
                                    rsNoteBasePcts.Close
                                    Set rsNoteBasePcts = Nothing
                                    Response.Write pctList
                                End If
                            %>">
                            <i class="fas fa-edit"></i> 编辑
                        </button>
                        <% If isManager Then %>
                            <% If isActiveVal <> 0 Then %>
                            <form method="post" style="display:inline;" onsubmit="return confirm('确定要禁用此香调吗？')">
                                <input type="hidden" name="action" value="delete">
                                <input type="hidden" name="noteId" value="<%= rsNotes("NoteID") %>">
                                <button type="submit" class="action-btn reject">
                                    <i class="fas fa-ban"></i> 禁用
                                </button>
                            </form>
                            <% Else %>
                            <form method="post" style="display:inline;" onsubmit="return confirm('确定要恢复此香调吗？')">
                                <input type="hidden" name="action" value="restore">
                                <input type="hidden" name="noteId" value="<%= rsNotes("NoteID") %>">
                                <button type="submit" class="action-btn approve">
                                    <i class="fas fa-undo"></i> 恢复
                                </button>
                            </form>
                            <% End If %>
                        <% End If %>
                    </div>
                </div>
            </div>
            <% rsNotes.MoveNext %>
            <% Loop %>
            <% Else %>
            <div class="empty-state">
                <i class="fas fa-leaf"></i>
                <h3>暂无香调数据</h3>
                <p>点击"新增香调"按钮创建第一个香调</p>
            </div>
            <% End If %>
        </div>
    </div>
    
    <!-- 添加/编辑香调模态框 -->
    <div id="noteModal" class="admin-modal">
        <div class="admin-modal-content">
            <div class="admin-modal-header">
                <h3 id="modalTitle" class="admin-modal-title">新增香调</h3>
                <button class="admin-modal-close" onclick="closeModal()">&times;</button>
            </div>
            <form id="noteForm" method="post" onsubmit="return validateBaseNotes()">
                <div class="admin-modal-body">
                    <input type="hidden" id="formAction" name="action" value="add">
                    <input type="hidden" id="noteId" name="noteId" value="">
                    <input type="hidden" name="csrf_token" value="<%= Session("CSRFToken") %>">
                    
                    <div class="admin-form-group">
                        <label for="noteName" class="admin-form-label">香调名称 *</label>
                        <input type="text" id="noteName" name="noteName" class="admin-form-control" required placeholder="请输入香调名称">
                    </div>
                    
                    <div class="admin-form-row">
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label for="noteType" class="admin-form-label">类型 *</label>
                                <select id="noteType" name="noteType" class="admin-form-control" required>
                                    <option value="前调">前调 (Top Note)</option>
                                    <option value="中调">中调 (Middle Note)</option>
                                    <option value="后调">后调 (Base Note)</option>
                                </select>
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
                    
                    <div class="admin-form-row">
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label for="recommendedPercentage" class="admin-form-label">推荐百分比 (%)</label>
                                <input type="number" id="recommendedPercentage" name="recommendedPercentage" min="0" max="100" class="admin-form-control" value="0" placeholder="建议配比">
                            </div>
                        </div>
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label for="priceAddition" class="admin-form-label">价格加成 (¥)</label>
                                <input type="number" id="priceAddition" name="priceAddition" step="0.01" min="0" class="admin-form-control" value="0" placeholder="附加价格">
                            </div>
                        </div>
                    </div>
                    
                    <div class="admin-form-row">
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label for="isBaseNote" class="admin-form-label">是否基香</label>
                                <select id="isBaseNote" name="isBaseNote" class="admin-form-control">
                                    <option value="0">否</option>
                                    <option value="1">是</option>
                                </select>
                            </div>
                        </div>
                        <div class="admin-form-col">
                            <div class="admin-form-group" id="baseNoteField">
                                <label class="admin-form-label">关联基香 <small style="color:#f44336;">(必选，配比总和需100%)</small></label>
                                <div class="base-note-checkbox-group">
                                    <%
                                    Dim rsBaseNotesForNote
                                    Set rsBaseNotesForNote = ExecuteQuery("SELECT BaseNoteID, BaseNoteName FROM BaseNotes WHERE IsActive <> 0 ORDER BY BaseNoteName")
                                    If Not rsBaseNotesForNote Is Nothing Then
                                        Do While Not rsBaseNotesForNote.EOF
                                    %>
                                    <label class="base-note-checkbox-item">
                                        <input type="checkbox" name="baseNoteSelect" value="<%= rsBaseNotesForNote("BaseNoteID") %>" onclick="togglePctInput(this)">
                                        <span><%= HTMLEncode(rsBaseNotesForNote("BaseNoteName")) %></span>
                                        <input type="number" name="baseNotePct_<%= rsBaseNotesForNote("BaseNoteID") %>" class="base-note-pct" min="0" max="100" step="0.01" placeholder="%" disabled value="0" oninput="updatePctSummary()" style="width:60px;margin-left:6px;padding:3px 6px;background:#252538;border:1px solid #3a3a3a;color:#e0e0e0;border-radius:4px;font-size:12px;">
                                    </label>
                                    <%
                                            rsBaseNotesForNote.MoveNext
                                        Loop
                                        rsBaseNotesForNote.Close
                                        Set rsBaseNotesForNote = Nothing
                                    End If
                                    %>
                                </div>
                                <small style="color: #bbb; display: block; margin-top: 5px;">选择该香调包含的基香成分（必选，配比总和必须等于100%）</small>
                                <div class="base-note-pct-summary" id="baseNotePctSummary" style="margin-top:8px;padding:8px 12px;border-radius:6px;font-size:13px;display:none;"></div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="admin-form-group">
                        <label class="admin-form-label">图片</label>
                        <div class="image-upload-wrapper">
                            <div class="image-preview" id="imagePreview_note">
                                <img id="previewImg_note" src="" alt="预览" style="display:none;">
                                <div class="image-placeholder" id="placeholder_note">
                                    <i class="fas fa-cloud-upload-alt"></i>
                                    <span>点击上传或拖拽图片</span>
                                </div>
                            </div>
                            <input type="file" id="fileInput_note" accept="image/jpeg,image/png,image/gif,image/webp,image/svg+xml" style="display:none;">
                            <div class="image-upload-actions">
                                <button type="button" class="admin-btn admin-btn-info btn-sm" onclick="document.getElementById('fileInput_note').click();">
                                    <i class="fas fa-upload"></i> 选择图片
                                </button>
                                <button type="button" class="admin-btn admin-btn-secondary btn-sm" onclick="toggleUrlInput_note()">
                                    <i class="fas fa-link"></i> 输入URL
                                </button>
                            </div>
                            <div id="urlInputWrapper_note" style="display:none; margin-top:8px;">
                                <input type="text" id="manualUrl_note" class="admin-form-control" placeholder="输入图片URL地址" style="font-size:13px;">
                                <button type="button" class="admin-btn admin-btn-secondary btn-sm" onclick="applyManualUrl_note()" style="margin-top:4px;">确认</button>
                            </div>
                            <div class="upload-progress" id="uploadProgress_note" style="display:none;">
                                <div class="progress-bar" id="progressBar_note"></div>
                                <span class="progress-text" id="progressText_note">上传中...</span>
                            </div>
                            <div style="font-size:11px;color:#888;margin-top:6px;">如果原图超过 180KB，将自动压缩后再上传</div>
                            <input type="hidden" name="imageURL" id="imageURL_note" value="/images/default-note.jpg">
                        </div>
                    </div>
                </div>
                <div class="admin-modal-footer">
                    <button type="button" class="action-btn" onclick="closeModal()">取消</button>
                    <button type="submit" class="action-btn edit">
                        <i class="fas fa-save"></i> 保存
                    </button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
        // 初始化 Doughnut 图表
        document.addEventListener('DOMContentLoaded', function() {
            var ctx = document.getElementById('noteTypeChart').getContext('2d');
            var topCount = <%= topCount %>;
            var middleCount = <%= middleCount %>;
            var baseCount = <%= baseCount %>;
            
            // 如果没有数据，显示占位
            var dataValues = [topCount, middleCount, baseCount];
            var hasData = dataValues.some(function(v) { return v > 0; });
            
            if (!hasData) {
                dataValues = [1, 1, 1]; // 占位显示
            }
            
            new Chart(ctx, {
                type: 'doughnut',
                data: {
                    labels: ['前调', '中调', '后调'],
                    datasets: [{
                        data: dataValues,
                        backgroundColor: ['#FF9800', '#4CAF50', '#2196F3'],
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
                            backgroundColor: 'rgba(45, 45, 68, 0.95)',
                            titleColor: '#fff',
                            bodyColor: '#e0e0e0',
                            borderColor: 'rgba(255,255,255,0.1)',
                            borderWidth: 1,
                            padding: 12,
                            callbacks: {
                                label: function(context) {
                                    var label = context.label || '';
                                    var value = context.parsed;
                                    var total = context.dataset.data.reduce(function(a, b) { return a + b; }, 0);
                                    var percentage = total > 0 ? Math.round((value / total) * 100) : 0;
                                    return label + ': ' + value + ' (' + percentage + '%)';
                                }
                            }
                        }
                    }
                }
            });
        });
        
        function showAddForm() {
            document.getElementById('modalTitle').textContent = '新增香调';
            document.getElementById('formAction').value = 'add';
            document.getElementById('noteId').value = '';
            document.getElementById('noteName').value = '';
            document.getElementById('noteType').value = '前调';
            document.getElementById('recommendedPercentage').value = '0';
            document.getElementById('priceAddition').value = '0';
            document.getElementById('isActive').value = '1';
            document.getElementById('isBaseNote').value = '0';
            document.getElementById('imageURL_note').value = '/images/default-note.jpg';
            document.getElementById('previewImg_note').style.display = 'none';
            document.getElementById('placeholder_note').style.display = 'flex';
            document.getElementById('fileInput_note').value = '';
            var checkboxes = document.querySelectorAll('input[name="baseNoteSelect"]');
            checkboxes.forEach(function(cb) {
                cb.checked = false;
            });
            // V9: 重置所有百分比输入框
            var pctInputs = document.querySelectorAll('.base-note-pct');
            pctInputs.forEach(function(pi) {
                pi.value = '0';
                pi.disabled = true;
            });
            document.getElementById('noteModal').style.display = 'block';
            updatePctSummary();
        }
        
        function showEditForm(button) {
            var id = button.getAttribute('data-id');
            var name = button.getAttribute('data-name');
            var type = button.getAttribute('data-type');
            var price = button.getAttribute('data-price');
            var percent = button.getAttribute('data-percent');
            var active = button.getAttribute('data-active');
            var image = button.getAttribute('data-image');
            var isBaseNote = button.getAttribute('data-isbasenote');
            var baseNoteIds = button.getAttribute('data-basenoteids') || '';
            var baseNotePcts = button.getAttribute('data-basenotepcts') || '';
            
            document.getElementById('modalTitle').textContent = '编辑香调';
            document.getElementById('formAction').value = 'edit';
            document.getElementById('noteId').value = id;
            document.getElementById('noteName').value = name;
            document.getElementById('noteType').value = type;
            document.getElementById('recommendedPercentage').value = percent || '0';
            document.getElementById('priceAddition').value = price || '0';
            document.getElementById('isActive').value = active || '1';
            document.getElementById('isBaseNote').value = isBaseNote || '0';
            var noteImageVal = image || '/images/default-note.jpg';
            document.getElementById('imageURL_note').value = noteImageVal;
            if (noteImageVal && noteImageVal !== '/images/default-note.jpg') {
                document.getElementById('previewImg_note').src = noteImageVal;
                document.getElementById('previewImg_note').style.display = 'block';
                document.getElementById('placeholder_note').style.display = 'none';
            } else {
                document.getElementById('previewImg_note').style.display = 'none';
                document.getElementById('placeholder_note').style.display = 'flex';
            }
            // 清空所有基香复选框和百分比
            var checkboxes = document.querySelectorAll('input[name="baseNoteSelect"]');
            checkboxes.forEach(function(cb) {
                cb.checked = false;
            });
            // V9: 也重置百分比
            var pctInputs = document.querySelectorAll('.base-note-pct');
            pctInputs.forEach(function(pi) {
                pi.value = '0';
                pi.disabled = true;
            });
            
            if (baseNoteIds) {
                var ids = baseNoteIds.split(',');
                ids.forEach(function(id) {
                    var cb = document.querySelector('input[name="baseNoteSelect"][value="' + id.trim() + '"]');
                    if (cb) {
                        cb.checked = true;
                        // V9: 启用百分比输入框
                        var pctInput = document.querySelector('input[name="baseNotePct_' + id.trim() + '"]');
                        if (pctInput) pctInput.disabled = false;
                    }
                });
            }
            
            // V9: 填充百分比值
            if (baseNotePcts) {
                var pairs = baseNotePcts.split(',');
                pairs.forEach(function(pair) {
                    var parts = pair.split(':');
                    if (parts.length === 2) {
                        var pctInput = document.querySelector('input[name="baseNotePct_' + parts[0].trim() + '"]');
                        if (pctInput) pctInput.value = parseFloat(parts[1]) || 0;
                    }
                });
            }
            
            document.getElementById('noteModal').style.display = 'block';
            updatePctSummary();
        }
        
        function closeModal() {
            document.getElementById('noteModal').style.display = 'none';
        }
        
        // V9: 切换基香百分比输入框的启用/禁用
        function togglePctInput(checkbox) {
            var baseNoteId = checkbox.value;
            var pctInput = document.querySelector('input[name="baseNotePct_' + baseNoteId + '"]');
            if (pctInput) {
                pctInput.disabled = !checkbox.checked;
                if (!checkbox.checked) pctInput.value = '0';
            }
            updatePctSummary();
        }
        
        // 实时更新配比总和
        function updatePctSummary() {
            var total = 0;
            var checkboxes = document.querySelectorAll('input[name="baseNoteSelect"]:checked');
            checkboxes.forEach(function(cb) {
                var pctInput = document.querySelector('input[name="baseNotePct_' + cb.value + '"]');
                if (pctInput) total += parseFloat(pctInput.value) || 0;
            });
            var summary = document.getElementById('baseNotePctSummary');
            if (summary) {
                summary.style.display = 'block';
                var checkedCount = checkboxes.length;
                if (checkedCount === 0) {
                    summary.innerHTML = '<span style="color:#f44336;">⚠ 必须关联至少一个基香</span>';
                    summary.style.background = 'rgba(244,67,54,0.1)';
                    summary.style.border = '1px solid rgba(244,67,54,0.3)';
                } else if (Math.abs(total - 100) > 0.01) {
                    summary.innerHTML = '基香配比总和: <strong style="color:#f44336;">' + total.toFixed(2) + '%</strong> (需等于100%)';
                    summary.style.background = 'rgba(244,67,54,0.1)';
                    summary.style.border = '1px solid rgba(244,67,54,0.3)';
                } else {
                    summary.innerHTML = '基香配比总和: <strong style="color:#4caf50;">100%</strong> ✓';
                    summary.style.background = 'rgba(76,175,80,0.1)';
                    summary.style.border = '1px solid rgba(76,175,80,0.3)';
                }
            }
        }
        
        // 表单提交验证
        function validateBaseNotes() {
            var checkboxes = document.querySelectorAll('input[name="baseNoteSelect"]:checked');
            if (checkboxes.length === 0) {
                alert('必须关联至少一个基香，基香配比总和必须为100%');
                return false;
            }
            var total = 0;
            checkboxes.forEach(function(cb) {
                var pctInput = document.querySelector('input[name="baseNotePct_' + cb.value + '"]');
                if (pctInput) total += parseFloat(pctInput.value) || 0;
            });
            if (Math.abs(total - 100) > 0.01) {
                alert('基香配比总和必须为100%，当前为' + total.toFixed(2) + '%');
                return false;
            }
            return true;
        }
        
        window.onclick = function(event) {
            var modal = document.getElementById('noteModal');
            if (event.target == modal) {
                modal.style.display = 'none';
            }
        }

        // 图片压缩函数 - note
        function compressImage_note(file, maxSizeKB, callback) {
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

        // 图片上传 - note
        document.getElementById('fileInput_note').addEventListener('change', function(e) {
            var file = e.target.files[0];
            if (!file) return;
            var maxSize = 5 * 1024 * 1024;
            if (file.size > maxSize) { alert('文件大小不能超过5MB'); return; }
            var allowedTypes = ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'image/svg+xml'];
            if (allowedTypes.indexOf(file.type) === -1) { alert('仅支持 JPG/PNG/GIF/WebP/SVG 格式'); return; }
            var reader = new FileReader();
            reader.onload = function(ev) {
                document.getElementById('previewImg_note').src = ev.target.result;
                document.getElementById('previewImg_note').style.display = 'block';
                document.getElementById('placeholder_note').style.display = 'none';
            };
            reader.readAsDataURL(file);
            compressImage_note(file, 180, function(fileToUpload, wasCompressed) {
                uploadImage_note(fileToUpload, 'note', wasCompressed);
            });
        });

        function uploadImage_note(file, uploadType, wasCompressed) {
            var formData = new FormData();
            formData.append('file', file);
            formData.append('type', uploadType);
            var csrfInput = document.querySelector('input[name="csrf_token"]');
            if (csrfInput) formData.append('csrf_token', csrfInput.value);
            var progressDiv = document.getElementById('uploadProgress_note');
            var progressBar = document.getElementById('progressBar_note');
            var progressText = document.getElementById('progressText_note');
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
                        document.getElementById('imageURL_note').value = resp.url;
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
            var wrapper = document.getElementById('imagePreview_note').parentElement;
            wrapper.addEventListener('dragover', function(e) { e.preventDefault(); wrapper.classList.add('dragover'); });
            wrapper.addEventListener('dragleave', function() { wrapper.classList.remove('dragover'); });
            wrapper.addEventListener('drop', function(e) {
                e.preventDefault();
                wrapper.classList.remove('dragover');
                var file = e.dataTransfer.files[0];
                if (file) {
                    document.getElementById('fileInput_note').files = e.dataTransfer.files;
                    document.getElementById('fileInput_note').dispatchEvent(new Event('change'));
                }
            });
            document.getElementById('imagePreview_note').addEventListener('click', function() {
                document.getElementById('fileInput_note').click();
            });
        })();

        function toggleUrlInput_note() {
            var el = document.getElementById('urlInputWrapper_note');
            el.style.display = el.style.display === 'none' ? 'block' : 'none';
        }

        function applyManualUrl_note() {
            var url = document.getElementById('manualUrl_note').value.trim();
            if (url) {
                document.getElementById('imageURL_note').value = url;
                document.getElementById('previewImg_note').src = url;
                document.getElementById('previewImg_note').style.display = 'block';
                document.getElementById('placeholder_note').style.display = 'none';
                document.getElementById('urlInputWrapper_note').style.display = 'none';
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
Call CloseConnection()
%>
