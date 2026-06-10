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

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Function SafeSQL(str)
    If IsNull(str) Or str = "" Then SafeSQL = "" Else SafeSQL = Replace(str, "'", "''")
End Function

Function GetScalar(sql)
    Dim rs, val : val = 0
    On Error Resume Next
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        If Not rs Is Nothing Then
            If Not rs.EOF Then val = rs(0)
            If IsNull(val) Then val = 0
            rs.Close
        End If
    Else : Err.Clear
    End If
    Set rs = Nothing
    GetScalar = val
End Function

Dim action, msg, msgType
action = Trim(Request.Form("action"))
msg = Trim(Request.QueryString("msg"))
msgType = "success"
If InStr(msg, "失败") > 0 Or InStr(msg, "错误") > 0 Then msgType = "error"

' ========== POST 处理 ==========
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If action = "toggle_active" Then
        Dim bnID, bnActive
        bnID = SafeNum(Request.Form("base_note_id"))
        bnActive = SafeNum(Request.Form("is_active"))
        If bnID > 0 Then
            conn.Execute "UPDATE BaseNotes SET IsActive=" & bnActive & " WHERE BaseNoteID=" & bnID
            If bnActive = 1 Then
                Response.Redirect "base_note_production.asp?msg=基香已激活"
            Else
                Response.Redirect "base_note_production.asp?msg=基香已停用"
            End If
            Response.End
        End If
    ElseIf action = "update_ingredients" Then
        Dim uiID, uiIngredients
        uiID = SafeNum(Request.Form("base_note_id"))
        uiIngredients = Trim(Request.Form("ingredients"))
        If uiID > 0 And uiIngredients <> "" Then
            conn.Execute "UPDATE BaseNotes SET Ingredients='" & SafeSQL(uiIngredients) & "' WHERE BaseNoteID=" & uiID
            Response.Redirect "base_note_production.asp?msg=成分已更新"
            Response.End
        End If
    End If
End If

' ========== 统计 ==========
Dim bnTotal, bnInactive, bnNoteCount
bnTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM BaseNotes"))
bnActive = SafeNum(GetScalar("SELECT COUNT(*) FROM BaseNotes WHERE IsActive=1"))
bnInactive = bnTotal - bnActive
bnNoteCount = SafeNum(GetScalar("SELECT COUNT(DISTINCT NoteID) FROM FragranceNotes WHERE BaseNoteID IS NOT NULL"))

' ========== 基香列表 ==========
Dim rsBaseNotes
Set rsBaseNotes = conn.Execute("SELECT bn.*, (SELECT COUNT(*) FROM FragranceNotes WHERE BaseNoteID=bn.BaseNoteID) AS NoteUsageCount FROM BaseNotes bn ORDER BY bn.IsActive DESC, bn.BaseNoteName ASC")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>基香生产 - 半成品生产中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --accent: #2196F3; --input-bg: #2d2d44; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #1a1a2e; color: var(--text); min-height: 100vh; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-header { margin-bottom: 25px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #2196F3; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 20px; border-radius: 12px; text-align: center; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .num { font-size: 28px; font-weight: bold; display: block; }
        .stat-card .label { font-size: 12px; color: #888; display: block; margin-top: 5px; }
        .stat-card .num.c-blue { color: #2196F3; }
        .stat-card .num.c-green { color: #4CAF50; }
        .stat-card .num.c-orange { color: #FF9800; }
        .stat-card .num.c-gray { color: #888; }
        
        .card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; margin-bottom: 25px; border: 1px solid rgba(255,255,255,0.06); }
        .card-header { padding: 14px 20px; background: rgba(33,150,243,0.08); border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 15px; color: #e0e0e0; display: flex; justify-content: space-between; align-items: center; }
        .card-body { padding: 16px 20px; overflow-x: auto; }
        
        table { width: 100%; border-collapse: collapse; }
        th { background: rgba(33,150,243,0.15); color: #64b5f6; padding: 12px 15px; text-align: left; font-weight: 600; font-size: 13px; }
        td { padding: 12px 15px; border-bottom: 1px solid rgba(255,255,255,0.04); color: #e0e0e0; font-size: 14px; }
        tr:hover td { background: rgba(255,255,255,0.03); }
        
        .status-badge { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 500; }
        .status-active { background: rgba(76,175,80,0.15); color: #81c784; }
        .status-inactive { background: rgba(244,67,54,0.15); color: #e57373; }
        
        .alert { padding: 15px 20px; border-radius: 8px; margin-bottom: 20px; }
        .alert-success { background: rgba(76,175,80,0.15); color: #81c784; border: 1px solid rgba(76,175,80,0.3); }
        .alert-error { background: rgba(244,67,54,0.15); color: #e57373; border: 1px solid rgba(244,67,54,0.3); }
        
        .ingredient-tags { display: flex; flex-wrap: wrap; gap: 4px; }
        .ingredient-tag { background: rgba(33,150,243,0.12); color: #64b5f6; padding: 2px 10px; border-radius: 10px; font-size: 12px; }
        
        .expandable-row { cursor: pointer; }
        .detail-row { display: none; }
        .detail-row.show { display: table-row; }
        .detail-row td { background: rgba(0,0,0,0.15); padding: 20px; }
        
        .text-center { text-align: center; }
        .text-muted { color: #888; font-size: 13px; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-vial"></i> 基香生产管理</h2>
        </div>
        
        <% If msg <> "" Then %>
        <div class="alert alert-<%=msgType%>"><%=Server.HTMLEncode(msg)%></div>
        <% End If %>
        
        <!-- 统计卡片 -->
        <div class="stats-grid">
            <div class="stat-card"><span class="num c-blue"><%=bnTotal%></span><span class="label">基香总数</span></div>
            <div class="stat-card"><span class="num c-green"><%=bnActive%></span><span class="label">已激活</span></div>
            <div class="stat-card"><span class="num c-gray"><%=bnInactive%></span><span class="label">已停用</span></div>
            <div class="stat-card"><span class="num c-orange"><%=bnNoteCount%></span><span class="label">关联香调</span></div>
        </div>
        
        <!-- 基香列表 -->
        <div class="card">
            <div class="card-header">基香列表</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>基香名称</th><th>成分</th><th>关联香调数</th><th>状态</th><th>操作</th></tr></thead>
                    <tbody>
                    <%
                    If Not rsBaseNotes Is Nothing Then
                        Do While Not rsBaseNotes.EOF
                            Dim bID, bName, bIngredients, bActive, bNoteUsage
                            bID = rsBaseNotes("BaseNoteID")
                            bName = CStr(rsBaseNotes("BaseNoteName") & "")
                            bIngredients = CStr(rsBaseNotes("Ingredients") & "")
                            bActive = SafeNum(rsBaseNotes("IsActive"))
                            bNoteUsage = SafeNum(rsBaseNotes("NoteUsageCount"))
                            
                            Dim ingArr, ingItem
                            ingArr = Split(bIngredients, ",")
                    %>
                        <tr>
                            <td><strong><%=Server.HTMLEncode(bName)%></strong></td>
                            <td>
                                <div class="ingredient-tags">
                                <% For Each ingItem In ingArr
                                    ingItem = Trim(ingItem)
                                    If ingItem <> "" Then %>
                                    <span class="ingredient-tag"><%=Server.HTMLEncode(ingItem)%></span>
                                <% End If
                                Next %>
                                </div>
                            </td>
                            <td><%=bNoteUsage%></td>
                            <td>
                                <span class="status-badge <%=IIF(bActive=1,"status-active","status-inactive")%>">
                                    <%=IIF(bActive=1,"激活","停用")%>
                                </span>
                            </td>
                            <td>
                                <form method="post" style="display:inline;">
                                    <input type="hidden" name="action" value="toggle_active">
                                    <input type="hidden" name="base_note_id" value="<%=bID%>">
                                    <input type="hidden" name="is_active" value="<%=IIF(bActive=1,0,1)%>">
                                    <% If bActive = 1 Then %>
                                    <button type="submit" class="btn btn-warning btn-sm">停用</button>
                                    <% Else %>
                                    <button type="submit" class="btn btn-success btn-sm">激活</button>
                                    <% End If %>
                                </form>
                                <button class="btn btn-primary btn-sm" onclick="openEditModal(<%=bID%>,'<%=Server.HTMLEncode(bName)%>','<%=Server.HTMLEncode(Replace(bIngredients,"'","\'"))%>')">编辑成分</button>
                            </td>
                        </tr>
                    <%
                            rsBaseNotes.MoveNext
                        Loop
                        rsBaseNotes.Close
                    End If
                    Set rsBaseNotes = Nothing
                    %>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    
    <!-- 编辑成分弹窗 -->
    <div id="editModal" style="display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.7);z-index:1000;">
        <div style="background:linear-gradient(135deg,#2d2d44,#1e1e32);width:90%;max-width:550px;margin:80px auto;padding:30px;border-radius:15px;border:1px solid rgba(255,255,255,0.06);">
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;padding-bottom:15px;border-bottom:1px solid rgba(255,255,255,0.06);">
                <h3 style="margin:0;font-size:18px;">编辑基香成分 - <span id="editName"></span></h3>
                <button onclick="closeEditModal()" style="background:none;border:none;color:#888;font-size:24px;cursor:pointer;">&times;</button>
            </div>
            <form method="post">
                <input type="hidden" name="action" value="update_ingredients">
                <input type="hidden" name="base_note_id" id="editID">
                <div style="margin-bottom:18px;">
                    <label style="display:block;margin-bottom:6px;font-weight:600;color:#e0e0e0;font-size:13px;">成分列表（逗号分隔）</label>
                    <textarea id="editIngredients" name="ingredients" rows="5" style="width:100%;padding:10px 12px;background:#2d2d44;border:1px solid rgba(255,255,255,0.12);border-radius:6px;color:#e0e0e0;font-size:14px;"></textarea>
                </div>
                <div style="display:flex;justify-content:flex-end;gap:10px;">
                    <button type="button" class="btn btn--neutral" onclick="closeEditModal()">取消</button>
                    <button type="submit" class="btn btn-primary">保存</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
    function openEditModal(id, name, ingredients) {
        document.getElementById('editID').value = id;
        document.getElementById('editName').innerText = name;
        document.getElementById('editIngredients').value = ingredients;
        document.getElementById('editModal').style.display = 'block';
    }
    function closeEditModal() {
        document.getElementById('editModal').style.display = 'none';
    }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
