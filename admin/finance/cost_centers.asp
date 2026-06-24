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

Function GetScalar(sql)
    Dim rs, val : val = 0
    On Error Resume Next
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        If Not rs Is Nothing Then
            If Not rs.EOF Then
                val = rs(0)
                rs.Close
            End If
        End If
    Else
        Err.Clear
    End If
    Set rs = Nothing
    GetScalar = val
End Function

' V10.1: 移除运行时DDL - CostCenters表应由 deploy.asp 预先创建
If Not TableExists_CC("CostCenters") Then
    Response.Write "<div style='padding:40px;color:#f44336;background:#1a1a2e;font-family:sans-serif;'><h2>表缺失</h2><p>CostCenters 表不存在，请先运行 <a href='/setup/deploy.asp' style='color:#00bcd4;'>系统部署工具</a> 创建数据库表并初始化种子数据。</p></div>"
    Response.End
End If

' POST 处理
Dim action : action = Request.Form("action")
If action = "save" Then
    Dim ccID : ccID = SafeNum(Request.Form("centerID"))
    Dim ccCode : ccCode = Replace(Request.Form("centerCode"),"'","''")
    Dim ccName : ccName = Replace(Request.Form("centerName"),"'","''")
    Dim ccType : ccType = Replace(Request.Form("centerType"),"'","''")
    Dim ccBudget : ccBudget = SafeNum(Request.Form("budgetAmount"))
    Dim ccNotes : ccNotes = Replace(Request.Form("notes"),"'","''")
    
    If ccID > 0 Then
        conn.Execute "UPDATE CostCenters SET CenterCode='" & ccCode & "', CenterName='" & ccName & "', CenterType='" & ccType & "', BudgetAmount=" & ccBudget & ", Notes='" & ccNotes & "', UpdatedAt=GETDATE() WHERE CenterID=" & ccID
    Else
        conn.Execute "INSERT INTO CostCenters (CenterCode, CenterName, CenterType, BudgetAmount, Notes) VALUES ('" & ccCode & "','" & ccName & "','" & ccType & "'," & ccBudget & ",'" & ccNotes & "')"
    End If
    Response.Redirect "cost_centers.asp?msg=" & Server.URLEncode("保存成功")
    Response.End
ElseIf action = "toggle" Then
    Dim tID : tID = SafeNum(Request.Form("centerID"))
    If tID > 0 Then conn.Execute "UPDATE CostCenters SET IsActive = CASE WHEN IsActive=1 THEN 0 ELSE 1 END WHERE CenterID=" & tID
    Response.Redirect "cost_centers.asp?msg=" & Server.URLEncode("状态切换成功")
    Response.End
End If

' 加载成本中心列表
Dim rsCC, ccList(), ccCount : ccCount = 0
Set rsCC = Server.CreateObject("ADODB.Recordset")
rsCC.CursorLocation = 3 ' adUseClient - 支持 MoveLast/RecordCount
rsCC.Open "SELECT * FROM CostCenters ORDER BY CenterType, CenterID", conn, 1, 1
If Not rsCC Is Nothing Then
    If Not rsCC.EOF Then
        rsCC.MoveLast : ccCount = rsCC.RecordCount : rsCC.MoveFirst
        ReDim ccList(ccCount - 1, 7)
        Dim cci : cci = 0
        Do While Not rsCC.EOF
            ccList(cci, 0) = rsCC("CenterID")
            ccList(cci, 1) = rsCC("CenterCode")
            ccList(cci, 2) = rsCC("CenterName")
            ccList(cci, 3) = rsCC("CenterType")
            ccList(cci, 4) = rsCC("BudgetAmount")
            ccList(cci, 5) = rsCC("IsActive")
            ccList(cci, 6) = rsCC("Notes") & ""
            ccList(cci, 7) = GetScalar("SELECT ISNULL(SUM(DebitAmount-CreditAmount),0) FROM GLTransactions WHERE CenterID=" & rsCC("CenterID"))
            cci = cci + 1
            rsCC.MoveNext
        Loop
    End If
    rsCC.Close : Set rsCC = Nothing
End If

' 统计
Dim ccTypeCount : ccTypeCount = GetScalar("SELECT COUNT(DISTINCT CenterType) FROM CostCenters")
Dim ccTotalBudget : ccTotalBudget = GetScalar("SELECT ISNULL(SUM(BudgetAmount),0) FROM CostCenters WHERE IsActive=1")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>成本中心设置 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; font-family: 'Segoe UI',Arial,sans-serif; }
        .main-content { margin-left: 250px; padding: 30px; min-height: 100vh; }
        .page-header { margin-bottom: 25px; }
        .page-title { color: #fff; font-size: 24px; margin: 0 0 8px; }
        .breadcrumb { color: #888; font-size: 13px; }
        .breadcrumb a { color: #00bcd4; text-decoration: none; }
        .stats-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); padding: 20px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); text-align: center; }
        .stat-card .num { font-size: 28px; font-weight: 700; color: #00bcd4; }
        .stat-card .label { font-size: 13px; color: #888; margin-top: 5px; }
        .card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); margin-bottom: 20px; overflow: hidden; }
        .card-header { padding: 15px 20px; border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 16px; display: flex; justify-content: space-between; align-items: center; }
        .card-body { padding: 20px; }
        table { width: 100%; border-collapse: collapse; }
        th { text-align: left; padding: 12px; background: rgba(0,188,212,0.06); border-bottom: 1px solid rgba(255,255,255,0.06); font-size: 13px; color: #999; }
        td { padding: 12px; border-bottom: 1px solid rgba(255,255,255,0.04); font-size: 14px; }
        tr:hover { background: rgba(255,255,255,0.02); }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; }
        .badge-active { background: rgba(76,175,80,0.2); color: #4CAF50; }
        .badge-inactive { background: rgba(158,158,158,0.2); color: #9E9E9E; }
        .badge-procurement { background: rgba(255,152,0,0.2); color: #FFB74D; }
        .badge-production { background: rgba(33,150,243,0.2); color: #64B5F6; }
        .badge-logistics { background: rgba(0,188,212,0.2); color: #80DEEA; }
        .badge-marketing { background: rgba(156,39,176,0.2); color: #CE93D8; }
        .badge-admin { background: rgba(158,158,158,0.2); color: #BDBDBD; }
        .badge-rnd { background: rgba(76,175,80,0.2); color: #81C784; }
        .modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.7); z-index: 1000; }
        .modal-content { background: linear-gradient(135deg, #2d2d44, #1e1e32); width: 90%; max-width: 500px; margin: 60px auto; padding: 30px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.1); }
        .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }
        .modal-header h3 { color: #fff; margin: 0; }
        .close-btn { background: none; border: none; color: #888; font-size: 24px; cursor: pointer; }
        .form-group { margin-bottom: 15px; }
        .form-group label { display: block; color: #b0b0b0; margin-bottom: 6px; font-size: 13px; }
        .form-group input, .form-group select { width: 100%; padding: 10px; background: #1e1e32; border: 1px solid rgba(255,255,255,0.1); border-radius: 6px; color: #e0e0e0; font-size: 14px; box-sizing: border-box; }
        .form-actions { text-align: right; margin-top: 20px; padding-top: 15px; border-top: 1px solid rgba(255,255,255,0.06); }
        .amount { font-weight: 600; color: #00bcd4; }
        .amount.negative { color: #f44336; }
        .msg { padding: 12px 20px; border-radius: 6px; margin-bottom: 16px; }
        .msg-success { background: rgba(76,175,80,0.15); color: #4CAF50; border: 1px solid rgba(76,175,80,0.3); }
        @media (max-width: 768px) { .main-content { margin-left: 0; } .stats-grid { grid-template-columns: 1fr; } }
    </style>
</head>
<body data-theme="operation-dark">
<!--#include file="includes/nav.asp"-->
<div class="main-content">
    <div class="page-header">
        <h2 class="page-title"><i class="fas fa-building"></i> 成本中心设置</h2>
        <div class="breadcrumb"><a href="index.asp">财务中心</a> / 成本中心设置</div>
    </div>
    <% If Request.QueryString("msg") <> "" Then %>
    <div class="msg msg-success"><i class="fas fa-check-circle"></i> <%= Server.HTMLEncode(Request.QueryString("msg")) %></div>
    <% End If %>
    
    <div class="stats-grid">
        <div class="stat-card"><div class="num"><%= ccCount %></div><div class="label">成本中心总数</div></div>
        <div class="stat-card"><div class="num"><%= ccTypeCount %></div><div class="label">类型数量</div></div>
        <div class="stat-card"><div class="num">¥<%= FormatNumber(ccTotalBudget, 0) %></div><div class="label">总预算金额</div></div>
    </div>
    
    <div class="card">
        <div class="card-header">
            <span><i class="fas fa-list"></i> 成本中心列表</span>
            <button class="btn btn-primary" onclick="openModal()"><i class="fas fa-plus"></i> 新增成本中心</button>
        </div>
        <div class="card-body">
            <table>
                <thead>
                    <tr><th>编码</th><th>名称</th><th>类型</th><th>预算金额</th><th>实际发生</th><th>状态</th><th>操作</th></tr>
                </thead>
                <tbody>
                    <% If ccCount > 0 Then
                        For cci = 0 To ccCount - 1
                            Dim ctBadge : ctBadge = "badge-admin"
                            Select Case ccList(cci, 3)
                                Case "Procurement" : ctBadge = "badge-procurement"
                                Case "Production" : ctBadge = "badge-production"
                                Case "Logistics" : ctBadge = "badge-logistics"
                                Case "Marketing" : ctBadge = "badge-marketing"
                                Case "R&D" : ctBadge = "badge-rnd"
                            End Select
                            Dim cb : cb = SafeNum(ccList(cci, 7))
                    %>
                    <tr>
                        <td><strong><%= ccList(cci, 1) %></strong></td>
                        <td><%= ccList(cci, 2) %></td>
                        <td><span class="badge <%= ctBadge %>"><%= ccList(cci, 3) %></span></td>
                        <td class="amount">¥<%= FormatNumber(SafeNum(ccList(cci, 4)), 0) %></td>
                        <td class="amount <%= IIf(cb < 0, "negative", "") %>">¥<%= FormatNumber(cb, 0) %></td>
                        <td><span class="badge <%= IIf(ccList(cci, 5) = 1, "badge-active", "badge-inactive") %>"><%= IIf(ccList(cci, 5) = 1, "启用", "禁用") %></span></td>
                        <td>
                            <button class="btn btn-edit btn-sm" onclick="editCenter('<%= ccList(cci, 0) %>','<%= ccList(cci, 1) %>','<%= Server.HTMLEncode(ccList(cci, 2)) %>','<%= ccList(cci, 3) %>','<%= ccList(cci, 4) %>','<%= Server.HTMLEncode(ccList(cci, 6)) %>')"><i class="fas fa-edit"></i></button>
                            <form method="post" style="display:inline"><input type="hidden" name="action" value="toggle"><input type="hidden" name="centerID" value="<%= ccList(cci, 0) %>">
                            <button type="submit" class="btn btn--warning btn--sm"><i class="fas fa-exchange-alt"></i></button></form>
                        </td>
                    </tr>
                    <% Next
                    Else %>
                    <tr><td colspan="7" style="text-align:center;padding:40px;color:#666;">暂无成本中心数据</td></tr>
                    <% End If %>
                </tbody>
            </table>
        </div>
    </div>
</div>

<!-- 编辑弹窗 -->
<div id="centerModal" class="modal">
    <div class="modal-content">
        <div class="modal-header">
            <h3 id="modalTitle"><i class="fas fa-plus-circle"></i> 新增成本中心</h3>
            <button class="close-btn" onclick="closeModal()">&times;</button>
        </div>
        <form method="post" action="cost_centers.asp">
            <input type="hidden" name="action" value="save">
            <input type="hidden" name="centerID" id="editCenterID">
            <div class="form-group"><label>中心编码 *</label><input type="text" name="centerCode" id="editCode" required></div>
            <div class="form-group"><label>中心名称 *</label><input type="text" name="centerName" id="editName" required></div>
            <div class="form-group"><label>类型 *</label>
                <select name="centerType" id="editType">
                    <option value="Procurement">采购</option><option value="Production">生产</option>
                    <option value="Logistics">物流</option><option value="Marketing">市场</option>
                    <option value="Admin">行政</option><option value="R&D">研发</option><option value="Department">部门</option>
                </select>
            </div>
            <div class="form-group"><label>预算金额</label><input type="number" name="budgetAmount" id="editBudget" step="0.01" value="0"></div>
            <div class="form-group"><label>备注</label><input type="text" name="notes" id="editNotes"></div>
            <div class="form-actions">
                <button type="button" class="btn btn-outline" onclick="closeModal()">取消</button>
                <button type="submit" class="btn btn-primary"><i class="fas fa-save"></i> 保存</button>
            </div>
        </form>
    </div>
</div>
<script>
function openModal(){ document.getElementById('modalTitle').innerHTML='<i class="fas fa-plus-circle"></i> 新增成本中心'; document.getElementById('editCenterID').value=''; document.getElementById('editCode').value=''; document.getElementById('editName').value=''; document.getElementById('editType').value='Procurement'; document.getElementById('editBudget').value='0'; document.getElementById('editNotes').value=''; document.getElementById('centerModal').style.display='block'; }
function editCenter(id,code,name,type,budget,notes){ document.getElementById('modalTitle').innerHTML='<i class="fas fa-edit"></i> 编辑成本中心'; document.getElementById('editCenterID').value=id; document.getElementById('editCode').value=code; document.getElementById('editName').value=name; document.getElementById('editType').value=type; document.getElementById('editBudget').value=budget; document.getElementById('editNotes').value=notes; document.getElementById('centerModal').style.display='block'; }
function closeModal(){ document.getElementById('centerModal').style.display='none'; }
window.onclick=function(e){ if(e.target.classList.contains('modal')) e.target.style.display='none'; }
</script>
</body>
</html>
<%
' 健壮的表存在检查（使用 sys.tables 而非 On Error Resume Next）
Function TableExists_CC(tblName)
    Dim rs, exists : exists = False
    On Error Resume Next
    Set rs = conn.Execute("SELECT 1 FROM sys.tables WHERE name='" & Replace(tblName,"'","''") & "'")
    If Err.Number = 0 Then
        If Not rs Is Nothing Then
            If Not rs.EOF Then exists = True
            rs.Close
        End If
    End If
    Err.Clear : Set rs = Nothing
    On Error GoTo 0
    TableExists_CC = exists
End Function

Call CloseConnection()
%>
