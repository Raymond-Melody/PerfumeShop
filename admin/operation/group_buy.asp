<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<!--#include file="../../includes/dal.asp"-->
<%
Call OpenConnection()

Dim gbAction, gbActionMsg, gbActionResult
gbAction = Request.QueryString("action")
If gbAction = "" Then gbAction = Request.Form("action")
gbActionMsg = ""
gbActionResult = True

' 新增/编辑拼团计划
If gbAction = "save" And Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim gbEditId, gbProductID, gbTeamSize, gbGroupPrice, gbMinUnit, gbMaxUnit, gbStart, gbEnd, gbDuration, gbSort
    gbEditId = Request.Form("edit_id")
    gbProductID = CLng(Request.Form("product_id"))
    gbTeamSize = CLng(Request.Form("team_size"))
    gbGroupPrice = CDbl(Request.Form("group_price"))
    gbMinUnit = Request.Form("min_unit")
    gbMaxUnit = Request.Form("max_unit")
    If gbMinUnit = "" Then gbMinUnit = "0"
    If gbMaxUnit = "" Then gbMaxUnit = "0"
    gbStart = Request.Form("start_time")
    gbEnd = Request.Form("end_time")
    gbDuration = CLng(Request.Form("duration_hours"))
    gbSort = Request.Form("sort_order")
    If gbSort = "" Or Not IsNumeric(gbSort) Then gbSort = 0

    If gbProductID <= 0 Or gbGroupPrice <= 0 Or gbTeamSize <= 1 Then
        gbActionMsg = "请完整填写必填字段"
        gbActionResult = False
    ElseIf CDate(gbEnd) <= CDate(gbStart) Then
        gbActionMsg = "结束时间必须晚于开始时间"
        gbActionResult = False
    Else
        On Error Resume Next
        If gbEditId <> "" And IsNumeric(gbEditId) Then
            conn.Execute "UPDATE GroupBuyPlans SET ProductID=" & gbProductID & ", TeamSize=" & gbTeamSize & _
                       ", GroupPrice=" & gbGroupPrice & ", MinUnit=" & gbMinUnit & ", MaxUnit=" & gbMaxUnit & _
                       ", StartTime='" & SafeSQL(gbStart) & "', EndTime='" & SafeSQL(gbEnd) & _
                       "', DurationHours=" & gbDuration & ", SortOrder=" & gbSort & " WHERE PlanID=" & gbEditId
            If Err.Number = 0 Then gbActionMsg = "更新成功" Else gbActionMsg = "更新失败: " & Err.Description : gbActionResult = False
        Else
            conn.Execute "INSERT INTO GroupBuyPlans (ProductID, TeamSize, GroupPrice, MinUnit, MaxUnit, StartTime, EndTime, DurationHours, SortOrder) VALUES (" & _
                       gbProductID & "," & gbTeamSize & "," & gbGroupPrice & "," & gbMinUnit & "," & gbMaxUnit & ",'" & _
                       SafeSQL(gbStart) & "','" & SafeSQL(gbEnd) & "'," & gbDuration & "," & gbSort & ")"
            If Err.Number = 0 Then gbActionMsg = "创建成功" Else gbActionMsg = "创建失败: " & Err.Description : gbActionResult = False
        End If
        On Error GoTo 0
    End If
End If

' 删除
If gbAction = "delete" Then
    Dim gbDelId : gbDelId = Request.QueryString("id")
    If IsNumeric(gbDelId) Then
        conn.Execute "DELETE FROM GroupBuyPlans WHERE PlanID = " & gbDelId
        gbActionMsg = "已删除"
    End If
End If

' 切换状态
If gbAction = "toggle" Then
    Dim gbTogId : gbTogId = Request.QueryString("id")
    If IsNumeric(gbTogId) Then
        conn.Execute "UPDATE GroupBuyPlans SET IsActive = CASE WHEN IsActive = 1 THEN 0 ELSE 1 END WHERE PlanID = " & gbTogId
        gbActionMsg = "状态已切换"
    End If
End If

' 统计数据
Dim gbTotalPlans, gbActivePlans, gbTotalGroups, gbSuccessGroups, gbTotalParticipants
gbTotalPlans = GetScalar("SELECT COUNT(*) FROM GroupBuyPlans")
gbActivePlans = GetScalar("SELECT COUNT(*) FROM GroupBuyPlans WHERE IsActive = 1 AND GETDATE() >= StartTime AND GETDATE() <= EndTime")
gbTotalGroups = GetScalar("SELECT COUNT(*) FROM GroupBuyOrders")
gbSuccessGroups = GetScalar("SELECT COUNT(*) FROM GroupBuyOrders WHERE Status = 1")
gbTotalParticipants = GetScalar("SELECT COUNT(*) FROM GroupBuyParticipants")

' 获取编辑中的计划
Dim gbEditData, gbIsEditing
gbEditId = Request.QueryString("edit_id")
Set gbEditData = Nothing
gbIsEditing = False
If gbEditId <> "" And IsNumeric(gbEditId) Then
    Set gbEditData = conn.Execute("SELECT * FROM GroupBuyPlans WHERE PlanID = " & gbEditId)
    If Not gbEditData Is Nothing Then
        If Not gbEditData.EOF Then
            gbIsEditing = True
        End If
    End If
End If

' 预提取编辑数据到变量（VBScript IIf/And不短路，需提前取值）
Dim eGBGroupPrice, eGBDuration, eGBStartTime, eGBEndTime, eGBMinUnit, eGBMaxUnit, eGBSort
If gbIsEditing Then
    eGBGroupPrice = gbEditData("GroupPrice")
    eGBDuration = gbEditData("DurationHours")
    eGBStartTime = FormatDateTime(gbEditData("StartTime"), 0)
    eGBEndTime = FormatDateTime(gbEditData("EndTime"), 0)
    eGBMinUnit = gbEditData("MinUnit")
    eGBMaxUnit = gbEditData("MaxUnit")
    eGBSort = gbEditData("SortOrder")
Else
    eGBGroupPrice = ""
    eGBDuration = "24"
    eGBStartTime = FormatDateTime(Now(), 0)
    eGBEndTime = FormatDateTime(DateAdd("d", 7, Now()), 0)
    eGBMinUnit = "1"
    eGBMaxUnit = "0"
    eGBSort = "0"
End If

' 所有计划列表
Dim rsPlans : Set rsPlans = DAL_GetList("SELECT gp.*, p.ProductName, p.BasePrice FROM GroupBuyPlans gp INNER JOIN Products p ON gp.ProductID = p.ProductID ORDER BY gp.StartTime DESC", Null)

' 产品列表
Dim rsGBProducts : Set rsGBProducts = DAL_GetList("SELECT ProductID, ProductName, BasePrice FROM Products WHERE IsActive <> 0 ORDER BY ProductName", Null)

' 团列表（含参团情况）
Dim rsGroups : Set rsGroups = DAL_GetList("SELECT g.GroupID, g.PlanID, g.GroupSN, g.InitiatorID, g.Status, g.CurrentSize, g.TargetSize, g.CreatedAt, g.CompletedAt, u.Username AS InitiatorName, p.ProductName FROM GroupBuyOrders g LEFT JOIN Users u ON g.InitiatorID = u.UserID INNER JOIN GroupBuyPlans gp ON g.PlanID = gp.PlanID INNER JOIN Products p ON gp.ProductID = p.ProductID ORDER BY g.CreatedAt DESC", Null)
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>拼团管理 - 运营管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; }
        .main-content { padding: 24px; }
        .page-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }
        .page-title { font-size: 22px; color: #fff; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #7c4dff; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(5, 1fr); gap: 16px; margin-bottom: 24px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); padding: 20px; border-radius: 10px; border: 1px solid rgba(255,255,255,0.05); text-align: center; }
        .stat-value { font-size: 28px; font-weight: bold; color: #7c4dff; }
        .stat-label { font-size: 13px; color: #888; margin-top: 4px; }
        
        .panel { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; padding: 24px; border: 1px solid rgba(255,255,255,0.05); margin-bottom: 24px; }
        .panel h3 { color: #fff; margin: 0 0 16px; font-size: 18px; display: flex; align-items: center; gap: 8px; }
        .panel h3 i { color: #7c4dff; }
        
        .form-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 12px; margin-bottom: 12px; }
        .form-group { display: flex; flex-direction: column; gap: 4px; }
        .form-group label { font-size: 12px; color: #888; font-weight: 500; }
        .form-group input, .form-group select, .form-group textarea {
            padding: 8px 12px; border: 1px solid #3a3a4a; border-radius: 6px; background: #1a1a2e; color: #e0e0e0; font-size: 13px;
        }
        .form-group input:focus, .form-group select:focus { border-color: #7c4dff; outline: none; }
        
        .btn { padding: 8px 18px; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; font-weight: 500; transition: all 0.2s; text-decoration: none; display: inline-block; }
        .btn-primary { background: linear-gradient(135deg, #7c4dff, #651fff); color: #fff; }
        .btn-primary:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(124,77,255,0.3); }
        .btn-danger { background: #c62828; color: #fff; }
        .btn-sm { padding: 4px 12px; font-size: 11px; }
        .btn-outline { background: transparent; border: 1px solid #555; color: #ccc; }
        
        .gb-table { width: 100%; border-collapse: collapse; }
        .gb-table th { text-align: left; padding: 10px 12px; background: rgba(0,0,0,0.2); color: #888; font-size: 11px; text-transform: uppercase; }
        .gb-table td { padding: 10px 12px; border-bottom: 1px solid rgba(255,255,255,0.05); font-size: 13px; }
        .gb-table tr:hover td { background: rgba(255,255,255,0.02); }
        
        .badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; }
        .badge-active { background: rgba(76,175,80,0.2); color: #4CAF50; }
        .badge-inactive { background: rgba(158,158,158,0.2); color: #9e9e9e; }
        .badge-pending { background: rgba(255,193,7,0.2); color: #FFC107; }
        .badge-success { background: rgba(76,175,80,0.2); color: #4CAF50; }
        .badge-failed { background: rgba(244,67,54,0.2); color: #F44336; }
        .badge-refunded { background: rgba(96,125,139,0.2); color: #90A4AE; }
        
        .alert { padding: 12px 16px; border-radius: 8px; margin-bottom: 16px; }
        .alert-success { background: rgba(76,175,80,0.15); color: #81c784; border: 1px solid rgba(76,175,80,0.3); }
        .alert-error { background: rgba(244,67,54,0.15); color: #ef9a9a; border: 1px solid rgba(244,67,54,0.3); }
        
        @media (max-width: 768px) { .stats-grid { grid-template-columns: repeat(2, 1fr); } .form-row { grid-template-columns: 1fr; } }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-users"></i> 拼团活动管理</h2>
        </div>
        
        <% If gbActionMsg <> "" Then %>
        <div class="alert <% If gbActionResult Then %>alert-success<% Else %>alert-error<% End If %>">
            <i class="fas fa-<% If gbActionResult Then %>check-circle<% Else %>exclamation-circle<% End If %>"></i> <%= gbActionMsg %>
        </div>
        <% End If %>
        
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-value"><%= gbTotalPlans %></div>
                <div class="stat-label">全部计划</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%= gbActivePlans %></div>
                <div class="stat-label">进行中</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%= gbTotalGroups %></div>
                <div class="stat-label">总开团数</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%= gbSuccessGroups %></div>
                <div class="stat-label">已成团</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%= gbTotalParticipants %></div>
                <div class="stat-label">参团人次</div>
            </div>
        </div>
        
        <!-- 拼团计划表单 -->
        <div class="panel">
            <h3><i class="fas fa-<% If gbIsEditing Then %>edit<% Else %>plus-circle<% End If %>"></i> <% If gbIsEditing Then %>编辑拼团计划<% Else %>创建拼团计划<% End If %></h3>
            <form method="post">
                <% If gbIsEditing Then %>
                <input type="hidden" name="edit_id" value="<%= gbEditData("PlanID") %>">
                <% End If %>
                <div class="form-row">
                    <div class="form-group">
                        <label>选择产品 *</label>
                        <select name="product_id" required>
                            <option value="">-- 请选择产品 --</option>
                            <%
                            If Not rsGBProducts Is Nothing Then
                                Do While Not rsGBProducts.EOF
                                    Dim gpSelID : gpSelID = rsGBProducts("ProductID")
                                    Dim gpSelName : gpSelName = rsGBProducts("ProductName")
                                    Dim gpSelPrice : gpSelPrice = rsGBProducts("BasePrice")
                                    Dim gpSelSelected : gpSelSelected = ""
                                    If gbIsEditing Then
                                        If CLng(gbEditData("ProductID")) = CLng(gpSelID) Then gpSelSelected = " selected"
                                    End If
                            %>
                            <option value="<%= gpSelID %>"<%= gpSelSelected %>><%= Server.HTMLEncode(gpSelName) %> (&yen;<%= FormatNumber(gpSelPrice, 2) %>)</option>
                            <%
                                    rsGBProducts.MoveNext
                                Loop
                                rsGBProducts.MoveFirst
                            End If
                            %>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>成团人数 *</label>
                        <select name="team_size" required>
                            <%
                            Dim tsVals : tsVals = Array(2, 3, 5, 10)
                            Dim tsLabels : tsLabels = Array("2人团", "3人团", "5人团", "10人团")
                            Dim tsi, tsSel
                            For tsi = 0 To UBound(tsVals)
                                tsSel = ""
                                If gbIsEditing Then
                                    If CLng(gbEditData("TeamSize")) = tsVals(tsi) Then tsSel = " selected"
                                End If
                            %>
                            <option value="<%= tsVals(tsi) %>"<%= tsSel %>><%= tsLabels(tsi) %></option>
                            <% Next %>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>拼团价(每人) *</label>
                        <input type="number" name="group_price" step="0.01" value="<%= eGBGroupPrice %>" placeholder="如: 199.00" required>
                    </div>
                    <div class="form-group">
                        <label>成团有效期(小时)</label>
                        <input type="number" name="duration_hours" value="<%= eGBDuration %>" min="1">
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>开始时间 *</label>
                        <input type="datetime-local" name="start_time" value="<%= eGBStartTime %>" required>
                    </div>
                    <div class="form-group">
                        <label>结束时间 *</label>
                        <input type="datetime-local" name="end_time" value="<%= eGBEndTime %>" required>
                    </div>
                    <div class="form-group">
                        <label>最低开团数</label>
                        <input type="number" name="min_unit" value="<%= eGBMinUnit %>" min="0">
                    </div>
                    <div class="form-group">
                        <label>最高开团数(0=不限)</label>
                        <input type="number" name="max_unit" value="<%= eGBMaxUnit %>" min="0">
                    </div>
                    <div class="form-group">
                        <label>排序</label>
                        <input type="number" name="sort_order" value="<%= eGBSort %>" min="0">
                    </div>
                </div>
                <div style="margin-top:12px;display:flex;gap:8px;">
                    <button type="submit" name="action" value="save" class="btn btn-primary">
                        <i class="fas fa-save"></i> <% If gbIsEditing Then %>更新<% Else %>创建<% End If %>拼团计划
                    </button>
                    <% If gbIsEditing Then %>
                    <a href="group_buy.asp" class="btn btn-outline">取消编辑</a>
                    <% End If %>
                </div>
            </form>
        </div>
        
        <!-- 拼团计划列表 -->
        <div class="panel">
            <h3><i class="fas fa-list"></i> 拼团计划列表</h3>
            <table class="gb-table">
                <thead>
                    <tr><th>ID</th><th>产品</th><th>原价</th><th>拼团价</th><th>人数</th><th>有效期</th><th>时间</th><th>状态</th><th>操作</th></tr>
                </thead>
                <tbody>
                    <%
                    If Not rsPlans Is Nothing Then
                        Do While Not rsPlans.EOF
                            Dim plID, plName, plBasePrice, plGroupPrice, plTeamSize, plDuration, plStart, plEnd, plActive, plSort
                            plID = rsPlans("PlanID")
                            plName = rsPlans("ProductName")
                            plBasePrice = rsPlans("BasePrice")
                            plGroupPrice = rsPlans("GroupPrice")
                            plTeamSize = rsPlans("TeamSize")
                            plDuration = rsPlans("DurationHours")
                            plStart = rsPlans("StartTime")
                            plEnd = rsPlans("EndTime")
                            plActive = CBool(rsPlans("IsActive"))
                            plSort = rsPlans("SortOrder")
                            
                            Dim plStatus, plBadge
                            If Not plActive Then
                                plStatus = "已禁用" : plBadge = "badge-inactive"
                            ElseIf CDate(Now()) < CDate(plStart) Then
                                plStatus = "未开始" : plBadge = "badge-pending"
                            ElseIf CDate(Now()) > CDate(plEnd) Then
                                plStatus = "已结束" : plBadge = "badge-failed"
                            Else
                                plStatus = "进行中" : plBadge = "badge-active"
                            End If
                            
                            Dim plSave : plSave = CDbl(plBasePrice) - CDbl(plGroupPrice)
                    %>
                    <tr>
                        <td><%= plID %></td>
                        <td><strong><%= Server.HTMLEncode(plName) %></strong></td>
                        <td>&yen;<%= FormatNumber(plBasePrice, 2) %></td>
                        <td style="color:#7c4dff;font-weight:700;">&yen;<%= FormatNumber(plGroupPrice, 2) %></td>
                        <td><%= plTeamSize %>人</td>
                        <td><%= plDuration %>小时</td>
                        <td style="font-size:11px;"><%= FormatDateTime(plStart, 0) %><br>~<%= FormatDateTime(plEnd, 0) %></td>
                        <td><span class="badge <%= plBadge %>"><%= plStatus %></span></td>
                        <td>
                            <a href="?edit_id=<%= plID %>" class="btn btn-sm btn-primary"><i class="fas fa-edit"></i></a>
                            <a href="?action=toggle&id=<%= plID %>" class="btn btn-sm btn-outline"><i class="fas fa-power-off"></i></a>
                            <a href="?action=delete&id=<%= plID %>" class="btn btn-sm btn-danger" onclick="return confirm('确认删除？')"><i class="fas fa-trash"></i></a>
                        </td>
                    </tr>
                    <%
                            rsPlans.MoveNext
                        Loop
                    End If
                    %>
                </tbody>
            </table>
            <% If rsPlans Is Nothing Or rsPlans.EOF Then %>
            <div class="empty-state" style="padding:40px;text-align:center;color:#888;">
                <i class="fas fa-users" style="font-size:2rem;margin-bottom:10px;display:block;"></i>
                <p>暂无拼团计划</p>
            </div>
            <% End If %>
        </div>
        
        <!-- 拼团记录列表 -->
        <div class="panel">
            <h3><i class="fas fa-history"></i> 拼团记录</h3>
            <table class="gb-table">
                <thead>
                    <tr><th>团编号</th><th>产品</th><th>团长</th><th>进度</th><th>状态</th><th>创建时间</th><th>完成时间</th></tr>
                </thead>
                <tbody>
                    <%
                    If Not rsGroups Is Nothing Then
                        Do While Not rsGroups.EOF
                            Dim grpID, grpSN, grpInitiator, grpStatus, grpCur, grpTar, grpCreated, grpCompleted, grpInitiatorName, grpProductName
                            grpID = rsGroups("GroupID")
                            grpSN = rsGroups("GroupSN")
                            grpInitiatorName = rsGroups("InitiatorName")
                            grpStatus = CLng(rsGroups("Status"))
                            grpCur = rsGroups("CurrentSize")
                            grpTar = rsGroups("TargetSize")
                            grpCreated = rsGroups("CreatedAt")
                            grpCompleted = rsGroups("CompletedAt")
                            grpProductName = rsGroups("ProductName")
                            
                            Dim grpStLabel, grpStBadge
                            Select Case grpStatus
                                Case 0: grpStLabel = "进行中" : grpStBadge = "badge-pending"
                                Case 1: grpStLabel = "已成团" : grpStBadge = "badge-success"
                                Case 2: grpStLabel = "已失效" : grpStBadge = "badge-failed"
                                Case 3: grpStLabel = "已退款" : grpStBadge = "badge-refunded"
                            End Select
                            
                            If IsNull(grpInitiatorName) Then grpInitiatorName = "—"
                            If IsNull(grpProductName) Then grpProductName = "—"
                    %>
                    <tr>
                        <td><code style="background:rgba(255,255,255,0.05);padding:2px 6px;border-radius:3px;font-size:11px;"><%= grpSN %></code></td>
                        <td><%= Server.HTMLEncode(grpProductName) %></td>
                        <td><%= Server.HTMLEncode(grpInitiatorName) %></td>
                        <td><%= grpCur %>/<%= grpTar %>人</td>
                        <td><span class="badge <%= grpStBadge %>"><%= grpStLabel %></span></td>
                        <td style="font-size:11px;"><%= FormatDateTime(grpCreated, 0) %></td>
                        <td style="font-size:11px;"><%= IIf(IsNull(grpCompleted), "—", FormatDateTime(grpCompleted, 0)) %></td>
                    </tr>
                    <%
                            rsGroups.MoveNext
                        Loop
                    End If
                    %>
                </tbody>
            </table>
            <% If rsGroups Is Nothing Or rsGroups.EOF Then %>
            <div class="empty-state" style="padding:40px;text-align:center;color:#888;">
                <p>暂无拼团记录</p>
            </div>
            <% End If %>
        </div>
    </div>
</body>
</html>
<%
If Not rsPlans Is Nothing Then
    If rsPlans.State = 1 Then rsPlans.Close
    Set rsPlans = Nothing
End If
If Not rsGBProducts Is Nothing Then
    If rsGBProducts.State = 1 Then rsGBProducts.Close
    Set rsGBProducts = Nothing
End If
If Not rsGroups Is Nothing Then
    If rsGroups.State = 1 Then rsGroups.Close
    Set rsGroups = Nothing
End If
If Not gbEditData Is Nothing Then
    If gbEditData.State = 1 Then gbEditData.Close
    Set gbEditData = Nothing
End If
Call CloseConnection()
%>
