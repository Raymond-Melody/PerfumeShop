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

Dim subAction, subActionMsg, subActionResult
subAction = Request.QueryString("action")
If subAction = "" Then subAction = Request.Form("action")
subActionMsg = ""
subActionResult = True

' 新增/编辑订阅计划
If subAction = "save" And Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim sEditId, sName, sPeriod, sPrice, sSample, sFull, sFreeShip, sCancelFee, sDesc, sSort
    sEditId = Request.Form("edit_id")
    sName = Trim(Request.Form("plan_name"))
    sPeriod = Request.Form("period")
    sPrice = CDbl(Request.Form("price"))
    sSample = CLng(Request.Form("sample_count"))
    sFull = CLng(Request.Form("fullsize_count"))
    sFreeShip = IIf(Request.Form("free_shipping") = "1", "1", "0")
    sCancelFee = CDbl(Request.Form("cancellation_fee"))
    sDesc = Trim(Request.Form("description"))
    sSort = Request.Form("sort_order")
    If sSort = "" Or Not IsNumeric(sSort) Then sSort = 0

    If sName = "" Or sPeriod = "" Or sPrice <= 0 Then
        subActionMsg = "请填写计划名称、周期和价格"
        subActionResult = False
    Else
        On Error Resume Next
        If sEditId <> "" And IsNumeric(sEditId) Then
            conn.Execute "UPDATE SubscriptionPlans SET PlanName='" & SafeSQL(sName) & "', Period='" & SafeSQL(sPeriod) & _
                       "', Price=" & sPrice & ", SampleCount=" & sSample & ", FullSizeCount=" & sFull & _
                       ", FreeShipping=" & sFreeShip & ", CancellationFee=" & sCancelFee & _
                       ", Description='" & SafeSQL(sDesc) & "', SortOrder=" & sSort & " WHERE PlanID=" & sEditId
            If Err.Number = 0 Then subActionMsg = "更新成功" Else subActionMsg = "更新失败: " & Err.Description : subActionResult = False
        Else
            conn.Execute "INSERT INTO SubscriptionPlans (PlanName, Period, Price, SampleCount, FullSizeCount, FreeShipping, CancellationFee, Description, SortOrder) VALUES ('" & _
                       SafeSQL(sName) & "','" & SafeSQL(sPeriod) & "'," & sPrice & "," & sSample & "," & sFull & "," & _
                       sFreeShip & "," & sCancelFee & ",'" & SafeSQL(sDesc) & "'," & sSort & ")"
            If Err.Number = 0 Then subActionMsg = "创建成功" Else subActionMsg = "创建失败: " & Err.Description : subActionResult = False
        End If
        On Error GoTo 0
    End If
End If

' 删除
If subAction = "delete" Then
    Dim sDelId : sDelId = Request.QueryString("id")
    If IsNumeric(sDelId) Then
        conn.Execute "DELETE FROM SubscriptionPlans WHERE PlanID = " & sDelId
        subActionMsg = "已删除"
    End If
End If

' 切换状态
If subAction = "toggle" Then
    Dim sTogId : sTogId = Request.QueryString("id")
    If IsNumeric(sTogId) Then
        conn.Execute "UPDATE SubscriptionPlans SET IsActive = CASE WHEN IsActive = 1 THEN 0 ELSE 1 END WHERE PlanID = " & sTogId
        subActionMsg = "状态已切换"
    End If
End If

' 统计数据
Dim subTotalPlans, subActivePlans, subTotalSubs, subActiveSubs
subTotalPlans = GetScalar("SELECT COUNT(*) FROM SubscriptionPlans")
subActivePlans = GetScalar("SELECT COUNT(*) FROM SubscriptionPlans WHERE IsActive = 1")
subTotalSubs = GetScalar("SELECT COUNT(*) FROM UserSubscriptions")
subActiveSubs = GetScalar("SELECT COUNT(*) FROM UserSubscriptions WHERE Status = 0")

' 获取编辑中的计划
Dim subEditData, subEditId, subIsEditing
subEditId = Request.QueryString("edit_id")
Set subEditData = Nothing
subIsEditing = False
If subEditId <> "" And IsNumeric(subEditId) Then
    Set subEditData = conn.Execute("SELECT * FROM SubscriptionPlans WHERE PlanID = " & subEditId)
    If Not subEditData Is Nothing Then
        If Not subEditData.EOF Then
            subIsEditing = True
        End If
    End If
End If

' 预提取编辑数据到变量（VBScript IIf不短路，需提前取值）
Dim eSubPlanName, eSubPeriod, eSubPrice, eSubSample, eSubFull, eSubCancelFee, eSubSort, eSubDesc, eSubFreeShip
If subIsEditing Then
    eSubPlanName = subEditData("PlanName")
    eSubPeriod = subEditData("Period")
    eSubPrice = subEditData("Price")
    eSubSample = subEditData("SampleCount")
    eSubFull = subEditData("FullSizeCount")
    eSubCancelFee = subEditData("CancellationFee")
    eSubSort = subEditData("SortOrder")
    eSubDesc = subEditData("Description")
    eSubFreeShip = CBool(subEditData("FreeShipping"))
Else
    eSubPlanName = ""
    eSubPeriod = "monthly"
    eSubPrice = ""
    eSubSample = "3"
    eSubFull = "1"
    eSubCancelFee = "0"
    eSubSort = "0"
    eSubDesc = ""
    eSubFreeShip = False
End If

' 计划列表
Dim rsSubPlans : Set rsSubPlans = DAL_GetList("SELECT * FROM SubscriptionPlans ORDER BY SortOrder ASC, PlanID ASC", Null)

' 活跃订阅列表
Dim rsActiveSubs : Set rsActiveSubs = DAL_GetList("SELECT us.SubscriptionID, us.UserID, us.Status, us.StartDate, us.NextDeliveryDate, us.TotalDeliveries, us.AutoRenew, sp.PlanName, sp.Period, u.Username " & _
    "FROM UserSubscriptions us INNER JOIN SubscriptionPlans sp ON us.PlanID = sp.PlanID LEFT JOIN Users u ON us.UserID = u.UserID " & _
    "WHERE us.Status = 0 ORDER BY us.NextDeliveryDate ASC", Null)

Function PeriodLabel(period)
    Select Case LCase(period)
        Case "monthly": PeriodLabel = "月度"
        Case "quarterly": PeriodLabel = "季度"
        Case "yearly": PeriodLabel = "年度"
        Case Else: PeriodLabel = period
    End Select
End Function
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>订阅计划管理 - 运营管理中心</title>
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
        .page-title i { color: #00c853; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 24px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); padding: 20px; border-radius: 10px; border: 1px solid rgba(255,255,255,0.05); text-align: center; }
        .stat-value { font-size: 28px; font-weight: bold; color: #00c853; }
        .stat-label { font-size: 13px; color: #888; margin-top: 4px; }
        
        .panel { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; padding: 24px; border: 1px solid rgba(255,255,255,0.05); margin-bottom: 24px; }
        .panel h3 { color: #fff; margin: 0 0 16px; font-size: 18px; display: flex; align-items: center; gap: 8px; }
        .panel h3 i { color: #00c853; }
        
        .form-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 12px; margin-bottom: 12px; }
        .form-group { display: flex; flex-direction: column; gap: 4px; }
        .form-group label { font-size: 12px; color: #888; font-weight: 500; }
        .form-group input, .form-group select, .form-group textarea {
            padding: 8px 12px; border: 1px solid #3a3a4a; border-radius: 6px; background: #1a1a2e; color: #e0e0e0; font-size: 13px;
        }
        .form-group input:focus, .form-group select:focus { border-color: #00c853; outline: none; }
        .form-group textarea { resize: vertical; min-height: 60px; }
        
        .btn { padding: 8px 18px; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; font-weight: 500; transition: all 0.2s; text-decoration: none; display: inline-block; }
        .btn-primary { background: linear-gradient(135deg, #00c853, #009624); color: #fff; }
        .btn-primary:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(0,200,83,0.3); }
        .btn-danger { background: #c62828; color: #fff; }
        .btn-sm { padding: 4px 12px; font-size: 11px; }
        .btn-outline { background: transparent; border: 1px solid #555; color: #ccc; }
        
        .sub-table { width: 100%; border-collapse: collapse; }
        .sub-table th { text-align: left; padding: 10px 12px; background: rgba(0,0,0,0.2); color: #888; font-size: 11px; text-transform: uppercase; }
        .sub-table td { padding: 10px 12px; border-bottom: 1px solid rgba(255,255,255,0.05); font-size: 13px; }
        .sub-table tr:hover td { background: rgba(255,255,255,0.02); }
        
        .badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; }
        .badge-active { background: rgba(76,175,80,0.2); color: #4CAF50; }
        .badge-inactive { background: rgba(158,158,158,0.2); color: #9e9e9e; }
        .badge-monthly { background: rgba(33,150,243,0.2); color: #64B5F6; }
        .badge-quarterly { background: rgba(156,39,176,0.2); color: #CE93D8; }
        .badge-yearly { background: rgba(255,152,0,0.2); color: #FFB74D; }
        
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
            <h2 class="page-title"><i class="fas fa-box-open"></i> 订阅计划管理</h2>
        </div>
        
        <% If subActionMsg <> "" Then %>
        <div class="alert <% If subActionResult Then %>alert-success<% Else %>alert-error<% End If %>">
            <i class="fas fa-<% If subActionResult Then %>check-circle<% Else %>exclamation-circle<% End If %>"></i> <%= subActionMsg %>
        </div>
        <% End If %>
        
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-value"><%= subTotalPlans %></div>
                <div class="stat-label">全部计划</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%= subActivePlans %></div>
                <div class="stat-label">生效中</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%= subTotalSubs %></div>
                <div class="stat-label">总订阅数</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%= subActiveSubs %></div>
                <div class="stat-label">活跃订阅</div>
            </div>
        </div>
        
        <!-- 创建/编辑计划 -->
        <div class="panel">
            <h3><i class="fas fa-<% If subIsEditing Then %>edit<% Else %>plus-circle<% End If %>"></i> <% If subIsEditing Then %>编辑计划<% Else %>创建订阅计划<% End If %></h3>
            <form method="post">
                <% If subIsEditing Then %>
                <input type="hidden" name="edit_id" value="<%= subEditData("PlanID") %>">
                <% End If %>
                <div class="form-row">
                    <div class="form-group">
                        <label>计划名称 *</label>
                        <input type="text" name="plan_name" value="<%= eSubPlanName %>" placeholder="如: 月度探索盒" required>
                    </div>
                    <div class="form-group">
                        <label>周期 *</label>
                        <select name="period" required>
                            <%
                            Dim selPeriod : selPeriod = eSubPeriod
                            %>
                            <option value="monthly"<% If selPeriod = "monthly" Then %> selected<% End If %>>月度</option>
                            <option value="quarterly"<% If selPeriod = "quarterly" Then %> selected<% End If %>>季度</option>
                            <option value="yearly"<% If selPeriod = "yearly" Then %> selected<% End If %>>年度</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>每期价格 *</label>
                        <input type="number" name="price" step="0.01" value="<%= eSubPrice %>" placeholder="如: 199.00" required>
                    </div>
                    <div class="form-group">
                        <label>小样数量</label>
                        <input type="number" name="sample_count" value="<%= eSubSample %>" min="0">
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>正装数量</label>
                        <input type="number" name="fullsize_count" value="<%= eSubFull %>" min="0">
                    </div>
                    <div class="form-group">
                        <label>取消费用</label>
                        <input type="number" name="cancellation_fee" step="0.01" value="<%= eSubCancelFee %>">
                    </div>
                    <div class="form-group">
                        <label>排序</label>
                        <input type="number" name="sort_order" value="<%= eSubSort %>">
                    </div>
                    <div class="form-group">
                        <label>&nbsp;</label>
                        <label style="display:flex;align-items:center;gap:6px;font-size:13px;">
                            <input type="checkbox" name="free_shipping" value="1"<% If eSubFreeShip Or Not subIsEditing Then %> checked<% End If %>> 包邮
                        </label>
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group" style="grid-column: 1 / -1;">
                        <label>计划描述</label>
                        <textarea name="description" placeholder="描述订阅计划内容..."><%= eSubDesc %></textarea>
                    </div>
                </div>
                <div style="margin-top:12px;display:flex;gap:8px;">
                    <button type="submit" name="action" value="save" class="btn btn-primary">
                        <i class="fas fa-save"></i> <% If subIsEditing Then %>更新<% Else %>创建<% End If %>计划
                    </button>
                    <% If subIsEditing Then %>
                    <a href="subscription_plans.asp" class="btn btn-outline">取消编辑</a>
                    <% End If %>
                </div>
            </form>
        </div>
        
        <!-- 计划列表 -->
        <div class="panel">
            <h3><i class="fas fa-list"></i> 订阅计划列表</h3>
            <table class="sub-table">
                <thead>
                    <tr><th>ID</th><th>计划名称</th><th>周期</th><th>价格</th><th>小样/正装</th><th>包邮</th><th>取消费</th><th>状态</th><th>操作</th></tr>
                </thead>
                <tbody>
                    <%
                    If Not rsSubPlans Is Nothing Then
                        Do While Not rsSubPlans.EOF
                            Dim spID, spName, spPeriod, spPrice, spSample, spFull, spFree, spCancelFee, spActive, spSort
                            spID = rsSubPlans("PlanID")
                            spName = rsSubPlans("PlanName")
                            spPeriod = rsSubPlans("Period")
                            spPrice = rsSubPlans("Price")
                            spSample = rsSubPlans("SampleCount")
                            spFull = rsSubPlans("FullSizeCount")
                            spFree = CBool(rsSubPlans("FreeShipping"))
                            spCancelFee = rsSubPlans("CancellationFee")
                            spActive = CBool(rsSubPlans("IsActive"))
                            spSort = rsSubPlans("SortOrder")
                            
                            Dim spPerBadge
                            Select Case LCase(spPeriod)
                                Case "monthly": spPerBadge = "badge-monthly"
                                Case "quarterly": spPerBadge = "badge-quarterly"
                                Case "yearly": spPerBadge = "badge-yearly"
                                Case Else: spPerBadge = "badge-inactive"
                            End Select
                    %>
                    <tr>
                        <td><%= spID %></td>
                        <td><strong><%= Server.HTMLEncode(spName) %></strong></td>
                        <td><span class="badge <%= spPerBadge %>"><%= PeriodLabel(spPeriod) %></span></td>
                        <td style="color:#00c853;font-weight:700;">&yen;<%= FormatNumber(spPrice, 2) %></td>
                        <td><%= spSample %>+<%= spFull %></td>
                        <td><% If spFree Then %>✅<% Else %>❌<% End If %></td>
                        <td>&yen;<%= FormatNumber(spCancelFee, 2) %></td>
                        <td><span class="badge <% If spActive Then %>badge-active<% Else %>badge-inactive<% End If %>"><% If spActive Then %>启用<% Else %>禁用<% End If %></span></td>
                        <td>
                            <a href="?edit_id=<%= spID %>" class="btn btn-sm btn-primary"><i class="fas fa-edit"></i></a>
                            <a href="?action=toggle&id=<%= spID %>" class="btn btn-sm btn-outline"><i class="fas fa-power-off"></i></a>
                            <a href="?action=delete&id=<%= spID %>" class="btn btn-sm btn-danger" onclick="return confirm('确认删除？')"><i class="fas fa-trash"></i></a>
                        </td>
                    </tr>
                    <%
                            rsSubPlans.MoveNext
                        Loop
                    End If
                    %>
                </tbody>
            </table>
            <%
            Dim subPlansEmpty : subPlansEmpty = True
            If Not rsSubPlans Is Nothing Then
                If Not rsSubPlans.EOF Then subPlansEmpty = False
            End If
            If subPlansEmpty Then
            %>
            <div style="padding:40px;text-align:center;color:#888;">暂无订阅计划</div>
            <% End If %>
        </div>
        
        <!-- 活跃订阅列表 -->
        <div class="panel">
            <h3><i class="fas fa-users"></i> 活跃订阅用户</h3>
            <table class="sub-table">
                <thead>
                    <tr><th>ID</th><th>用户</th><th>计划</th><th>开始日期</th><th>下次配送</th><th>已完成</th><th>自动续费</th></tr>
                </thead>
                <tbody>
                    <%
                    If Not rsActiveSubs Is Nothing Then
                        Do While Not rsActiveSubs.EOF
                            Dim asID, asUser, asPlan, asStart, asNext, asCount, asAuto, asUsername
                            asID = rsActiveSubs("SubscriptionID")
                            asUser = rsActiveSubs("UserID")
                            asPlan = rsActiveSubs("PlanName")
                            asStart = rsActiveSubs("StartDate")
                            asNext = rsActiveSubs("NextDeliveryDate")
                            asCount = rsActiveSubs("TotalDeliveries")
                            asAuto = CBool(rsActiveSubs("AutoRenew"))
                            asUsername = rsActiveSubs("Username")
                            If IsNull(asUsername) Then asUsername = "—"
                    %>
                    <tr>
                        <td><%= asID %></td>
                        <td><%= Server.HTMLEncode(asUsername) %></td>
                        <td><%= Server.HTMLEncode(asPlan) %></td>
                        <td style="font-size:11px;"><%= FormatDateTime(asStart, 0) %></td>
                        <td style="font-size:11px;color:#00c853;"><%= FormatDateTime(asNext, 0) %></td>
                        <td><%= asCount %> 次</td>
                        <td><% If asAuto Then %><span class="badge badge-active">是</span><% Else %><span class="badge badge-inactive">否</span><% End If %></td>
                    </tr>
                    <%
                            rsActiveSubs.MoveNext
                        Loop
                    End If
                    %>
                </tbody>
            </table>
            <%
            Dim activeSubsEmpty : activeSubsEmpty = True
            If Not rsActiveSubs Is Nothing Then
                If Not rsActiveSubs.EOF Then activeSubsEmpty = False
            End If
            If activeSubsEmpty Then
            %>
            <div style="padding:40px;text-align:center;color:#888;">暂无活跃订阅</div>
            <% End If %>
        </div>
    </div>
</body>
</html>
<%
If Not rsSubPlans Is Nothing Then
    If rsSubPlans.State = 1 Then rsSubPlans.Close
    Set rsSubPlans = Nothing
End If
If Not rsActiveSubs Is Nothing Then
    If rsActiveSubs.State = 1 Then rsActiveSubs.Close
    Set rsActiveSubs = Nothing
End If
If Not subEditData Is Nothing Then
    If subEditData.State = 1 Then subEditData.Close
    Set subEditData = Nothing
End If
Call CloseConnection()
%>
