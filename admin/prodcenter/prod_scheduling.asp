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

' 确保 OrderItems 表存在（V8新增表）
On Error Resume Next
conn.Execute "IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='OrderItems') CREATE TABLE OrderItems (OrderItemID INT IDENTITY(1,1) PRIMARY KEY, OrderID INT NOT NULL, ProductID INT NULL, Quantity INT DEFAULT 1, UnitPrice DECIMAL(10,2) DEFAULT 0, CreatedAt DATETIME DEFAULT GETDATE())"
Err.Clear
On Error GoTo 0

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

Dim msg, msgType
msg = Trim(Request.QueryString("msg"))
msgType = "success"
If InStr(msg, "失败") > 0 Or InStr(msg, "错误") > 0 Then msgType = "error"

' ========== POST 处理：批量排产 ==========
Dim batchResult, batchSuccess, batchFail
batchResult = "" : batchSuccess = 0 : batchFail = 0

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim scAction : scAction = Trim(Request.Form("action"))
    
    If scAction = "batch_schedule" Or scAction = "batch_schedule_date" Then
        Dim orderIds, dateFrom, dateTo
        Dim scPriority, scAssignedTo
        scPriority = SafeNum(Request.Form("batch_priority"))
        If scPriority = 0 Then scPriority = 1
        scAssignedTo = Trim(Request.Form("batch_assigned_to"))
        
        ' 构建订单ID列表
        If scAction = "batch_schedule" Then
            orderIds = Trim(Request.Form("order_ids"))
        Else
            ' 按日期范围获取订单
            dateFrom = Trim(Request.Form("date_from"))
            dateTo = Trim(Request.Form("date_to"))
            If dateFrom = "" Then dateFrom = Year(Date()) & "-" & Right("0" & Month(Date()), 2) & "-" & Right("0" & Day(Date()), 2)
            If dateTo = "" Then dateTo = Year(Date()) & "-" & Right("0" & Month(Date()), 2) & "-" & Right("0" & Day(Date()), 2)
            
            Dim rsDateRange
            Set rsDateRange = conn.Execute("SELECT STRING_AGG(OrderID, ',') AS IDS FROM Orders WHERE Status IN ('Paid','Processing') AND NOT EXISTS (SELECT 1 FROM ProductionOrders po WHERE po.OrderID=Orders.OrderID AND po.Status<>'Cancelled') AND CAST(CreatedAt AS DATE) BETWEEN '" & SafeSQL(dateFrom) & "' AND '" & SafeSQL(dateTo) & "'")
            If Not rsDateRange Is Nothing Then
                If Not rsDateRange.EOF Then
                    orderIds = rsDateRange("IDS") & ""
                End If
                rsDateRange.Close
            End If
            Set rsDateRange = Nothing
        End If
        
        If orderIds <> "" Then
            Dim idArray, i, currentOrderId
            idArray = Split(orderIds, ",")
            Dim batchNo
            batchNo = "BATCH" & Year(Now) & Right("0"&Month(Now),2) & Right("0"&Day(Now),2) & Right("0"&Hour(Now),2) & Right("0"&Minute(Now),2)
            
            For i = 0 To UBound(idArray)
                currentOrderId = Trim(idArray(i))
                If IsNumeric(currentOrderId) And CLng(currentOrderId) > 0 Then
                    On Error Resume Next
                    
                    ' 获取订单的产品和配方信息
                    Dim rsOrderInfo
                    Set rsOrderInfo = conn.Execute("SELECT TOP 1 o.OrderNo, oi.ProductID, p.RecipeID, p.ProductName, oi.Quantity FROM Orders o LEFT JOIN OrderItems oi ON o.OrderID=oi.OrderID LEFT JOIN Products p ON oi.ProductID=p.ProductID WHERE o.OrderID=" & currentOrderId)
                    
                    Dim scProductName, scRecipeID, scQty, scOrderNo
                    scProductName = "" : scRecipeID = 0 : scQty = 1 : scOrderNo = ""
                    If Not rsOrderInfo Is Nothing Then
                        If Not rsOrderInfo.EOF Then
                            scOrderNo = rsOrderInfo("OrderNo") & ""
                            scProductName = rsOrderInfo("ProductName") & ""
                            scRecipeID = SafeNum(rsOrderInfo("RecipeID"))
                            scQty = SafeNum(rsOrderInfo("Quantity"))
                            If scQty = 0 Then scQty = 1
                        End If
                        rsOrderInfo.Close
                    End If
                    Set rsOrderInfo = Nothing
                    
                    ' 检查是否已存在有效工单
                    Dim existingPO
                    existingPO = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE OrderID=" & currentOrderId & " AND Status<>'Cancelled'"))
                    
                    If existingPO = 0 Then
                        Dim scWorkNo
                        scWorkNo = "PO" & Year(Now) & Right("0"&Month(Now),2) & Right("0"&Day(Now),2) & Right("0"&Hour(Now),2) & Right("0"&Minute(Now),2) & Right("0"&Second(Now),2) & Right("00" & i, 2)
                        
                        conn.Execute "INSERT INTO ProductionOrders (WorkOrderNo, OrderID, RecipeID, RecipeName, PlannedQty, Priority, Status, BatchNo, Notes, AssignedTo, CreatedAt, UpdatedAt) VALUES ('" & _
                            scWorkNo & "'," & currentOrderId & "," & scRecipeID & ",'" & SafeSQL(scProductName) & "'," & scQty & "," & scPriority & ",'Pending','" & SafeSQL(batchNo) & "','批量排产','" & SafeSQL(scAssignedTo) & "',GETDATE(),GETDATE())"
                        
                        ' 插入生产日志
                        conn.Execute "INSERT INTO ProductionLogs (ProductionID, Status, Notes, CreatedBy, CreatedAt) SELECT SCOPE_IDENTITY(), 'Pending', '批量排产生成工单，批次号:" & SafeSQL(batchNo) & "', '" & SafeSQL(Session("AdminUsername")) & "', GETDATE()"
                        
                        ' 更新订单状态
                        conn.Execute "UPDATE Orders SET Status='Processing', UpdatedAt=GETDATE() WHERE OrderID=" & currentOrderId
                        
                        If Err.Number = 0 Then
                            batchSuccess = batchSuccess + 1
                        Else
                            batchFail = batchFail + 1
                            Err.Clear
                        End If
                    Else
                        batchFail = batchFail + 1
                    End If
                    On Error GoTo 0
                End If
            Next
            
            batchResult = "批量排产完成：成功 " & batchSuccess & " 条，跳过 " & batchFail & " 条（可能已存在工单）"
            If batchFail > 0 Then msgType = "error"
        Else
            batchResult = "未找到符合条件的订单"
            msgType = "error"
        End If
    End If
End If

If batchResult <> "" Then msg = batchResult

' ========== 数据查询 ==========

' 确保批次号字段存在
On Error Resume Next
conn.Execute "SELECT BatchNo FROM ProductionOrders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE ProductionOrders ADD BatchNo NVARCHAR(30)"
On Error GoTo 0

' 统计
Dim scPending, scInProgress, scToday, scTotalOrders
scPending = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE Status='Pending'"))
scInProgress = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE Status='InProgress'"))
scToday = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE Status='Completed' AND CompletedAt >= CAST(GETDATE() AS DATE)"))
scTotalOrders = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders o WHERE o.Status IN ('Paid','Processing') AND NOT EXISTS (SELECT 1 FROM ProductionOrders po WHERE po.OrderID=o.OrderID AND po.Status<>'Cancelled')"))

' 待排产订单（含产品信息）
Dim rsPendingOrders
Set rsPendingOrders = conn.Execute("SELECT TOP 50 o.OrderID, o.OrderNo, o.Status, o.CreatedAt, o.TotalAmount, u.Username, " & _
    "(SELECT TOP 1 p.ProductName FROM OrderItems oi LEFT JOIN Products p ON oi.ProductID=p.ProductID WHERE oi.OrderID=o.OrderID) AS ProductName " & _
    "FROM Orders o LEFT JOIN Users u ON o.UserID=u.UserID " & _
    "WHERE o.Status IN ('Paid','Processing') AND NOT EXISTS (SELECT 1 FROM ProductionOrders po WHERE po.OrderID=o.OrderID AND po.Status<>'Cancelled') " & _
    "ORDER BY o.CreatedAt ASC")

' 排产中的工单
Dim rsScheduled
Set rsScheduled = conn.Execute("SELECT TOP 30 po.*, o.OrderNo FROM ProductionOrders po LEFT JOIN Orders o ON po.OrderID=o.OrderID WHERE po.Status IN ('Pending','InProgress') ORDER BY po.Priority DESC, po.CreatedAt ASC")

' 可用负责人列表（生产角色）
Dim rsAssignees
Set rsAssignees = conn.Execute("SELECT u.AdminID, u.FullName FROM AdminUsers u LEFT JOIN AdminRoles r ON u.RoleID=r.RoleID WHERE r.RoleCode LIKE 'PROD%' OR r.RoleCode='SUPER_ADMIN' ORDER BY u.FullName")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>排产调度 - 产品生产中心</title>
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
        .page-header { margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px; }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #2196F3; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 20px; border-radius: 12px; text-align: center; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .num { font-size: 28px; font-weight: bold; }
        .stat-card .label { font-size: 12px; color: #888; margin-top: 5px; }
        
        .alert { padding: 15px 20px; border-radius: 8px; margin-bottom: 20px; font-size: 14px; }
        .alert-success { background: rgba(76,175,80,0.12); color: #81c784; border: 1px solid rgba(76,175,80,0.25); }
        .alert-error { background: rgba(244,67,54,0.12); color: #e57373; border: 1px solid rgba(244,67,54,0.25); }
        
        .toolbar { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; padding: 16px 20px; margin-bottom: 20px; border: 1px solid rgba(255,255,255,0.06); display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
        .toolbar .section-divider { width: 1px; height: 30px; background: rgba(255,255,255,0.08); margin: 0 8px; }
        .toolbar label { font-size: 12px; color: #888; white-space: nowrap; }
        .toolbar select, .toolbar input { padding: 8px 12px; background: #2d2d44; border: 1px solid rgba(255,255,255,0.12); border-radius: 6px; color: #e0e0e0; font-size: 13px; }
        .toolbar select:focus, .toolbar input:focus { outline: none; border-color: #2196F3; }
        .toolbar input[type="date"] { color-scheme: dark; }
        

        .info-cards { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        .card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; border: 1px solid rgba(255,255,255,0.06); overflow: hidden; }
        .card-header { padding: 14px 20px; font-weight: 600; font-size: 15px; color: #e0e0e0; border-bottom: 1px solid rgba(255,255,255,0.06); display: flex; justify-content: space-between; align-items: center; }
        .card-body { padding: 16px 20px; overflow-x: auto; }
        
        table { width: 100%; border-collapse: collapse; }
        th { padding: 10px 12px; text-align: left; font-weight: 600; font-size: 12px; color: #888; border-bottom: 1px solid rgba(255,255,255,0.04); white-space: nowrap; }
        td { padding: 10px 12px; border-bottom: 1px solid rgba(255,255,255,0.03); color: #e0e0e0; font-size: 13px; }
        tr.selected td { background: rgba(33,150,243,0.1); }
        
        .status-badge { display: inline-block; padding: 3px 10px; border-radius: 10px; font-size: 11px; }
        .badge-paid { background: rgba(33,150,243,0.15); color: #64b5f6; }
        .badge-progress { background: rgba(255,152,0,0.15); color: #ffb74d; }
        .badge-pending { background: rgba(255,152,0,0.12); color: #ffb74d; }
        
        .text-center { text-align: center; }
        .text-muted { color: #888; font-size: 13px; }
        
        .batch-summary { background: rgba(33,150,243,0.08); border: 1px solid rgba(33,150,243,0.2); border-radius: 8px; padding: 12px 16px; margin-bottom: 15px; font-size: 13px; display: none; }
        .batch-summary.active { display: block; }
        .batch-summary .count { font-weight: bold; color: #64b5f6; }
        
        .date-range-panel { display: none; align-items: center; gap: 10px; margin-top: 8px; }
        .date-range-panel.active { display: flex; }
        
        .checkbox-cell { width: 40px; text-align: center; }
        .checkbox-cell input[type="checkbox"] { width: 16px; height: 16px; accent-color: #2196F3; cursor: pointer; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-calendar-alt"></i> 排产调度</h2>
        </div>
        
        <% If msg <> "" Then %>
        <div class="alert alert-<%=msgType%>"><i class="fas fa-<%=IIf(msgType="success","check-circle","exclamation-circle")%>"></i> <%=Server.HTMLEncode(msg)%></div>
        <% End If %>
        
        <div class="stats-grid">
            <div class="stat-card"><span class="num" style="color:#FF9800;"><%=scTotalOrders%></span><span class="label">待排产订单</span></div>
            <div class="stat-card"><span class="num" style="color:#FF9800;"><%=scPending%></span><span class="label">待生产工单</span></div>
            <div class="stat-card"><span class="num" style="color:#2196F3;"><%=scInProgress%></span><span class="label">进行中</span></div>
            <div class="stat-card"><span class="num" style="color:#4CAF50;"><%=scToday%></span><span class="label">今日完成</span></div>
        </div>
        
        <!-- 批量操作工具栏 -->
        <form method="post" id="batchForm">
            <input type="hidden" name="action" id="batchAction" value="batch_schedule">
            <input type="hidden" name="order_ids" id="selectedOrderIds" value="">
            
            <div class="toolbar">
                <i class="fas fa-cogs" style="color:#2196F3;"></i>
                <label>排产操作：</label>
                <button type="button" class="btn btn-success" onclick="doBatchSchedule()" id="btnBatchSchedule" disabled><i class="fas fa-play"></i> 批量排产（选中订单）</button>
                <button type="button" class="btn btn-outline" onclick="toggleDateRange()"><i class="fas fa-calendar-week"></i> 按日期范围排产</button>
                
                <span class="section-divider"></span>
                <label>默认优先级：</label>
                <select name="batch_priority" style="width:100px;">
                    <option value="1">普通</option>
                    <option value="2" selected>优先</option>
                    <option value="3">紧急</option>
                </select>
                
                <label>负责人：</label>
                <select name="batch_assigned_to">
                    <option value="">-- 未指定 --</option>
                    <% If Not rsAssignees Is Nothing Then
                        Do While Not rsAssignees.EOF %>
                    <option value="<%=rsAssignees("FullName") & ""%>"><%=rsAssignees("FullName") & ""%></option>
                    <%      rsAssignees.MoveNext
                        Loop
                        rsAssignees.Close
                    End If
                    Set rsAssignees = Nothing %>
                </select>
                
                <div id="dateRangePanel" class="date-range-panel">
                    <span class="section-divider"></span>
                    <label>从：</label>
                    <input type="date" name="date_from" value="<%=Year(Date()) & "-" & Right("0" & Month(Date()), 2) & "-" & Right("0" & Day(Date()), 2)%>">
                    <label>至：</label>
                    <input type="date" name="date_to" value="<%=Year(Date()) & "-" & Right("0" & Month(Date()), 2) & "-" & Right("0" & Day(Date()), 2)%>">
                    <button type="button" class="btn btn-warning" onclick="doDateRangeSchedule()"><i class="fas fa-calendar-check"></i> 执行日期范围排产</button>
                </div>
            </div>
            
            <div class="batch-summary" id="batchSummary">
                <i class="fas fa-check-square"></i> 已选择 <span class="count" id="selectedCount">0</span> 个订单
            </div>
        </form>
        
        <div class="info-cards">
            <!-- 待排产订单 -->
            <div class="card">
                <div class="card-header" style="background:rgba(255,152,0,0.08);">
                    <span><i class="fas fa-clock"></i> 待排产订单</span>
                    <span style="font-size:12px;color:#888;"><label style="cursor:pointer;"><input type="checkbox" id="selectAll" onclick="toggleSelectAll()" style="margin-right:5px;">全选</label></span>
                </div>
                <div class="card-body">
                    <table>
                        <thead><tr><th class="checkbox-cell"><input type="checkbox" id="selectAllTop" onclick="toggleSelectAll()"></th><th>订单号</th><th>客户</th><th>产品</th><th>金额</th><th>状态</th><th>时间</th></tr></thead>
                        <tbody>
                        <%
                        Dim scRow1 : scRow1 = 0
                        If Not rsPendingOrders Is Nothing Then
                            Do While Not rsPendingOrders.EOF
                                scRow1 = scRow1 + 1
                                Dim scOid : scOid = rsPendingOrders("OrderID")
                        %>
                            <tr id="row_<%=scOid%>">
                                <td class="checkbox-cell"><input type="checkbox" value="<%=scOid%>" class="order-checkbox" onchange="updateSelection()"></td>
                                <td><strong><%=rsPendingOrders("OrderNo") & ""%></strong></td>
                                <td><%=rsPendingOrders("Username") & ""%></td>
                                <td><%=Left(rsPendingOrders("ProductName") & "", 20)%></td>
                                <td>¥<%=FormatNumber(SafeNum(rsPendingOrders("TotalAmount")),2)%></td>
                                <td><span class="status-badge badge-paid"><%=rsPendingOrders("Status")%></span></td>
                                <td class="text-muted"><%=IIF(IsNull(rsPendingOrders("CreatedAt")),"",Left(rsPendingOrders("CreatedAt"),10))%></td>
                            </tr>
                        <%
                                rsPendingOrders.MoveNext
                            Loop
                            rsPendingOrders.Close
                        End If
                        Set rsPendingOrders = Nothing
                        If scRow1 = 0 Then %>
                            <tr><td colspan="7" class="text-center text-muted" style="padding:40px;"><i class="fas fa-check-circle" style="font-size:32px;color:rgba(76,175,80,0.3);display:block;margin-bottom:10px;"></i>暂无待排产订单</td></tr>
                        <% End If %>
                        </tbody>
                    </table>
                </div>
            </div>
            
            <!-- 排产中的工单 -->
            <div class="card">
                <div class="card-header" style="background:rgba(33,150,243,0.08);"><i class="fas fa-tasks"></i> 排产进行中</div>
                <div class="card-body">
                    <table>
                        <thead><tr><th>工单号</th><th>订单号</th><th>配方</th><th>计划量</th><th>状态</th><th>优先级</th><th>负责人</th></tr></thead>
                        <tbody>
                        <%
                        Dim scRow2 : scRow2 = 0
                        If Not rsScheduled Is Nothing Then
                            Do While Not rsScheduled.EOF
                                scRow2 = scRow2 + 1
                        %>
                            <tr>
                                <td><strong><%=rsScheduled("WorkOrderNo") & ""%></strong></td>
                                <td><%=rsScheduled("OrderNo") & ""%></td>
                                <td><%=rsScheduled("RecipeName") & ""%></td>
                                <td><%=rsScheduled("PlannedQty") & ""%></td>
                                <td><span class="status-badge <%=IIF(rsScheduled("Status")&""="Pending","badge-pending","badge-progress")%>"><%=rsScheduled("Status")%></span></td>
                                <td><%=IIF(IsNull(rsScheduled("Priority")),"-",rsScheduled("Priority"))%></td>
                                <td><%=rsScheduled("AssignedTo") & ""%></td>
                            </tr>
                        <%
                                rsScheduled.MoveNext
                            Loop
                            rsScheduled.Close
                        End If
                        Set rsScheduled = Nothing
                        If scRow2 = 0 Then %>
                            <tr><td colspan="7" class="text-center text-muted" style="padding:40px;">暂无进行中工单</td></tr>
                        <% End If %>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    
    <script>
    function toggleSelectAll() {
        var checkboxes = document.querySelectorAll('.order-checkbox');
        var selectAll = document.getElementById('selectAll');
        var selectAllTop = document.getElementById('selectAllTop');
        var isChecked = selectAll.checked || selectAllTop.checked;
        
        selectAll.checked = !selectAll.checked;
        selectAllTop.checked = selectAll.checked;
        
        checkboxes.forEach(function(cb) { cb.checked = selectAll.checked; });
        updateSelection();
    }
    
    function updateSelection() {
        var checkboxes = document.querySelectorAll('.order-checkbox:checked');
        var ids = [];
        checkboxes.forEach(function(cb) { ids.push(cb.value); });
        
        document.getElementById('selectedOrderIds').value = ids.join(',');
        document.getElementById('selectedCount').textContent = ids.length;
        document.getElementById('btnBatchSchedule').disabled = (ids.length === 0);
        
        var summary = document.getElementById('batchSummary');
        if (ids.length > 0) {
            summary.classList.add('active');
        } else {
            summary.classList.remove('active');
        }
        
        // 高亮选中行
        document.querySelectorAll('.order-checkbox').forEach(function(cb) {
            var row = document.getElementById('row_' + cb.value);
            if (row) {
                if (cb.checked) { row.classList.add('selected'); }
                else { row.classList.remove('selected'); }
            }
        });
    }
    
    function doBatchSchedule() {
        var ids = document.getElementById('selectedOrderIds').value;
        if (!ids) { alert('请至少选择一个订单'); return; }
        if (!confirm('确认要为选中的订单批量创建生产工单吗？\n\n选中订单数：' + document.getElementById('selectedCount').textContent)) return;
        document.getElementById('batchAction').value = 'batch_schedule';
        document.getElementById('batchForm').submit();
    }
    
    function toggleDateRange() {
        var panel = document.getElementById('dateRangePanel');
        panel.classList.toggle('active');
    }
    
    function doDateRangeSchedule() {
        if (!confirm('确认要为指定日期范围内的所有待排产订单创建工单吗？')) return;
        document.getElementById('batchAction').value = 'batch_schedule_date';
        document.getElementById('batchForm').submit();
    }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
