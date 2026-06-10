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

' ========== 确保 ProductionOrders 必要字段存在 ==========
On Error Resume Next
conn.Execute "SELECT PlannedQty FROM ProductionOrders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE ProductionOrders ADD PlannedQty INT DEFAULT 0"
conn.Execute "SELECT AssignedTo FROM ProductionOrders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE ProductionOrders ADD AssignedTo NVARCHAR(50)"
conn.Execute "SELECT RecipeName FROM ProductionOrders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE ProductionOrders ADD RecipeName NVARCHAR(200)"
conn.Execute "SELECT BatchNo FROM ProductionOrders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE ProductionOrders ADD BatchNo NVARCHAR(30)"
conn.Execute "SELECT ProductionID FROM ProductionLogs WHERE 1=0"
If Err.Number <> 0 Then Err.Clear

' 确保 OrderItems 表存在（V8新增表）
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
        If Not rs Is Nothing Then If Not rs.EOF Then val = rs(0) : rs.Close
    Else : Err.Clear
    End If
    Set rs = Nothing : GetScalar = val
End Function

Dim action, msg, msgType
action = Trim(Request.Form("action"))
msg = Trim(Request.QueryString("msg"))
msgType = "success"
If InStr(msg, "失败") > 0 Or InStr(msg, "错误") > 0 Then msgType = "error"

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If action = "update_status" Then
        Dim poID, poStatus
        poID = SafeNum(Request.Form("production_id"))
        poStatus = Trim(Request.Form("new_status"))
        If poID > 0 And poStatus <> "" Then
            Dim extraSQL : extraSQL = ""
            If poStatus = "InProgress" Then extraSQL = ", StartedAt=GETDATE()"
            If poStatus = "Completed" Then extraSQL = ", CompletedAt=GETDATE()"
            conn.Execute "UPDATE ProductionOrders SET Status='" & SafeSQL(poStatus) & "'" & extraSQL & ", UpdatedAt=GETDATE() WHERE ProductionID=" & poID
            
            conn.Execute "INSERT INTO ProductionLogs (ProductionID, Status, CreatedBy, Notes, CreatedAt) VALUES (" & _
                poID & ",'" & SafeSQL(poStatus) & "','" & SafeSQL(Session("AdminUsername")) & "','状态更新为: " & SafeSQL(poStatus) & "',GETDATE())"
            Response.Redirect "production_management.asp?msg=状态已更新"
            Response.End
        End If
    ElseIf action = "create_po" Then
        Dim coOrderID, coRecipeID, coQty, coPriority, coNotes, coRecipeName, coAssignedTo
        coOrderID = SafeNum(Request.Form("order_id"))
        coRecipeID = SafeNum(Request.Form("recipe_id"))
        coQty = SafeNum(Request.Form("planned_qty"))
        coPriority = SafeNum(Request.Form("priority"))
        coNotes = Trim(Request.Form("notes"))
        coRecipeName = Trim(Request.Form("recipe_name"))
        coAssignedTo = Trim(Request.Form("assigned_to"))
        
        If coQty > 0 Then
            Dim coWorkNo
            coWorkNo = "PO" & Year(Now) & Right("0"&Month(Now),2) & Right("0"&Day(Now),2) & Right("0"&Hour(Now),2) & Right("0"&Minute(Now),2) & Right("0"&Second(Now),2)
            conn.Execute "INSERT INTO ProductionOrders (WorkOrderNo, OrderID, RecipeID, RecipeName, PlannedQty, Priority, Status, Notes, AssignedTo, CreatedAt, UpdatedAt) VALUES ('" & _
                coWorkNo & "'," & coOrderID & "," & coRecipeID & ",'" & SafeSQL(coRecipeName) & "'," & coQty & "," & coPriority & ",'Pending','" & SafeSQL(coNotes) & "','" & SafeSQL(coAssignedTo) & "',GETDATE(),GETDATE())"
            
            ' 插入生产日志
            conn.Execute "INSERT INTO ProductionLogs (ProductionID, Status, Notes, CreatedBy, CreatedAt) SELECT SCOPE_IDENTITY(), 'Pending', '手动创建工单', '" & SafeSQL(Session("AdminUsername")) & "', GETDATE()"
            
            ' 更新订单状态
            If coOrderID > 0 Then
                conn.Execute "UPDATE Orders SET Status='Processing', UpdatedAt=GETDATE() WHERE OrderID=" & coOrderID
            End If
            
            Response.Redirect "production_management.asp?msg=工单已创建：" & coWorkNo
            Response.End
        End If
    End If
End If

' 统计
Dim pmPending, pmProgress, pmCompleted, pmTotal
pmPending = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE Status='Pending'"))
pmProgress = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE Status='InProgress'"))
pmCompleted = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE Status='Completed'"))
pmTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders"))

' 预填充订单ID（来自URL参数）
Dim pmPreOrderID
pmPreOrderID = SafeNum(Request.QueryString("orderId"))
If pmPreOrderID = 0 Then pmPreOrderID = SafeNum(Request.QueryString("id"))

' 待排产订单下拉数据
Dim rsOrderOptions
Set rsOrderOptions = conn.Execute("SELECT o.OrderID, o.OrderNo, oi.ProductID, p.ProductName, p.RecipeID, oi.Quantity FROM Orders o LEFT JOIN OrderItems oi ON o.OrderID=oi.OrderID LEFT JOIN Products p ON oi.ProductID=p.ProductID WHERE o.Status IN ('Paid','Processing') AND NOT EXISTS (SELECT 1 FROM ProductionOrders po WHERE po.OrderID=o.OrderID AND po.Status<>'Cancelled') ORDER BY o.CreatedAt ASC")

' 可用负责人列表
Dim rsAssigneesPM
Set rsAssigneesPM = conn.Execute("SELECT u.AdminID, u.FullName FROM AdminUsers u LEFT JOIN AdminRoles r ON u.RoleID=r.RoleID WHERE r.RoleCode LIKE 'PROD%' OR r.RoleCode='SUPER_ADMIN' ORDER BY u.FullName")

' 生产工单列表
Dim filterStatus
filterStatus = Trim(Request.QueryString("status"))
Dim pmSQL
pmSQL = "SELECT po.*, o.OrderNo FROM ProductionOrders po LEFT JOIN Orders o ON po.OrderID=o.OrderID WHERE 1=1"
If filterStatus <> "" Then pmSQL = pmSQL & " AND po.Status='" & SafeSQL(filterStatus) & "'"
pmSQL = pmSQL & " ORDER BY po.CreatedAt DESC"
Dim rsPM
Set rsPM = conn.Execute(pmSQL)

' 构建订单选项JSON数据（简化：避免VBScript内嵌引号转义问题）
Dim orderDataJSON : orderDataJSON = ""
Dim Q : Q = Chr(34)  ' 双引号字符
If Not rsOrderOptions Is Nothing Then
    Do While Not rsOrderOptions.EOF
        If orderDataJSON <> "" Then orderDataJSON = orderDataJSON & ","
        Dim safeProduct : safeProduct = Replace(Replace(rsOrderOptions("ProductName") & "", Q, ""), vbCr, "") & ""
        safeProduct = Replace(safeProduct, vbLf, "")
        orderDataJSON = orderDataJSON & Q & rsOrderOptions("OrderID") & Q & ":{" & Q & "product" & Q & ":" & Q & safeProduct & Q & "," & Q & "recipe" & Q & ":" & SafeNum(rsOrderOptions("RecipeID")) & "," & Q & "qty" & Q & ":" & SafeNum(rsOrderOptions("Quantity")) & "}"
        rsOrderOptions.MoveNext
    Loop
    If Not rsOrderOptions.EOF Then rsOrderOptions.MoveFirst
End If
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>生产工单管理 - 产品生产中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --accent: #4CAF50; --input-bg: #2d2d44; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #1a1a2e; color: var(--text); min-height: 100vh; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-header { margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px; }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #4CAF50; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 20px; border-radius: 12px; text-align: center; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .num { font-size: 28px; font-weight: bold; }
        .stat-card .label { font-size: 12px; color: #888; margin-top: 5px; }
        
        .filter-tabs { display: flex; gap: 8px; margin-bottom: 15px; }
        .filter-tab { padding: 6px 16px; border-radius: 16px; font-size: 13px; text-decoration: none; color: #888; background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.08); }
        .filter-tab:hover, .filter-tab.active { color: #fff; background: #4CAF50; border-color: #4CAF50; }
        
        .card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; margin-bottom: 25px; border: 1px solid rgba(255,255,255,0.06); }
        .card-header { padding: 14px 20px; background: rgba(76,175,80,0.08); border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 15px; color: #e0e0e0; display: flex; justify-content: space-between; }
        .card-body { padding: 16px 20px; overflow-x: auto; }
        
        table { width: 100%; border-collapse: collapse; }
        th { background: rgba(76,175,80,0.15); color: #81c784; padding: 12px 15px; text-align: left; font-weight: 600; font-size: 13px; }
        td { padding: 12px 15px; border-bottom: 1px solid rgba(255,255,255,0.04); color: #e0e0e0; font-size: 14px; }
        tr:hover td { background: rgba(255,255,255,0.03); }
        
        .status-badge { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 500; }
        .status-pending { background: rgba(255,152,0,0.15); color: #ffb74d; }
        .status-progress { background: rgba(33,150,243,0.15); color: #64b5f6; }
        .status-completed { background: rgba(76,175,80,0.15); color: #81c784; }
        

        .alert { padding: 15px 20px; border-radius: 8px; margin-bottom: 20px; }
        .alert-success { background: rgba(76,175,80,0.15); color: #81c784; border: 1px solid rgba(76,175,80,0.3); }
        .alert-error { background: rgba(244,67,54,0.15); color: #e57373; border: 1px solid rgba(244,67,54,0.3); }
        
        .modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.7); z-index: 1000; }
        .modal-content { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); width: 90%; max-width: 550px; margin: 80px auto; padding: 30px; border-radius: 15px; border: 1px solid rgba(255,255,255,0.06); }
        .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.06); }
        .modal-header h3 { margin: 0; font-size: 18px; }
        .modal-close { background: none; border: none; color: #888; font-size: 24px; cursor: pointer; }
        .modal-footer { display: flex; justify-content: flex-end; gap: 10px; margin-top: 25px; }
        
        .form-group { margin-bottom: 18px; }
        .form-group label { display: block; margin-bottom: 6px; font-weight: 600; color: #e0e0e0; font-size: 13px; }
        .form-group input, .form-group select { width: 100%; padding: 10px 12px; background: #2d2d44; border: 1px solid rgba(255,255,255,0.12); border-radius: 6px; color: #e0e0e0; font-size: 14px; }
        .form-group input:focus, .form-group select:focus { outline: none; border-color: #4CAF50; }
        
        .text-center { text-align: center; }
        .text-muted { color: #888; font-size: 13px; }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-clipboard-list"></i> 生产工单管理</h2>
            <button class="btn btn-primary" onclick="openCreateModal()"><i class="fas fa-plus"></i> 新建工单</button>
        </div>
        
        <% If msg <> "" Then %><div class="alert alert-<%=msgType%>"><%=Server.HTMLEncode(msg)%></div><% End If %>
        
        <div class="stats-grid">
            <div class="stat-card"><span class="num" style="color:#4CAF50;"><%=pmTotal%></span><span class="label">总工单</span></div>
            <div class="stat-card"><span class="num" style="color:#FF9800;"><%=pmPending%></span><span class="label">待生产</span></div>
            <div class="stat-card"><span class="num" style="color:#2196F3;"><%=pmProgress%></span><span class="label">生产中</span></div>
            <div class="stat-card"><span class="num" style="color:#888;"><%=pmCompleted%></span><span class="label">已完成</span></div>
        </div>
        
        <div class="filter-tabs">
            <a href="production_management.asp" class="filter-tab <%=IIF(filterStatus="","active","")%>">全部</a>
            <a href="?status=Pending" class="filter-tab <%=IIF(filterStatus="Pending","active","")%>">待生产</a>
            <a href="?status=InProgress" class="filter-tab <%=IIF(filterStatus="InProgress","active","")%>">生产中</a>
            <a href="?status=Completed" class="filter-tab <%=IIF(filterStatus="Completed","active","")%>">已完成</a>
        </div>
        
        <div class="card">
            <div class="card-header">生产工单列表</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>工单号</th><th>订单号</th><th>配方名</th><th>计划量</th><th>优先级</th><th>状态</th><th>负责人</th><th>创建时间</th><th>操作</th></tr></thead>
                    <tbody>
                    <%
                    Dim pmRowCount : pmRowCount = 0
                    If Not rsPM Is Nothing Then
                        Do While Not rsPM.EOF
                            pmRowCount = pmRowCount + 1
                            Dim pmStatus : pmStatus = CStr(rsPM("Status") & "")
                    %>
                        <tr>
                            <td><strong><%=rsPM("WorkOrderNo") & ""%></strong></td>
                            <td><%=rsPM("OrderNo") & ""%></td>
                            <td><%=rsPM("RecipeName") & ""%></td>
                            <td><%=rsPM("PlannedQty") & ""%></td>
                            <td><%=IIF(IsNull(rsPM("Priority")),"",rsPM("Priority"))%></td>
                            <td><span class="status-badge status-<%=LCase(pmStatus)%>"><%=pmStatus%></span></td>
                            <td><%=rsPM("AssignedTo") & ""%></td>
                            <td class="text-muted"><%=IIF(IsNull(rsPM("CreatedAt")),"",Left(rsPM("CreatedAt"),10))%></td>
                            <td>
                                <% If pmStatus = "Pending" Then %>
                                <form method="post" style="display:inline;">
                                    <input type="hidden" name="action" value="update_status">
                                    <input type="hidden" name="production_id" value="<%=rsPM("ProductionID")%>">
                                    <input type="hidden" name="new_status" value="InProgress">
                                    <button type="submit" class="btn btn-primary btn-sm">开始生产</button>
                                </form>
                                <% ElseIf pmStatus = "InProgress" Then %>
                                <form method="post" style="display:inline;">
                                    <input type="hidden" name="action" value="update_status">
                                    <input type="hidden" name="production_id" value="<%=rsPM("ProductionID")%>">
                                    <input type="hidden" name="new_status" value="Completed">
                                    <button type="submit" class="btn btn-warning btn-sm">完成生产</button>
                                </form>
                                <% End If %>
                            </td>
                        </tr>
                    <%
                            rsPM.MoveNext
                        Loop
                        rsPM.Close
                    End If
                    Set rsPM = Nothing
                    If pmRowCount = 0 Then %>
                        <tr><td colspan="9" class="text-center text-muted" style="padding:40px;">暂无生产工单</td></tr>
                    <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    
    <!-- 新建工单弹窗 -->
    <div id="createModal" class="modal">
        <div class="modal-content">
            <div class="modal-header"><h3>新建生产工单</h3><button class="modal-close" onclick="closeModal('createModal')">&times;</button></div>
            <form method="post">
                <input type="hidden" name="action" value="create_po">
                <input type="hidden" name="recipe_name" id="recipeName" value="">
                <div class="form-group"><label>关联订单</label><select name="order_id" id="orderSelect" onchange="onOrderChange()"><option value="0">-- 不关联订单（独立工单）--</option><%
If Not rsOrderOptions Is Nothing Then
    Do While Not rsOrderOptions.EOF
%><option value="<%=rsOrderOptions("OrderID")%>" <%=IIf(pmPreOrderID=rsOrderOptions("OrderID"),"selected","")%>><%=rsOrderOptions("OrderNo") & ""%> - <%=Left(rsOrderOptions("ProductName") & "", 25)%></option><%
        rsOrderOptions.MoveNext
    Loop
    rsOrderOptions.Close
End If
Set rsOrderOptions = Nothing
%></select></div>
                <div class="form-group"><label>配方ID</label><input type="number" name="recipe_id" id="recipeId" value="0"></div>
                <div class="form-group"><label>产品名称</label><input type="text" id="productNamePreview" readonly style="background:#1a1a2e;color:#888;"></div>
                <div class="form-group"><label>计划产量</label><input type="number" name="planned_qty" id="plannedQty" required min="1" value="1"></div>
                <div class="form-group"><label>负责人</label><select name="assigned_to"><option value="">-- 未指定 --</option><%
If Not rsAssigneesPM Is Nothing Then
    Do While Not rsAssigneesPM.EOF
%><option value="<%=rsAssigneesPM("FullName") & ""%>"><%=rsAssigneesPM("FullName") & ""%></option><%
        rsAssigneesPM.MoveNext
    Loop
    rsAssigneesPM.Close
End If
Set rsAssigneesPM = Nothing
%></select></div>
                <div class="form-group"><label>优先级</label><select name="priority"><option value="1">普通</option><option value="2" selected>优先</option><option value="3">紧急</option></select></div>
                <div class="form-group"><label>备注</label><input type="text" name="notes"></div>
                <div class="modal-footer">
                    <button type="button" class="btn btn--neutral" onclick="closeModal('createModal')">取消</button>
                    <button type="submit" class="btn btn-primary">创建工单</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
    var orderData = {<%=orderDataJSON%>};
    
    function openCreateModal() {
        document.getElementById('createModal').style.display = 'block';
        // 如果URL有预选订单，自动触发
        var sel = document.getElementById('orderSelect');
        if (sel.value != '0') onOrderChange();
    }
    
    function closeModal(id) { document.getElementById(id).style.display = 'none'; }
    window.onclick = function(event) { if (event.target.classList.contains('modal')) event.target.style.display = 'none'; }
    
    function onOrderChange() {
        var orderId = document.getElementById('orderSelect').value;
        var info = orderData[orderId];
        if (info) {
            document.getElementById('recipeId').value = info.recipe || 0;
            document.getElementById('recipeName').value = info.product || '';
            document.getElementById('productNamePreview').value = info.product || '';
            document.getElementById('plannedQty').value = info.qty || 1;
        } else {
            document.getElementById('recipeId').value = 0;
            document.getElementById('recipeName').value = '';
            document.getElementById('productNamePreview').value = '';
            document.getElementById('plannedQty').value = 1;
        }
    }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
