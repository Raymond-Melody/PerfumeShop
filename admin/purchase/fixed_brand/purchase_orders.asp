<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<%
Call OpenConnection()
%>
<!--#include file="includes/db_setup.asp"-->
<%
' ========== 辅助函数 ==========
Function GetStatusName(statusCode)
    Select Case statusCode
        Case "Draft"          : GetStatusName = "草稿"
        Case "Submitted"      : GetStatusName = "待审批"
        Case "Approved"       : GetStatusName = "已审批"
        Case "Ordered"        : GetStatusName = "已下单"
        Case "PartialReceived": GetStatusName = "部分收货"
        Case "Received"       : GetStatusName = "已收货"
        Case "Completed"      : GetStatusName = "已完成"
        Case "Rejected"       : GetStatusName = "已拒绝"
        Case "Cancelled"      : GetStatusName = "已取消"
        Case Else             : GetStatusName = statusCode
    End Select
End Function

Function GetStatusClass(statusCode)
    Select Case statusCode
        Case "Draft"          : GetStatusClass = "status-draft"
        Case "Submitted"      : GetStatusClass = "status-submitted"
        Case "Approved"       : GetStatusClass = "status-approved"
        Case "Ordered"        : GetStatusClass = "status-ordered"
        Case "PartialReceived": GetStatusClass = "status-partial"
        Case "Received"       : GetStatusClass = "status-received"
        Case "Completed"      : GetStatusClass = "status-completed"
        Case "Rejected"       : GetStatusClass = "status-rejected"
        Case "Cancelled"      : GetStatusClass = "status-cancelled"
        Case Else             : GetStatusClass = "status-draft"
    End Select
End Function

Function GeneratePurchaseNo()
    Dim today, prefix, countNum
    today = Date()
    prefix = "FBPO-" & Year(today) & Right("0" & Month(today), 2) & Right("0" & Day(today), 2) & "-"
    countNum = SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandPurchaseOrders WHERE PurchaseNo LIKE '" & prefix & "%'"))
    GeneratePurchaseNo = prefix & Right("000" & (countNum + 1), 3)
End Function

' ========== 消息变量 ==========
Dim msg, msgType
msg = ""
msgType = "success"

' ========== POST处理 ==========
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not ValidateCSRFToken() Then
        msg = "安全令牌验证失败"
        msgType = "error"
    Else
        Dim postAction : postAction = Trim(Request.Form("action"))
        
        If postAction = "create" Then
            Dim newSupplierID : newSupplierID = SafeNum(Request.Form("supplier_id"))
            Dim newSupplierName : newSupplierName = Trim(Request.Form("supplier_name"))
            ' 从数据库验证供应商名称，防止前端伪造
            If newSupplierID > 0 Then
                Dim dbSupplierName : dbSupplierName = CStr(GetScalar("SELECT SupplierName FROM Suppliers WHERE SupplierID=" & newSupplierID) & "")
                If dbSupplierName <> "" Then newSupplierName = dbSupplierName
            End If
            Dim newExpectedDt : newExpectedDt = Trim(Request.Form("expected_date"))
            Dim newRemarks : newRemarks = SafeSQL(Trim(Request.Form("remarks")))
            Dim detailCount : detailCount = SafeNum(Request.Form("detail_count"))
            
            If newSupplierID <= 0 Then
                msg = "请选择供应商"
                msgType = "error"
            ElseIf detailCount <= 0 Then
                msg = "请至少添加一个产品"
                msgType = "error"
            Else
                Dim newPONo : newPONo = GeneratePurchaseNo()
                Dim newTotalAmount : newTotalAmount = 0
                
                ' 先计算总金额
                Dim di
                For di = 1 To detailCount
                    Dim dQty : dQty = SafeNum(Request.Form("detail_qty_" & di))
                    Dim dPrice : dPrice = SafeNum(Request.Form("detail_price_" & di))
                    newTotalAmount = newTotalAmount + (dQty * dPrice)
                Next
                
                Call BeginTransaction()
                
                Dim insOrderSQL : insOrderSQL = "INSERT INTO FixedBrandPurchaseOrders (PurchaseNo, SupplierID, SupplierName, TotalAmount, Status, ExpectedDate, Remarks, CreatedBy) VALUES ('" & _
                    newPONo & "', " & newSupplierID & ", '" & SafeSQL(newSupplierName) & "', " & newTotalAmount & ", 'Draft', " & _
                    IIf(newExpectedDt <> "", "'" & newExpectedDt & "'", "NULL") & ", '" & newRemarks & "', '" & SafeSQL(Session("AdminName")) & "')"
                
                If ExecuteNonQuery(insOrderSQL) Then
                    Dim newPID : newPID = SafeNum(GetScalar("SELECT MAX(PurchaseID) FROM FixedBrandPurchaseOrders"))
                    
                    If newPID > 0 Then
                        Dim allDetailOK : allDetailOK = True
                        For di = 1 To detailCount
                            Dim dFPID : dFPID = SafeNum(Request.Form("detail_fpid_" & di))
                            Dim dPName : dPName = SafeSQL(Trim(Request.Form("detail_name_" & di)))
                            Dim dPSpec : dPSpec = SafeSQL(Trim(Request.Form("detail_spec_" & di)))
                            dQty = SafeNum(Request.Form("detail_qty_" & di))
                            dPrice = SafeNum(Request.Form("detail_price_" & di))
                            Dim dSubTotal : dSubTotal = dQty * dPrice
                            Dim dExpDt : dExpDt = Trim(Request.Form("detail_expected_" & di))
                            
                            Dim insDetailSQL : insDetailSQL = "INSERT INTO FixedBrandPurchaseDetails (PurchaseID, FixedProductID, ProductName, Specification, Quantity, UnitPrice, SubTotal, ExpectedDate) VALUES (" & _
                                newPID & ", " & dFPID & ", '" & dPName & "', '" & dPSpec & "', " & dQty & ", " & dPrice & ", " & dSubTotal & ", " & _
                                IIf(dExpDt <> "", "'" & dExpDt & "'", "NULL") & ")"
                            
                            If Not ExecuteNonQuery(insDetailSQL) Then
                                allDetailOK = False
                                Exit For
                            End If
                        Next
                        
                        If allDetailOK Then
                            Call CommitTransaction()
                            msg = "采购订单 " & newPONo & " 创建成功"
                            msgType = "success"
                        Else
                            Call RollbackTransaction()
                            msg = "明细添加失败，订单已回滚"
                            msgType = "error"
                        End If
                    Else
                        Call RollbackTransaction()
                        msg = "订单创建失败"
                        msgType = "error"
                    End If
                Else
                    Call RollbackTransaction()
                    msg = "订单创建失败：" & Session("LastDBError")
                    msgType = "error"
                End If
            End If
            
        ElseIf postAction = "update_status" Then
            Dim updPID : updPID = SafeNum(Request.Form("purchase_id"))
            Dim newStatus : newStatus = Trim(Request.Form("new_status"))
            
            If updPID > 0 And newStatus <> "" Then
                Dim updSQL : updSQL = "UPDATE FixedBrandPurchaseOrders SET Status='" & newStatus & "', UpdatedAt=GETDATE()"
                
                If newStatus = "Approved" Then
                    updSQL = updSQL & ", ApprovedBy='" & SafeSQL(Session("AdminName")) & "', ApprovedAt=GETDATE()"
                End If
                
                updSQL = updSQL & " WHERE PurchaseID=" & updPID
                
                If ExecuteNonQuery(updSQL) Then
                    If newStatus = "Completed" Then
                        ' 完成时更新成本分摊的利润数据
                        On Error Resume Next
                        Call ExecuteNonQuery("UPDATE FixedBrandCostAllocation SET ProfitAmount = ISNULL(SalePrice,0) * ISNULL(Quantity,0) - ISNULL(TotalCost,0), ProfitRate = CASE WHEN ISNULL(SalePrice,0) * ISNULL(Quantity,0) > 0 THEN ((ISNULL(SalePrice,0) * ISNULL(Quantity,0) - ISNULL(TotalCost,0)) / (ISNULL(SalePrice,0) * ISNULL(Quantity,0))) * 100 ELSE 0 END FROM FixedBrandCostAllocation WHERE PurchaseID=" & updPID)
                        If Err.Number <> 0 Then
                            Err.Clear
                            ' FROM 语法不兼容时降级为子查询方式
                            Call ExecuteNonQuery("UPDATE FixedBrandCostAllocation SET ProfitAmount = ISNULL(SalePrice,0) * ISNULL(Quantity,0) - ISNULL(TotalCost,0), ProfitRate = CASE WHEN ISNULL(SalePrice,0) * ISNULL(Quantity,0) > 0 THEN ((ISNULL(SalePrice,0) * ISNULL(Quantity,0) - ISNULL(TotalCost,0)) / (ISNULL(SalePrice,0) * ISNULL(Quantity,0))) * 100 ELSE 0 END WHERE PurchaseID=" & updPID)
                        End If
                        On Error GoTo 0
                    End If
                    msg = "状态更新成功"
                    msgType = "success"
                Else
                    msg = "状态更新失败"
                    msgType = "error"
                End If
            End If
            
        ElseIf postAction = "allocate_cost" Then
            ' 将采购成本按产品分摊到FixedBrandCostAllocation
            Dim allocPID : allocPID = SafeNum(Request.Form("purchase_id"))
            If allocPID > 0 Then
                Dim allocCount : allocCount = 0
                Dim rsDetail : Set rsDetail = conn.Execute("SELECT d.*, p.PurchaseNo FROM FixedBrandPurchaseDetails d JOIN FixedBrandPurchaseOrders p ON d.PurchaseID=p.PurchaseID WHERE d.PurchaseID=" & allocPID)
                If Not rsDetail Is Nothing Then
                    Do While Not rsDetail.EOF
                        Dim fpID : fpID = SafeNum(rsDetail("FixedProductID"))
                        Dim salePrice : salePrice = SafeNum(GetScalar("SELECT SalePrice FROM FixedBrandProducts WHERE FixedProductID=" & fpID))
                        Dim costPerUnit : costPerUnit = SafeNum(rsDetail("UnitPrice"))
                        Dim rcvQty : rcvQty = SafeNum(rsDetail("ReceivedQty"))
                        
                        If rcvQty > 0 Then
                            ' 查找近期的客户订单关联（优先匹配产品ID，查找最近30天的订单）
                            Dim relatedOrderID : relatedOrderID = SafeNum(GetScalar("SELECT TOP 1 od.OrderID FROM OrderDetails od JOIN Orders o ON od.OrderID=o.OrderID WHERE od.ProductID=(SELECT ISNULL(ProductID,0) FROM FixedBrandProducts WHERE FixedProductID=" & fpID & ") AND o.OrderDate >= DATEADD(DAY,-30,GETDATE()) ORDER BY o.OrderDate DESC"))
                            ' 将采购成本按产品分摊到 FixedBrandCostAllocation
                            Call ExecuteNonQuery("INSERT INTO FixedBrandCostAllocation (OrderID, OrderNo, PurchaseID, PurchaseNo, FixedProductID, ProductName, CostPerUnit, Quantity, TotalCost, SalePrice) VALUES (" & IIf(relatedOrderID>0, relatedOrderID, "0") & ", '" & IIf(relatedOrderID>0, GetScalar("SELECT OrderNo FROM Orders WHERE OrderID=" & relatedOrderID), "PENDING") & "', " & allocPID & ", '" & SafeSQL(CStr(rsDetail("PurchaseNo"))) & "', " & fpID & ", '" & SafeSQL(CStr(rsDetail("ProductName"))) & "', " & costPerUnit & ", " & rcvQty & ", " & (costPerUnit * rcvQty) & ", " & salePrice & ")")
                            allocCount = allocCount + 1
                        End If
                        rsDetail.MoveNext
                    Loop
                    rsDetail.Close
                End If
                Set rsDetail = Nothing
                msg = "已分摊 " & allocCount & " 项成本记录"
                msgType = "success"
            End If
        End If
    End If
End If

' ========== 查看订单详情 ==========
Dim viewID : viewID = SafeNum(Request.QueryString("view"))
Dim viewOrderData, viewDetailRS
If viewID > 0 Then
    Set viewOrderRS = conn.Execute("SELECT * FROM FixedBrandPurchaseOrders WHERE PurchaseID=" & viewID)
    If Not viewOrderRS Is Nothing Then
        If Not viewOrderRS.EOF Then
            Set viewOrderData = viewOrderRS
        Else
            viewOrderRS.Close
            Set viewOrderRS = Nothing
        End If
    End If
    If IsObject(viewOrderData) Then
        If Not viewOrderData Is Nothing Then
            Set viewDetailRS = conn.Execute("SELECT * FROM FixedBrandPurchaseDetails WHERE PurchaseID=" & viewID & " ORDER BY DetailID")
        End If
    End If
End If

' ========== 查询订单列表 ==========
Dim statusFilter : statusFilter = Trim(Request.QueryString("status"))
Dim dateFrom : dateFrom = Trim(Request.QueryString("date_from"))
Dim dateTo : dateTo = Trim(Request.QueryString("date_to"))

Dim whereSQL : whereSQL = " WHERE 1=1"
If statusFilter <> "" Then
    whereSQL = whereSQL & " AND Status='" & SafeSQL(statusFilter) & "'"
End If
If dateFrom <> "" Then
    whereSQL = whereSQL & " AND OrderDate >= '" & dateFrom & "'"
End If
If dateTo <> "" Then
    whereSQL = whereSQL & " AND OrderDate < DATEADD(DAY,1,'" & dateTo & "')"
End If

Dim sqlOrders : sqlOrders = "SELECT * FROM FixedBrandPurchaseOrders" & whereSQL & " ORDER BY PurchaseID DESC"
Dim rsOrders : Set rsOrders = conn.Execute(sqlOrders)

' ========== 统计 ==========
Dim totalOrders : totalOrders = SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandPurchaseOrders"))
Dim pendingApproval : pendingApproval = SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandPurchaseOrders WHERE Status='Submitted'"))
Dim pendingReceipt : pendingReceipt = SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandPurchaseOrders WHERE Status IN ('Ordered','PartialReceived')"))
Dim monthlyAmount : monthlyAmount = SafeNum(GetScalar("SELECT ISNULL(SUM(TotalAmount),0) FROM FixedBrandPurchaseOrders WHERE (Status='Ordered' OR Status='Received' OR Status='Completed' OR Status='PartialReceived') AND Month(OrderDate)=Month(GETDATE()) AND Year(OrderDate)=Year(GETDATE())"))

' ========== 供应商列表 ==========
Dim rsSuppliers : Set rsSuppliers = conn.Execute("SELECT SupplierID, SupplierName FROM Suppliers ORDER BY SupplierName")

' ========== 产品列表(供表单选择) ==========
Dim rsFBProducts : Set rsFBProducts = conn.Execute("SELECT FixedProductID, ProductCode, ProductName, Specification, UnitPrice, SupplierName FROM FixedBrandProducts WHERE Status='Active' ORDER BY ProductName")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>品牌定香采购订单 - 采购管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; }
        .main-content { margin-left: 270px; padding: 25px; min-height: 100vh; }
        
        .page-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }
        .page-title { font-size: 20px; font-weight: 600; color: #fff; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #FF9800; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 10px; padding: 18px; border: 1px solid rgba(255,255,255,0.05); }
        .stat-icon { width: 40px; height: 40px; border-radius: 8px; display: flex; align-items: center; justify-content: center; font-size: 16px; margin-bottom: 10px; }
        .stat-value { font-size: 22px; font-weight: 700; color: #fff; }
        .stat-label { font-size: 12px; color: #888; margin-top: 4px; }
        
        .toolbar { display: flex; gap: 10px; align-items: center; margin-bottom: 20px; flex-wrap: wrap; }
        .filter-select { padding: 8px 14px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.1); background: #2d2d44; color: #e0e0e0; font-size: 13px; }
        .filter-input { padding: 8px 14px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.1); background: #2d2d44; color: #e0e0e0; width: 140px; font-size: 13px; }
        
        .data-table { width: 100%; border-collapse: collapse; background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 10px; overflow: hidden; }
        .data-table th, .data-table td { padding: 12px 15px; text-align: left; border-bottom: 1px solid rgba(255,255,255,0.05); font-size: 13px; }
        .data-table th { color: #888; font-size: 11px; text-transform: uppercase; font-weight: 600; background: rgba(0,0,0,0.2); }
        .data-table td { color: #ccc; }
        .data-table tr:hover td { background: rgba(255,255,255,0.03); }
        
        .status-badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 500; }
        .status-draft { background: rgba(158,158,158,0.2); color: #9E9E9E; }
        .status-submitted { background: rgba(255,152,0,0.2); color: #FF9800; }
        .status-approved { background: rgba(33,150,243,0.2); color: #2196F3; }
        .status-ordered { background: rgba(156,39,176,0.2); color: #9C27B0; }
        .status-partial { background: rgba(0,188,212,0.2); color: #00BCD4; }
        .status-received { background: rgba(76,175,80,0.2); color: #4CAF50; }
        .status-completed { background: rgba(76,175,80,0.2); color: #4CAF50; }
        .status-rejected { background: rgba(244,67,54,0.2); color: #F44336; }
        .status-cancelled { background: rgba(158,158,158,0.2); color: #9E9E9E; }
        
        .action-btns { display: flex; gap: 4px; flex-wrap: wrap; }
        
        .modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.6); z-index: 1000; justify-content: center; align-items: center; }
        .modal-overlay.show { display: flex; }
        .modal-box { background: #2d2d44; border-radius: 12px; padding: 25px; width: 700px; max-height: 85vh; overflow-y: auto; border: 1px solid rgba(255,255,255,0.1); box-shadow: 0 10px 40px rgba(0,0,0,0.5); }
        .modal-box h3 { color: #fff; font-size: 18px; margin: 0 0 20px; }
        .form-group { margin-bottom: 14px; }
        .form-group label { display: block; font-size: 12px; color: #888; margin-bottom: 5px; }
        .form-group input, .form-group select, .form-group textarea { width: 100%; padding: 9px 12px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.1); background: #1a1a2e; color: #e0e0e0; font-size: 13px; box-sizing: border-box; }
        .form-row { display: flex; gap: 12px; }
        .form-row .form-group { flex: 1; }
        .modal-actions { display: flex; gap: 10px; justify-content: flex-end; margin-top: 20px; }
        
        .detail-row { display: flex; gap: 8px; align-items: center; margin-bottom: 8px; padding: 10px; background: rgba(255,255,255,0.03); border-radius: 8px; }
        .detail-row select { flex: 2; padding: 8px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.1); background: #1a1a2e; color: #e0e0e0; font-size: 12px; }
        .detail-row input { flex: 1; padding: 8px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.1); background: #1a1a2e; color: #e0e0e0; font-size: 12px; }
        .detail-row .detail-subtotal { flex: 0.8; text-align: right; font-weight: 600; color: #FF9800; font-size: 13px; }
        
        .detail-view { display: none; }
        .detail-view.show { display: block; }
    </style>
</head>
<body data-theme="purchase-dark">
    <!--#include file="../includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-file-invoice"></i> 品牌定香采购订单</h2>
            <div class="breadcrumb" style="font-size:13px;color:#888;">
                <a href="index.asp" style="color:#FF9800;text-decoration:none;">品牌定香采购</a> / 采购订单
            </div>
        </div>
        
        <% If msg <> "" Then %>
        <div style="padding:12px 20px; border-radius:8px; margin-bottom:20px; font-size:14px; background:<%=IIf(msgType="success","rgba(76,175,80,0.15)","rgba(244,67,54,0.15)")%>; color:<%=IIf(msgType="success","#4CAF50","#F44336")%>; border:1px solid <%=IIf(msgType="success","rgba(76,175,80,0.3)","rgba(244,67,54,0.3)")%>;">
            <i class="fas fa-<%=IIf(msgType="success","check-circle","exclamation-circle")%>"></i> <%= msg %>
        </div>
        <% End If %>
        
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-icon" style="background:linear-gradient(135deg,#FF9800,#F57C00);"><i class="fas fa-file-invoice"></i></div>
                <div class="stat-value"><%= totalOrders %></div>
                <div class="stat-label">订单总数</div>
            </div>
            <div class="stat-card">
                <div class="stat-icon" style="background:linear-gradient(135deg,#FF5722,#D84315);"><i class="fas fa-clock"></i></div>
                <div class="stat-value"><%= pendingApproval %></div>
                <div class="stat-label">待审批</div>
            </div>
            <div class="stat-card">
                <div class="stat-icon" style="background:linear-gradient(135deg,#2196F3,#1565C0);"><i class="fas fa-truck-loading"></i></div>
                <div class="stat-value"><%= pendingReceipt %></div>
                <div class="stat-label">待收货</div>
            </div>
            <div class="stat-card">
                <div class="stat-icon" style="background:linear-gradient(135deg,#4CAF50,#388E3C);"><i class="fas fa-yen-sign"></i></div>
                <div class="stat-value">¥<%= FormatNumber(monthlyAmount, 0) %></div>
                <div class="stat-label">本月采购额</div>
            </div>
        </div>
        
        <div class="toolbar">
            <button class="btn btn--primary" onclick="openCreateModal()"><i class="fas fa-plus"></i> 创建订单</button>
            <form method="get" style="display:flex;gap:8px;margin-left:auto;">
                <select name="status" class="filter-select" onchange="this.form.submit()">
                    <option value="">全部状态</option>
                    <option value="Draft" <% If statusFilter="Draft" Then %>selected<% End If %>>草稿</option>
                    <option value="Submitted" <% If statusFilter="Submitted" Then %>selected<% End If %>>待审批</option>
                    <option value="Approved" <% If statusFilter="Approved" Then %>selected<% End If %>>已审批</option>
                    <option value="Ordered" <% If statusFilter="Ordered" Then %>selected<% End If %>>已下单</option>
                    <option value="PartialReceived" <% If statusFilter="PartialReceived" Then %>selected<% End If %>>部分收货</option>
                    <option value="Received" <% If statusFilter="Received" Then %>selected<% End If %>>已收货</option>
                    <option value="Completed" <% If statusFilter="Completed" Then %>selected<% End If %>>已完成</option>
                </select>
                <input type="date" name="date_from" class="filter-input" value="<%=Server.HTMLEncode(dateFrom)%>" onchange="this.form.submit()">
                <input type="date" name="date_to" class="filter-input" value="<%=Server.HTMLEncode(dateTo)%>" onchange="this.form.submit()">
                <% If statusFilter <> "" Or dateFrom <> "" Or dateTo <> "" Then %>
                <a href="purchase_orders.asp" class="btn btn--neutral btn--sm"><i class="fas fa-times"></i></a>
                <% End If %>
            </form>
        </div>
        
        <table class="data-table">
            <thead>
                <tr>
                    <th>采购单号</th>
                    <th>供应商</th>
                    <th>金额</th>
                    <th>状态</th>
                    <th>下单日期</th>
                    <th>预计到货</th>
                    <th>操作</th>
                </tr>
            </thead>
            <tbody>
                <% If Not rsOrders Is Nothing Then
                    If Not rsOrders.EOF Then
                        Do While Not rsOrders.EOF
                            Dim orderStatus : orderStatus = CStr(rsOrders("Status"))
                %>
                <tr>
                    <td><span style="font-family:Consolas,monospace;color:#FF9800;"><%= Server.HTMLEncode(CStr(rsOrders("PurchaseNo"))) %></span></td>
                    <td><%= Server.HTMLEncode(CStr(rsOrders("SupplierName") & "")) %></td>
                    <td>¥<%= FormatNumber(SafeNum(rsOrders("TotalAmount")), 2) %></td>
                    <td><span class="status-badge <%= GetStatusClass(orderStatus) %>"><%= GetStatusName(orderStatus) %></span></td>
                    <td><%= SafeFormatDateTime(rsOrders("OrderDate"), 2) %></td>
                    <td><%= SafeFormatDateTime(rsOrders("ExpectedDate"), 2) %></td>
                    <td>
                        <div class="action-btns">
                            <button class="btn btn--primary btn--xs" onclick="viewOrder(<%= rsOrders("PurchaseID") %>)" title="查看详情"><i class="fas fa-eye"></i></button>
                            <% If orderStatus = "Draft" Then %>
                            <button class="btn btn--warning btn--xs" onclick="updateStatus(<%= rsOrders("PurchaseID") %>, 'Submitted')" title="提交审批"><i class="fas fa-paper-plane"></i></button>
                            <% End If %>
                            <% If orderStatus = "Submitted" Then %>
                            <button class="btn btn--success btn--xs" onclick="updateStatus(<%= rsOrders("PurchaseID") %>, 'Approved')" title="审批通过"><i class="fas fa-check"></i></button>
                            <button class="btn btn--danger btn--xs" onclick="updateStatus(<%= rsOrders("PurchaseID") %>, 'Rejected')" title="拒绝"><i class="fas fa-times"></i></button>
                            <% End If %>
                            <% If orderStatus = "Approved" Then %>
                            <button class="btn btn--info btn--xs" onclick="updateStatus(<%= rsOrders("PurchaseID") %>, 'Ordered')" title="确认下单"><i class="fas fa-shopping-cart"></i></button>
                            <% End If %>
                            <% If orderStatus = "Ordered" Or orderStatus = "PartialReceived" Then %>
                            <a href="receiving.asp?purchase_id=<%= rsOrders("PurchaseID") %>" class="btn btn--success btn--xs" title="收货入库"><i class="fas fa-box"></i></a>
                            <% End If %>
                            <% If orderStatus = "Received" Then %>
                            <button class="btn btn--success btn--xs" onclick="updateStatus(<%= rsOrders("PurchaseID") %>, 'Completed')" title="完成"><i class="fas fa-flag-checkered"></i></button>
                            <button class="btn btn--info btn--xs" onclick="allocateCost(<%= rsOrders("PurchaseID") %>)" title="分摊成本"><i class="fas fa-calculator"></i></button>
                            <% End If %>
                            <% If orderStatus = "Draft" Or orderStatus = "Submitted" Then %>
                            <button class="btn btn--neutral btn--xs" onclick="updateStatus(<%= rsOrders("PurchaseID") %>, 'Cancelled')" title="取消"><i class="fas fa-ban"></i></button>
                            <% End If %>
                        </div>
                    </td>
                </tr>
                <%
                            rsOrders.MoveNext
                        Loop
                    Else
                %>
                <tr><td colspan="7" style="text-align:center;padding:40px;color:#666;">
                    <i class="fas fa-inbox" style="font-size:28px;display:block;margin-bottom:10px;"></i>暂无采购订单
                </td></tr>
                <%      End If
                    rsOrders.Close
                    Set rsOrders = Nothing
                Else
                %>
                <tr><td colspan="7" style="text-align:center;padding:40px;color:#666;">数据加载失败</td></tr>
                <% End If %>
            </tbody>
        </table>
    </div>
    
    <!-- 创建订单弹窗 -->
    <div class="modal-overlay" id="createModal">
        <div class="modal-box">
            <h3><i class="fas fa-plus-circle" style="color:#FF9800;"></i> 创建采购订单</h3>
            <form method="post" id="createForm">
                <input type="hidden" name="csrf_token" value="<%= Session("CSRFToken") %>">
                <input type="hidden" name="action" value="create">
                <input type="hidden" name="detail_count" id="detailCount" value="1">
                
                <div class="form-row">
                    <div class="form-group">
                        <label>供应商</label>
                        <select name="supplier_id" id="supplierSelect" onchange="updateSupplier()" required>
                            <option value="0">-- 选择供应商 --</option>
                            <% If Not rsSuppliers Is Nothing Then
                                Do While Not rsSuppliers.EOF %>
                            <option value="<%= rsSuppliers("SupplierID") %>" data-name="<%= Server.HTMLEncode(CStr(rsSuppliers("SupplierName"))) %>"><%= Server.HTMLEncode(CStr(rsSuppliers("SupplierName"))) %></option>
                            <%      rsSuppliers.MoveNext
                                Loop
                                rsSuppliers.Close
                                Set rsSuppliers = Nothing
                            End If %>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>预计到货日期</label>
                        <input type="date" name="expected_date">
                    </div>
                </div>
                <div class="form-group">
                    <label>备注</label>
                    <textarea name="remarks" rows="2" placeholder="订单备注..."></textarea>
                </div>
                
                <div style="margin-top:16px;">
                    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
                        <span style="color:#888;font-size:13px;"><i class="fas fa-list"></i> 订单明细</span>
                        <button type="button" class="btn btn--success btn--xs" onclick="addDetailRow()"><i class="fas fa-plus"></i> 添加行</button>
                    </div>
                    <div id="detailRows">
                        <div class="detail-row" id="row-1">
                            <input type="hidden" name="detail_fpid_1" value="0">
                            <select name="detail_product_1" onchange="onProductSelect(this, 1)" style="flex:2;">
                                <option value="">-- 选择产品 --</option>
                                <% If Not rsFBProducts Is Nothing Then
                                    Do While Not rsFBProducts.EOF %>
                                <option value="<%= rsFBProducts("FixedProductID") %>" data-price="<%= SafeNum(rsFBProducts("UnitPrice")) %>" data-spec="<%= Server.HTMLEncode(CStr(rsFBProducts("Specification") & "")) %>" data-name="<%= Server.HTMLEncode(CStr(rsFBProducts("ProductName"))) %>"><%= Server.HTMLEncode(CStr(rsFBProducts("ProductCode"))) %> - <%= Server.HTMLEncode(CStr(rsFBProducts("ProductName"))) %></option>
                                <%      rsFBProducts.MoveNext
                                    Loop
                                    rsFBProducts.Close
                                    Set rsFBProducts = Nothing
                                End If %>
                            </select>
                            <input type="hidden" name="detail_name_1" value="">
                            <input type="hidden" name="detail_spec_1" value="">
                            <input type="number" name="detail_qty_1" placeholder="数量" min="1" value="1" onchange="calcSubtotal(1)" style="flex:0.6;">
                            <input type="number" name="detail_price_1" placeholder="单价" step="0.01" value="0" onchange="calcSubtotal(1)" style="flex:0.8;">
                            <span class="detail-subtotal" id="subtotal-1">¥0.00</span>
                            <input type="hidden" name="detail_expected_1" value="">
                        </div>
                    </div>
                    <div style="text-align:right;margin-top:10px;font-size:16px;color:#fff;">
                        合计：<span style="color:#FF9800;font-weight:700;" id="totalAmount">¥0.00</span>
                    </div>
                </div>
                
                <input type="hidden" name="supplier_name" id="hiddenSupplierName" value="">
                
                <div class="modal-actions">
                    <button type="button" class="btn btn--neutral" onclick="closeCreateModal()">取消</button>
                    <button type="submit" class="btn btn--primary">创建订单</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- 查看订单详情弹窗 -->
    <% If IsObject(viewOrderData) Then
        If Not viewOrderData Is Nothing Then %>
    <div class="modal-overlay show" id="viewModal">
        <div class="modal-box" style="width:800px;">
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;">
                <h3 style="margin:0;"><i class="fas fa-file-invoice" style="color:#FF9800;"></i> 采购订单详情</h3>
                <button class="btn btn--neutral btn--sm" onclick="closeViewModal()"><i class="fas fa-times"></i></button>
            </div>
            
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:20px;padding:15px;background:rgba(255,255,255,0.03);border-radius:8px;">
                <div><span style="color:#888;font-size:12px;">采购单号</span><br><span style="font-family:Consolas,monospace;color:#FF9800;font-size:15px;"><%= Server.HTMLEncode(CStr(viewOrderData("PurchaseNo"))) %></span></div>
                <div><span style="color:#888;font-size:12px;">供应商</span><br><span style="color:#fff;"><%= Server.HTMLEncode(CStr(viewOrderData("SupplierName") & "")) %></span></div>
                <div><span style="color:#888;font-size:12px;">状态</span><br><span class="status-badge <%= GetStatusClass(CStr(viewOrderData("Status"))) %>"><%= GetStatusName(CStr(viewOrderData("Status"))) %></span></div>
                <div><span style="color:#888;font-size:12px;">订单金额</span><br><span style="color:#4CAF50;font-size:18px;font-weight:700;">¥<%= FormatNumber(SafeNum(viewOrderData("TotalAmount")), 2) %></span></div>
                <div><span style="color:#888;font-size:12px;">下单日期</span><br><span style="color:#ccc;"><%= SafeFormatDateTime(viewOrderData("OrderDate"), 2) %></span></div>
                <div><span style="color:#888;font-size:12px;">预计到货</span><br><span style="color:#ccc;"><%= SafeFormatDateTime(viewOrderData("ExpectedDate"), 2) %></span></div>
                <div><span style="color:#888;font-size:12px;">创建人</span><br><span style="color:#ccc;"><%= Server.HTMLEncode(CStr(viewOrderData("CreatedBy") & "")) %></span></div>
                <div><span style="color:#888;font-size:12px;">备注</span><br><span style="color:#ccc;"><%= Server.HTMLEncode(CStr(viewOrderData("Remarks") & "")) %></span></div>
            </div>
            
            <h4 style="color:#fff;margin:0 0 10px;font-size:14px;"><i class="fas fa-list"></i> 订单明细</h4>
            <table class="data-table">
                <thead>
                    <tr>
                        <th>产品名称</th>
                        <th>规格</th>
                        <th>数量</th>
                        <th>单价</th>
                        <th>已收货</th>
                        <th>小计</th>
                    </tr>
                </thead>
                <tbody>
                    <% If Not viewDetailRS Is Nothing Then
                        Dim viewDetailTotal : viewDetailTotal = 0
                        Do While Not viewDetailRS.EOF
                            Dim vdQty : vdQty = SafeNum(viewDetailRS("Quantity"))
                            Dim vdPrice : vdPrice = SafeNum(viewDetailRS("UnitPrice"))
                            Dim vdSub : vdSub = vdQty * vdPrice
                            viewDetailTotal = viewDetailTotal + vdSub
                    %>
                    <tr>
                        <td><%= Server.HTMLEncode(CStr(viewDetailRS("ProductName"))) %></td>
                        <td><%= Server.HTMLEncode(CStr(viewDetailRS("Specification") & "")) %></td>
                        <td><%= vdQty %></td>
                        <td>¥<%= FormatNumber(vdPrice, 2) %></td>
                        <td style="color:<%=IIf(SafeNum(viewDetailRS("ReceivedQty"))>=vdQty,"#4CAF50","#FF9800")%>;"><%= SafeNum(viewDetailRS("ReceivedQty")) %>/<%= vdQty %></td>
                        <td style="font-weight:600;">¥<%= FormatNumber(vdSub, 2) %></td>
                    </tr>
                    <%
                            viewDetailRS.MoveNext
                        Loop
                        viewDetailRS.Close
                        Set viewDetailRS = Nothing
                    %>
                    <tr style="background:rgba(76,175,80,0.08);">
                        <td colspan="5" style="text-align:right;font-weight:600;color:#fff;">合计</td>
                        <td style="font-weight:700;color:#4CAF50;font-size:15px;">¥<%= FormatNumber(viewDetailTotal, 2) %></td>
                    </tr>
                    <% End If %>
                </tbody>
            </table>
            
            <div class="modal-actions" style="margin-top:20px;">
                <% Dim vStatus : vStatus = CStr(viewOrderData("Status")) %>
                <% If vStatus = "Draft" Then %>
                <button class="btn btn--warning" onclick="viewUpdateStatus(<%= viewID %>, 'Submitted')"><i class="fas fa-paper-plane"></i> 提交审批</button>
                <% End If %>
                <% If vStatus = "Submitted" Then %>
                <button class="btn btn--success" onclick="viewUpdateStatus(<%= viewID %>, 'Approved')"><i class="fas fa-check"></i> 审批通过</button>
                <button class="btn btn--danger" onclick="viewUpdateStatus(<%= viewID %>, 'Rejected')"><i class="fas fa-times"></i> 拒绝</button>
                <% End If %>
                <% If vStatus = "Approved" Then %>
                <button class="btn btn--info" onclick="viewUpdateStatus(<%= viewID %>, 'Ordered')"><i class="fas fa-shopping-cart"></i> 确认下单</button>
                <% End If %>
                <% If vStatus = "Ordered" Or vStatus = "PartialReceived" Then %>
                <a href="receiving.asp?purchase_id=<%= viewID %>" class="btn btn--success"><i class="fas fa-box"></i> 收货入库</a>
                <% End If %>
                <% If vStatus = "Received" Then %>
                <button class="btn btn--success" onclick="viewUpdateStatus(<%= viewID %>, 'Completed')"><i class="fas fa-flag-checkered"></i> 完成</button>
                <button class="btn btn--info" onclick="allocateCost(<%= viewID %>)"><i class="fas fa-calculator"></i> 分摊成本</button>
                <% End If %>
                <% If vStatus = "Draft" Or vStatus = "Submitted" Then %>
                <button class="btn btn--neutral" onclick="viewUpdateStatus(<%= viewID %>, 'Cancelled')"><i class="fas fa-ban"></i> 取消订单</button>
                <% End If %>
                <button type="button" class="btn btn--neutral" onclick="closeViewModal()">关闭</button>
            </div>
        </div>
    </div>
    <%
        viewOrderData.Close
        Set viewOrderData = Nothing
    End If
End If
    %>
    
    <script>
        var detailRowCounter = 1;  // 用于生成唯一行ID，只增不减
        
        // 获取当前实际行数
        function getActualRowCount() {
            return document.querySelectorAll('#detailRows .detail-row').length;
        }
        
        function openCreateModal() {
            document.getElementById('createModal').classList.add('show');
        }
        
        function closeCreateModal() {
            document.getElementById('createModal').classList.remove('show');
        }
        
        function updateSupplier() {
            var sel = document.getElementById('supplierSelect');
            var name = sel.options[sel.selectedIndex].getAttribute('data-name');
            document.getElementById('hiddenSupplierName').value = name || '';
        }
        
        function addDetailRow() {
            detailRowCounter++;
            var newIdx = detailRowCounter;  // 使用唯一递增ID
            document.getElementById('detailCount').value = getActualRowCount() + 1;
            
            var row = document.createElement('div');
            row.className = 'detail-row';
            row.id = 'row-' + newIdx;
            row.setAttribute('data-row-index', newIdx);
            
            var selectHTML = document.querySelector('#row-1 select').outerHTML;
            selectHTML = selectHTML.replace(/detail_product_1/g, 'detail_product_' + newIdx);
            selectHTML = selectHTML.replace(/onProductSelect\(this, 1\)/g, 'onProductSelect(this, ' + newIdx + ')');
            // 清除可能从第一行克隆的选中状态
            selectHTML = selectHTML.replace(/ selected(?=[ >])/gi, '');
            
            row.innerHTML = 
                '<input type="hidden" name="detail_fpid_' + newIdx + '" value="0">' +
                selectHTML +
                '<input type="hidden" name="detail_name_' + newIdx + '" value="">' +
                '<input type="hidden" name="detail_spec_' + newIdx + '" value="">' +
                '<input type="number" name="detail_qty_' + newIdx + '" placeholder="数量" min="1" value="1" onchange="calcSubtotal(' + newIdx + ')" style="flex:0.6;">' +
                '<input type="number" name="detail_price_' + newIdx + '" placeholder="单价" step="0.01" value="0" onchange="calcSubtotal(' + newIdx + ')" style="flex:0.8;">' +
                '<span class="detail-subtotal" id="subtotal-' + newIdx + '">¥0.00</span>' +
                '<input type="hidden" name="detail_expected_' + newIdx + '" value="">' +
                '<button type="button" class="btn btn--danger btn--xs" onclick="removeDetailRow(' + newIdx + ')" style="padding:4px 8px;"><i class="fas fa-times"></i></button>';
            
            document.getElementById('detailRows').appendChild(row);
        }
        
        function removeDetailRow(index) {
            var row = document.getElementById('row-' + index);
            if (!row) return;
            row.remove();
            
            // 重新编号所有剩余行，确保服务器端按连续序号处理
            var rows = document.querySelectorAll('#detailRows .detail-row');
            for (var i = 0; i < rows.length; i++) {
                var newSeq = i + 1;
                var oldIdx = rows[i].getAttribute('data-row-index');
                rows[i].id = 'row-' + newSeq;
                rows[i].setAttribute('data-row-index', newSeq);
                
                // 更新所有 input/select 的 name 属性
                var inputs = rows[i].querySelectorAll('input, select, span.detail-subtotal');
                for (var j = 0; j < inputs.length; j++) {
                    var name = inputs[j].getAttribute('name');
                    if (name && name.indexOf('detail_') === 0) {
                        var newName = name.replace(/detail_(\w+)_\d+$/, 'detail_$1_' + newSeq);
                        inputs[j].setAttribute('name', newName);
                    }
                    // 更新 onchange 等属性中的序号
                    var onchange = inputs[j].getAttribute('onchange');
                    if (onchange && onchange.indexOf('calcSubtotal') !== -1) {
                        onchange = onchange.replace(/calcSubtotal\(\d+\)/, 'calcSubtotal(' + newSeq + ')');
                        inputs[j].setAttribute('onchange', onchange);
                    }
                    // 更新 subtotal span 的 id
                    if (inputs[j].className && inputs[j].className.indexOf('detail-subtotal') !== -1) {
                        inputs[j].id = 'subtotal-' + newSeq;
                    }
                }
                // 更新 select 的 onchange
                var sel = rows[i].querySelector('select');
                if (sel) {
                    var selOnchange = sel.getAttribute('onchange');
                    if (selOnchange) {
                        selOnchange = selOnchange.replace(/onProductSelect\(this, \d+\)/, 'onProductSelect(this, ' + newSeq + ')');
                        sel.setAttribute('onchange', selOnchange);
                    }
                }
                // 更新删除按钮的 onclick
                var delBtn = rows[i].querySelector('button.btn--danger');
                if (delBtn) {
                    var delOnclick = delBtn.getAttribute('onclick');
                    delOnclick = delOnclick.replace(/removeDetailRow\(\d+\)/, 'removeDetailRow(' + newSeq + ')');
                    delBtn.setAttribute('onclick', delOnclick);
                }
            }
            
            // 更新实际行数
            document.getElementById('detailCount').value = rows.length;
            calcTotal();
        }
        
        function onProductSelect(sel, index) {
            var opt = sel.options[sel.selectedIndex];
            var fpid = opt.value;
            var price = opt.getAttribute('data-price') || '0';
            var name = opt.getAttribute('data-name') || '';
            var spec = opt.getAttribute('data-spec') || '';
            
            document.querySelector('input[name="detail_fpid_' + index + '"]').value = fpid;
            document.querySelector('input[name="detail_name_' + index + '"]').value = name;
            document.querySelector('input[name="detail_spec_' + index + '"]').value = spec;
            document.querySelector('input[name="detail_price_' + index + '"]').value = price;
            
            calcSubtotal(index);
        }
        
        function calcSubtotal(index) {
            var qty = parseFloat(document.querySelector('input[name="detail_qty_' + index + '"]').value) || 0;
            var price = parseFloat(document.querySelector('input[name="detail_price_' + index + '"]').value) || 0;
            var subtotal = qty * price;
            document.getElementById('subtotal-' + index).textContent = '¥' + subtotal.toFixed(2);
            calcTotal();
        }
        
        function calcTotal() {
            var total = 0;
            var rows = document.querySelectorAll('#detailRows .detail-row');
            for (var i = 0; i < rows.length; i++) {
                var seq = rows[i].getAttribute('data-row-index');
                var qty = parseFloat(rows[i].querySelector('input[name="detail_qty_' + seq + '"]').value) || 0;
                var price = parseFloat(rows[i].querySelector('input[name="detail_price_' + seq + '"]').value) || 0;
                total += qty * price;
            }
            document.getElementById('totalAmount').textContent = '¥' + total.toFixed(2);
        }
        
        function updateStatus(pid, newStatus) {
            var statusNames = {
                'Submitted': '提交审批', 'Approved': '审批通过', 'Rejected': '拒绝',
                'Ordered': '确认下单', 'Completed': '标记完成', 'Cancelled': '取消订单'
            };
            if (confirm('确定要' + (statusNames[newStatus] || newStatus) + '吗？')) {
                var f = document.createElement('form');
                f.method = 'post';
                f.innerHTML = '<input type="hidden" name="csrf_token" value="<%= Session("CSRFToken") %>"><input type="hidden" name="action" value="update_status"><input type="hidden" name="purchase_id" value="' + pid + '"><input type="hidden" name="new_status" value="' + newStatus + '">';
                document.body.appendChild(f);
                f.submit();
            }
        }
        
        function allocateCost(pid) {
            if (confirm('确定要分摊该订单的成本吗？')) {
                var f = document.createElement('form');
                f.method = 'post';
                f.innerHTML = '<input type="hidden" name="csrf_token" value="<%= Session("CSRFToken") %>"><input type="hidden" name="action" value="allocate_cost"><input type="hidden" name="purchase_id" value="' + pid + '">';
                document.body.appendChild(f);
                f.submit();
            }
        }
        
        function viewOrder(pid) {
            window.location.href = 'purchase_orders.asp?view=' + pid;
        }
        
        function closeViewModal() {
            window.location.href = 'purchase_orders.asp';
        }
        
        function viewUpdateStatus(pid, newStatus) {
            var statusNames = {
                'Submitted': '提交审批', 'Approved': '审批通过', 'Rejected': '拒绝',
                'Ordered': '确认下单', 'Completed': '标记完成', 'Cancelled': '取消订单'
            };
            if (confirm('确定要' + (statusNames[newStatus] || newStatus) + '吗？')) {
                var f = document.createElement('form');
                f.method = 'post';
                f.innerHTML = '<input type="hidden" name="csrf_token" value="<%= Session("CSRFToken") %>"><input type="hidden" name="action" value="update_status"><input type="hidden" name="purchase_id" value="' + pid + '"><input type="hidden" name="new_status" value="' + newStatus + '">';
                document.body.appendChild(f);
                f.submit();
            }
        }
        
        document.getElementById('createModal').addEventListener('click', function(e) { if (e.target === this) closeCreateModal(); });
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
