<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
' ========== SafeNum函数：安全处理数值空值 ==========
Function SafeNum(val)
    If IsNull(val) Then
        SafeNum = 0
    ElseIf val = "" Then
        SafeNum = 0
    ElseIf Not IsNumeric(val) Then
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

' ========== SafeDiv函数：安全除法，防止除零 ==========
Function SafeDiv(numerator, denominator)
    If SafeNum(denominator) = 0 Then
        SafeDiv = 0
    Else
        SafeDiv = SafeNum(numerator) / SafeNum(denominator)
    End If
End Function

Call OpenConnection()

' ========== 权限检查 ==========
Dim canReview
canReview = False
If Session("AdminRoleCode") = "FIN_MANAGER" Then
    canReview = True
ElseIf Session("AdminRoleCode") = "SUPER_ADMIN" Then
    canReview = True
End If

' ========== 获取当前Tab ==========
Dim currentTab
currentTab = Request.QueryString("tab")
If currentTab = "" Then currentTab = "pending"

' ========== 获取筛选参数 ==========
Dim filterCategory, filterStartDate, filterEndDate, filterPurchaseNo
filterCategory = Trim(Request.QueryString("category"))
filterStartDate = Trim(Request.QueryString("startDate"))
filterEndDate = Trim(Request.QueryString("endDate"))
filterPurchaseNo = Trim(Request.QueryString("purchaseNo"))

' ========== 处理POST请求（PRG模式）==========
Dim action, msg, errMsg
action = Trim(Request.Form("action"))
msg = ""
errMsg = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If action = "approve" AND canReview Then
        ' 审核通过
        Dim approvePurchaseID, reviewAmount, costAllocation, reviewComments
        approvePurchaseID = SafeNum(Request.Form("purchaseID"))
        reviewAmount = SafeNum(Request.Form("reviewAmount"))
        costAllocation = Trim(Request.Form("costAllocation"))
        reviewComments = Trim(Request.Form("reviewComments"))
        
        If approvePurchaseID = 0 Then
            errMsg = "无效的采购单ID"
        ElseIf reviewAmount <= 0 Then
            errMsg = "审核金额必须大于0"
        ElseIf costAllocation = "" Then
            errMsg = "请选择成本归类"
        Else
            ' 开始事务处理
            On Error Resume Next
            
            ' 1. 更新采购订单状态
            Dim updateOrderSQL
            updateOrderSQL = "UPDATE PurchaseOrders SET Status='FinanceApproved', ApprovedBy=" & SafeNum(Session("AdminID")) & ", ApprovedAt=GETDATE(), UpdatedAt= GETDATE() WHERE PurchaseID=" & approvePurchaseID
            ExecuteNonQuery updateOrderSQL
            
            If Err.Number = 0 Then
                ' 2. 插入审核记录
                Dim insertReviewSQL
                insertReviewSQL = "INSERT INTO PurchaseCostReview (PurchaseID, ReviewerID, ReviewStatus, ReviewAmount, CostAllocation, ReviewComments, ReviewedAt, CreatedAt) VALUES (" & _
                    approvePurchaseID & ", " & SafeNum(Session("AdminID")) & ", 'Approved', " & reviewAmount & ", '" & SafeSQL(costAllocation) & "', '" & SafeSQL(reviewComments) & "', GETDATE(), GETDATE())"
                ExecuteNonQuery insertReviewSQL
            End If
            
            If Err.Number = 0 Then
                ' 3. 关联到ProductCosts表（根据成本归类）
                Call LinkToProductCosts(approvePurchaseID, reviewAmount, costAllocation)
            End If
            
            If Err.Number = 0 Then
                Call LogAdminAction("采购成本审核通过", "finance", "PurchaseOrders", approvePurchaseID, "金额:" & reviewAmount & ",归类:" & costAllocation)
                Response.Redirect "purchase_review.asp?tab=pending&msg=" & Server.URLEncode("审核通过成功")
                Response.End
            Else
                errMsg = "审核失败: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
        
        If errMsg <> "" Then
            Response.Redirect "purchase_review.asp?tab=pending&error=" & Server.URLEncode(errMsg)
            Response.End
        End If
        
    ElseIf action = "reject" AND canReview Then
        ' 审核驳回
        Dim rejectPurchaseID, rejectReason
        rejectPurchaseID = SafeNum(Request.Form("purchaseID"))
        rejectReason = Trim(Request.Form("rejectReason"))
        
        If rejectPurchaseID = 0 Then
            errMsg = "无效的采购单ID"
        ElseIf rejectReason = "" Then
            errMsg = "请填写驳回原因"
        Else
            On Error Resume Next
            
            ' 1. 更新采购订单状态为已拒绝
            Dim rejectOrderSQL
            rejectOrderSQL = "UPDATE PurchaseOrders SET Status='Rejected', UpdatedAt= GETDATE() WHERE PurchaseID=" & rejectPurchaseID
            ExecuteNonQuery rejectOrderSQL
            
            If Err.Number = 0 Then
                ' 2. 插入审核记录
                Dim insertRejectSQL
                insertRejectSQL = "INSERT INTO PurchaseCostReview (PurchaseID, ReviewerID, ReviewStatus, ReviewAmount, CostAllocation, ReviewComments, ReviewedAt, CreatedAt) VALUES (" & _
                    rejectPurchaseID & ", " & SafeNum(Session("AdminID")) & ", 'Rejected', 0, '', '" & SafeSQL(rejectReason) & "', GETDATE(), GETDATE())"
                ExecuteNonQuery insertRejectSQL
            End If
            
            If Err.Number = 0 Then
                Call LogAdminAction("采购成本审核驳回", "finance", "PurchaseOrders", rejectPurchaseID, "原因:" & rejectReason)
                Response.Redirect "purchase_review.asp?tab=pending&msg=" & Server.URLEncode("审核驳回成功")
                Response.End
            Else
                errMsg = "驳回失败: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
        
        If errMsg <> "" Then
            Response.Redirect "purchase_review.asp?tab=pending&error=" & Server.URLEncode(errMsg)
            Response.End
        End If
    End If
End If

' ========== 关联到ProductCosts表的子程序 ==========
Sub LinkToProductCosts(purchaseID, amount, allocation)
    On Error Resume Next
    
    ' 获取采购订单详情中的产品信息
    Dim detailSQL, detailRS
    detailSQL = "SELECT ItemName, Quantity, ProductID FROM PurchaseOrderDetails WHERE PurchaseID=" & purchaseID
    Set detailRS = ExecuteQuery(detailSQL)
    
    If Not detailRS Is Nothing Then
        Do While Not detailRS.EOF
            Dim productID, itemName, quantity
            itemName = ""
            quantity = 0
            productID = 0
            
            ' 安全获取字段值
            On Error Resume Next
            If Not IsNull(detailRS("ItemName")) Then
                itemName = CStr(detailRS("ItemName"))
            End If
            If Not IsNull(detailRS("Quantity")) Then
                quantity = SafeNum(detailRS("Quantity"))
            End If
            If Not IsNull(detailRS("ProductID")) Then
                productID = SafeNum(detailRS("ProductID"))
            End If
            On Error GoTo 0
            
            ' 如果有ProductID，则关联成本
            If productID > 0 AND quantity > 0 Then
                Dim unitCost
                unitCost = SafeDiv(amount, quantity)
                
                Dim costType
                If allocation = "PRODUCT_COST" Then
                    costType = "Purchase"
                ElseIf allocation = "OPERATION_COST" Then
                    costType = "Operation"
                ElseIf allocation = "MARKETING_COST" Then
                    costType = "Marketing"
                Else
                    costType = "Other"
                End If
                
                Dim insertCostSQL
                insertCostSQL = "INSERT INTO ProductCosts (ProductID, CostType, UnitCost, TotalCost, Quantity, EffectiveDate, ReferenceID, ReferenceType, CreatedAt) VALUES (" & _
                    productID & ", '" & costType & "', " & unitCost & ", " & amount & ", " & quantity & ", GETDATE(), " & purchaseID & ", 'PurchaseOrder', GETDATE())"
                ExecuteNonQuery insertCostSQL
            End If
            
            detailRS.MoveNext
        Loop
        detailRS.Close
        Set detailRS = Nothing
    End If
    
    On Error GoTo 0
End Sub

' ========== 获取状态中文名称 ==========
Function GetStatusName(statusCode)
    If statusCode = "Draft" Then
        GetStatusName = "草稿"
    ElseIf statusCode = "Submitted" Then
        GetStatusName = "待审核"
    ElseIf statusCode = "FinanceApproved" Then
        GetStatusName = "已审批"
    ElseIf statusCode = "Ordered" Then
        GetStatusName = "已下单"
    ElseIf statusCode = "Received" Then
        GetStatusName = "已收货"
    ElseIf statusCode = "Completed" Then
        GetStatusName = "已完成"
    ElseIf statusCode = "Rejected" Then
        GetStatusName = "已拒绝"
    Else
        GetStatusName = statusCode
    End If
End Function

' ========== 获取状态样式类 ==========
Function GetStatusClass(statusCode)
    If statusCode = "Draft" Then
        GetStatusClass = "status-draft"
    ElseIf statusCode = "Submitted" Then
        GetStatusClass = "status-submitted"
    ElseIf statusCode = "FinanceApproved" Then
        GetStatusClass = "status-approved"
    ElseIf statusCode = "Ordered" Then
        GetStatusClass = "status-ordered"
    ElseIf statusCode = "Received" Then
        GetStatusClass = "status-received"
    ElseIf statusCode = "Completed" Then
        GetStatusClass = "status-completed"
    ElseIf statusCode = "Rejected" Then
        GetStatusClass = "status-rejected"
    Else
        GetStatusClass = "status-draft"
    End If
End Function

' ========== 获取分类名称 ==========
Function GetCategoryName(catCode)
    If catCode = "RAW" Then
        GetCategoryName = "原材料"
    ElseIf catCode = "BASE" Then
        GetCategoryName = "基香原料"
    ElseIf catCode = "PACK" Then
        GetCategoryName = "包装材料"
    ElseIf catCode = "MARKET" Then
        GetCategoryName = "营销物料"
    Else
        GetCategoryName = catCode
    End If
End Function

' ========== 获取供应商名称（带错误保护）==========
Function GetSupplierName(supplierID)
    Dim sql, rs, result
    result = "未知供应商"
    On Error Resume Next
    sql = "SELECT SupplierName FROM Suppliers WHERE SupplierID=" & SafeNum(supplierID)
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        If Not rs.EOF Then
            If Not IsNull(rs("SupplierName")) Then
                result = CStr(rs("SupplierName"))
            End If
        End If
        rs.Close
        Set rs = Nothing
    Else
        Err.Clear
    End If
    On Error GoTo 0
    GetSupplierName = result
End Function

' ========== 获取提交人名称 ==========
Function GetCreatorName(adminID)
    Dim sql, rs, result
    result = "未知"
    On Error Resume Next
    sql = "SELECT AdminName FROM AdminUsers WHERE AdminID=" & SafeNum(adminID)
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        If Not rs.EOF Then
            If Not IsNull(rs("AdminName")) Then
                result = CStr(rs("AdminName"))
            End If
        End If
        rs.Close
        Set rs = Nothing
    Else
        Err.Clear
    End If
    On Error GoTo 0
    GetCreatorName = result
End Function

Call LogAdminAction("访问采购成本审核", "finance", "", "", "")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>采购成本审核 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        .tab-container { margin-bottom: 25px; }
        .tab-nav { display: flex; gap: 5px; border-bottom: 2px solid #3a3a4a; margin-bottom: 25px; }
        .tab-nav a { 
            padding: 15px 25px; color: #888; text-decoration: none; 
            border-bottom: 3px solid transparent; transition: all 0.3s;
            display: flex; align-items: center; gap: 8px;
        }
        .tab-nav a:hover { color: #e0e0e0; background: #2d2d44; }
        .tab-nav a.active { color: #00bcd4; border-bottom-color: #00bcd4; background: #2d2d44; }
        
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        
        .filter-bar { 
            display: flex; gap: 15px; margin-bottom: 20px; flex-wrap: wrap;
            background: #2d2d44; padding: 15px; border-radius: 8px;
        }
        .filter-bar select, .filter-bar input { 
            padding: 10px 15px; border: 1px solid rgba(255,255,255,0.15); 
            border-radius: 6px; background: #1a1a2e; color: #e0e0e0;
        }
        .filter-bar button {
            padding: 10px 20px; background: #00bcd4; color: white; border: none;
            border-radius: 6px; cursor: pointer;
        }
        .filter-bar button:hover { background: #00838f; }
        
        .data-table { width: 100%; border-collapse: collapse; background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; overflow: hidden; }
        .data-table th, .data-table td { padding: 15px; text-align: left; border-bottom: 1px solid rgba(255,255,255,0.06); }
        .data-table th { background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); color: white; font-weight: 600; }
        .data-table td { color: #e0e0e0; }
        .data-table tr:hover { background: rgba(255,255,255,0.05); }
        .data-table a { color: #00bcd4; text-decoration: none; }
        .data-table a:hover { text-decoration: underline; }
        
        .status-badge { 
            display: inline-block; padding: 4px 12px; border-radius: 12px; 
            font-size: 12px; font-weight: 500;
        }
        .status-submitted { background: #1a237e; color: #7986cb; }
        .status-approved { background: #1b5e20; color: #81c784; }
        .status-rejected { background: #5e1b1b; color: #e57373; }
        .status-draft { background: #424242; color: #bdbdbd; }
        
        .btn-small { 
            padding: 6px 12px; border: none; border-radius: 6px; 
            cursor: pointer; font-size: 12px; margin-right: 5px;
        }
        .btn-view { background: #00bcd4; color: white; }
        .btn-approve { background: #4CAF50; color: white; }
        .btn-reject { background: #ff5252; color: white; }
        .btn-small:hover { opacity: 0.9; }
        
        .modal {
            display: none; position: fixed; z-index: 1000; left: 0; top: 0;
            width: 100%; height: 100%; background: rgba(0,0,0,0.7);
        }
        .modal.active { display: flex; align-items: center; justify-content: center; }
        .modal-content {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; width: 90%; max-width: 700px;
            max-height: 90vh; overflow-y: auto; border: 1px solid rgba(255,255,255,0.06);
        }
        .modal-header {
            padding: 20px; border-bottom: 1px solid rgba(255,255,255,0.06);
            display: flex; justify-content: space-between; align-items: center;
        }
        .modal-header h3 { color: #e0e0e0; margin: 0; }
        .modal-close {
            background: none; border: none; color: #888; font-size: 24px;
            cursor: pointer;
        }
        .modal-close:hover { color: #e0e0e0; }
        .modal-body { padding: 20px; }
        .modal-footer {
            padding: 20px; border-top: 1px solid rgba(255,255,255,0.06);
            display: flex; justify-content: flex-end; gap: 10px;
        }
        
        .info-grid {
            display: grid; grid-template-columns: repeat(2, 1fr); gap: 15px;
            margin-bottom: 20px;
        }
        .info-item {
            background: #1a1a2e; padding: 15px; border-radius: 8px;
        }
        .info-label { color: #888; font-size: 12px; margin-bottom: 5px; }
        .info-value { color: #e0e0e0; font-size: 14px; font-weight: 500; }
        
        .detail-table {
            width: 100%; border-collapse: collapse; margin-top: 15px;
            background: #1a1a2e; border-radius: 8px; overflow: hidden;
        }
        .detail-table th, .detail-table td {
            padding: 12px; text-align: left; border-bottom: 1px solid #3a3a4a;
        }
        .detail-table th { color: #888; font-weight: 500; }
        .detail-table td { color: #e0e0e0; }
        
        .form-group { margin-bottom: 20px; }
        .form-group label {
            display: block; margin-bottom: 8px; color: #b0b0b0; font-weight: 500;
        }
        .form-group input, .form-group select, .form-group textarea {
            width: 100%; padding: 12px 15px; border: 2px solid #3a3a4a;
            border-radius: 8px; font-size: 14px; background: #1a1a2e; color: #e0e0e0;
            box-sizing: border-box;
        }
        .form-group textarea { resize: vertical; min-height: 80px; }
        
        .btn-primary {
            background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%);
            color: white; padding: 12px 30px; border: none; border-radius: 8px;
            font-size: 14px; cursor: pointer;
        }
        .btn-danger {
            background: #ff5252; color: white; padding: 12px 30px;
            border: none; border-radius: 8px; font-size: 14px; cursor: pointer;
        }
        .btn-secondary {
            background: #3a3a4a; color: #e0e0e0; padding: 12px 30px;
            border: none; border-radius: 8px; font-size: 14px; cursor: pointer;
        }
        .btn-primary:hover, .btn-danger:hover, .btn-secondary:hover { opacity: 0.9; }
        
        .alert {
            padding: 15px; border-radius: 8px; margin-bottom: 20px;
            display: flex; align-items: center; gap: 10px;
        }
        .alert-success { background: #1b5e20; color: #81c784; }
        .alert-error { background: #5e1b1b; color: #e57373; }
        
        .readonly-mask { 
            position: relative; pointer-events: none; opacity: 0.7;
        }
        .readonly-mask::after { 
            content: "无权限"; position: absolute; top: 50%; left: 50%; 
            transform: translate(-50%, -50%); background: rgba(0,0,0,0.8);
            padding: 10px 20px; border-radius: 6px; color: #888;
        }
        
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin-bottom: 30px; }
        .stat-card { 
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; padding: 25px; 
            text-align: center; border: 1px solid rgba(255,255,255,0.06);
        }
        .stat-value { font-size: 36px; font-weight: bold; margin-bottom: 8px; color: #00bcd4; }
        .stat-label { color: #888; font-size: 14px; }
        
        @media (max-width: 1200px) {
            .stats-grid { grid-template-columns: repeat(2, 1fr); }
            .info-grid { grid-template-columns: 1fr; }
        }
        @media (max-width: 768px) {
            .stats-grid { grid-template-columns: 1fr; }
            .filter-bar { flex-direction: column; }
        }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-clipboard-check"></i> 采购成本审核</h2>
            <div class="breadcrumb">
                <a href="index.asp">财务中心</a> / <span>采购审核</span>
            </div>
        </div>
        
        <% If Request.QueryString("msg") <> "" Then %>
        <div class="alert alert-success"><i class="fas fa-check-circle"></i> <%= Server.HTMLEncode(Request.QueryString("msg")) %></div>
        <% End If %>
        
        <% If Request.QueryString("error") <> "" Then %>
        <div class="alert alert-error"><i class="fas fa-times-circle"></i> <%= Server.HTMLEncode(Request.QueryString("error")) %></div>
        <% End If %>
        
        <!-- Tab导航 -->
        <div class="tab-container">
            <div class="tab-nav">
                <a href="?tab=pending" class="<%= IIf(currentTab="pending", "active", "") %>"><i class="fas fa-clock"></i> 待审核</a>
                <a href="?tab=history" class="<%= IIf(currentTab="history", "active", "") %>"><i class="fas fa-history"></i> 审核历史</a>
            </div>
        </div>
        
        <!-- Tab 1: 待审核列表 -->
        <div class="tab-content <%= IIf(currentTab="pending", "active", "") %>">
            <!-- 筛选栏 -->
            <form method="get" action="purchase_review.asp" class="filter-bar">
                <input type="hidden" name="tab" value="pending">
                <input type="text" name="purchaseNo" placeholder="采购单号" value="<%= Server.HTMLEncode(filterPurchaseNo) %>">
                <select name="category">
                    <option value="">全部分类</option>
                    <option value="RAW" <%= IIf(filterCategory="RAW", "selected", "") %>>原材料</option>
                    <option value="BASE" <%= IIf(filterCategory="BASE", "selected", "") %>>基香原料</option>
                    <option value="PACK" <%= IIf(filterCategory="PACK", "selected", "") %>>包装材料</option>
                    <option value="MARKET" <%= IIf(filterCategory="MARKET", "selected", "") %>>营销物料</option>
                </select>
                <input type="date" name="startDate" placeholder="开始日期" value="<%= Server.HTMLEncode(filterStartDate) %>">
                <input type="date" name="endDate" placeholder="结束日期" value="<%= Server.HTMLEncode(filterEndDate) %>">
                <button type="submit"><i class="fas fa-search"></i> 筛选</button>
                <button type="button" onclick="location.href='purchase_review.asp?tab=pending'" style="background: #3a3a4a;"><i class="fas fa-undo"></i> 重置</button>
            </form>
            
            <!-- 待审核列表 -->
            <table class="data-table">
                <thead>
                    <tr>
                        <th>采购单号</th>
                        <th>供应商</th>
                        <th>分类</th>
                        <th>总金额</th>
                        <th>提交日期</th>
                        <th>提交人</th>
                        <th>状态</th>
                        <th>操作</th>
                    </tr>
                </thead>
                <tbody>
                <% 
                Dim pendingSQL, pendingRS
                pendingSQL = "SELECT PurchaseID, PurchaseNo, SupplierID, CategoryCode, TotalAmount, Status, CreatedBy, CreatedAt FROM PurchaseOrders WHERE Status='Submitted'"
                
                If filterCategory <> "" Then
                    pendingSQL = pendingSQL & " AND CategoryCode='" & SafeSQL(filterCategory) & "'"
                End If
                If filterPurchaseNo <> "" Then
                    pendingSQL = pendingSQL & " AND PurchaseNo LIKE '%" & SafeSQL(filterPurchaseNo) & "%'"
                End If
                If filterStartDate <> "" Then
                    pendingSQL = pendingSQL & " AND CreatedAt >= #" & filterStartDate & "'"
                End If
                If filterEndDate <> "" Then
                    pendingSQL = pendingSQL & " AND CreatedAt < DATEADD(day, 1, #" & filterEndDate & "#)"
                End If
                
                pendingSQL = pendingSQL & " ORDER BY CreatedAt DESC"
                
                Set pendingRS = ExecuteQuery(pendingSQL)
                
                If Not pendingRS Is Nothing Then
                    Do While Not pendingRS.EOF
                        Dim pID, pNo, pSupplierID, pCategory, pAmount, pStatus, pCreator, pDate
                        pID = 0
                        pNo = ""
                        pSupplierID = 0
                        pCategory = ""
                        pAmount = 0
                        pStatus = ""
                        pCreator = 0
                        pDate = ""
                        
                        On Error Resume Next
                        If Not IsNull(pendingRS("PurchaseID")) Then pID = SafeNum(pendingRS("PurchaseID"))
                        If Not IsNull(pendingRS("PurchaseNo")) Then pNo = CStr(pendingRS("PurchaseNo"))
                        If Not IsNull(pendingRS("SupplierID")) Then pSupplierID = SafeNum(pendingRS("SupplierID"))
                        If Not IsNull(pendingRS("CategoryCode")) Then pCategory = CStr(pendingRS("CategoryCode"))
                        If Not IsNull(pendingRS("TotalAmount")) Then pAmount = SafeNum(pendingRS("TotalAmount"))
                        If Not IsNull(pendingRS("Status")) Then pStatus = CStr(pendingRS("Status"))
                        If Not IsNull(pendingRS("CreatedBy")) Then pCreator = SafeNum(pendingRS("CreatedBy"))
                        If Not IsNull(pendingRS("CreatedAt")) Then pDate = CStr(pendingRS("CreatedAt"))
                        On Error GoTo 0
                %>
                    <tr>
                        <td><a href="javascript:void(0)" onclick="openDetailModal(<%= pID %>)"><%= Server.HTMLEncode(pNo) %></a></td>
                        <td><%= Server.HTMLEncode(GetSupplierName(pSupplierID)) %></td>
                        <td><%= Server.HTMLEncode(GetCategoryName(pCategory)) %></td>
                        <td>¥<%= FormatNumber(pAmount, 2) %></td>
                        <td><%= Server.HTMLEncode(pDate) %></td>
                        <td><%= Server.HTMLEncode(GetCreatorName(pCreator)) %></td>
                        <td><span class="status-badge <%= GetStatusClass(pStatus) %>"><%= GetStatusName(pStatus) %></span></td>
                        <td>
                            <button class="btn-small btn-view" onclick="openDetailModal(<%= pID %>)"><i class="fas fa-eye"></i> 查看</button>
                            <% If canReview Then %>
                            <button class="btn-small btn-approve" onclick="openApproveModal(<%= pID %>, <%= pAmount %>)"><i class="fas fa-check"></i> 通过</button>
                            <button class="btn-small btn-reject" onclick="openRejectModal(<%= pID %>)"><i class="fas fa-times"></i> 驳回</button>
                            <% End If %>
                        </td>
                    </tr>
                <% 
                        pendingRS.MoveNext
                    Loop
                    pendingRS.Close
                    Set pendingRS = Nothing
                End If
                %>
                </tbody>
            </table>
        </div>
        
        <!-- Tab 2: 审核历史 -->
        <div class="tab-content <%= IIf(currentTab="history", "active", "") %>">
            <!-- 筛选栏 -->
            <form method="get" action="purchase_review.asp" class="filter-bar">
                <input type="hidden" name="tab" value="history">
                <select name="category">
                    <option value="">全部分类</option>
                    <option value="RAW" <%= IIf(filterCategory="RAW", "selected", "") %>>原材料</option>
                    <option value="BASE" <%= IIf(filterCategory="BASE", "selected", "") %>>基香原料</option>
                    <option value="PACK" <%= IIf(filterCategory="PACK", "selected", "") %>>包装材料</option>
                    <option value="MARKET" <%= IIf(filterCategory="MARKET", "selected", "") %>>营销物料</option>
                </select>
                <input type="date" name="startDate" placeholder="开始日期" value="<%= Server.HTMLEncode(filterStartDate) %>">
                <input type="date" name="endDate" placeholder="结束日期" value="<%= Server.HTMLEncode(filterEndDate) %>">
                <button type="submit"><i class="fas fa-search"></i> 筛选</button>
                <button type="button" onclick="location.href='purchase_review.asp?tab=history'" style="background: #3a3a4a;"><i class="fas fa-undo"></i> 重置</button>
            </form>
            
            <table class="data-table">
                <thead>
                    <tr>
                        <th>采购单号</th>
                        <th>供应商</th>
                        <th>分类</th>
                        <th>审核金额</th>
                        <th>成本归类</th>
                        <th>审核结果</th>
                        <th>审核人</th>
                        <th>审核时间</th>
                    </tr>
                </thead>
                <tbody>
                <% 
                Dim historySQL, historyRS
                historySQL = "SELECT r.PurchaseID, r.ReviewAmount, r.CostAllocation, r.ReviewStatus, r.ReviewComments, r.ReviewedAt, r.ReviewerID, " & _
                    "p.PurchaseNo, p.SupplierID, p.CategoryCode " & _
                    "FROM PurchaseCostReview r INNER JOIN PurchaseOrders p ON r.PurchaseID = p.PurchaseID WHERE 1=1"
                
                If filterCategory <> "" Then
                    historySQL = historySQL & " AND p.CategoryCode='" & SafeSQL(filterCategory) & "'"
                End If
                If filterStartDate <> "" Then
                    historySQL = historySQL & " AND r.ReviewedAt >= #" & filterStartDate & "'"
                End If
                If filterEndDate <> "" Then
                    historySQL = historySQL & " AND r.ReviewedAt < DATEADD(day, 1, #" & filterEndDate & "#)"
                End If
                
                historySQL = historySQL & " ORDER BY r.ReviewedAt DESC"
                
                Set historyRS = ExecuteQuery(historySQL)
                
                If Not historyRS Is Nothing Then
                    Do While Not historyRS.EOF
                        Dim hPurchaseNo, hSupplierID, hCategory, hAmount, hAllocation, hStatus, hReviewer, hDate, hReviewStatus
                        hPurchaseNo = ""
                        hSupplierID = 0
                        hCategory = ""
                        hAmount = 0
                        hAllocation = ""
                        hStatus = ""
                        hReviewer = 0
                        hDate = ""
                        hReviewStatus = ""
                        
                        On Error Resume Next
                        If Not IsNull(historyRS("PurchaseNo")) Then hPurchaseNo = CStr(historyRS("PurchaseNo"))
                        If Not IsNull(historyRS("SupplierID")) Then hSupplierID = SafeNum(historyRS("SupplierID"))
                        If Not IsNull(historyRS("CategoryCode")) Then hCategory = CStr(historyRS("CategoryCode"))
                        If Not IsNull(historyRS("ReviewAmount")) Then hAmount = SafeNum(historyRS("ReviewAmount"))
                        If Not IsNull(historyRS("CostAllocation")) Then hAllocation = CStr(historyRS("CostAllocation"))
                        If Not IsNull(historyRS("ReviewStatus")) Then hReviewStatus = CStr(historyRS("ReviewStatus"))
                        If Not IsNull(historyRS("ReviewerID")) Then hReviewer = SafeNum(historyRS("ReviewerID"))
                        If Not IsNull(historyRS("ReviewedAt")) Then hDate = CStr(historyRS("ReviewedAt"))
                        On Error GoTo 0
                        
                        Dim allocationName
                        If hAllocation = "PRODUCT_COST" Then
                            allocationName = "产品成本"
                        ElseIf hAllocation = "OPERATION_COST" Then
                            allocationName = "运营成本"
                        ElseIf hAllocation = "MARKETING_COST" Then
                            allocationName = "营销成本"
                        Else
                            allocationName = hAllocation
                        End If
                        
                        Dim reviewStatusClass, reviewStatusName
                        If hReviewStatus = "Approved" Then
                            reviewStatusClass = "status-approved"
                            reviewStatusName = "已通过"
                        ElseIf hReviewStatus = "Rejected" Then
                            reviewStatusClass = "status-rejected"
                            reviewStatusName = "已驳回"
                        Else
                            reviewStatusClass = "status-draft"
                            reviewStatusName = hReviewStatus
                        End If
                %>
                    <tr>
                        <td><%= Server.HTMLEncode(hPurchaseNo) %></td>
                        <td><%= Server.HTMLEncode(GetSupplierName(hSupplierID)) %></td>
                        <td><%= Server.HTMLEncode(GetCategoryName(hCategory)) %></td>
                        <td>¥<%= FormatNumber(hAmount, 2) %></td>
                        <td><%= Server.HTMLEncode(allocationName) %></td>
                        <td><span class="status-badge <%= reviewStatusClass %>"><%= reviewStatusName %></span></td>
                        <td><%= Server.HTMLEncode(GetCreatorName(hReviewer)) %></td>
                        <td><%= Server.HTMLEncode(hDate) %></td>
                    </tr>
                <% 
                        historyRS.MoveNext
                    Loop
                    historyRS.Close
                    Set historyRS = Nothing
                End If
                %>
                </tbody>
            </table>
        </div>
    </div>
    
    <!-- 详情弹窗 -->
    <div id="detailModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-file-invoice"></i> 采购订单详情</h3>
                <button class="modal-close" onclick="closeModal('detailModal')">&times;</button>
            </div>
            <div class="modal-body" id="detailContent">
                <!-- 动态加载内容 -->
            </div>
            <div class="modal-footer">
                <button class="btn-secondary" onclick="closeModal('detailModal')">关闭</button>
            </div>
        </div>
    </div>
    
    <!-- 审核通过弹窗 -->
    <div id="approveModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-check-circle"></i> 审核通过</h3>
                <button class="modal-close" onclick="closeModal('approveModal')">&times;</button>
            </div>
            <form method="post" action="purchase_review.asp?tab=pending">
                <%= GetCSRFTokenField() %>
                <input type="hidden" name="action" value="approve">
                <input type="hidden" name="purchaseID" id="approvePurchaseID">
                <div class="modal-body">
                    <div class="form-group">
                        <label>审核金额</label>
                        <input type="number" name="reviewAmount" id="approveAmount" step="0.01" required>
                    </div>
                    <div class="form-group">
                        <label>成本归类</label>
                        <select name="costAllocation" required>
                            <option value="">请选择</option>
                            <option value="PRODUCT_COST">产品成本</option>
                            <option value="OPERATION_COST">运营成本</option>
                            <option value="MARKETING_COST">营销成本</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>审核意见</label>
                        <textarea name="reviewComments" placeholder="请输入审核意见（可选）"></textarea>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn-secondary" onclick="closeModal('approveModal')">取消</button>
                    <button type="submit" class="btn-primary"><i class="fas fa-check"></i> 确认通过</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- 审核驳回弹窗 -->
    <div id="rejectModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-times-circle"></i> 审核驳回</h3>
                <button class="modal-close" onclick="closeModal('rejectModal')">&times;</button>
            </div>
            <form method="post" action="purchase_review.asp?tab=pending">
                <%= GetCSRFTokenField() %>
                <input type="hidden" name="action" value="reject">
                <input type="hidden" name="purchaseID" id="rejectPurchaseID">
                <div class="modal-body">
                    <div class="form-group">
                        <label>驳回原因 <span style="color: #ff5252;">*</span></label>
                        <textarea name="rejectReason" required placeholder="请详细说明驳回原因，以便采购人员修改"></textarea>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn-secondary" onclick="closeModal('rejectModal')">取消</button>
                    <button type="submit" class="btn-danger"><i class="fas fa-times"></i> 确认驳回</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
        function openModal(modalId) {
            document.getElementById(modalId).classList.add('active');
        }
        
        function closeModal(modalId) {
            document.getElementById(modalId).classList.remove('active');
        }
        
        function openDetailModal(purchaseID) {
            // 通过AJAX加载详情
            var xhr = new XMLHttpRequest();
            xhr.open('GET', 'purchase_review_detail.asp?id=' + purchaseID, true);
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4 && xhr.status === 200) {
                    document.getElementById('detailContent').innerHTML = xhr.responseText;
                    openModal('detailModal');
                }
            };
            xhr.send();
        }
        
        function openApproveModal(purchaseID, amount) {
            document.getElementById('approvePurchaseID').value = purchaseID;
            document.getElementById('approveAmount').value = amount.toFixed(2);
            openModal('approveModal');
        }
        
        function openRejectModal(purchaseID) {
            document.getElementById('rejectPurchaseID').value = purchaseID;
            openModal('rejectModal');
        }
        
        // 点击弹窗外部关闭
        document.querySelectorAll('.modal').forEach(function(modal) {
            modal.addEventListener('click', function(e) {
                if (e.target === this) {
                    this.classList.remove('active');
                }
            });
        });
    </script>
</body>
</html>
