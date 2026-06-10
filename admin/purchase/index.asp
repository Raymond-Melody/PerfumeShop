<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<%
Call OpenConnection()

' ========== SafeNum函数：安全处理数值空值 ==========
Function SafeNum(val)
    If IsNull(val) Then
        SafeNum = 0
    ElseIf val = "" Then
        SafeNum = 0
    ElseIf Not IsNumeric(val) Then
        SafeNum = 0
    Else
        SafeNum = CDbl(val)
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

' ========== 获取当前Tab（默认RawMaterial）==========
Dim activeTab
activeTab = Request.QueryString("tab")
If activeTab = "" Then activeTab = "RawMaterial"

' ========== 辅助函数：获取品类名称 ==========
Function GetOrderTypeName(ot)
    If ot = "RawMaterial" Then
        GetOrderTypeName = "原料采购"
    ElseIf ot = "Packaging" Then
        GetOrderTypeName = "包装物采购"
    ElseIf ot = "Bottle" Then
        GetOrderTypeName = "瓶子采购"
    ElseIf ot = "Printing" Then
        GetOrderTypeName = "印刷品采购"
    ElseIf ot = "SprayHead" Then
        GetOrderTypeName = "喷头采购"
    Else
        GetOrderTypeName = ot
    End If
End Function

' ========== 全局统计卡片 ==========
Dim pendingApproval, monthlyAmount, supplierCount, pendingReceipt
pendingApproval = GetScalar("SELECT COUNT(*) FROM PurchaseOrders WHERE Status='Submitted'")
monthlyAmount = GetScalar("SELECT SUM(CAST(ISNULL(TotalAmount,0) AS FLOAT)) FROM PurchaseOrders WHERE (Status='Ordered' OR Status='Received' OR Status='Completed') AND Month(OrderDate)=Month(GETDATE()) AND Year(OrderDate)=Year(GETDATE())")
On Error Resume Next : supplierCount = GetScalar("SELECT COUNT(*) FROM Suppliers") : If Err.Number <> 0 Then supplierCount = 0 : Err.Clear : On Error GoTo 0
pendingReceipt = GetScalar("SELECT COUNT(*) FROM PurchaseOrders WHERE Status='Ordered'")

' ========== 按品类统计 ==========
Dim rawPOCount, rawPOAmount, rawLowStock, rawTotal
rawPOCount = GetScalar("SELECT COUNT(*) FROM PurchaseOrders WHERE OrderType='RawMaterial' AND Status IN ('Ordered','PartialReceived','Received')")
rawPOAmount = GetScalar("SELECT SUM(CAST(ISNULL(TotalAmount,0) AS FLOAT)) FROM PurchaseOrders WHERE OrderType='RawMaterial' AND Month(OrderDate)=Month(GETDATE()) AND Year(OrderDate)=Year(GETDATE())")
rawLowStock = SafeNum(GetScalar("SELECT COUNT(*) FROM RawMaterialInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0"))
rawTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM RawMaterialInventory"))

Dim packPOCount, packPOAmount, packLowStock, packTotal
packPOCount = GetScalar("SELECT COUNT(*) FROM PurchaseOrders WHERE OrderType='Packaging' AND Status IN ('Ordered','PartialReceived','Received')")
packPOAmount = GetScalar("SELECT SUM(CAST(ISNULL(TotalAmount,0) AS FLOAT)) FROM PurchaseOrders WHERE OrderType='Packaging' AND Month(OrderDate)=Month(GETDATE()) AND Year(OrderDate)=Year(GETDATE())")
On Error Resume Next : packLowStock = SafeNum(GetScalar("SELECT COUNT(*) FROM PackagingInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0")) : If Err.Number <> 0 Then packLowStock = 0 : Err.Clear
packTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM PackagingInventory")) : If Err.Number <> 0 Then packTotal = 0 : Err.Clear : On Error GoTo 0

Dim bottlePOCount, bottlePOAmount, bottleLowStock, bottleTotal
bottlePOCount = GetScalar("SELECT COUNT(*) FROM PurchaseOrders WHERE OrderType='Bottle' AND Status IN ('Ordered','PartialReceived','Received')")
bottlePOAmount = GetScalar("SELECT SUM(CAST(ISNULL(TotalAmount,0) AS FLOAT)) FROM PurchaseOrders WHERE OrderType='Bottle' AND Month(OrderDate)=Month(GETDATE()) AND Year(OrderDate)=Year(GETDATE())")
On Error Resume Next : bottleLowStock = SafeNum(GetScalar("SELECT COUNT(*) FROM BottleStyles WHERE StockQty <= SafetyStock AND SafetyStock > 0")) : If Err.Number <> 0 Then bottleLowStock = 0 : Err.Clear
bottleTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM BottleStyles")) : If Err.Number <> 0 Then bottleTotal = 0 : Err.Clear : On Error GoTo 0

Dim printPOCount, printPOAmount, printLowStock, printTotal
printPOCount = GetScalar("SELECT COUNT(*) FROM PurchaseOrders WHERE OrderType='Printing' AND Status IN ('Ordered','PartialReceived','Received')")
printPOAmount = GetScalar("SELECT SUM(CAST(ISNULL(TotalAmount,0) AS FLOAT)) FROM PurchaseOrders WHERE OrderType='Printing' AND Month(OrderDate)=Month(GETDATE()) AND Year(OrderDate)=Year(GETDATE())")
On Error Resume Next : printLowStock = SafeNum(GetScalar("SELECT COUNT(*) FROM PrintingInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0")) : If Err.Number <> 0 Then printLowStock = 0 : Err.Clear
printTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM PrintingInventory")) : If Err.Number <> 0 Then printTotal = 0 : Err.Clear : On Error GoTo 0

Dim sprayPOCount, sprayPOAmount, sprayLowStock, sprayTotal
sprayPOCount = GetScalar("SELECT COUNT(*) FROM PurchaseOrders WHERE OrderType='SprayHead' AND Status IN ('Ordered','PartialReceived','Received')")
sprayPOAmount = GetScalar("SELECT SUM(CAST(ISNULL(TotalAmount,0) AS FLOAT)) FROM PurchaseOrders WHERE OrderType='SprayHead' AND Month(OrderDate)=Month(GETDATE()) AND Year(OrderDate)=Year(GETDATE())")
On Error Resume Next : sprayLowStock = SafeNum(GetScalar("SELECT COUNT(*) FROM SprayHeadInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0")) : If Err.Number <> 0 Then sprayLowStock = 0 : Err.Clear
sprayTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM SprayHeadInventory")) : If Err.Number <> 0 Then sprayTotal = 0 : Err.Clear : On Error GoTo 0

' ========== 最近采购单（按当前Tab过滤）==========
Dim rsRecent
Set rsRecent = ExecuteQuery("SELECT TOP 5 PurchaseID, OrderNo, SupplierName, TotalAmount, Status, OrderDate FROM PurchaseOrders WHERE OrderType='" & activeTab & "' ORDER BY OrderDate DESC")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>采购概览 - 采购管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        /* 暗色主题基础 */
        body {
            background: #1a1a2e;
            color: #e0e0e0;
        }
        
        /* 顶部统计卡片 */
        .stats-section {
            margin-bottom: 30px;
        }
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 20px;
        }
        .stats-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 25px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            border: 1px solid rgba(255,255,255,0.05);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        .stats-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 6px 25px rgba(0,0,0,0.4);
        }
        .stats-header {
            display: flex;
            align-items: center;
            margin-bottom: 15px;
        }
        .stats-icon {
            width: 48px;
            height: 48px;
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 20px;
            margin-right: 12px;
        }
        /* 采购管理中心使用橙色系 */
        .stats-icon.pending { background: linear-gradient(135deg, #FF9800 0%, #F57C00 100%); }
        .stats-icon.amount { background: linear-gradient(135deg, #FF5722 0%, #D84315 100%); }
        .stats-icon.supplier { background: linear-gradient(135deg, #FFB74D 0%, #FF9800 100%); }
        .stats-icon.receipt { background: linear-gradient(135deg, #FFA726 0%, #EF6C00 100%); }
        .stats-label {
            font-size: 13px;
            color: #888;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .stats-value {
            font-size: 28px;
            font-weight: 700;
            color: #fff;
            margin-top: 5px;
        }
        .stats-sub {
            font-size: 12px;
            color: #666;
            margin-top: 8px;
        }
        
        /* 品类Tab导航 */
        .tab-nav {
            display: flex;
            gap: 0;
            margin-bottom: 20px;
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            overflow: hidden;
            border: 1px solid rgba(255,255,255,0.05);
        }
        .tab-link {
            flex: 1;
            padding: 18px 20px;
            text-align: center;
            color: #888;
            text-decoration: none;
            font-size: 15px;
            font-weight: 500;
            transition: all 0.3s ease;
            border-bottom: 3px solid transparent;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
        }
        .tab-link:hover { color: #e0e0e0; background: rgba(255,255,255,0.03); }
        .tab-link.active {
            color: #fff;
            border-bottom-color: #FF9800;
            background: rgba(255,152,0,0.1);
        }
        .tab-link.raw i { color: #4CAF50; }
        .tab-link.pack i { color: #2196F3; }
        .tab-link.bottle i { color: #9C27B0; }
        .tab-link.print i { color: #00BCD4; }
        .tab-link.spray i { color: #FF5722; }
        .tab-badge {
            display: inline-block;
            padding: 2px 8px;
            border-radius: 10px;
            font-size: 11px;
            font-weight: 600;
            background: rgba(255,255,255,0.1);
        }
        
        /* 中部快捷入口 */
        .quick-section {
            margin-bottom: 30px;
        }
        .section-title {
            font-size: 18px;
            font-weight: 600;
            margin-bottom: 20px;
            color: #fff;
            display: flex;
            align-items: center;
        }
        .section-title i {
            margin-right: 10px;
            color: #FF9800;
        }
        .quick-grid {
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 20px;
        }
        .quick-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 30px;
            text-align: center;
            border: 1px solid rgba(255,255,255,0.05);
            transition: all 0.3s ease;
            cursor: pointer;
            text-decoration: none;
            color: inherit;
            display: block;
        }
        .quick-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 10px 30px rgba(255,152,0,0.2);
            border-color: rgba(255,152,0,0.3);
        }
        .quick-icon {
            width: 70px;
            height: 70px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            margin: 0 auto 15px;
            font-size: 28px;
            color: white;
        }
        .quick-icon.orders { background: linear-gradient(135deg, #FF9800 0%, #F57C00 100%); }
        .quick-icon.suppliers { background: linear-gradient(135deg, #FF5722 0%, #D84315 100%); }
        .quick-icon.prices { background: linear-gradient(135deg, #FFB74D 0%, #FF9800 100%); }
        .quick-icon.receiving { background: linear-gradient(135deg, #4CAF50 0%, #388E3C 100%); }
        .quick-icon.fixedbrand { background: linear-gradient(135deg, #E91E63 0%, #AD1457 100%); }
        .quick-title {
            font-size: 18px;
            font-weight: 600;
            color: #fff;
            margin-bottom: 8px;
        }
        .quick-desc {
            font-size: 13px;
            color: #888;
        }
        
        /* 品类统计卡片 */
        .category-stats-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 15px;
            margin-bottom: 30px;
        }
        .cat-stat-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 10px;
            padding: 20px;
            border: 1px solid rgba(255,255,255,0.05);
        }
        .cat-stat-label { font-size: 12px; color: #888; margin-bottom: 8px; }
        .cat-stat-value { font-size: 24px; font-weight: 700; color: #fff; }
        
        /* 底部最近采购单 */
        .recent-section {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 20px;
            border: 1px solid rgba(255,255,255,0.05);
        }
        .recent-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 15px;
        }
        .recent-table th,
        .recent-table td {
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid rgba(255,255,255,0.05);
        }
        .recent-table th {
            font-size: 12px;
            color: #888;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            font-weight: 600;
        }
        .recent-table td {
            font-size: 14px;
            color: #e0e0e0;
        }
        .recent-table tr:hover td {
            background: rgba(255,255,255,0.02);
        }
        .status-badge {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: 500;
        }
        .status-submitted { background: rgba(255,152,0,0.2); color: #FF9800; }
        .status-ordered { background: rgba(33,150,243,0.2); color: #2196F3; }
        .status-received { background: rgba(76,175,80,0.2); color: #4CAF50; }
        .status-completed { background: rgba(156,39,176,0.2); color: #9C27B0; }
        .status-cancelled { background: rgba(244,67,54,0.2); color: #F44336; }
        .empty-row {
            text-align: center;
            color: #666;
            padding: 40px;
        }
        
        /* 响应式 */
        @media (max-width: 1200px) {
            .stats-grid { grid-template-columns: repeat(2, 1fr); }
            .quick-grid { grid-template-columns: 1fr; }
        }
        @media (max-width: 768px) {
            .stats-grid { grid-template-columns: 1fr; }
        }
    </style>
</head>
<body data-theme="purchase-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-shopping-cart"></i> 采购概览</h2>
            <div class="breadcrumb">
                <a href="index.asp">采购中心</a> / <span>概览</span>
            </div>
        </div>
        
        <!-- 顶部统计卡片 -->
        <div class="stats-section">
            <div class="stats-grid">
                <div class="stats-card">
                    <div class="stats-header">
                        <div class="stats-icon pending"><i class="fas fa-clock"></i></div>
                        <div class="stats-label">待审批采购单</div>
                    </div>
                    <div class="stats-value"><%= pendingApproval %></div>
                    <div class="stats-sub">等待审批的采购申请</div>
                </div>
                <div class="stats-card">
                    <div class="stats-header">
                        <div class="stats-icon amount"><i class="fas fa-yen-sign"></i></div>
                        <div class="stats-label">本月采购额</div>
                    </div>
                    <div class="stats-value">¥<%= FormatNumber(SafeNum(monthlyAmount), 0) %></div>
                    <div class="stats-sub">本月已确认订单总额</div>
                </div>
                <div class="stats-card">
                    <div class="stats-header">
                        <div class="stats-icon supplier"><i class="fas fa-truck"></i></div>
                        <div class="stats-label">供应商数</div>
                    </div>
                    <div class="stats-value"><%= supplierCount %></div>
                    <div class="stats-sub">合作供应商总数</div>
                </div>
                <div class="stats-card">
                    <div class="stats-header">
                        <div class="stats-icon receipt"><i class="fas fa-box"></i></div>
                        <div class="stats-label">待收货数</div>
                    </div>
                    <div class="stats-value"><%= pendingReceipt %></div>
                    <div class="stats-sub">已下单待收货订单</div>
                </div>
            </div>
        </div>
        
        <!-- 品类Tab导航 -->
        <div class="tab-nav">
            <a href="?tab=RawMaterial" class="tab-link raw <%= IIf(activeTab="RawMaterial","active","") %>">
                <i class="fas fa-flask"></i> 原料采购
                <span class="tab-badge"><%= rawPOCount %> 进行中</span>
            </a>
            <a href="?tab=Packaging" class="tab-link pack <%= IIf(activeTab="Packaging","active","") %>">
                <i class="fas fa-box"></i> 包装物采购
                <span class="tab-badge"><%= packPOCount %> 进行中</span>
            </a>
            <a href="?tab=Bottle" class="tab-link bottle <%= IIf(activeTab="Bottle","active","") %>">
                <i class="fas fa-wine-bottle"></i> 瓶子采购
                <span class="tab-badge"><%= bottlePOCount %> 进行中</span>
            </a>
            <a href="?tab=Printing" class="tab-link print <%= IIf(activeTab="Printing","active","") %>">
                <i class="fas fa-print"></i> 印刷品采购
                <span class="tab-badge"><%= printPOCount %> 进行中</span>
            </a>
            <a href="?tab=SprayHead" class="tab-link spray <%= IIf(activeTab="SprayHead","active","") %>">
                <i class="fas fa-spray-can"></i> 喷头采购
                <span class="tab-badge"><%= sprayPOCount %> 进行中</span>
            </a>
        </div>
        
        <!-- 当前品类统计 -->
        <% Dim catPOCount, catPOAmount, catLowStock, catTotal, catIcon, catColor
        If activeTab = "RawMaterial" Then
            catPOCount = rawPOCount : catPOAmount = rawPOAmount : catLowStock = rawLowStock : catTotal = rawTotal
            catIcon = "fa-flask" : catColor = "#4CAF50"
        ElseIf activeTab = "Packaging" Then
            catPOCount = packPOCount : catPOAmount = packPOAmount : catLowStock = packLowStock : catTotal = packTotal
            catIcon = "fa-box" : catColor = "#2196F3"
        ElseIf activeTab = "Bottle" Then
            catPOCount = bottlePOCount : catPOAmount = bottlePOAmount : catLowStock = bottleLowStock : catTotal = bottleTotal
            catIcon = "fa-wine-bottle" : catColor = "#9C27B0"
        ElseIf activeTab = "Printing" Then
            catPOCount = printPOCount : catPOAmount = printPOAmount : catLowStock = printLowStock : catTotal = printTotal
            catIcon = "fa-print" : catColor = "#00BCD4"
        Else
            catPOCount = sprayPOCount : catPOAmount = sprayPOAmount : catLowStock = sprayLowStock : catTotal = sprayTotal
            catIcon = "fa-spray-can" : catColor = "#FF5722"
        End If %>
        <div class="category-stats-grid">
            <div class="cat-stat-card">
                <div class="cat-stat-label"><i class="fas fa-file-invoice" style="color:<%=catColor%>"></i> 进行中订单</div>
                <div class="cat-stat-value"><%= catPOCount %></div>
            </div>
            <div class="cat-stat-card">
                <div class="cat-stat-label"><i class="fas fa-yen-sign" style="color:<%=catColor%>"></i> 本月采购额</div>
                <div class="cat-stat-value">¥<%= FormatNumber(SafeNum(catPOAmount), 0) %></div>
            </div>
            <div class="cat-stat-card" <% If catLowStock > 0 Then %>style="border:1px solid rgba(255,152,0,0.5);"<% End If %>>
                <div class="cat-stat-label"><i class="fas fa-exclamation-triangle" style="color:#FF9800"></i> 低库存预警</div>
                <div class="cat-stat-value" <% If catLowStock > 0 Then %>style="color:#FF9800;"<% End If %>><%= catLowStock %></div>
            </div>
            <div class="cat-stat-card">
                <div class="cat-stat-label"><i class="fas fa-list-alt" style="color:<%=catColor%>"></i> 库存品种</div>
                <div class="cat-stat-value"><%= catTotal %></div>
            </div>
        </div>
        
        <!-- 中部快捷入口 -->
        <div class="quick-section">
            <div class="section-title"><i class="fas fa-th-large"></i> 快捷入口</div>
            <div class="quick-grid">
                <a href="purchase_orders.asp" class="quick-card">
                    <div class="quick-icon orders"><i class="fas fa-file-invoice"></i></div>
                    <div class="quick-title">采购订单</div>
                    <div class="quick-desc">创建、审批和跟踪采购订单</div>
                </a>
                <a href="supplier_management.asp" class="quick-card">
                    <div class="quick-icon suppliers"><i class="fas fa-truck"></i></div>
                    <div class="quick-title">供应商管理</div>
                    <div class="quick-desc">维护供应商信息和合作关系</div>
                </a>
                <a href="price_management.asp" class="quick-card">
                    <div class="quick-icon prices"><i class="fas fa-tags"></i></div>
                    <div class="quick-title">价格管理</div>
                    <div class="quick-desc">管理采购价格和价格历史</div>
                </a>
                <a href="receiving.asp" class="quick-card">
                    <div class="quick-icon receiving"><i class="fas fa-clipboard-check"></i></div>
                    <div class="quick-title">收货入库</div>
                    <div class="quick-desc">采购到货验收与入库</div>
                </a>
                <a href="fixed_brand/index.asp" class="quick-card">
                    <div class="quick-icon fixedbrand"><i class="fas fa-boxes"></i></div>
                    <div class="quick-title">品牌定香采购</div>
                    <div class="quick-desc">品牌成品独立采购与管理</div>
                </a>
            </div>
        </div>
        
        <!-- 底部最近采购单 -->
        <div class="recent-section">
            <div class="section-title"><i class="fas fa-history"></i> 最近采购单 - <%= GetOrderTypeName(activeTab) %></div>
            <table class="recent-table">
                <thead>
                    <tr>
                        <th>订单号</th>
                        <th>供应商</th>
                        <th>金额</th>
                        <th>状态</th>
                        <th>日期</th>
                    </tr>
                </thead>
                <tbody>
                    <% 
                    ' 使用嵌套If检查Recordset，避免Or不短路问题
                    If rsRecent Is Nothing Then
                    %>
                    <tr>
                        <td colspan="5" class="empty-row">
                            <i class="fas fa-inbox" style="font-size: 24px; margin-bottom: 10px; display: block;"></i>
                            暂无采购订单数据
                        </td>
                    </tr>
                    <% 
                    Else
                        If rsRecent.EOF Then
                    %>
                    <tr>
                        <td colspan="5" class="empty-row">
                            <i class="fas fa-inbox" style="font-size: 24px; margin-bottom: 10px; display: block;"></i>
                            暂无采购订单数据
                        </td>
                    </tr>
                    <% 
                        Else
                            Do While Not rsRecent.EOF
                                Dim statusClass, statusText
                                Dim orderStatus
                                orderStatus = CStr(rsRecent("Status"))
                                If orderStatus = "Submitted" Then
                                    statusClass = "status-submitted"
                                    statusText = "待审批"
                                ElseIf orderStatus = "Ordered" Then
                                    statusClass = "status-ordered"
                                    statusText = "已下单"
                                ElseIf orderStatus = "Received" Then
                                    statusClass = "status-received"
                                    statusText = "已收货"
                                ElseIf orderStatus = "Completed" Then
                                    statusClass = "status-completed"
                                    statusText = "已完成"
                                ElseIf orderStatus = "Cancelled" Then
                                    statusClass = "status-cancelled"
                                    statusText = "已取消"
                                Else
                                    statusClass = "status-submitted"
                                    statusText = orderStatus
                                End If
                    %>
                    <tr>
                        <td><%= Server.HTMLEncode(CStr(rsRecent("OrderNo"))) %></td>
                        <td><%= Server.HTMLEncode(CStr(rsRecent("SupplierName"))) %></td>
                        <td>¥<%= FormatNumber(SafeNum(rsRecent("TotalAmount")), 2) %></td>
                        <td><span class="status-badge <%= statusClass %>"><%= statusText %></span></td>
                        <td><% If IsDate(rsRecent("OrderDate")) Then Response.Write FormatDateTime(rsRecent("OrderDate"), 2) End If %></td>
                    </tr>
                    <% 
                                rsRecent.MoveNext
                            Loop
                            rsRecent.Close
                            Set rsRecent = Nothing
                        End If
                    End If
                    %>
                </tbody>
            </table>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
