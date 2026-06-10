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
' ========== 仪表盘统计 ==========
Dim totalProducts : totalProducts = SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandProducts WHERE Status='Active'"))
Dim pendingOrders : pendingOrders = SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandPurchaseOrders WHERE Status IN ('Submitted','Approved')"))
Dim pendingReceipt : pendingReceipt = SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandPurchaseOrders WHERE Status IN ('Ordered','PartialReceived')"))
Dim completedOrders : completedOrders = SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandPurchaseOrders WHERE Status IN ('Received','Completed')"))
Dim monthlyCost : monthlyCost = SafeNum(GetScalar("SELECT ISNULL(SUM(TotalAmount),0) FROM FixedBrandPurchaseOrders WHERE Status IN ('Ordered','PartialReceived','Received','Completed') AND Month(OrderDate)=Month(GETDATE()) AND Year(OrderDate)=Year(GETDATE())"))
Dim lowStockCount : lowStockCount = SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandInventory WHERE StockQty <= SafetyStock"))
Dim totalInventory : totalInventory = SafeNum(GetScalar("SELECT ISNULL(SUM(StockQty),0) FROM FixedBrandInventory"))
Dim totalCost : totalCost = SafeNum(GetScalar("SELECT ISNULL(SUM(TotalCost),0) FROM FixedBrandCostAllocation"))
Dim totalProfit : totalProfit = SafeNum(GetScalar("SELECT ISNULL(SUM(ProfitAmount),0) FROM FixedBrandCostAllocation"))

' ========== 最近订单 ==========
Dim rsRecent : Set rsRecent = conn.Execute("SELECT TOP 5 * FROM FixedBrandPurchaseOrders ORDER BY PurchaseID DESC")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>品牌定香采购 - 采购管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; }
        .main-content { margin-left: 270px; padding: 25px; min-height: 100vh; }
        
        .page-header { margin-bottom: 25px; }
        .page-title { font-size: 22px; font-weight: 600; color: #fff; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #FF9800; }
        .page-subtitle { font-size: 13px; color: #888; margin-top: 5px; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; padding: 22px; border: 1px solid rgba(255,255,255,0.05); cursor: pointer; text-decoration: none; color: inherit; display: block; transition: all 0.3s; }
        .stat-card:hover { transform: translateY(-3px); box-shadow: 0 6px 25px rgba(0,0,0,0.4); }
        .stat-icon { width: 48px; height: 48px; border-radius: 10px; display: flex; align-items: center; justify-content: center; font-size: 20px; margin-bottom: 12px; }
        .stat-value { font-size: 26px; font-weight: 700; color: #fff; }
        .stat-label { font-size: 13px; color: #888; margin-top: 4px; }
        
        .quick-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 30px; }
        .quick-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; padding: 25px; text-align: center; border: 1px solid rgba(255,255,255,0.05); transition: all 0.3s; text-decoration: none; color: inherit; }
        .quick-card:hover { transform: translateY(-4px); box-shadow: 0 8px 30px rgba(255,152,0,0.2); border-color: rgba(255,152,0,0.3); }
        .quick-icon { width: 60px; height: 60px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 0 auto 12px; font-size: 24px; color: white; }
        .quick-title { font-size: 16px; font-weight: 600; color: #fff; margin-bottom: 6px; }
        .quick-desc { font-size: 12px; color: #888; }
        
        .bottom-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        .panel { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 10px; padding: 20px; border: 1px solid rgba(255,255,255,0.05); }
        .panel h3 { color: #fff; font-size: 15px; margin: 0 0 15px; display: flex; align-items: center; gap: 8px; }
        .panel h3 i { color: #FF9800; }
        
        .data-table { width: 100%; border-collapse: collapse; }
        .data-table th, .data-table td { padding: 10px 14px; text-align: left; border-bottom: 1px solid rgba(255,255,255,0.05); font-size: 13px; }
        .data-table th { color: #888; font-size: 11px; font-weight: 600; }
        .data-table td { color: #ccc; }
        .data-table tr:hover td { background: rgba(255,255,255,0.02); }
        
        .status-badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 500; }
        .status-draft { background: rgba(158,158,158,0.2); color: #9E9E9E; }
        .status-submitted { background: rgba(255,152,0,0.2); color: #FF9800; }
        .status-approved { background: rgba(33,150,243,0.2); color: #2196F3; }
        .status-ordered { background: rgba(156,39,176,0.2); color: #9C27B0; }
        .status-received { background: rgba(76,175,80,0.2); color: #4CAF50; }
        .status-completed { background: rgba(76,175,80,0.2); color: #4CAF50; }
        
        .warning-badge { display: inline-block; padding: 3px 8px; border-radius: 10px; font-size: 11px; background: rgba(244,67,54,0.2); color: #F44336; }
    </style>
</head>
<body data-theme="purchase-dark">
    <!--#include file="../includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-boxes"></i> 品牌定香采购概览</h2>
            <p class="page-subtitle">品牌定香产品采购、入库、成本与利润全流程管理</p>
        </div>
        
        <div class="stats-grid">
            <div class="stat-card" onclick="location.href='product_management.asp'">
                <div class="stat-icon" style="background:linear-gradient(135deg,#4CAF50,#388E3C);"><i class="fas fa-cubes"></i></div>
                <div class="stat-value"><%= totalProducts %></div>
                <div class="stat-label">活跃产品数</div>
                <% If lowStockCount > 0 Then %>
                <span class="warning-badge"><%= lowStockCount %> 低库存</span>
                <% End If %>
            </div>
            <div class="stat-card" onclick="location.href='purchase_orders.asp?status=Submitted'">
                <div class="stat-icon" style="background:linear-gradient(135deg,#FF9800,#F57C00);"><i class="fas fa-clock"></i></div>
                <div class="stat-value"><%= pendingOrders %></div>
                <div class="stat-label">待处理订单</div>
            </div>
            <div class="stat-card" onclick="location.href='receiving.asp'">
                <div class="stat-icon" style="background:linear-gradient(135deg,#2196F3,#1565C0);"><i class="fas fa-truck-loading"></i></div>
                <div class="stat-value"><%= pendingReceipt %></div>
                <div class="stat-label">待收货订单</div>
            </div>
            <div class="stat-card" onclick="location.href='cost_profit.asp'">
                <div class="stat-icon" style="background:linear-gradient(135deg,#9C27B0,#6A1B9A);"><i class="fas fa-yen-sign"></i></div>
                <div class="stat-value">¥<%= FormatNumber(monthlyCost, 0) %></div>
                <div class="stat-label">本月采购额</div>
            </div>
        </div>
        
        <div class="stats-grid" style="grid-template-columns:repeat(4,1fr);margin-bottom:30px;">
            <div class="stat-card">
                <div class="stat-value" style="font-size:20px;"><%= completedOrders %></div>
                <div class="stat-label">已完成订单</div>
            </div>
            <div class="stat-card">
                <div class="stat-value" style="font-size:20px;"><%= totalInventory %></div>
                <div class="stat-label">总库存量(件)</div>
            </div>
            <div class="stat-card">
                <div class="stat-value" style="font-size:20px;">¥<%= FormatNumber(totalCost, 0) %></div>
                <div class="stat-label">已分摊总成本</div>
            </div>
            <div class="stat-card">
                <div class="stat-value" style="font-size:20px;color:<%=IIf(totalProfit>=0,"#4CAF50","#F44336")%>;">¥<%= FormatNumber(totalProfit, 0) %></div>
                <div class="stat-label">累计利润</div>
            </div>
        </div>
        
        <div class="quick-grid">
            <a href="product_management.asp" class="quick-card">
                <div class="quick-icon" style="background:linear-gradient(135deg,#4CAF50,#388E3C);"><i class="fas fa-boxes"></i></div>
                <div class="quick-title">产品管理</div>
                <div class="quick-desc">管理品牌定香产品目录和价格</div>
            </a>
            <a href="purchase_orders.asp" class="quick-card">
                <div class="quick-icon" style="background:linear-gradient(135deg,#FF9800,#F57C00);"><i class="fas fa-file-invoice"></i></div>
                <div class="quick-title">采购订单</div>
                <div class="quick-desc">创建、审批和跟踪采购订单</div>
            </a>
            <a href="receiving.asp" class="quick-card">
                <div class="quick-icon" style="background:linear-gradient(135deg,#2196F3,#1565C0);"><i class="fas fa-clipboard-check"></i></div>
                <div class="quick-title">收货入库</div>
                <div class="quick-desc">到货验收与库存更新</div>
            </a>
            <a href="replenishment.asp" class="quick-card">
                <div class="quick-icon" style="background:linear-gradient(135deg,#9C27B0,#6A1B9A);"><i class="fas fa-robot"></i></div>
                <div class="quick-title">智能补货</div>
                <div class="quick-desc">低库存预警与自动补货建议</div>
            </a>
        </div>
        
        <div class="bottom-grid">
            <div class="panel">
                <h3><i class="fas fa-history"></i> 最近采购订单</h3>
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>采购单号</th>
                            <th>供应商</th>
                            <th>金额</th>
                            <th>状态</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% If Not rsRecent Is Nothing Then
                            If Not rsRecent.EOF Then
                                Do While Not rsRecent.EOF
                                    Dim rStatus : rStatus = CStr(rsRecent("Status"))
                                    Dim rClass : rClass = "status-draft"
                                    Dim rName : rName = "草稿"
                                    Select Case rStatus
                                        Case "Submitted" : rClass = "status-submitted" : rName = "待审批"
                                        Case "Approved" : rClass = "status-approved" : rName = "已审批"
                                        Case "Ordered" : rClass = "status-ordered" : rName = "已下单"
                                        Case "Received" : rClass = "status-received" : rName = "已收货"
                                        Case "Completed" : rClass = "status-completed" : rName = "已完成"
                                    End Select
                        %>
                        <tr style="cursor:pointer;" onclick="location.href='purchase_orders.asp'">
                            <td style="font-family:Consolas,monospace;color:#FF9800;font-size:12px;"><%= Server.HTMLEncode(CStr(rsRecent("PurchaseNo"))) %></td>
                            <td><%= Server.HTMLEncode(CStr(rsRecent("SupplierName") & "")) %></td>
                            <td>¥<%= FormatNumber(SafeNum(rsRecent("TotalAmount")), 2) %></td>
                            <td><span class="status-badge <%= rClass %>"><%= rName %></span></td>
                        </tr>
                        <%
                                    rsRecent.MoveNext
                                Loop
                            Else
                        %>
                        <tr><td colspan="4" style="text-align:center;color:#666;padding:20px;">暂无订单</td></tr>
                        <%  End If
                        End If
                        If Not rsRecent Is Nothing Then
                            If rsRecent.State = 1 Then rsRecent.Close
                            Set rsRecent = Nothing
                        End If %>
                    </tbody>
                </table>
            </div>
            
            <div class="panel">
                <h3><i class="fas fa-chart-line"></i> 快速成本概览</h3>
                <div style="display:flex;flex-direction:column;gap:15px;">
                    <div style="display:flex;justify-content:space-between;align-items:center;padding:10px;background:rgba(255,255,255,0.03);border-radius:8px;">
                        <span style="font-size:13px;color:#888;">库存资产价值</span>
                        <span style="font-size:15px;font-weight:600;color:#fff;">¥<%= FormatNumber(SafeNum(GetScalar("SELECT ISNULL(SUM(StockQty*AvgUnitCost),0) FROM FixedBrandInventory")), 0) %></span>
                    </div>
                    <div style="display:flex;justify-content:space-between;align-items:center;padding:10px;background:rgba(255,255,255,0.03);border-radius:8px;">
                        <span style="font-size:13px;color:#888;">已分摊成本</span>
                        <span style="font-size:15px;font-weight:600;color:#FF9800;">¥<%= FormatNumber(totalCost, 0) %></span>
                    </div>
                    <div style="display:flex;justify-content:space-between;align-items:center;padding:10px;background:rgba(255,255,255,0.03);border-radius:8px;">
                        <span style="font-size:13px;color:#888;">累计利润</span>
                        <span style="font-size:15px;font-weight:600;color:<%=IIf(totalProfit>=0,"#4CAF50","#F44336")%>;">¥<%= FormatNumber(totalProfit, 0) %></span>
                    </div>
                    <a href="cost_profit.asp" class="btn btn--primary btn--sm" style="align-self:flex-end;">查看详细分析</a>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
