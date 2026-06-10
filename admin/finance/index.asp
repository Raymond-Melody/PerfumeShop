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

' 判断当前用户角色 (isManager 已在 nav.asp 中声明)
isManager = False
If Session("AdminRoleCode") = "FIN_MANAGER" Or Session("AdminRoleCode") = "SUPER_ADMIN" Then
    isManager = True
End If

' ========== 顶部：关键指标卡片 ==========
' 总营收：SUM(Orders.TotalAmount) WHERE Status IN ('Paid','Processing','Shipped','Completed')
Dim totalRevenue
totalRevenue = GetScalar("SELECT CAST(IIF(SUM(TotalAmount) IS NULL, 0, SUM(TotalAmount)) AS FLOAT) FROM Orders WHERE Status IN ('Paid','Processing','Shipped','Completed')")

' 总成本：SUM(Orders.CostAmount)
Dim totalCost
totalCost = GetScalar("SELECT CAST(IIF(SUM(CostAmount) IS NULL, 0, SUM(CostAmount)) AS FLOAT) FROM Orders WHERE Status IN ('Paid','Processing','Shipped','Completed')")

' 总退款：SUM(RefundRecords.RefundAmount WHERE Status='Completed')
Dim totalRefund
totalRefund = GetScalar("SELECT IIF(SUM(RefundAmount) IS NULL, 0, SUM(RefundAmount)) FROM RefundRecords WHERE Status='Completed'")

' 总利润 = 总营收 - 总成本 - 总退款
Dim totalProfit
totalProfit = CDbl("0" & totalRevenue) - CDbl("0" & totalCost) - CDbl("0" & totalRefund)

' 利润率
Dim profitMargin
If CDbl("0" & totalRevenue) > 0 Then
    profitMargin = (totalProfit / CDbl("0" & totalRevenue)) * 100
Else
    profitMargin = 0
End If

' 利润率颜色
Dim marginColor, marginClass
If profitMargin >= 20 Then
    marginColor = "#4CAF50"
    marginClass = "margin-green"
ElseIf profitMargin >= 10 Then
    marginColor = "#FFC107"
    marginClass = "margin-yellow"
Else
    marginColor = "#F44336"
    marginClass = "margin-red"
End If

' 待对账数：COUNT(*) FROM ReconciliationLogs WHERE Status NOT IN ('Matched','Resolved')
Dim pendingRecon
pendingRecon = GetScalar("SELECT COUNT(*) FROM ReconciliationLogs WHERE Status NOT IN ('Matched','Resolved')")

' ========== 底部：异常预警汇总 ==========
' 成本异动预警（ProductCosts 本月vs上月波动>5%的商品数）
Dim currentMonth, lastMonth
currentMonth = Year(Now) & Right("0" & Month(Now), 2)
If Month(Now) = 1 Then
    lastMonth = (Year(Now) - 1) & "12"
Else
    lastMonth = Year(Now) & Right("0" & (Month(Now) - 1), 2)
End If

' 成本异动预警（简化查询，避免Access复杂子查询问题）
Dim costAlertCount, rsCost
' 循环变量声明移到循环外（VBScript限制）
Dim pid, currCost, prevCost, costDiff
Set rsCost = ExecuteQuery("SELECT ProductID FROM ProductCosts WHERE Format(CreatedAt,'yyyymm') = '" & currentMonth & "' GROUP BY ProductID")
costAlertCount = 0
If Not rsCost Is Nothing Then
    Do While Not rsCost.EOF
        pid = rsCost("ProductID")
        currCost = GetScalar("SELECT IIF(SUM(TotalCost) IS NULL, 0, SUM(TotalCost)) FROM ProductCosts WHERE ProductID = " & pid & " AND Format(CreatedAt,'yyyymm') = '" & currentMonth & "'")
        prevCost = GetScalar("SELECT IIF(SUM(TotalCost) IS NULL, 0, SUM(TotalCost)) FROM ProductCosts WHERE ProductID = " & pid & " AND Format(CreatedAt,'yyyymm') = '" & lastMonth & "'")
        If CDbl("0" & prevCost) > 0 Then
            costDiff = Abs(CDbl("0" & currCost) - CDbl("0" & prevCost)) / CDbl("0" & prevCost) * 100
            If costDiff > 5 Then
                costAlertCount = costAlertCount + 1
            End If
        End If
        rsCost.MoveNext
    Loop
    rsCost.Close
    Set rsCost = Nothing
End If

' 对账异常（ReconciliationLogs 未解决的异常数）
Dim reconAlertCount
reconAlertCount = GetScalar("SELECT COUNT(*) FROM ReconciliationLogs WHERE Status IN ('Exception','Pending')")

' 预算超支（BudgetPlans 执行率>90%的类目数）
Dim budgetAlertCount
' Access SQL不支持IIF函数在查询中，使用条件表达式替代
budgetAlertCount = GetScalar("SELECT COUNT(*) FROM BudgetPlans WHERE (ActualAmount / IIF(BudgetAmount=0,1,BudgetAmount)) * 100 > 90 AND Status = 'Active'")

' 资金预警（FundAccounts 余额低于阈值的账户数）
Dim fundAlertCount
fundAlertCount = GetScalar("SELECT COUNT(*) FROM FundAccounts WHERE CurrentBalance < AlertThreshold AND Status = 'Active'")

' 预警总数
Dim totalAlerts
totalAlerts = CLng(costAlertCount) + CLng(reconAlertCount) + CLng(budgetAlertCount) + CLng(fundAlertCount)

' ========== 供应链对接：库存价值汇总 ==========
Dim invRawValue, invNoteValue, invProductValue, totalInvValue
invRawValue = 0 : invNoteValue = 0 : invProductValue = 0 : totalInvValue = 0
On Error Resume Next
invRawValue = SafeNum(GetScalar("SELECT SUM(StockQty * UnitPrice) FROM RawMaterialInventory"))
invNoteValue = SafeNum(GetScalar("SELECT SUM(StockQuantity * UnitCost) FROM NoteInventory"))
invProductValue = SafeNum(GetScalar("SELECT SUM(StockQty * UnitCost) FROM ProductInventory"))
totalInvValue = invRawValue + invNoteValue + invProductValue

Dim pendingPayable, pendingReceivable
pendingPayable = SafeNum(GetScalar("SELECT COUNT(*) FROM PurchaseReceipts WHERE Status IN ('Pending','Partial')"))
pendingReceivable = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders WHERE Status IN ('Paid','Processing') AND ShippingStatus<>'Delivered'"))
On Error GoTo 0

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>财务概览 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        /* 暗色主题基础 */
        body {
            background: #1a1a2e;
            color: #e0e0e0;
        }
        
        /* 顶部关键指标卡片 */
        .kpi-section {
            margin-bottom: 30px;
        }
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 20px;
        }
        .kpi-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 25px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            border: 1px solid rgba(255,255,255,0.05);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        .kpi-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 6px 25px rgba(0,0,0,0.4);
        }
        .kpi-header {
            display: flex;
            align-items: center;
            margin-bottom: 15px;
        }
        .kpi-icon {
            width: 48px;
            height: 48px;
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 20px;
            margin-right: 12px;
        }
        .kpi-icon.revenue { background: linear-gradient(135deg, #4CAF50 0%, #2e7d32 100%); }
        .kpi-icon.profit { background: linear-gradient(135deg, #2196F3 0%, #1565c0 100%); }
        .kpi-icon.margin { background: linear-gradient(135deg, #FF9800 0%, #e65100 100%); }
        .kpi-icon.recon { background: linear-gradient(135deg, #9C27B0 0%, #6a1b9a 100%); }
        .kpi-label {
            font-size: 13px;
            color: #888;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .kpi-value {
            font-size: 28px;
            font-weight: 700;
            color: #fff;
            margin-top: 5px;
        }
        .kpi-sub {
            font-size: 12px;
            color: #888;
            margin-top: 8px;
        }
        
        /* 利润率红绿灯 */
        .margin-green { color: #4CAF50 !important; }
        .margin-yellow { color: #FFC107 !important; }
        .margin-red { color: #F44336 !important; }
        
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
            color: #00bcd4;
        }
        .quick-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
        }
        .quick-column {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 20px;
            border: 1px solid rgba(255,255,255,0.05);
        }
        .column-header {
            font-size: 14px;
            font-weight: 600;
            color: #fff;
            margin-bottom: 15px;
            padding-bottom: 10px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
            display: flex;
            align-items: center;
        }
        .column-header i {
            margin-right: 8px;
            width: 20px;
            text-align: center;
        }
        .quick-item {
            display: flex;
            align-items: flex-start;
            padding: 12px;
            border-radius: 8px;
            transition: all 0.2s ease;
            text-decoration: none;
            color: inherit;
            margin-bottom: 8px;
        }
        .quick-item:hover {
            background: rgba(255,255,255,0.05);
        }
        .quick-item-icon {
            width: 36px;
            height: 36px;
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-right: 12px;
            font-size: 14px;
            flex-shrink: 0;
        }
        .quick-item-icon.blue { background: rgba(33,150,243,0.2); color: #2196F3; }
        .quick-item-icon.green { background: rgba(76,175,80,0.2); color: #4CAF50; }
        .quick-item-icon.orange { background: rgba(255,152,0,0.2); color: #FF9800; }
        .quick-item-icon.purple { background: rgba(156,39,176,0.2); color: #9C27B0; }
        .quick-item-icon.red { background: rgba(244,67,54,0.2); color: #F44336; }
        .quick-item-icon.teal { background: rgba(0,150,136,0.2); color: #009688; }
        .quick-item-content {
            flex: 1;
        }
        .quick-item-title {
            font-size: 14px;
            font-weight: 500;
            color: #fff;
            margin-bottom: 3px;
        }
        .quick-item-desc {
            font-size: 11px;
            color: #888;
        }
        
        /* 底部异常预警 */
        .alert-section {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 20px;
            border: 1px solid rgba(255,255,255,0.05);
        }
        .alert-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 20px;
            margin-top: 15px;
        }
        .alert-item {
            display: flex;
            align-items: center;
            padding: 15px;
            background: rgba(0,0,0,0.2);
            border-radius: 8px;
            border-left: 3px solid;
        }
        .alert-item.normal {
            border-left-color: #4CAF50;
        }
        .alert-item.warning {
            border-left-color: #F44336;
            animation: pulse 2s infinite;
        }
        @keyframes pulse {
            0%, 100% { box-shadow: 0 0 0 0 rgba(244,67,54,0.4); }
            50% { box-shadow: 0 0 0 8px rgba(244,67,54,0); }
        }
        .alert-icon {
            width: 40px;
            height: 40px;
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-right: 12px;
            font-size: 16px;
        }
        .alert-icon.normal { background: rgba(76,175,80,0.2); color: #4CAF50; }
        .alert-icon.warning { background: rgba(244,67,54,0.2); color: #F44336; }
        .alert-content {
            flex: 1;
        }
        .alert-label {
            font-size: 12px;
            color: #888;
        }
        .alert-value {
            font-size: 20px;
            font-weight: 700;
            color: #fff;
        }
        .alert-value.normal { color: #4CAF50; }
        .alert-value.warning { color: #F44336; }
        
        /* 响应式 */
        @media (max-width: 1200px) {
            .kpi-grid { grid-template-columns: repeat(2, 1fr); }
            .quick-grid { grid-template-columns: 1fr; }
            .alert-grid { grid-template-columns: repeat(2, 1fr); }
        }
        @media (max-width: 768px) {
            .kpi-grid { grid-template-columns: 1fr; }
            .alert-grid { grid-template-columns: 1fr; }
        }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-chart-line"></i> 财务概览</h2>
            <div class="breadcrumb">
                <a href="index.asp">财务中心</a> / <span>概览</span>
            </div>
        </div>
        
        <!-- 顶部：关键指标卡片 -->
        <div class="kpi-section">
            <div class="kpi-grid">
                <div class="kpi-card">
                    <div class="kpi-header">
                        <div class="kpi-icon revenue"><i class="fas fa-yen-sign"></i></div>
                        <div class="kpi-label">总营收</div>
                    </div>
                    <div class="kpi-value">¥<%= FormatNumber(CDbl("0" & totalRevenue), 0) %></div>
                    <div class="kpi-sub">已付款订单总额</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-header">
                        <div class="kpi-icon profit"><i class="fas fa-wallet"></i></div>
                        <div class="kpi-label">总利润</div>
                    </div>
                    <div class="kpi-value">¥<%= FormatNumber(totalProfit, 0) %></div>
                    <div class="kpi-sub">营收 - 成本 - 退款</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-header">
                        <div class="kpi-icon margin"><i class="fas fa-percentage"></i></div>
                        <div class="kpi-label">利润率</div>
                    </div>
                    <div class="kpi-value <%= marginClass %>"><%= FormatNumber(profitMargin, 1) %>%</div>
                    <div class="kpi-sub">
                        <% If profitMargin >= 20 Then %>
                            <i class="fas fa-check-circle" style="color:#4CAF50"></i> 健康
                        <% ElseIf profitMargin >= 10 Then %>
                            <i class="fas fa-exclamation-circle" style="color:#FFC107"></i> 一般
                        <% Else %>
                            <i class="fas fa-times-circle" style="color:#F44336"></i> 需关注
                        <% End If %>
                    </div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-header">
                        <div class="kpi-icon recon"><i class="fas fa-balance-scale"></i></div>
                        <div class="kpi-label">待对账</div>
                    </div>
                    <div class="kpi-value"><%= pendingRecon %></div>
                    <div class="kpi-sub">未匹配/未解决记录</div>
                </div>
            </div>
        </div>
        
        <!-- 供应链对接：库存价值看板 -->
        <div class="kpi-section">
            <div class="section-title"><i class="fas fa-link" style="color:#00bcd4;"></i> 供应链对接</div>
            <div class="kpi-grid">
                <div class="kpi-card">
                    <div class="kpi-header">
                        <div class="kpi-icon" style="background:linear-gradient(135deg,#FF9800,#F57C00);"><i class="fas fa-cubes"></i></div>
                        <div class="kpi-label">原材料库存价值</div>
                    </div>
                    <div class="kpi-value">¥<%=FormatNumber(invRawValue,0)%></div>
                    <div class="kpi-sub">RawMaterialInventory</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-header">
                        <div class="kpi-icon" style="background:linear-gradient(135deg,#9C27B0,#7B1FA2);"><i class="fas fa-wine-bottle"></i></div>
                        <div class="kpi-label">香调库存价值</div>
                    </div>
                    <div class="kpi-value">¥<%=FormatNumber(invNoteValue,0)%></div>
                    <div class="kpi-sub">NoteInventory</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-header">
                        <div class="kpi-icon" style="background:linear-gradient(135deg,#2196F3,#1976D2);"><i class="fas fa-box"></i></div>
                        <div class="kpi-label">成品库存价值</div>
                    </div>
                    <div class="kpi-value">¥<%=FormatNumber(invProductValue,0)%></div>
                    <div class="kpi-sub">ProductInventory</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-header">
                        <div class="kpi-icon" style="background:linear-gradient(135deg,#4CAF50,#388E3C);"><i class="fas fa-chart-pie"></i></div>
                        <div class="kpi-label">库存总价值</div>
                    </div>
                    <div class="kpi-value">¥<%=FormatNumber(totalInvValue,0)%></div>
                    <div class="kpi-sub">应付款<%=pendingPayable%>单 | 应收款<%=pendingReceivable%>单</div>
                </div>
            </div>
        </div>
        
        <!-- 中部：三层快捷入口 -->
        <div class="quick-section">
            <div class="section-title"><i class="fas fa-th-large"></i> 快捷入口</div>
            <div class="quick-grid">
                <!-- 第一列：财务中台 -->
                <div class="quick-column">
                    <div class="column-header"><i class="fas fa-cogs" style="color:#2196F3"></i> 财务中台</div>
                    <a href="cost_management.asp" class="quick-item">
                        <div class="quick-item-icon blue"><i class="fas fa-coins"></i></div>
                        <div class="quick-item-content">
                            <div class="quick-item-title">成本管理</div>
                            <div class="quick-item-desc">商品成本录入与核算</div>
                        </div>
                    </a>
                    <a href="expense_allocation.asp" class="quick-item">
                        <div class="quick-item-icon green"><i class="fas fa-calculator"></i></div>
                        <div class="quick-item-content">
                            <div class="quick-item-title">费用分摊</div>
                            <div class="quick-item-desc">间接费用分摊计算</div>
                        </div>
                    </a>
                    <a href="reconciliation.asp" class="quick-item">
                        <div class="quick-item-icon purple"><i class="fas fa-balance-scale"></i></div>
                        <div class="quick-item-content">
                            <div class="quick-item-title">对账中心</div>
                            <div class="quick-item-desc">收支对账与差异处理</div>
                        </div>
                    </a>
                </div>
                
                <!-- 第二列：资金台账 -->
                <div class="quick-column">
                    <div class="column-header"><i class="fas fa-wallet" style="color:#4CAF50"></i> 资金台账</div>
                    <a href="fund_dashboard.asp" class="quick-item">
                        <div class="quick-item-icon green"><i class="fas fa-wallet"></i></div>
                        <div class="quick-item-content">
                            <div class="quick-item-title">资金看板</div>
                            <div class="quick-item-desc">账户余额与资金流向</div>
                        </div>
                    </a>
                    <a href="transactions.asp" class="quick-item">
                        <div class="quick-item-icon blue"><i class="fas fa-exchange-alt"></i></div>
                        <div class="quick-item-content">
                            <div class="quick-item-title">流水管理</div>
                            <div class="quick-item-desc">资金流水查询与导出</div>
                        </div>
                    </a>
                </div>
                
                <!-- 第三列：报表分析 -->
                <div class="quick-column">
                    <div class="column-header"><i class="fas fa-chart-pie" style="color:#FF9800"></i> 报表分析</div>
                    <a href="profit_report.asp" class="quick-item">
                        <div class="quick-item-icon orange"><i class="fas fa-chart-line"></i></div>
                        <div class="quick-item-content">
                            <div class="quick-item-title">经营利润表</div>
                            <div class="quick-item-desc">月度/季度利润分析</div>
                        </div>
                    </a>
                    <a href="product_analysis.asp" class="quick-item">
                        <div class="quick-item-icon teal"><i class="fas fa-box-open"></i></div>
                        <div class="quick-item-content">
                            <div class="quick-item-title">单品分析</div>
                            <div class="quick-item-desc">单品盈利与成本分析</div>
                        </div>
                    </a>
                    <a href="marketing_stats.asp" class="quick-item">
                        <div class="quick-item-icon purple"><i class="fas fa-bullhorn"></i></div>
                        <div class="quick-item-content">
                            <div class="quick-item-title">营销统计</div>
                            <div class="quick-item-desc">营销投入产出分析</div>
                        </div>
                    </a>
                    <a href="budget_management.asp" class="quick-item">
                        <div class="quick-item-icon red"><i class="fas fa-chart-pie"></i></div>
                        <div class="quick-item-content">
                            <div class="quick-item-title">预算管理</div>
                            <div class="quick-item-desc">预算编制与执行监控</div>
                        </div>
                    </a>
                </div>
            </div>
        </div>
        
        <!-- 底部：异常预警汇总 -->
        <div class="alert-section">
            <div class="section-title"><i class="fas fa-bell"></i> 异常预警汇总</div>
            <div class="alert-grid">
                <div class="alert-item <%= IIf(CLng(costAlertCount) > 0, "warning", "normal") %>">
                    <div class="alert-icon <%= IIf(CLng(costAlertCount) > 0, "warning", "normal") %>"><i class="fas fa-coins"></i></div>
                    <div class="alert-content">
                        <div class="alert-label">成本异动预警</div>
                        <div class="alert-value <%= IIf(CLng(costAlertCount) > 0, "warning", "normal") %>"><%= costAlertCount %></div>
                    </div>
                </div>
                <div class="alert-item <%= IIf(CLng(reconAlertCount) > 0, "warning", "normal") %>">
                    <div class="alert-icon <%= IIf(CLng(reconAlertCount) > 0, "warning", "normal") %>"><i class="fas fa-balance-scale"></i></div>
                    <div class="alert-content">
                        <div class="alert-label">对账异常</div>
                        <div class="alert-value <%= IIf(CLng(reconAlertCount) > 0, "warning", "normal") %>"><%= reconAlertCount %></div>
                    </div>
                </div>
                <div class="alert-item <%= IIf(CLng(budgetAlertCount) > 0, "warning", "normal") %>">
                    <div class="alert-icon <%= IIf(CLng(budgetAlertCount) > 0, "warning", "normal") %>"><i class="fas fa-chart-pie"></i></div>
                    <div class="alert-content">
                        <div class="alert-label">预算超支预警</div>
                        <div class="alert-value <%= IIf(CLng(budgetAlertCount) > 0, "warning", "normal") %>"><%= budgetAlertCount %></div>
                    </div>
                </div>
                <div class="alert-item <%= IIf(CLng(fundAlertCount) > 0, "warning", "normal") %>">
                    <div class="alert-icon <%= IIf(CLng(fundAlertCount) > 0, "warning", "normal") %>"><i class="fas fa-wallet"></i></div>
                    <div class="alert-content">
                        <div class="alert-label">资金预警</div>
                        <div class="alert-value <%= IIf(CLng(fundAlertCount) > 0, "warning", "normal") %>"><%= fundAlertCount %></div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
