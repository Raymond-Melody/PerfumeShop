<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
' 安全数值转换函数
Function SafeNum(val)
    If IsNull(val) Or IsEmpty(val) Or val = "" Then
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

' 安全除法函数
Function SafeDiv(numerator, denominator)
    Dim n, d
    n = SafeNum(numerator)
    d = SafeNum(denominator)
    If d = 0 Then
        SafeDiv = 0
    Else
        SafeDiv = n / d
    End If
End Function

' 安全格式化函数
Function SafeFormat(val, decimals)
    SafeFormat = FormatNumber(SafeNum(val), decimals)
End Function

Call OpenConnection()

' ============================================
' 权限验证：FIN_MANAGER 和 FIN_STAFF 均可查看
' ============================================
Dim roleCode
roleCode = Session("AdminRoleCode")
If roleCode <> "FIN_MANAGER" And roleCode <> "FIN_STAFF" And roleCode <> "SUPER_ADMIN" Then
    Response.Redirect "/admin/unauthorized.asp?module=finance"
    Response.End
End If

' ============================================
' 时间维度处理
' ============================================
Dim timeRange, startDate, endDate
Dim viewMode, groupBy
viewMode = LCase(Trim(Request.QueryString("viewMode")))
If viewMode = "" Then viewMode = "month"

Select Case viewMode
    Case "day"
        groupBy = "CAST(OrderDate AS DATE)"
        If startDate = "" Then startDate = SafeFormatDateTime(DateAdd("d", -30, Date()), 2)
        If endDate = "" Then endDate = SafeFormatDateTime(Date(), 2)
    Case "week"
        groupBy = "Year(OrderDate) & '-' & DatePart('ww', OrderDate)"
        If startDate = "" Then startDate = SafeFormatDateTime(DateAdd("ww", -12, Date()), 2)
        If endDate = "" Then endDate = SafeFormatDateTime(Date(), 2)
    Case Else ' month
        viewMode = "month"
        groupBy = "Year(OrderDate) & '-' & Month(OrderDate)"
        If startDate = "" Then startDate = SafeFormatDateTime(DateSerial(Year(Date()), Month(Date()), 1), 2)
        If endDate = "" Then endDate = SafeFormatDateTime(Date(), 2)
End Select

' 获取查询参数
startDate = Trim(Request.QueryString("startDate"))
endDate = Trim(Request.QueryString("endDate"))

' 默认日期范围
If startDate = "" Then startDate = SafeFormatDateTime(DateSerial(Year(Date()), Month(Date()), 1), 2)
If endDate = "" Then endDate = SafeFormatDateTime(Date(), 2)

Dim safeStart, safeEnd
safeStart = SafeSQL(startDate)
safeEnd = SafeSQL(endDate)

' ============================================
' 核心利润表数据查询
' ============================================

' 1. 订单金额 - SUM(Orders.TotalAmount) WHERE Status IN ('Paid','Processing','Shipped','Completed')
Dim orderAmount
orderAmount = GetScalar("SELECT CAST(ISNULL(SUM(TotalAmount), 0) AS FLOAT) FROM Orders WHERE OrderDate >= '" & safeStart & "' AND OrderDate <= '" & safeEnd & "' AND Status IN ('Paid','Processing','Shipped','Completed')")
If IsNull(orderAmount) Or orderAmount = "" Then orderAmount = 0

' 2. 退货金额 - SUM(RefundRecords.RefundAmount) WHERE Status='Completed'
Dim refundAmount
refundAmount = GetScalar("SELECT ISNULL(SUM(RefundAmount), 0) FROM RefundRecords WHERE CompletedAt >= '" & safeStart & "' AND CompletedAt <= '" & safeEnd & "' AND Status='Completed'")
If IsNull(refundAmount) Or refundAmount = "" Then refundAmount = 0

' 3. 净销售额
Dim netSales
netSales = SafeNum(orderAmount) - SafeNum(refundAmount)

' 4. 商品成本 - SUM(Orders.CostAmount)
Dim productCost
productCost = GetScalar("SELECT CAST(ISNULL(SUM(CostAmount), 0) AS FLOAT) FROM Orders WHERE OrderDate >= '" & safeStart & "' AND OrderDate <= '" & safeEnd & "' AND Status IN ('Paid','Processing','Shipped','Completed')")
If IsNull(productCost) Or productCost = "" Then productCost = 0

' 5. 运费成本 - SUM(ExpenseRecords.Amount) WHERE ExpenseType='Shipping'
Dim shippingCost
shippingCost = GetScalar("SELECT ISNULL(SUM(Amount), 0) FROM ExpenseRecords WHERE CreatedAt >= '" & safeStart & "' AND CreatedAt <= '" & safeEnd & "' AND ExpenseType='Shipping'")
If IsNull(shippingCost) Or shippingCost = "" Then shippingCost = 0

' 6. 平台扣点 - SUM(ExpenseRecords.Amount) WHERE ExpenseType='PlatformFee'
Dim platformFee
platformFee = GetScalar("SELECT ISNULL(SUM(Amount), 0) FROM ExpenseRecords WHERE CreatedAt >= '" & safeStart & "' AND CreatedAt <= '" & safeEnd & "' AND ExpenseType='PlatformFee'")
If IsNull(platformFee) Or platformFee = "" Then platformFee = 0

' 7. 推广费用 - SUM(ExpenseRecords.Amount) WHERE ExpenseType='Promotion'
Dim promotionCost
promotionCost = GetScalar("SELECT ISNULL(SUM(Amount), 0) FROM ExpenseRecords WHERE CreatedAt >= '" & safeStart & "' AND CreatedAt <= '" & safeEnd & "' AND ExpenseType='Promotion'")
If IsNull(promotionCost) Or promotionCost = "" Then promotionCost = 0

' 8. 其他费用
Dim otherCost
otherCost = GetScalar("SELECT ISNULL(SUM(Amount), 0) FROM ExpenseRecords WHERE CreatedAt >= '" & safeStart & "' AND CreatedAt <= '" & safeEnd & "' AND ExpenseType NOT IN ('Shipping','PlatformFee','Promotion')")
If IsNull(otherCost) Or otherCost = "" Then otherCost = 0

' 9. 毛利润 = 净销售额 - 商品成本
Dim grossProfit
grossProfit = netSales - productCost

' 10. 边际贡献 = 净销售额 - 商品成本 - 运费 - 平台扣点 - 推广费
Dim contributionMargin
contributionMargin = netSales - productCost - shippingCost - platformFee - promotionCost

' 11. 毛利率 = 毛利润 / 净销售额 x 100%
Dim grossProfitRate
If netSales > 0 Then grossProfitRate = (grossProfit / netSales) * 100 Else grossProfitRate = 0

' 12. 费用率 = 推广费 / 净销售额 x 100%
Dim expenseRate
If netSales > 0 Then expenseRate = (promotionCost / netSales) * 100 Else expenseRate = 0

' ============================================
' 月度利润趋势数据（最近6个月）
' ============================================
Dim trendData(5, 4) ' 6个月，每月4个值：年月、净销售额、毛利润、边际贡献
Dim trendMonths(5)
Dim i, j
' 循环内使用的变量声明移到循环外（VBScript限制）
Dim targetMonth, targetYear
Dim monthStart, monthEnd
Dim safeMonthStart, safeMonthEnd
Dim monthNetSales, monthRefund
Dim monthProductCost, monthGrossProfit
Dim monthShipping, monthPlatform, monthPromotion, monthContribution

For i = 0 To 5
    targetMonth = Month(DateAdd("m", -i, Date()))
    targetYear = Year(DateAdd("m", -i, Date()))
    trendMonths(i) = targetYear & "-" & targetMonth
    
    monthStart = DateSerial(targetYear, targetMonth, 1)
    monthEnd = DateSerial(targetYear, targetMonth + 1, 0)
    
    safeMonthStart = SafeSQL(SafeFormatDateTime(monthStart, 2))
    safeMonthEnd = SafeSQL(SafeFormatDateTime(monthEnd, 2))
    
    ' 净销售额
    monthNetSales = GetScalar("SELECT CAST(ISNULL(SUM(TotalAmount), 0) AS FLOAT) FROM Orders WHERE OrderDate >= '" & safeMonthStart & "' AND OrderDate <= '" & safeMonthEnd & "' AND Status IN ('Paid','Processing','Shipped','Completed')")
    If IsNull(monthNetSales) Then monthNetSales = 0
    
    monthRefund = GetScalar("SELECT ISNULL(SUM(RefundAmount), 0) FROM RefundRecords WHERE CompletedAt >= '" & safeMonthStart & "' AND CompletedAt <= '" & safeMonthEnd & "' AND Status='Completed'")
    If IsNull(monthRefund) Then monthRefund = 0
    
    monthNetSales = SafeNum(monthNetSales) - SafeNum(monthRefund)
    
    ' 毛利润
    monthProductCost = GetScalar("SELECT CAST(ISNULL(SUM(CostAmount), 0) AS FLOAT) FROM Orders WHERE OrderDate >= '" & safeMonthStart & "' AND OrderDate <= '" & safeMonthEnd & "' AND Status IN ('Paid','Processing','Shipped','Completed')")
    If IsNull(monthProductCost) Then monthProductCost = 0
    
    monthGrossProfit = monthNetSales - SafeNum(monthProductCost)
    
    ' 边际贡献
    monthShipping = GetScalar("SELECT ISNULL(SUM(Amount), 0) FROM ExpenseRecords WHERE CreatedAt >= '" & safeMonthStart & "' AND CreatedAt <= '" & safeMonthEnd & "' AND ExpenseType='Shipping'")
    monthPlatform = GetScalar("SELECT ISNULL(SUM(Amount), 0) FROM ExpenseRecords WHERE CreatedAt >= '" & safeMonthStart & "' AND CreatedAt <= '" & safeMonthEnd & "' AND ExpenseType='PlatformFee'")
    monthPromotion = GetScalar("SELECT ISNULL(SUM(Amount), 0) FROM ExpenseRecords WHERE CreatedAt >= '" & safeMonthStart & "' AND CreatedAt <= '" & safeMonthEnd & "' AND ExpenseType='Promotion'")
    If IsNull(monthShipping) Then monthShipping = 0
    If IsNull(monthPlatform) Then monthPlatform = 0
    If IsNull(monthPromotion) Then monthPromotion = 0
    monthContribution = monthNetSales - SafeNum(monthProductCost) - SafeNum(monthShipping) - SafeNum(monthPlatform) - SafeNum(monthPromotion)
    
    trendData(i, 0) = targetYear & "年" & targetMonth & "月"
    trendData(i, 1) = monthNetSales
    trendData(i, 2) = monthGrossProfit
    trendData(i, 3) = monthContribution
Next

' ============================================
' 上月同比数据
' ============================================
Dim lastMonthStart, lastMonthEnd
lastMonthStart = DateSerial(Year(Date()), Month(Date()) - 1, 1)
lastMonthEnd = DateSerial(Year(Date()), Month(Date()), 0)

Dim safeLastStart, safeLastEnd
safeLastStart = SafeSQL(SafeFormatDateTime(lastMonthStart, 2))
safeLastEnd = SafeSQL(SafeFormatDateTime(lastMonthEnd, 2))

' 上月订单金额
Dim lastOrderAmount
lastOrderAmount = GetScalar("SELECT CAST(ISNULL(SUM(TotalAmount), 0) AS FLOAT) FROM Orders WHERE OrderDate >= '" & safeLastStart & "' AND OrderDate <= '" & safeLastEnd & "' AND Status IN ('Paid','Processing','Shipped','Completed')")
If IsNull(lastOrderAmount) Then lastOrderAmount = 0

' 上月退货金额
Dim lastRefundAmount
lastRefundAmount = GetScalar("SELECT ISNULL(SUM(RefundAmount), 0) FROM RefundRecords WHERE CompletedAt >= '" & safeLastStart & "' AND CompletedAt <= '" & safeLastEnd & "' AND Status='Completed'")
If IsNull(lastRefundAmount) Then lastRefundAmount = 0

' 上月净销售额
Dim lastNetSales
lastNetSales = SafeNum(lastOrderAmount) - SafeNum(lastRefundAmount)

' 上月商品成本
Dim lastProductCost
lastProductCost = GetScalar("SELECT CAST(ISNULL(SUM(CostAmount), 0) AS FLOAT) FROM Orders WHERE OrderDate >= '" & safeLastStart & "' AND OrderDate <= '" & safeLastEnd & "' AND Status IN ('Paid','Processing','Shipped','Completed')")
If IsNull(lastProductCost) Then lastProductCost = 0

' 上月毛利润
Dim lastGrossProfit
lastGrossProfit = lastNetSales - SafeNum(lastProductCost)

' 上月推广费用
Dim lastPromotionCost
lastPromotionCost = GetScalar("SELECT ISNULL(SUM(Amount), 0) FROM ExpenseRecords WHERE CreatedAt >= '" & safeLastStart & "' AND CreatedAt <= '" & safeLastEnd & "' AND ExpenseType='Promotion'")
If IsNull(lastPromotionCost) Then lastPromotionCost = 0

' 上月费用率
Dim lastExpenseRate
If lastNetSales > 0 Then lastExpenseRate = (SafeNum(lastPromotionCost) / lastNetSales) * 100 Else lastExpenseRate = 0

' 计算增长率
Function CalcGrowth(current, last)
    If SafeNum(last) = 0 Then
        If SafeNum(current) > 0 Then CalcGrowth = 100 Else CalcGrowth = 0
    Else
        CalcGrowth = ((SafeNum(current) - SafeNum(last)) / SafeNum(last)) * 100
    End If
End Function

Dim netSalesGrowth, grossProfitGrowth, expenseRateChange
netSalesGrowth = CalcGrowth(netSales, lastNetSales)
grossProfitGrowth = CalcGrowth(grossProfit, lastGrossProfit)
expenseRateChange = expenseRate - lastExpenseRate

' 记录日志
Call LogAdminAction("查看经营利润表", "finance", "", "", safeStart & " 至 " & safeEnd)
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>经营利润表 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        /* 暗色主题 */
        .main-content { background: #1a1a2e; min-height: 100vh; color: #e0e0e0; }
        .page-title { color: #fff; }
        
        /* 筛选栏 */
        .filter-bar { 
            background: #2d2d44; 
            padding: 20px; 
            border-radius: 12px; 
            margin-bottom: 25px; 
            box-shadow: 0 4px 15px rgba(0,0,0,0.3); 
        }
        .filter-bar form { display: flex; gap: 15px; align-items: center; flex-wrap: wrap; }
        .filter-bar label { color: #a0aec0; }
        .filter-bar input, .filter-bar select { 
            padding: 10px 15px; 
            border: 1px solid rgba(255,255,255,0.15); 
            border-radius: 8px; 
            background: #1a1a2e; 
            color: #e0e0e0;
        }
        .filter-bar input:focus, .filter-bar select:focus { 
            border-color: #00bcd4; 
            outline: none; 
        }
        .view-mode-btn {
            padding: 8px 16px;
            border: 1px solid #3d445c;
            background: #1a1a2e;
            color: #a0aec0;
            border-radius: 6px;
            cursor: pointer;
            transition: all 0.3s;
        }
        .view-mode-btn.active, .view-mode-btn:hover {
            background: #00bcd4;
            color: #fff;
            border-color: #00bcd4;
        }
        
        /* 报表区块 */
        .report-section { 
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); 
            border-radius: 12px; 
            padding: 25px; 
            box-shadow: 0 4px 15px rgba(0,0,0,0.3); 
            margin-bottom: 25px; 
        }
        .report-section h3 { 
            font-size: 18px; 
            color: #fff; 
            margin-bottom: 20px; 
            display: flex; 
            align-items: center; 
            gap: 10px; 
            padding-bottom: 15px; 
            border-bottom: 1px solid #3d445c; 
        }
        
        /* 利润表表格 */
        .profit-table { width: 100%; border-collapse: collapse; }
        .profit-table th { 
            background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); 
            color: white; 
            padding: 15px; 
            text-align: left; 
            font-weight: 600;
        }
        .profit-table td { 
            padding: 15px; 
            border-bottom: 1px solid #3d445c; 
            color: #e0e0e0;
        }
        .profit-table tr:hover { background: #2d3347; }
        .profit-table .highlight-row { 
            background: rgba(102, 126, 234, 0.15) !important; 
            font-weight: bold;
        }
        .profit-table .total-row { 
            background: rgba(76, 175, 80, 0.15) !important; 
            font-weight: bold;
            font-size: 1.05em;
        }
        .profit-table .negative { color: #ff6b6b; }
        .profit-table .positive { color: #4CAF50; }
        .profit-table .amount { text-align: right; font-family: 'Consolas', monospace; }
        .profit-table .rate { text-align: right; color: #a0aec0; }
        
        /* 红绿灯指示器 */
        .traffic-light {
            display: inline-flex;
            align-items: center;
            gap: 8px;
        }
        .light {
            width: 12px;
            height: 12px;
            border-radius: 50%;
            box-shadow: 0 0 8px currentColor;
        }
        .light.green { background: #4CAF50; color: #4CAF50; }
        .light.yellow { background: #FFC107; color: #FFC107; }
        .light.red { background: #F44336; color: #F44336; }
        
        /* 图表网格 */
        .charts-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 25px; margin-bottom: 25px; }
        @media (max-width: 1024px) { .charts-grid { grid-template-columns: 1fr; } }
        
        /* 柱状图样式 */
        .bar-chart { display: flex; align-items: flex-end; justify-content: space-around; height: 250px; padding: 20px 0; }
        .bar-group { display: flex; flex-direction: column; align-items: center; gap: 8px; flex: 1; }
        .bar-stack { display: flex; align-items: flex-end; gap: 4px; height: 200px; }
        .bar {
            width: 24px;
            border-radius: 4px 4px 0 0;
            transition: all 0.3s;
            position: relative;
        }
        .bar:hover { opacity: 0.8; }
        .bar.net-sales { background: linear-gradient(to top, #00bcd4, #00838f); }
        .bar.gross-profit { background: linear-gradient(to top, #4CAF50, #388e3c); }
        .bar.contribution { background: linear-gradient(to top, #ff9800, #f57c00); }
        .bar-label { font-size: 11px; color: #a0aec0; text-align: center; }
        .bar-value {
            position: absolute;
            top: -20px;
            left: 50%;
            transform: translateX(-50%);
            font-size: 10px;
            color: #fff;
            white-space: nowrap;
            opacity: 0;
            transition: opacity 0.3s;
        }
        .bar:hover .bar-value { opacity: 1; }
        
        /* 饼图样式 */
        .pie-chart-container { display: flex; align-items: center; justify-content: center; gap: 40px; padding: 20px; }
        .pie-chart {
            width: 180px;
            height: 180px;
            border-radius: 50%;
            background: conic-gradient(
                #00bcd4 0deg var(--p1),
                #4CAF50 var(--p1) var(--p2),
                #ff9800 var(--p2) var(--p3),
                #e91e63 var(--p3) var(--p4),
                #9c27b0 var(--p4) 360deg
            );
            box-shadow: 0 4px 15px rgba(0,0,0,0.3);
        }
        .pie-legend { display: flex; flex-direction: column; gap: 10px; }
        .legend-item { display: flex; align-items: center; gap: 10px; font-size: 13px; }
        .legend-color { width: 16px; height: 16px; border-radius: 4px; }
        .legend-label { color: #a0aec0; }
        .legend-value { color: #fff; font-weight: bold; margin-left: auto; }
        
        /* 对比分析 */
        .compare-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; }
        @media (max-width: 768px) { .compare-grid { grid-template-columns: 1fr; } }
        .compare-card {
            background: #1a1a2e;
            border-radius: 10px;
            padding: 20px;
            text-align: center;
            border: 1px solid #3d445c;
        }
        .compare-card h4 { color: #a0aec0; font-size: 14px; margin-bottom: 10px; }
        .compare-card .current-value { font-size: 24px; font-weight: bold; color: #fff; margin-bottom: 8px; }
        .compare-card .change {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            font-size: 14px;
            padding: 4px 12px;
            border-radius: 20px;
        }
        .compare-card .change.up { background: rgba(76, 175, 80, 0.2); color: #4CAF50; }
        .compare-card .change.down { background: rgba(244, 67, 54, 0.2); color: #F44336; }
        
        /* 费用率柱状图 */
        .rate-chart { padding: 20px 0; }
        .rate-bar-item { display: flex; align-items: center; margin-bottom: 15px; }
        .rate-bar-label { width: 80px; color: #a0aec0; font-size: 12px; }
        .rate-bar-wrap { flex: 1; height: 24px; background: #1a1a2e; border-radius: 12px; overflow: hidden; margin: 0 15px; }
        .rate-bar-fill { height: 100%; border-radius: 12px; transition: width 0.5s ease; }
        .rate-bar-fill.green { background: linear-gradient(90deg, #4CAF50, #388e3c); }
        .rate-bar-fill.yellow { background: linear-gradient(90deg, #FFC107, #ff9800); }
        .rate-bar-fill.red { background: linear-gradient(90deg, #F44336, #e91e63); }
        .rate-bar-value { width: 60px; text-align: right; color: #fff; font-family: 'Consolas', monospace; }
        
        .no-data { text-align: center; padding: 40px; color: #666; }
        .no-data i { font-size: 48px; margin-bottom: 15px; display: block; color: #444; }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-chart-line"></i> 经营利润表</h2>
            <div class="breadcrumb">
                <a href="index.asp">财务中心</a> / <span>经营利润表</span>
            </div>
        </div>
        
        <!-- 时间维度选择器 -->
        <div class="filter-bar">
            <form method="get" action="profit_report.asp">
                <label>视图模式:</label>
                <a href="?viewMode=day&startDate=<%= Server.HTMLEncode(SafeFormatDateTime(DateAdd("d", -30, Date()), 2)) %>&endDate=<%= Server.HTMLEncode(SafeFormatDateTime(Date(), 2)) %>" 
                   class="view-mode-btn <%= IIf(viewMode="day", "active", "") %>">日报</a>
                <a href="?viewMode=week&startDate=<%= Server.HTMLEncode(SafeFormatDateTime(DateAdd("ww", -12, Date()), 2)) %>&endDate=<%= Server.HTMLEncode(SafeFormatDateTime(Date(), 2)) %>" 
                   class="view-mode-btn <%= IIf(viewMode="week", "active", "") %>">周报</a>
                <a href="?viewMode=month&startDate=<%= Server.HTMLEncode(SafeFormatDateTime(DateSerial(Year(Date()), Month(Date()), 1), 2)) %>&endDate=<%= Server.HTMLEncode(SafeFormatDateTime(Date(), 2)) %>" 
                   class="view-mode-btn <%= IIf(viewMode="month", "active", "") %>">月报</a>
                
                <label style="margin-left: 20px;">开始日期:</label>
                <input type="date" name="startDate" value="<%= Server.HTMLEncode(startDate) %>">
                <label>结束日期:</label>
                <input type="date" name="endDate" value="<%= Server.HTMLEncode(endDate) %>">
                <input type="hidden" name="viewMode" value="<%= Server.HTMLEncode(viewMode) %>">
                <button type="submit" class="admin-btn admin-btn-primary"><i class="fas fa-search"></i> 查询</button>
                <a href="profit_report.asp" class="admin-btn admin-btn-secondary"><i class="fas fa-redo"></i> 重置</a>
            </form>
        </div>
        
        <!-- 核心利润表 -->
        <div class="report-section">
            <h3><i class="fas fa-table" style="color: #00bcd4;"></i> 核心利润表 (<%= Server.HTMLEncode(startDate) %> 至 <%= Server.HTMLEncode(endDate) %>)</h3>
            <table class="profit-table">
                <thead>
                    <tr>
                        <th>指标项目</th>
                        <th class="amount">金额 (¥)</th>
                        <th class="rate">占比/比率</th>
                        <th>状态</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td><i class="fas fa-shopping-cart" style="color: #00bcd4;"></i> 订单金额</td>
                        <td class="amount positive"><%= FormatNumber(SafeNum(orderAmount), 2) %></td>
                        <td class="rate">-</td>
                        <td>-</td>
                    </tr>
                    <tr>
                        <td><i class="fas fa-undo" style="color: #ff6b6b;"></i> 退货金额</td>
                        <td class="amount negative">-<%= FormatNumber(SafeNum(refundAmount), 2) %></td>
                        <td class="rate"><%= SafeFormat(SafeDiv(refundAmount, orderAmount) * 100, 1) %>%</td>
                        <td>-</td>
                    </tr>
                    <tr class="highlight-row">
                        <td><i class="fas fa-calculator" style="color: #4CAF50;"></i> 净销售额</td>
                        <td class="amount positive"><%= FormatNumber(SafeNum(netSales), 2) %></td>
                        <td class="rate">100%</td>
                        <td>-</td>
                    </tr>
                    <tr>
                        <td colspan="4" style="height: 10px; background: transparent;"></td>
                    </tr>
                    <tr>
                        <td><i class="fas fa-box" style="color: #9c27b0;"></i> 商品成本</td>
                        <td class="amount negative">-<%= FormatNumber(SafeNum(productCost), 2) %></td>
                        <td class="rate"><%= SafeFormat(SafeDiv(productCost, netSales) * 100, 1) %>%</td>
                        <td>-</td>
                    </tr>
                    <tr class="highlight-row">
                        <td><i class="fas fa-coins" style="color: #ff9800;"></i> 毛利润</td>
                        <td class="amount <%= IIf(SafeNum(grossProfit)>=0, "positive", "negative") %>"><%= IIf(SafeNum(grossProfit)>=0, "", "-") %><%= FormatNumber(Abs(SafeNum(grossProfit)), 2) %></td>
                        <td class="rate"><%= FormatNumber(SafeNum(grossProfitRate), 1) %>%</td>
                        <td>
                            <div class="traffic-light">
                                <span class="light <%= IIf(SafeNum(grossProfitRate)>=30, "green", IIf(SafeNum(grossProfitRate)>=15, "yellow", "red")) %>"></span>
                                <%= IIf(SafeNum(grossProfitRate)>=30, "优秀", IIf(SafeNum(grossProfitRate)>=15, "良好", "需关注")) %>
                            </div>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="4" style="height: 10px; background: transparent;"></td>
                    </tr>
                    <tr>
                        <td><i class="fas fa-truck" style="color: #2196F3;"></i> 运费成本</td>
                        <td class="amount negative">-<%= FormatNumber(SafeNum(shippingCost), 2) %></td>
                        <td class="rate"><%= SafeFormat(SafeDiv(shippingCost, netSales) * 100, 1) %>%</td>
                        <td>-</td>
                    </tr>
                    <tr>
                        <td><i class="fas fa-percentage" style="color: #e91e63;"></i> 平台扣点</td>
                        <td class="amount negative">-<%= FormatNumber(SafeNum(platformFee), 2) %></td>
                        <td class="rate"><%= SafeFormat(SafeDiv(platformFee, netSales) * 100, 1) %>%</td>
                        <td>-</td>
                    </tr>
                    <tr>
                        <td><i class="fas fa-bullhorn" style="color: #ff5722;"></i> 推广费用</td>
                        <td class="amount negative">-<%= FormatNumber(SafeNum(promotionCost), 2) %></td>
                        <td class="rate"><%= FormatNumber(SafeNum(expenseRate), 1) %>%</td>
                        <td>
                            <div class="traffic-light">
                                <span class="light <%= IIf(SafeNum(expenseRate)<10, "green", IIf(SafeNum(expenseRate)<20, "yellow", "red")) %>"></span>
                                <%= IIf(SafeNum(expenseRate)<10, "健康", IIf(SafeNum(expenseRate)<20, "警告", "过高")) %>
                            </div>
                        </td>
                    </tr>
                    <tr>
                        <td><i class="fas fa-ellipsis-h" style="color: #607d8b;"></i> 其他费用</td>
                        <td class="amount negative">-<%= FormatNumber(SafeNum(otherCost), 2) %></td>
                        <td class="rate"><%= SafeFormat(SafeDiv(otherCost, netSales) * 100, 1) %>%</td>
                        <td>-</td>
                    </tr>
                    <tr class="total-row">
                        <td><i class="fas fa-chart-pie" style="color: #4CAF50;"></i> 边际贡献</td>
                        <td class="amount <%= IIf(SafeNum(contributionMargin)>=0, "positive", "negative") %>"><%= IIf(SafeNum(contributionMargin)>=0, "", "-") %><%= FormatNumber(Abs(SafeNum(contributionMargin)), 2) %></td>
                        <td class="rate"><%= SafeFormat(SafeDiv(contributionMargin, netSales) * 100, 1) %>%</td>
                        <td>
                            <div class="traffic-light">
                                <span class="light <%= IIf(SafeNum(contributionMargin)>0, "green", "red") %>"></span>
                                <%= IIf(SafeNum(contributionMargin)>0, "盈利", "亏损") %>
                            </div>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>
        
        <!-- 月度利润趋势图 -->
        <div class="report-section">
            <h3><i class="fas fa-chart-bar" style="color: #4CAF50;"></i> 月度利润趋势（最近6个月）</h3>
            <div class="bar-chart">
                <% 
                Dim maxVal
                maxVal = 0
                For i = 0 To 5
                    If SafeNum(trendData(i, 1)) > maxVal Then maxVal = SafeNum(trendData(i, 1))
                    If SafeNum(trendData(i, 2)) > maxVal Then maxVal = SafeNum(trendData(i, 2))
                    If SafeNum(trendData(i, 3)) > maxVal Then maxVal = SafeNum(trendData(i, 3))
                Next
                If maxVal = 0 Then maxVal = 1
                
                Dim h1, h2, h3
                For i = 5 To 0 Step -1
                    h1 = (SafeNum(trendData(i, 1)) / maxVal) * 180
                    h2 = (SafeNum(trendData(i, 2)) / maxVal) * 180
                    h3 = (SafeNum(trendData(i, 3)) / maxVal) * 180
                    If h1 < 2 Then h1 = 2
                    If h2 < 2 Then h2 = 2
                    If h3 < 2 Then h3 = 2
                %>
                <div class="bar-group">
                    <div class="bar-stack">
                        <div class="bar net-sales" style="height: <%= h1 %>px;">
                            <span class="bar-value">¥<%= FormatNumber(SafeNum(trendData(i, 1))/1000, 0) %>k</span>
                        </div>
                        <div class="bar gross-profit" style="height: <%= h2 %>px;">
                            <span class="bar-value">¥<%= FormatNumber(SafeNum(trendData(i, 2))/1000, 0) %>k</span>
                        </div>
                        <div class="bar contribution" style="height: <%= h3 %>px;">
                            <span class="bar-value">¥<%= FormatNumber(SafeNum(trendData(i, 3))/1000, 0) %>k</span>
                        </div>
                    </div>
                    <div class="bar-label"><%= trendData(i, 0) %></div>
                </div>
                <% Next %>
            </div>
            <div style="display: flex; justify-content: center; gap: 30px; margin-top: 15px;">
                <div class="legend-item">
                    <div class="legend-color" style="background: linear-gradient(to right, #00bcd4, #00838f);"></div>
                    <span class="legend-label">净销售额</span>
                </div>
                <div class="legend-item">
                    <div class="legend-color" style="background: linear-gradient(to right, #4CAF50, #388e3c);"></div>
                    <span class="legend-label">毛利润</span>
                </div>
                <div class="legend-item">
                    <div class="legend-color" style="background: linear-gradient(to right, #ff9800, #f57c00);"></div>
                    <span class="legend-label">边际贡献</span>
                </div>
            </div>
        </div>
        
        <div class="charts-grid">
            <!-- 利润构成分析（饼图） -->
            <div class="report-section" style="margin-bottom: 0;">
                <h3><i class="fas fa-chart-pie" style="color: #ff9800;"></i> 成本构成分析</h3>
                <% 
                Dim totalCost, p1, p2, p3, p4, p1deg, p2deg, p3deg, p4deg
                totalCost = SafeNum(productCost) + SafeNum(shippingCost) + SafeNum(platformFee) + SafeNum(promotionCost) + SafeNum(otherCost)
                If totalCost = 0 Then totalCost = 1
                
                p1 = (SafeNum(productCost) / totalCost) * 100
                p2 = (SafeNum(shippingCost) / totalCost) * 100
                p3 = (SafeNum(platformFee) / totalCost) * 100
                p4 = (SafeNum(promotionCost) / totalCost) * 100
                
                p1deg = (p1 / 100) * 360
                p2deg = ((p1 + p2) / 100) * 360
                p3deg = ((p1 + p2 + p3) / 100) * 360
                p4deg = ((p1 + p2 + p3 + p4) / 100) * 360
                %>
                <div class="pie-chart-container">
                    <div class="pie-chart" style="--p1: <%= p1deg %>deg; --p2: <%= p2deg %>deg; --p3: <%= p3deg %>deg; --p4: <%= p4deg %>deg;"></div>
                    <div class="pie-legend">
                        <div class="legend-item">
                            <div class="legend-color" style="background: #00bcd4;"></div>
                            <span class="legend-label">商品成本</span>
                            <span class="legend-value"><%= FormatNumber(p1, 1) %>%</span>
                        </div>
                        <div class="legend-item">
                            <div class="legend-color" style="background: #4CAF50;"></div>
                            <span class="legend-label">运费</span>
                            <span class="legend-value"><%= FormatNumber(p2, 1) %>%</span>
                        </div>
                        <div class="legend-item">
                            <div class="legend-color" style="background: #ff9800;"></div>
                            <span class="legend-label">平台扣点</span>
                            <span class="legend-value"><%= FormatNumber(p3, 1) %>%</span>
                        </div>
                        <div class="legend-item">
                            <div class="legend-color" style="background: #e91e63;"></div>
                            <span class="legend-label">推广费</span>
                            <span class="legend-value"><%= FormatNumber(p4, 1) %>%</span>
                        </div>
                        <div class="legend-item">
                            <div class="legend-color" style="background: #9c27b0;"></div>
                            <span class="legend-label">其他</span>
                            <span class="legend-value"><%= FormatNumber(100 - p1 - p2 - p3 - p4, 1) %>%</span>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- 费用率趋势 -->
            <div class="report-section" style="margin-bottom: 0;">
                <h3><i class="fas fa-percent" style="color: #e91e63;"></i> 各月费用率变化</h3>
                <div class="rate-chart">
                    <% 
                    Dim monthRate
                    For i = 5 To 0 Step -1
                        If CDbl(trendData(i, 1)) > 0 Then
                            ' 估算费用率（基于推广费用占比）
                            If SafeNum(trendData(i, 1)) <> 0 Then monthRate = ((SafeNum(trendData(i, 1)) - SafeNum(trendData(i, 3))) / SafeNum(trendData(i, 1))) * 50 Else monthRate = 0
                        Else
                            monthRate = 0
                        End If
                        If monthRate > 100 Then monthRate = 100
                        If monthRate < 0 Then monthRate = 0
                    %>
                    <div class="rate-bar-item">
                        <div class="rate-bar-label"><%= trendData(i, 0) %></div>
                        <div class="rate-bar-wrap">
                            <div class="rate-bar-fill <%= IIf(monthRate<10, "green", IIf(monthRate<20, "yellow", "red")) %>" style="width: <%= monthRate %>%;"></div>
                        </div>
                        <div class="rate-bar-value"><%= FormatNumber(monthRate, 1) %>%</div>
                    </div>
                    <% Next %>
                </div>
            </div>
        </div>
        
        <!-- 对比分析 -->
        <div class="report-section">
            <h3><i class="fas fa-exchange-alt" style="color: #2196F3;"></i> 本月 vs 上月 对比分析</h3>
            <div class="compare-grid">
                <div class="compare-card">
                    <h4>净销售额</h4>
                    <div class="current-value">¥<%= SafeFormat(netSales/1000, 1) %>k</div>
                    <div class="change <%= IIf(netSalesGrowth>=0, "up", "down") %>">
                        <i class="fas fa-arrow-<%= IIf(netSalesGrowth>=0, "up", "down") %>"></i>
                        <%= FormatNumber(Abs(netSalesGrowth), 1) %>%
                    </div>
                </div>
                <div class="compare-card">
                    <h4>毛利润</h4>
                    <div class="current-value">¥<%= SafeFormat(grossProfit/1000, 1) %>k</div>
                    <div class="change <%= IIf(grossProfitGrowth>=0, "up", "down") %>">
                        <i class="fas fa-arrow-<%= IIf(grossProfitGrowth>=0, "up", "down") %>"></i>
                        <%= FormatNumber(Abs(grossProfitGrowth), 1) %>%
                    </div>
                </div>
                <div class="compare-card">
                    <h4>费用率变化</h4>
                    <div class="current-value"><%= FormatNumber(expenseRate, 1) %>%</div>
                    <div class="change <%= IIf(expenseRateChange<=0, "up", "down") %>">
                        <i class="fas fa-arrow-<%= IIf(expenseRateChange<=0, "up", "down") %>"></i>
                        <%= FormatNumber(Abs(expenseRateChange), 1) %>%
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
