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

Call OpenConnection()

' 权限检查：FIN_MANAGER 和 FIN_STAFF 均可查看
Dim roleCode
roleCode = Session("AdminRoleCode")
If Not (Left(roleCode, 3) = "FIN" Or roleCode = "SUPER_ADMIN") Then
    Response.Redirect "/admin/unauthorized.asp?module=finance"
    Response.End
End If

' 获取当前Tab
Dim currentTab
currentTab = Trim(Request.QueryString("tab"))
If currentTab = "" Then currentTab = "breakeven"

' Tab 1: 获取商品列表用于下拉选择
Dim rsProducts
Set rsProducts = ExecuteQuery("SELECT ProductID, ProductName, BasePrice, BOMCost, UnitCost FROM Products WHERE IsActive = 1 ORDER BY ProductName")

' Tab 1: 获取选中商品信息
Dim selectedProductId, productPrice, productBOMCost, productName
selectedProductId = Trim(Request.QueryString("pid"))
productPrice = 0
productBOMCost = 0
productName = ""

If selectedProductId <> "" And IsNumeric(selectedProductId) Then
    Dim rsSelected
    Set rsSelected = ExecuteQuery("SELECT ProductName, BasePrice, BOMCost FROM Products WHERE ProductID = " & CLng(selectedProductId))
    If Not rsSelected Is Nothing And Not rsSelected.EOF Then
        productName = rsSelected("ProductName")
        productPrice = CDbl("0" & rsSelected("BasePrice"))
        productBOMCost = CDbl("0" & rsSelected("BOMCost"))
        rsSelected.Close
    End If
    Set rsSelected = Nothing
End If

' Tab 1: 获取用户输入参数
Dim categoryRate, conversionRate, fixedRate
categoryRate = Trim(Request.QueryString("crate"))
conversionRate = Trim(Request.QueryString("convr"))
fixedRate = Trim(Request.QueryString("frate"))

If categoryRate = "" Then categoryRate = "5"
If conversionRate = "" Then conversionRate = "3"
If fixedRate = "" Then fixedRate = "8"

Dim catRateVal, convRateVal, fixRateVal
catRateVal = CDbl("0" & categoryRate) / 100
convRateVal = CDbl("0" & conversionRate) / 100
fixRateVal = CDbl("0" & fixedRate) / 100

' Tab 2: 成本异动分析数据
Dim rsCostChange
Dim currentMonth, lastMonth
currentMonth = Year(Date()) & "-" & Right("0" & Month(Date()), 2)
lastMonth = Year(DateAdd("m", -1, Date())) & "-" & Right("0" & Month(DateAdd("m", -1, Date())), 2)

' 使用子查询替代COUNT(DISTINCT) - Access兼容性
Set rsCostChange = ExecuteQuery(_
    "SELECT p.ProductID, p.ProductName, " & _
    "(SELECT SUM(TotalCost) FROM ProductCosts WHERE ProductID = p.ProductID AND Period = '" & lastMonth & "-01" & "') AS LastMonthCost, " & _
    "(SELECT SUM(TotalCost) FROM ProductCosts WHERE ProductID = p.ProductID AND Period = '" & currentMonth & "-01" & "') AS CurrentMonthCost " & _
    "FROM Products p " & _
    "WHERE p.IsActive = 1 " & _
    "ORDER BY p.ProductName")

' Tab 3: TOP10 最盈利商品
Dim rsTopProfit
Set rsTopProfit = ExecuteQuery(_
    "SELECT TOP 10 p.ProductID, p.ProductName, " & _
    "SUM(od.Subtotal) AS TotalRevenue, " & _
    "SUM(od.Quantity * p.UnitCost) AS TotalCost, " & _
    "SUM(od.Subtotal) - SUM(od.Quantity * p.UnitCost) AS Profit " & _
    "FROM (Products p " & _
    "INNER JOIN OrderDetails od ON p.ProductID = od.ProductID) " & _
    "INNER JOIN Orders o ON od.OrderID = o.OrderID " & _
    "WHERE o.Status IN ('Paid', 'Shipped', 'Delivered') " & _
    "GROUP BY p.ProductID, p.ProductName " & _
    "ORDER BY SUM(od.Subtotal) - SUM(od.Quantity * p.UnitCost) DESC")

' Tab 3: TOP10 最亏损商品（利润最低的10个）
Dim rsBottomProfit
Set rsBottomProfit = ExecuteQuery(_
    "SELECT TOP 10 p.ProductID, p.ProductName, " & _
    "SUM(od.Subtotal) AS TotalRevenue, " & _
    "SUM(od.Quantity * p.UnitCost) AS TotalCost, " & _
    "SUM(od.Subtotal) - SUM(od.Quantity * p.UnitCost) AS Profit " & _
    "FROM (Products p " & _
    "INNER JOIN OrderDetails od ON p.ProductID = od.ProductID) " & _
    "INNER JOIN Orders o ON od.OrderID = o.OrderID " & _
    "WHERE o.Status IN ('Paid', 'Shipped', 'Delivered') " & _
    "GROUP BY p.ProductID, p.ProductName " & _
    "ORDER BY SUM(od.Subtotal) - SUM(od.Quantity * p.UnitCost) ASC")

' Tab 4: 商品毛利率分布
Dim rsMarginDist
Set rsMarginDist = ExecuteQuery(_
    "SELECT p.ProductID, p.ProductName, p.BasePrice, p.UnitCost, " & _
    "IIF(p.BasePrice > 0, (p.BasePrice - p.UnitCost) / p.BasePrice * 100, 0) AS MarginRate " & _
    "FROM Products p WHERE p.IsActive = 1 AND p.BasePrice > 0")

' 统计各区间商品数量
Dim cntNeg, cnt0_10, cnt10_20, cnt20_30, cnt30_50, cnt50plus, marginRate
cntNeg = 0 : cnt0_10 = 0 : cnt10_20 = 0 : cnt20_30 = 0 : cnt30_50 = 0 : cnt50plus = 0

If Not rsMarginDist Is Nothing Then
    Do While Not rsMarginDist.EOF
        marginRate = CDbl("0" & rsMarginDist("MarginRate"))
        If marginRate < 0 Then
            cntNeg = cntNeg + 1
        ElseIf marginRate < 10 Then
            cnt0_10 = cnt0_10 + 1
        ElseIf marginRate < 20 Then
            cnt10_20 = cnt10_20 + 1
        ElseIf marginRate < 30 Then
            cnt20_30 = cnt20_30 + 1
        ElseIf marginRate < 50 Then
            cnt30_50 = cnt30_50 + 1
        Else
            cnt50plus = cnt50plus + 1
        End If
        rsMarginDist.MoveNext
    Loop
    rsMarginDist.MoveFirst
End If

Call LogAdminAction("查看单品分析", "finance", "Products", "", "")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>单品分析 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        :root {
            --bg-dark: #1a1a2e;
            --bg-card: #2d2d44;
            --bg-hover: #1e1e32;
            --text-primary: #e0e0e0;
            --text-secondary: #888;
            --accent: #00bcd4;
            --accent-light: #4dd0e1;
            --success: #4CAF50;
            --warning: #ff9800;
            --danger: #f44336;
            --info: #2196F3;
        }
        
        body {
            background: var(--bg-dark);
            color: var(--text-primary);
        }
        
        .main-content {
            background: var(--bg-dark);
        }
        
        .page-title {
            color: var(--text-primary);
        }
        
        .breadcrumb {
            color: var(--text-secondary);
        }
        
        .breadcrumb a {
            color: var(--info);
        }
        
        /* Tab导航 */
        .tab-nav {
            display: flex;
            gap: 5px;
            margin-bottom: 25px;
            background: var(--bg-card);
            padding: 10px;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.3);
        }
        
        .tab-nav a {
            padding: 12px 24px;
            color: var(--text-secondary);
            text-decoration: none;
            border-radius: 8px;
            transition: all 0.3s ease;
            font-weight: 500;
        }
        
        .tab-nav a:hover {
            background: var(--bg-hover);
            color: var(--text-primary);
        }
        
        .tab-nav a.active {
            background: linear-gradient(135deg, var(--accent) 0%, var(--accent-light) 100%);
            color: white;
        }
        
        /* 卡片样式 */
        .card {
            background: var(--bg-card);
            border-radius: 16px;
            padding: 25px;
            margin-bottom: 25px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
        }
        
        .card-title {
            font-size: 18px;
            font-weight: 600;
            margin-bottom: 20px;
            color: var(--text-primary);
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .card-title i {
            color: var(--accent);
        }
        
        /* 表单样式 */
        .form-group {
            margin-bottom: 20px;
        }
        
        .form-group label {
            display: block;
            margin-bottom: 8px;
            color: var(--text-secondary);
            font-size: 14px;
        }
        
        .form-control {
            width: 100%;
            padding: 12px 15px;
            background: var(--bg-dark);
            border: 1px solid var(--bg-hover);
            border-radius: 10px;
            color: var(--text-primary);
            font-size: 14px;
            transition: all 0.3s ease;
        }
        
        .form-control:focus {
            outline: none;
            border-color: var(--accent);
            box-shadow: 0 0 0 3px rgba(233, 69, 96, 0.1);
        }
        
        select.form-control {
            cursor: pointer;
        }
        
        /* 计算器布局 */
        .calculator-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 25px;
        }
        
        .input-section {
            background: var(--bg-dark);
            padding: 20px;
            border-radius: 12px;
        }
        
        .result-section {
            background: linear-gradient(135deg, var(--bg-hover) 0%, var(--bg-card) 100%);
            padding: 20px;
            border-radius: 12px;
            border: 1px solid rgba(233, 69, 96, 0.2);
        }
        
        .metric-box {
            background: var(--bg-dark);
            padding: 20px;
            border-radius: 12px;
            margin-bottom: 15px;
            text-align: center;
        }
        
        .metric-label {
            font-size: 12px;
            color: var(--text-secondary);
            margin-bottom: 8px;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        
        .metric-value {
            font-size: 28px;
            font-weight: 700;
            color: var(--accent);
        }
        
        .metric-value.success {
            color: var(--success);
        }
        
        .metric-value.warning {
            color: var(--warning);
        }
        
        .conclusion-box {
            background: linear-gradient(135deg, rgba(233, 69, 96, 0.2) 0%, rgba(233, 69, 96, 0.05) 100%);
            border: 1px solid var(--accent);
            border-radius: 12px;
            padding: 20px;
            margin-top: 20px;
            text-align: center;
        }
        
        .conclusion-text {
            font-size: 16px;
            color: var(--text-primary);
            line-height: 1.6;
        }
        
        .conclusion-highlight {
            color: var(--accent);
            font-weight: 700;
            font-size: 20px;
        }
        
        /* 表格样式 */
        .data-table {
            width: 100%;
            border-collapse: collapse;
        }
        
        .data-table th {
            background: var(--bg-hover);
            color: var(--text-primary);
            padding: 15px;
            text-align: left;
            font-weight: 600;
            font-size: 13px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .data-table td {
            padding: 15px;
            border-bottom: 1px solid var(--bg-hover);
            color: var(--text-primary);
        }
        
        .data-table tr:hover {
            background: var(--bg-hover);
        }
        
        .data-table tr.warning-row {
            background: rgba(244, 67, 54, 0.1);
            border-left: 4px solid var(--danger);
        }
        
        .data-table tr.warning-row:hover {
            background: rgba(244, 67, 54, 0.2);
        }
        
        /* 排行榜样式 */
        .ranking-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 25px;
        }
        
        .rank-item {
            display: flex;
            align-items: center;
            padding: 15px;
            background: var(--bg-dark);
            border-radius: 10px;
            margin-bottom: 10px;
            transition: all 0.3s ease;
        }
        
        .rank-item:hover {
            transform: translateX(5px);
            background: var(--bg-hover);
        }
        
        .rank-number {
            width: 36px;
            height: 36px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: 700;
            font-size: 14px;
            margin-right: 15px;
        }
        
        .rank-number.top3 {
            background: linear-gradient(135deg, #FFD700 0%, #FFA500 100%);
            color: #1a1a2e;
        }
        
        .rank-number.normal {
            background: var(--bg-hover);
            color: var(--text-secondary);
        }
        
        .rank-info {
            flex: 1;
        }
        
        .rank-name {
            font-weight: 500;
            color: var(--text-primary);
            margin-bottom: 4px;
        }
        
        .rank-value {
            font-size: 12px;
            color: var(--text-secondary);
        }
        
        .rank-profit {
            font-weight: 700;
            font-size: 16px;
        }
        
        .rank-profit.positive {
            color: var(--success);
        }
        
        .rank-profit.negative {
            color: var(--danger);
        }
        
        /* 图表容器 */
        .chart-container {
            position: relative;
            height: 300px;
            margin: 20px 0;
        }
        
        .chart-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 25px;
        }
        
        /* 分布统计 */
        .dist-grid {
            display: grid;
            grid-template-columns: repeat(6, 1fr);
            gap: 15px;
            margin: 20px 0;
        }
        
        .dist-item {
            background: var(--bg-dark);
            padding: 20px;
            border-radius: 12px;
            text-align: center;
            transition: all 0.3s ease;
        }
        
        .dist-item:hover {
            transform: translateY(-5px);
            background: var(--bg-hover);
        }
        
        .dist-range {
            font-size: 12px;
            color: var(--text-secondary);
            margin-bottom: 10px;
        }
        
        .dist-count {
            font-size: 32px;
            font-weight: 700;
            color: var(--accent);
        }
        
        .dist-item.danger .dist-count {
            color: var(--danger);
        }
        
        .dist-item.warning .dist-count {
            color: var(--warning);
        }
        
        .dist-item.success .dist-count {
            color: var(--success);
        }
        

        /* 下拉选择归因 */
        .attribution-select {
            padding: 8px 12px;
            background: var(--bg-dark);
            border: 1px solid var(--bg-hover);
            border-radius: 6px;
            color: var(--text-primary);
            font-size: 13px;
            cursor: pointer;
        }
        
        /* 低毛利警告 */
        .low-margin-alert {
            background: rgba(244, 67, 54, 0.1);
            border: 1px solid var(--danger);
            border-radius: 12px;
            padding: 15px 20px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .low-margin-alert i {
            color: var(--danger);
            font-size: 20px;
        }
        
        /* 响应式 */
        @media (max-width: 768px) {
            .calculator-grid,
            .ranking-grid,
            .chart-grid {
                grid-template-columns: 1fr;
            }
            
            .dist-grid {
                grid-template-columns: repeat(3, 1fr);
            }
            
            .tab-nav {
                flex-wrap: wrap;
            }
        }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-box-open"></i> 单品分析</h2>
            <div class="breadcrumb">
                <a href="index.asp">财务中心</a> / <span>单品分析</span>
            </div>
        </div>
        
        <!-- Tab导航 -->
        <div class="tab-nav">
            <a href="?tab=breakeven" class="<%= IIf(currentTab="breakeven", "active", "") %>"><i class="fas fa-calculator"></i> 盈亏平衡模型</a>
            <a href="?tab=costchange" class="<%= IIf(currentTab="costchange", "active", "") %>"><i class="fas fa-exchange-alt"></i> 成本异动分析</a>
            <a href="?tab=ranking" class="<%= IIf(currentTab="ranking", "active", "") %>"><i class="fas fa-trophy"></i> 利润排行榜</a>
            <a href="?tab=margin" class="<%= IIf(currentTab="margin", "active", "") %>"><i class="fas fa-chart-pie"></i> 毛利率分布</a>
        </div>
        
        <% If currentTab = "breakeven" Then %>
        <!-- Tab 1: 单品盈亏平衡模型 -->
        <div class="card">
            <div class="card-title"><i class="fas fa-calculator"></i> 单品盈亏平衡计算器</div>
            <div class="calculator-grid">
                <div class="input-section">
                    <form method="get" action="product_analysis.asp">
                        <input type="hidden" name="tab" value="breakeven">
                        
                        <div class="form-group">
                            <label><i class="fas fa-box"></i> 选择商品</label>
                            <select name="pid" class="form-control" onchange="this.form.submit()">
                                <option value="">-- 请选择商品 --</option>
                                <% If Not rsProducts Is Nothing Then %>
                                <% Do While Not rsProducts.EOF %>
                                <option value="<%= rsProducts("ProductID") %>" <%= IIf(CStr(rsProducts("ProductID"))=selectedProductId, "selected", "") %>>
                                    <%= Server.HTMLEncode(rsProducts("ProductName")) %> (¥<%= FormatNumber(CDbl("0" & rsProducts("BasePrice")), 2) %>)
                                </option>
                                <% rsProducts.MoveNext %>
                                <% Loop %>
                                <% rsProducts.MoveFirst %>
                                <% End If %>
                            </select>
                        </div>
                        
                        <% If selectedProductId <> "" Then %>
                        <div class="form-group">
                            <label><i class="fas fa-tag"></i> 类目扣点率 (%)</label>
                            <input type="number" name="crate" class="form-control" value="<%= categoryRate %>" step="0.1" min="0" max="100">
                        </div>
                        
                        <div class="form-group">
                            <label><i class="fas fa-percentage"></i> 预估转化率 (%)</label>
                            <input type="number" name="convr" class="form-control" value="<%= conversionRate %>" step="0.1" min="0" max="100">
                        </div>
                        
                        <div class="form-group">
                            <label><i class="fas fa-coins"></i> 固定费用率 (%)</label>
                            <input type="number" name="frate" class="form-control" value="<%= fixedRate %>" step="0.1" min="0" max="100">
                        </div>
                        
                        <button type="submit" class="btn btn-primary">
                            <i class="fas fa-sync-alt"></i> 重新计算
                        </button>
                        <% End If %>
                    </form>
                </div>
                
                <% If selectedProductId <> "" Then
                    Dim grossMargin, breakEvenROI, maxPromoFee, netMargin
                    
                    If productPrice > 0 Then
                        grossMargin = (productPrice - productBOMCost) / productPrice
                    Else
                        grossMargin = 0
                    End If
                    
                    netMargin = grossMargin - fixRateVal
                    
                    If netMargin > 0 Then
                        breakEvenROI = 1 / netMargin
                        maxPromoFee = productPrice * netMargin
                    Else
                        breakEvenROI = 0
                        maxPromoFee = 0
                    End If
                %>
                <div class="result-section">
                    <div class="metric-box">
                        <div class="metric-label">当前商品</div>
                        <div style="font-size: 18px; color: var(--text-primary);"><%= Server.HTMLEncode(productName) %></div>
                        <div style="font-size: 14px; color: var(--text-secondary); margin-top: 5px;">
                            售价: ¥<%= FormatNumber(productPrice, 2) %> | BOM成本: ¥<%= FormatNumber(productBOMCost, 2) %>
                        </div>
                    </div>
                    
                    <div class="metric-box">
                        <div class="metric-label">毛利率</div>
                        <div class="metric-value <%= IIf(grossMargin*100>=30, "success", IIf(grossMargin*100>=10, "warning", "")) %>">
                            <%= FormatNumber(grossMargin * 100, 2) %>%
                        </div>
                    </div>
                    
                    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px;">
                        <div class="metric-box">
                            <div class="metric-label">保本ROI</div>
                            <div class="metric-value <%= IIf(breakEvenROI>0, "", "warning") %>">
                                <%= IIf(breakEvenROI > 0, FormatNumber(breakEvenROI, 2), "N/A") %>
                            </div>
                        </div>
                        
                        <div class="metric-box">
                            <div class="metric-label">推广费上限</div>
                            <div class="metric-value <%= IIf(maxPromoFee>0, "success", "warning") %>">
                                ¥<%= FormatNumber(maxPromoFee, 2) %>
                            </div>
                        </div>
                    </div>
                    
                    <% If maxPromoFee > 0 Then %>
                    <div class="conclusion-box">
                        <div class="conclusion-text">
                            <i class="fas fa-exclamation-triangle" style="color: var(--accent);"></i><br>
                            该商品推广费出到 <span class="conclusion-highlight">¥<%= FormatNumber(maxPromoFee, 2) %></span> 会亏本
                        </div>
                    </div>
                    <% Else %>
                    <div class="conclusion-box">
                        <div class="conclusion-text" style="color: var(--danger);">
                            <i class="fas fa-times-circle"></i><br>
                            毛利率过低，无法覆盖固定费用，建议优化成本结构
                        </div>
                    </div>
                    <% End If %>
                </div>
                <% End If %>
            </div>
        </div>
        <% End If %>
        
        <% If currentTab = "costchange" Then %>
        <!-- Tab 2: SKU成本异动分析 -->
        <div class="card">
            <div class="card-title"><i class="fas fa-exchange-alt"></i> SKU成本异动分析 (<%= lastMonth %> vs <%= currentMonth %>)</div>
            
            <table class="data-table">
                <thead>
                    <tr>
                        <th>商品名称</th>
                        <th>上月成本</th>
                        <th>本月成本</th>
                        <th>波动金额</th>
                        <th>波动率</th>
                        <th>状态</th>
                        <th>归因</th>
                    </tr>
                </thead>
                <tbody>
                    <% 
                    Dim hasCostData, lastCost, currCost, costDiff, costChangeRate, isWarning
                    hasCostData = False
                    If Not rsCostChange Is Nothing Then
                    Do While Not rsCostChange.EOF
                        lastCost = CDbl("0" & rsCostChange("LastMonthCost"))
                        currCost = CDbl("0" & rsCostChange("CurrentMonthCost"))
                        
                        If lastCost > 0 Or currCost > 0 Then
                            hasCostData = True
                            costDiff = currCost - lastCost
                            If lastCost > 0 Then
                                costChangeRate = (costDiff / lastCost) * 100
                            Else
                                costChangeRate = 0
                            End If
                            
                            isWarning = False
                            If lastCost > 0 And Abs(costChangeRate) > 5 Then
                                isWarning = True
                            End If
                    %>
                    <tr class="<%= IIf(isWarning, "warning-row", "") %>">
                        <td><%= Server.HTMLEncode(rsCostChange("ProductName")) %></td>
                        <td>¥<%= FormatNumber(lastCost, 2) %></td>
                        <td>¥<%= FormatNumber(currCost, 2) %></td>
                        <td style="color: <%= IIf(costDiff>0, "#f44336", "#4CAF50") %>">
                            <%= IIf(costDiff>0, "+", "") %><%= FormatNumber(costDiff, 2) %>
                        </td>
                        <td style="color: <%= IIf(costChangeRate>0, "#f44336", "#4CAF50") %>">
                            <%= IIf(costChangeRate>0, "+", "") %><%= FormatNumber(costChangeRate, 2) %>%
                        </td>
                        <td>
                            <% If isWarning Then %>
                            <span style="color: #f44336;"><i class="fas fa-exclamation-circle"></i> 异常</span>
                            <% Else %>
                            <span style="color: #4CAF50;"><i class="fas fa-check-circle"></i> 正常</span>
                            <% End If %>
                        </td>
                        <td>
                            <select class="attribution-select">
                                <option value="">-- 选择归因 --</option>
                                <option value="purchase">采购价上涨</option>
                                <option value="logistics">物流涨价</option>
                                <option value="material">原料涨价</option>
                                <option value="other">其他</option>
                            </select>
                        </td>
                    </tr>
                    <% 
                        End If
                        rsCostChange.MoveNext
                    Loop
                    End If
                    %>
                </tbody>
            </table>
            
            <% If Not hasCostData Then %>
            <div style="text-align: center; padding: 40px; color: var(--text-secondary);">
                <i class="fas fa-inbox" style="font-size: 48px; margin-bottom: 15px; display: block;"></i>
                暂无成本数据，请在 ProductCosts 表中维护成本信息
            </div>
            <% End If %>
        </div>
        <% End If %>
        
        <% If currentTab = "ranking" Then %>
        <!-- Tab 3: 单品利润排行榜 -->
        <div class="ranking-grid">
            <div class="card">
                <div class="card-title"><i class="fas fa-trophy"></i> TOP10 最盈利商品</div>
                <% 
                Dim rankNum, hasTopData
                rankNum = 0
                hasTopData = False
                If Not rsTopProfit Is Nothing Then
                Do While Not rsTopProfit.EOF
                    hasTopData = True
                    rankNum = rankNum + 1
                %>
                <div class="rank-item">
                    <div class="rank-number <%= IIf(rankNum<=3, "top3", "normal") %>"><%= rankNum %></div>
                    <div class="rank-info">
                        <div class="rank-name"><%= Server.HTMLEncode(rsTopProfit("ProductName")) %></div>
                        <div class="rank-value">销售额: ¥<%= FormatNumber(SafeNum(rsTopProfit("TotalRevenue")), 0) %></div>
                    </div>
                    <div class="rank-profit positive">+¥<%= FormatNumber(SafeNum(rsTopProfit("Profit")), 0) %></div>
                </div>
                <% 
                    rsTopProfit.MoveNext
                Loop
                End If
                
                If Not hasTopData Then
                %>
                <div style="text-align: center; padding: 40px; color: var(--text-secondary);">
                    <i class="fas fa-inbox" style="font-size: 48px; margin-bottom: 15px; display: block;"></i>
                    暂无订单数据
                </div>
                <% End If %>
            </div>
            
            <div class="card">
                <div class="card-title"><i class="fas fa-arrow-down"></i> TOP10 最亏损商品</div>
                <% 
                rankNum = 0
                hasTopData = False
                If Not rsBottomProfit Is Nothing Then
                Do While Not rsBottomProfit.EOF
                    hasTopData = True
                    rankNum = rankNum + 1
                %>
                <div class="rank-item">
                    <div class="rank-number <%= IIf(rankNum<=3, "top3", "normal") %>"><%= rankNum %></div>
                    <div class="rank-info">
                        <div class="rank-name"><%= Server.HTMLEncode(rsBottomProfit("ProductName")) %></div>
                        <div class="rank-value">销售额: ¥<%= FormatNumber(SafeNum(rsBottomProfit("TotalRevenue")), 0) %></div>
                    </div>
                    <div class="rank-profit negative">¥<%= FormatNumber(SafeNum(rsBottomProfit("Profit")), 0) %></div>
                </div>
                <% 
                    rsBottomProfit.MoveNext
                Loop
                End If
                
                If Not hasTopData Then
                %>
                <div style="text-align: center; padding: 40px; color: var(--text-secondary);">
                    <i class="fas fa-inbox" style="font-size: 48px; margin-bottom: 15px; display: block;"></i>
                    暂无订单数据
                </div>
                <% End If %>
            </div>
        </div>
        
        <!-- 利润排行图表 -->
        <div class="card">
            <div class="card-title"><i class="fas fa-chart-bar"></i> 利润排行可视化</div>
            <div class="chart-grid">
                <div class="chart-container">
                    <canvas id="topProfitChart"></canvas>
                </div>
                <div class="chart-container">
                    <canvas id="bottomProfitChart"></canvas>
                </div>
            </div>
        </div>
        <% End If %>
        
        <% If currentTab = "margin" Then %>
        <!-- Tab 4: 商品毛利率分布 -->
        <div class="card">
            <div class="card-title"><i class="fas fa-chart-pie"></i> 毛利率分布统计</div>
            
            <div class="dist-grid">
                <div class="dist-item danger">
                    <div class="dist-range">< 0%</div>
                    <div class="dist-count"><%= cntNeg %></div>
                    <div style="font-size: 12px; color: var(--text-secondary);">亏损商品</div>
                </div>
                <div class="dist-item warning">
                    <div class="dist-range">0-10%</div>
                    <div class="dist-count"><%= cnt0_10 %></div>
                    <div style="font-size: 12px; color: var(--text-secondary);">低毛利</div>
                </div>
                <div class="dist-item">
                    <div class="dist-range">10-20%</div>
                    <div class="dist-count"><%= cnt10_20 %></div>
                    <div style="font-size: 12px; color: var(--text-secondary);">一般</div>
                </div>
                <div class="dist-item">
                    <div class="dist-range">20-30%</div>
                    <div class="dist-count"><%= cnt20_30 %></div>
                    <div style="font-size: 12px; color: var(--text-secondary);">良好</div>
                </div>
                <div class="dist-item success">
                    <div class="dist-range">30-50%</div>
                    <div class="dist-count"><%= cnt30_50 %></div>
                    <div style="font-size: 12px; color: var(--text-secondary);">优秀</div>
                </div>
                <div class="dist-item success">
                    <div class="dist-range">> 50%</div>
                    <div class="dist-count"><%= cnt50plus %></div>
                    <div style="font-size: 12px; color: var(--text-secondary);">超高</div>
                </div>
            </div>
            
            <div class="chart-grid">
                <div class="chart-container">
                    <canvas id="marginBarChart"></canvas>
                </div>
                <div class="chart-container">
                    <canvas id="marginPieChart"></canvas>
                </div>
            </div>
        </div>
        
        <!-- 低毛利商品列表 -->
        <div class="card">
            <div class="card-title"><i class="fas fa-exclamation-triangle"></i> 低毛利商品列表（毛利率 < 10%，需关注）</div>
            
            <% If cntNeg + cnt0_10 > 0 Then %>
            <div class="low-margin-alert">
                <i class="fas fa-info-circle"></i>
                <span>发现 <%= cntNeg + cnt0_10 %> 个毛利率低于10%的商品，建议优化定价策略或降低成本</span>
            </div>
            <% End If %>
            
            <table class="data-table">
                <thead>
                    <tr>
                        <th>商品名称</th>
                        <th>售价</th>
                        <th>成本</th>
                        <th>毛利</th>
                        <th>毛利率</th>
                        <th>状态</th>
                    </tr>
                </thead>
                <tbody>
                    <% 
                    Dim hasLowMargin : hasLowMargin = False
                    If Not rsMarginDist Is Nothing Then
                    Do While Not rsMarginDist.EOF
                        Dim pMargin
                        pMargin = CDbl("0" & rsMarginDist("MarginRate"))
                        If pMargin < 10 Then
                            hasLowMargin = True
                    %>
                    <tr class="<%= IIf(pMargin < 0, "warning-row", "") %>">
                        <td><%= Server.HTMLEncode(rsMarginDist("ProductName")) %></td>
                        <td>¥<%= FormatNumber(SafeNum(rsMarginDist("BasePrice")), 2) %></td>
                        <td>¥<%= FormatNumber(SafeNum(rsMarginDist("UnitCost")), 2) %></td>
                        <td>¥<%= FormatNumber(SafeNum(rsMarginDist("BasePrice")) - SafeNum(rsMarginDist("UnitCost")), 2) %></td>
                        <td style="color: <%= IIf(pMargin < 0, "#f44336", "#ff9800") %>">
                            <%= FormatNumber(pMargin, 2) %>%
                        </td>
                        <td>
                            <% If pMargin < 0 Then %>
                            <span style="color: #f44336;"><i class="fas fa-times-circle"></i> 亏损</span>
                            <% Else %>
                            <span style="color: #ff9800;"><i class="fas fa-exclamation-circle"></i> 低毛利</span>
                            <% End If %>
                        </td>
                    </tr>
                    <% 
                        End If
                        rsMarginDist.MoveNext
                    Loop
                    End If
                    %>
                </tbody>
            </table>
            
            <% If Not hasLowMargin Then %>
            <div style="text-align: center; padding: 40px; color: var(--text-secondary);">
                <i class="fas fa-check-circle" style="font-size: 48px; margin-bottom: 15px; display: block; color: var(--success);"></i>
                恭喜！所有商品毛利率均在10%以上
            </div>
            <% End If %>
        </div>
        <% End If %>
    </div>
    
    <% If currentTab = "ranking" Then %>
    <script>
        // TOP10 盈利商品图表
        const topProfitCtx = document.getElementById('topProfitChart').getContext('2d');
        new Chart(topProfitCtx, {
            type: 'bar',
            data: {
                labels: [<% If Not rsTopProfit Is Nothing Then rsTopProfit.MoveFirst %><% Do While Not rsTopProfit.EOF %>'<%= Replace(Server.HTMLEncode(rsTopProfit("ProductName")), "'", "\'") %>'<%= IIf(Not rsTopProfit.EOF And Not rsTopProfit.AbsolutePosition >= 10, ",", "") %><% rsTopProfit.MoveNext %><% Loop %>],
                datasets: [{
                    label: '利润 (¥)',
                    data: [<% If Not rsTopProfit Is Nothing Then rsTopProfit.MoveFirst %><% Do While Not rsTopProfit.EOF %><%= CDbl("0" & rsTopProfit("Profit")) %><%= IIf(Not rsTopProfit.EOF And Not rsTopProfit.AbsolutePosition >= 10, ",", "") %><% rsTopProfit.MoveNext %><% Loop %>],
                    backgroundColor: 'rgba(76, 175, 80, 0.8)',
                    borderColor: '#4CAF50',
                    borderWidth: 1,
                    borderRadius: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    title: {
                        display: true,
                        text: 'TOP10 盈利商品',
                        color: '#eaeaea',
                        font: { size: 14 }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        grid: { color: 'rgba(255,255,255,0.1)' },
                        ticks: { color: '#a0a0a0' }
                    },
                    x: {
                        grid: { display: false },
                        ticks: { 
                            color: '#a0a0a0',
                            maxRotation: 45
                        }
                    }
                }
            }
        });
        
        // TOP10 亏损商品图表
        const bottomProfitCtx = document.getElementById('bottomProfitChart').getContext('2d');
        new Chart(bottomProfitCtx, {
            type: 'bar',
            data: {
                labels: [<% If Not rsBottomProfit Is Nothing Then rsBottomProfit.MoveFirst %><% Do While Not rsBottomProfit.EOF %>'<%= Replace(Server.HTMLEncode(rsBottomProfit("ProductName")), "'", "\'") %>'<%= IIf(Not rsBottomProfit.EOF And Not rsBottomProfit.AbsolutePosition >= 10, ",", "") %><% rsBottomProfit.MoveNext %><% Loop %>],
                datasets: [{
                    label: '利润 (¥)',
                    data: [<% If Not rsBottomProfit Is Nothing Then rsBottomProfit.MoveFirst %><% Do While Not rsBottomProfit.EOF %><%= rsBottomProfit("Profit") %><%= IIf(Not rsBottomProfit.EOF And Not rsBottomProfit.AbsolutePosition >= 10, ",", "") %><% rsBottomProfit.MoveNext %><% Loop %>],
                    backgroundColor: 'rgba(244, 67, 54, 0.8)',
                    borderColor: '#f44336',
                    borderWidth: 1,
                    borderRadius: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    title: {
                        display: true,
                        text: 'TOP10 亏损商品',
                        color: '#eaeaea',
                        font: { size: 14 }
                    }
                },
                scales: {
                    y: {
                        grid: { color: 'rgba(255,255,255,0.1)' },
                        ticks: { color: '#a0a0a0' }
                    },
                    x: {
                        grid: { display: false },
                        ticks: { 
                            color: '#a0a0a0',
                            maxRotation: 45
                        }
                    }
                }
            }
        });
    </script>
    <% End If %>
    
    <% If currentTab = "margin" Then %>
    <script>
        // 毛利率分布柱状图
        const marginBarCtx = document.getElementById('marginBarChart').getContext('2d');
        new Chart(marginBarCtx, {
            type: 'bar',
            data: {
                labels: ['<0%', '0-10%', '10-20%', '20-30%', '30-50%', '>50%'],
                datasets: [{
                    label: '商品数量',
                    data: [<%= cntNeg %>, <%= cnt0_10 %>, <%= cnt10_20 %>, <%= cnt20_30 %>, <%= cnt30_50 %>, <%= cnt50plus %>],
                    backgroundColor: [
                        'rgba(244, 67, 54, 0.8)',
                        'rgba(255, 152, 0, 0.8)',
                        'rgba(33, 150, 243, 0.8)',
                        'rgba(0, 150, 136, 0.8)',
                        'rgba(76, 175, 80, 0.8)',
                        'rgba(156, 39, 176, 0.8)'
                    ],
                    borderColor: [
                        '#f44336',
                        '#ff9800',
                        '#2196F3',
                        '#009688',
                        '#4CAF50',
                        '#9C27B0'
                    ],
                    borderWidth: 1,
                    borderRadius: 8
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    title: {
                        display: true,
                        text: '毛利率分布 (柱状图)',
                        color: '#eaeaea',
                        font: { size: 14 }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        grid: { color: 'rgba(255,255,255,0.1)' },
                        ticks: { color: '#a0a0a0' }
                    },
                    x: {
                        grid: { display: false },
                        ticks: { color: '#a0a0a0' }
                    }
                }
            }
        });
        
        // 毛利率分布饼图
        const marginPieCtx = document.getElementById('marginPieChart').getContext('2d');
        new Chart(marginPieCtx, {
            type: 'doughnut',
            data: {
                labels: ['<0%', '0-10%', '10-20%', '20-30%', '30-50%', '>50%'],
                datasets: [{
                    data: [<%= cntNeg %>, <%= cnt0_10 %>, <%= cnt10_20 %>, <%= cnt20_30 %>, <%= cnt30_50 %>, <%= cnt50plus %>],
                    backgroundColor: [
                        '#f44336',
                        '#ff9800',
                        '#2196F3',
                        '#009688',
                        '#4CAF50',
                        '#9C27B0'
                    ],
                    borderColor: '#16213e',
                    borderWidth: 2
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: {
                        display: true,
                        text: '毛利率分布 (占比)',
                        color: '#eaeaea',
                        font: { size: 14 }
                    },
                    legend: {
                        position: 'right',
                        labels: { color: '#a0a0a0' }
                    }
                }
            }
        });
    </script>
    <% End If %>
</body>
</html>
<%
If Not rsProducts Is Nothing Then rsProducts.Close
If Not rsCostChange Is Nothing Then rsCostChange.Close
If Not rsTopProfit Is Nothing Then rsTopProfit.Close
If Not rsBottomProfit Is Nothing Then rsBottomProfit.Close
If Not rsMarginDist Is Nothing Then rsMarginDist.Close

Set rsProducts = Nothing
Set rsCostChange = Nothing
Set rsTopProfit = Nothing
Set rsBottomProfit = Nothing
Set rsMarginDist = Nothing

Call CloseConnection()
%>
