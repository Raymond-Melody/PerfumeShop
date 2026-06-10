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

' 安全数值转换（避免数据库空值/类型导致CDbl失败）
Function SafeCDbl(val)
    If IsNumeric(val) Then SafeCDbl = CDbl(val) Else SafeCDbl = 0
End Function

' V8：关联成本中心
On Error Resume Next
conn.Execute "SELECT CenterID FROM BudgetItems WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE BudgetItems ADD CenterID INT NULL"
On Error GoTo 0

' ============================================
' 权限检查
' ============================================
' isManager 已在 nav.asp 中声明
Dim canEdit
canEdit = False

If Session("AdminRoleCode") = "FIN_MANAGER" OR Session("AdminRoleCode") = "SUPER_ADMIN" Then
    isManager = True
    canEdit = True
End If

' ============================================
' 预算类目定义
' ============================================
Dim budgetCategories(5), categoryLabels(5)
budgetCategories(0) = "Promotion" : categoryLabels(0) = "推广费用"
budgetCategories(1) = "Shipping" : categoryLabels(1) = "运费"
budgetCategories(2) = "Platform" : categoryLabels(2) = "平台费用"
budgetCategories(3) = "Purchase" : categoryLabels(3) = "采购成本"
budgetCategories(4) = "Salary" : categoryLabels(4) = "人工成本"
budgetCategories(5) = "Other" : categoryLabels(5) = "其他费用"

' ============================================
' 获取当前Tab和月份参数
' ============================================
Dim currentTab, currentMonth
currentTab = Trim(Request.QueryString("tab"))
If currentTab = "" Then currentTab = "1"
currentMonth = Trim(Request.QueryString("month"))
If currentMonth = "" Then currentMonth = Year(Date()) & "-" & Right("0" & Month(Date()), 2)

Dim safeMonth
safeMonth = SafeSQL(currentMonth)

' ============================================
' 处理预算保存请求
' ============================================
Dim saveMessage, saveSuccess
saveMessage = ""
saveSuccess = False

If Request.ServerVariables("REQUEST_METHOD") = "POST" AND canEdit Then
    Dim actionType
    actionType = Trim(Request.Form("action"))
    
    If actionType = "save_budget" Then
        Dim saveMonth, catIndex, budgetAmount
        saveMonth = SafeSQL(Trim(Request.Form("budget_month")))
        
        If saveMonth <> "" Then
            Dim i, catName, amountVal
            ' 循环变量声明移到循环外（VBScript限制）
            Dim existsCheck, saveSQL
            For i = 0 To UBound(budgetCategories)
                catName = budgetCategories(i)
                amountVal = Trim(Request.Form("budget_" & catName))
                
                If IsNumeric(amountVal) Then
                    ' 检查是否已存在记录
                    existsCheck = GetScalar("SELECT COUNT(*) FROM BudgetPlans WHERE BudgetMonth = '" & saveMonth & "' AND Category = '" & catName & "'")
                    If CLng(existsCheck) > 0 Then
                        saveSQL = "UPDATE BudgetPlans SET BudgetAmount = " & CDbl(amountVal) & ", UpdatedAt = GETDATE() WHERE BudgetMonth = '" & saveMonth & "' AND Category = '" & catName & "'"
                    Else
                        saveSQL = "INSERT INTO BudgetPlans (BudgetMonth, Category, BudgetAmount, CreatedAt, UpdatedAt) VALUES ('" & saveMonth & "', '" & catName & "', " & CDbl(amountVal) & ", GETDATE(), GETDATE())"
                    End If
                    
                    Call ExecuteNonQuery(saveSQL)
                End If
            Next
            
            saveSuccess = True
            saveMessage = "预算保存成功！"
            Call LogAdminAction("保存月度预算", "finance", "BudgetPlans", "", saveMonth)
        End If
    End If
End If

' ============================================
' Tab 1: 获取月度预算数据
' ============================================
Dim rsBudget, rsLastMonthActual
Set rsBudget = ExecuteQuery("SELECT Category, BudgetAmount FROM BudgetPlans WHERE BudgetMonth = '" & safeMonth & "'")

' 获取上月实际数据
Dim lastMonth, lastMonthYear, lastMonthNum
lastMonthNum = Month(Date()) - 1
lastMonthYear = Year(Date())
If lastMonthNum = 0 Then
    lastMonthNum = 12
    lastMonthYear = lastMonthYear - 1
End If
lastMonth = lastMonthYear & "-" & Right("0" & lastMonthNum, 2)

' ============================================
' Tab 2: 费用执行率数据
' ============================================
' 获取本月预算总计
Dim monthBudgetTotal
monthBudgetTotal = GetScalar("SELECT IIF(SUM(BudgetAmount) IS NULL, 0, SUM(BudgetAmount)) FROM BudgetPlans WHERE BudgetMonth = '" & safeMonth & "'")

' 获取本月实际消耗（按类型）
Dim rsActualExpense
Set rsActualExpense = ExecuteQuery(_
    "SELECT ExpenseType, SUM(Amount) AS TotalAmount FROM ExpenseRecords " & _
    "WHERE ExpenseMonth = '" & safeMonth & "' GROUP BY ExpenseType")

' 获取本月GMV
Dim monthGMV
monthGMV = GetScalar("SELECT CAST(IIF(SUM(TotalAmount) IS NULL, 0, SUM(TotalAmount)) AS FLOAT) FROM Orders WHERE Status = 'Paid' AND OrderDate >= '" & safeMonth & "-01' AND OrderDate < DATEADD(month, 1, '" & safeMonth & "-01')")

' ============================================
' Tab 3: 现金流预测数据
' ============================================
' 期初余额
Dim openingBalance
openingBalance = SafeCDbl(GetScalar("SELECT IIF(SUM(AvailableBalance) IS NULL, 0, SUM(AvailableBalance)) FROM FundAccounts WHERE IsActive = 1"))

' 最近7天预计回款
Dim expectedCollection
expectedCollection = SafeCDbl(GetScalar("SELECT CAST(IIF(SUM(TotalAmount) IS NULL, 0, SUM(TotalAmount)) AS FLOAT) FROM Orders WHERE Status = 'Paid' AND OrderDate >= DATEADD(day, -7, CAST(GETDATE() AS DATE))"))

' --- 预计算 SVG 图表数据（确保所有计算在主逻辑块完成，避免 HTML 体内渲染中断）---
' 先构建 budgetData（原本在 HTML 体内，现提前至此供 SVG 计算使用）
Dim budgetData
Set budgetData = CreateObject("Scripting.Dictionary")
If Not rsBudget Is Nothing Then
    Do While Not rsBudget.EOF
        budgetData.Add rsBudget("Category"), rsBudget("BudgetAmount")
        rsBudget.MoveNext
    Loop
    rsBudget.Close
End If

Dim chartDailyData(29), chartMinVal, chartMaxVal, chartSafeLine
Dim chartPurchaseDaily, chartFixedDaily, chartPurchaseBudget, chartFixedBudget
Dim chartPoints, chartAreaPath, chartSafeLineY

chartPurchaseBudget = 0
chartFixedBudget = 0
If budgetData.Exists("Purchase") Then chartPurchaseBudget = CDbl(budgetData("Purchase"))
If budgetData.Exists("Salary") Then chartFixedBudget = chartFixedBudget + CDbl(budgetData("Salary"))
If budgetData.Exists("Platform") Then chartFixedBudget = chartFixedBudget + CDbl(budgetData("Platform"))

chartPurchaseDaily = chartPurchaseBudget / 30
chartFixedDaily = chartFixedBudget / 30
chartSafeLine = chartFixedDaily * 60

Dim chartDayIdx, chartProjIncome, chartProjExpense, chartDayBalance
For chartDayIdx = 0 To 29
    chartProjIncome = expectedCollection / 7
    chartProjExpense = chartPurchaseDaily + chartFixedDaily
    chartDayBalance = openingBalance + (chartProjIncome * (chartDayIdx + 1)) - (chartProjExpense * (chartDayIdx + 1))
    chartDailyData(chartDayIdx) = chartDayBalance
Next

chartMinVal = chartDailyData(0)
chartMaxVal = chartDailyData(0)
For chartDayIdx = 1 To 29
    If chartDailyData(chartDayIdx) < chartMinVal Then chartMinVal = chartDailyData(chartDayIdx)
    If chartDailyData(chartDayIdx) > chartMaxVal Then chartMaxVal = chartDailyData(chartDayIdx)
Next
If chartMinVal < 0 Then chartMinVal = 0
If chartMaxVal = chartMinVal Then chartMaxVal = chartMinVal + 1000

Dim chartXPos, chartYPos
chartPoints = ""
For chartDayIdx = 0 To 29
    chartXPos = 50 + (chartDayIdx * 25)
    chartYPos = 200 - ((chartDailyData(chartDayIdx) - chartMinVal) / (chartMaxVal - chartMinVal) * 180)
    If chartDayIdx = 0 Then
        chartPoints = chartXPos & "," & chartYPos
        chartAreaPath = "M" & chartXPos & ",200 L" & chartXPos & "," & chartYPos
    Else
        chartPoints = chartPoints & " " & chartXPos & "," & chartYPos
        chartAreaPath = chartAreaPath & " L" & chartXPos & "," & chartYPos
    End If
Next
chartAreaPath = chartAreaPath & " L" & (50 + 29 * 25) & ",200 Z"

If chartSafeLine >= chartMinVal AND chartSafeLine <= chartMaxVal Then
    chartSafeLineY = 200 - ((chartSafeLine - chartMinVal) / (chartMaxVal - chartMinVal) * 180)
Else
    chartSafeLineY = 200
End If
' --- SVG 预计算结束 ---

' ============================================
' Tab 4: 季度/年度汇总数据
' ============================================
Dim currentYear, currentQuarter
currentYear = Year(Date())
currentQuarter = Int((Month(Date()) - 1) / 3) + 1

Dim rsQuarterlyBudget
Set rsQuarterlyBudget = ExecuteQuery(_
    "SELECT Category, SUM(BudgetAmount) AS TotalBudget FROM BudgetPlans " & _
    "WHERE Year(BudgetMonth) = " & currentYear & " GROUP BY Category")

Call LogAdminAction("查看预算管理", "finance", "", "", "Tab: " & currentTab)
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>预算管理 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root {
            --dark-bg: #1a1a2e;
            --dark-card: #2d2d44;
            --dark-border: rgba(255,255,255,0.06);
            --text-primary: #e0e0e0;
            --text-secondary: #888;
            --accent-blue: #00bcd4;
            --accent-green: #69f0ae;
            --accent-yellow: #ffd54f;
            --accent-red: #ff5252;
            --accent-purple: #e040fb;
        }
        
        body {
            background: var(--dark-bg);
            color: var(--text-primary);
        }
        
        .main-content {
            background: var(--dark-bg);
        }
        
        .page-header {
            background: var(--dark-card);
            padding: 20px 30px;
            border-radius: 12px;
            margin-bottom: 25px;
            border: 1px solid var(--dark-border);
        }
        
        .page-title {
            color: var(--text-primary);
            margin: 0;
        }
        
        /* Tab导航 */
        .tab-nav {
            display: flex;
            gap: 5px;
            margin-bottom: 25px;
            background: var(--dark-card);
            padding: 10px;
            border-radius: 12px;
            border: 1px solid var(--dark-border);
        }
        
        .tab-btn {
            padding: 12px 24px;
            background: transparent;
            border: none;
            color: var(--text-secondary);
            cursor: pointer;
            border-radius: 8px;
            transition: all 0.3s;
            font-size: 14px;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .tab-btn:hover {
            background: rgba(79, 195, 247, 0.1);
            color: var(--accent-blue);
        }
        
        .tab-btn.active {
            background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%);
            color: white;
        }
        
        /* Tab内容 */
        .tab-content {
            display: none;
        }
        
        .tab-content.active {
            display: block;
        }
        
        /* 卡片样式 */
        .budget-card {
            background: var(--dark-card);
            border-radius: 12px;
            padding: 25px;
            margin-bottom: 25px;
            border: 1px solid var(--dark-border);
        }
        
        .card-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            padding-bottom: 15px;
            border-bottom: 1px solid var(--dark-border);
        }
        
        .card-title {
            font-size: 18px;
            color: var(--text-primary);
            margin: 0;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        /* 表格样式 */
        .budget-table {
            width: 100%;
            border-collapse: collapse;
        }
        
        .budget-table th {
            background: rgba(0, 188, 212, 0.2);
            color: var(--accent-blue);
            padding: 15px;
            text-align: left;
            font-weight: 600;
            border-bottom: 2px solid var(--dark-border);
        }
        
        .budget-table td {
            padding: 15px;
            border-bottom: 1px solid var(--dark-border);
            color: var(--text-primary);
        }
        
        .budget-table tr:hover {
            background: rgba(255,255,255,0.03);
        }
        
        /* 输入框样式 */
        .budget-input {
            background: var(--dark-bg);
            border: 1px solid var(--dark-border);
            color: var(--text-primary);
            padding: 10px 15px;
            border-radius: 6px;
            width: 150px;
            font-size: 14px;
        }
        
        .budget-input:focus {
            outline: none;
            border-color: var(--accent-blue);
            box-shadow: 0 0 0 3px rgba(79, 195, 247, 0.1);
        }
        
        .budget-input:disabled {
            background: rgba(255,255,255,0.05);
            cursor: not-allowed;
        }
        
        /* 月份选择器 */
        .month-selector {
            display: flex;
            align-items: center;
            gap: 15px;
            margin-bottom: 20px;
        }
        
        .month-input {
            background: var(--dark-bg);
            border: 1px solid var(--dark-border);
            color: var(--text-primary);
            padding: 10px 15px;
            border-radius: 6px;
            font-size: 14px;
        }
        
        /* 进度条 */
        .progress-container {
            width: 150px;
            height: 20px;
            background: var(--dark-bg);
            border-radius: 10px;
            overflow: hidden;
            position: relative;
        }
        
        .progress-fill {
            height: 100%;
            border-radius: 10px;
            transition: width 0.5s ease;
        }
        
        .progress-fill.blue { background: linear-gradient(90deg, #2196F3, #03a9f4); }
        .progress-fill.green { background: linear-gradient(90deg, #4CAF50, #8bc34a); }
        .progress-fill.yellow { background: linear-gradient(90deg, #ff9800, #ffc107); }
        .progress-fill.red { background: linear-gradient(90deg, #f44336, #ff5722); }
        
        .progress-text {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            font-size: 11px;
            font-weight: bold;
            color: white;
            text-shadow: 0 1px 2px rgba(0,0,0,0.5);
        }
        
        /* 统计卡片 */
        .stats-row {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 20px;
            margin-bottom: 25px;
        }
        
        .stat-box {
            background: var(--dark-card);
            border-radius: 12px;
            padding: 20px;
            border: 1px solid var(--dark-border);
            text-align: center;
        }
        
        .stat-box-icon {
            width: 50px;
            height: 50px;
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            margin: 0 auto 15px;
            font-size: 20px;
        }
        
        .stat-box-icon.blue { background: rgba(79, 195, 247, 0.2); color: var(--accent-blue); }
        .stat-box-icon.green { background: rgba(105, 240, 174, 0.2); color: var(--accent-green); }
        .stat-box-icon.yellow { background: rgba(255, 213, 79, 0.2); color: var(--accent-yellow); }
        .stat-box-icon.purple { background: rgba(224, 64, 251, 0.2); color: var(--accent-purple); }
        
        .stat-box-value {
            font-size: 24px;
            font-weight: bold;
            color: var(--text-primary);
            margin-bottom: 5px;
        }
        
        .stat-box-label {
            font-size: 13px;
            color: var(--text-secondary);
        }
        
        /* 资金水位图 */
        .cashflow-chart {
            height: 300px;
            position: relative;
            background: var(--dark-bg);
            border-radius: 8px;
            padding: 20px;
            margin-top: 20px;
        }
        
        .chart-grid {
            position: absolute;
            top: 20px;
            left: 50px;
            right: 20px;
            bottom: 50px;
            border-left: 1px solid var(--dark-border);
            border-bottom: 1px solid var(--dark-border);
        }
        
        .chart-line {
            position: absolute;
            bottom: 50px;
            left: 50px;
            right: 20px;
            height: 200px;
        }
        
        .chart-area {
            fill: url(#gradientBlue);
            opacity: 0.6;
        }
        
        .chart-path {
            fill: none;
            stroke: var(--accent-blue);
            stroke-width: 2;
        }
        
        .safe-line {
            position: absolute;
            left: 50px;
            right: 20px;
            border-top: 2px dashed var(--accent-red);
            opacity: 0.7;
        }
        
        /* 预警提示 */
        .alert-box {
            background: rgba(255, 82, 82, 0.1);
            border: 1px solid var(--accent-red);
            border-radius: 8px;
            padding: 15px 20px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 12px;
            color: var(--accent-red);
        }
        
        .alert-box.warning {
            background: rgba(255, 213, 79, 0.1);
            border-color: var(--accent-yellow);
            color: var(--accent-yellow);
        }
        
        /* 模态框 */
        .modal {
            display: none;
            position: fixed;
            z-index: 3000;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.7);
        }
        
        .modal-content {
            background: var(--dark-card);
            margin: 10% auto;
            padding: 0;
            border-radius: 12px;
            width: 90%;
            max-width: 500px;
            border: 1px solid var(--dark-border);
        }
        
        .modal-header {
            padding: 20px 25px;
            border-bottom: 1px solid var(--dark-border);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .modal-title {
            margin: 0;
            color: var(--text-primary);
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .modal-body {
            padding: 25px;
            color: var(--text-secondary);
        }
        
        .modal-footer {
            padding: 15px 25px;
            border-top: 1px solid var(--dark-border);
            display: flex;
            justify-content: flex-end;
            gap: 10px;
        }
        
        .close-btn {
            background: none;
            border: none;
            color: var(--text-secondary);
            font-size: 24px;
            cursor: pointer;
        }
        
        /* 按钮样式覆盖 */
        .admin-btn-primary {
            background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%);
        }
        
        .admin-btn-success {
            background: linear-gradient(135deg, #4CAF50, #388e3c);
        }
        
        /* 汇总报表样式 */
        .summary-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 25px;
        }
        
        .trend-indicator {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            font-size: 12px;
            padding: 3px 8px;
            border-radius: 4px;
        }
        
        .trend-up {
            color: var(--accent-green);
            background: rgba(105, 240, 174, 0.1);
        }
        
        .trend-down {
            color: var(--accent-red);
            background: rgba(255, 82, 82, 0.1);
        }
        
        .amount-positive { color: var(--accent-green); }
        .amount-negative { color: var(--accent-red); }
        
        /* 响应式 */
        @media (max-width: 768px) {
            .stats-row { grid-template-columns: repeat(2, 1fr); }
            .summary-grid { grid-template-columns: 1fr; }
            .tab-nav { flex-wrap: wrap; }
        }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-chart-pie"></i> 预算管理中心</h2>
            <div style="color: var(--text-secondary); font-size: 14px;">
                <% If isManager Then %>
                <span style="color: var(--accent-green);"><i class="fas fa-user-shield"></i> 管理权限</span>
                <% Else %>
                <span style="color: var(--accent-blue);"><i class="fas fa-eye"></i> 只读权限</span>
                <% End If %>
            </div>
        </div>
        
        <% If saveMessage <> "" Then %>
        <div class="alert-box <%= IIf(saveSuccess, "warning", "") %>">
            <i class="fas <%= IIf(saveSuccess, "fa-check-circle", "fa-exclamation-triangle") %>"></i>
            <%= saveMessage %>
        </div>
        <% End If %>
        
        <!-- Tab导航 -->
        <div class="tab-nav">
            <button type="button" class="tab-btn <%= IIf(currentTab="1", "active", "") %>" onclick="switchTab('1')">
                <i class="fas fa-edit"></i> 月度预算编制
            </button>
            <button type="button" class="tab-btn <%= IIf(currentTab="2", "active", "") %>" onclick="switchTab('2')">
                <i class="fas fa-tachometer-alt"></i> 费用执行率
            </button>
            <button type="button" class="tab-btn <%= IIf(currentTab="3", "active", "") %>" onclick="switchTab('3')">
                <i class="fas fa-water"></i> 现金流预测
            </button>
            <button type="button" class="tab-btn <%= IIf(currentTab="4", "active", "") %>" onclick="switchTab('4')">
                <i class="fas fa-file-invoice-dollar"></i> 预算汇总报表
            </button>
        </div>
        
        <!-- Tab 1: 月度预算编制 -->
        <div id="tab1" class="tab-content <%= IIf(currentTab="1", "active", "") %>">
            <div class="budget-card">
                <div class="card-header">
                    <h3 class="card-title"><i class="fas fa-calendar-alt" style="color: var(--accent-blue);"></i> 月度预算编制</h3>
                    <% If canEdit Then %>
                    <button type="button" class="admin-btn admin-btn-primary" onclick="saveBudget()">
                        <i class="fas fa-save"></i> 批量保存
                    </button>
                    <% End If %>
                </div>
                
                <form id="budgetForm" method="post" action="?tab=1&month=<%= Server.URLEncode(currentMonth) %>">
                    <input type="hidden" name="action" value="save_budget">
                    <div class="month-selector">
                        <label style="color: var(--text-secondary);">预算月份：</label>
                        <input type="month" name="budget_month" class="month-input" value="<%= Server.HTMLEncode(currentMonth) %>" onchange="changeMonth(this.value)">
                    </div>
                    
                    <table class="budget-table">
                        <thead>
                            <tr>
                                <th>预算类目</th>
                                <th>预算金额</th>
                                <th>上月实际</th>
                                <th>建议值</th>
                                <th>状态</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% 
                            ' budgetData 已在顶部预计算块中创建，此处直接使用
                            Dim j, catKey, catLabel, budgetVal, lastActual, suggestion
                            For j = 0 To UBound(budgetCategories)
                                catKey = budgetCategories(j)
                                catLabel = categoryLabels(j)
                                
                                If budgetData.Exists(catKey) Then
                                    budgetVal = budgetData(catKey)
                                Else
                                    budgetVal = 0
                                End If
                                
                                ' 模拟上月实际和建议值
                                lastActual = 0
                                suggestion = Round(budgetVal * 0.95, 2)
                            %>
                            <tr>
                                <td>
                                    <i class="fas fa-tag" style="color: var(--accent-purple); margin-right: 8px;"></i>
                                    <%= catLabel %>
                                </td>
                                <td>
                                    <input type="number" name="budget_<%= catKey %>" class="budget-input" 
                                           value="<%= budgetVal %>" step="0.01" min="0"
                                           <%= IIf(canEdit, "", "disabled") %>>
                                </td>
                                <td style="color: var(--text-secondary);">¥<%= FormatNumber(lastActual, 2) %></td>
                                <td style="color: var(--accent-green);">¥<%= FormatNumber(suggestion, 2) %></td>
                                <td>
                                    <% If budgetVal > 0 Then %>
                                    <span style="color: var(--accent-green);"><i class="fas fa-check-circle"></i> 已设定</span>
                                    <% Else %>
                                    <span style="color: var(--text-secondary);"><i class="fas fa-circle"></i> 未设定</span>
                                    <% End If %>
                                </td>
                            </tr>
                            <% Next %>
                        </tbody>
                        <tfoot>
                            <tr style="background: rgba(102, 126, 234, 0.1);">
                                <td style="font-weight: bold;">合计</td>
                                <td style="font-weight: bold; color: var(--accent-blue);">
                                    ¥<%= FormatNumber(monthBudgetTotal, 2) %>
                                </td>
                                <td colspan="3"></td>
                            </tr>
                        </tfoot>
                    </table>
                </form>
            </div>
        </div>
        
        <!-- Tab 2: 费用执行率看板 -->
        <div id="tab2" class="tab-content <%= IIf(currentTab="2", "active", "") %>">
            <!-- 统计卡片 -->
            <div class="stats-row">
                <div class="stat-box">
                    <div class="stat-box-icon blue"><i class="fas fa-wallet"></i></div>
                    <div class="stat-box-value">¥<%= FormatNumber(monthBudgetTotal, 0) %></div>
                    <div class="stat-box-label">本月预算</div>
                </div>
                <div class="stat-box">
                    <div class="stat-box-icon yellow"><i class="fas fa-coins"></i></div>
                    <div class="stat-box-value" id="totalActual">¥0</div>
                    <div class="stat-box-label">实际消耗</div>
                </div>
                <div class="stat-box">
                    <div class="stat-box-icon green"><i class="fas fa-chart-line"></i></div>
                    <div class="stat-box-value">¥<%= FormatNumber(monthGMV, 0) %></div>
                    <div class="stat-box-label">产出GMV</div>
                </div>
                <div class="stat-box">
                    <div class="stat-box-icon purple"><i class="fas fa-percentage"></i></div>
                    <div class="stat-box-value" id="overallROI">0%</div>
                    <div class="stat-box-label">综合ROI</div>
                </div>
            </div>
            
            <div class="budget-card">
                <div class="card-header">
                    <h3 class="card-title"><i class="fas fa-tachometer-alt" style="color: var(--accent-yellow);"></i> 费用执行率看板</h3>
                    <span style="color: var(--text-secondary); font-size: 13px;">
                        数据月份：<%= currentMonth %>
                    </span>
                </div>
                
                <table class="budget-table">
                    <thead>
                        <tr>
                            <th>类目</th>
                            <th>预算金额</th>
                            <th>实际消耗</th>
                            <th>执行率</th>
                            <th>产出GMV</th>
                            <th>ROI</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% 
                        Dim actualData, totalActual, totalBudgetForROI
                        Set actualData = CreateObject("Scripting.Dictionary")
                        If Not rsActualExpense Is Nothing Then
                            Do While Not rsActualExpense.EOF
                                actualData.Add rsActualExpense("ExpenseType"), rsActualExpense("TotalAmount")
                                rsActualExpense.MoveNext
                            Loop
                            rsActualExpense.Close
                        End If
                        
                        totalActual = 0
                        totalBudgetForROI = 0
                        
                        Dim k, execRate, roiVal, actualAmt, budgetAmtForCat
                        Dim hasAlert : hasAlert = False
                        
                        For k = 0 To UBound(budgetCategories)
                            catKey = budgetCategories(k)
                            catLabel = categoryLabels(k)
                            
                            ' 获取预算
                            If budgetData.Exists(catKey) Then
                                budgetAmtForCat = CDbl(budgetData(catKey))
                            Else
                                budgetAmtForCat = 0
                            End If
                            
                            ' 获取实际消耗
                            If actualData.Exists(catKey) Then
                                actualAmt = CDbl(actualData(catKey))
                            Else
                                actualAmt = 0
                            End If
                            
                            totalActual = totalActual + actualAmt
                            If budgetAmtForCat > 0 Then
                                totalBudgetForROI = totalBudgetForROI + budgetAmtForCat
                            End If
                            
                            ' 计算执行率
                            If budgetAmtForCat > 0 Then
                                execRate = Round((actualAmt / budgetAmtForCat) * 100, 1)
                            Else
                                execRate = 0
                            End If
                            
                            ' 计算ROI（简化计算：GMV按比例分配）
                            If actualAmt > 0 Then
                                Dim allocatedGMV
                                If monthBudgetTotal > 0 Then
                                    allocatedGMV = monthGMV * (budgetAmtForCat / monthBudgetTotal)
                                Else
                                    allocatedGMV = 0
                                End If
                                roiVal = Round((allocatedGMV - actualAmt) / actualAmt * 100, 1)
                            Else
                                allocatedGMV = 0
                                roiVal = 0
                            End If
                            
                            ' 检查预警条件
                            Dim alertClass : alertClass = ""
                            If execRate >= 90 AND roiVal < 20 Then
                                hasAlert = True
                                alertClass = "style='background: rgba(255, 82, 82, 0.1);'"
                            End If
                            
                            ' 进度条颜色
                            Dim progressClass
                            If execRate < 70 Then
                                progressClass = "blue"
                            ElseIf execRate < 90 Then
                                progressClass = "green"
                            ElseIf execRate <= 100 Then
                                progressClass = "yellow"
                            Else
                                progressClass = "red"
                            End If
                        %>
                        <tr <%= alertClass %>>
                            <td><%= catLabel %></td>
                            <td>¥<%= FormatNumber(budgetAmtForCat, 2) %></td>
                            <td>¥<%= FormatNumber(actualAmt, 2) %></td>
                            <td>
                                <div class="progress-container">
                                    <div class="progress-fill <%= progressClass %>" style="width: <%= IIf(execRate > 100, 100, execRate) %>%"></div>
                                    <span class="progress-text"><%= execRate %>%</span>
                                </div>
                            </td>
                            <td>¥<%= FormatNumber(allocatedGMV, 2) %></td>
                            <td style="color: <%= IIf(roiVal >= 0, "#69f0ae", "#ff5252") %>">
                                <%= roiVal %>%
                            </td>
                        </tr>
                        <% Next %>
                    </tbody>
                    <tfoot>
                        <tr style="background: rgba(102, 126, 234, 0.1);">
                            <td style="font-weight: bold;">合计</td>
                            <td style="font-weight: bold;">¥<%= FormatNumber(monthBudgetTotal, 2) %></td>
                            <td style="font-weight: bold;">¥<%= FormatNumber(totalActual, 2) %></td>
                            <td>
                                <div class="progress-container">
                                    <% 
                                    Dim totalExecRate
                                    If monthBudgetTotal > 0 Then
                                        totalExecRate = Round((totalActual / monthBudgetTotal) * 100, 1)
                                    Else
                                        totalExecRate = 0
                                    End If
                                    %>
                                    <div class="progress-fill <%= IIf(totalExecRate > 100, "red", IIf(totalExecRate > 90, "yellow", "green")) %>" 
                                         style="width: <%= IIf(totalExecRate > 100, 100, totalExecRate) %>%"></div>
                                    <span class="progress-text"><%= totalExecRate %>%</span>
                                </div>
                            </td>
                            <td style="font-weight: bold;">¥<%= FormatNumber(monthGMV, 2) %></td>
                            <td style="font-weight: bold;">
                                <% 
                                Dim totalROI
                                If totalActual > 0 Then
                                    totalROI = Round((monthGMV - totalActual) / totalActual * 100, 1)
                                Else
                                    totalROI = 0
                                End If
                                %>
                                <span style="color: <%= IIf(totalROI >= 0, "#69f0ae", "#ff5252") %>"><%= totalROI %>%</span>
                            </td>
                        </tr>
                    </tfoot>
                </table>
                
                <% If hasAlert Then %>
                <div class="alert-box" style="margin-top: 20px;">
                    <i class="fas fa-exclamation-triangle"></i>
                    <div>
                        <strong>预算超支预警</strong><br>
                        <span style="font-size: 13px;">检测到部分类目执行率已达90%且ROI低于20%，请关注资金使用情况。</span>
                    </div>
                </div>
                <% End If %>
            </div>
        </div>
        
        <!-- Tab 3: 现金流滚动预测 -->
        <div id="tab3" class="tab-content <%= IIf(currentTab="3", "active", "") %>">
            <div class="stats-row">
                <div class="stat-box">
                    <div class="stat-box-icon blue"><i class="fas fa-piggy-bank"></i></div>
                    <div class="stat-box-value">¥<%= FormatNumber(openingBalance, 0) %></div>
                    <div class="stat-box-label">期初余额</div>
                </div>
                <div class="stat-box">
                    <div class="stat-box-icon green"><i class="fas fa-arrow-down"></i></div>
                    <div class="stat-box-value">¥<%= FormatNumber(expectedCollection, 0) %></div>
                    <div class="stat-box-label">预计回款(7天)</div>
                </div>
                <div class="stat-box">
                    <div class="stat-box-icon yellow"><i class="fas fa-arrow-up"></i></div>
                    <div class="stat-box-value" id="projectedExpense">计算中...</div>
                    <div class="stat-box-label">预计支出(30天)</div>
                </div>
                <div class="stat-box">
                    <div class="stat-box-icon purple"><i class="fas fa-chart-area"></i></div>
                    <div class="stat-box-value" id="projectedBalance">计算中...</div>
                    <div class="stat-box-label">期末预测</div>
                </div>
            </div>
            
            <div class="budget-card">
                <div class="card-header">
                    <h3 class="card-title"><i class="fas fa-water" style="color: var(--accent-blue);"></i> 未来30天资金水位预测</h3>
                    <div style="display: flex; align-items: center; gap: 20px;">
                        <span style="display: flex; align-items: center; gap: 5px; font-size: 13px; color: var(--text-secondary);">
                            <span style="width: 20px; height: 3px; background: var(--accent-blue); display: inline-block;"></span>
                            预测余额
                        </span>
                        <span style="display: flex; align-items: center; gap: 5px; font-size: 13px; color: var(--text-secondary);">
                            <span style="width: 20px; height: 3px; background: var(--accent-red); border-top: 2px dashed; display: inline-block;"></span>
                            安全线
                        </span>
                    </div>
                </div>
                
                <div class="cashflow-chart" id="cashflowChart">
                    <svg width="100%" height="100%" viewBox="0 0 800 250" preserveAspectRatio="none">
                        <defs>
                            <linearGradient id="gradientBlue" x1="0%" y1="0%" x2="0%" y2="100%">
                                <stop offset="0%" style="stop-color:#4fc3f7;stop-opacity:0.6" />
                                <stop offset="100%" style="stop-color:#4fc3f7;stop-opacity:0.1" />
                            </linearGradient>
                        </defs>
                        
                        <!-- 网格线 -->
                        <line x1="50" y1="20" x2="50" y2="200" stroke="#3a4055" stroke-width="1"/>
                        <line x1="50" y1="200" x2="775" y2="200" stroke="#3a4055" stroke-width="1"/>
                        
                        <!-- Y轴标签 -->
                        <text x="40" y="205" fill="#9fa8da" font-size="10" text-anchor="end">0</text>
                        <text x="40" y="20" fill="#9fa8da" font-size="10" text-anchor="end">¥<%= FormatNumber(chartMaxVal/1000, 0) %>k</text>
                        
                        <!-- 安全线 -->
                        <line x1="50" y1="<%= chartSafeLineY %>" x2="775" y2="<%= chartSafeLineY %>" 
                              stroke="#ff5252" stroke-width="2" stroke-dasharray="5,5"/>
                        
                        <!-- 面积图 -->
                        <path d="<%= chartAreaPath %>" fill="url(#gradientBlue)"/>
                        
                        <!-- 折线 -->
                        <polyline points="<%= chartPoints %>" fill="none" stroke="#4fc3f7" stroke-width="2"/>
                        
                        <!-- 数据点 -->
                        <% 
                        Dim svgPtIdx, svgPtX, svgPtY
                        For svgPtIdx = 0 To 29 Step 5
                            svgPtX = 50 + (svgPtIdx * 25)
                            svgPtY = 200 - ((chartDailyData(svgPtIdx) - chartMinVal) / (chartMaxVal - chartMinVal) * 180)
                        %>
                        <circle cx="<%= svgPtX %>" cy="<%= svgPtY %>" r="3" fill="#4fc3f7"/>
                        <text x="<%= svgPtX %>" y="220" fill="#9fa8da" font-size="9" text-anchor="middle">D<%= svgPtIdx + 1 %></text>
                        <% Next %>
                    </svg>
                </div>
                
                <div style="margin-top: 20px; padding: 15px; background: rgba(79, 195, 247, 0.1); border-radius: 8px;">
                    <h4 style="margin: 0 0 10px 0; color: var(--accent-blue);"><i class="fas fa-info-circle"></i> 预测说明</h4>
                    <ul style="margin: 0; padding-left: 20px; color: var(--text-secondary); font-size: 13px; line-height: 1.8;">
                        <li><strong>期初余额：</strong>当前所有资金账户可用余额总和</li>
                        <li><strong>预计回款：</strong>基于最近7天已付款订单金额（假设T+7结算周期）</li>
                        <li><strong>预计采购：</strong>本月采购预算按日均分摊</li>
                        <li><strong>预计固定开支：</strong>人工+平台费用按日均分摊</li>
                        <li><strong>安全线：</strong>固定开支的2倍作为最低资金储备</li>
                    </ul>
                </div>
            </div>
        </div>
        
        <!-- Tab 4: 预算汇总报表 -->
        <div id="tab4" class="tab-content <%= IIf(currentTab="4", "active", "") %>">
            <div class="summary-grid">
                <div class="budget-card">
                    <div class="card-header">
                        <h3 class="card-title"><i class="fas fa-calendar" style="color: var(--accent-green);"></i> <%= currentYear %>年 季度预算汇总</h3>
                    </div>
                    <table class="budget-table">
                        <thead>
                            <tr>
                                <th>季度</th>
                                <th>预算总额</th>
                                <th>实际支出</th>
                                <th>执行率</th>
                                <th>同比</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% 
                            Dim q, qBudget, qActual, qExecRate
                            For q = 1 To 4
                                ' 季度预算
                                qBudget = GetScalar("SELECT IIF(SUM(BudgetAmount) IS NULL, 0, SUM(BudgetAmount)) FROM BudgetPlans WHERE Year(BudgetMonth) = " & currentYear & " AND Month(BudgetMonth) BETWEEN " & ((q-1)*3+1) & " AND " & (q*3))
                                ' 季度实际
                                qActual = GetScalar("SELECT IIF(SUM(Amount) IS NULL, 0, SUM(Amount)) FROM ExpenseRecords WHERE Year(ExpenseMonth) = " & currentYear & " AND Month(ExpenseMonth) BETWEEN " & ((q-1)*3+1) & " AND " & (q*3))
                                ' 执行率
                                If qBudget > 0 Then
                                    qExecRate = Round((qActual / qBudget) * 100, 1)
                                Else
                                    qExecRate = 0
                                End If
                            %>
                            <tr>
                                <td>第<%= q %>季度</td>
                                <td>¥<%= FormatNumber(qBudget, 2) %></td>
                                <td>¥<%= FormatNumber(qActual, 2) %></td>
                                <td>
                                    <div class="progress-container" style="width: 100px;">
                                        <div class="progress-fill <%= IIf(qExecRate > 100, "red", IIf(qExecRate > 90, "yellow", "green")) %>" 
                                             style="width: <%= IIf(qExecRate > 100, 100, qExecRate) %>%"></div>
                                        <span class="progress-text"><%= qExecRate %>%</span>
                                    </div>
                                </td>
                                <td>
                                    <% If q < currentQuarter Then %>
                                    <span class="trend-indicator trend-up">
                                        <i class="fas fa-arrow-up"></i> --
                                    </span>
                                    <% Else %>
                                    <span class="trend-indicator" style="color: var(--text-secondary);">-</span>
                                    <% End If %>
                                </td>
                            </tr>
                            <% Next %>
                        </tbody>
                    </table>
                </div>
                
                <div class="budget-card">
                    <div class="card-header">
                        <h3 class="card-title"><i class="fas fa-tags" style="color: var(--accent-purple);"></i> 年度类目汇总</h3>
                    </div>
                    <table class="budget-table">
                        <thead>
                            <tr>
                                <th>类目</th>
                                <th>年度预算</th>
                                <th>已执行</th>
                                <th>剩余预算</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% 
                            Dim yearBudget, yearActual, yearRemaining
                            For k = 0 To UBound(budgetCategories)
                                catKey = budgetCategories(k)
                                catLabel = categoryLabels(k)
                                
                                yearBudget = GetScalar("SELECT IIF(SUM(BudgetAmount) IS NULL, 0, SUM(BudgetAmount)) FROM BudgetPlans WHERE Year(BudgetMonth) = " & currentYear & " AND Category = '" & catKey & "'")
                                yearActual = GetScalar("SELECT IIF(SUM(Amount) IS NULL, 0, SUM(Amount)) FROM ExpenseRecords WHERE Year(ExpenseMonth) = " & currentYear & " AND ExpenseType = '" & catKey & "'")
                                yearRemaining = yearBudget - yearActual
                            %>
                            <tr>
                                <td><%= catLabel %></td>
                                <td>¥<%= FormatNumber(yearBudget, 2) %></td>
                                <td style="color: <%= IIf(yearActual > yearBudget, "#ff5252", "#69f0ae") %>">¥<%= FormatNumber(yearActual, 2) %></td>
                                <td style="color: <%= IIf(yearRemaining < 0, "#ff5252", "var(--text-primary)") %>">¥<%= FormatNumber(yearRemaining, 2) %></td>
                            </tr>
                            <% Next %>
                        </tbody>
                        <tfoot>
                            <tr style="background: rgba(102, 126, 234, 0.1);">
                                <td style="font-weight: bold;">年度合计</td>
                                <td style="font-weight: bold;">
                                    ¥<%= FormatNumber(CDbl("0" & GetScalar("SELECT IIF(SUM(BudgetAmount) IS NULL, 0, SUM(BudgetAmount)) FROM BudgetPlans WHERE Year(BudgetMonth) = " & currentYear)), 2) %>
                                </td>
                                <td style="font-weight: bold;">
                                    ¥<%= FormatNumber(CDbl("0" & GetScalar("SELECT IIF(SUM(Amount) IS NULL, 0, SUM(Amount)) FROM ExpenseRecords WHERE Year(ExpenseMonth) = " & currentYear)), 2) %>
                                </td>
                                <td style="font-weight: bold;">
                                    <% 
                                    Dim totalYearBudget, totalYearActual, totalRemaining
                                    totalYearBudget = SafeCDbl(GetScalar("SELECT IIF(SUM(BudgetAmount) IS NULL, 0, SUM(BudgetAmount)) FROM BudgetPlans WHERE Year(BudgetMonth) = " & currentYear))
                                    totalYearActual = SafeCDbl(GetScalar("SELECT IIF(SUM(Amount) IS NULL, 0, SUM(Amount)) FROM ExpenseRecords WHERE Year(ExpenseMonth) = " & currentYear))
                                    totalRemaining = totalYearBudget - totalYearActual
                                    %>
                                    <span style="color: <%= IIf(totalRemaining < 0, "#ff5252", "#69f0ae") %>">¥<%= FormatNumber(totalRemaining, 2) %></span>
                                </td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
            
            <div class="budget-card" style="margin-top: 25px;">
                <div class="card-header">
                    <h3 class="card-title"><i class="fas fa-chart-bar" style="color: var(--accent-blue);"></i> 预算执行趋势分析</h3>
                </div>
                <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px;">
                    <div style="text-align: center; padding: 20px; background: var(--dark-bg); border-radius: 8px;">
                        <div style="font-size: 32px; font-weight: bold; color: var(--accent-blue); margin-bottom: 10px;">
                            <%= Round((totalYearActual / IIf(totalYearBudget > 0, totalYearBudget, 1)) * 100, 1) %>%
                        </div>
                        <div style="color: var(--text-secondary); font-size: 14px;">年度预算执行率</div>
                    </div>
                    <div style="text-align: center; padding: 20px; background: var(--dark-bg); border-radius: 8px;">
                        <div style="font-size: 32px; font-weight: bold; color: var(--accent-green); margin-bottom: 10px;">
                            <%= currentQuarter %>
                        </div>
                        <div style="color: var(--text-secondary); font-size: 14px;">当前季度</div>
                    </div>
                    <div style="text-align: center; padding: 20px; background: var(--dark-bg); border-radius: 8px;">
                        <div style="font-size: 32px; font-weight: bold; color: var(--accent-purple); margin-bottom: 10px;">
                            <%= Month(Date()) %>
                        </div>
                        <div style="color: var(--text-secondary); font-size: 14px;">当前月份</div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    
    <!-- 预警模态框 -->
    <div id="alertModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3 class="modal-title"><i class="fas fa-exclamation-triangle" style="color: var(--accent-red);"></i> 预算超支预警</h3>
                <button type="button" class="close-btn" onclick="closeModal()">&times;</button>
            </div>
            <div class="modal-body">
                <p>检测到以下预算类目存在异常：</p>
                <ul id="alertList" style="line-height: 2;">
                    <li>执行率超过90%且ROI低于20%</li>
                </ul>
                <p style="margin-top: 15px; color: var(--text-secondary);">建议立即审查相关支出，优化资金配置。</p>
            </div>
            <div class="modal-footer">
                <button type="button" class="admin-btn admin-btn-outline" onclick="closeModal()">稍后处理</button>
                <button type="button" class="admin-btn admin-btn-primary" onclick="switchTab('2'); closeModal();">查看详情</button>
            </div>
        </div>
    </div>

    <script>
        // 切换Tab（保留月份参数）
        function switchTab(tabId) {
            var params = new URLSearchParams(window.location.search);
            params.set('tab', tabId);
            var newUrl = window.location.pathname + '?' + params.toString();
            window.location.href = newUrl;
        }
        
        // 切换月份（保留tab参数）
        function changeMonth(month) {
            var params = new URLSearchParams(window.location.search);
            params.set('month', month);
            var newUrl = window.location.pathname + '?' + params.toString();
            window.location.href = newUrl;
        }
        
        // 保存预算
        function saveBudget() {
            var form = document.getElementById('budgetForm');
            if (form) form.submit();
        }
        
        // 模态框控制
        function showModal() {
            var modal = document.getElementById('alertModal');
            if (modal) modal.style.display = 'block';
        }
        
        function closeModal() {
            var modal = document.getElementById('alertModal');
            if (modal) modal.style.display = 'none';
        }
        
        // 点击模态框外部关闭（使用addEventListener避免覆盖其他点击处理）
        document.addEventListener('click', function(event) {
            var modal = document.getElementById('alertModal');
            if (modal && event.target == modal) {
                modal.style.display = 'none';
            }
        });
        
        // 页面加载完成后更新统计值和现金流预测值
        document.addEventListener('DOMContentLoaded', function() {
            <% If currentTab = "2" Then %>
            var elActual = document.getElementById('totalActual');
            var elROI = document.getElementById('overallROI');
            if (elActual) elActual.textContent = '¥<%= FormatNumber(totalActual, 0) %>';
            if (elROI) elROI.textContent = '<%= totalROI %>%';
            <% End If %>
            
            <% If currentTab = "3" Then %>
            <% 
            Dim projectedExpense30, projectedEnding
            projectedExpense30 = (chartPurchaseDaily + chartFixedDaily) * 30
            projectedEnding = openingBalance + expectedCollection - projectedExpense30
            %>
            var elExpense = document.getElementById('projectedExpense');
            var elBalance = document.getElementById('projectedBalance');
            if (elExpense) elExpense.textContent = '¥<%= FormatNumber(projectedExpense30, 0) %>';
            if (elBalance) elBalance.textContent = '¥<%= FormatNumber(projectedEnding, 0) %>';
            <% End If %>
        });
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
