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
    If SafeNum(denominator) = 0 Then
        SafeDiv = 0
    Else
        SafeDiv = SafeNum(numerator) / SafeNum(denominator)
    End If
End Function

Call OpenConnection()

' 权限检查 - FIN_MANAGER 和 FIN_STAFF 均可查看
Dim canEditCampaign
If Session("AdminRoleCode") = "FIN_MANAGER" OR Session("AdminRoleCode") = "SUPER_ADMIN" Then
    canEditCampaign = True
Else
    canEditCampaign = False
End If

' 获取当前Tab
Dim currentTab
currentTab = Trim(Request.QueryString("tab"))
If currentTab = "" Then currentTab = "channel"

' 日期范围处理
Dim startDate, endDate, cpName, cpChannel, cpStart, cpEnd, cpBudget, cpSpent, cpRevenue, cpOrders, cpClicks, cpROI, cpConv
Dim cpStatus, cpStatusClass, cpStatusName, roiClass, aovYear, aovMonth, aovCount, aovRev, aovAvg
Dim bgPeriod, bgBudget, bgActual, bgGMV, bgROI, bgRate, bgStatus, bgStatusClass
startDate = Trim(Request.QueryString("startDate"))
endDate = Trim(Request.QueryString("endDate"))
If startDate = "" Then startDate = SafeFormatDateTime(DateAdd("m", -6, Date()), 2)
If endDate = "" Then endDate = SafeFormatDateTime(Date(), 2)

Dim safeStart, safeEnd
safeStart = SafeSQL(startDate)
safeEnd = SafeSQL(endDate)

' ============================================
' Tab 1: 渠道效果统计
' ============================================
Dim rsChannel, totalChannelRevenue, totalChannelOrders
totalChannelRevenue = GetScalar("SELECT CAST(IIF(SUM(TotalAmount) IS NULL, 0, SUM(TotalAmount)) AS FLOAT) FROM Orders WHERE OrderDate >= '" & safeStart & "' AND OrderDate <= '" & safeEnd & "'")
totalChannelOrders = GetScalar("SELECT COUNT(*) FROM Orders WHERE OrderDate >= '" & safeStart & "' AND OrderDate <= '" & safeEnd & "'")
If IsNull(totalChannelRevenue) Then totalChannelRevenue = 0
If IsNull(totalChannelOrders) Then totalChannelOrders = 0

' ============================================
' Tab 2: 营销活动数据
' ============================================
Dim rsCampaigns
' 使用字段别名兼容运营模块的MarketingCampaigns表结构
' 运营模块字段: CampaignType, DiscountValue, MinPurchase, TotalSales, ParticipantCount, IsActive
' 映射到财务字段: ChannelSource, BudgetAmount, SpentAmount, Revenue, OrderCount, Status
On Error Resume Next
Set rsCampaigns = ExecuteQuery(_
    "SELECT CampaignID, CampaignName, " & _
    "CampaignType AS ChannelSource, " & _
    "StartDate, EndDate, " & _
    "DiscountValue AS BudgetAmount, " & _
    "MinPurchase AS SpentAmount, " & _
    "TotalSales AS Revenue, " & _
    "ParticipantCount AS OrderCount, " & _
    "0 AS ClickCount, " & _
    "IIF(IsActive=1, 'Active', 'Ended') AS Status " & _
    "FROM MarketingCampaigns ORDER BY StartDate DESC")
If Err.Number <> 0 Then
    Err.Clear
    Set rsCampaigns = ExecuteQuery("SELECT * FROM MarketingCampaigns ORDER BY StartDate DESC")
End If
On Error GoTo 0

' ============================================
' Tab 3: 用户消费行为分析
' ============================================
' 总用户数
Dim totalUsers
totalUsers = GetScalar("SELECT COUNT(*) FROM Users")

' 有2+订单的用户数（复购用户）- 使用子查询替代COUNT(DISTINCT)
Dim repeatUsers
repeatUsers = GetScalar("SELECT COUNT(*) FROM (SELECT UserID FROM Orders WHERE UserID IS NOT NULL GROUP BY UserID HAVING COUNT(*) >= 2) AS T")

' 复购率
Dim repeatRate
If CDbl(totalUsers) > 0 Then
    repeatRate = Round((CDbl(repeatUsers) / CDbl(totalUsers)) * 100, 1)
Else
    repeatRate = 0
End If

' 最近6个月客单价趋势
Dim rsAvgOrderValue
Set rsAvgOrderValue = ExecuteQuery(_
    "SELECT Year(OrderDate) AS Y, Month(OrderDate) AS M, COUNT(*) AS OrderCount, SUM(CAST(TotalAmount AS FLOAT)) AS Revenue " & _
    "FROM Orders WHERE OrderDate >= DATEADD(month, -6, CAST(GETDATE() AS DATE)) " & _
    "GROUP BY Year(OrderDate), Month(OrderDate) " & _
    "ORDER BY Year(OrderDate) DESC, Month(OrderDate) DESC")

' 消费频次分布
Dim freq1, freq2_3, freq4_5, freq6plus
freq1 = GetScalar("SELECT COUNT(*) FROM (SELECT UserID FROM Orders WHERE UserID IS NOT NULL GROUP BY UserID HAVING COUNT(*) = 1) AS T")
freq2_3 = GetScalar("SELECT COUNT(*) FROM (SELECT UserID FROM Orders WHERE UserID IS NOT NULL GROUP BY UserID HAVING COUNT(*) >= 2 AND COUNT(*) <= 3) AS T")
freq4_5 = GetScalar("SELECT COUNT(*) FROM (SELECT UserID FROM Orders WHERE UserID IS NOT NULL GROUP BY UserID HAVING COUNT(*) >= 4 AND COUNT(*) <= 5) AS T")
freq6plus = GetScalar("SELECT COUNT(*) FROM (SELECT UserID FROM Orders WHERE UserID IS NOT NULL GROUP BY UserID HAVING COUNT(*) >= 6) AS T")

Dim totalFreqUsers
If CDbl(freq1) + CDbl(freq2_3) + CDbl(freq4_5) + CDbl(freq6plus) > 0 Then
    totalFreqUsers = CDbl(freq1) + CDbl(freq2_3) + CDbl(freq4_5) + CDbl(freq6plus)
Else
    totalFreqUsers = 1
End If

' ============================================
' Tab 4: 转化率漏斗
' ============================================
' 注册用户数
Dim funnelRegister
funnelRegister = GetScalar("SELECT COUNT(*) FROM Users")

' 下单用户数（去重）
Dim funnelOrderUsers
funnelOrderUsers = GetScalar("SELECT COUNT(*) FROM (SELECT DISTINCT UserID FROM Orders WHERE UserID IS NOT NULL) AS T")

' 付款用户数（去重）
Dim funnelPaidUsers
funnelPaidUsers = GetScalar("SELECT COUNT(*) FROM (SELECT DISTINCT UserID FROM Orders WHERE UserID IS NOT NULL AND Status IN ('Paid','Completed')) AS T")

' 复购用户数（有2+订单的用户）
Dim funnelRepeatUsers
funnelRepeatUsers = GetScalar("SELECT COUNT(*) FROM (SELECT UserID FROM Orders WHERE UserID IS NOT NULL GROUP BY UserID HAVING COUNT(*) >= 2) AS T")

' ============================================
' Tab 5: 营销费用与预算对比
' ============================================
Dim rsBudget
Set rsBudget = ExecuteQuery(_
    "SELECT Period, BudgetAmount, ActualAmount, GMVAmount, ROI " & _
    "FROM BudgetPlans WHERE Category='Promotion' ORDER BY Period DESC")

' 计算总预算和总实际支出
Dim totalBudget, totalActual, totalGMV
totalBudget = GetScalar("SELECT IIF(SUM(BudgetAmount) IS NULL, 0, SUM(BudgetAmount)) FROM BudgetPlans WHERE Category='Promotion'")
totalActual = GetScalar("SELECT IIF(SUM(ActualAmount) IS NULL, 0, SUM(ActualAmount)) FROM BudgetPlans WHERE Category='Promotion'")
totalGMV = GetScalar("SELECT IIF(SUM(GMVAmount) IS NULL, 0, SUM(GMVAmount)) FROM BudgetPlans WHERE Category='Promotion'")

' 处理新建活动提交
Dim actionMsg
actionMsg = ""
If Request.Form("action") = "create_campaign" AND canEditCampaign Then
    Dim cName, cChannel, cStart, cEnd, cBudget
    cName = Trim(Request.Form("campaignName"))
    cChannel = Trim(Request.Form("channelSource"))
    cStart = Trim(Request.Form("startDate"))
    cEnd = Trim(Request.Form("endDate"))
    cBudget = Trim(Request.Form("budgetAmount"))
    
    If cName <> "" AND cStart <> "" AND cEnd <> "" Then
        Dim sqlInsert
        ' 使用运营模块MarketingCampaigns表的字段结构
        ' CampaignType=渠道, DiscountValue=预算, MinPurchase=花费, TotalSales=收入, ParticipantCount=订单数
        sqlInsert = "INSERT INTO MarketingCampaigns (CampaignName, CampaignType, [Description], StartDate, EndDate, DiscountValue, MinPurchase, TotalSales, ParticipantCount, IsActive, CreatedAt) VALUES (" & _
            "'" & SafeSQL(cName) & "', " & _
            "'" & SafeSQL(cChannel) & "', " & _
            "'营销活动', " & _
            "'" & SafeSQL(cStart) & "', " & _
            "'" & SafeSQL(cEnd) & "', " & _
            CDbl("0" & cBudget) & ", 0, 0, 0, True, GETDATE())"
        
        If ExecuteNonQuery(sqlInsert) Then
            actionMsg = "活动创建成功！"
        Else
            actionMsg = "活动创建失败！"
        End If
    End If
    
    Response.Redirect "marketing_stats.asp?tab=campaigns&msg=" & Server.URLEncode(actionMsg)
    Response.End
End If

Call LogAdminAction("查看营销数据统计", "finance", "MarketingCampaigns", "", safeStart & " 至 " & safeEnd)
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>营销数据统计 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <style>
        /* 暗色主题覆盖 */
        body { background: #1a1a2e; color: #e0e0e0; }
        .main-content { background: #1a1a2e; }
        .page-title { color: #fff; }
        .breadcrumb { color: #888; }
        .breadcrumb a { color: #00bcd4; }
        
        /* Tab 样式 */
        .stats-tabs { display: flex; gap: 10px; margin-bottom: 25px; border-bottom: 2px solid #2d3142; padding-bottom: 15px; }
        .stats-tab { 
            padding: 12px 24px; 
            background: #2d2d44; 
            border-radius: 8px; 
            cursor: pointer; 
            color: #888; 
            transition: all 0.3s;
            border: none;
            font-size: 14px;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .stats-tab:hover { background: #1e1e32; color: #fff; }
        .stats-tab.active { 
            background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); 
            color: white; 
        }
        
        /* 筛选栏 */
        .filter-bar { 
            background: #2d2d44; 
            padding: 20px; 
            border-radius: 12px; 
            margin-bottom: 25px; 
            box-shadow: 0 4px 20px rgba(0,0,0,0.3); 
        }
        .filter-bar form { display: flex; gap: 15px; align-items: center; flex-wrap: wrap; }
        .filter-bar label { color: #aaa; }
        .filter-bar input { 
            padding: 10px 15px; 
            background: #1a1a2e; 
            border: 1px solid rgba(255,255,255,0.15); 
            border-radius: 8px; 
            color: #fff;
        }
        .filter-bar input:focus { border-color: #00bcd4; outline: none; }
        
        /* 卡片样式 */
        .stat-card { 
            background: #2d2d44; 
            border-radius: 12px; 
            padding: 25px; 
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            transition: transform 0.3s;
        }
        .stat-card:hover { transform: translateY(-3px); }
        .stat-card-header { 
            display: flex; 
            align-items: center; 
            gap: 12px; 
            margin-bottom: 15px;
            color: #888;
            font-size: 14px;
        }
        .stat-card-value { 
            font-size: 32px; 
            font-weight: bold; 
            color: #fff;
        }
        .stat-card-value.green { color: #4ade80; }
        .stat-card-value.blue { color: #60a5fa; }
        .stat-card-value.orange { color: #fb923c; }
        .stat-card-value.purple { color: #a78bfa; }
        
        /* 数据表格 */
        .data-table { 
            width: 100%; 
            border-collapse: collapse; 
            background: #2d2d44;
            border-radius: 12px;
            overflow: hidden;
        }
        .data-table th { 
            background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); 
            color: white; 
            padding: 15px; 
            text-align: left; 
        }
        .data-table td { 
            padding: 12px 15px; 
            border-bottom: 1px solid #3d4252; 
            color: #e0e0e0;
        }
        .data-table tr:hover { background: #2d3142; }
        .data-table tfoot td { 
            font-weight: bold; 
            background: #1a1a2e; 
            color: #fff;
        }
        
        /* 图表容器 */
        .chart-card {
            background: #2d2d44;
            border-radius: 12px;
            padding: 25px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            margin-bottom: 25px;
        }
        .chart-card h3 {
            color: #fff;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        /* 饼图容器 */
        .pie-chart-container {
            display: flex;
            align-items: center;
            gap: 40px;
        }
        .pie-chart {
            width: 200px;
            height: 200px;
            border-radius: 50%;
            background: conic-gradient(
                #00bcd4 0deg var(--p1),
                #4ade80 var(--p1) var(--p2),
                #fb923c var(--p2) var(--p3),
                #a78bfa var(--p3) var(--p4),
                #f472b6 var(--p4) 360deg
            );
        }
        .pie-legend {
            display: flex;
            flex-direction: column;
            gap: 10px;
        }
        .legend-item {
            display: flex;
            align-items: center;
            gap: 10px;
            color: #ccc;
        }
        .legend-color {
            width: 16px;
            height: 16px;
            border-radius: 4px;
        }
        
        /* ROI 颜色编码 */
        .roi-high { color: #4ade80; font-weight: bold; }
        .roi-medium { color: #facc15; font-weight: bold; }
        .roi-low { color: #f87171; font-weight: bold; }
        
        /* 状态标签 */
        .status-badge {
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 500;
        }
        .status-active { background: #064e3b; color: #4ade80; }
        .status-ended { background: #7f1d1d; color: #f87171; }
        .status-ongoing { background: #1e3a8a; color: #60a5fa; }
        .status-overbudget { background: #7c2d12; color: #fb923c; }
        
        /* 漏斗图 */
        .funnel-container {
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 5px;
            padding: 20px;
        }
        .funnel-level {
            display: flex;
            align-items: center;
            justify-content: center;
            height: 60px;
            color: white;
            font-weight: bold;
            position: relative;
            transition: all 0.3s;
        }
        .funnel-level:hover { transform: scale(1.02); }
        .funnel-level-1 { 
            width: 400px; 
            background: linear-gradient(135deg, #00bcd4, #00838f);
            clip-path: polygon(0 0, 100% 0, 95% 100%, 5% 100%);
        }
        .funnel-level-2 { 
            width: 360px; 
            background: linear-gradient(135deg, #4ade80, #22c55e);
            clip-path: polygon(0 0, 100% 0, 94% 100%, 6% 100%);
        }
        .funnel-level-3 { 
            width: 320px; 
            background: linear-gradient(135deg, #fb923c, #ea580c);
            clip-path: polygon(0 0, 100% 0, 93% 100%, 7% 100%);
        }
        .funnel-level-4 { 
            width: 280px; 
            background: linear-gradient(135deg, #a78bfa, #7c3aed);
            clip-path: polygon(0 0, 100% 0, 92% 100%, 8% 100%);
        }
        .funnel-info {
            display: flex;
            flex-direction: column;
            align-items: center;
            line-height: 1.3;
        }
        .funnel-label { font-size: 14px; opacity: 0.9; }
        .funnel-value { font-size: 20px; }
        .funnel-rate {
            position: absolute;
            right: -80px;
            color: #888;
            font-size: 13px;
        }
        
        /* 进度条 */
        .progress-bar { 
            height: 20px; 
            background: #1a1a2e; 
            border-radius: 10px; 
            overflow: hidden; 
            display: inline-block; 
            vertical-align: middle;
            width: 120px;
        }
        .progress-fill { 
            height: 100%; 
            border-radius: 10px; 
            transition: width 0.3s;
        }
        .progress-fill.blue { background: linear-gradient(90deg, #00bcd4, #00838f); }
        .progress-fill.green { background: linear-gradient(90deg, #4ade80, #22c55e); }
        .progress-fill.orange { background: linear-gradient(90deg, #fb923c, #ea580c); }
        .progress-fill.red { background: linear-gradient(90deg, #f87171, #dc2626); }
        
        /* 网格布局 */
        .stats-grid { 
            display: grid; 
            grid-template-columns: repeat(4, 1fr); 
            gap: 20px; 
            margin-bottom: 25px; 
        }
        .charts-grid { 
            display: grid; 
            grid-template-columns: 1fr 1fr; 
            gap: 25px; 
            margin-bottom: 25px; 
        }
        
        /* 无数据提示 */
        .no-data { 
            text-align: center; 
            padding: 40px; 
            color: #666; 
        }
        .no-data i { font-size: 48px; margin-bottom: 15px; display: block; color: #444; }
        
        /* 预警提示 */
        .alert-warning {
            background: rgba(251, 146, 60, 0.1);
            border: 1px solid rgba(251, 146, 60, 0.3);
            color: #fb923c;
            padding: 12px 20px;
            border-radius: 8px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        /* 响应式 */
        @media (max-width: 1200px) {
            .stats-grid { grid-template-columns: repeat(2, 1fr); }
            .charts-grid { grid-template-columns: 1fr; }
        }
        @media (max-width: 768px) {
            .stats-grid { grid-template-columns: 1fr; }
            .stats-tabs { flex-wrap: wrap; }
            .funnel-level-1 { width: 280px; }
            .funnel-level-2 { width: 250px; }
            .funnel-level-3 { width: 220px; }
            .funnel-level-4 { width: 190px; }
        }
        
        /* 隐藏内容 */
        .tab-content { display: none; }
        .tab-content.active { display: block; }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-chart-pie"></i> 营销数据统计</h2>
            <div class="breadcrumb">
                <a href="index.asp">财务中心</a> / <span>营销统计</span>
            </div>
        </div>
        
        <% If Request.QueryString("msg") <> "" Then %>
        <div class="alert-warning">
            <i class="fas fa-info-circle"></i>
            <%= Server.HTMLEncode(Request.QueryString("msg")) %>
        </div>
        <% End If %>
        
        <!-- Tab 导航 -->
        <div class="stats-tabs">
            <a href="?tab=channel&startDate=<%= Server.URLEncode(startDate) %>&endDate=<%= Server.URLEncode(endDate) %>" class="stats-tab <%= IIf(currentTab="channel", "active", "") %>">
                <i class="fas fa-sitemap"></i> 渠道效果
            </a>
            <a href="?tab=campaigns&startDate=<%= Server.URLEncode(startDate) %>&endDate=<%= Server.URLEncode(endDate) %>" class="stats-tab <%= IIf(currentTab="campaigns", "active", "") %>">
                <i class="fas fa-bullhorn"></i> 营销活动ROI
            </a>
            <a href="?tab=behavior&startDate=<%= Server.URLEncode(startDate) %>&endDate=<%= Server.URLEncode(endDate) %>" class="stats-tab <%= IIf(currentTab="behavior", "active", "") %>">
                <i class="fas fa-users"></i> 用户行为
            </a>
            <a href="?tab=funnel&startDate=<%= Server.URLEncode(startDate) %>&endDate=<%= Server.URLEncode(endDate) %>" class="stats-tab <%= IIf(currentTab="funnel", "active", "") %>">
                <i class="fas fa-filter"></i> 转化漏斗
            </a>
            <a href="?tab=budget&startDate=<%= Server.URLEncode(startDate) %>&endDate=<%= Server.URLEncode(endDate) %>" class="stats-tab <%= IIf(currentTab="budget", "active", "") %>">
                <i class="fas fa-wallet"></i> 预算对比
            </a>
        </div>
        
        <!-- 日期筛选 -->
        <div class="filter-bar">
            <form method="get" action="marketing_stats.asp">
                <input type="hidden" name="tab" value="<%= currentTab %>">
                <label><i class="fas fa-calendar"></i> 开始日期:</label>
                <input type="date" name="startDate" value="<%= Server.HTMLEncode(startDate) %>">
                <label><i class="fas fa-calendar"></i> 结束日期:</label>
                <input type="date" name="endDate" value="<%= Server.HTMLEncode(endDate) %>">
                <button type="submit" class="admin-btn admin-btn-primary"><i class="fas fa-search"></i> 查询</button>
                <a href="marketing_stats.asp?tab=<%= currentTab %>" class="admin-btn" style="background:#3d4252;color:#fff;"><i class="fas fa-redo"></i> 重置</a>
            </form>
        </div>
        
        <!-- ============================================
             Tab 1: 渠道效果统计
             ============================================ -->
        <div class="tab-content <%= IIf(currentTab="channel", "active", "") %>">
            <div class="stats-grid">
                <div class="stat-card">
                    <div class="stat-card-header"><i class="fas fa-shopping-cart"></i> 总订单数</div>
                    <div class="stat-card-value blue"><%= totalChannelOrders %></div>
                </div>
                <div class="stat-card">
                    <div class="stat-card-header"><i class="fas fa-yen-sign"></i> 总销售额</div>
                    <div class="stat-card-value green">¥<%= FormatNumber(SafeNum(totalChannelRevenue), 0) %></div>
                </div>
                <div class="stat-card">
                    <div class="stat-card-header"><i class="fas fa-calculator"></i> 平均客单价</div>
                    <div class="stat-card-value orange">
                        <% If SafeNum(totalChannelOrders) > 0 Then %>
                        ¥<%= FormatNumber(SafeNum(totalChannelRevenue) / SafeNum(totalChannelOrders), 2) %>
                        <% Else %>¥0.00<% End If %>
                    </div>
                </div>
                <div class="stat-card">
                    <div class="stat-card-header"><i class="fas fa-layer-group"></i> 渠道数量</div>
                    <div class="stat-card-value purple">
                        <% 
                        Dim channelCount
                        ' Access不支持COUNT(DISTINCT)，使用子查询替代
                        channelCount = GetScalar("SELECT COUNT(*) FROM (SELECT DISTINCT ChannelSource FROM Orders WHERE OrderDate >= '" & safeStart & "' AND OrderDate <= '" & safeEnd & "' AND ChannelSource IS NOT NULL)")
                        Response.Write channelCount
                        %>
                    </div>
                </div>
            </div>
            
            <div class="charts-grid">
                <div class="chart-card">
                    <h3><i class="fas fa-table" style="color:#00bcd4;"></i> 渠道效果明细</h3>
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th>渠道来源</th>
                                <th>订单数</th>
                                <th>销售额</th>
                                <th>平均客单价</th>
                                <th>占比</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% 
                            Set rsChannel = ExecuteQuery(_
                                "SELECT ChannelSource, COUNT(*) AS OrderCount, SUM(CAST(TotalAmount AS FLOAT)) AS Revenue " & _
                                "FROM Orders WHERE OrderDate >= '" & safeStart & "' AND OrderDate <= '" & safeEnd & "' " & _
                                "GROUP BY ChannelSource ORDER BY SUM(CAST(TotalAmount AS FLOAT)) DESC")
                            
                            Dim chTotalOrders, chTotalRevenue
                            ' 循环变量声明移到循环外（VBScript限制）
                            Dim chSource, chOrders, chRevenue, chAvg, chPct
                            chTotalOrders = 0
                            chTotalRevenue = 0
                            
                            If Not rsChannel Is Nothing Then
                            Do While Not rsChannel.EOF
                                chSource = rsChannel("ChannelSource") & ""
                                If chSource = "" Then chSource = "未分类"
                                chOrders = rsChannel("OrderCount")
                                chRevenue = rsChannel("Revenue")
                                If IsNull(chRevenue) Then chRevenue = 0
                                If CDbl(chOrders) > 0 Then chAvg = CDbl(chRevenue) / CDbl(chOrders) Else chAvg = 0
                                If CDbl(totalChannelRevenue) > 0 Then chPct = Round((CDbl(chRevenue) / CDbl(totalChannelRevenue)) * 100, 1) Else chPct = 0
                                chTotalOrders = chTotalOrders + CLng(chOrders)
                                chTotalRevenue = chTotalRevenue + CDbl(chRevenue)
                            %>
                            <tr>
                                <td><i class="fas fa-globe" style="color:#00bcd4;margin-right:8px;"></i><%= chSource %></td>
                                <td><%= chOrders %></td>
                                <td style="color:#4ade80;">¥<%= FormatNumber(CDbl(chRevenue), 2) %></td>
                                <td>¥<%= FormatNumber(chAvg, 2) %></td>
                                <td>
                                    <div class="progress-bar">
                                        <div class="progress-fill blue" style="width: <%= chPct %>%"></div>
                                    </div>
                                    <span style="margin-left:8px;"><%= chPct %>%</span>
                                </td>
                            </tr>
                            <% 
                                rsChannel.MoveNext
                            Loop
                            rsChannel.Close
                            End If
                            %>
                        </tbody>
                        <% If chTotalOrders > 0 Then %>
                        <tfoot>
                            <tr>
                                <td>合计</td>
                                <td><%= chTotalOrders %></td>
                                <td>¥<%= FormatNumber(chTotalRevenue, 2) %></td>
                                <td>¥<%= FormatNumber(IIf(chTotalOrders>0, chTotalRevenue/chTotalOrders, 0), 2) %></td>
                                <td>100%</td>
                            </tr>
                        </tfoot>
                        <% End If %>
                    </table>
                </div>
                
                <div class="chart-card">
                    <h3><i class="fas fa-chart-pie" style="color:#4ade80;"></i> 渠道销售占比</h3>
                    <% 
                    ' 重新查询用于饼图
                    Set rsChannel = ExecuteQuery(_
                        "SELECT TOP 5 ChannelSource, SUM(CAST(TotalAmount AS FLOAT)) AS Revenue " & _
                        "FROM Orders WHERE OrderDate >= '" & safeStart & "' AND OrderDate <= '" & safeEnd & "' " & _
                        "GROUP BY ChannelSource ORDER BY SUM(CAST(TotalAmount AS FLOAT)) DESC")
                    
                    Dim pieData(4, 1), pieTotal, pieIdx
                    pieTotal = 0
                    pieIdx = 0
                    If Not rsChannel Is Nothing Then
                    Do While Not rsChannel.EOF AND pieIdx <= 4
                        pieData(pieIdx, 0) = rsChannel("ChannelSource") & ""
                        If pieData(pieIdx, 0) = "" Then pieData(pieIdx, 0) = "未分类"
                        pieData(pieIdx, 1) = CDbl("0" & rsChannel("Revenue"))
                        pieTotal = pieTotal + pieData(pieIdx, 1)
                        pieIdx = pieIdx + 1
                        rsChannel.MoveNext
                    Loop
                    rsChannel.Close
                    End If
                    
                    ' 计算角度
                    Dim p1, p2, p3, p4, p5
                    If pieTotal > 0 Then
                        p1 = (SafeNum(pieData(0, 1)) / pieTotal) * 360
                        p2 = p1 + (SafeNum(pieData(1, 1)) / pieTotal) * 360
                        p3 = p2 + (SafeNum(pieData(2, 1)) / pieTotal) * 360
                        p4 = p3 + (SafeNum(pieData(3, 1)) / pieTotal) * 360
                    Else
                        p1 = 72: p2 = 144: p3 = 216: p4 = 288
                    End If
                    %>
                    <div class="pie-chart-container">
                        <div class="pie-chart" style="--p1:<%= p1 %>deg; --p2:<%= p2 %>deg; --p3:<%= p3 %>deg; --p4:<%= p4 %>deg;"></div>
                        <div class="pie-legend">
                            <div class="legend-item">
                                <div class="legend-color" style="background:#00bcd4;"></div>
                                <span><%= pieData(0, 0) %> (<%= Round(SafeDiv(SafeNum(pieData(0,1)), pieTotal)*100, 1) %>%)</span>
                            </div>
                            <div class="legend-item">
                                <div class="legend-color" style="background:#4ade80;"></div>
                                <span><%= pieData(1, 0) %> (<%= Round(SafeDiv(SafeNum(pieData(1,1)), pieTotal)*100, 1) %>%)</span>
                            </div>
                            <div class="legend-item">
                                <div class="legend-color" style="background:#fb923c;"></div>
                                <span><%= pieData(2, 0) %> (<%= Round(SafeDiv(SafeNum(pieData(2,1)), pieTotal)*100, 1) %>%)</span>
                            </div>
                            <div class="legend-item">
                                <div class="legend-color" style="background:#a78bfa;"></div>
                                <span><%= pieData(3, 0) %> (<%= Round(SafeDiv(SafeNum(pieData(3,1)), pieTotal)*100, 1) %>%)</span>
                            </div>
                            <div class="legend-item">
                                <div class="legend-color" style="background:#f472b6;"></div>
                                <span>其他 (<%= Round(SafeDiv(SafeNum(pieData(4,1)), pieTotal)*100, 1) %>%)</span>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- ============================================
             Tab 2: 营销活动管理与ROI
             ============================================ -->
        <div class="tab-content <%= IIf(currentTab="campaigns", "active", "") %>">
            <% If canEditCampaign Then %>
            <div style="margin-bottom: 20px;">
                <button class="admin-btn admin-btn-primary" onclick="showCampaignModal()">
                    <i class="fas fa-plus"></i> 新建活动
                </button>
            </div>
            <% End If %>
            
            <div class="chart-card">
                <h3><i class="fas fa-bullhorn" style="color:#00bcd4;"></i> 营销活动ROI分析</h3>
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>活动名称</th>
                            <th>渠道</th>
                            <th>时间范围</th>
                            <th>预算/已花费</th>
                            <th>产生收入</th>
                            <th>订单数</th>
                            <th>ROI</th>
                            <th>转化率</th>
                            <th>状态</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% If Not rsCampaigns Is Nothing Then %>
                        <% Do While Not rsCampaigns.EOF %>
                        <%
                        
                        cpName = rsCampaigns("CampaignName") & ""
                        If IsNull(rsCampaigns("ChannelSource")) Then
                            cpChannel = "-"
                        Else
                            cpChannel = CStr(rsCampaigns("ChannelSource"))
                            If cpChannel = "" Then cpChannel = "-"
                        End If
                        ' 日期字段安全处理
                        If IsNull(rsCampaigns("StartDate")) Or IsEmpty(rsCampaigns("StartDate")) Then
                            cpStart = Date()
                        Else
                            cpStart = rsCampaigns("StartDate")
                        End If
                        If IsNull(rsCampaigns("EndDate")) Or IsEmpty(rsCampaigns("EndDate")) Then
                            cpEnd = Date()
                        Else
                            cpEnd = rsCampaigns("EndDate")
                        End If
                        ' 数值字段使用SafeNum安全转换
                        cpBudget = SafeNum(rsCampaigns("BudgetAmount"))
                        cpSpent = SafeNum(rsCampaigns("SpentAmount"))
                        cpRevenue = SafeNum(rsCampaigns("Revenue"))
                        cpOrders = SafeNum(rsCampaigns("OrderCount"))
                        cpClicks = SafeNum(rsCampaigns("ClickCount"))
                        
                        ' 计算ROI
                        If cpSpent > 0 Then
                            cpROI = Round(cpRevenue / cpSpent, 2)
                        Else
                            cpROI = 0
                        End If
                        
                        ' 计算转化率
                        If cpClicks > 0 Then
                            cpConv = Round((CDbl(cpOrders) / CDbl(cpClicks)) * 100, 2)
                        Else
                            cpConv = 0
                        End If
                        
                        ' 状态判断
                        If Date() < cpStart Then
                            cpStatus = "pending"
                            cpStatusClass = "status-ended"
                            cpStatusName = "未开始"
                        ElseIf Date() > cpEnd Then
                            cpStatus = "ended"
                            cpStatusClass = "status-ended"
                            cpStatusName = "已结束"
                        ElseIf cpSpent > cpBudget AND cpBudget > 0 Then
                            cpStatus = "overbudget"
                            cpStatusClass = "status-overbudget"
                            cpStatusName = "超预算"
                        Else
                            cpStatus = "ongoing"
                            cpStatusClass = "status-ongoing"
                            cpStatusName = "进行中"
                        End If
                        
                        ' ROI颜色
                        If cpROI > 3 Then
                            roiClass = "roi-high"
                        ElseIf cpROI >= 1 Then
                            roiClass = "roi-medium"
                        Else
                            roiClass = "roi-low"
                        End If
                        %>
                        <tr>
                            <td><strong><%= cpName %></strong></td>
                            <td><%= cpChannel %></td>
                            <td><%= SafeFormatDateTime(cpStart, 2) %> ~ <%= SafeFormatDateTime(cpEnd, 2) %></td>
                            <td>
                                ¥<%= FormatNumber(cpSpent, 0) %> / ¥<%= FormatNumber(cpBudget, 0) %>
                                <% If cpBudget > 0 Then %>
                                <div class="progress-bar" style="width:80px;margin-top:5px;">
                                    <div class="progress-fill <%= IIf(cpSpent>cpBudget,"red","blue") %>" style="width:<%= IIf(cpSpent>cpBudget,100,Round((cpSpent/cpBudget)*100,0)) %>%"></div>
                                </div>
                                <% End If %>
                            </td>
                            <td style="color:#4ade80;">¥<%= FormatNumber(cpRevenue, 0) %></td>
                            <td><%= cpOrders %></td>
                            <td class="<%= roiClass %>"><%= cpROI %>x</td>
                            <td><%= cpConv %>%</td>
                            <td><span class="status-badge <%= cpStatusClass %>"><%= cpStatusName %></span></td>
                        </tr>
                        <% rsCampaigns.MoveNext %>
                        <% Loop %>
                        <% rsCampaigns.Close %>
                        <% End If %>
                    </tbody>
                </table>
            </div>
            
            <!-- ROI 说明 -->
            <div style="display:flex;gap:20px;margin-top:20px;">
                <div style="background:#252836;padding:15px 20px;border-radius:8px;display:flex;align-items:center;gap:10px;">
                    <span class="roi-high" style="font-size:18px;">ROI > 3</span>
                    <span style="color:#888;">优秀</span>
                </div>
                <div style="background:#252836;padding:15px 20px;border-radius:8px;display:flex;align-items:center;gap:10px;">
                    <span class="roi-medium" style="font-size:18px;">ROI 1-3</span>
                    <span style="color:#888;">良好</span>
                </div>
                <div style="background:#252836;padding:15px 20px;border-radius:8px;display:flex;align-items:center;gap:10px;">
                    <span class="roi-low" style="font-size:18px;">ROI < 1</span>
                    <span style="color:#888;">需优化</span>
                </div>
            </div>
        </div>
        
        <!-- ============================================
             Tab 3: 用户消费行为分析
             ============================================ -->
        <div class="tab-content <%= IIf(currentTab="behavior", "active", "") %>">
            <div class="stats-grid">
                <div class="stat-card">
                    <div class="stat-card-header"><i class="fas fa-users"></i> 总用户数</div>
                    <div class="stat-card-value blue"><%= totalUsers %></div>
                </div>
                <div class="stat-card">
                    <div class="stat-card-header"><i class="fas fa-redo"></i> 复购用户数</div>
                    <div class="stat-card-value green"><%= repeatUsers %></div>
                </div>
                <div class="stat-card">
                    <div class="stat-card-header"><i class="fas fa-percentage"></i> 复购率</div>
                    <div class="stat-card-value orange"><%= repeatRate %>%</div>
                </div>
                <div class="stat-card">
                    <div class="stat-card-header"><i class="fas fa-shopping-bag"></i> 人均订单</div>
                    <div class="stat-card-value purple">
                        <% If CDbl(totalUsers) > 0 Then %>
                        <%= Round(CDbl(totalChannelOrders) / CDbl(totalUsers), 2) %>
                        <% Else %>0<% End If %>
                    </div>
                </div>
            </div>
            
            <div class="charts-grid">
                <div class="chart-card">
                    <h3><i class="fas fa-chart-line" style="color:#00bcd4;"></i> 最近6个月客单价趋势</h3>
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th>月份</th>
                                <th>订单数</th>
                                <th>销售额</th>
                                <th>平均客单价</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% If Not rsAvgOrderValue Is Nothing Then %>
                            <% Do While Not rsAvgOrderValue.EOF %>
                            <%
                            aovYear = rsAvgOrderValue("Y")
                            aovMonth = rsAvgOrderValue("M")
                            aovCount = rsAvgOrderValue("OrderCount")
                            aovRev = rsAvgOrderValue("Revenue")
                            If IsNull(aovRev) Then aovRev = 0
                            If CDbl(aovCount) > 0 Then aovAvg = CDbl(aovRev) / CDbl(aovCount) Else aovAvg = 0
                            %>
                            <tr>
                                <td><%= aovYear %>年<%= aovMonth %>月</td>
                                <td><%= aovCount %></td>
                                <td style="color:#4ade80;">¥<%= FormatNumber(CDbl(aovRev), 0) %></td>
                                <td style="color:#fb923c;font-weight:bold;">¥<%= FormatNumber(aovAvg, 2) %></td>
                            </tr>
                            <% rsAvgOrderValue.MoveNext %>
                            <% Loop %>
                            <% rsAvgOrderValue.Close %>
                            <% End If %>
                        </tbody>
                    </table>
                </div>
                
                <div class="chart-card">
                    <h3><i class="fas fa-chart-bar" style="color:#4ade80;"></i> 消费频次分布</h3>
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th>消费频次</th>
                                <th>用户数</th>
                                <th>占比</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td><i class="fas fa-user" style="color:#00bcd4;margin-right:8px;"></i>1次</td>
                                <td><%= freq1 %></td>
                                <td>
                                    <div class="progress-bar">
                                        <div class="progress-fill blue" style="width:<%= Round(SafeDiv(freq1,totalFreqUsers)*100,0) %>"></div>
                                    </div>
                                    <span style="margin-left:8px;"><%= Round(SafeDiv(freq1,totalFreqUsers)*100,1) %>%</span>
                                </td>
                            </tr>
                            <tr>
                                <td><i class="fas fa-user-friends" style="color:#4ade80;margin-right:8px;"></i>2-3次</td>
                                <td><%= freq2_3 %></td>
                                <td>
                                    <div class="progress-bar">
                                        <div class="progress-fill green" style="width:<%= Round(SafeDiv(freq2_3,totalFreqUsers)*100,0) %>"></div>
                                    </div>
                                    <span style="margin-left:8px;"><%= Round(SafeDiv(freq2_3,totalFreqUsers)*100,1) %>%</span>
                                </td>
                            </tr>
                            <tr>
                                <td><i class="fas fa-users" style="color:#fb923c;margin-right:8px;"></i>4-5次</td>
                                <td><%= freq4_5 %></td>
                                <td>
                                    <div class="progress-bar">
                                        <div class="progress-fill orange" style="width:<%= Round(SafeDiv(freq4_5,totalFreqUsers)*100,0) %>"></div>
                                    </div>
                                    <span style="margin-left:8px;"><%= Round(SafeDiv(freq4_5,totalFreqUsers)*100,1) %>%</span>
                                </td>
                            </tr>
                            <tr>
                                <td><i class="fas fa-crown" style="color:#a78bfa;margin-right:8px;"></i>6次以上</td>
                                <td><%= freq6plus %></td>
                                <td>
                                    <div class="progress-bar">
                                        <div class="progress-fill" style="width:<%= Round(SafeDiv(freq6plus,totalFreqUsers)*100,0) %>%;background:linear-gradient(90deg,#a78bfa,#7c3aed);"></div>
                                    </div>
                                    <span style="margin-left:8px;"><%= Round(SafeDiv(freq6plus,totalFreqUsers)*100,1) %>%</span>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        
        <!-- ============================================
             Tab 4: 转化率漏斗
             ============================================ -->
        <div class="tab-content <%= IIf(currentTab="funnel", "active", "") %>">
            <div class="chart-card">
                <h3><i class="fas fa-filter" style="color:#00bcd4;"></i> 用户转化漏斗</h3>
                <div class="funnel-container">
                    <div class="funnel-level funnel-level-1">
                        <div class="funnel-info">
                            <div class="funnel-label">注册用户</div>
                            <div class="funnel-value"><%= funnelRegister %></div>
                        </div>
                    </div>
                    <div class="funnel-level funnel-level-2">
                        <div class="funnel-info">
                            <div class="funnel-label">下单用户</div>
                            <div class="funnel-value"><%= funnelOrderUsers %></div>
                        </div>
                        <div class="funnel-rate">
                            <% If CDbl(funnelRegister) > 0 Then %>
                            转化率 <%= Round((CDbl(funnelOrderUsers)/CDbl(funnelRegister))*100, 1) %>%
                            <% Else %>-%<% End If %>
                        </div>
                    </div>
                    <div class="funnel-level funnel-level-3">
                        <div class="funnel-info">
                            <div class="funnel-label">付款用户</div>
                            <div class="funnel-value"><%= funnelPaidUsers %></div>
                        </div>
                        <div class="funnel-rate">
                            <% If CDbl(funnelOrderUsers) > 0 Then %>
                            转化率 <%= Round((CDbl(funnelPaidUsers)/CDbl(funnelOrderUsers))*100, 1) %>%
                            <% Else %>-%<% End If %>
                        </div>
                    </div>
                    <div class="funnel-level funnel-level-4">
                        <div class="funnel-info">
                            <div class="funnel-label">复购用户</div>
                            <div class="funnel-value"><%= funnelRepeatUsers %></div>
                        </div>
                        <div class="funnel-rate">
                            <% If CDbl(funnelPaidUsers) > 0 Then %>
                            转化率 <%= Round((CDbl(funnelRepeatUsers)/CDbl(funnelPaidUsers))*100, 1) %>%
                            <% Else %>-%<% End If %>
                        </div>
                    </div>
                </div>
            </div>
            
            <div class="stats-grid" style="margin-top:25px;">
                <div class="stat-card">
                    <div class="stat-card-header"><i class="fas fa-mouse-pointer"></i> 注册→下单转化率</div>
                    <div class="stat-card-value blue">
                        <% If CDbl(funnelRegister) > 0 Then %>
                        <%= Round((CDbl(funnelOrderUsers)/CDbl(funnelRegister))*100, 1) %>%
                        <% Else %>0%<% End If %>
                    </div>
                </div>
                <div class="stat-card">
                    <div class="stat-card-header"><i class="fas fa-credit-card"></i> 下单→付款转化率</div>
                    <div class="stat-card-value green">
                        <% If CDbl(funnelOrderUsers) > 0 Then %>
                        <%= Round((CDbl(funnelPaidUsers)/CDbl(funnelOrderUsers))*100, 1) %>%
                        <% Else %>0%<% End If %>
                    </div>
                </div>
                <div class="stat-card">
                    <div class="stat-card-header"><i class="fas fa-redo-alt"></i> 付款→复购转化率</div>
                    <div class="stat-card-value orange">
                        <% If CDbl(funnelPaidUsers) > 0 Then %>
                        <%= Round((CDbl(funnelRepeatUsers)/CDbl(funnelPaidUsers))*100, 1) %>%
                        <% Else %>0%<% End If %>
                    </div>
                </div>
                <div class="stat-card">
                    <div class="stat-card-header"><i class="fas fa-chart-line"></i> 整体转化率</div>
                    <div class="stat-card-value purple">
                        <% If CDbl(funnelRegister) > 0 Then %>
                        <%= Round((CDbl(funnelRepeatUsers)/CDbl(funnelRegister))*100, 1) %>%
                        <% Else %>0%<% End If %>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- ============================================
             Tab 5: 营销费用与预算对比
             ============================================ -->
        <div class="tab-content <%= IIf(currentTab="budget", "active", "") %>">
            <div class="stats-grid">
                <div class="stat-card">
                    <div class="stat-card-header"><i class="fas fa-wallet"></i> 总预算</div>
                    <div class="stat-card-value blue">¥<%= FormatNumber(SafeNum(totalBudget), 0) %></div>
                </div>
                <div class="stat-card">
                    <div class="stat-card-header"><i class="fas fa-money-bill-wave"></i> 实际支出</div>
                    <div class="stat-card-value <%= IIf(SafeNum(totalActual)>SafeNum(totalBudget),"roi-low","green") %>">¥<%= FormatNumber(SafeNum(totalActual), 0) %></div>
                </div>
                <div class="stat-card">
                    <div class="stat-card-header"><i class="fas fa-percentage"></i> 执行率</div>
                    <div class="stat-card-value <%= IIf(SafeNum(totalActual)>SafeNum(totalBudget),"roi-low","orange") %>">
                        <% If SafeNum(totalBudget) > 0 Then %>
                        <%= Round((SafeNum(totalActual)/SafeNum(totalBudget))*100, 1) %>%
                        <% Else %>0%<% End If %>
                    </div>
                </div>
                <div class="stat-card">
                    <div class="stat-card-header"><i class="fas fa-chart-line"></i> 总ROI</div>
                    <div class="stat-card-value <%= IIf(SafeNum(totalGMV)>SafeNum(totalActual),"roi-high","roi-low") %>">
                        <% If SafeNum(totalActual) > 0 Then %>
                        <%= Round(SafeNum(totalGMV)/SafeNum(totalActual), 2) %>x
                        <% Else %>0x<% End If %>
                    </div>
                </div>
            </div>
            
            <% If SafeNum(totalActual) > SafeNum(totalBudget) AND SafeNum(totalBudget) > 0 Then %>
            <div class="alert-warning">
                <i class="fas fa-exclamation-triangle"></i>
                警告：营销费用已超出预算 ¥<%= FormatNumber(SafeNum(totalActual)-SafeNum(totalBudget), 0) %>，超支率 <%= Round(((SafeNum(totalActual)-SafeNum(totalBudget))/SafeNum(totalBudget))*100, 1) %>%
            </div>
            <% End If %>
            
            <div class="chart-card">
                <h3><i class="fas fa-table" style="color:#00bcd4;"></i> 月度预算执行情况</h3>
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>月份</th>
                            <th>预算</th>
                            <th>实际支出</th>
                            <th>执行率</th>
                            <th>GMV</th>
                            <th>ROI</th>
                            <th>状态</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% If Not rsBudget Is Nothing Then %>
                        <% Do While Not rsBudget.EOF %>
                        <%
                        
                        bgPeriod = rsBudget("Period") & ""
                        ' 数值字段使用SafeNum安全转换
                        bgBudget = SafeNum(rsBudget("BudgetAmount"))
                        bgActual = SafeNum(rsBudget("ActualAmount"))
                        bgGMV = SafeNum(rsBudget("GMVAmount"))
                        bgROI = SafeNum(rsBudget("ROI"))
                        
                        If CDbl(bgBudget) > 0 Then
                            bgRate = Round((bgActual / bgBudget) * 100, 1)
                        Else
                            bgRate = 0
                        End If
                        
                        If bgActual > bgBudget AND bgBudget > 0 Then
                            bgStatus = "超支"
                            bgStatusClass = "roi-low"
                        ElseIf bgRate >= 90 Then
                            bgStatus = "预警"
                            bgStatusClass = "roi-medium"
                        Else
                            bgStatus = "正常"
                            bgStatusClass = "roi-high"
                        End If
                        %>
                        <tr>
                            <td><i class="fas fa-calendar-alt" style="color:#00bcd4;margin-right:8px;"></i><%= bgPeriod %></td>
                            <td>¥<%= FormatNumber(bgBudget, 0) %></td>
                            <td style="<%= IIf(bgActual>bgBudget,"color:#f87171;font-weight:bold","") %>">¥<%= FormatNumber(bgActual, 0) %></td>
                            <td>
                                <div class="progress-bar">
                                    <div class="progress-fill <%= IIf(bgActual>bgBudget,"red",IIf(bgRate>=90,"orange","blue")) %>" style="width:<%= IIf(bgRate>100,100,bgRate) %>"></div>
                                </div>
                                <span style="margin-left:8px;"><%= bgRate %>%</span>
                            </td>
                            <td style="color:#4ade80;">¥<%= FormatNumber(bgGMV, 0) %></td>
                            <td class="<%= IIf(bgROI>=1.5,"roi-high",IIf(bgROI>=1,"roi-medium","roi-low")) %>"><%= bgROI %>x</td>
                            <td class="<%= bgStatusClass %>"><%= bgStatus %></td>
                        </tr>
                        <% rsBudget.MoveNext %>
                        <% Loop %>
                        <% rsBudget.Close %>
                        <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    
    <% If canEditCampaign Then %>
    <!-- 新建活动模态框 -->
    <div id="campaignModal" class="admin-modal">
        <div class="admin-modal-content" style="background:#252836;color:#fff;">
            <div class="admin-modal-header" style="border-color:#3d4252;">
                <h3 class="admin-modal-title"><i class="fas fa-plus"></i> 新建营销活动</h3>
                <button class="admin-modal-close" onclick="closeCampaignModal()">&times;</button>
            </div>
            <form method="post" action="marketing_stats.asp?tab=campaigns">
                <input type="hidden" name="action" value="create_campaign">
                <div class="admin-modal-body">
                    <div class="admin-form-group">
                        <label class="admin-form-label">活动名称</label>
                        <input type="text" name="campaignName" class="admin-form-control" style="background:#2d2d44;border-color:rgba(255,255,255,0.15);color:#e0e0e0;" required>
                    </div>
                    <div class="admin-form-group">
                        <label class="admin-form-label">渠道来源</label>
                        <input type="text" name="channelSource" class="admin-form-control" style="background:#2d2d44;border-color:rgba(255,255,255,0.15);color:#e0e0e0;" placeholder="如：微信、抖音、百度等">
                    </div>
                    <div class="admin-form-row">
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label class="admin-form-label">开始日期</label>
                                <input type="date" name="startDate" class="admin-form-control" style="background:#2d2d44;border-color:rgba(255,255,255,0.15);color:#e0e0e0;" required>
                            </div>
                        </div>
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label class="admin-form-label">结束日期</label>
                                <input type="date" name="endDate" class="admin-form-control" style="background:#2d2d44;border-color:rgba(255,255,255,0.15);color:#e0e0e0;" required>
                            </div>
                        </div>
                    </div>
                    <div class="admin-form-group">
                        <label class="admin-form-label">预算金额 (¥)</label>
                        <input type="number" name="budgetAmount" class="admin-form-control" style="background:#2d2d44;border-color:rgba(255,255,255,0.15);color:#e0e0e0;" value="0" min="0" step="0.01">
                    </div>
                </div>
                <div class="admin-modal-footer" style="background:#1e1e32;border-color:rgba(255,255,255,0.06);">
                    <button type="button" class="admin-btn btn--neutral" onclick="closeCampaignModal()">取消</button>
                    <button type="submit" class="admin-btn admin-btn-primary">创建活动</button>
                </div>
            </form>
        </div>
    </div>
    <% End If %>
    
    <script>
        function showCampaignModal() {
            document.getElementById('campaignModal').style.display = 'block';
        }
        
        function closeCampaignModal() {
            document.getElementById('campaignModal').style.display = 'none';
        }
        
        // 点击模态框外部关闭
        window.onclick = function(event) {
            var modal = document.getElementById('campaignModal');
            if (event.target == modal) {
                modal.style.display = 'none';
            }
        }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
