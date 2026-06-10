<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<!--#include file="../../includes/cost_engine.asp"-->
<%
' ============================================
' 综合报表中心 - Comprehensive Report Center
' 企业三流闭环：信息流、物流、资金流
' ============================================

Function SafeNum(val)
    If IsNull(val) Or IsEmpty(val) Or val = "" Then SafeNum = 0 Else On Error Resume Next: SafeNum = CDbl(val): If Err.Number <> 0 Then SafeNum = 0: Err.Clear: End If
End Function

Function GetScalarVal(sql)
    Dim rs, val: val = 0
    Set rs = conn.Execute(sql)
    If Not rs Is Nothing Then If Not rs.EOF Then val = SafeNum(rs(0)): rs.Close
    Set rs = Nothing: GetScalarVal = val
End Function

Call OpenConnection()

' 日期范围
Dim sd, ed
sd = Trim(Request.QueryString("sd"))
ed = Trim(Request.QueryString("ed"))
If sd = "" Then sd = DateAdd("m", -3, Date())
If ed = "" Then ed = Date()
Dim safeSD, safeED
safeSD = SafeSQL(FormatDateTime(sd, 2))
safeED = SafeSQL(FormatDateTime(ed, 2))

' ============================================
' 1. 信息流 - 核心业务数据
' ============================================
Dim totalOrders, totalRevenue, totalCustomers, newCustomers, avgOrderValue, orderGrowth
totalOrders = GetScalarVal("SELECT COUNT(*) FROM Orders WHERE Status NOT IN ('Cancelled') AND CreatedAt BETWEEN '" & safeSD & "' AND '" & safeED & "'")
totalRevenue = GetScalarVal("SELECT ISNULL(SUM(CAST(TotalAmount AS FLOAT)),0) FROM Orders WHERE Status IN ('Paid','Completed') AND CreatedAt BETWEEN '" & safeSD & "' AND '" & safeED & "'")
totalCustomers = GetScalarVal("SELECT COUNT(DISTINCT UserID) FROM Orders WHERE Status NOT IN ('Cancelled') AND CreatedAt BETWEEN '" & safeSD & "' AND '" & safeED & "'")
newCustomers = GetScalarVal("SELECT COUNT(*) FROM Users WHERE CreatedAt BETWEEN '" & safeSD & "' AND '" & safeED & "'")
If totalOrders > 0 Then avgOrderValue = totalRevenue / totalOrders Else avgOrderValue = 0

' 订单趋势（按月）
Dim rsMonthlyOrders
Set rsMonthlyOrders = conn.Execute("SELECT Year(CreatedAt) AS Y, Month(CreatedAt) AS M, COUNT(*) AS OCnt, ISNULL(SUM(CAST(TotalAmount AS FLOAT)),0) AS ORev " & _
    "FROM Orders WHERE Status NOT IN ('Cancelled') AND CreatedAt BETWEEN '" & safeSD & "' AND '" & safeED & "' " & _
    "GROUP BY Year(CreatedAt), Month(CreatedAt) ORDER BY Y, M")

' 产品线分布
Dim rsProdLine
Set rsProdLine = conn.Execute("SELECT p.ProductType, COUNT(DISTINCT od.OrderID) AS OCnt, ISNULL(SUM(od.Subtotal),0) AS Rev, COUNT(od.DetailID) AS UnitCnt " & _
    "FROM OrderDetails od JOIN Orders o ON od.OrderID=o.OrderID LEFT JOIN Products p ON od.ProductID=p.ProductID " & _
    "WHERE o.Status IN ('Paid','Completed') AND o.CreatedAt BETWEEN '" & safeSD & "' AND '" & safeED & "' " & _
    "GROUP BY p.ProductType ORDER BY Rev DESC")

' ============================================
' 2. 物流 - 进产销存数据
' ============================================
Dim rawStockValue, noteStockValue, prodStockValue, totalInvValue
rawStockValue = GetScalarVal("SELECT ISNULL(SUM(StockQty * UnitPrice),0) FROM RawMaterialInventory WHERE StockQty > 0")
noteStockValue = GetScalarVal("SELECT ISNULL(SUM(ni.StockQuantity * ISNULL(fn.PriceAddition,0)),0) FROM NoteInventory ni LEFT JOIN FragranceNotes fn ON ni.NoteID=fn.NoteID WHERE ni.StockQuantity > 0")
prodStockValue = GetScalarVal("SELECT ISNULL(SUM(StockQty * UnitCost),0) FROM ProductInventory WHERE StockQty > 0")
totalInvValue = rawStockValue + noteStockValue + prodStockValue

Dim inProduction, pendingShip, completedProd
inProduction = GetScalarVal("SELECT COUNT(*) FROM ProductionOrders WHERE Status='InProgress'")
pendingShip = GetScalarVal("SELECT COUNT(*) FROM Orders WHERE Status IN ('Paid','Processing') AND ShippingStatus IS NULL")
completedProd = GetScalarVal("SELECT COUNT(*) FROM ProductionOrders WHERE Status='Completed'")

' 原料库存周转（低库存项数）
Dim lowStockCount
lowStockCount = GetScalarVal("SELECT COUNT(*) FROM RawMaterialInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0")

' ============================================
' 3. 资金流 - 财务数据
' ============================================
Dim totalCost, totalProfit, profitMargin, pendingPayable, pendingReceivable
totalCost = GetScalarVal("SELECT ISNULL(SUM(CAST(CostAmount AS FLOAT)),0) FROM Orders WHERE Status IN ('Paid','Completed') AND CreatedAt BETWEEN '" & safeSD & "' AND '" & safeED & "'")
totalProfit = totalRevenue - totalCost
If totalRevenue > 0 Then profitMargin = (totalProfit / totalRevenue) * 100 Else profitMargin = 0
pendingPayable = GetScalarVal("SELECT ISNULL(SUM(CAST(TotalAmount AS FLOAT)),0) FROM PurchaseOrders WHERE Status IN ('Pending','Approved')")
pendingReceivable = GetScalarVal("SELECT ISNULL(SUM(CAST(TotalAmount AS FLOAT)),0) FROM Orders WHERE Status='Paid' AND ShippingStatus<>'Delivered'")

' 利润率月度趋势
Dim rsProfitTrend
Set rsProfitTrend = conn.Execute("SELECT Year(CreatedAt) AS Y, Month(CreatedAt) AS M, " & _
    "ISNULL(SUM(CAST(TotalAmount AS FLOAT)),0) AS Rev, ISNULL(SUM(CAST(CostAmount AS FLOAT)),0) AS Cost " & _
    "FROM Orders WHERE Status IN ('Paid','Completed') AND CreatedAt BETWEEN '" & safeSD & "' AND '" & safeED & "' " & _
    "GROUP BY Year(CreatedAt), Month(CreatedAt) ORDER BY Y, M")

Dim totalCostOverall, totalRevOverall
totalCostOverall = GetScalarVal("SELECT ISNULL(SUM(CAST(CostAmount AS FLOAT)),0) FROM Orders WHERE Status IN ('Paid','Completed')")
totalRevOverall = GetScalarVal("SELECT ISNULL(SUM(CAST(TotalAmount AS FLOAT)),0) FROM Orders WHERE Status IN ('Paid','Completed')")

' ============================================
' 4. 闭环KPI
' ============================================
Dim invTurnover, orderToCash, customerLTV, profitPerOrder
' 库存周转率 = 销售成本 / 平均库存价值
Dim annualizedCOGS
annualizedCOGS = GetScalarVal("SELECT ISNULL(SUM(CAST(CostAmount AS FLOAT)),0) FROM Orders WHERE Status IN ('Paid','Completed') AND CreatedAt >= DATEADD(year, -1, GETDATE())")
If totalInvValue > 0 Then invTurnover = annualizedCOGS / totalInvValue Else invTurnover = 0

' 订单到现金周期（天）
Dim avgFulfillmentDays
avgFulfillmentDays = GetScalarVal("SELECT ISNULL(AVG(DATEDIFF(day, CreatedAt, ISNULL(DeliveredAt, GETDATE()))),0) FROM Orders WHERE DeliveredAt IS NOT NULL AND CreatedAt BETWEEN '" & safeSD & "' AND '" & safeED & "'")

' 客户生命周期价值
If totalCustomers > 0 Then customerLTV = totalRevenue / totalCustomers Else customerLTV = 0

' 每单利润
If totalOrders > 0 Then profitPerOrder = totalProfit / totalOrders Else profitPerOrder = 0

' 预计算闭环指标样式
Dim invColor, invClass, invIcon, invText
If invTurnover >= 6 Then
    invColor = "#81c784": invClass = "c-green": invIcon = "fa-check-circle": invText = "优秀"
ElseIf invTurnover >= 3 Then
    invColor = "#ffb74d": invClass = "c-orange": invIcon = "fa-exclamation-circle": invText = "正常"
Else
    invColor = "#e57373": invClass = "c-red": invIcon = "fa-times-circle": invText = "偏低"
End If

Dim fulfillColor, fulfillClass, fulfillIcon, fulfillText
If avgFulfillmentDays <= 3 And avgFulfillmentDays > 0 Then
    fulfillColor = "#81c784": fulfillClass = "c-green": fulfillIcon = "fa-check-circle": fulfillText = "高效"
ElseIf avgFulfillmentDays <= 7 Then
    fulfillColor = "#ffb74d": fulfillClass = "c-orange": fulfillIcon = "fa-exclamation-circle": fulfillText = "正常"
Else
    fulfillColor = "#e57373": fulfillClass = "c-red": fulfillIcon = "fa-times-circle": fulfillText = "需优化"
End If

Dim ltvColor, ltvClass, ltvIcon, ltvText
If customerLTV >= 500 Then
    ltvColor = "#81c784": ltvClass = "c-green": ltvIcon = "fa-check-circle": ltvText = "优质"
ElseIf customerLTV >= 200 Then
    ltvColor = "#ffb74d": ltvClass = "c-orange": ltvIcon = "fa-exclamation-circle": ltvText = "一般"
Else
    ltvColor = "#e57373": ltvClass = "c-red": ltvIcon = "fa-times-circle": ltvText = "偏低"
End If

Dim ppoColor, ppoClass, ppoIcon, ppoText
If profitPerOrder >= 50 Then
    ppoColor = "#81c784": ppoClass = "c-green": ppoIcon = "fa-check-circle": ppoText = "健康"
ElseIf profitPerOrder >= 20 Then
    ppoColor = "#ffb74d": ppoClass = "c-orange": ppoIcon = "fa-exclamation-circle": ppoText = "正常"
Else
    ppoColor = "#e57373": ppoClass = "c-red": ppoIcon = "fa-times-circle": ppoText = "需改善"
End If

Call LogAdminAction("查看综合报表", "finance", "comprehensive_report", "", safeSD & " 至 " & safeED)
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>综合报表中心 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background:#1a1a2e; color:#e0e0e0; }
        .main-content { padding:30px; margin-left:260px; }
        .page-title { font-size:24px; display:flex; align-items:center; gap:12px; margin-bottom:25px; }
        .page-title i { color:#00bcd4; }
        .breadcrumb { color:#888; font-size:14px; margin-bottom:20px; }
        .breadcrumb a { color:#00bcd4; text-decoration:none; }
        
        /* 筛选 */
        .filter-bar { background:linear-gradient(135deg,#2d2d44,#1e1e32); border-radius:12px; padding:20px; margin-bottom:25px; border:1px solid rgba(255,255,255,0.06); }
        .filter-bar form { display:flex; gap:15px; align-items:center; flex-wrap:wrap; }
        .filter-bar label { color:#888; font-size:13px; }
        .filter-bar input { padding:8px 12px; border:1px solid #3a3a4a; border-radius:6px; background:#1a1a2e; color:#e0e0e0; }
        /* 三流标题 */
        .flow-section { margin-bottom:30px; }
        .flow-header { display:flex; align-items:center; gap:12px; margin-bottom:20px; padding:15px 20px; background:linear-gradient(135deg,#2d2d44,#1e1e32); border-radius:10px; border-left:4px solid; }
        .flow-header i { font-size:24px; }
        .flow-header h3 { margin:0; font-size:18px; }
        .flow-header p { margin:5px 0 0; font-size:12px; color:#888; }
        .flow-info { border-left-color:#2196F3; } .flow-info i { color:#2196F3; } .flow-info h3 { color:#64b5f6; }
        .flow-logistics { border-left-color:#FF9800; } .flow-logistics i { color:#FF9800; } .flow-logistics h3 { color:#ffb74d; }
        .flow-capital { border-left-color:#4CAF50; } .flow-capital i { color:#4CAF50; } .flow-capital h3 { color:#81c784; }
        .flow-closed { border-left-color:#9C27B0; } .flow-closed i { color:#9C27B0; } .flow-closed h3 { color:#ce93d8; }
        
        /* KPI网格 */
        .kpi-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(160px,1fr)); gap:12px; margin-bottom:20px; }
        .kpi-card { background:#1a1a2e; border-radius:10px; padding:15px; text-align:center; border:1px solid rgba(255,255,255,0.04); }
        .kpi-card .lbl { font-size:11px; color:#888; margin-bottom:5px; }
        .kpi-card .val { font-size:22px; font-weight:700; }
        .kpi-card .sub { font-size:10px; color:#666; margin-top:3px; }
        
        /* 表格 */
        .data-table { width:100%; border-collapse:collapse; font-size:13px; background:#1a1a2e; border-radius:8px; overflow:hidden; }
        .data-table th { background:rgba(0,188,212,0.15); color:#00bcd4; padding:10px; text-align:left; border-bottom:1px solid #3a3a4a; font-weight:600; }
        .data-table td { padding:8px 10px; border-bottom:1px solid rgba(255,255,255,0.04); }
        .data-table tr:hover td { background:rgba(255,255,255,0.03); }
        
        .section-card { background:linear-gradient(135deg,#2d2d44,#1e1e32); border-radius:12px; padding:20px; margin-bottom:20px; border:1px solid rgba(255,255,255,0.06); }
        .section-subtitle { color:#e0e0e0; font-size:15px; margin:0 0 15px; display:flex; align-items:center; gap:8px; }
        
        .mb-20 { margin-bottom:20px; }
        .text-green { color:#81c784; } .text-blue { color:#64b5f6; } .text-orange { color:#ffb74d; } .text-purple { color:#ce93d8; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    <div class="main-content">
        <div class="breadcrumb">
            <a href="index.asp">财务中心</a> / <span>综合报表中心</span>
        </div>
        <h2 class="page-title"><i class="fas fa-chart-pie"></i> 综合报表中心</h2>
        
        <!-- 日期筛选 -->
        <div class="filter-bar">
            <form method="get" action="comprehensive_report.asp">
                <label>开始:</label>
                <input type="date" name="sd" value="<%= FormatDateTime(sd,2) %>">
                <label>结束:</label>
                <input type="date" name="ed" value="<%= FormatDateTime(ed,2) %>">
                <button type="submit" class="btn btn-primary"><i class="fas fa-search"></i> 查询</button>
            </form>
        </div>
        
        <!-- ====== 信息流 ====== -->
        <div class="flow-section">
            <div class="flow-header flow-info">
                <i class="fas fa-info-circle"></i>
                <div><h3>信息流 - 核心业务数据</h3><p>订单趋势、客户获取、产品销售，反映业务健康状况</p></div>
            </div>
            <div class="kpi-grid">
                <div class="kpi-card"><div class="lbl">总订单</div><div class="val text-blue"><%= totalOrders %></div><div class="sub">所选周期</div></div>
                <div class="kpi-card"><div class="lbl">总营收</div><div class="val text-green">¥<%= FormatNumber(totalRevenue,0) %></div><div class="sub">已支付</div></div>
                <div class="kpi-card"><div class="lbl">客户数</div><div class="val text-orange"><%= totalCustomers %></div><div class="sub">下单客户</div></div>
                <div class="kpi-card"><div class="lbl">新客户</div><div class="val text-blue"><%= newCustomers %></div><div class="sub">新注册</div></div>
                <div class="kpi-card"><div class="lbl">客单价</div><div class="val text-purple">¥<%= FormatNumber(avgOrderValue,2) %></div><div class="sub">平均</div></div>
            </div>
            
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;">
                <!-- 月度订单趋势 -->
                <div class="section-card">
                    <h4 class="section-subtitle"><i class="fas fa-calendar-alt text-blue"></i> 月度订单趋势</h4>
                    <table class="data-table">
                        <thead><tr><th>月份</th><th>订单数</th><th>营收</th></tr></thead>
                        <tbody>
                            <% If Not rsMonthlyOrders Is Nothing Then
                                Dim hasMO: hasMO = False
                                Do While Not rsMonthlyOrders.EOF
                                    hasMO = True %>
                            <tr>
                                <td><%= rsMonthlyOrders("Y") %>年<%= rsMonthlyOrders("M") %>月</td>
                                <td class="text-blue"><%= rsMonthlyOrders("OCnt") %></td>
                                <td class="text-green">¥<%= FormatNumber(rsMonthlyOrders("ORev"),0) %></td>
                            </tr>
                            <%  rsMonthlyOrders.MoveNext: Loop
                                If Not hasMO Then %>
                            <tr><td colspan="3" style="text-align:center;color:#888;padding:20px;">暂无数据</td></tr>
                            <%  End If
                                rsMonthlyOrders.Close
                            End If %>
                        </tbody>
                    </table>
                </div>
                
                <!-- 产品线分布 -->
                <div class="section-card">
                    <h4 class="section-subtitle"><i class="fas fa-layer-group text-orange"></i> 产品线业绩分布</h4>
                    <% Dim totalPLRev: totalPLRev = 0 %>
                    <table class="data-table">
                        <thead><tr><th>产品线</th><th>订单数</th><th>销量</th><th>营收</th><th>占比</th></tr></thead>
                        <tbody>
                            <% If Not rsProdLine Is Nothing Then
                                Dim plTotal, hasPL: hasPL = False
                                Do While Not rsProdLine.EOF
                                    hasPL = True
                                    plTotal = plTotal + SafeNum(rsProdLine("Rev"))
                                    rsProdLine.MoveNext
                                Loop
                                rsProdLine.MoveFirst
                                Do While Not rsProdLine.EOF
                                    Dim pType, pRev
                                    pType = rsProdLine("ProductType")
                                    If IsNull(pType) Then pType = "未分类"
                                    pRev = SafeNum(rsProdLine("Rev"))
                                    totalPLRev = totalPLRev + pRev
                            %>
                            <tr>
                                <td><%= Server.HTMLEncode(pType) %></td>
                                <td><%= rsProdLine("OCnt") %></td>
                                <td><%= rsProdLine("UnitCnt") %></td>
                                <td class="text-green">¥<%= FormatNumber(pRev,0) %></td>
                                <td><%= IIf(plTotal > 0, Round(pRev/plTotal*100,1), 0) %>%</td>
                            </tr>
                            <%  rsProdLine.MoveNext: Loop
                                rsProdLine.Close
                                If Not hasPL Then %>
                            <tr><td colspan="5" style="text-align:center;color:#888;padding:20px;">暂无数据</td></tr>
                            <%  End If
                            End If %>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        
        <!-- ====== 物流 ====== -->
        <div class="flow-section">
            <div class="flow-header flow-logistics">
                <i class="fas fa-truck"></i>
                <div><h3>物流 - 进产销存状态</h3><p>原料采购、香调生产、成品库存、物流配送，全链路可视化</p></div>
            </div>
            <div class="kpi-grid">
                <div class="kpi-card"><div class="lbl">原料库存价值</div><div class="val text-green">¥<%= FormatNumber(rawStockValue,0) %></div><div class="sub"><%= IIf(lowStockCount > 0, lowStockCount & " 项低库存", "库存正常") %></div></div>
                <div class="kpi-card"><div class="lbl">香调库存价值</div><div class="val text-blue">¥<%= FormatNumber(noteStockValue,0) %></div><div class="sub">半成品</div></div>
                <div class="kpi-card"><div class="lbl">成品库存价值</div><div class="val text-orange">¥<%= FormatNumber(prodStockValue,0) %></div><div class="sub">可售成品</div></div>
                <div class="kpi-card"><div class="lbl">总库存价值</div><div class="val text-purple">¥<%= FormatNumber(totalInvValue,0) %></div><div class="sub">占压资金</div></div>
                <div class="kpi-card"><div class="lbl">生产中</div><div class="val text-blue"><%= inProduction %></div><div class="sub">生产订单</div></div>
                <div class="kpi-card"><div class="lbl">待发货</div><div class="val text-orange"><%= pendingShip %></div><div class="sub">销售订单</div></div>
            </div>
        </div>
        
        <!-- ====== 资金流 ====== -->
        <div class="flow-section">
            <div class="flow-header flow-capital">
                <i class="fas fa-dollar-sign"></i>
                <div><h3>资金流 - 财务分析</h3><p>收入成本对比、利润趋势、应付应收，反映资金健康状况</p></div>
            </div>
            <div class="kpi-grid">
                <div class="kpi-card"><div class="lbl">周期营收</div><div class="val text-green">¥<%= FormatNumber(totalRevenue,0) %></div><div class="sub"><%= FormatDateTime(sd,2) %> - <%= FormatDateTime(ed,2) %></div></div>
                <div class="kpi-card"><div class="lbl">周期成本</div><div class="val text-orange">¥<%= FormatNumber(totalCost,0) %></div><div class="sub">已支付订单</div></div>
                <div class="kpi-card"><div class="lbl">周期利润</div><div class="val text-blue">¥<%= FormatNumber(totalProfit,0) %></div><div class="sub">营收-成本</div></div>
                <div class="kpi-card"><div class="lbl">利润率</div><div class="val <%= IIf(profitMargin >= 20, "text-green", IIf(profitMargin >= 10, "text-orange", "text-orange")) %>%"><%= FormatNumber(profitMargin,1) %>%</div><div class="sub"><%= IIf(profitMargin >= 20, "健康", IIf(profitMargin >= 10, "一般", "偏低")) %></div></div>
                <div class="kpi-card"><div class="lbl">应付(未结)</div><div class="val text-orange">¥<%= FormatNumber(pendingPayable,0) %></div><div class="sub">采购订单</div></div>
                <div class="kpi-card"><div class="lbl">应收(在途)</div><div class="val text-green">¥<%= FormatNumber(pendingReceivable,0) %></div><div class="sub">已付未发货</div></div>
            </div>
            
            <!-- 月度利润趋势 -->
            <div class="section-card">
                <h4 class="section-subtitle"><i class="fas fa-chart-line text-green"></i> 月度利润趋势</h4>
                <table class="data-table">
                    <thead><tr><th>月份</th><th>营收</th><th>成本</th><th>利润</th><th>利润率</th></tr></thead>
                    <tbody>
                        <% If Not rsProfitTrend Is Nothing Then
                            Dim hasPT: hasPT = False
                            Do While Not rsProfitTrend.EOF
                                hasPT = True
                                Dim pCost, pProfit, pMargin
                                pRev = SafeNum(rsProfitTrend("Rev"))
                                pCost = SafeNum(rsProfitTrend("Cost"))
                                pProfit = pRev - pCost
                                If pRev > 0 Then pMargin = (pProfit / pRev) * 100 Else pMargin = 0
                        %>
                        <tr>
                            <td><%= rsProfitTrend("Y") %>年<%= rsProfitTrend("M") %>月</td>
                            <td class="text-green">¥<%= FormatNumber(pRev,0) %></td>
                            <td class="text-orange">¥<%= FormatNumber(pCost,0) %></td>
                            <td class="text-blue">¥<%= FormatNumber(pProfit,0) %></td>
                            <td><%= FormatNumber(pMargin,1) %>%</td>
                        </tr>
                        <%  rsProfitTrend.MoveNext: Loop
                            rsProfitTrend.Close
                            If Not hasPT Then %>
                        <tr><td colspan="5" style="text-align:center;color:#888;padding:20px;">暂无数据</td></tr>
                        <%  End If
                        End If %>
                    </tbody>
                </table>
            </div>
        </div>
        
        <!-- ====== 闭环关键指标 ====== -->
        <div class="flow-section">
            <div class="flow-header flow-closed">
                <i class="fas fa-infinity"></i>
                <div><h3>闭环关键指标</h3><p>企业三流闭环效率评估，信息流→物流→资金流的端到端效能</p></div>
            </div>
            <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:15px;">
                <div class="section-card" style="text-align:center;">
                    <div style="font-size:11px;color:#888;margin-bottom:5px;">库存周转率</div>
                    <div style="font-size:32px;font-weight:700;color:#81c784;"><%= FormatNumber(invTurnover,1) %></div>
                    <div style="font-size:11px;color:#888;margin-top:5px;">年化销售成本/平均库存价值</div>
                    <div class="<%= invClass %>" style="margin-top:10px;font-size:12px;">
                        <i class="fas <%= invIcon %>"></i>
                        <%= invText %>
                    </div>
                </div>
                <div class="section-card" style="text-align:center;">
                    <div style="font-size:11px;color:#888;margin-bottom:5px;">订单履约周期</div>
                    <div style="font-size:32px;font-weight:700;color:#64b5f6;"><%= FormatNumber(avgFulfillmentDays,0) %></div>
                    <div style="font-size:11px;color:#888;margin-top:5px;">平均下单到签收天数</div>
                    <div class="<%= fulfillClass %>" style="margin-top:10px;font-size:12px;">
                        <i class="fas <%= fulfillIcon %>"></i>
                        <%= fulfillText %>
                    </div>
                </div>
                <div class="section-card" style="text-align:center;">
                    <div style="font-size:11px;color:#888;margin-bottom:5px;">客户生命周期价值</div>
                    <div style="font-size:32px;font-weight:700;color:#ce93d8;">¥<%= FormatNumber(customerLTV,0) %></div>
                    <div style="font-size:11px;color:#888;margin-top:5px;">每客户平均营收贡献</div>
                    <div class="<%= ltvClass %>" style="margin-top:10px;font-size:12px;">
                        <i class="fas <%= ltvIcon %>"></i>
                        <%= ltvText %>
                    </div>
                </div>
                <div class="section-card" style="text-align:center;">
                    <div style="font-size:11px;color:#888;margin-bottom:5px;">每单平均利润</div>
                    <div style="font-size:32px;font-weight:700;color:#ffb74d;">¥<%= FormatNumber(profitPerOrder,2) %></div>
                    <div style="font-size:11px;color:#888;margin-top:5px;">扣除成本后每单净利润</div>
                    <div class="<%= ppoClass %>" style="margin-top:10px;font-size:12px;">
                        <i class="fas <%= ppoIcon %>"></i>
                        <%= ppoText %>
                    </div>
                </div>
            </div>
            
            <!-- 三流闭环示意图 -->
            <div class="section-card" style="margin-top:20px;">
                <h4 class="section-subtitle"><i class="fas fa-project-diagram text-purple"></i> 企业三流闭环</h4>
                <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:15px;text-align:center;padding:20px 0;">
                    <div style="background:#1a1a2e;border-radius:10px;padding:20px;border-top:3px solid #2196F3;">
                        <i class="fas fa-info-circle" style="font-size:36px;color:#2196F3;margin-bottom:10px;"></i>
                        <h5 style="color:#64b5f6;margin:10px 0;">信息流</h5>
                        <p style="color:#888;font-size:12px;">订单数据 → 生产指令<br>客户需求 → 产品配置<br>销售预测 → 采购计划</p>
                    </div>
                    <div style="background:#1a1a2e;border-radius:10px;padding:20px;border-top:3px solid #FF9800;">
                        <i class="fas fa-truck" style="font-size:36px;color:#FF9800;margin-bottom:10px;"></i>
                        <h5 style="color:#ffb74d;margin:10px 0;">物流</h5>
                        <p style="color:#888;font-size:12px;">原料采购入库<br>香调生产加工<br>成品仓储配送</p>
                    </div>
                    <div style="background:#1a1a2e;border-radius:10px;padding:20px;border-top:3px solid #4CAF50;">
                        <i class="fas fa-dollar-sign" style="font-size:36px;color:#4CAF50;margin-bottom:10px;"></i>
                        <h5 style="color:#81c784;margin:10px 0;">资金流</h5>
                        <p style="color:#888;font-size:12px;">采购付款 → 生产成本<br>销售收入 → 利润核算<br>应付管理 → 应收管理</p>
                    </div>
                </div>
                <div style="text-align:center;padding:10px;background:rgba(0,188,212,0.08);border-radius:8px;">
                    <i class="fas fa-sync-alt" style="color:#00bcd4;margin-right:8px;"></i>
                    <span style="color:#888;font-size:13px;">三流实时同步，数据驱动决策</span>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>