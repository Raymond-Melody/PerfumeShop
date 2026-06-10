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
Function GeneratePurchaseNo()
    Dim today, prefix, countNum
    today = Date()
    prefix = "FBPO-" & Year(today) & Right("0" & Month(today), 2) & Right("0" & Day(today), 2) & "-"
    countNum = SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandPurchaseOrders WHERE PurchaseNo LIKE '" & prefix & "%'"))
    GeneratePurchaseNo = prefix & Right("000" & (countNum + 1), 3)
End Function

' ========== 双模式参数推算函数 ==========

' 从 FixedBrandCostAllocation 表按时间窗口统计日均销量
Function GetStatisticalDailySales(fpid, days)
    Dim sql : sql = "SELECT ISNULL(SUM(Quantity),0) / CAST(" & days & " AS DECIMAL(10,2)) FROM FixedBrandCostAllocation WHERE FixedProductID=" & fpid & " AND AllocatedAt >= DATEADD(DAY, -" & days & ", GETDATE())"
    GetStatisticalDailySales = SafeNum(GetScalar(sql))
End Function

' 从历史收货记录计算平均交货天数（取近12个月）
Function GetStatisticalLeadTime(fpid)
    Dim sql : sql = "SELECT ISNULL(AVG(CAST(DATEDIFF(DAY, po.OrderDate, r.ReceiptDate) AS DECIMAL(10,2))),0) FROM FixedBrandReceipts r JOIN FixedBrandPurchaseOrders po ON r.PurchaseID=po.PurchaseID JOIN FixedBrandReceiptDetails rd ON r.ReceiptID=rd.ReceiptID WHERE rd.FixedProductID=" & fpid & " AND r.ReceiptDate >= DATEADD(MONTH, -12, GETDATE())"
    Dim val : val = SafeNum(GetScalar(sql))
    If val <= 0 Then val = 7
    GetStatisticalLeadTime = CInt(val)
End Function

' 计算近12个月中有销售记录的月份数
Function GetConsecutiveDataMonths(fpid)
    Dim sql : sql = "SELECT COUNT(DISTINCT CONVERT(NVARCHAR(7), AllocatedAt, 120)) FROM FixedBrandCostAllocation WHERE FixedProductID=" & fpid & " AND AllocatedAt >= DATEADD(MONTH, -12, GETDATE())"
    GetConsecutiveDataMonths = SafeNum(GetScalar(sql))
End Function

' 统计推算安全库存：基于需求波动（日均销量 × 交货周期 × 缓冲系数）
Function GetStatisticalSafetyStock(fpid, dailySales, leadDays)
    ' 缓冲系数基准 0.5，可根据实际波动调整
    Dim safety : safety = CInt(CDbl(dailySales) * CDbl(leadDays) * 0.5)
    If safety < 5 Then safety = 5
    GetStatisticalSafetyStock = safety
End Function

' 判断统计模式是否可用：连续数据月数 >= 阈值
Function IsStatisticalModeReady(consecutiveMonths, thresholdMonths)
    If consecutiveMonths >= thresholdMonths Then
        IsStatisticalModeReady = True
    Else
        IsStatisticalModeReady = False
    End If
End Function

' ========== 消息 ==========
Dim msg, msgType
msg = ""
msgType = "success"

' ========== 参数配置 ==========
Dim replenishDays : replenishDays = SafeNum(Request.QueryString("days"))
If replenishDays <= 0 Then replenishDays = 30
Dim autoGen : autoGen = Trim(Request.QueryString("auto"))
' 统计模式最低数据月数阈值（默认6个月，可由URL参数调整）
Dim statThreshold : statThreshold = SafeNum(Request.QueryString("threshold"))
If statThreshold <= 0 Then statThreshold = 3

' ========== POST: 一键生成补货单 ==========
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not ValidateCSRFToken() Then
        msg = "安全令牌验证失败"
        msgType = "error"
    Else
        Dim postAction : postAction = Trim(Request.Form("action"))
        
        If postAction = "bulk_generate" Then
            If Not isManager Then
                msg = "权限不足：仅管理员和采购经理可生成补货单"
                msgType = "error"
            Else
                Dim genCount : genCount = SafeNum(Request.Form("gen_count"))
            Dim generatedOrders : generatedOrders = 0
            
            If genCount > 0 Then
                Dim gi
                For gi = 1 To genCount
                    Dim gFPID : gFPID = SafeNum(Request.Form("gen_fpid_" & gi))
                    Dim gSID : gSID = SafeNum(Request.Form("gen_supplier_" & gi))
                    Dim gSName : gSName = Trim(Request.Form("gen_supplier_name_" & gi))
                    Dim gPName : gPName = SafeSQL(Trim(Request.Form("gen_name_" & gi)))
                    Dim gSpec : gSpec = SafeSQL(Trim(Request.Form("gen_spec_" & gi)))
                    Dim gQty : gQty = SafeNum(Request.Form("gen_qty_" & gi))
                    Dim gPrice : gPrice = SafeNum(Request.Form("gen_price_" & gi))
                    
                    If gFPID > 0 And gQty > 0 Then
                        Call BeginTransaction()
                        
                        Dim gPONo : gPONo = GeneratePurchaseNo()
                        Dim gTotal : gTotal = gQty * gPrice
                        
                        Dim insGOrder : insGOrder = "INSERT INTO FixedBrandPurchaseOrders (PurchaseNo, SupplierID, SupplierName, TotalAmount, Status, CreatedBy) VALUES ('" & _
                            gPONo & "', " & IIf(gSID > 0, gSID, "NULL") & ", '" & SafeSQL(gSName) & "', " & gTotal & ", 'Draft', '" & SafeSQL(Session("AdminName")) & "')"
                        
                        If ExecuteNonQuery(insGOrder) Then
                            Dim gNewPID : gNewPID = SafeNum(GetScalar("SELECT MAX(PurchaseID) FROM FixedBrandPurchaseOrders"))
                            If gNewPID > 0 Then
                                If ExecuteNonQuery("INSERT INTO FixedBrandPurchaseDetails (PurchaseID, FixedProductID, ProductName, Specification, Quantity, UnitPrice, SubTotal) VALUES (" & gNewPID & ", " & gFPID & ", '" & gPName & "', '" & gSpec & "', " & gQty & ", " & gPrice & ", " & gTotal & ")") Then
                                    Call CommitTransaction()
                                    generatedOrders = generatedOrders + 1
                                Else
                                    Call RollbackTransaction()
                                End If
                            Else
                                Call RollbackTransaction()
                            End If
                        Else
                            Call RollbackTransaction()
                        End If
                    End If
                Next
                
                If generatedOrders > 0 Then
                    msg = "成功生成 " & generatedOrders & " 个补货采购单"
                    msgType = "success"
                Else
                    msg = "未生成任何补货单，请检查数据"
                    msgType = "error"
                End If
            End If
            End If
        End If
    End If
End If

' ========== 查询低库存产品 ==========
Dim sqlLowStock : sqlLowStock = "SELECT fp.FixedProductID, fp.ProductCode, fp.ProductName, fp.Specification, fp.UnitPrice, fp.SupplierID, fp.SupplierName, fp.MinOrderQty, fp.LeadTimeDays, fp.SafetyStockManual, fp.LeadTimeDaysManual, " & _
    "ISNULL(fi.StockQty,0) AS StockQty, ISNULL(fi.SafetyStock,10) AS SafetyStock, ISNULL(fi.TotalSold,0) AS TotalSold, " & _
    "ISNULL(fi.ParamMode,'Manual') AS ParamMode, ISNULL(fi.DailySalesAvg,0) AS DailySalesAvg, ISNULL(fi.ConsecutiveDataMonths,0) AS ConsecutiveDataMonths " & _
    "FROM FixedBrandProducts fp LEFT JOIN FixedBrandInventory fi ON fp.FixedProductID=fi.FixedProductID " & _
    "WHERE fp.Status='Active' AND ISNULL(fi.StockQty,0) <= ISNULL(fi.SafetyStock,10) " & _
    "ORDER BY (ISNULL(fi.StockQty,0) * 1.0 / NULLIF(ISNULL(fi.SafetyStock,10),0)) ASC"

Dim rsLowStock : Set rsLowStock = conn.Execute(sqlLowStock)

' ========== 统计 ==========
Dim totalProducts : totalProducts = SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandProducts WHERE Status='Active'"))
Dim lowStockCount : lowStockCount = SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandInventory WHERE StockQty <= SafetyStock"))
Dim outOfStock : outOfStock = SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandInventory WHERE StockQty <= 0"))
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>品牌定香智能补货 - 采购管理中心</title>
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
        
        .stats-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 10px; padding: 18px; border: 1px solid rgba(255,255,255,0.05); }
        .stat-icon { width: 40px; height: 40px; border-radius: 8px; display: flex; align-items: center; justify-content: center; font-size: 16px; margin-bottom: 10px; }
        .stat-value { font-size: 22px; font-weight: 700; color: #fff; }
        .stat-label { font-size: 12px; color: #888; margin-top: 4px; }
        
        .config-bar { display: flex; gap: 12px; align-items: center; margin-bottom: 20px; padding: 15px 20px; background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 10px; border: 1px solid rgba(255,255,255,0.05); }
        .config-bar label { font-size: 13px; color: #888; }
        .config-bar select { padding: 8px 14px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.1); background: #1a1a2e; color: #e0e0e0; font-size: 13px; }
        
        .data-table { width: 100%; border-collapse: collapse; background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 10px; overflow: hidden; }
        .data-table th, .data-table td { padding: 12px 14px; text-align: left; border-bottom: 1px solid rgba(255,255,255,0.05); font-size: 13px; }
        .data-table th { color: #888; font-size: 11px; text-transform: uppercase; font-weight: 600; background: rgba(0,0,0,0.2); }
        .data-table td { color: #ccc; }
        .data-table tr:hover td { background: rgba(255,255,255,0.03); }
        
        .stock-critical { color: #F44336; font-weight: 700; }
        .stock-low { color: #FF9800; font-weight: 600; }
        .stock-warning { color: #FFC107; }
        .stock-ok { color: #4CAF50; }
        
        .suggest-qty { display: inline-block; padding: 2px 10px; border-radius: 12px; background: rgba(76,175,80,0.15); color: #4CAF50; font-weight: 600; }
        .qty-input { padding: 6px 10px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.1); background: #1a1a2e; color: #e0e0e0; width: 70px; text-align: center; font-size: 13px; }
        
        .chart-bar { height: 20px; border-radius: 10px; min-width: 4px; transition: width 0.4s ease; }
        
        .mode-badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 10px; font-weight: 600; }
        .mode-manual { background: rgba(255,152,0,0.2); color: #FF9800; }
        .mode-auto { background: rgba(76,175,80,0.2); color: #4CAF50; }
        .mode-fallback { background: rgba(244,67,54,0.2); color: #F44336; }
    </style>
</head>
<body data-theme="purchase-dark">
    <!--#include file="../includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-robot"></i> 品牌定香智能补货</h2>
            <div class="breadcrumb" style="font-size:13px;color:#888;">
                <a href="index.asp" style="color:#FF9800;text-decoration:none;">品牌定香采购</a> / 智能补货
            </div>
        </div>
        
        <% If msg <> "" Then %>
        <div style="padding:12px 20px; border-radius:8px; margin-bottom:20px; font-size:14px; background:<%=IIf(msgType="success","rgba(76,175,80,0.15)","rgba(244,67,54,0.15)")%>; color:<%=IIf(msgType="success","#4CAF50","#F44336")%>; border:1px solid <%=IIf(msgType="success","rgba(76,175,80,0.3)","rgba(244,67,54,0.3)")%>;">
            <i class="fas fa-<%=IIf(msgType="success","check-circle","exclamation-circle")%>"></i> <%= msg %>
        </div>
        <% End If %>
        
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-icon" style="background:linear-gradient(135deg,#4CAF50,#388E3C);"><i class="fas fa-cubes"></i></div>
                <div class="stat-value"><%= totalProducts %></div>
                <div class="stat-label">活跃产品总数</div>
            </div>
            <div class="stat-card" style="border:1px solid rgba(255,152,0,0.3);">
                <div class="stat-icon" style="background:linear-gradient(135deg,#FF9800,#F57C00);"><i class="fas fa-exclamation-triangle"></i></div>
                <div class="stat-value" style="color:#FF9800;"><%= lowStockCount %></div>
                <div class="stat-label">低库存产品</div>
            </div>
            <div class="stat-card" style="border:1px solid rgba(244,67,54,0.3);">
                <div class="stat-icon" style="background:linear-gradient(135deg,#F44336,#C62828);"><i class="fas fa-times-circle"></i></div>
                <div class="stat-value" style="color:#F44336;"><%= outOfStock %></div>
                <div class="stat-label">已断货产品</div>
            </div>
        </div>
        
        <div class="config-bar">
            <label><i class="fas fa-chart-line"></i> 销量统计周期：</label>
            <form method="get" style="display:flex;gap:10px;align-items:center;">
                <select name="days" onchange="this.form.submit()">
                    <option value="7" <% If replenishDays=7 Then %>selected<% End If %>>近7天</option>
                    <option value="14" <% If replenishDays=14 Then %>selected<% End If %>>近14天</option>
                    <option value="30" <% If replenishDays=30 Then %>selected<% End If %>>近30天</option>
                    <option value="60" <% If replenishDays=60 Then %>selected<% End If %>>近60天</option>
                    <option value="90" <% If replenishDays=90 Then %>selected<% End If %>>近90天</option>
                </select>
                <select name="threshold" onchange="this.form.submit()" title="统计模式最低数据月数">
                    <option value="3" <% If statThreshold=3 Then %>selected<% End If %>>需≥3月数据</option>
                    <option value="6" <% If statThreshold=6 Then %>selected<% End If %>>需≥6月数据</option>
                    <option value="12" <% If statThreshold=12 Then %>selected<% End If %>>需≥12月数据</option>
                </select>
            </form>
            <span style="color:#666;font-size:12px;margin-left:auto;">
                <i class="fas fa-cog" style="color:#FF9800;"></i> 补货建议 = (日均销量 × 交货周期) + 安全库存 - 当前库存
            </span>
            <span style="font-size:11px;color:#888;"> | 模式：</span>
            <span class="mode-badge mode-manual"><i class="fas fa-user-edit"></i> 人工</span>
            <span class="mode-badge mode-auto" style="margin-left:4px;"><i class="fas fa-database"></i> 统计</span>
        </div>
        
        <% If Not rsLowStock Is Nothing Then
            If Not rsLowStock.EOF Then
                Dim itemIdx : itemIdx = 0
        %>
        <form method="post">
            <input type="hidden" name="csrf_token" value="<%= Session("CSRFToken") %>">
            <input type="hidden" name="action" value="bulk_generate">
            
            <table class="data-table">
                <thead>
                    <tr>
                        <th>产品编码</th>
                        <th>产品名称</th>
                        <th>当前库存</th>
                        <th>安全库存</th>
                        <th>库存率</th>
                        <th>模式</th>
                        <th>日均销量</th>
                        <th>交货周期</th>
                        <th>建议补货</th>
                        <th>批量补货数量</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                        Do While Not rsLowStock.EOF
                            itemIdx = itemIdx + 1
                            Dim fpID : fpID = SafeNum(rsLowStock("FixedProductID"))
                            Dim stkQty : stkQty = SafeNum(rsLowStock("StockQty"))
                            Dim sfStk : sfStk = SafeNum(rsLowStock("SafetyStock"))
                            Dim minQty : minQty = SafeNum(rsLowStock("MinOrderQty"))
                            Dim leadDays : leadDays = SafeNum(rsLowStock("LeadTimeDays"))
                            Dim totalSold : totalSold = SafeNum(rsLowStock("TotalSold"))
                            Dim unitPrice : unitPrice = SafeNum(rsLowStock("UnitPrice"))
                            Dim pName : pName = CStr(rsLowStock("ProductName"))
                            Dim pCode : pCode = CStr(rsLowStock("ProductCode"))
                            Dim pSpec : pSpec = CStr(rsLowStock("Specification") & "")
                            Dim sID : sID = SafeNum(rsLowStock("SupplierID"))
                            Dim sName : sName = CStr(rsLowStock("SupplierName") & "")
                            
                            ' 双模式参数
                            Dim paramMode : paramMode = CStr(rsLowStock("ParamMode") & "")
                            If paramMode = "" Then paramMode = "Manual"
                            Dim savedDailyAvg : savedDailyAvg = SafeNum(rsLowStock("DailySalesAvg"))
                            Dim consecutiveMonths : consecutiveMonths = SafeNum(rsLowStock("ConsecutiveDataMonths"))
                            Dim safetyStockManual : safetyStockManual = SafeNum(rsLowStock("SafetyStockManual"))
                            Dim leadDaysManual : leadDaysManual = SafeNum(rsLowStock("LeadTimeDaysManual"))
                            
                            ' 库存率
                            Dim stockRatio : stockRatio = 0
                            If sfStk > 0 Then stockRatio = (stkQty / sfStk) * 100
                            If stockRatio > 100 Then stockRatio = 100
                            
                            ' ========== 双模式参数计算 ==========
                            Dim useStatistical : useStatistical = False
                            Dim actualDailySales : actualDailySales = 0
                            Dim actualLeadDays : actualLeadDays = leadDays
                            Dim actualSafetyStock : actualSafetyStock = sfStk
                            Dim modeLabel : modeLabel = "人工"
                            Dim modeIcon : modeIcon = "fa-user-edit"
                            Dim modeClass : modeClass = "mode-manual"
                            
                            If paramMode = "Auto" Then
                                ' 检查是否有足够的历史数据
                                consecutiveMonths = GetConsecutiveDataMonths(fpID)
                                If IsStatisticalModeReady(consecutiveMonths, statThreshold) Then
                                    useStatistical = True
                                    ' 从 FixedBrandCostAllocation 获取实际日均销量
                                    actualDailySales = GetStatisticalDailySales(fpID, replenishDays)
                                    ' 从历史收货记录推算交货周期
                                    actualLeadDays = GetStatisticalLeadTime(fpID)
                                    ' 基于需求波动计算安全库存
                                    actualSafetyStock = GetStatisticalSafetyStock(fpID, actualDailySales, actualLeadDays)
                                    ' 仅当距上次计算超过1天才更新持久化统计值
                                    Dim lastCalcDate : lastCalcDate = GetScalar("SELECT ISNULL(LastAutoCalcDate, '1900-01-01') FROM FixedBrandInventory WHERE FixedProductID=" & fpID)
                                    If DateDiff("d", lastCalcDate, Now()) >= 1 Then
                                        Call ExecuteNonQuery("UPDATE FixedBrandInventory SET DailySalesAvg=" & actualDailySales & ", ConsecutiveDataMonths=" & consecutiveMonths & ", LastAutoCalcDate=GETDATE() WHERE FixedProductID=" & fpID)
                                    End If
                                    modeLabel = "统计(" & consecutiveMonths & "月)"
                                    modeIcon = "fa-database"
                                    modeClass = "mode-auto"
                                Else
                                    ' 自动模式但数据不足，降级使用人工参数
                                    modeLabel = "人工(数据不足)"
                                    modeIcon = "fa-exclamation-triangle"
                                    modeClass = "mode-fallback"
                                    actualDailySales = (totalSold / replenishDays)
                                    If actualDailySales < 0 Then actualDailySales = 0
                                    If safetyStockManual > 0 Then actualSafetyStock = safetyStockManual
                                    If leadDaysManual > 0 Then actualLeadDays = leadDaysManual
                                End If
                            Else
                                ' 人工模式：使用人工设定值，日均销量从成本分摊表获取
                                actualDailySales = GetStatisticalDailySales(fpID, replenishDays)
                                If actualDailySales <= 0 Then actualDailySales = 0
                                If savedDailyAvg > 0 Then actualDailySales = savedDailyAvg
                                If safetyStockManual > 0 Then actualSafetyStock = safetyStockManual
                                If leadDaysManual > 0 Then actualLeadDays = leadDaysManual
                            End If
                            
                            ' 建议补货量 = (日均销量 * 交货周期) + 安全库存 - 当前库存
                            Dim suggestQty : suggestQty = CInt((actualDailySales * actualLeadDays) + actualSafetyStock - stkQty)
                            If suggestQty < minQty Then suggestQty = minQty
                            If suggestQty < 1 Then suggestQty = minQty
                            
                            Dim stockClass : stockClass = "stock-ok"
                            If stkQty <= 0 Then
                                stockClass = "stock-critical"
                            ElseIf stockRatio < 50 Then
                                stockClass = "stock-low"
                            ElseIf stockRatio < 100 Then
                                stockClass = "stock-warning"
                            End If
                    %>
                    <tr>
                        <td><span style="font-family:Consolas,monospace;color:#FF9800;"><%= Server.HTMLEncode(pCode) %></span></td>
                        <td><%= Server.HTMLEncode(pName) %></td>
                        <td class="<%= stockClass %>"><%= stkQty %></td>
                        <td><%= actualSafetyStock %><% If useStatistical Then %><span style="color:#4CAF50;font-size:10px;"> (推算)</span><% End If %></td>
                        <td>
                            <div style="display:flex;align-items:center;gap:6px;">
                                <div style="flex:1;height:8px;background:rgba(255,255,255,0.06);border-radius:4px;overflow:hidden;">
                                    <div class="chart-bar" style="width:<%= stockRatio %>%;background:<%=IIf(stockRatio<=0,"#F44336",IIf(stockRatio<50,"#FF9800",IIf(stockRatio<100,"#FFC107","#4CAF50")))%>;"></div>
                                </div>
                                <span style="font-size:11px;color:#888;"><%= CInt(stockRatio) %>%</span>
                            </div>
                        </td>
                        <td><span class="mode-badge <%= modeClass %>" title="<%= modeLabel %>"><i class="fas <%= modeIcon %>"></i></span></td>
                        <td><%= FormatNumber(actualDailySales, 1) %>/天</td>
                        <td><%= actualLeadDays %>天<% If useStatistical Then %><span style="color:#4CAF50;font-size:10px;">(推算)</span><% End If %></td>
                        <td><span class="suggest-qty"><%= suggestQty %></span></td>
                        <td>
                            <input type="hidden" name="gen_fpid_<%= itemIdx %>" value="<%= fpID %>">
                            <input type="hidden" name="gen_supplier_<%= itemIdx %>" value="<%= sID %>">
                            <input type="hidden" name="gen_supplier_name_<%= itemIdx %>" value="<%= Server.HTMLEncode(sName) %>">
                            <input type="hidden" name="gen_name_<%= itemIdx %>" value="<%= Server.HTMLEncode(pName) %>">
                            <input type="hidden" name="gen_spec_<%= itemIdx %>" value="<%= Server.HTMLEncode(pSpec) %>">
                            <input type="hidden" name="gen_price_<%= itemIdx %>" value="<%= unitPrice %>">
                            <input type="number" name="gen_qty_<%= itemIdx %>" class="qty-input" value="<%= suggestQty %>" min="<%= minQty %>">
                            <span style="font-size:11px;color:#666;margin-left:4px;">¥<%= FormatNumber(unitPrice*suggestQty,0) %></span>
                        </td>
                    </tr>
                    <%
                            rsLowStock.MoveNext
                        Loop
                    %>
                    <input type="hidden" name="gen_count" value="<%= itemIdx %>">
                </tbody>
            </table>
            
            <div style="margin-top:20px;text-align:right;">
                <button type="submit" class="btn btn--primary" onclick="return confirm('确定要一键生成 ' + <%= itemIdx %> + ' 个补货采购单吗？')">
                    <i class="fas fa-magic"></i> 一键生成补货单
                </button>
            </div>
        </form>
        <%  Else %>
        <div style="text-align:center;padding:60px;color:#666;">
            <i class="fas fa-check-circle" style="font-size:48px;display:block;margin-bottom:15px;color:#4CAF50;"></i>
            <p>所有产品库存充足，无需补货</p>
        </div>
        <%  End If
            rsLowStock.Close
            Set rsLowStock = Nothing
        End If %>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
