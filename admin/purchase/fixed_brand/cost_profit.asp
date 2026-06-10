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

' ========== 消息 ==========
Dim msg, msgType
msg = ""
msgType = "success"

' ========== POST: 更新利润计算 ==========
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not ValidateCSRFToken() Then
        msg = "安全令牌验证失败"
        msgType = "error"
    ElseIf Trim(Request.Form("action")) = "recalc" Then
        Call ExecuteNonQuery("UPDATE FixedBrandCostAllocation SET ProfitAmount = ISNULL(SalePrice,0) * ISNULL(Quantity,0) - ISNULL(TotalCost,0), ProfitRate = CASE WHEN ISNULL(SalePrice,0) * ISNULL(Quantity,0) > 0 THEN ((ISNULL(SalePrice,0) * ISNULL(Quantity,0) - ISNULL(TotalCost,0)) / (ISNULL(SalePrice,0) * ISNULL(Quantity,0))) * 100 ELSE 0 END")
        msg = "利润数据已重新计算"
        msgType = "success"
    ElseIf Trim(Request.Form("action")) = "sync_to_products" Then
        ' 将 FixedBrandInventory 的加权平均成本同步到 Products.UnitCost
        Dim syncCount : syncCount = 0
        Dim rsSync
        Set rsSync = conn.Execute("SELECT fbp.ProductID, ISNULL(fbi.AvgUnitCost, fbp.UnitPrice) AS Cost FROM FixedBrandProducts fbp LEFT JOIN FixedBrandInventory fbi ON fbp.FixedProductID = fbi.FixedProductID WHERE fbp.ProductID IS NOT NULL AND fbp.ProductID > 0 AND fbp.Status = 'Active'")
        If Not rsSync Is Nothing Then
            Do While Not rsSync.EOF
                Dim syncPID : syncPID = SafeNum(rsSync("ProductID"))
                Dim syncCost : syncCost = SafeNum(rsSync("Cost"))
                If syncPID > 0 And syncCost > 0 Then
                    Call ExecuteNonQuery("UPDATE Products SET UnitCost=" & syncCost & ", UpdatedAt=GETDATE() WHERE ProductID=" & syncPID)
                    syncCount = syncCount + 1
                End If
                rsSync.MoveNext
            Loop
            rsSync.Close
        End If
        Set rsSync = Nothing
        msg = "已同步 " & syncCount & " 个品牌定香产品成本到产品库"
        msgType = "success"
    End If
End If

' ========== 总览统计 ==========
Dim totalCost : totalCost = SafeNum(GetScalar("SELECT ISNULL(SUM(TotalCost),0) FROM FixedBrandCostAllocation"))
Dim totalRevenue : totalRevenue = SafeNum(GetScalar("SELECT ISNULL(SUM(SalePrice * Quantity),0) FROM FixedBrandCostAllocation"))
Dim totalProfit : totalProfit = SafeNum(GetScalar("SELECT ISNULL(SUM(ProfitAmount),0) FROM FixedBrandCostAllocation"))
Dim avgProfitRate : avgProfitRate = SafeNum(GetScalar("SELECT ISNULL(AVG(ProfitRate),0) FROM FixedBrandCostAllocation WHERE Quantity > 0"))

Dim totalPurchases : totalPurchases = SafeNum(GetScalar("SELECT ISNULL(SUM(TotalAmount),0) FROM FixedBrandPurchaseOrders WHERE Status IN ('Received','Completed')"))
Dim totalInventory : totalInventory = SafeNum(GetScalar("SELECT ISNULL(SUM(StockQty),0) FROM FixedBrandInventory"))
Dim inventoryValue : inventoryValue = SafeNum(GetScalar("SELECT ISNULL(SUM(StockQty * AvgUnitCost),0) FROM FixedBrandInventory"))

' ========== 按产品汇总成本 ==========
Dim sqlProductCost : sqlProductCost = "SELECT fp.FixedProductID, fp.ProductCode, fp.ProductName, fp.Specification, fp.UnitPrice AS CurrentPrice, fp.SalePrice, " & _
    "ISNULL(fi.StockQty,0) AS StockQty, ISNULL(fi.AvgUnitCost,0) AS AvgCost, ISNULL(fi.TotalPurchased,0) AS TotalPurchased, ISNULL(fi.TotalSold,0) AS TotalSold, ISNULL(fi.ParamMode,'Manual') AS ParamMode, " & _
    "ISNULL(ca.TotalCost,0) AS AllocatedCost, ISNULL(ca.TotalRevenue,0) AS Revenue, ISNULL(ca.TotalProfit,0) AS Profit " & _
    "FROM FixedBrandProducts fp " & _
    "LEFT JOIN FixedBrandInventory fi ON fp.FixedProductID=fi.FixedProductID " & _
    "LEFT JOIN (SELECT FixedProductID, ISNULL(SUM(TotalCost),0) AS TotalCost, ISNULL(SUM(SalePrice * Quantity),0) AS TotalRevenue, ISNULL(SUM(ProfitAmount),0) AS TotalProfit FROM FixedBrandCostAllocation GROUP BY FixedProductID) ca ON fp.FixedProductID=ca.FixedProductID " & _
    "WHERE fp.Status='Active' ORDER BY ISNULL(ca.TotalProfit,0) DESC"

Dim rsProductCost : Set rsProductCost = conn.Execute(sqlProductCost)

' ========== 成本分摊明细(最近30条) ==========
Dim sqlAllocation : sqlAllocation = "SELECT TOP 30 * FROM FixedBrandCostAllocation ORDER BY AllocatedAt DESC"
Dim rsAllocation : Set rsAllocation = conn.Execute(sqlAllocation)

' ========== 采购成本趋势(近12个月) ==========
Dim monthlyCosts(11), monthlyLabels(11)
Dim m
For m = 0 To 11
    monthlyLabels(m) = Month(DateAdd("m", -11 + m, Date())) & "月"
    Dim monthStart : monthStart = DateSerial(Year(DateAdd("m", -11 + m, Date())), Month(DateAdd("m", -11 + m, Date())), 1)
    Dim monthEnd : monthEnd = DateSerial(Year(DateAdd("m", -10 + m, Date())), Month(DateAdd("m", -10 + m, Date())), 1)
    monthlyCosts(m) = SafeNum(GetScalar("SELECT ISNULL(SUM(TotalAmount),0) FROM FixedBrandPurchaseOrders WHERE (Status='Received' OR Status='Completed') AND OrderDate >= '" & FormatDateTime(monthStart, 2) & "' AND OrderDate < '" & FormatDateTime(monthEnd, 2) & "'"))
Next

Dim maxMonthlyCost : maxMonthlyCost = 1
For m = 0 To 11
    If monthlyCosts(m) > maxMonthlyCost Then maxMonthlyCost = monthlyCosts(m)
Next
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>成本追踪与利润分析 - 采购管理中心</title>
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
        
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 10px; padding: 18px; border: 1px solid rgba(255,255,255,0.05); }
        .stat-icon { width: 40px; height: 40px; border-radius: 8px; display: flex; align-items: center; justify-content: center; font-size: 16px; margin-bottom: 10px; }
        .stat-value { font-size: 22px; font-weight: 700; color: #fff; }
        .stat-label { font-size: 12px; color: #888; margin-top: 4px; }
        
        .grid-2col { display: grid; grid-template-columns: 1fr 1fr; gap: 25px; margin-bottom: 25px; }
        .panel { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 10px; padding: 20px; border: 1px solid rgba(255,255,255,0.05); }
        .panel h3 { color: #fff; font-size: 16px; margin: 0 0 15px; display: flex; align-items: center; gap: 8px; }
        
        .data-table { width: 100%; border-collapse: collapse; }
        .data-table th, .data-table td { padding: 10px 12px; text-align: left; border-bottom: 1px solid rgba(255,255,255,0.05); font-size: 13px; }
        .data-table th { color: #888; font-size: 11px; text-transform: uppercase; font-weight: 600; }
        .data-table td { color: #ccc; }
        .data-table tr:hover td { background: rgba(255,255,255,0.02); }
        
        .profit-positive { color: #4CAF50; font-weight: 600; }
        .profit-negative { color: #F44336; font-weight: 600; }
        .profit-zero { color: #888; }
        
        .bar-container { margin-bottom: 8px; }
        .bar-label { display: flex; justify-content: space-between; font-size: 11px; color: #888; margin-bottom: 3px; }
        .bar-track { height: 22px; background: rgba(255,255,255,0.04); border-radius: 4px; overflow: hidden; position: relative; }
        .bar-fill { height: 100%; border-radius: 4px; transition: width 0.5s ease; position: absolute; left: 0; top: 0; }
        .bar-value { position: absolute; right: 8px; top: 50%; transform: translateY(-50%); font-size: 11px; font-weight: 600; color: #fff; }
    </style>
</head>
<body data-theme="purchase-dark">
    <!--#include file="../includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-chart-pie"></i> 成本追踪与利润分析</h2>
            <div class="breadcrumb" style="font-size:13px;color:#888;">
                <a href="index.asp" style="color:#FF9800;text-decoration:none;">品牌定香采购</a> / 成本利润
            </div>
        </div>
        
        <% If msg <> "" Then %>
        <div style="padding:12px 20px; border-radius:8px; margin-bottom:20px; font-size:14px; background:<%=IIf(msgType="success","rgba(76,175,80,0.15)","rgba(244,67,54,0.15)")%>; color:<%=IIf(msgType="success","#4CAF50","#F44336")%>; border:1px solid <%=IIf(msgType="success","rgba(76,175,80,0.3)","rgba(244,67,54,0.3)")%>;">
            <i class="fas fa-<%=IIf(msgType="success","check-circle","exclamation-circle")%>"></i> <%= msg %>
        </div>
        <% End If %>
        
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-icon" style="background:linear-gradient(135deg,#FF9800,#F57C00);"><i class="fas fa-yen-sign"></i></div>
                <div class="stat-value">¥<%= FormatNumber(totalCost, 0) %></div>
                <div class="stat-label">已分摊总成本</div>
            </div>
            <div class="stat-card">
                <div class="stat-icon" style="background:linear-gradient(135deg,#2196F3,#1565C0);"><i class="fas fa-chart-line"></i></div>
                <div class="stat-value">¥<%= FormatNumber(totalRevenue, 0) %></div>
                <div class="stat-label">预期收入</div>
            </div>
            <div class="stat-card">
                <div class="stat-icon" style="background:linear-gradient(135deg,#4CAF50,#388E3C);"><i class="fas fa-coins"></i></div>
                <div class="stat-value" style="color:<%=IIf(totalProfit>=0,"#4CAF50","#F44336")%>;">¥<%= FormatNumber(totalProfit, 0) %></div>
                <div class="stat-label">总利润</div>
            </div>
            <div class="stat-card">
                <div class="stat-icon" style="background:linear-gradient(135deg,#9C27B0,#6A1B9A);"><i class="fas fa-percentage"></i></div>
                <div class="stat-value"><%= FormatNumber(avgProfitRate, 1) %>%</div>
                <div class="stat-label">平均利润率</div>
            </div>
        </div>
        
        <!-- 库存价值卡片 -->
        <div class="stats-grid" style="grid-template-columns:repeat(3,1fr);">
            <div class="stat-card">
                <div class="stat-value">¥<%= FormatNumber(totalPurchases, 0) %></div>
                <div class="stat-label">累计采购总额</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%= totalInventory %></div>
                <div class="stat-label">当前总库存(件)</div>
            </div>
            <div class="stat-card">
                <div class="stat-value">¥<%= FormatNumber(inventoryValue, 0) %></div>
                <div class="stat-label">库存资产价值</div>
            </div>
        </div>
        
        <div class="grid-2col">
            <!-- 按产品利润排行 -->
            <div class="panel">
                <h3><i class="fas fa-trophy" style="color:#FF9800;"></i> 产品利润排行</h3>
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>产品名称</th>
                            <th>参数</th>
                            <th>成本</th>
                            <th>收入</th>
                            <th>利润</th>
                            <th>利润率</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% If Not rsProductCost Is Nothing Then
                            If Not rsProductCost.EOF Then
                                Dim pcIdx : pcIdx = 0
                                Do While Not rsProductCost.EOF And pcIdx < 10
                                    pcIdx = pcIdx + 1
                                    Dim pcCost : pcCost = SafeNum(rsProductCost("AllocatedCost"))
                                    Dim pcRev : pcRev = SafeNum(rsProductCost("Revenue"))
                                    Dim pcProfit : pcProfit = SafeNum(rsProductCost("Profit"))
                                    Dim pcRate : pcRate = 0
                                    If pcRev > 0 Then pcRate = (pcProfit / pcRev) * 100
                        %>
                        <tr>
                            <td><%= Server.HTMLEncode(CStr(rsProductCost("ProductName"))) %></td>
                            <td><% 
                                Dim pcMode : pcMode = CStr(rsProductCost("ParamMode") & "")
                                If pcMode = "" Then pcMode = "Manual"
                                If pcMode = "Auto" Then
                                    Response.Write "<span style='color:#4CAF50;font-size:11px;'><i class='fas fa-database'></i> 统计</span>"
                                Else
                                    Response.Write "<span style='color:#FF9800;font-size:11px;'><i class='fas fa-user-edit'></i> 人工</span>"
                                End If
                            %></td>
                            <td>¥<%= FormatNumber(pcCost, 2) %></td>
                            <td>¥<%= FormatNumber(pcRev, 2) %></td>
                            <td class="<%=IIf(pcProfit>=0,"profit-positive","profit-negative")%>">¥<%= FormatNumber(pcProfit, 2) %></td>
                            <td class="<%=IIf(pcRate>=0,"profit-positive","profit-negative")%>"><%= FormatNumber(pcRate, 1) %>%</td>
                        </tr>
                        <%
                                    rsProductCost.MoveNext
                                Loop
                            Else
                        %>
                        <tr><td colspan="6" style="text-align:center;color:#666;padding:30px;">暂无成本数据</td></tr>
                        <%  End If
                        End If %>
                    </tbody>
                </table>
            </div>
            
            <!-- 采购成本月度趋势 -->
            <div class="panel">
                <h3><i class="fas fa-chart-bar" style="color:#2196F3;"></i> 采购成本月度趋势</h3>
                <% For m = 0 To 11
                    Dim barWidth : barWidth = 0
                    If maxMonthlyCost > 0 Then barWidth = (monthlyCosts(m) / maxMonthlyCost) * 100
                %>
                <div class="bar-container">
                    <div class="bar-label">
                        <span><%= monthlyLabels(m) %></span>
                        <span>¥<%= FormatNumber(monthlyCosts(m), 0) %></span>
                    </div>
                    <div class="bar-track">
                        <div class="bar-fill" style="width:<%= barWidth %>%;background:linear-gradient(90deg,#FF9800,#FF5722);"></div>
                    </div>
                </div>
                <% Next %>
                
                <div style="margin-top:15px; padding:12px; background:rgba(255,255,255,0.03); border-radius:8px;">
                    <div style="display:flex;justify-content:space-between;font-size:12px;color:#888;">
                        <span><i class="fas fa-arrow-up" style="color:#4CAF50;"></i> 最高: ¥<%= FormatNumber(maxMonthlyCost, 0) %></span>
                        <span><i class="fas fa-calendar"></i> 近12个月趋势</span>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- 成本分摊明细 -->
        <div class="panel" style="margin-top:0;">
            <div style="display:flex;justify-content:space-between;align-items:center;">
                <h3 style="margin:0 0 15px;"><i class="fas fa-list-alt" style="color:#4CAF50;"></i> 成本分摊明细</h3>
                <div style="display:flex;gap:10px;">
                    <form method="post" style="display:inline;">
                        <input type="hidden" name="csrf_token" value="<%= Session("CSRFToken") %>">
                        <input type="hidden" name="action" value="recalc">
                        <button type="submit" class="btn btn--info btn--sm"><i class="fas fa-sync-alt"></i> 重新计算利润</button>
                    </form>
                    <form method="post" style="display:inline;" onsubmit="return confirm('确定要将所有品牌定香产品的采购成本同步到产品库吗？')">
                        <input type="hidden" name="csrf_token" value="<%= Session("CSRFToken") %>">
                        <input type="hidden" name="action" value="sync_to_products">
                        <button type="submit" class="btn btn--warning btn--sm"><i class="fas fa-link"></i> 同步成本到产品库</button>
                    </form>
                </div>
            </div>
            <table class="data-table">
                <thead>
                    <tr>
                        <th>时间</th>
                        <th>采购单</th>
                        <th>产品</th>
                        <th>单位成本</th>
                        <th>数量</th>
                        <th>总成本</th>
                        <th>售价</th>
                        <th>利润</th>
                    </tr>
                </thead>
                <tbody>
                    <% If Not rsAllocation Is Nothing Then
                        If Not rsAllocation.EOF Then
                            Do While Not rsAllocation.EOF
                                Dim alProfit : alProfit = SafeNum(rsAllocation("ProfitAmount"))
                    %>
                    <tr>
                        <td style="font-size:12px;color:#888;"><%= SafeFormatDateTime(rsAllocation("AllocatedAt"), 2) %></td>
                        <td style="font-family:Consolas,monospace;font-size:12px;"><%= Server.HTMLEncode(CStr(rsAllocation("PurchaseNo"))) %></td>
                        <td><%= Server.HTMLEncode(CStr(rsAllocation("ProductName"))) %></td>
                        <td>¥<%= FormatNumber(SafeNum(rsAllocation("CostPerUnit")), 2) %></td>
                        <td><%= SafeNum(rsAllocation("Quantity")) %></td>
                        <td>¥<%= FormatNumber(SafeNum(rsAllocation("TotalCost")), 2) %></td>
                        <td>¥<%= FormatNumber(SafeNum(rsAllocation("SalePrice")), 2) %></td>
                        <td class="<%=IIf(alProfit>=0,"profit-positive","profit-negative")%>">¥<%= FormatNumber(alProfit, 2) %></td>
                    </tr>
                    <%
                                rsAllocation.MoveNext
                            Loop
                        Else
                    %>
                    <tr><td colspan="8" style="text-align:center;color:#666;padding:30px;">暂无成本分摊数据</td></tr>
                    <%  End If
                    End If
                    If Not rsAllocation Is Nothing Then
                        rsAllocation.Close
                        Set rsAllocation = Nothing
                    End If %>
                </tbody>
            </table>
        </div>
    </div>
</body>
</html>
<%
If Not rsProductCost Is Nothing Then
    rsProductCost.Close
    Set rsProductCost = Nothing
End If
Call CloseConnection()
%>
