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
' 风控管理体系 - Risk Control Dashboard
' 功能: 异常订单识别、库存资金预警、SKU健康度、客户信用、成本异常告警
' ============================================

Function SafeNum(val)
    If IsNull(val) Or IsEmpty(val) Or val = "" Then SafeNum = 0 Else On Error Resume Next: SafeNum = CDbl(val): If Err.Number <> 0 Then SafeNum = 0: Err.Clear: End If
End Function

Call OpenConnection()
conn.CommandTimeout = 60

' 增加脚本超时时间 (默认90秒，此页数据量大设300秒)
Server.ScriptTimeout = 300

' 页面执行耗时记录
Dim pageStartTime
pageStartTime = Timer()

' 权限检查
Dim canEdit
canEdit = (Session("AdminRoleCode") = "FIN_MANAGER" OR Session("AdminRoleCode") = "SUPER_ADMIN")

Dim currentTab
currentTab = Request.QueryString("tab")
If currentTab = "" Then currentTab = "orders"

Call LogAdminAction("查看风控管理", "finance", "risk_control", "", "")

' ============================================
' 1. 异常订单识别
' ============================================
' 高风险阈值: 订单金额>5000, 同一地址多次下单, 高频率下单, 异常数量
' 安全读取ADODB字段值（处理VT_ERROR、Decimal、Null）
Function SafeN(rs, f)
    On Error Resume Next
    Dim v : v = rs(f)
    If Err.Number <> 0 Or VarType(v) = 10 Or IsNull(v) Or IsEmpty(v) Then
        SafeN = 0
        Err.Clear
    Else
        SafeN = CDbl("0" & v)
    End If
End Function

Function SafeS(rs, f)
    On Error Resume Next
    Dim v : v = rs(f)
    If Err.Number <> 0 Or VarType(v) = 10 Or IsNull(v) Then
        SafeS = ""
        Err.Clear
    Else
        SafeS = CStr(v)
    End If
End Function

Function GetAnomalyOrders()
    Dim rs, sql, result
    Set result = Server.CreateObject("Scripting.Dictionary")
    On Error Resume Next
    
    ' 大额订单 (高于平均客单价3倍) — 使用预聚合子查询避免笛卡尔积
    sql = "SELECT TOP 50 o.OrderID, o.OrderNo, o.TotalAmount, o.Status, o.UserID, u.Username, " & _
          "o.CreatedAt, o.ShippingAddress, o.ShippingPhone, " & _
          "ISNULL(oc.UserOrderCount, 0) AS UserOrderCount " & _
          "FROM Orders o " & _
          "LEFT JOIN Users u ON o.UserID=u.UserID " & _
          "LEFT JOIN (SELECT UserID, COUNT(*) AS UserOrderCount FROM Orders WHERE Status NOT IN ('Cancelled') GROUP BY UserID) oc ON o.UserID=oc.UserID " & _
          "WHERE o.TotalAmount > 3000 AND o.Status IN ('Pending','Processing','Paid') " & _
          "ORDER BY o.TotalAmount DESC"
    Set rs = conn.Execute(sql)
    If Not rs Is Nothing Then
        Dim items, item, arrIdx, arrCap
        arrIdx = -1
        arrCap = 49
        ReDim items(arrCap)
        Do While Not rs.EOF
            arrIdx = arrIdx + 1
            If arrIdx > arrCap Then
                arrCap = arrCap + 50
                ReDim Preserve items(arrCap)
            End If
            Set item = Server.CreateObject("Scripting.Dictionary")
            Err.Clear
            item.Add "OrderID", SafeN(rs, "OrderID")
            item.Add "OrderNo", SafeS(rs, "OrderNo")
            item.Add "TotalAmount", SafeNum(rs("TotalAmount"))
            item.Add "Status", SafeS(rs, "Status")
            item.Add "UserID", SafeN(rs, "UserID")
            item.Add "Username", SafeS(rs, "Username")
            item.Add "CreatedAt", SafeS(rs, "CreatedAt")
            Dim svAddr : svAddr = SafeS(rs, "ShippingAddress")
            item.Add "Address", svAddr
            Dim svPhone : svPhone = SafeS(rs, "ShippingPhone")
            item.Add "Phone", svPhone
            item.Add "UserOrderCount", SafeN(rs, "UserOrderCount")
            item.Add "RiskType", "大额订单"
            item.Add "RiskLevel", IIf(SafeNum(rs("TotalAmount")) > 10000, "high", "medium")
            Set items(arrIdx) = item
            rs.MoveNext
        Loop
        If arrIdx >= 0 Then ReDim Preserve items(arrIdx) Else items = Array()
        rs.Close
        result.Add "high_amount", items
    End If
    Err.Clear
    Set rs = Nothing
    
    ' 同一IP/电话高频下单（通过相同地址和电话判断）
    sql = "SELECT TOP 50 o.ShippingPhone, o.ShippingAddress, COUNT(*) AS OrderCount, " & _
          "SUM(o.TotalAmount) AS TotalSpent, MAX(o.CreatedAt) AS LastOrder " & _
          "FROM Orders o WHERE o.Status NOT IN ('Cancelled') " & _
          "GROUP BY o.ShippingPhone, o.ShippingAddress HAVING COUNT(*) >= 3 " & _
          "ORDER BY COUNT(*) DESC"
    Set rs = conn.Execute(sql)
    If Not rs Is Nothing Then
        arrIdx = -1
        arrCap = 49
        ReDim items(arrCap)
        Do While Not rs.EOF
            Err.Clear
            arrIdx = arrIdx + 1
            If arrIdx > arrCap Then
                arrCap = arrCap + 50
                ReDim Preserve items(arrCap)
            End If
            Set item = Server.CreateObject("Scripting.Dictionary")
            Dim rpPhone : rpPhone = SafeS(rs, "ShippingPhone")
            item.Add "Phone", rpPhone
            Dim rpAddr : rpAddr = SafeS(rs, "ShippingAddress")
            item.Add "Address", rpAddr
            Dim ocVal : ocVal = SafeN(rs, "OrderCount")
            item.Add "OrderCount", ocVal
            item.Add "TotalSpent", SafeNum(rs("TotalSpent"))
            item.Add "LastOrder", SafeS(rs, "LastOrder")
            item.Add "RiskType", "重复下单"
            item.Add "RiskLevel", IIf(CInt(ocVal) > 5, "high", "medium")
            Set items(arrIdx) = item
            rs.MoveNext
        Loop
        If arrIdx >= 0 Then ReDim Preserve items(arrIdx) Else items = Array()
        rs.Close
        result.Add "repeat_orders", items
    End If
    Err.Clear
    Set rs = Nothing
    
    ' 异常状态订单（长时间未支付/退款频繁）
    sql = "SELECT TOP 50 o.OrderID, o.OrderNo, o.UserID, u.Username, o.TotalAmount, o.Status, o.CreatedAt, " & _
          "DATEDIFF(day, o.CreatedAt, GETDATE()) AS PendingDays " & _
          "FROM Orders o LEFT JOIN Users u ON o.UserID=u.UserID " & _
          "WHERE o.Status = 'Pending' AND DATEDIFF(day, o.CreatedAt, GETDATE()) > 3 " & _
          "ORDER BY PendingDays DESC"
    Set rs = conn.Execute(sql)
    If Not rs Is Nothing Then
        arrIdx = -1
        arrCap = 49
        ReDim items(arrCap)
        Do While Not rs.EOF
            Err.Clear
            arrIdx = arrIdx + 1
            If arrIdx > arrCap Then
                arrCap = arrCap + 50
                ReDim Preserve items(arrCap)
            End If
            Set item = Server.CreateObject("Scripting.Dictionary")
            item.Add "OrderID", SafeN(rs, "OrderID")
            item.Add "OrderNo", SafeS(rs, "OrderNo")
            item.Add "UserID", SafeN(rs, "UserID")
            item.Add "Username", SafeS(rs, "Username")
            item.Add "TotalAmount", SafeNum(rs("TotalAmount"))
            item.Add "Status", SafeS(rs, "Status")
            item.Add "CreatedAt", SafeS(rs, "CreatedAt")
            Dim pdVal : pdVal = SafeN(rs, "PendingDays")
            item.Add "PendingDays", pdVal
            item.Add "RiskType", "长期待支付"
            item.Add "RiskLevel", IIf(CInt(pdVal) > 7, "high", "low")
            Set items(arrIdx) = item
            rs.MoveNext
        Loop
        If arrIdx >= 0 Then ReDim Preserve items(arrIdx) Else items = Array()
        rs.Close
        result.Add "stale_orders", items
    End If
    Err.Clear
    Set rs = Nothing
    
    Set GetAnomalyOrders = result
End Function

' ============================================
' 2. 库存占压资金预警
' ============================================
Function GetInventoryCapitalRisk()
    Dim rs, sql, result
    Set result = Server.CreateObject("Scripting.Dictionary")
    On Error Resume Next
    
    ' 原料库存价值
    Set rs = conn.Execute("SELECT COUNT(*) AS ItemCount, ISNULL(SUM(StockQty * UnitPrice),0) AS TotalValue FROM RawMaterialInventory WHERE StockQty > 0")
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            result.Add "raw_count", SafeNum(rs("ItemCount"))
            result.Add "raw_value", SafeNum(rs("TotalValue"))
        End If
        rs.Close
    End If
    Set rs = Nothing
    
    ' 香调库存价值
    Set rs = conn.Execute("SELECT COUNT(*) AS ItemCount, ISNULL(SUM(ni.StockQuantity * ISNULL(fn.PriceAddition,0)),0) AS TotalValue FROM NoteInventory ni LEFT JOIN FragranceNotes fn ON ni.NoteID=fn.NoteID WHERE ni.StockQuantity > 0")
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            result.Add "note_count", SafeNum(rs("ItemCount"))
            result.Add "note_value", SafeNum(rs("TotalValue"))
        End If
        rs.Close
    End If
    Set rs = Nothing
    
    ' 成品库存价值
    Set rs = conn.Execute("SELECT COUNT(*) AS ItemCount, ISNULL(SUM(StockQty * UnitCost),0) AS TotalValue FROM ProductInventory WHERE StockQty > 0")
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            result.Add "product_count", SafeNum(rs("ItemCount"))
            result.Add "product_value", SafeNum(rs("TotalValue"))
        End If
        rs.Close
    End If
    Set rs = Nothing
    
    ' 慢流动原料 (库存量大但最近无采购/出库记录)
    Set rs = conn.Execute("SELECT TOP 20 r.MaterialID, r.ItemName, r.ItemCode, r.StockQty, r.SafetyStock, r.UnitPrice, " & _
                          "(r.StockQty * r.UnitPrice) AS InventoryValue, r.UpdatedAt, " & _
                          "DATEDIFF(day, ISNULL(r.UpdatedAt, r.LastPurchaseDate), GETDATE()) AS DaysInactive " & _
                          "FROM RawMaterialInventory r WHERE r.StockQty > 0 AND r.UnitPrice > 0 " & _
                          "ORDER BY (r.StockQty * r.UnitPrice) DESC")
    If Not rs Is Nothing Then
        arrIdx = -1
        arrCap = 49
        ReDim items(arrCap)
        Do While Not rs.EOF
            Err.Clear
            arrIdx = arrIdx + 1
            If arrIdx > arrCap Then
                arrCap = arrCap + 50
                ReDim Preserve items(arrCap)
            End If
            Set item = Server.CreateObject("Scripting.Dictionary")
            item.Add "MaterialID", SafeN(rs, "MaterialID")
            item.Add "ItemName", SafeS(rs, "ItemName")
            item.Add "ItemCode", SafeS(rs, "ItemCode")
            item.Add "StockQty", SafeNum(rs("StockQty"))
            item.Add "SafetyStock", SafeNum(rs("SafetyStock"))
            item.Add "UnitPrice", SafeNum(rs("UnitPrice"))
            item.Add "InventoryValue", SafeNum(rs("InventoryValue"))
            item.Add "UpdatedAt", SafeS(rs, "UpdatedAt")
            item.Add "DaysInactive", SafeNum(rs("DaysInactive"))
            Set items(arrIdx) = item
            rs.MoveNext
        Loop
        If arrIdx >= 0 Then ReDim Preserve items(arrIdx) Else items = Array()
        rs.Close
        result.Add "slow_moving", items
    End If
    Set rs = Nothing
    Err.Clear
    
    ' 总库存占压资金
    Dim totalInv
    totalInv = SafeNum(result("raw_value")) + SafeNum(result("note_value")) + SafeNum(result("product_value"))
    result.Add "total_inventory_value", totalInv
    
    ' 最近30天销售额
    Set rs = conn.Execute("SELECT ISNULL(SUM(CAST(TotalAmount AS FLOAT)),0) AS Sales30 FROM Orders WHERE Status IN ('Paid','Completed') AND CreatedAt >= DATEADD(day, -30, GETDATE())")
    If Not rs Is Nothing Then
        If Not rs.EOF Then result.Add "sales_30d", SafeNum(rs("Sales30"))
        rs.Close
    End If
    Set rs = Nothing
    Err.Clear
    
    Set GetInventoryCapitalRisk = result
End Function

' ============================================
' 3. SKU健康度评分
' ============================================
Function GetSKUHealthScore()
    Dim rs, sql, result
    Set result = Server.CreateObject("Scripting.Dictionary")
    On Error Resume Next
    
    sql = "SELECT TOP 200 p.ProductID, p.ProductName, p.ProductType, p.BasePrice, p.UnitCost, p.BOMCost, " & _
          "p.IsActive, " & _
          "ISNULL(od_agg.SalesCount, 0) AS SalesCount, " & _
          "ISNULL(od_agg.TotalSold, 0) AS TotalSold, " & _
          "ISNULL(od_agg.TotalRevenue, 0) AS TotalRevenue, " & _
          "ISNULL(pr_agg.AvgRating, 0) AS AvgRating, " & _
          "ISNULL(uf_agg.FavCount, 0) AS FavCount " & _
          "FROM Products p " & _
          "LEFT JOIN (SELECT od.ProductID, COUNT(*) AS SalesCount, SUM(od.Quantity) AS TotalSold, SUM(od.Subtotal) AS TotalRevenue " & _
          "FROM OrderDetails od JOIN Orders o ON od.OrderID=o.OrderID WHERE o.Status IN ('Paid','Completed') GROUP BY od.ProductID) od_agg ON p.ProductID=od_agg.ProductID " & _
          "LEFT JOIN (SELECT ProductID, AVG(CAST(Rating AS FLOAT)) AS AvgRating FROM ProductReviews WHERE Status='Approved' GROUP BY ProductID) pr_agg ON p.ProductID=pr_agg.ProductID " & _
          "LEFT JOIN (SELECT ProductID, COUNT(*) AS FavCount FROM UserFavorites GROUP BY ProductID) uf_agg ON p.ProductID=uf_agg.ProductID " & _
          "ORDER BY p.ProductID"
    Set rs = conn.Execute(sql)
    If Not rs Is Nothing Then
        arrIdx = -1
        arrCap = 49
        ReDim items(arrCap)
        Do While Not rs.EOF
            Err.Clear
            arrIdx = arrIdx + 1
            If arrIdx > arrCap Then
                arrCap = arrCap + 50
                ReDim Preserve items(arrCap)
            End If
            Set item = Server.CreateObject("Scripting.Dictionary")
            Dim pid, pName, pType, basePrice, unitCost, salesCount, totalSold, avgRating, totalRevenue, favCount
            pid = SafeN(rs, "ProductID")
            pName = SafeS(rs, "ProductName")
            pType = SafeS(rs, "ProductType")
            basePrice = SafeNum(rs("BasePrice"))
            unitCost = SafeNum(rs("UnitCost"))
            salesCount = SafeNum(rs("SalesCount"))
            totalSold = SafeNum(rs("TotalSold"))
            avgRating = SafeNum(rs("AvgRating"))
            totalRevenue = SafeNum(rs("TotalRevenue"))
            favCount = SafeNum(rs("FavCount"))
            
            item.Add "ProductID", pid
            item.Add "ProductName", pName
            item.Add "ProductType", pType
            item.Add "BasePrice", basePrice
            item.Add "UnitCost", unitCost
            
            ' 计算评分 (0-100)
            Dim score, salesScore, profitScore, ratingScore, favScore
            
            ' 销售得分 (30分): 基于销售数量
            If totalSold >= 100 Then
                salesScore = 30
            ElseIf totalSold >= 50 Then
                salesScore = 25
            ElseIf totalSold >= 20 Then
                salesScore = 20
            ElseIf totalSold >= 10 Then
                salesScore = 15
            ElseIf totalSold >= 5 Then
                salesScore = 10
            ElseIf totalSold >= 1 Then
                salesScore = 5
            Else
                salesScore = 0
            End If
            
            ' 利润得分 (30分): 基于利润率
            If basePrice > 0 And unitCost > 0 Then
                Dim margin
                margin = ((basePrice - unitCost) / basePrice) * 100
                If margin >= 50 Then
                    profitScore = 30
                ElseIf margin >= 40 Then
                    profitScore = 25
                ElseIf margin >= 30 Then
                    profitScore = 20
                ElseIf margin >= 20 Then
                    profitScore = 15
                ElseIf margin >= 10 Then
                    profitScore = 10
                ElseIf margin > 0 Then
                    profitScore = 5
                Else
                    profitScore = 0
                End If
            Else
                profitScore = 0
            End If
            
            ' 评分得分 (20分): 基于用户评分
            If avgRating >= 4.5 Then
                ratingScore = 20
            ElseIf avgRating >= 4.0 Then
                ratingScore = 16
            ElseIf avgRating >= 3.5 Then
                ratingScore = 12
            ElseIf avgRating >= 3.0 Then
                ratingScore = 8
            ElseIf avgRating >= 2.0 Then
                ratingScore = 4
            ElseIf avgRating > 0 Then
                ratingScore = 2
            Else
                ratingScore = 5
            End If
            
            ' 收藏得分 (20分): 基于收藏数量
            If favCount >= 50 Then
                favScore = 20
            ElseIf favCount >= 30 Then
                favScore = 16
            ElseIf favCount >= 20 Then
                favScore = 12
            ElseIf favCount >= 10 Then
                favScore = 8
            ElseIf favCount >= 5 Then
                favScore = 5
            ElseIf favCount >= 1 Then
                favScore = 2
            Else
                favScore = 0
            End If
            
            score = salesScore + profitScore + ratingScore + favScore
            
            item.Add "SalesCount", salesCount
            item.Add "TotalSold", totalSold
            item.Add "AvgRating", Round(avgRating, 1)
            item.Add "TotalRevenue", totalRevenue
            item.Add "FavCount", favCount
            item.Add "HealthScore", score
            
            ' 健康度等级
            If score >= 80 Then
                item.Add "HealthLevel", "healthy"
            ElseIf score >= 60 Then
                item.Add "HealthLevel", "warning"
            ElseIf score >= 40 Then
                item.Add "HealthLevel", "risk"
            Else
                item.Add "HealthLevel", "critical"
            End If
            
            Set items(arrIdx) = item
            rs.MoveNext
        Loop
        If arrIdx >= 0 Then ReDim Preserve items(arrIdx) Else items = Array()
        rs.Close
        result.Add "sku_scores", items
    End If
    Set rs = Nothing
    Err.Clear
    
    ' 统计各等级数量
    Dim healthy, warning, risk, critical
    healthy = 0: warning = 0: risk = 0: critical = 0
    If result.Exists("sku_scores") Then
        Dim skuItems
        skuItems = result("sku_scores")
        For Each item In skuItems
            Select Case item("HealthLevel")
                Case "healthy": healthy = healthy + 1
                Case "warning": warning = warning + 1
                Case "risk": risk = risk + 1
                Case "critical": critical = critical + 1
            End Select
        Next
    End If
    result.Add "healthy_count", healthy
    result.Add "warning_count", warning
    result.Add "risk_count", risk
    result.Add "critical_count", critical
    
    Set GetSKUHealthScore = result
End Function

' ============================================
' 4. 客户信用评级
' ============================================
Function GetCustomerCreditRating()
    Dim rs, sql, result
    Set result = Server.CreateObject("Scripting.Dictionary")
    On Error Resume Next
    
    sql = "SELECT TOP 200 u.UserID, u.Username, u.Email, u.CreatedAt, " & _
          "ISNULL(ord_agg.TotalOrders, 0) AS TotalOrders, " & _
          "ISNULL(ord_agg.TotalSpent, 0) AS TotalSpent, " & _
          "ISNULL(ord_agg.ReturnCount, 0) AS ReturnCount, " & _
          "ISNULL(ord_agg.CompletedOrders, 0) AS CompletedOrders, " & _
          "ISNULL(ord_agg.UnpaidOrders, 0) AS UnpaidOrders, " & _
          "ISNULL(ord_agg.CancelCount, 0) AS CancelCount, " & _
          "ISNULL(pr_agg.AvgRating, 0) AS AvgRating " & _
          "FROM Users u " & _
          "LEFT JOIN (SELECT UserID, SUM(CASE WHEN Status NOT IN ('Cancelled') THEN 1 ELSE 0 END) AS TotalOrders, " & _
          "SUM(CASE WHEN Status IN ('Paid','Completed') THEN TotalAmount ELSE 0 END) AS TotalSpent, " & _
          "SUM(CASE WHEN Status='Returned' THEN 1 ELSE 0 END) AS ReturnCount, " & _
          "SUM(CASE WHEN Status IN ('Paid','Completed') AND ShippingStatus='Delivered' THEN 1 ELSE 0 END) AS CompletedOrders, " & _
          "SUM(CASE WHEN DATEDIFF(day, CreatedAt, GETDATE()) > 7 AND Status='Pending' THEN 1 ELSE 0 END) AS UnpaidOrders, " & _
          "SUM(CASE WHEN Status='Cancelled' THEN 1 ELSE 0 END) AS CancelCount " & _
          "FROM Orders GROUP BY UserID) ord_agg ON u.UserID=ord_agg.UserID " & _
          "LEFT JOIN (SELECT UserID, AVG(CAST(Rating AS FLOAT)) AS AvgRating FROM ProductReviews WHERE Status='Approved' GROUP BY UserID) pr_agg ON u.UserID=pr_agg.UserID " & _
          "WHERE u.IsActive=1 ORDER BY ord_agg.TotalSpent DESC"
    Set rs = conn.Execute(sql)
    If Not rs Is Nothing Then
        arrIdx = -1
        arrCap = 49
        ReDim items(arrCap)
        Do While Not rs.EOF
            Err.Clear
            arrIdx = arrIdx + 1
            If arrIdx > arrCap Then
                arrCap = arrCap + 50
                ReDim Preserve items(arrCap)
            End If
            Set item = Server.CreateObject("Scripting.Dictionary")
            Dim uid, uname, totalOrders, totalSpent, returnCount, completedOrders, avgRate, unpaidOrders, cancelCount
            uid = SafeN(rs, "UserID")
            uname = SafeS(rs, "Username")
            totalOrders = SafeNum(rs("TotalOrders"))
            totalSpent = SafeNum(rs("TotalSpent"))
            returnCount = SafeNum(rs("ReturnCount"))
            completedOrders = SafeNum(rs("CompletedOrders"))
            avgRate = SafeNum(rs("AvgRating"))
            unpaidOrders = SafeNum(rs("UnpaidOrders"))
            cancelCount = SafeNum(rs("CancelCount"))
            
            item.Add "UserID", uid
            item.Add "Username", uname
            item.Add "TotalOrders", totalOrders
            item.Add "TotalSpent", totalSpent
            item.Add "ReturnCount", returnCount
            item.Add "CompletedOrders", completedOrders
            item.Add "AvgRating", Round(avgRate, 1)
            item.Add "Email", SafeS(rs, "Email")
            item.Add "UnpaidOrders", unpaidOrders
            item.Add "CancelCount", cancelCount
            
            ' 信用评级: A/B/C/D
            Dim credit, returnRate, cancelRate
            If totalOrders > 0 Then
                returnRate = returnCount / totalOrders
                cancelRate = cancelCount / totalOrders
            Else
                returnRate = 0
                cancelRate = 0
            End If
            
            If totalSpent >= 5000 And returnRate < 0.1 And cancelRate < 0.15 And unpaidOrders = 0 Then
                credit = "A级(优质)"
            ElseIf totalSpent >= 1000 And returnRate < 0.2 And cancelRate < 0.25 And unpaidOrders <= 1 Then
                credit = "B级(良好)"
            ElseIf totalOrders > 0 And returnRate < 0.3 Then
                credit = "C级(一般)"
            Else
                credit = "D级(关注)"
            End If
            item.Add "CreditLevel", credit
            
            Set items(arrIdx) = item
            rs.MoveNext
        Loop
        If arrIdx >= 0 Then ReDim Preserve items(arrIdx) Else items = Array()
        rs.Close
        result.Add "customers", items
    End If
    Set rs = Nothing
    Err.Clear
    
    ' 信用评级统计
    Dim aCount, bCount, cCount, dCount
    aCount = 0: bCount = 0: cCount = 0: dCount = 0
    If result.Exists("customers") Then
        Dim custItems
        custItems = result("customers")
        For Each item In custItems
            Select Case Left(item("CreditLevel"), 1)
                Case "A": aCount = aCount + 1
                Case "B": bCount = bCount + 1
                Case "C": cCount = cCount + 1
                Case "D": dCount = dCount + 1
            End Select
        Next
    End If
    result.Add "a_count", aCount
    result.Add "b_count", bCount
    result.Add "c_count", cCount
    result.Add "d_count", dCount
    
    Set GetCustomerCreditRating = result
End Function

' ============================================
' 5. 成本异常波动告警
' ============================================
Function GetCostAnomalyAlert()
    Dim rs, sql, result
    Set result = Server.CreateObject("Scripting.Dictionary")
    On Error Resume Next
    
    ' 获取所有产品当月和上月成本对比（使用日期范围避免全表扫描）
    sql = "SELECT TOP 200 p.ProductID, p.ProductName, p.UnitCost, p.BOMCost, " & _
          "ISNULL(pc_curr.TotalCost, 0) AS CurrentCost, " & _
          "ISNULL(pc_last.TotalCost, 0) AS LastCost " & _
          "FROM Products p " & _
          "LEFT JOIN (SELECT ProductID, SUM(TotalCost) AS TotalCost FROM ProductCosts " & _
          "WHERE CreatedAt >= DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1) " & _
          "AND CreatedAt < DATEADD(month, 1, DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1)) " & _
          "GROUP BY ProductID) pc_curr ON p.ProductID=pc_curr.ProductID " & _
          "LEFT JOIN (SELECT ProductID, SUM(TotalCost) AS TotalCost FROM ProductCosts " & _
          "WHERE CreatedAt >= DATEADD(month, -1, DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1)) " & _
          "AND CreatedAt < DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1) " & _
          "GROUP BY ProductID) pc_last ON p.ProductID=pc_last.ProductID " & _
          "WHERE p.IsActive=1 ORDER BY p.ProductID"
    Set rs = conn.Execute(sql)
    If Not rs Is Nothing Then
        arrIdx = -1
        arrCap = 49
        ReDim items(arrCap)
        Do While Not rs.EOF
            Err.Clear
            Dim currCost, lastCost, variance, varianceRate
            currCost = SafeNum(rs("CurrentCost"))
            lastCost = SafeNum(rs("LastCost"))
            
            If currCost > 0 Or lastCost > 0 Then
                If lastCost > 0 Then
                    variance = currCost - lastCost
                    varianceRate = (variance / lastCost) * 100
                Else
                    variance = currCost
                    varianceRate = 100
                End If
                
                ' 仅记录波动超过5%的
                If Abs(varianceRate) > 5 Then
                    arrIdx = arrIdx + 1
                    If arrIdx > arrCap Then
                        arrCap = arrCap + 50
                        ReDim Preserve items(arrCap)
                    End If
                    Set item = Server.CreateObject("Scripting.Dictionary")
                    item.Add "ProductID", SafeN(rs, "ProductID")
                    item.Add "ProductName", SafeS(rs, "ProductName")
                    item.Add "CurrentCost", currCost
                    item.Add "LastCost", lastCost
                    item.Add "Variance", Round(variance, 2)
                    item.Add "VarianceRate", Round(varianceRate, 1)
                    item.Add "AlertLevel", IIf(Abs(varianceRate) > 20, "high", IIf(Abs(varianceRate) > 10, "medium", "low"))
                    Set items(arrIdx) = item
                End If
            End If
            rs.MoveNext
        Loop
        If arrIdx >= 0 Then ReDim Preserve items(arrIdx) Else items = Array()
        rs.Close
        result.Add "cost_alerts", items
    End If
    Set rs = Nothing
    Err.Clear
    
    ' 原材料价格波动
    Set rs = conn.Execute("SELECT TOP 20 sp.ItemCode, sp.ItemName, sp.UnitPrice, sp.SupplierID, s.SupplierName, " & _
                          "sp.CreatedAt FROM SupplierPrices sp LEFT JOIN Suppliers s ON sp.SupplierID=s.SupplierID " & _
                          "WHERE sp.IsActive=1 ORDER BY sp.CreatedAt DESC")
    If Not rs Is Nothing Then
        arrIdx = -1
        arrCap = 49
        ReDim items(arrCap)
        Do While Not rs.EOF
            Err.Clear
            arrIdx = arrIdx + 1
            If arrIdx > arrCap Then
                arrCap = arrCap + 50
                ReDim Preserve items(arrCap)
            End If
            Set item = Server.CreateObject("Scripting.Dictionary")
            item.Add "ItemCode", SafeS(rs, "ItemCode")
            item.Add "ItemName", SafeS(rs, "ItemName")
            item.Add "UnitPrice", SafeNum(rs("UnitPrice"))
            item.Add "SupplierID", SafeN(rs, "SupplierID")
            item.Add "SupplierName", SafeS(rs, "SupplierName")
            item.Add "CreatedAt", SafeS(rs, "CreatedAt")
            Set items(arrIdx) = item
            rs.MoveNext
        Loop
        If arrIdx >= 0 Then ReDim Preserve items(arrIdx) Else items = Array()
        rs.Close
        result.Add "supplier_prices", items
    End If
    Set rs = Nothing
    Err.Clear
    
    Set GetCostAnomalyAlert = result
End Function

' ============================================
' 执行数据采集
' ============================================
Dim anomalyOrders, invCapital, skuHealth, creditRating, costAlert

' 初始化为空Dictionary（按需加载时未加载的Tab保持空字典）
Set anomalyOrders = Server.CreateObject("Scripting.Dictionary")
Set invCapital = Server.CreateObject("Scripting.Dictionary")
Set skuHealth = Server.CreateObject("Scripting.Dictionary")
Set creditRating = Server.CreateObject("Scripting.Dictionary")
Set costAlert = Server.CreateObject("Scripting.Dictionary")

' 安全初始化——如果函数执行失败则返回空Dictionary避免后续报错
Dim dataLoadErrors
Set dataLoadErrors = Server.CreateObject("Scripting.Dictionary")

Function SafeGetData(getFunction, funcName)
    On Error Resume Next
    Dim result
    Set result = getFunction
    If Err.Number <> 0 Then
        dataLoadErrors.Add funcName, "执行错误 " & Err.Number & ": " & Err.Description
        Set result = Server.CreateObject("Scripting.Dictionary")
        Err.Clear
    ElseIf result.Count = 0 Then
        dataLoadErrors.Add funcName, "查询无数据返回"
    End If
    On Error GoTo 0
    Set SafeGetData = result
End Function

' 按Tab按需加载——只执行当前Tab需要的数据查询
If currentTab = "overview" Then
    ' 概览页：仅执行轻量级COUNT/SUM聚合查询，不加载完整数据集
    Dim rsOvw, skuCntSql, credCntSql, costCntSql
    On Error Resume Next
    
    ' === 大额异常订单数 ===
    Err.Clear
    Set rsOvw = conn.Execute("SELECT COUNT(*) AS Cnt FROM Orders WHERE TotalAmount > 3000 AND Status IN ('Pending','Processing','Paid')")
    If Err.Number = 0 And Not rsOvw Is Nothing Then
        If Not rsOvw.EOF Then anomalyOrders.Add "high_amount_count", SafeNum(rsOvw("Cnt"))
        rsOvw.Close
    End If
    If Err.Number <> 0 Then Err.Clear
    Set rsOvw = Nothing
    
    ' === 库存价值——原料 ===
    Err.Clear
    Set rsOvw = conn.Execute("SELECT ISNULL(SUM(StockQty * UnitPrice),0) AS Val FROM RawMaterialInventory WHERE StockQty > 0")
    If Err.Number = 0 And Not rsOvw Is Nothing Then
        If Not rsOvw.EOF Then invCapital.Add "raw_value", SafeNum(rsOvw("Val"))
        rsOvw.Close
    End If
    If Err.Number <> 0 Then Err.Clear
    Set rsOvw = Nothing
    
    ' === 库存价值——香调 ===
    Err.Clear
    Set rsOvw = conn.Execute("SELECT ISNULL(SUM(ni.StockQuantity * ISNULL(fn.PriceAddition,0)),0) AS Val FROM NoteInventory ni LEFT JOIN FragranceNotes fn ON ni.NoteID=fn.NoteID WHERE ni.StockQuantity > 0")
    If Err.Number = 0 And Not rsOvw Is Nothing Then
        If Not rsOvw.EOF Then invCapital.Add "note_value", SafeNum(rsOvw("Val"))
        rsOvw.Close
    End If
    If Err.Number <> 0 Then Err.Clear
    Set rsOvw = Nothing
    
    ' === 库存价值——成品 ===
    Err.Clear
    Set rsOvw = conn.Execute("SELECT ISNULL(SUM(StockQty * UnitCost),0) AS Val FROM ProductInventory WHERE StockQty > 0")
    If Err.Number = 0 And Not rsOvw Is Nothing Then
        If Not rsOvw.EOF Then invCapital.Add "product_value", SafeNum(rsOvw("Val"))
        rsOvw.Close
    End If
    If Err.Number <> 0 Then Err.Clear
    Set rsOvw = Nothing
    
    invCapital.Add "total_inventory_value", SafeNum(invCapital("raw_value")) + SafeNum(invCapital("note_value")) + SafeNum(invCapital("product_value"))
    
    ' === 近30天销售额 ===
    Err.Clear
    Set rsOvw = conn.Execute("SELECT ISNULL(SUM(CAST(TotalAmount AS FLOAT)),0) AS Val FROM Orders WHERE Status IN ('Paid','Completed') AND CreatedAt >= DATEADD(day,-30,GETDATE())")
    If Err.Number = 0 And Not rsOvw Is Nothing Then
        If Not rsOvw.EOF Then invCapital.Add "sales_30d", SafeNum(rsOvw("Val"))
        rsOvw.Close
    End If
    If Err.Number <> 0 Then Err.Clear
    Set rsOvw = Nothing
    
    ' === SKU健康度统计——SQL层计算评分并聚合为计数 ===
    skuCntSql = "SELECT " & _
        "SUM(CASE WHEN sc>=80 THEN 1 ELSE 0 END) AS hc," & _
        "SUM(CASE WHEN sc>=60 AND sc<80 THEN 1 ELSE 0 END) AS wc," & _
        "SUM(CASE WHEN sc>=40 AND sc<60 THEN 1 ELSE 0 END) AS rc," & _
        "SUM(CASE WHEN sc<40 THEN 1 ELSE 0 END) AS cc " & _
        "FROM (SELECT " & _
        "CASE WHEN ISNULL(oa.TS,0)>=100 THEN 30 WHEN ISNULL(oa.TS,0)>=50 THEN 25 " & _
        "WHEN ISNULL(oa.TS,0)>=20 THEN 20 WHEN ISNULL(oa.TS,0)>=10 THEN 15 " & _
        "WHEN ISNULL(oa.TS,0)>=5 THEN 10 WHEN ISNULL(oa.TS,0)>=1 THEN 5 ELSE 0 END+" & _
        "CASE WHEN p.BasePrice>0 AND p.UnitCost>0 THEN " & _
        "CASE WHEN (p.BasePrice-p.UnitCost)*100/p.BasePrice>=50 THEN 30 " & _
        "WHEN (p.BasePrice-p.UnitCost)*100/p.BasePrice>=40 THEN 25 " & _
        "WHEN (p.BasePrice-p.UnitCost)*100/p.BasePrice>=30 THEN 20 " & _
        "WHEN (p.BasePrice-p.UnitCost)*100/p.BasePrice>=20 THEN 15 " & _
        "WHEN (p.BasePrice-p.UnitCost)*100/p.BasePrice>=10 THEN 10 " & _
        "WHEN (p.BasePrice-p.UnitCost)*100/p.BasePrice>0 THEN 5 ELSE 0 END " & _
        "ELSE 0 END+" & _
        "CASE WHEN ISNULL(pa.AR,0)>=4.5 THEN 20 WHEN ISNULL(pa.AR,0)>=4.0 THEN 16 " & _
        "WHEN ISNULL(pa.AR,0)>=3.5 THEN 12 WHEN ISNULL(pa.AR,0)>=3.0 THEN 8 " & _
        "WHEN ISNULL(pa.AR,0)>=2.0 THEN 4 WHEN ISNULL(pa.AR,0)>0 THEN 2 ELSE 5 END+" & _
        "CASE WHEN ISNULL(fa.FC,0)>=50 THEN 20 WHEN ISNULL(fa.FC,0)>=30 THEN 16 " & _
        "WHEN ISNULL(fa.FC,0)>=20 THEN 12 WHEN ISNULL(fa.FC,0)>=10 THEN 8 " & _
        "WHEN ISNULL(fa.FC,0)>=5 THEN 5 WHEN ISNULL(fa.FC,0)>=1 THEN 2 ELSE 0 END " & _
        "AS sc FROM Products p " & _
        "LEFT JOIN (SELECT od.ProductID,SUM(od.Quantity) AS TS FROM OrderDetails od " & _
        "JOIN Orders o ON od.OrderID=o.OrderID WHERE o.Status IN ('Paid','Completed') " & _
        "GROUP BY od.ProductID) oa ON p.ProductID=oa.ProductID " & _
        "LEFT JOIN (SELECT ProductID,AVG(CAST(Rating AS FLOAT)) AS AR FROM ProductReviews " & _
        "WHERE Status='Approved' GROUP BY ProductID) pa ON p.ProductID=pa.ProductID " & _
        "LEFT JOIN (SELECT ProductID,COUNT(*) AS FC FROM UserFavorites " & _
        "GROUP BY ProductID) fa ON p.ProductID=fa.ProductID) t"
    Err.Clear
    Set rsOvw = conn.Execute(skuCntSql)
    If Err.Number = 0 And Not rsOvw Is Nothing Then
        If Not rsOvw.EOF Then
            skuHealth.Add "healthy_count", SafeNum(rsOvw("hc"))
            skuHealth.Add "warning_count", SafeNum(rsOvw("wc"))
            skuHealth.Add "risk_count", SafeNum(rsOvw("rc"))
            skuHealth.Add "critical_count", SafeNum(rsOvw("cc"))
        End If
        rsOvw.Close
    End If
    If Err.Number <> 0 Then Err.Clear
    Set rsOvw = Nothing
    
    ' === 客户信用评级统计——SQL层聚合 ===
    credCntSql = "SELECT " & _
        "SUM(CASE WHEN TS>=5000 AND RR<0.1 AND CR<0.15 AND UP=0 THEN 1 ELSE 0 END) AS ac," & _
        "SUM(CASE WHEN NOT(TS>=5000 AND RR<0.1 AND CR<0.15 AND UP=0) AND TS>=1000 AND RR<0.2 AND CR<0.25 AND UP<=1 THEN 1 ELSE 0 END) AS bc," & _
        "SUM(CASE WHEN NOT(TS>=5000 AND RR<0.1 AND CR<0.15 AND UP=0) AND NOT(TS>=1000 AND RR<0.2 AND CR<0.25 AND UP<=1) AND TN>0 AND RR<0.3 THEN 1 ELSE 0 END) AS cc2," & _
        "SUM(CASE WHEN NOT(TS>=5000 AND RR<0.1 AND CR<0.15 AND UP=0) AND NOT(TS>=1000 AND RR<0.2 AND CR<0.25 AND UP<=1) AND NOT(TN>0 AND RR<0.3) THEN 1 ELSE 0 END) AS dc " & _
        "FROM (SELECT u.UserID,ISNULL(oa.TN,0) AS TN,ISNULL(oa.TS,0) AS TS," & _
        "CASE WHEN ISNULL(oa.TN,0)>0 THEN CAST(ISNULL(oa.RC,0) AS FLOAT)/oa.TN ELSE 0 END AS RR," & _
        "CASE WHEN ISNULL(oa.TN,0)>0 THEN CAST(ISNULL(oa.CN,0) AS FLOAT)/oa.TN ELSE 0 END AS CR," & _
        "ISNULL(oa.UP,0) AS UP FROM Users u " & _
        "LEFT JOIN (SELECT UserID," & _
        "SUM(CASE WHEN Status NOT IN ('Cancelled') THEN 1 ELSE 0 END) AS TN," & _
        "SUM(CASE WHEN Status IN ('Paid','Completed') THEN TotalAmount ELSE 0 END) AS TS," & _
        "SUM(CASE WHEN Status='Returned' THEN 1 ELSE 0 END) AS RC," & _
        "SUM(CASE WHEN DATEDIFF(day,CreatedAt,GETDATE())>7 AND Status='Pending' THEN 1 ELSE 0 END) AS UP," & _
        "SUM(CASE WHEN Status='Cancelled' THEN 1 ELSE 0 END) AS CN " & _
        "FROM Orders GROUP BY UserID) oa ON u.UserID=oa.UserID WHERE u.IsActive=1) r"
    Err.Clear
    Set rsOvw = conn.Execute(credCntSql)
    If Err.Number = 0 And Not rsOvw Is Nothing Then
        If Not rsOvw.EOF Then
            creditRating.Add "a_count", SafeNum(rsOvw("ac"))
            creditRating.Add "b_count", SafeNum(rsOvw("bc"))
            creditRating.Add "c_count", SafeNum(rsOvw("cc2"))
            creditRating.Add "d_count", SafeNum(rsOvw("dc"))
        End If
        rsOvw.Close
    End If
    If Err.Number <> 0 Then Err.Clear
    Set rsOvw = Nothing
    
    ' === 成本异动告警数 ===
    costCntSql = "SELECT COUNT(*) AS Cnt FROM (" & _
        "SELECT CASE WHEN ISNULL(lc.TC,0)>0 THEN ABS(ISNULL(cc2.TC,0)-lc.TC)*100.0/lc.TC " & _
        "WHEN ISNULL(cc2.TC,0)>0 THEN 100 ELSE 0 END AS VR " & _
        "FROM Products p " & _
        "LEFT JOIN (SELECT ProductID,SUM(TotalCost) AS TC FROM ProductCosts " & _
        "WHERE CreatedAt>=DATEFROMPARTS(YEAR(GETDATE()),MONTH(GETDATE()),1) " & _
        "AND CreatedAt<DATEADD(month,1,DATEFROMPARTS(YEAR(GETDATE()),MONTH(GETDATE()),1)) " & _
        "GROUP BY ProductID) cc2 ON p.ProductID=cc2.ProductID " & _
        "LEFT JOIN (SELECT ProductID,SUM(TotalCost) AS TC FROM ProductCosts " & _
        "WHERE CreatedAt>=DATEADD(month,-1,DATEFROMPARTS(YEAR(GETDATE()),MONTH(GETDATE()),1)) " & _
        "AND CreatedAt<DATEFROMPARTS(YEAR(GETDATE()),MONTH(GETDATE()),1) " & _
        "GROUP BY ProductID) lc ON p.ProductID=lc.ProductID " & _
        "WHERE p.IsActive=1 AND (ISNULL(cc2.TC,0)>0 OR ISNULL(lc.TC,0)>0)" & _
        ") sub WHERE VR > 5"
    Err.Clear
    Set rsOvw = conn.Execute(costCntSql)
    If Err.Number = 0 And Not rsOvw Is Nothing Then
        If Not rsOvw.EOF Then costAlert.Add "alert_count", SafeNum(rsOvw("Cnt"))
        rsOvw.Close
    End If
    If Err.Number <> 0 Then Err.Clear
    Set rsOvw = Nothing
    
    On Error GoTo 0
ElseIf currentTab = "orders" Then
    Set anomalyOrders = SafeGetData(GetAnomalyOrders, "异常订单识别")
ElseIf currentTab = "inventory" Then
    Set invCapital = SafeGetData(GetInventoryCapitalRisk, "库存资金预警")
ElseIf currentTab = "sku" Then
    Set skuHealth = SafeGetData(GetSKUHealthScore, "SKU健康度")
ElseIf currentTab = "credit" Then
    Set creditRating = SafeGetData(GetCustomerCreditRating, "客户信用评级")
ElseIf currentTab = "cost" Then
    Set costAlert = SafeGetData(GetCostAnomalyAlert, "成本异常告警")
End If

' 统计告警总数（仅概览页计算完整统计）
Dim totalAlerts, totalSKU, healthyPct, warningPct, riskPct, costAlertCount
totalAlerts = 0
totalSKU = 0
healthyPct = 0: warningPct = 0: riskPct = 0
costAlertCount = 0

If currentTab = "overview" Then
    ' 大额告警（轻量级查询已存入count）
    totalAlerts = SafeNum(anomalyOrders("high_amount_count"))
    ' SKU健康度风险
    totalAlerts = totalAlerts + SafeNum(skuHealth("risk_count")) + SafeNum(skuHealth("critical_count"))
    ' 成本告警（轻量级查询已存入alert_count）
    costAlertCount = SafeNum(costAlert("alert_count"))
    totalAlerts = totalAlerts + costAlertCount

    ' 预计算SKU健康度百分比
    totalSKU = SafeNum(skuHealth("healthy_count")) + SafeNum(skuHealth("warning_count")) + SafeNum(skuHealth("risk_count")) + SafeNum(skuHealth("critical_count"))
    If totalSKU > 0 Then
        healthyPct = Round(SafeNum(skuHealth("healthy_count")) / totalSKU * 100, 0)
        warningPct = Round(SafeNum(skuHealth("warning_count")) / totalSKU * 100, 0)
        riskPct = Round(SafeNum(skuHealth("risk_count")) / totalSKU * 100, 0)
    End If
End If
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>风控管理 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background:#1a1a2e; color:#e0e0e0; }
        .main-content { padding:30px; margin-left:260px; }
        .page-title { font-size:24px; display:flex; align-items:center; gap:12px; margin-bottom:25px; }
        .page-title i { color:#00bcd4; }
        .breadcrumb { color:#888; font-size:14px; margin-bottom:20px; }
        .breadcrumb a { color:#00bcd4; text-decoration:none; }
        
        /* 告警横幅 */
        .alert-banner { 
            background:linear-gradient(135deg,#5e1b1b,#3a0e0e); border:1px solid rgba(244,67,54,0.3);
            border-radius:12px; padding:20px; margin-bottom:25px; display:flex; align-items:center; gap:15px;
        }
        .alert-banner i { font-size:32px; color:#e57373; }
        .alert-banner .alert-text { flex:1; }
        .alert-banner .alert-text h3 { color:#e57373; margin:0 0 5px; font-size:18px; }
        .alert-banner .alert-text p { color:#aaa; margin:0; font-size:13px; }
        .alert-banner .alert-count { 
            background:#f44336; color:white; border-radius:50%; width:50px; height:50px;
            display:flex; align-items:center; justify-content:center; font-size:24px; font-weight:700;
        }
        
        /* Tab导航 */
        .tab-nav { display:flex; border-bottom:2px solid #3a3a4a; margin-bottom:25px; }
        .tab-nav a { 
            padding:15px 25px; color:#888; text-decoration:none; font-size:14px;
            border-bottom:3px solid transparent; transition:all 0.3s;
            display:flex; align-items:center; gap:8px;
        }
        .tab-nav a:hover { color:#e0e0e0; }
        .tab-nav a.active { color:#00bcd4; border-bottom-color:#00bcd4; }
        .tab-nav a .tab-badge { 
            background:#f44336; color:white; border-radius:10px; padding:2px 8px; font-size:11px; margin-left:5px;
        }
        
        /* 内容卡片 */
        .content-card { background:linear-gradient(135deg,#2d2d44,#1e1e32); border-radius:12px; padding:25px; border:1px solid rgba(255,255,255,0.06); }
        .section-title { color:#00bcd4; font-size:16px; margin:0 0 20px; display:flex; align-items:center; gap:10px; }
        
        /* 概览卡片网格 */
        .overview-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(200px,1fr)); gap:15px; margin-bottom:25px; }
        .overview-card { background:linear-gradient(135deg,#2d2d44,#1e1e32); border-radius:12px; padding:20px; border:1px solid rgba(255,255,255,0.06); text-align:center; }
        .overview-card .card-icon { font-size:28px; margin-bottom:10px; }
        .overview-card .card-value { font-size:28px; font-weight:700; margin:8px 0; }
        .overview-card .card-label { font-size:12px; color:#888; }
        
        /* 表格 */
        .data-table { width:100%; border-collapse:collapse; font-size:13px; }
        .data-table th { background:#1a1a2e; color:#888; font-weight:600; padding:10px; text-align:left; border-bottom:1px solid #3a3a4a; }
        .data-table td { padding:8px 10px; border-bottom:1px solid rgba(255,255,255,0.04); }
        .data-table tr:hover td { background:rgba(255,255,255,0.03); }
        
        /* 风险标签 */
        .risk-high { background:#5e1b1b; color:#e57373; }
        .risk-medium { background:#5e4b1b; color:#ffb74d; }
        .risk-low { background:#1b3a5e; color:#64b5f6; }
        .risk-healthy { background:#1b5e20; color:#81c784; }
        .risk-warning { background:#5e4b1b; color:#ffb74d; }
        .risk-critical { background:#5e1b1b; color:#e57373; }
        .badge { display:inline-block; padding:3px 10px; border-radius:10px; font-size:11px; font-weight:600; }
        .badge-a { background:#1b5e20; color:#81c784; }
        .badge-b { background:#1b3a5e; color:#64b5f6; }
        .badge-c { background:#5e4b1b; color:#ffb74d; }
        .badge-d { background:#5e1b1b; color:#e57373; }
        
        /* 进度条 */
        .health-bar { height:6px; border-radius:3px; margin-top:5px; margin-bottom:10px; }
        .health-bar-fill { height:100%; border-radius:3px; transition:width 0.5s; }
        .hb-green { background:#1b5e20; } .hbf-green { background:#81c784; }
        .hb-orange { background:#5e4b1b; } .hbf-orange { background:#ffb74d; }
        .hb-red { background:#5e1b1b; } .hbf-red { background:#e57373; }
        .c-green { color:#81c784; } .c-orange { color:#ffb74d; } .c-red { color:#e57373; }
        
        /* 概览小卡片 */
        .mini-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(140px,1fr)); gap:12px; margin-bottom:20px; }
        .mini-card { background:#1a1a2e; border-radius:8px; padding:15px; text-align:center; border:1px solid rgba(255,255,255,0.04); }
        .mini-card .val { font-size:20px; font-weight:700; }
        .mini-card .lbl { font-size:11px; color:#888; margin-top:5px; }
        
        /* 查询区域 */
        .filter-bar { background:rgba(255,255,255,0.03); border-radius:8px; padding:15px; margin-bottom:20px; }
        .filter-bar form { display:flex; gap:15px; align-items:center; flex-wrap:wrap; }
        .filter-bar label { color:#888; font-size:13px; }
        .filter-bar select, .filter-bar input { padding:8px 12px; border:1px solid #3a3a4a; border-radius:6px; background:#1a1a2e; color:#e0e0e0; font-size:13px; }
.text-muted { color:#666; }
        .t-mar { font-weight:600; }
        .t-hl { font-weight:700; }
        .t-var { font-weight:600; }
        .t-vrate { font-weight:600; }
        
        .info-row { display:flex; justify-content:space-between; padding:8px 0; border-bottom:1px solid rgba(255,255,255,0.04); }
        .info-row .label { color:#888; font-size:13px; }
        .info-row .value { color:#e0e0e0; font-weight:600; }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    <div class="main-content">
        <div class="breadcrumb">
            <a href="index.asp">财务中心</a> / <span>风控管理</span>
        </div>
        <h2 class="page-title"><i class="fas fa-shield-alt"></i> 风控管理体系</h2>
        
        <% If currentTab = "overview" Then %>
        <!-- 告警横幅 -->
        <div class="alert-banner">
            <i class="fas fa-exclamation-triangle"></i>
            <div class="alert-text">
                <h3>风控概览</h3>
                <p>系统实时监控订单异常、库存资金、SKU健康度、客户信用及成本波动，共发现 <strong style="color:#e57373;"><%= totalAlerts %></strong> 项风险事项需关注</p>
            </div>
            <div class="alert-count"><%= totalAlerts %></div>
        </div>
        <% End If %>
        
        <!-- Tab导航 -->
        <div class="tab-nav">
            <a href="?tab=overview" class="<%= IIf(currentTab="overview","active","") %>"><i class="fas fa-tachometer-alt"></i> 风控概览</a>
            <a href="?tab=orders" class="<%= IIf(currentTab="orders","active","") %>"><i class="fas fa-shopping-cart"></i> 异常订单</a>
            <a href="?tab=inventory" class="<%= IIf(currentTab="inventory","active","") %>"><i class="fas fa-warehouse"></i> 库存资金</a>
            <a href="?tab=sku" class="<%= IIf(currentTab="sku","active","") %>"><i class="fas fa-boxes"></i> SKU健康度</a>
            <a href="?tab=credit" class="<%= IIf(currentTab="credit","active","") %>"><i class="fas fa-user-check"></i> 客户信用</a>
            <a href="?tab=cost" class="<%= IIf(currentTab="cost","active","") %>"><i class="fas fa-chart-line"></i> 成本告警</a>
        </div>
        
        <div class="content-card">
        
        <% If currentTab = "overview" Then %>
            <!-- ====== Tab 1: 风控概览 ====== -->
            <h3 class="section-title"><i class="fas fa-tachometer-alt"></i> 风控综合概览</h3>
            
            <div class="overview-grid">
                <div class="overview-card" style="border-top:3px solid #f44336;">
                    <div class="card-icon" style="color:#f44336;"><i class="fas fa-exclamation-circle"></i></div>
                    <div class="card-value" style="color:#e57373;"><%= SafeNum(skuHealth("critical_count")) + SafeNum(skuHealth("risk_count")) %></div>
                    <div class="card-label">SKU健康度告警</div>
                </div>
                <div class="overview-card" style="border-top:3px solid #FF9800;">
                    <div class="card-icon" style="color:#FF9800;"><i class="fas fa-dollar-sign"></i></div>
                    <div class="card-value" style="color:#ffb74d;">¥<%= FormatNumber(SafeNum(invCapital("total_inventory_value")),0) %></div>
                    <div class="card-label">库存占压资金</div>
                </div>
                <div class="overview-card" style="border-top:3px solid #2196F3;">
                    <div class="card-icon" style="color:#2196F3;"><i class="fas fa-users"></i></div>
                    <div class="card-value" style="color:#64b5f6;"><%= SafeNum(creditRating("d_count")) %></div>
                    <div class="card-label">D级关注客户</div>
                </div>
                <div class="overview-card" style="border-top:3px solid #9C27B0;">
                    <div class="card-icon" style="color:#9C27B0;"><i class="fas fa-chart-line"></i></div>
                    <div class="card-value" style="color:#ce93d8;">
                        <%= costAlertCount %>
                    </div>
                    <div class="card-label">成本异动告警</div>
                </div>
            </div>
            
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:20px;">
                <!-- SKU健康度分布 -->
                <div style="background:#1a1a2e;border-radius:10px;padding:20px;border:1px solid rgba(255,255,255,0.04);">
                    <h4 style="color:#e0e0e0;margin:0 0 15px;font-size:15px;"><i class="fas fa-chart-pie"></i> SKU健康度分布</h4>
                    <div class="info-row"><span class="label">健康 (≥80分)</span><span class="value c-green"><%= SafeNum(skuHealth("healthy_count")) %> 个</span></div>
                    <div class="health-bar hb-green"><div class="health-bar-fill hbf-green" data-pct="<%= healthyPct %>"></div></div>
                    <div class="info-row"><span class="label">预警 (60-79分)</span><span class="value c-orange"><%= SafeNum(skuHealth("warning_count")) %> 个</span></div>
                    <div class="health-bar hb-orange"><div class="health-bar-fill hbf-orange" data-pct="<%= warningPct %>"></div></div>
                    <div class="info-row"><span class="label">风险 (40-59分)</span><span class="value c-red"><%= SafeNum(skuHealth("risk_count")) %> 个</span></div>
                    <div class="health-bar hb-red"><div class="health-bar-fill hbf-red" data-pct="<%= riskPct %>"></div></div>
                    <div class="info-row"><span class="label">危机 (&lt;40分)</span><span class="value c-red"><%= SafeNum(skuHealth("critical_count")) %> 个</span></div>
                </div>
                
                <!-- 客户信用分布 -->
                <div style="background:#1a1a2e;border-radius:10px;padding:20px;border:1px solid rgba(255,255,255,0.04);">
                    <h4 style="color:#e0e0e0;margin:0 0 15px;font-size:15px;"><i class="fas fa-users"></i> 客户信用评级分布</h4>
                    <div class="info-row"><span class="label">A级(优质)</span><span class="value" style="color:#81c784;"><%= SafeNum(creditRating("a_count")) %> 人</span></div>
                    <div class="info-row"><span class="label">B级(良好)</span><span class="value" style="color:#64b5f6;"><%= SafeNum(creditRating("b_count")) %> 人</span></div>
                    <div class="info-row"><span class="label">C级(一般)</span><span class="value" style="color:#ffb74d;"><%= SafeNum(creditRating("c_count")) %> 人</span></div>
                    <div class="info-row"><span class="label">D级(关注)</span><span class="value" style="color:#e57373;"><%= SafeNum(creditRating("d_count")) %> 人</span></div>
                </div>
            </div>
            
            <!-- 库存资金概览 -->
            <div class="mini-grid">
                <div class="mini-card"><div class="val" style="color:#4CAF50;">¥<%= FormatNumber(SafeNum(invCapital("raw_value")),0) %></div><div class="lbl">原料库存价值</div></div>
                <div class="mini-card"><div class="val" style="color:#2196F3;">¥<%= FormatNumber(SafeNum(invCapital("note_value")),0) %></div><div class="lbl">香调库存价值</div></div>
                <div class="mini-card"><div class="val" style="color:#FF9800;">¥<%= FormatNumber(SafeNum(invCapital("product_value")),0) %></div><div class="lbl">成品库存价值</div></div>
                <div class="mini-card"><div class="val" style="color:#9C27B0;">¥<%= FormatNumber(SafeNum(invCapital("sales_30d")),0) %></div><div class="lbl">近30天销售额</div></div>
            </div>
            
        <% ElseIf currentTab = "orders" Then %>
            <!-- ====== Tab 2: 异常订单 ====== -->
            <h3 class="section-title"><i class="fas fa-shopping-cart"></i> 异常订单识别</h3>
            <p style="color:#888;font-size:13px;margin-bottom:20px;">系统自动识别以下异常订单类型：大额订单(&gt;¥3,000)、高频重复下单(&gt;3次)、滞留超时订单(待支付&gt;3天)</p>
            
            <%
            Dim orderTabs
            orderTabs = Array("high_amount", "repeat_orders", "stale_orders")
            Dim orderTabNames, orderTabIcons, orderActiveTab
            orderTabNames = Array("大额订单", "重复下单", "长期待支付")
            orderTabIcons = Array("fa-money-bill-wave", "fa-copy", "fa-clock")
            orderActiveTab = Request.QueryString("otab")
            If orderActiveTab = "" Then orderActiveTab = "high_amount"
            
            For i = 0 To UBound(orderTabs)
                If orderActiveTab = orderTabs(i) Then
            %>
            <div style="margin-bottom:20px;">
                <h4 style="color:#e0e0e0;margin:0 0 15px;font-size:14px;"><i class="fas <%= orderTabIcons(i) %>"></i> <%= orderTabNames(i) %></h4>
                <table class="data-table">
                    <thead>
                        <tr>
                            <% If orderActiveTab = "high_amount" Then %>
                            <th>订单号</th><th>用户</th><th>金额</th><th>状态</th><th>历史订单数</th><th>创建时间</th><th>风险等级</th>
                            <% ElseIf orderActiveTab = "repeat_orders" Then %>
                            <th>电话</th><th>地址</th><th>订单数</th><th>总金额</th><th>最后下单</th><th>风险等级</th>
                            <% Else %>
                            <th>订单号</th><th>用户</th><th>金额</th><th>状态</th><th>滞留天数</th><th>创建时间</th><th>风险等级</th>
                            <% End If %>
                        </tr>
                    </thead>
                    <tbody>
                        <% 
                        Dim orderItems
                        If anomalyOrders.Exists(orderActiveTab) Then
                            orderItems = anomalyOrders(orderActiveTab)
                            If UBound(orderItems) >= LBound(orderItems) Then
                                On Error Resume Next
                                For Each oItem In orderItems
                        %>
                        <tr>
                            <% If orderActiveTab = "high_amount" Then %>
                            <td><%= oItem("OrderNo") %></td>
                            <td><%= Server.HTMLEncode(oItem("Username")) %></td>
                            <td style="color:#ffb74d;font-weight:600;">¥<%= FormatNumber(oItem("TotalAmount"),2) %></td>
                            <td><%= oItem("Status") %></td>
                            <td><%= oItem("UserOrderCount") %></td>
                            <td><%= oItem("CreatedAt") %></td>
                            <td><span class="badge <%= IIf(oItem("RiskLevel")="high","risk-high","risk-medium") %>"><%= oItem("RiskLevel") %></span></td>
                            <% ElseIf orderActiveTab = "repeat_orders" Then %>
                            <td><%= oItem("Phone") %></td>
                            <td title="<%= Server.HTMLEncode(oItem("Address")) %>"><%= Left(oItem("Address"), 20) & IIf(Len(oItem("Address"))>20,"...","") %></td>
                            <td><%= oItem("OrderCount") %></td>
                            <td>¥<%= FormatNumber(oItem("TotalSpent"),2) %></td>
                            <td><%= oItem("LastOrder") %></td>
                            <td><span class="badge <%= IIf(oItem("RiskLevel")="high","risk-high","risk-medium") %>"><%= oItem("RiskLevel") %></span></td>
                            <% Else %>
                            <td><%= oItem("OrderNo") %></td>
                            <td><%= Server.HTMLEncode(oItem("Username")) %></td>
                            <td>¥<%= FormatNumber(oItem("TotalAmount"),2) %></td>
                            <td><span class="badge risk-medium"><%= oItem("Status") %></span></td>
                            <td style="color:#f44336;"><%= oItem("PendingDays") %> 天</td>
                            <td><%= oItem("CreatedAt") %></td>
                            <td><span class="badge <%= IIf(oItem("RiskLevel")="high","risk-high","risk-low") %>"><%= oItem("RiskLevel") %></span></td>
                            <% End If %>
                        </tr>
                        <% 
                                Next
                                On Error GoTo 0
                            Else
                        %>
                        <tr><td colspan="10" style="text-align:center;color:#888;padding:30px;"><i class="fas fa-check-circle" style="color:#4CAF50;"></i> 未发现此类异常</td></tr>
                        <% 
                            End If
                        Else
                        %>
                        <tr><td colspan="10" style="text-align:center;color:#888;padding:30px;"><i class="fas fa-check-circle" style="color:#4CAF50;"></i> 未发现此类异常</td></tr>
                        <% End If %>
                    </tbody>
                </table>
            </div>
            <% 
                    Exit For
                End If
            Next 
            %>
            
            <div style="display:flex;gap:10px;margin-top:15px;flex-wrap:wrap;">
                <% For i = 0 To UBound(orderTabs) %>
                <a href="?tab=orders&otab=<%= orderTabs(i) %>" class="btn <%= IIf(orderActiveTab=orderTabs(i),"btn-primary","btn" & IIf(false,"","")) %>" style="<%= IIf(orderActiveTab=orderTabs(i),"","background:#3a3a4a;color:#e0e0e0;") %>">
                    <i class="fas <%= orderTabIcons(i) %>"></i> <%= orderTabNames(i) %>
                </a>
                <% Next %>
            </div>
        
        <% ElseIf currentTab = "inventory" Then %>
            <!-- ====== Tab 3: 库存资金 ====== -->
            <h3 class="section-title"><i class="fas fa-warehouse"></i> 库存占压资金预警</h3>
            
            <div class="mini-grid">
                <div class="mini-card"><div class="val" style="color:#4CAF50;"><%= SafeNum(invCapital("raw_count")) %></div><div class="lbl">原料种类</div></div>
                <div class="mini-card"><div class="val" style="color:#2196F3;"><%= SafeNum(invCapital("note_count")) %></div><div class="lbl">香调种类</div></div>
                <div class="mini-card"><div class="val" style="color:#FF9800;"><%= SafeNum(invCapital("product_count")) %></div><div class="lbl">成品种类</div></div>
                <div class="mini-card"><div class="val" style="color:#e57373;font-size:16px;">¥<%= FormatNumber(SafeNum(invCapital("total_inventory_value")),0) %></div><div class="lbl">总库存价值</div></div>
            </div>
            
            <h4 style="color:#e0e0e0;margin:20px 0 15px;font-size:14px;"><i class="fas fa-hourglass-half" style="color:#FF9800;"></i> 高库存值原料（Top 20）</h4>
            <table class="data-table">
                <thead><tr><th>原料名称</th><th>编码</th><th>库存量</th><th>单价</th><th>库存价值</th><th>安全库存</th><th>未动天数</th><th>建议</th></tr></thead>
                <tbody>
                    <% If invCapital.Exists("slow_moving") Then
                        Dim invItems, daysInact
                        invItems = invCapital("slow_moving")
                        If IsArray(invItems) Then
                        If UBound(invItems) >= LBound(invItems) Then
                            Dim invI
                            On Error Resume Next
                            For invI = 0 To UBound(invItems)
                                If IsObject(invItems(invI)) Then
                                    Set invItem = invItems(invI)
                    %>
                    <tr<%= IIf(SafeNum(invItem("DaysInactive")) > 90, " style='background:rgba(244,67,54,0.1);'", "") %>>
                        <td><%= Server.HTMLEncode(invItem("ItemName")) %></td>
                        <td style="color:#888;"><%= invItem("ItemCode") %></td>
                        <td><%= FormatNumber(invItem("StockQty"),0) %></td>
                        <td>¥<%= FormatNumber(invItem("UnitPrice"),2) %></td>
                        <td style="color:#ffb74d;font-weight:600;">¥<%= FormatNumber(invItem("InventoryValue"),2) %></td>
                        <td><%= FormatNumber(invItem("SafetyStock"),0) %></td>
                        <td>
                            <% daysInact = CInt(SafeNum(invItem("DaysInactive"))) %>
                            <span class="badge <%= IIf(daysInact > 90, "risk-high", IIf(daysInact > 30, "risk-medium", "risk-low")) %>"><%= daysInact %>天</span>
                        </td>
                        <td style="color:#888;font-size:12px;">
                            <%= IIf(SafeNum(invItem("StockQty")) > SafeNum(invItem("SafetyStock")) * 3, "库存偏高", IIf(SafeNum(invItem("StockQty")) < SafeNum(invItem("SafetyStock")), "低于安全库存", "正常")) %>
                        </td>
                    </tr>
                    <% 
                                End If
                            Next
                            On Error GoTo 0
                        End If
                        Else
                    %>
                    <tr><td colspan="8" style="text-align:center;color:#888;padding:30px;"><i class="fas fa-box-open" style="color:#888;margin-right:8px;"></i>暂无慢流动原料数据</td></tr>
                    <% 
                        End If
                    Else 
                    %>
                    <tr><td colspan="8" style="text-align:center;color:#888;padding:30px;"><i class="fas fa-box-open" style="color:#888;margin-right:8px;"></i>暂无慢流动原料数据</td></tr>
                    <% End If %>
                </tbody>
            </table>
        
        <% ElseIf currentTab = "sku" Then %>
            <!-- ====== Tab 4: SKU健康度 ====== -->
            <h3 class="section-title"><i class="fas fa-boxes"></i> SKU健康度评分</h3>
            <p style="color:#888;font-size:13px;margin-bottom:20px;">评分维度：销售量(30分) + 利润率(30分) + 用户评分(20分) + 收藏量(20分) = 总分</p>
            
            <div class="mini-grid">
                <div class="mini-card"><div class="val" style="color:#81c784;"><%= SafeNum(skuHealth("healthy_count")) %></div><div class="lbl">健康(≥80)</div></div>
                <div class="mini-card"><div class="val" style="color:#ffb74d;"><%= SafeNum(skuHealth("warning_count")) %></div><div class="lbl">预警(60-79)</div></div>
                <div class="mini-card"><div class="val" style="color:#e57373;"><%= SafeNum(skuHealth("risk_count")) %></div><div class="lbl">风险(40-59)</div></div>
                <div class="mini-card"><div class="val" style="color:#f44336;"><%= SafeNum(skuHealth("critical_count")) %></div><div class="lbl">危机(&lt;40)</div></div>
            </div>
            
            <table class="data-table">
                <thead><tr><th>产品名称</th><th>类型</th><th>销量</th><th>销售额</th><th>利润率</th><th>评分</th><th>收藏</th><th>健康分</th><th>等级</th></tr></thead>
                <tbody>
                    <% If skuHealth.Exists("sku_scores") Then
                        Dim skuItems2
                        skuItems2 = skuHealth("sku_scores")
                        If UBound(skuItems2) >= LBound(skuItems2) Then
                            On Error Resume Next
                            For Each skuItem In skuItems2
                                Dim hlColor
                                Select Case skuItem("HealthLevel")
                                    Case "healthy": hlColor = "#81c784"
                                    Case "warning": hlColor = "#ffb74d"
                                    Case "risk": hlColor = "#e57373"
                                    Case "critical": hlColor = "#f44336"
                                End Select
                                Dim marginPct, marginColor2, healthBadgeClass
                                If SafeNum(skuItem("BasePrice")) > 0 And SafeNum(skuItem("UnitCost")) > 0 Then
                                    marginPct = Round(((SafeNum(skuItem("BasePrice")) - SafeNum(skuItem("UnitCost"))) / SafeNum(skuItem("BasePrice"))) * 100, 1)
                                Else
                                    marginPct = 0
                                End If
                                If marginPct > 30 Then
                                    marginColor2 = "#81c784"
                                ElseIf marginPct > 10 Then
                                    marginColor2 = "#ffb74d"
                                Else
                                    marginColor2 = "#e57373"
                                End If
                                Select Case skuItem("HealthLevel")
                                    Case "healthy": healthBadgeClass = "risk-healthy"
                                    Case "warning": healthBadgeClass = "risk-warning"
                                    Case "risk": healthBadgeClass = "risk-medium"
                                    Case "critical": healthBadgeClass = "risk-critical"
                                End Select
                    %>
                    <tr>
                        <td><%= Server.HTMLEncode(skuItem("ProductName")) %></td>
                        <td><span style="color:#888;"><%= skuItem("ProductType") %></span></td>
                        <td><%= skuItem("TotalSold") %></td>
                        <td>¥<%= FormatNumber(skuItem("TotalRevenue"),0) %></td>
                        <td class="t-mar"><%= marginPct %>%</td>
                        <td><%= skuItem("AvgRating") %></td>
                        <td><%= skuItem("FavCount") %></td>
                        <td class="t-hl"><%= skuItem("HealthScore") %></td>
                        <td><span class="badge <%= healthBadgeClass %>"><%= skuItem("HealthLevel") %></span></td>
                    </tr>
                    <% 
                            Next
                            On Error GoTo 0
                        End If
                    End If 
                    %>
                </tbody>
            </table>
        
        <% ElseIf currentTab = "credit" Then %>
            <!-- ====== Tab 5: 客户信用 ====== -->
            <h3 class="section-title"><i class="fas fa-user-check"></i> 客户信用评级</h3>
            <p style="color:#888;font-size:13px;margin-bottom:20px;">评级依据：累计消费金额、退货率、取消率、未支付订单数</p>
            
            <div class="mini-grid">
                <div class="mini-card"><div class="val" style="color:#81c784;"><%= SafeNum(creditRating("a_count")) %></div><div class="lbl">A级(优质)</div></div>
                <div class="mini-card"><div class="val" style="color:#64b5f6;"><%= SafeNum(creditRating("b_count")) %></div><div class="lbl">B级(良好)</div></div>
                <div class="mini-card"><div class="val" style="color:#ffb74d;"><%= SafeNum(creditRating("c_count")) %></div><div class="lbl">C级(一般)</div></div>
                <div class="mini-card"><div class="val" style="color:#e57373;"><%= SafeNum(creditRating("d_count")) %></div><div class="lbl">D级(关注)</div></div>
            </div>
            
            <table class="data-table">
                <thead><tr><th>用户名</th><th>总订单</th><th>总消费</th><th>退货数</th><th>取消数</th><th>未支付</th><th>评分</th><th>信用等级</th></tr></thead>
                <tbody>
                    <% If creditRating.Exists("customers") Then
                        Dim custItems2
                        custItems2 = creditRating("customers")
                        If UBound(custItems2) >= LBound(custItems2) Then
                            On Error Resume Next
                            For Each custItem In custItems2
                                Dim badgeClass
                                Select Case Left(custItem("CreditLevel"), 1)
                                    Case "A": badgeClass = "badge-a"
                                    Case "B": badgeClass = "badge-b"
                                    Case "C": badgeClass = "badge-c"
                                    Case "D": badgeClass = "badge-d"
                                End Select
                    %>
                    <tr>
                        <td><%= Server.HTMLEncode(custItem("Username")) %></td>
                        <td><%= custItem("TotalOrders") %></td>
                        <td>¥<%= FormatNumber(custItem("TotalSpent"),0) %></td>
                        <td><%= custItem("ReturnCount") %></td>
                        <td><%= custItem("CancelCount") %></td>
                        <td><%= custItem("UnpaidOrders") %></td>
                        <td><%= custItem("AvgRating") %></td>
                        <td><span class="badge <%= badgeClass %>"><%= custItem("CreditLevel") %></span></td>
                    </tr>
                    <% 
                            Next
                            On Error GoTo 0
                        End If
                    End If 
                    %>
                </tbody>
            </table>
        
        <% ElseIf currentTab = "cost" Then %>
            <!-- ====== Tab 6: 成本告警 ====== -->
            <h3 class="section-title"><i class="fas fa-chart-line"></i> 成本异常波动告警</h3>
            <p style="color:#888;font-size:13px;margin-bottom:20px;">监控商品成本月度波动，波动超过5%时触发告警，超过20%为高风险</p>
            
            <h4 style="color:#e0e0e0;margin:0 0 15px;font-size:14px;"><i class="fas fa-exclamation-triangle" style="color:#f44336;"></i> 产品成本异动</h4>
            <table class="data-table">
                <thead><tr><th>产品名称</th><th>上月成本</th><th>本月成本</th><th>变动额</th><th>变动率</th><th>告警级别</th></tr></thead>
                <tbody>
                    <% If costAlert.Exists("cost_alerts") Then
                        Dim costAlertItems2, alertColor, alertClass, vColor, vrColor
                        costAlertItems2 = costAlert("cost_alerts")
                        If IsArray(costAlertItems2) Then
                        If UBound(costAlertItems2) >= LBound(costAlertItems2) Then
                            Dim costI2
                            On Error Resume Next
                            For costI2 = 0 To UBound(costAlertItems2)
                                If IsObject(costAlertItems2(costI2)) Then
                                    Set costItem2 = costAlertItems2(costI2)
                                Select Case costItem2("AlertLevel")
                                    Case "high": alertColor = "#f44336": alertClass = "risk-high"
                                    Case "medium": alertColor = "#ffb74d": alertClass = "risk-medium"
                                    Case "low": alertColor = "#64b5f6": alertClass = "risk-low"
                                End Select
                                If SafeNum(costItem2("Variance")) > 0 Then vColor = "#f44336" Else vColor = "#81c784"
                                vrColor = alertColor
                    %>
                    <tr>
                        <td><%= Server.HTMLEncode(costItem2("ProductName")) %></td>
                        <td>¥<%= FormatNumber(costItem2("LastCost"),2) %></td>
                        <td>¥<%= FormatNumber(costItem2("CurrentCost"),2) %></td>
                        <td class="t-var"><%= IIf(SafeNum(costItem2("Variance"))>0,"+","") %><%= FormatNumber(costItem2("Variance"),2) %></td>
                        <td class="t-vrate"><%= IIf(SafeNum(costItem2("VarianceRate"))>0,"+","") %><%= costItem2("VarianceRate") %>%</td>
                        <td><span class="badge <%= alertClass %>"><%= costItem2("AlertLevel") %></span></td>
                    </tr>
                    <% 
                                End If
                            Next
                            On Error GoTo 0
                        End If
                        Else
                    %>
                    <tr><td colspan="6" style="text-align:center;color:#888;padding:30px;"><i class="fas fa-check-circle" style="color:#4CAF50;"></i> 未发现成本异常波动</td></tr>
                    <% 
                        End If
                    Else 
                    %>
                    <tr><td colspan="6" style="text-align:center;color:#888;padding:30px;"><i class="fas fa-chart-bar" style="color:#888;margin-right:8px;"></i>暂无成本波动数据</td></tr>
                    <% End If %>
                </tbody>
            </table>
            
            <h4 style="color:#e0e0e0;margin:25px 0 15px;font-size:14px;"><i class="fas fa-truck" style="color:#2196F3;"></i> 最近供应商报价</h4>
            <table class="data-table">
                <thead><tr><th>物料编码</th><th>物料名称</th><th>单价</th><th>供应商</th><th>报价日期</th></tr></thead>
                <tbody>
                    <% If costAlert.Exists("supplier_prices") Then
                        Dim supItems
                        supItems = costAlert("supplier_prices")
                        If IsArray(supItems) Then
                        If UBound(supItems) >= LBound(supItems) Then
                            Dim supI
                            On Error Resume Next
                            For supI = 0 To UBound(supItems)
                                If IsObject(supItems(supI)) Then
                                    Set supItem = supItems(supI)
                    %>
                    <tr>
                        <td><%= supItem("ItemCode") %></td>
                        <td><%= Server.HTMLEncode(supItem("ItemName")) %></td>
                        <td style="color:#4CAF50;">¥<%= FormatNumber(supItem("UnitPrice"),2) %></td>
                        <td><%= Server.HTMLEncode(supItem("SupplierName")) %></td>
                        <td><%= supItem("CreatedAt") %></td>
                    </tr>
                    <% 
                                End If
                            Next
                            On Error GoTo 0
                        End If
                        End If
                    Else 
                    %>
                    <tr><td colspan="5" style="text-align:center;color:#888;padding:30px;"><i class="fas fa-truck-loading" style="color:#888;margin-right:8px;"></i>暂无供应商报价数据</td></tr>
                    <% End If %>
                </tbody>
            </table>
        <% End If %>
        </div>
    </div>
<script>
// 设置健康度进度条宽度
document.querySelectorAll('.health-bar-fill').forEach(function(el) {
    var pct = el.getAttribute('data-pct');
    if (pct) el.style.width = pct + '%';
});
</script>
<!-- 页面耗时与数据时间 -->
<% If dataLoadErrors.Count > 0 Then %>
<div style="text-align:center;padding:12px;margin:0 30px 5px;background:linear-gradient(135deg,#5e1b1b,#3a0e0e);border:1px solid rgba(244,67,54,0.3);border-radius:8px;font-size:12px;color:#e57373;">
    <i class="fas fa-exclamation-triangle"></i> <strong>数据加载异常 (</strong><%= dataLoadErrors.Count %> 项<strong>)</strong>
    <% Dim errKey: For Each errKey In dataLoadErrors %>
    <span style="margin-left:15px;color:#aaa;"><%= errKey %>: <%= dataLoadErrors(errKey) %></span>
    <% Next %>
</div>
<% End If %>
<div style="text-align:center;padding:15px;color:#555;font-size:11px;border-top:1px solid #2a2a3a;margin-top:20px;">
    <i class="fas fa-clock"></i> 页面生成耗时: <%= Round(Timer() - pageStartTime, 2) %>s | 数据更新时间: <%= Now() %>
</div>
</body>
</html>
<%
Call CloseConnection()
%>