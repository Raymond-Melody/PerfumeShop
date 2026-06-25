<%
' ============================================
' V17.0 DAL - 财务数据访问层
' 依赖: dal.asp, connection.asp
' 用法: <!--#include file="dal_finance.asp"-->
' 涵盖：收入报表、利润分析、支付统计、成本中心
' ============================================

' ============================================
' 月度收入汇总
' 返回: Year, Month, OrderCount, Revenue, AvgOrderValue
' ============================================
Function DAL_Fin_GetMonthlyRevenue(startDate, endDate)
    Dim sql, params(1)
    sql = "SELECT Year(OrderDate) AS Y, Month(OrderDate) AS M, " & _
          "COUNT(*) AS OrderCount, SUM(CAST(TotalAmount AS FLOAT)) AS Revenue " & _
          "FROM Orders WHERE OrderDate >= @StartDate AND OrderDate <= @EndDate " & _
          "GROUP BY Year(OrderDate), Month(OrderDate) " & _
          "ORDER BY Year(OrderDate) DESC, Month(OrderDate) DESC"
    params(0) = Array("@StartDate", DAL_adVarChar, 20, startDate)
    params(1) = Array("@EndDate", DAL_adVarChar, 20, endDate)
    Set DAL_Fin_GetMonthlyRevenue = DAL_GetList(sql, params)
End Function

' ============================================
' 订单金额分布统计
' 返回各金额区间的订单数量
' ============================================
Sub DAL_Fin_GetOrderDistribution(startDate, endDate, ByRef cnt0_50, ByRef cnt50_100, ByRef cnt100_200, ByRef cnt200_500, ByRef cnt500plus)
    Dim sql, params(1)
    params(0) = Array("@StartDate", DAL_adVarChar, 20, startDate)
    params(1) = Array("@EndDate", DAL_adVarChar, 20, endDate)
    
    cnt0_50 = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM Orders WHERE TotalAmount >= 0 AND TotalAmount < 50 AND OrderDate >= @StartDate AND OrderDate <= @EndDate", params, 0))
    cnt50_100 = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM Orders WHERE TotalAmount >= 50 AND TotalAmount < 100 AND OrderDate >= @StartDate AND OrderDate <= @EndDate", params, 0))
    cnt100_200 = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM Orders WHERE TotalAmount >= 100 AND TotalAmount < 200 AND OrderDate >= @StartDate AND OrderDate <= @EndDate", params, 0))
    cnt200_500 = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM Orders WHERE TotalAmount >= 200 AND TotalAmount < 500 AND OrderDate >= @StartDate AND OrderDate <= @EndDate", params, 0))
    cnt500plus = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM Orders WHERE TotalAmount >= 500 AND OrderDate >= @StartDate AND OrderDate <= @EndDate", params, 0))
End Sub

' ============================================
' 支付方式占比统计
' ============================================
Function DAL_Fin_GetPaymentStats(startDate, endDate)
    Dim sql, params(1)
    sql = "SELECT PaymentMethod, COUNT(*) AS Cnt, SUM(CAST(TotalAmount AS FLOAT)) AS Total " & _
          "FROM Orders WHERE OrderDate >= @StartDate AND OrderDate <= @EndDate " & _
          "GROUP BY PaymentMethod"
    params(0) = Array("@StartDate", DAL_adVarChar, 20, startDate)
    params(1) = Array("@EndDate", DAL_adVarChar, 20, endDate)
    Set DAL_Fin_GetPaymentStats = DAL_GetList(sql, params)
End Function

' ============================================
' 总收入（指定日期范围）
' ============================================
Function DAL_Fin_GetTotalRevenue(startDate, endDate)
    Dim sql, params(1)
    sql = "SELECT ISNULL(SUM(CAST(TotalAmount AS FLOAT)), 0) FROM Orders " & _
          "WHERE OrderDate >= @StartDate AND OrderDate <= @EndDate"
    params(0) = Array("@StartDate", DAL_adVarChar, 20, startDate)
    params(1) = Array("@EndDate", DAL_adVarChar, 20, endDate)
    DAL_Fin_GetTotalRevenue = CDbl(DAL_GetScalar(sql, params, 0))
End Function

' ============================================
' 订单总数（指定日期范围）
' ============================================
Function DAL_Fin_GetOrderCount(startDate, endDate)
    Dim sql, params(1)
    sql = "SELECT COUNT(*) FROM Orders WHERE OrderDate >= @StartDate AND OrderDate <= @EndDate"
    params(0) = Array("@StartDate", DAL_adVarChar, 20, startDate)
    params(1) = Array("@EndDate", DAL_adVarChar, 20, endDate)
    DAL_Fin_GetOrderCount = CLng(DAL_GetScalar(sql, params, 0))
End Function

' ============================================
' 利润统计（收入 - 成本）
' ============================================
Function DAL_Fin_GetProfit(startDate, endDate)
    Dim sql, params(1)
    sql = "SELECT ISNULL(SUM(CAST(TotalAmount AS FLOAT)), 0) - " & _
          "ISNULL(SUM(CAST(TotalCost AS FLOAT)), 0) FROM Orders " & _
          "WHERE OrderDate >= @StartDate AND OrderDate <= @EndDate"
    params(0) = Array("@StartDate", DAL_adVarChar, 20, startDate)
    params(1) = Array("@EndDate", DAL_adVarChar, 20, endDate)
    DAL_Fin_GetProfit = CDbl(DAL_GetScalar(sql, params, 0))
End Function

' ============================================
' 按订单状态统计
' ============================================
Function DAL_Fin_GetOrderStatusStats(startDate, endDate)
    Dim sql, params(1)
    sql = "SELECT Status, COUNT(*) AS Cnt, SUM(CAST(TotalAmount AS FLOAT)) AS Total " & _
          "FROM Orders WHERE OrderDate >= @StartDate AND OrderDate <= @EndDate " & _
          "GROUP BY Status ORDER BY Cnt DESC"
    params(0) = Array("@StartDate", DAL_adVarChar, 20, startDate)
    params(1) = Array("@EndDate", DAL_adVarChar, 20, endDate)
    Set DAL_Fin_GetOrderStatusStats = DAL_GetList(sql, params)
End Function

' ============================================
' 获取最近订单列表（后台财务用）
' ============================================
Function DAL_Fin_GetRecentOrders(limit, status)
    Dim sql
    sql = "SELECT TOP " & CLng(limit) & " OrderID, OrderNo, UserID, TotalAmount, " & _
          "Status, PaymentMethod, OrderDate, ShippingStatus " & _
          "FROM Orders WHERE 1=1"
    
    If status <> "" Then
        sql = sql & " AND Status='" & Replace(status, "'", "''") & "'"
    End If
    sql = sql & " ORDER BY OrderDate DESC"
    
    Set DAL_Fin_GetRecentOrders = DAL_GetList(sql, Null)
End Function

' ============================================
' 按产品类型统计销售额
' ============================================
Function DAL_Fin_GetSalesByProductType(startDate, endDate)
    Dim sql, params(1)
    sql = "SELECT p.ProductType, COUNT(DISTINCT o.OrderID) AS OrderCount, " & _
          "SUM(oi.Quantity) AS TotalQty, SUM(oi.UnitPrice * oi.Quantity) AS Revenue " & _
          "FROM Orders o INNER JOIN OrderItems oi ON o.OrderID=oi.OrderID " & _
          "INNER JOIN Products p ON oi.ProductID=p.ProductID " & _
          "WHERE o.OrderDate >= @StartDate AND o.OrderDate <= @EndDate " & _
          "GROUP BY p.ProductType ORDER BY Revenue DESC"
    params(0) = Array("@StartDate", DAL_adVarChar, 20, startDate)
    params(1) = Array("@EndDate", DAL_adVarChar, 20, endDate)
    Set DAL_Fin_GetSalesByProductType = DAL_GetList(sql, params)
End Function

' ============================================
' 订单详情（含商品明细）
' ============================================
Function DAL_Fin_GetOrderDetail(orderId)
    Dim sql, params(0)
    sql = "SELECT o.*, oi.OrderItemID, oi.ProductID, oi.ProductName, " & _
          "oi.Quantity, oi.UnitPrice, oi.SubTotal " & _
          "FROM Orders o LEFT JOIN OrderItems oi ON o.OrderID=oi.OrderID " & _
          "WHERE o.OrderID=@OrderID"
    params(0) = Array("@OrderID", DAL_adInteger, 0, CLng(orderId))
    Set DAL_Fin_GetOrderDetail = DAL_GetList(sql, params)
End Function

' ============================================
' 按地区统计销售额
' ============================================
Function DAL_Fin_GetSalesByRegion(startDate, endDate)
    Dim sql, params(1)
    sql = "SELECT oa.Province, COUNT(DISTINCT o.OrderID) AS OrderCount, " & _
          "SUM(CAST(o.TotalAmount AS FLOAT)) AS Revenue " & _
          "FROM Orders o INNER JOIN OrderAddresses oa ON o.OrderID=oa.OrderID " & _
          "WHERE o.OrderDate >= @StartDate AND o.OrderDate <= @EndDate " & _
          "GROUP BY oa.Province ORDER BY Revenue DESC"
    params(0) = Array("@StartDate", DAL_adVarChar, 20, startDate)
    params(1) = Array("@EndDate", DAL_adVarChar, 20, endDate)
    Set DAL_Fin_GetSalesByRegion = DAL_GetList(sql, params)
End Function

' ============================================
' 客户消费排名
' ============================================
Function DAL_Fin_GetTopCustomers(limit, startDate, endDate)
    Dim sql, params(2)
    sql = "SELECT TOP " & CLng(limit) & " o.UserID, u.Username, u.FullName, " & _
          "COUNT(*) AS OrderCount, SUM(CAST(o.TotalAmount AS FLOAT)) AS TotalSpent " & _
          "FROM Orders o INNER JOIN Users u ON o.UserID=u.UserID " & _
          "WHERE o.OrderDate >= @StartDate AND o.OrderDate <= @EndDate " & _
          "GROUP BY o.UserID, u.Username, u.FullName " & _
          "ORDER BY TotalSpent DESC"
    params(0) = Array("@StartDate", DAL_adVarChar, 20, startDate)
    params(1) = Array("@EndDate", DAL_adVarChar, 20, endDate)
    Set DAL_Fin_GetTopCustomers = DAL_GetList(sql, params)
End Function

' ============================================
' 成本中心 - 获取成本汇总
' ============================================
Function DAL_Fin_GetCostSummary(startDate, endDate)
    Dim sql, params(1)
    sql = "SELECT CostCategory, SUM(Amount) AS TotalAmount, COUNT(*) AS ItemCount " & _
          "FROM CostRecords WHERE RecordDate >= @StartDate AND RecordDate <= @EndDate " & _
          "GROUP BY CostCategory ORDER BY TotalAmount DESC"
    params(0) = Array("@StartDate", DAL_adVarChar, 20, startDate)
    params(1) = Array("@EndDate", DAL_adVarChar, 20, endDate)
    Set DAL_Fin_GetCostSummary = DAL_GetList(sql, params)
End Function

' ============================================
' 获取指定日期范围的订单（分页，用于报表导出）
' ============================================
Function DAL_Fin_GetOrderList(startDate, endDate, statusFilter, page, pageSize, ByRef pageInfo)
    Dim sql, params(), paramCount
    
    sql = "SELECT OrderID, OrderNo, UserID, TotalAmount, Status, " & _
          "PaymentMethod, ShippingStatus, OrderDate, PaidAt, ShippedAt " & _
          "FROM Orders WHERE 1=1"
    paramCount = -1
    ReDim params(0)
    
    If startDate <> "" Then
        sql = sql & " AND OrderDate >= @StartDate"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@StartDate", DAL_adVarChar, 20, startDate)
    End If
    
    If endDate <> "" Then
        sql = sql & " AND OrderDate <= @EndDate"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@EndDate", DAL_adVarChar, 20, endDate & " 23:59:59")
    End If
    
    If statusFilter <> "" Then
        sql = sql & " AND Status=@Status"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@Status", DAL_adVarChar, 30, statusFilter)
    End If
    
    sql = sql & " ORDER BY OrderDate DESC"
    
    If paramCount >= 0 Then
        Set DAL_Fin_GetOrderList = DAL_GetListPaged(sql, params, page, pageSize, pageInfo)
    Else
        Set DAL_Fin_GetOrderList = DAL_GetListPaged(sql, Null, page, pageSize, pageInfo)
    End If
End Function

' ============================================
' 获取仪表盘关键指标
' ============================================
Sub DAL_Fin_GetDashboardMetrics(ByRef todayRevenue, ByRef monthRevenue, ByRef pendingOrders, ByRef totalProducts)
    todayRevenue = CDbl(DAL_GetScalar( _
        "SELECT ISNULL(SUM(CAST(TotalAmount AS FLOAT)), 0) FROM Orders WHERE CAST(OrderDate AS DATE)=CAST(GETDATE() AS DATE)", Null, 0))
    monthRevenue = CDbl(DAL_GetScalar( _
        "SELECT ISNULL(SUM(CAST(TotalAmount AS FLOAT)), 0) FROM Orders WHERE OrderDate >= DATEADD(day, 1, EOMONTH(GETDATE(), -1))", Null, 0))
    pendingOrders = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM Orders WHERE Status='Pending' OR Status='Processing'", Null, 0))
    totalProducts = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM Products WHERE IsActive=1", Null, 0))
End Sub

' ============================================
' 年度对比 - 按月统计
' ============================================
Function DAL_Fin_GetYearlyComparison(year)
    Dim sql, params(0)
    sql = "SELECT Month(OrderDate) AS M, COUNT(*) AS OrderCount, " & _
          "SUM(CAST(TotalAmount AS FLOAT)) AS Revenue " & _
          "FROM Orders WHERE Year(OrderDate)=@Year " & _
          "GROUP BY Month(OrderDate) ORDER BY M"
    params(0) = Array("@Year", DAL_adInteger, 0, CInt(year))
    Set DAL_Fin_GetYearlyComparison = DAL_GetList(sql, params)
End Function
%>
