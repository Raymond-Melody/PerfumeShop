<%
' ============================================
' V15.0 DAL - 订单数据访问层
' 依赖: dal.asp, connection.asp
' 用法: <!--#include file="dal_orders.asp"-->
' ============================================

' ============================================
' 根据ID获取订单
' ============================================
Function DAL_Orders_GetByID(orderId)
    Dim sql, params(0)
    sql = "SELECT * FROM Orders WHERE OrderID=@OrderID"
    params(0) = Array("@OrderID", DAL_adInteger, 0, CLng(orderId))
    Set DAL_Orders_GetByID = DAL_GetRow(sql, params)
End Function

' ============================================
' 根据订单号获取订单
' ============================================
Function DAL_Orders_GetByOrderNo(orderNo)
    Dim sql, params(0)
    sql = "SELECT * FROM Orders WHERE OrderNo=@OrderNo"
    params(0) = Array("@OrderNo", DAL_adVarChar, 50, orderNo)
    Set DAL_Orders_GetByOrderNo = DAL_GetRow(sql, params)
End Function

' ============================================
' 获取用户订单列表（分页）
' ============================================
Function DAL_Orders_GetByUserID(userId, page, pageSize, ByRef pageInfo)
    Dim sql, params(0)
    sql = "SELECT o.OrderID, o.OrderNo, o.TotalAmount, o.Status, o.ShippingStatus, " & _
          "o.PaymentMethod, o.CreatedAt, o.ShippedAt, o.DeliveredAt, " & _
          "o.TrackingNumber, o.ShippingCompany " & _
          "FROM Orders o WHERE o.UserID=@UserID " & _
          "ORDER BY o.CreatedAt DESC, o.OrderID DESC"
    params(0) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
    Set DAL_Orders_GetByUserID = DAL_GetListPaged(sql, params, page, pageSize, pageInfo)
End Function

' ============================================
' 获取订单列表（管理端，支持状态筛选分页）
' ============================================
Function DAL_Orders_GetList(status, shippingStatus, search, page, pageSize, ByRef pageInfo)
    Dim sql, params(), paramCount
    
    sql = "SELECT o.OrderID, o.OrderNo, o.UserID, u.Username, " & _
          "o.TotalAmount, o.Status, o.ShippingStatus, " & _
          "o.PaymentMethod, o.ShippingName, o.ShippingPhone, " & _
          "o.CreatedAt, o.UpdatedAt, o.ShippedAt, o.DeliveredAt, " & _
          "o.TrackingNumber, o.ShippingCompany, o.ChannelSource " & _
          "FROM Orders o " & _
          "LEFT JOIN Users u ON o.UserID=u.UserID WHERE 1=1"
    
    paramCount = -1
    ReDim params(0)
    
    ' 订单状态筛选
    If Not IsNull(status) And status <> "" Then
        sql = sql & " AND o.Status=@Status"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@Status", DAL_adVarChar, 20, status)
    End If
    
    ' 配送状态筛选
    If Not IsNull(shippingStatus) And shippingStatus <> "" Then
        sql = sql & " AND o.ShippingStatus=@ShippingStatus"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@ShippingStatus", DAL_adVarChar, 20, shippingStatus)
    End If
    
    ' 搜索（订单号或用户名）
    If Not IsNull(search) And search <> "" Then
        sql = sql & " AND (o.OrderNo LIKE '%' + @Search + '%' OR u.Username LIKE '%' + @Search + '%')"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@Search", DAL_adVarChar, 100, search)
    End If
    
    sql = sql & " ORDER BY o.CreatedAt DESC, o.OrderID DESC"
    
    If paramCount >= 0 Then
        Set DAL_Orders_GetList = DAL_GetListPaged(sql, params, page, pageSize, pageInfo)
    Else
        Set DAL_Orders_GetList = DAL_GetListPaged(sql, Null, page, pageSize, pageInfo)
    End If
End Function

' ============================================
' 获取订单详情（OrderDetails行）
' ============================================
Function DAL_Orders_GetDetails(orderId)
    Dim sql, params(0)
    sql = "SELECT od.DetailID, od.OrderID, od.ProductID, od.ProductName, " & _
          "od.Quantity, od.UnitPrice, od.Subtotal, " & _
          "od.CustomLabel, od.BottleName, od.VolumeML, od.VolumeName, " & _
          "od.BaseNoteName, od.MiddleNoteName, od.TopNoteName " & _
          "FROM OrderDetails od " & _
          "WHERE od.OrderID=@OrderID " & _
          "ORDER BY od.DetailID ASC"
    params(0) = Array("@OrderID", DAL_adInteger, 0, CLng(orderId))
    Set DAL_Orders_GetDetails = DAL_GetList(sql, params)
End Function

' ============================================
' 更新订单状态
' ============================================
Function DAL_Orders_UpdateStatus(orderId, newStatus)
    Dim sql, params(1)
    sql = "UPDATE Orders SET Status=@Status, UpdatedAt=GETDATE() WHERE OrderID=@OrderID"
    params(0) = Array("@Status", DAL_adVarChar, 20, newStatus)
    params(1) = Array("@OrderID", DAL_adInteger, 0, CLng(orderId))
    DAL_Orders_UpdateStatus = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' 更新配送信息
' ============================================
Function DAL_Orders_UpdateShipping(orderId, shippingStatus, trackingNumber, shippingCompany)
    Dim sql, params(3)
    sql = "UPDATE Orders SET ShippingStatus=@ShippingStatus, " & _
          "TrackingNumber=@TrackingNumber, ShippingCompany=@ShippingCompany, " & _
          "UpdatedAt=GETDATE()"
    
    ' 如果状态为"已发货"，记录发货时间
    If shippingStatus = "Shipped" Then
        sql = sql & ", ShippedAt=GETDATE()"
    End If
    ' 如果状态为"已送达"，记录送达时间
    If shippingStatus = "Delivered" Then
        sql = sql & ", DeliveredAt=GETDATE()"
    End If
    
    sql = sql & " WHERE OrderID=@OrderID"
    
    params(0) = Array("@ShippingStatus", DAL_adVarChar, 20, shippingStatus)
    params(1) = Array("@TrackingNumber", DAL_adVarChar, 100, trackingNumber)
    params(2) = Array("@ShippingCompany", DAL_adVarChar, 50, shippingCompany)
    params(3) = Array("@OrderID", DAL_adInteger, 0, CLng(orderId))
    DAL_Orders_UpdateShipping = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' 获取各状态订单数量（仪表盘用）
' ============================================
Function DAL_Orders_GetCountByStatus(status)
    Dim sql, params(0)
    sql = "SELECT COUNT(*) FROM Orders WHERE Status=@Status"
    params(0) = Array("@Status", DAL_adVarChar, 20, status)
    DAL_Orders_GetCountByStatus = CLng(DAL_GetScalar(sql, params, 0))
End Function

' ============================================
' 获取订单状态统计汇总
' ============================================
Function DAL_Orders_GetStatusSummary()
    Dim sql
    sql = "SELECT Status, COUNT(*) AS OrderCount, ISNULL(SUM(TotalAmount), 0) AS TotalAmount " & _
          "FROM Orders GROUP BY Status ORDER BY Status"
    Set DAL_Orders_GetStatusSummary = DAL_GetList(sql, Null)
End Function

' ============================================
' 获取营收统计（按日期范围）
' ============================================
Function DAL_Orders_GetRevenueStats(startDate, endDate)
    Dim sql, params(1)
    sql = "SELECT ISNULL(COUNT(*), 0) AS OrderCount, " & _
          "ISNULL(SUM(TotalAmount), 0) AS TotalRevenue, " & _
          "ISNULL(SUM(ShippingFee), 0) AS TotalShippingFee, " & _
          "ISNULL(SUM(CostAmount), 0) AS TotalCost, " & _
          "ISNULL(SUM(ProfitAmount), 0) AS TotalProfit, " & _
          "ISNULL(SUM(RefundAmount), 0) AS TotalRefund " & _
          "FROM Orders WHERE Status <> 'Cancelled' " & _
          "AND CreatedAt >= @StartDate AND CreatedAt < DATEADD(DAY, 1, @EndDate)"
    params(0) = Array("@StartDate", DAL_adDBTimeStamp, 0, startDate)
    params(1) = Array("@EndDate", DAL_adDBTimeStamp, 0, endDate)
    Set DAL_Orders_GetRevenueStats = DAL_GetRow(sql, params)
End Function

' ============================================
' 获取每日营收趋势
' ============================================
Function DAL_Orders_GetDailyRevenue(days)
    Dim sql
    If IsNull(days) Or days < 1 Then days = 30
    sql = "SELECT CAST(CreatedAt AS DATE) AS OrderDate, " & _
          "COUNT(*) AS OrderCount, " & _
          "ISNULL(SUM(TotalAmount), 0) AS DailyRevenue " & _
          "FROM Orders WHERE Status <> 'Cancelled' " & _
          "AND CreatedAt >= DATEADD(DAY, -" & CLng(days) & ", GETDATE()) " & _
          "GROUP BY CAST(CreatedAt AS DATE) " & _
          "ORDER BY OrderDate ASC"
    Set DAL_Orders_GetDailyRevenue = DAL_GetList(sql, Null)
End Function

' ============================================
' 检查订单是否存在
' ============================================
Function DAL_Orders_Exists(orderId)
    DAL_Orders_Exists = DAL_Exists("Orders", "OrderID=@OrderID", _
        Array(Array("@OrderID", DAL_adInteger, 0, CLng(orderId))))
End Function

' ============================================
' 获取订单支付记录
' ============================================
Function DAL_Orders_GetPayments(orderId)
    Dim sql, params(0)
    sql = "SELECT RecordID, OrderID, OrderNo, Amount, Fee, NetAmount, " & _
          "PaymentMethod, TransactionNo, TransactionType, Status, " & _
          "Category, Remark, ReconcileStatus, CreatedAt " & _
          "FROM PaymentRecords WHERE OrderID=@OrderID ORDER BY CreatedAt DESC"
    params(0) = Array("@OrderID", DAL_adInteger, 0, CLng(orderId))
    Set DAL_Orders_GetPayments = DAL_GetList(sql, params)
End Function

' ============================================
' 获取订单退款记录
' ============================================
Function DAL_Orders_GetRefunds(orderId)
    Dim sql, params(0)
    sql = "SELECT RefundID, OrderID, OrderNo, RefundAmount, RefundNo, " & _
          "RefundReason, Status, CostWriteBack, CreatedAt, ApprovedAt, CompletedAt " & _
          "FROM RefundRecords WHERE OrderID=@OrderID ORDER BY CreatedAt DESC"
    params(0) = Array("@OrderID", DAL_adInteger, 0, CLng(orderId))
    Set DAL_Orders_GetRefunds = DAL_GetList(sql, params)
End Function

' ============================================
' 获取近期订单（首页仪表盘用）
' ============================================
Function DAL_Orders_GetRecent(limit)
    Dim sql
    If IsNull(limit) Or limit < 1 Then limit = 10
    sql = "SELECT TOP " & CLng(limit) & " o.OrderID, o.OrderNo, o.TotalAmount, " & _
          "o.Status, o.ShippingStatus, o.PaymentMethod, o.CreatedAt, " & _
          "u.Username " & _
          "FROM Orders o " & _
          "LEFT JOIN Users u ON o.UserID=u.UserID " & _
          "ORDER BY o.CreatedAt DESC"
    Set DAL_Orders_GetRecent = DAL_GetList(sql, Null)
End Function

' ============================================
' 获取用户最近一笔订单
' ============================================
Function DAL_Orders_GetLatestByUser(userId)
    Dim sql, params(0)
    sql = "SELECT TOP 1 * FROM Orders WHERE UserID=@UserID ORDER BY CreatedAt DESC"
    params(0) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
    Set DAL_Orders_GetLatestByUser = DAL_GetRow(sql, params)
End Function
%>