<%
' ============================================
' V17.0 DAL - 结算流程数据访问层
' 依赖: dal.asp, connection.asp
' 用法: <!--#include file="dal_checkout.asp"-->
' 涵盖: Orders, OrderItems, OrderAddresses
' ============================================

' ============================================
' 创建订单
' ============================================
Function DAL_Checkout_CreateOrder(userId, orderNo, totalAmount, shippingFee, discountAmount, paymentMethod, shippingAddress, notes)
    Dim sql, params(7)
    sql = "INSERT INTO Orders (UserID, OrderNo, TotalAmount, ShippingFee, DiscountAmount, " & _
          "PaymentMethod, ShippingAddress, Notes, Status, OrderDate) " & _
          "VALUES (@UserID, @OrderNo, @TotalAmount, @ShippingFee, @DiscountAmount, " & _
          "@PaymentMethod, @ShippingAddress, @Notes, 'Pending', GETDATE()); SELECT SCOPE_IDENTITY()"
    params(0) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
    params(1) = Array("@OrderNo", DAL_adVarChar, 50, Left(orderNo, 50))
    params(2) = Array("@TotalAmount", DAL_adCurrency, 0, CDbl(totalAmount))
    params(3) = Array("@ShippingFee", DAL_adCurrency, 0, CDbl(shippingFee))
    params(4) = Array("@DiscountAmount", DAL_adCurrency, 0, CDbl(discountAmount))
    params(5) = Array("@PaymentMethod", DAL_adVarChar, 30, Left(paymentMethod, 30))
    params(6) = Array("@ShippingAddress", DAL_adVarChar, 500, Left(shippingAddress, 500))
    params(7) = Array("@Notes", DAL_adVarChar, 500, Left(notes, 500))
    DAL_Checkout_CreateOrder = CLng(DAL_GetScalar(sql, params, 0))
End Function

' ============================================
' 添加订单商品明细
' ============================================
Function DAL_Checkout_AddOrderItem(orderId, productId, productName, quantity, unitPrice, subTotal, volumeInfo, bottleInfo, customLabel)
    Dim sql, params(9)
    sql = "INSERT INTO OrderItems (OrderID, ProductID, ProductName, Quantity, " & _
          "UnitPrice, SubTotal, VolumeInfo, BottleInfo, CustomLabel) " & _
          "VALUES (@OrderID, @ProductID, @ProductName, @Quantity, " & _
          "@UnitPrice, @SubTotal, @VolumeInfo, @BottleInfo, @CustomLabel); SELECT SCOPE_IDENTITY()"
    params(0) = Array("@OrderID", DAL_adInteger, 0, CLng(orderId))
    params(1) = Array("@ProductID", DAL_adInteger, 0, CLng(productId))
    params(2) = Array("@ProductName", DAL_adVarChar, 200, Left(productName, 200))
    params(3) = Array("@Quantity", DAL_adInteger, 0, CInt(quantity))
    params(4) = Array("@UnitPrice", DAL_adCurrency, 0, CDbl(unitPrice))
    params(5) = Array("@SubTotal", DAL_adCurrency, 0, CDbl(subTotal))
    params(6) = Array("@VolumeInfo", DAL_adVarChar, 100, Left(volumeInfo, 100))
    params(7) = Array("@BottleInfo", DAL_adVarChar, 100, Left(bottleInfo, 100))
    params(8) = Array("@CustomLabel", DAL_adVarChar, 200, Left(customLabel, 200))
    DAL_Checkout_AddOrderItem = CLng(DAL_GetScalar(sql, params, 0))
End Function

' ============================================
' 保存订单地址
' ============================================
Function DAL_Checkout_SaveAddress(orderId, province, city, district, address, zipCode, receiverName, receiverPhone)
    Dim sql, params(8)
    sql = "INSERT INTO OrderAddresses (OrderID, Province, City, District, " & _
          "Address, ZipCode, ReceiverName, ReceiverPhone) " & _
          "VALUES (@OrderID, @Province, @City, @District, " & _
          "@Address, @ZipCode, @ReceiverName, @ReceiverPhone); SELECT SCOPE_IDENTITY()"
    params(0) = Array("@OrderID", DAL_adInteger, 0, CLng(orderId))
    params(1) = Array("@Province", DAL_adVarChar, 50, Left(province, 50))
    params(2) = Array("@City", DAL_adVarChar, 50, Left(city, 50))
    params(3) = Array("@District", DAL_adVarChar, 50, Left(district, 50))
    params(4) = Array("@Address", DAL_adVarChar, 200, Left(address, 200))
    params(5) = Array("@ZipCode", DAL_adVarChar, 10, Left(zipCode, 10))
    params(6) = Array("@ReceiverName", DAL_adVarChar, 50, Left(receiverName, 50))
    params(7) = Array("@ReceiverPhone", DAL_adVarChar, 20, Left(receiverPhone, 20))
    DAL_Checkout_SaveAddress = CLng(DAL_GetScalar(sql, params, 0))
End Function

' ============================================
' 根据订单号获取订单
' ============================================
Function DAL_Checkout_GetByOrderNo(orderNo)
    Dim sql, params(0)
    sql = "SELECT * FROM Orders WHERE OrderNo=@OrderNo"
    params(0) = Array("@OrderNo", DAL_adVarChar, 50, orderNo)
    Set DAL_Checkout_GetByOrderNo = DAL_GetRow(sql, params)
End Function

' ============================================
' 根据ID获取订单
' ============================================
Function DAL_Checkout_GetByID(orderId)
    Dim sql, params(0)
    sql = "SELECT * FROM Orders WHERE OrderID=@OrderID"
    params(0) = Array("@OrderID", DAL_adInteger, 0, CLng(orderId))
    Set DAL_Checkout_GetByID = DAL_GetRow(sql, params)
End Function

' ============================================
' 获取订单商品明细
' ============================================
Function DAL_Checkout_GetItems(orderId)
    Dim sql, params(0)
    sql = "SELECT * FROM OrderItems WHERE OrderID=@OrderID"
    params(0) = Array("@OrderID", DAL_adInteger, 0, CLng(orderId))
    Set DAL_Checkout_GetItems = DAL_GetList(sql, params)
End Function

' ============================================
' 获取订单地址
' ============================================
Function DAL_Checkout_GetAddress(orderId)
    Dim sql, params(0)
    sql = "SELECT * FROM OrderAddresses WHERE OrderID=@OrderID"
    params(0) = Array("@OrderID", DAL_adInteger, 0, CLng(orderId))
    Set DAL_Checkout_GetAddress = DAL_GetRow(sql, params)
End Function

' ============================================
' 更新订单状态
' ============================================
Function DAL_Checkout_UpdateStatus(orderId, status)
    Dim sql, params(1)
    sql = "UPDATE Orders SET Status=@Status, UpdatedAt=GETDATE() WHERE OrderID=@OrderID"
    params(0) = Array("@Status", DAL_adVarChar, 30, Left(status, 30))
    params(1) = Array("@OrderID", DAL_adInteger, 0, CLng(orderId))
    DAL_Checkout_UpdateStatus = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' 更新订单支付状态
' ============================================
Function DAL_Checkout_UpdatePayment(orderId, paymentMethod, transactionId)
    Dim sql, params(2)
    sql = "UPDATE Orders SET PaymentMethod=@PaymentMethod, TransactionID=@TransactionID, " & _
          "PaidAt=GETDATE(), Status='Paid' WHERE OrderID=@OrderID"
    params(0) = Array("@PaymentMethod", DAL_adVarChar, 30, Left(paymentMethod, 30))
    params(1) = Array("@TransactionID", DAL_adVarChar, 100, Left(transactionId, 100))
    params(2) = Array("@OrderID", DAL_adInteger, 0, CLng(orderId))
    DAL_Checkout_UpdatePayment = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' 获取用户订单列表
' ============================================
Function DAL_Checkout_GetUserOrders(userId, page, pageSize, ByRef pageInfo)
    Dim sql, params(0)
    sql = "SELECT OrderID, OrderNo, TotalAmount, Status, PaymentMethod, " & _
          "ShippingStatus, OrderDate FROM Orders WHERE UserID=@UserID " & _
          "ORDER BY OrderDate DESC"
    params(0) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
    Set DAL_Checkout_GetUserOrders = DAL_GetListPaged(sql, params, page, pageSize, pageInfo)
End Function

' ============================================
' 获取所有订单（后台管理）
' ============================================
Function DAL_Checkout_GetAllOrders(statusFilter, search, page, pageSize, ByRef pageInfo)
    Dim sql, params(), paramCount
    
    sql = "SELECT o.OrderID, o.OrderNo, o.UserID, u.Username, " & _
          "o.TotalAmount, o.Status, o.PaymentMethod, o.ShippingStatus, o.OrderDate " & _
          "FROM Orders o LEFT JOIN Users u ON o.UserID=u.UserID WHERE 1=1"
    paramCount = -1
    ReDim params(0)
    
    If statusFilter <> "" Then
        sql = sql & " AND o.Status=@Status"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@Status", DAL_adVarChar, 30, statusFilter)
    End If
    
    If search <> "" Then
        sql = sql & " AND (o.OrderNo LIKE '%' + @Search + '%' OR u.Username LIKE '%' + @Search + '%')"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@Search", DAL_adVarChar, 100, search)
    End If
    
    sql = sql & " ORDER BY o.OrderDate DESC"
    
    If paramCount >= 0 Then
        Set DAL_Checkout_GetAllOrders = DAL_GetListPaged(sql, params, page, pageSize, pageInfo)
    Else
        Set DAL_Checkout_GetAllOrders = DAL_GetListPaged(sql, Null, page, pageSize, pageInfo)
    End If
End Function

' ============================================
' 合并购物车到订单（批量转移购物车项到订单）
' ============================================
Function DAL_Checkout_MergeCartToOrder(cartIds, orderId)
    Dim i, id, cartItem, sql, params(1)
    Dim successCount : successCount = 0
    
    If Not IsArray(cartIds) Then
        DAL_Checkout_MergeCartToOrder = 0
        Exit Function
    End If
    
    For i = 0 To UBound(cartIds)
        id = CLng(cartIds(i))
        If id > 0 Then
            ' 获取购物车项
            Set cartItem = DAL_Cart_GetByID(id)
            If Not cartItem Is Nothing Then
                ' 添加到订单明细
                If DAL_Checkout_AddOrderItem(orderId, _
                    cartItem("ProductID"), "", _
                    cartItem("Quantity"), cartItem("UnitPrice"), _
                    CDbl(cartItem("Quantity")) * CDbl(cartItem("UnitPrice")), _
                    "", "", "") > 0 Then
                    successCount = successCount + 1
                End If
                ' 删除购物车项
                DAL_Cart_Remove id
            End If
        End If
    Next
    
    DAL_Checkout_MergeCartToOrder = successCount
End Function
%>
