<!--#include file="payment_config.asp"-->
<%
' ============================================
' Payment Handler Module
' ============================================

' Create payment order (compatible with old version)
Function CreatePaymentOrder(userId, orderAmount, orderDesc, paymentMethod)
    CreatePaymentOrder = SafeCreatePaymentOrder(userId, orderAmount, orderDesc, paymentMethod, "", "", "")
End Function

' Safe create payment order function
Function SafeCreatePaymentOrder(userId, orderAmount, orderDesc, paymentMethod, shippingName, shippingPhone, shippingAddress)
    Dim orderId, orderNo, sql, insertSuccess
    orderId = 0
    
    ' Validate order amount
    If Not IsNumeric(orderAmount) Or orderAmount <= 0 Then
        Session("ErrorMessage") = "Invalid order amount"
        SafeCreatePaymentOrder = 0
        Exit Function
    End If
    
    ' Validate payment method
    If paymentMethod = "" Then
        Session("ErrorMessage") = "Payment method is empty"
        SafeCreatePaymentOrder = 0
        Exit Function
    End If
    
    If Not IsNumeric(paymentMethod) Then
        Session("ErrorMessage") = "Payment method is not numeric: " & paymentMethod
        SafeCreatePaymentOrder = 0
        Exit Function
    End If
    
    ' Check if payment method is valid
    Dim validPaymentMethod, pmValue
    pmValue = CLng(paymentMethod)
    validPaymentMethod = False
    If pmValue = 1 Then validPaymentMethod = True
    If pmValue = 2 Then validPaymentMethod = True
    If pmValue = 3 Then validPaymentMethod = True
    If pmValue = 4 Then validPaymentMethod = True
    If Not validPaymentMethod Then
        Session("ErrorMessage") = "Invalid payment method: " & paymentMethod
        SafeCreatePaymentOrder = 0
        Exit Function
    End If
    
    ' Validate user ID
    If Not IsNumeric(userId) Or userId <= 0 Then
        Session("ErrorMessage") = "Invalid user ID"
        SafeCreatePaymentOrder = 0
        Exit Function
    End If
    
    ' Validate order description
    If IsNull(orderDesc) Then
        orderDesc = ""
    End If
    
    ' Validate shipping info (optional)
    If IsNull(shippingName) Then shippingName = ""
    If IsNull(shippingPhone) Then shippingPhone = ""
    If IsNull(shippingAddress) Then shippingAddress = ""
    
    ' Generate order number
    orderNo = "ORD" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & _
              Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & _
              Right(userId, 4)
    
    ' Validate order number generation
    If orderNo = "" Then
        Session("ErrorMessage") = "Order number generation failed"
        SafeCreatePaymentOrder = 0
        Exit Function
    End If
    
    ' Build SQL for inserting order
    sql = "INSERT INTO Orders (OrderNo, UserID, TotalAmount, Notes, PaymentMethod, Status, ShippingName, ShippingPhone, ShippingAddress, CreatedAt) " & _
          "VALUES ('" & SafeSQL(orderNo) & "', " & CLng(userId) & ", " & CDbl(orderAmount) & ", '" & _
          SafeSQL(orderDesc) & "', " & CLng(paymentMethod) & ", 'Pending', '" & _
          SafeSQL(shippingName) & "', '" & SafeSQL(shippingPhone) & "', '" & SafeSQL(shippingAddress) & "', GETDATE())"
    
    ' Execute INSERT
    insertSuccess = ExecuteNonQuery(sql)
    
    If insertSuccess Then
        ' Get new order ID
        sql = "SELECT TOP 1 OrderID FROM Orders WHERE OrderNo = '" & SafeSQL(orderNo) & "' ORDER BY CreatedAt DESC"
        Dim rs
        Set rs = ExecuteQuery(sql)
        
        If Not rs Is Nothing Then
            If Not rs.EOF Then
                orderId = rs("OrderID")
            End If
            rs.Close
            Set rs = Nothing
        End If
    Else
        Dim dbError
        dbError = Session("LastDBError")
        Session("ErrorMessage") = "Database error: " & dbError
        SafeCreatePaymentOrder = 0
        Exit Function
    End If
    
    SafeCreatePaymentOrder = orderId
End Function

' Process WeChat Pay
Function ProcessWeChatPay(orderId)
    Dim result, updateSuccess
    result = False
    
    If Not WECHAT_PAY_ENABLED Then
        Session("ErrorMessage") = "WeChat Pay not enabled"
        ProcessWeChatPay = result
        Exit Function
    End If
    
    Dim rsOrder
    Set rsOrder = ExecuteQuery("SELECT OrderNo, TotalAmount, Notes FROM Orders WHERE OrderID = " & orderId)
    If rsOrder Is Nothing Or rsOrder.EOF Then
        Session("ErrorMessage") = "Order not found"
        ProcessWeChatPay = result
        Exit Function
    End If
    
    Dim orderNo, orderAmount, orderDesc
    orderNo = rsOrder("OrderNo")
    orderAmount = rsOrder("TotalAmount")
    orderDesc = rsOrder("Notes")
    rsOrder.Close
    Set rsOrder = Nothing
    
    If WECHAT_PAY_APPID = "" Or WECHAT_PAY_MCH_ID = "" Or WECHAT_PAY_KEY = "" Then
        result = True
        updateSuccess = UpdateOrderPaymentStatus(orderId, 1, "WECHAT-DEV-" & orderId)
        If Not updateSuccess Then
            result = False
        End If
    Else
        On Error Resume Next
        result = True
        updateSuccess = UpdateOrderPaymentStatus(orderId, 1, "WECHAT-API-" & orderId)
        If Not updateSuccess Then
            result = False
        End If
        On Error Goto 0
    End If
    
    ProcessWeChatPay = result
End Function

' Process Alipay
Function ProcessAlipay(orderId)
    Dim result, updateSuccess
    result = False
    
    If Not ALIPAY_ENABLED Then
        Session("ErrorMessage") = "Alipay not enabled"
        ProcessAlipay = result
        Exit Function
    End If
    
    Dim rsOrder
    Set rsOrder = ExecuteQuery("SELECT OrderNo, TotalAmount, Notes FROM Orders WHERE OrderID = " & orderId)
    If rsOrder Is Nothing Or rsOrder.EOF Then
        Session("ErrorMessage") = "Order not found"
        ProcessAlipay = result
        Exit Function
    End If
    
    Dim orderNo, orderAmount, orderDesc
    orderNo = rsOrder("OrderNo")
    orderAmount = rsOrder("TotalAmount")
    orderDesc = rsOrder("Notes")
    rsOrder.Close
    Set rsOrder = Nothing
    
    If ALIPAY_APP_ID = "" Or ALIPAY_PRIVATE_KEY = "" Or ALIPAY_PUBLIC_KEY = "" Then
        result = True
        updateSuccess = UpdateOrderPaymentStatus(orderId, 1, "ALIPAY-DEV-" & orderId)
        If Not updateSuccess Then
            result = False
        End If
    Else
        On Error Resume Next
        result = True
        updateSuccess = UpdateOrderPaymentStatus(orderId, 1, "ALIPAY-API-" & orderId)
        If Not updateSuccess Then
            result = False
        End If
        On Error Goto 0
    End If
    
    ProcessAlipay = result
End Function

' Process PayPal
Function ProcessPayPal(orderId)
    Dim result, updateSuccess
    result = False
    
    If Not PAYPAL_ENABLED Then
        Session("ErrorMessage") = "PayPal not enabled"
        ProcessPayPal = result
        Exit Function
    End If
    
    Dim rsOrder
    Set rsOrder = ExecuteQuery("SELECT OrderNo, TotalAmount, Notes FROM Orders WHERE OrderID = " & orderId)
    If rsOrder Is Nothing Or rsOrder.EOF Then
        Session("ErrorMessage") = "Order not found"
        ProcessPayPal = result
        Exit Function
    End If
    
    Dim orderNo, orderAmount, orderDesc
    orderNo = rsOrder("OrderNo")
    orderAmount = rsOrder("TotalAmount")
    orderDesc = rsOrder("Notes")
    rsOrder.Close
    Set rsOrder = Nothing
    
    If PAYPAL_CLIENT_ID = "" Or PAYPAL_SECRET = "" Then
        result = True
        updateSuccess = UpdateOrderPaymentStatus(orderId, 1, "PAYPAL-DEV-" & orderId)
        If Not updateSuccess Then
            result = False
        End If
    Else
        On Error Resume Next
        result = True
        updateSuccess = UpdateOrderPaymentStatus(orderId, 1, "PAYPAL-API-" & orderId)
        If Not updateSuccess Then
            result = False
        End If
        On Error Goto 0
    End If
    
    ProcessPayPal = result
End Function

' Process Cash on Delivery
Function ProcessCashOnDelivery(orderId)
    Dim result, updateSuccess
    result = False
    
    If Not COD_ENABLED Then
        Session("ErrorMessage") = "COD not enabled"
        ProcessCashOnDelivery = result
        Exit Function
    End If
    
    updateSuccess = UpdateOrderPaymentStatus(orderId, 1, "COD-" & orderId)
    If updateSuccess Then
        result = True
    End If
    
    ProcessCashOnDelivery = result
End Function

' Update order payment status
Function UpdateOrderPaymentStatus(orderId, status, transactionId)
    Dim sql, updateResult, currentNotesSql, rsCurrentNotes, currentNotes
    
    currentNotesSql = "SELECT Notes FROM Orders WHERE OrderID = " & orderId
    Set rsCurrentNotes = ExecuteQuery(currentNotesSql)
    
    If Not rsCurrentNotes Is Nothing Then
        If Not rsCurrentNotes.EOF Then
            currentNotes = rsCurrentNotes("Notes")
        End If
        rsCurrentNotes.Close
        Set rsCurrentNotes = Nothing
    End If
    
    If InStr(currentNotes, "Transaction:") = 0 Then
        currentNotes = currentNotes & " | Transaction: " & SafeSQL(transactionId)
    End If
    
    Dim statusText
    Select Case status
        Case 1
            statusText = "Paid"
        Case 2
            statusText = "Failed"
        Case 3
            statusText = "Refunded"
        Case Else
            statusText = "Pending"
    End Select
    
    sql = "UPDATE Orders SET Status = '" & statusText & "', Notes = '" & _
          SafeSQL(currentNotes) & "', UpdatedAt = GETDATE() WHERE OrderID = " & orderId
    
    On Error Resume Next
    updateResult = ExecuteNonQuery(sql)
    If Err.Number <> 0 Then
        Session("LastPaymentError") = "Update failed: " & Err.Description
        UpdateOrderPaymentStatus = False
    ElseIf Not updateResult Then
        Session("LastPaymentError") = "Update failed: DBError"
        UpdateOrderPaymentStatus = False
    Else
        Session("LastPaymentError") = ""
        UpdateOrderPaymentStatus = True
        ' 支付成功，自动创建生产订单
        If statusText = "Paid" Then
            Call AutoCreateProductionOrder(orderId)
        End If
    End If
    On Error Goto 0
End Function

' Verify payment callback
Function VerifyPaymentCallback(paymentMethod, callbackData)
    Dim result
    result = False
    
    Select Case paymentMethod
        Case PAYMENT_METHOD_WECHAT
            result = VerifyWeChatCallback(callbackData)
        Case PAYMENT_METHOD_ALIPAY
            result = VerifyAlipayCallback(callbackData)
        Case PAYMENT_METHOD_PAYPAL
            result = VerifyPayPalCallback(callbackData)
    End Select
    
    VerifyPaymentCallback = result
End Function

' Verify WeChat callback
Function VerifyWeChatCallback(callbackData)
    VerifyWeChatCallback = True
End Function

' Verify Alipay callback
Function VerifyAlipayCallback(callbackData)
    VerifyAlipayCallback = True
End Function

' Verify PayPal callback
Function VerifyPayPalCallback(callbackData)
    VerifyPayPalCallback = True
End Function

' Auto create production order after payment success
Function AutoCreateProductionOrder(orderId)
    On Error Resume Next
    
    Dim sqlCheck, rsCheck, existingCount
    Dim sqlTotal, rsTotal
    Dim sqlDetails, rsDetails
    Dim totalBottles, bottleIndex
    Dim workOrderPrefix, workOrderNo
    Dim detailId, qty, productId
    Dim sqlRecipe, rsRecipe
    Dim recipeId, recipeName, recipeCode, fullRecipeName
    Dim hasRecipe
    Dim sqlInsert, insertResult
    Dim sqlGetId, rsGetId, newProdId
    Dim sqlLog, logResult
    Dim i
    
    existingCount = 0
    totalBottles = 0
    bottleIndex = 0
    newProdId = 0
    hasRecipe = False
    
    ' 幂等检查：检查是否已存在该订单的生产订单
    sqlCheck = "SELECT COUNT(*) AS Cnt FROM ProductionOrders WHERE OrderID = " & CLng(orderId)
    Set rsCheck = ExecuteQuery(sqlCheck)
    
    If rsCheck Is Nothing Then
        AutoCreateProductionOrder = False
        Exit Function
    End If
    
    If Not rsCheck.EOF Then
        existingCount = rsCheck("Cnt")
    End If
    
    rsCheck.Close
    Set rsCheck = Nothing
    
    ' 如果已存在生产订单，则退出（幂等）
    If existingCount > 0 Then
        AutoCreateProductionOrder = True
        Exit Function
    End If
    
    ' 查询总瓶数
    sqlTotal = "SELECT SUM(Quantity) AS TotalQty FROM OrderDetails WHERE OrderID = " & CLng(orderId)
    Set rsTotal = ExecuteQuery(sqlTotal)
    
    If Not rsTotal Is Nothing Then
        If Not rsTotal.EOF Then
            If Not IsNull(rsTotal("TotalQty")) Then
                totalBottles = CLng(rsTotal("TotalQty"))
            End If
        End If
        rsTotal.Close
        Set rsTotal = Nothing
    End If
    
    ' 如果没有订单明细，直接返回成功
    If totalBottles = 0 Then
        AutoCreateProductionOrder = True
        Exit Function
    End If
    
    ' 查询订单明细
    sqlDetails = "SELECT DetailID, Quantity, ProductID FROM OrderDetails WHERE OrderID = " & CLng(orderId)
    Set rsDetails = ExecuteQuery(sqlDetails)
    
    If rsDetails Is Nothing Then
        AutoCreateProductionOrder = False
        Exit Function
    End If
    
    ' 生成工单编号前缀
    workOrderPrefix = "WO-" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "-"
    
    ' 遍历每个订单明细，为每瓶生成工单
    Do While Not rsDetails.EOF
        detailId = rsDetails("DetailID")
        qty = CLng(rsDetails("Quantity"))
        productId = rsDetails("ProductID")
        
        ' 获取产品配方信息
        hasRecipe = False
        recipeId = 0
        recipeName = ""
        recipeCode = ""
        
        sqlRecipe = "SELECT p.RecipeID, r.RecipeName, r.RecipeCode FROM (Products p LEFT JOIN Recipes r ON p.RecipeID = r.RecipeID) WHERE p.ProductID = " & CLng(productId)
        Set rsRecipe = ExecuteQuery(sqlRecipe)
        
        If Not rsRecipe Is Nothing Then
            If Not rsRecipe.EOF Then
                If Not IsNull(rsRecipe("RecipeID")) Then
                    hasRecipe = True
                    recipeId = rsRecipe("RecipeID")
                    If Not IsNull(rsRecipe("RecipeName")) Then
                        recipeName = rsRecipe("RecipeName")
                    End If
                    If Not IsNull(rsRecipe("RecipeCode")) Then
                        recipeCode = rsRecipe("RecipeCode")
                    End If
                End If
            End If
            rsRecipe.Close
            Set rsRecipe = Nothing
        End If
        
        ' 循环 Quantity 次，为每瓶创建工单
        For i = 1 To qty
            bottleIndex = bottleIndex + 1
            workOrderNo = workOrderPrefix & Right("00" & bottleIndex, 3)
            
            ' 构建 INSERT 语句
            If hasRecipe Then
                fullRecipeName = "[" & recipeCode & "] " & recipeName
                sqlInsert = "INSERT INTO ProductionOrders (OrderID, DetailID, WorkOrderNo, BottleIndex, TotalBottles, Status, Priority, RecipeID, RecipeName, CreatedAt, UpdatedAt) VALUES (" & _
                            CLng(orderId) & ", " & CLng(detailId) & ", '" & SafeSQL(workOrderNo) & "', " & bottleIndex & ", " & totalBottles & ", '待排产', 0, " & CLng(recipeId) & ", '" & SafeSQL(fullRecipeName) & "', GETDATE(), GETDATE())"
            Else
                sqlInsert = "INSERT INTO ProductionOrders (OrderID, DetailID, WorkOrderNo, BottleIndex, TotalBottles, Status, Priority, CreatedAt, UpdatedAt) VALUES (" & _
                            CLng(orderId) & ", " & CLng(detailId) & ", '" & SafeSQL(workOrderNo) & "', " & bottleIndex & ", " & totalBottles & ", '待排产', 0, GETDATE(), GETDATE())"
            End If
            
            insertResult = ExecuteNonQuery(sqlInsert)
            
            If Not insertResult Then
                rsDetails.Close
                Set rsDetails = Nothing
                AutoCreateProductionOrder = False
                Exit Function
            End If
            
            ' 获取新创建的 ProductionID（Access 不支持 LAST_INSERT_ID，使用 MAX）
            sqlGetId = "SELECT MAX(ProductionID) AS NewId FROM ProductionOrders WHERE OrderID = " & CLng(orderId)
            Set rsGetId = ExecuteQuery(sqlGetId)
            
            If rsGetId Is Nothing Then
                rsDetails.Close
                Set rsDetails = Nothing
                AutoCreateProductionOrder = False
                Exit Function
            End If
            
            newProdId = 0
            If Not rsGetId.EOF Then
                newProdId = rsGetId("NewId")
            End If
            
            rsGetId.Close
            Set rsGetId = Nothing
            
            If newProdId = 0 Then
                rsDetails.Close
                Set rsDetails = Nothing
                AutoCreateProductionOrder = False
                Exit Function
            End If
            
            ' 插入生产订单日志
            sqlLog = "INSERT INTO ProductionLogs (ProductionID, Status, Notes, CreatedBy, CreatedAt) VALUES (" & _
                     CLng(newProdId) & ", '待排产', '订单支付成功，系统自动创建生产工单 (第" & bottleIndex & "瓶/共" & totalBottles & "瓶)', 'SYSTEM', GETDATE())"
            
            logResult = ExecuteNonQuery(sqlLog)
        Next
        
        rsDetails.MoveNext
    Loop
    
    rsDetails.Close
    Set rsDetails = Nothing
    
    AutoCreateProductionOrder = True
    
    On Error Goto 0
End Function
%>