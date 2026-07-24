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

' V8：关联成本中心
On Error Resume Next
conn.Execute "SELECT CenterID FROM ExpenseRecords WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE ExpenseRecords ADD CenterID INT NULL"
conn.Execute "SELECT TOP 1 1 FROM CostCenters"
If Err.Number <> 0 Then Err.Clear : conn.Execute "CREATE TABLE CostCenters (CenterID INT IDENTITY(1,1) PRIMARY KEY, CenterCode NVARCHAR(50) NOT NULL UNIQUE, CenterName NVARCHAR(200) NOT NULL, CenterType NVARCHAR(50) DEFAULT 'Department', IsActive BIT DEFAULT 1)"
' V20：确保 Products 有 Weight/Volume 列（运费按规则计算所需）
If Err.Number <> 0 Then Err.Clear
conn.Execute "SELECT Weight FROM Products WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE Products ADD Weight DECIMAL(9,3) NULL"
If Err.Number <> 0 Then Err.Clear
conn.Execute "SELECT Volume FROM Products WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE Products ADD Volume DECIMAL(12,3) NULL"
On Error GoTo 0

' ========== 权限检查 ==========
Dim canEdit
canEdit = False
If Session("AdminRoleCode") = "FIN_MANAGER" OR Session("AdminRoleCode") = "SUPER_ADMIN" Then
    canEdit = True
End If

' ========== 安全保存配置项的函数（UPSERT模式）==========
Sub SaveConfig(key, value)
    Dim checkSQL, updateSQL, insertSQL
    checkSQL = "SELECT COUNT(*) FROM SiteSettings WHERE SettingKey = '" & SafeSQL(key) & "'"
    If GetScalar(checkSQL) > 0 Then
        updateSQL = "UPDATE SiteSettings SET SettingValue = '" & SafeSQL(value) & "' WHERE SettingKey = '" & SafeSQL(key) & "'"
        ExecuteNonQuery updateSQL
    Else
        insertSQL = "INSERT INTO SiteSettings (SettingKey, SettingValue) VALUES ('" & SafeSQL(key) & "', '" & SafeSQL(value) & "')"
        ExecuteNonQuery insertSQL
    End If
End Sub

' ========== 获取配置值 ==========
Function GetConfig(key)
    Dim val
    val = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey = '" & key & "'")
    If IsNull(val) Then val = ""
    GetConfig = val
End Function

Function GetConfigWithDefault(key, defaultValue)
    Dim val
    val = GetConfig(key)
    If val = "" Then val = defaultValue
    GetConfigWithDefault = val
End Function

' ========== V21: 分摊后回写订单费用金额并重算利润（集合式，幂等）==========
' 将 ExpenseRecords 按订单汇总写入 Orders.ExpenseAmount，并令
' ProfitAmount = TotalAmount - CostAmount - ShippingFee - ExpenseAmount（下限0），与成本引擎口径一致
Sub SyncAllOrderExpenseProfit()
    On Error Resume Next
    ' 回写各订单费用合计
    ExecuteNonQuery "UPDATE o SET o.ExpenseAmount = ISNULL(e.ExpSum,0) FROM Orders o " & _
        "INNER JOIN (SELECT OrderID, SUM(Amount) AS ExpSum FROM ExpenseRecords WHERE OrderID IS NOT NULL GROUP BY OrderID) e ON o.OrderID = e.OrderID"
    ' 重算利润（已含费用）
    ExecuteNonQuery "UPDATE Orders SET ProfitAmount = CASE WHEN (ISNULL(TotalAmount,0) - ISNULL(CostAmount,0) - ISNULL(ShippingFee,0) - ISNULL(ExpenseAmount,0)) < 0 " & _
        "THEN 0 ELSE (ISNULL(TotalAmount,0) - ISNULL(CostAmount,0) - ISNULL(ShippingFee,0) - ISNULL(ExpenseAmount,0)) END WHERE ISNULL(ExpenseAmount,0) > 0"
    On Error GoTo 0
End Sub

' ========== 处理表单提交 ==========
Dim action, msg, errMsg, activeTab
action = Request.Form("action")
msg = ""
errMsg = ""
activeTab = ""

If action = "save_config" AND canEdit Then
    If Not ValidateCSRFToken() Then
        errMsg = "安全验证失败"
    Else
        ' 运费分摊配置
        SaveConfig "ShippingAllocationMethod", Request.Form("shippingMethod")
        SaveConfig "ShippingFirstWeight", Request.Form("firstWeight")
        SaveConfig "ShippingFirstPrice", Request.Form("firstPrice")
        SaveConfig "ShippingContinueWeight", Request.Form("continueWeight")
        SaveConfig "ShippingContinuePrice", Request.Form("continuePrice")
        SaveConfig "ShippingVolumeFactor", Request.Form("volumeFactor")
        
        ' 平台费率配置
        SaveConfig "PlatformFeeAlipay", Request.Form("platformFeeAlipay")
        SaveConfig "PlatformFeeWechat", Request.Form("platformFeeWechat")
        SaveConfig "PlatformFeeStripe", Request.Form("platformFeeStripe")
        SaveConfig "PlatformFeePayPal", Request.Form("platformFeePayPal")
        SaveConfig "PlatformFeeUnionPay", Request.Form("platformFeeUnionPay")
        SaveConfig "PlatformFixedFee", Request.Form("platformFixedFee")
        
        Call LogAdminAction("配置费用分摊规则", "finance", "SiteSettings", "", "费用分摊规则配置更新")
        msg = "配置保存成功"
        activeTab = Request.Form("configScope")
    End If
End If

' ========== 运费分摊处理 ==========
If action = "allocate_shipping" AND canEdit Then
    If Not ValidateCSRFToken() Then
        errMsg = "安全验证失败"
    Else
        Dim orderId, startDate, endDate, shippingMethod
        orderId = Request.Form("orderId")
        startDate = Request.Form("startDate")
        endDate = Request.Form("endDate")
        shippingMethod = Request.Form("shippingMethod")
        
        ' 按运费规则自动计算并分摊
        Call AllocateShipping(orderId, startDate, endDate, shippingMethod)
        Call SyncAllOrderExpenseProfit()   ' V21: 回写费用并重算利润
        msg = "运费分摊完成"
        activeTab = "shipping"
    End If
End If

' ========== 平台费用分摊处理 ==========
If action = "allocate_platform_fee" AND canEdit Then
    If Not ValidateCSRFToken() Then
        errMsg = "安全验证失败"
    Else
        Dim pfStartDate, pfEndDate
        pfStartDate = Request.Form("startDate")
        pfEndDate = Request.Form("endDate")
        
        Call AllocatePlatformFee(pfStartDate, pfEndDate)
        Call SyncAllOrderExpenseProfit()   ' V21: 回写费用并重算利润
        msg = "平台费用分摊完成"
        activeTab = "platform"
    End If
End If

' ========== 推广费分摊处理 ==========
If action = "allocate_promotion" AND canEdit Then
    If Not ValidateCSRFToken() Then
        errMsg = "安全验证失败"
    Else
        Dim promoStartDate, promoEndDate, promoTotalAmount, promoGMV
        promoStartDate = Request.Form("startDate")
        promoEndDate = Request.Form("endDate")
        promoTotalAmount = CDbl("0" & Request.Form("totalAmount"))
        promoGMV = CDbl("0" & Request.Form("gmvAmount"))
        
        If promoTotalAmount <= 0 Then
            errMsg = "推广费用必须大于0"
        ElseIf promoGMV <= 0 Then
            errMsg = "有效成交额必须大于0"
        Else
            Call AllocatePromotion(promoStartDate, promoEndDate, promoTotalAmount, promoGMV)
            Call SyncAllOrderExpenseProfit()   ' V21: 回写费用并重算利润
            msg = "推广费分摊完成"
            activeTab = "promotion"
        End If
    End If
End If

' ========== 手动调整分摊金额 ==========
If action = "adjust_expense" AND canEdit Then
    If Not ValidateCSRFToken() Then
        errMsg = "安全验证失败"
    Else
        Dim expenseId, newAmount, adjustReason
        expenseId = Request.Form("expenseId")
        newAmount = CDbl("0" & Request.Form("newAmount"))
        adjustReason = Request.Form("adjustReason")
        
        If expenseId <> "" AND newAmount >= 0 Then
            Dim adjSQL
            adjSQL = "UPDATE ExpenseRecords SET Amount = " & newAmount & ", " & _
                     "ExpenseName = ExpenseName + ' [调整:" & SafeSQL(adjustReason) & "]' " & _
                     "WHERE ExpenseID = " & CLng(expenseId)
            If ExecuteNonQuery(adjSQL) Then
                msg = "分摊金额调整成功"
                activeTab = "results"
            Else
                errMsg = "调整失败"
            End If
        End If
    End If
End If

' ========== 运费分摊子程序（V20 重做：按运费规则自动计算）==========
Sub AllocateShipping(orderId, startDate, endDate, method)
    Dim whereClause, rsOrders, sql, orderIDVal, orderCreatedAt
    Dim rsItems, totalWeight, totalVolume, totalQty, itemWeight, itemVolume, quantity
    Dim productId, weight, volume, allocAmount
    Dim firstWeight, firstPrice, continueWeight, continuePrice, volumeFactor
    Dim defaultWeight, defaultVolume
    Dim chargeableWeight, orderFreight, skuCount, i, remainingAmount
    
    ' 获取运费规则配置
    firstWeight = CDbl("0" & GetConfigWithDefault("ShippingFirstWeight", "1"))
    firstPrice = CDbl("0" & GetConfigWithDefault("ShippingFirstPrice", "10"))
    continueWeight = CDbl("0" & GetConfigWithDefault("ShippingContinueWeight", "1"))
    continuePrice = CDbl("0" & GetConfigWithDefault("ShippingContinuePrice", "5"))
    volumeFactor = CDbl("0" & GetConfigWithDefault("ShippingVolumeFactor", "5000"))
    ' 商品缺失重量/体积时的默认值
    defaultWeight = CDbl("0" & GetConfigWithDefault("ShippingDefaultUnitWeight", "0.5"))
    defaultVolume = CDbl("0" & GetConfigWithDefault("ShippingDefaultUnitVolume", "750"))
    
    ' 构建查询条件（修正日期区间：含结束当天）
    whereClause = "Status = 'Paid'"
    If orderId <> "" Then
        whereClause = whereClause & " AND OrderID = " & CLng(orderId)
    End If
    If startDate <> "" Then
        whereClause = whereClause & " AND CreatedAt >= '" & startDate & "'"
    End If
    If endDate <> "" Then
        whereClause = whereClause & " AND CreatedAt < DATEADD(day, 1, '" & endDate & "')"
    End If
    
    ' 查询符合条件的订单（加选 CreatedAt 用于账期）
    sql = "SELECT OrderID, CreatedAt FROM Orders WHERE " & whereClause
    Set rsOrders = ExecuteQuery(sql)
    
    If Not rsOrders Is Nothing Then
        Do While Not rsOrders.EOF
            orderIDVal = rsOrders("OrderID").Value
            orderCreatedAt = rsOrders("CreatedAt").Value
            
            ' 获取订单所有SKU的重量/体积/数量信息
            sql = "SELECT od.ProductID, od.Quantity, p.Weight, p.Volume, p.ProductName " & _
                  "FROM (OrderDetails AS od INNER JOIN Products AS p ON od.ProductID = p.ProductID) " & _
                  "WHERE od.OrderID = " & orderIDVal
            Set rsItems = ExecuteQuery(sql)
            
            If Not rsItems Is Nothing Then
                totalWeight = 0
                totalVolume = 0
                totalQty = 0
                
                ' 第一遍：汇总计费重量和统计信息
                chargeableWeight = 0
                Do While Not rsItems.EOF
                    quantity = rsItems("Quantity").Value
                    If IsNull(quantity) Then quantity = 1
                    weight = rsItems("Weight").Value
                    If IsNull(weight) Or CDbl(weight) <= 0 Then weight = defaultWeight
                    volume = rsItems("Volume").Value
                    If IsNull(volume) Or CDbl(volume) <= 0 Then volume = defaultVolume
                    
                    ' 计费重量 = max(实际重量, 体积重/系数)
                    itemWeight = CDbl(weight) * quantity
                    itemVolume = (CDbl(volume) * quantity) / volumeFactor
                    If itemWeight > itemVolume Then
                        chargeableWeight = chargeableWeight + itemWeight
                    Else
                        chargeableWeight = chargeableWeight + itemVolume
                    End If
                    totalWeight = totalWeight + CDbl(weight) * quantity
                    totalVolume = totalVolume + CDbl(volume) * quantity
                    totalQty = totalQty + quantity
                    rsItems.MoveNext
                Loop
                
                ' 按运费规则计算订单运费
                If chargeableWeight > 0 Then
                    If chargeableWeight <= firstWeight Then
                        orderFreight = firstPrice
                    Else
                        Dim extraWeight
                        extraWeight = chargeableWeight - firstWeight
                        ' Ceil(extraWeight / continueWeight)
                        Dim extraUnits
                        extraUnits = Int(extraWeight / continueWeight)
                        If extraWeight <> extraUnits * continueWeight And continueWeight > 0 Then
                            extraUnits = extraUnits + 1
                        End If
                        orderFreight = firstPrice + extraUnits * continuePrice
                    End If
                Else
                    orderFreight = 0
                End If
                
                ' 幂等：先清除该订单旧的运费分摊记录
                If orderFreight > 0 Then
                    ExecuteNonQuery "DELETE FROM ExpenseRecords WHERE OrderID = " & orderIDVal & " AND ExpenseType = 'Shipping'"
                End If
                
                ' 第二遍：将订单运费分摊到各SKU
                If orderFreight > 0 Then
                    rsItems.MoveFirst
                    skuCount = rsItems.RecordCount  ' 使用客户端游标的实际行数
                    remainingAmount = orderFreight
                    i = 0
                    
                    Do While Not rsItems.EOF
                        quantity = rsItems("Quantity").Value
                        If IsNull(quantity) Then quantity = 1
                        weight = rsItems("Weight").Value
                        If IsNull(weight) Or CDbl(weight) <= 0 Then weight = defaultWeight
                        volume = rsItems("Volume").Value
                        If IsNull(volume) Or CDbl(volume) <= 0 Then volume = defaultVolume
                        productId = rsItems("ProductID").Value
                        
                        ' 按选定方法计算分摊比例
                        If method = "weight" AND totalWeight > 0 Then
                            allocAmount = orderFreight * (CDbl(weight) * quantity) / totalWeight
                        ElseIf method = "volume" AND totalVolume > 0 Then
                            allocAmount = orderFreight * (CDbl(volume) * quantity) / totalVolume
                        Else
                            ' 回退为按数量/平均分摊（缺失数据时保证有结果）
                            allocAmount = orderFreight / totalQty
                        End If
                        
                        ' 尾差：最后一行承担剩余金额
                        i = i + 1
                        If i = skuCount Then
                            allocAmount = remainingAmount
                        Else
                            allocAmount = Round(allocAmount, 2)
                            remainingAmount = remainingAmount - allocAmount
                        End If
                        
                        ' 写入分摊记录，账期取订单创建月份
                        If allocAmount > 0 Then
                            Dim shipPeriod
                            shipPeriod = ""
                            If Not IsNull(orderCreatedAt) And orderCreatedAt <> "" Then
                                shipPeriod = Year(orderCreatedAt) & "-" & Right("0" & Month(orderCreatedAt), 2)
                            Else
                                shipPeriod = Year(Now) & "-" & Right("0" & Month(Now), 2)
                            End If
                            
                            Dim shipSQL
                            shipSQL = "INSERT INTO ExpenseRecords (OrderID, ProductID, ExpenseType, ExpenseName, Amount, " & _
                                      "AllocationMethod, AllocationRatio, Period, CreatedAt) VALUES (" & _
                                      orderIDVal & ", " & productId & ", 'Shipping', '运费分摊', " & allocAmount & ", '" & _
                                      SafeSQL(method) & "', " & Round(allocAmount / orderFreight, 4) & ", '" & _
                                      shipPeriod & "', GETDATE())"
                            ExecuteNonQuery shipSQL
                        End If
                        
                        rsItems.MoveNext
                    Loop
                End If
                
                rsItems.Close
                Set rsItems = Nothing
            End If
            
            rsOrders.MoveNext
        Loop
        rsOrders.Close
        Set rsOrders = Nothing
    End If
End Sub

' ========== 平台费用分摊子程序（V20 修复：数字编码匹配 + 幂等 + 账期）==========
Sub AllocatePlatformFee(startDate, endDate)
    Dim whereClause, rsOrders, sql, orderIDVal, totalAmount, paymentMethod, orderCreatedAt
    Dim rsItems, platformFeeRate, fixedFee, feeAmount, allocAmount
    Dim i, remainingFee, skuCount
    
    ' 获取各支付方式费率
    Dim feeAlipay, feeWechat, feeStripe, feePayPal, feeUnionPay, feeFixed
    feeAlipay = CDbl("0" & GetConfigWithDefault("PlatformFeeAlipay", "0.6")) / 100
    feeWechat = CDbl("0" & GetConfigWithDefault("PlatformFeeWechat", "0.6")) / 100
    feeStripe = CDbl("0" & GetConfigWithDefault("PlatformFeeStripe", "2.9")) / 100
    feePayPal = CDbl("0" & GetConfigWithDefault("PlatformFeePayPal", "4.4")) / 100
    feeUnionPay = CDbl("0" & GetConfigWithDefault("PlatformFeeUnionPay", "0.6")) / 100
    feeFixed = CDbl("0" & GetConfigWithDefault("PlatformFixedFee", "0"))
    
    ' 构建查询条件（修正日期区间：含结束当天）
    whereClause = "Status = 'Paid'"
    If startDate <> "" Then
        whereClause = whereClause & " AND CreatedAt >= '" & startDate & "'"
    End If
    If endDate <> "" Then
        whereClause = whereClause & " AND CreatedAt < DATEADD(day, 1, '" & endDate & "')"
    End If
    
    ' 查询符合条件的订单（加选 CreatedAt 用于账期）
    sql = "SELECT OrderID, TotalAmount, PaymentMethod, CreatedAt FROM Orders WHERE " & whereClause
    Set rsOrders = ExecuteQuery(sql)
    
    If Not rsOrders Is Nothing Then
        Do While Not rsOrders.EOF
            orderIDVal = rsOrders("OrderID").Value
            totalAmount = CDbl("0" & rsOrders("TotalAmount").Value)
            paymentMethod = rsOrders("PaymentMethod").Value
            orderCreatedAt = rsOrders("CreatedAt").Value
            If IsNull(paymentMethod) Then paymentMethod = ""
            
            ' V20: 按数字支付编码匹配费率（优先），回退英文子串匹配（旧数据兼容）
            platformFeeRate = 0
            Dim pmTrimmed
            pmTrimmed = Trim(CStr(paymentMethod))
            Select Case pmTrimmed
                Case "1": platformFeeRate = feeWechat      ' 微信支付
                Case "2": platformFeeRate = feeAlipay      ' 支付宝
                Case "3": platformFeeRate = feePayPal      ' PayPal
                Case "4": platformFeeRate = 0              ' 货到付款（默认无平台扣点）
                Case Else
                    ' 旧数据兼容：英文子串匹配
                    If InStr(LCase(pmTrimmed), "alipay") > 0 Then
                        platformFeeRate = feeAlipay
                    ElseIf InStr(LCase(pmTrimmed), "wechat") > 0 Then
                        platformFeeRate = feeWechat
                    ElseIf InStr(LCase(pmTrimmed), "stripe") > 0 Then
                        platformFeeRate = feeStripe
                    ElseIf InStr(LCase(pmTrimmed), "paypal") > 0 Then
                        platformFeeRate = feePayPal
                    ElseIf InStr(LCase(pmTrimmed), "union") > 0 Then
                        platformFeeRate = feeUnionPay
                    End If
            End Select
            
            ' 计算平台费用
            feeAmount = totalAmount * platformFeeRate + feeFixed
            
            ' 幂等：先清除该订单旧的平台费分摊记录
            If feeAmount > 0 Then
                ExecuteNonQuery "DELETE FROM ExpenseRecords WHERE OrderID = " & orderIDVal & " AND ExpenseType = 'PlatformFee'"
            End If
            
            If feeAmount > 0 Then
                ' 获取订单明细并按金额比例分摊
                sql = "SELECT od.ProductID, od.Subtotal, p.ProductName " & _
                      "FROM (OrderDetails AS od INNER JOIN Products AS p ON od.ProductID = p.ProductID) " & _
                      "WHERE od.OrderID = " & orderIDVal
                Set rsItems = ExecuteQuery(sql)
                
                If Not rsItems Is Nothing Then
                    skuCount = rsItems.RecordCount  ' 使用实际 INNER JOIN 后的行数，避免商品缺失时少分摊
                    remainingFee = feeAmount
                    i = 0
                    Do While Not rsItems.EOF
                        Dim subtotal, productId2
                        subtotal = CDbl("0" & rsItems("Subtotal").Value)
                        productId2 = rsItems("ProductID").Value
                        
                        If totalAmount > 0 Then
                            allocAmount = feeAmount * subtotal / totalAmount
                        ElseIf skuCount > 0 Then
                            allocAmount = feeAmount / skuCount
                        Else
                            allocAmount = 0
                        End If
                        
                        ' 尾差：最后一行承担剩余金额
                        i = i + 1
                        If i = skuCount Then
                            allocAmount = remainingFee
                        Else
                            allocAmount = Round(allocAmount, 2)
                            remainingFee = remainingFee - allocAmount
                        End If
                        
                        If allocAmount > 0 Then
                            ' 账期取订单创建月份
                            Dim pfPeriod
                            pfPeriod = ""
                            If Not IsNull(orderCreatedAt) And orderCreatedAt <> "" Then
                                pfPeriod = Year(orderCreatedAt) & "-" & Right("0" & Month(orderCreatedAt), 2)
                            Else
                                pfPeriod = Year(Now) & "-" & Right("0" & Month(Now), 2)
                            End If
                            
                            Dim pfSQL
                            pfSQL = "INSERT INTO ExpenseRecords (OrderID, ProductID, ExpenseType, ExpenseName, Amount, " & _
                                      "AllocationMethod, AllocationRatio, Period, CreatedAt) VALUES (" & _
                                      orderIDVal & ", " & productId2 & ", 'PlatformFee', '平台扣点', " & allocAmount & ", " & _
                                      "'PaymentMethod', " & Round(platformFeeRate, 4) & ", '" & _
                                      pfPeriod & "', GETDATE())"
                            ExecuteNonQuery pfSQL
                        End If
                        
                        rsItems.MoveNext
                    Loop
                    rsItems.Close
                    Set rsItems = Nothing
                End If
            End If
            
            rsOrders.MoveNext
        Loop
        rsOrders.Close
        Set rsOrders = Nothing
    End If
End Sub

' ========== 推广费分摊子程序（V20 修复：幂等 + 账期 + 尾差）==========
Sub AllocatePromotion(startDate, endDate, totalPromoAmount, gmvAmount)
    Dim whereClause, rsOrders, sql, orderIDVal, orderAmount, orderCreatedAt
    Dim rsItems, allocRatio, allocAmount, i, remainingAmount
    
    ' 构建查询条件（修正日期区间：含结束当天）
    whereClause = "Status = 'Paid'"
    If startDate <> "" Then
        whereClause = whereClause & " AND CreatedAt >= '" & startDate & "'"
    End If
    If endDate <> "" Then
        whereClause = whereClause & " AND CreatedAt < DATEADD(day, 1, '" & endDate & "')"
    End If
    
    ' 查询符合条件的订单（加选 CreatedAt 用于账期）
    sql = "SELECT OrderID, TotalAmount, CreatedAt FROM Orders WHERE " & whereClause
    Set rsOrders = ExecuteQuery(sql)
    
    If Not rsOrders Is Nothing Then
        remainingAmount = totalPromoAmount
        
        ' 先统计订单总数
        Dim orderCount
        orderCount = 0
        Do While Not rsOrders.EOF
            orderCount = orderCount + 1
            rsOrders.MoveNext
        Loop
        rsOrders.MoveFirst
        
        i = 0
        Do While Not rsOrders.EOF
            orderIDVal = rsOrders("OrderID").Value
            orderAmount = CDbl("0" & rsOrders("TotalAmount").Value)
            orderCreatedAt = rsOrders("CreatedAt").Value
            
            ' 计算该订单应分摊的推广费
            ' 公式：单笔承担推广费 = (总消耗 / 有效成交额) x 订单金额
            If gmvAmount > 0 Then
                allocAmount = totalPromoAmount * orderAmount / gmvAmount
            Else
                allocAmount = 0
            End If
            
            ' 最后一个订单承担剩余金额
            i = i + 1
            If i = orderCount Then
                allocAmount = remainingAmount
            Else
                allocAmount = Round(allocAmount, 2)
                remainingAmount = remainingAmount - allocAmount
            End If
            
            ' 幂等：先清除该订单旧的推广费分摊记录
            If allocAmount > 0 Then
                ExecuteNonQuery "DELETE FROM ExpenseRecords WHERE OrderID = " & orderIDVal & " AND ExpenseType = 'Promotion'"
            End If
            
            ' 获取订单SKU并按金额比例分摊到SKU
            If allocAmount > 0 Then
                ' 账期取订单创建月份
                Dim promoPeriod
                promoPeriod = ""
                If Not IsNull(orderCreatedAt) And orderCreatedAt <> "" Then
                    promoPeriod = Year(orderCreatedAt) & "-" & Right("0" & Month(orderCreatedAt), 2)
                Else
                    promoPeriod = Year(Now) & "-" & Right("0" & Month(Now), 2)
                End If
                
                sql = "SELECT od.ProductID, od.Subtotal, p.ProductName " & _
                      "FROM (OrderDetails AS od INNER JOIN Products AS p ON od.ProductID = p.ProductID) " & _
                      "WHERE od.OrderID = " & orderIDVal
                Dim rsSKU
                Set rsSKU = ExecuteQuery(sql)
                
                If Not rsSKU Is Nothing Then
                    Dim skuActualCount, j, remainingSKUAmount
                    skuActualCount = rsSKU.RecordCount  ' 使用实际行数，避免商品缺失时少分摊
                    remainingSKUAmount = allocAmount
                    j = 0
                    Do While Not rsSKU.EOF
                        Dim skuSubtotal, skuProductId, skuAlloc
                        skuSubtotal = CDbl("0" & rsSKU("Subtotal").Value)
                        skuProductId = rsSKU("ProductID").Value
                        
                        If orderAmount > 0 Then
                            skuAlloc = allocAmount * skuSubtotal / orderAmount
                        ElseIf skuActualCount > 0 Then
                            skuAlloc = allocAmount / skuActualCount
                        Else
                            skuAlloc = 0
                        End If
                        
                        ' 尾差：最后一行承担剩余金额
                        j = j + 1
                        If j = skuActualCount Then
                            skuAlloc = remainingSKUAmount
                        Else
                            skuAlloc = Round(skuAlloc, 2)
                            remainingSKUAmount = remainingSKUAmount - skuAlloc
                        End If
                        
                        If skuAlloc > 0 Then
                            Dim promoSQL
                            promoSQL = "INSERT INTO ExpenseRecords (OrderID, ProductID, ExpenseType, ExpenseName, Amount, " & _
                                      "AllocationMethod, AllocationRatio, SourceOrderID, Period, CreatedAt) VALUES (" & _
                                      orderIDVal & ", " & skuProductId & ", 'Promotion', '推广费分摊', " & skuAlloc & ", " & _
                                      "'GMVRatio', " & Round(skuSubtotal / gmvAmount, 6) & ", 0, '" & _
                                      promoPeriod & "', GETDATE())"
                            ExecuteNonQuery promoSQL
                        End If
                        
                        rsSKU.MoveNext
                    Loop
                    rsSKU.Close
                    Set rsSKU = Nothing
                End If
            End If
            
            rsOrders.MoveNext
        Loop
        rsOrders.Close
        Set rsOrders = Nothing
    End If
End Sub

Call LogAdminAction("访问费用分摊引擎", "finance", "", "", "")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>费用分摊引擎 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root {
            --bg-dark: #1a1a2e;
            --bg-card: #2d2d44;
            --bg-hover: #1e1e32;
            --border-color: rgba(255,255,255,0.06);
            --text-primary: #e0e0e0;
            --text-secondary: #b0b0b0;
            --text-muted: #888;
            --accent-primary: #00bcd4;
            --accent-secondary: #00838f;
            --success: #4CAF50;
            --warning: #ffa726;
            --danger: #f44336;
            --info: #2196F3;
        }
        
        * { box-sizing: border-box; }
        
        body {
            background: var(--bg-dark);
            color: var(--text-primary);
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            margin: 0;
            padding: 0;
        }
        
        .main-content {
            margin-left: 250px;
            padding: 30px;
            min-height: 100vh;
        }
        
        .page-header {
            margin-bottom: 30px;
            padding-bottom: 20px;
            border-bottom: 1px solid var(--border-color);
        }
        
        .page-title {
            font-size: 28px;
            font-weight: 600;
            margin: 0 0 10px 0;
            display: flex;
            align-items: center;
            gap: 15px;
        }
        
        .page-title i { color: var(--accent-primary); }
        
        .breadcrumb {
            color: var(--text-muted);
            font-size: 14px;
        }
        
        .breadcrumb a {
            color: var(--accent-primary);
            text-decoration: none;
        }
        
        .breadcrumb a:hover { text-decoration: underline; }
        
        /* 消息提示 */
        .alert {
            padding: 15px 20px;
            border-radius: 8px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .alert-success {
            background: rgba(76, 175, 80, 0.15);
            border: 1px solid var(--success);
            color: var(--success);
        }
        
        .alert-error {
            background: rgba(244, 67, 54, 0.15);
            border: 1px solid var(--danger);
            color: var(--danger);
        }
        
        .alert-warning {
            background: rgba(255, 167, 38, 0.15);
            border: 1px solid var(--warning);
            color: var(--warning);
        }
        
        /* Tab 导航 */
        .tab-nav {
            display: flex;
            gap: 5px;
            margin-bottom: 25px;
            background: var(--bg-card);
            padding: 5px;
            border-radius: 12px;
            border: 1px solid var(--border-color);
        }
        
        .tab-btn {
            flex: 1;
            padding: 15px 20px;
            background: transparent;
            border: none;
            color: var(--text-secondary);
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            border-radius: 8px;
            transition: all 0.3s;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
        }
        
        .tab-btn:hover {
            background: var(--bg-hover);
            color: var(--text-primary);
        }
        
        .tab-btn.active {
            background: linear-gradient(135deg, var(--accent-primary), var(--accent-secondary));
            color: white;
        }
        
        /* Tab 内容 */
        .tab-content {
            display: none;
        }
        
        .tab-content.active {
            display: block;
            animation: fadeIn 0.3s ease;
        }
        
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        /* 卡片样式 */
        .card {
            background: var(--bg-card);
            border-radius: 12px;
            border: 1px solid var(--border-color);
            margin-bottom: 25px;
            overflow: hidden;
        }
        
        .card-header {
            padding: 20px 25px;
            border-bottom: 1px solid var(--border-color);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        
        .card-title {
            font-size: 18px;
            font-weight: 600;
            margin: 0;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .card-body {
            padding: 25px;
        }
        
        /* 表单样式 */
        .form-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
        }
        
        .form-group {
            margin-bottom: 20px;
        }
        
        .form-group label {
            display: block;
            margin-bottom: 8px;
            color: var(--text-secondary);
            font-size: 14px;
            font-weight: 500;
        }
        
        .form-group input,
        .form-group select,
        .form-group textarea {
            width: 100%;
            padding: 12px 15px;
            background: var(--bg-dark);
            border: 2px solid var(--border-color);
            border-radius: 8px;
            color: var(--text-primary);
            font-size: 14px;
            transition: border-color 0.3s;
        }
        
        .form-group input:focus,
        .form-group select:focus,
        .form-group textarea:focus {
            outline: none;
            border-color: var(--accent-primary);
        }
        
        .form-group input:read-only,
        .form-group select:disabled {
            background: var(--bg-hover);
            cursor: not-allowed;
        }
        
        .form-group .help-text {
            font-size: 12px;
            color: var(--text-muted);
            margin-top: 5px;
        }
        

        /* 表格样式 */
        .data-table {
            width: 100%;
            border-collapse: collapse;
        }
        
        .data-table th,
        .data-table td {
            padding: 15px;
            text-align: left;
            border-bottom: 1px solid var(--border-color);
        }
        
        .data-table th {
            background: var(--bg-dark);
            color: var(--text-secondary);
            font-weight: 600;
            font-size: 13px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .data-table tr:hover {
            background: var(--bg-hover);
        }
        
        .data-table td {
            color: var(--text-primary);
            font-size: 14px;
        }
        
        /* 统计卡片 */
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 25px;
        }
        
        .stat-card {
            background: var(--bg-card);
            border-radius: 12px;
            padding: 20px;
            border: 1px solid var(--border-color);
        }
        
        .stat-label {
            font-size: 13px;
            color: var(--text-muted);
            margin-bottom: 8px;
        }
        
        .stat-value {
            font-size: 24px;
            font-weight: 700;
            color: var(--text-primary);
        }
        
        .stat-value.shipping { color: #4CAF50; }
        .stat-value.platform { color: #2196F3; }
        .stat-value.promotion { color: #ffa726; }
        .stat-value.total { color: #9c27b0; }
        
        /* 配置区域 */
        .config-section {
            background: linear-gradient(135deg, rgba(102, 126, 234, 0.1), rgba(118, 75, 162, 0.1));
            border: 1px solid var(--accent-primary);
            border-radius: 12px;
            padding: 25px;
            margin-bottom: 25px;
        }
        
        .config-section h3 {
            margin: 0 0 20px 0;
            color: var(--accent-primary);
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        /* 费率配置网格 */
        .fee-config-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
        }
        
        .fee-config-item {
            background: var(--bg-dark);
            padding: 15px;
            border-radius: 8px;
            border: 1px solid var(--border-color);
        }
        
        .fee-config-item label {
            display: block;
            font-size: 13px;
            color: var(--text-secondary);
            margin-bottom: 8px;
        }
        
        .fee-config-item input {
            width: 100%;
            padding: 10px;
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 6px;
            color: var(--text-primary);
        }
        
        /* 只读遮罩 */
        .readonly-mask {
            position: relative;
        }
        
        .readonly-mask::after {
            content: "只读权限";
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(26, 26, 46, 0.85);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 16px;
            color: var(--text-muted);
            border-radius: 12px;
            pointer-events: none;
        }
        
        /* 进度提示 */
        .progress-overlay {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0, 0, 0, 0.8);
            z-index: 9999;
            align-items: center;
            justify-content: center;
        }
        
        .progress-overlay.active {
            display: flex;
        }
        
        .progress-box {
            background: var(--bg-card);
            padding: 40px;
            border-radius: 16px;
            text-align: center;
            border: 1px solid var(--border-color);
        }
        
        .progress-spinner {
            width: 60px;
            height: 60px;
            border: 4px solid var(--border-color);
            border-top-color: var(--accent-primary);
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin: 0 auto 20px;
        }
        
        @keyframes spin {
            to { transform: rotate(360deg); }
        }
        
        /* 分页 */
        .pagination {
            display: flex;
            justify-content: center;
            gap: 5px;
            margin-top: 20px;
        }
        
        .page-btn {
            padding: 8px 16px;
            background: var(--bg-dark);
            border: 1px solid var(--border-color);
            color: var(--text-secondary);
            border-radius: 6px;
            cursor: pointer;
            transition: all 0.3s;
        }
        
        .page-btn:hover {
            background: var(--bg-hover);
            color: var(--text-primary);
        }
        
        .page-btn.active {
            background: var(--accent-primary);
            color: white;
            border-color: var(--accent-primary);
        }
        
        /* 响应式 */
        @media (max-width: 768px) {
            .main-content {
                margin-left: 0;
                padding: 15px;
            }
            
            .tab-nav {
                flex-wrap: wrap;
            }
            
            .tab-btn {
                flex: 1 1 45%;
            }
            
            .form-grid {
                grid-template-columns: 1fr;
            }
        }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-calculator"></i> 费用分摊引擎</h2>
            <div class="breadcrumb">
                <a href="index.asp">财务中心</a> / <span>费用分摊</span>
            </div>
        </div>
        
        <% If msg <> "" Then %>
        <div class="alert alert-success">
            <i class="fas fa-check-circle"></i> <%= Server.HTMLEncode(msg) %>
        </div>
        <% End If %>
        
        <% If errMsg <> "" Then %>
        <div class="alert alert-error">
            <i class="fas fa-times-circle"></i> <%= Server.HTMLEncode(errMsg) %>
        </div>
        <% End If %>
        
        <% If Not canEdit Then %>
        <div class="alert alert-warning">
            <i class="fas fa-lock"></i> 您当前为只读权限，仅可查看分摊结果，无法执行分摊操作
        </div>
        <% End If %>
        
        <!-- Tab 导航 -->
        <%
        Dim tabShippingClass, tabPlatformClass, tabPromoClass, tabResultsClass
        If activeTab = "" Then activeTab = "shipping"
        If activeTab = "shipping" Then tabShippingClass = " active" Else tabShippingClass = ""
        If activeTab = "platform" Then tabPlatformClass = " active" Else tabPlatformClass = ""
        If activeTab = "promotion" Then tabPromoClass = " active" Else tabPromoClass = ""
        If activeTab = "results" Then tabResultsClass = " active" Else tabResultsClass = ""
        %>
        <div class="tab-nav">
            <button class="tab-btn<%= tabShippingClass %>" onclick="switchTab('shipping', event)">
                <i class="fas fa-truck"></i> 运费分摊
            </button>
            <button class="tab-btn<%= tabPlatformClass %>" onclick="switchTab('platform', event)">
                <i class="fas fa-percentage"></i> 平台扣点
            </button>
            <button class="tab-btn<%= tabPromoClass %>" onclick="switchTab('promotion', event)">
                <i class="fas fa-bullhorn"></i> 推广费分摊
            </button>
            <button class="tab-btn<%= tabResultsClass %>" onclick="switchTab('results', event)">
                <i class="fas fa-list-alt"></i> 分摊结果
            </button>
        </div>
        
        <!-- Tab 1: 运费分摊 -->
        <div id="tab-shipping" class="tab-content<%= tabShippingClass %>">
            <div class="config-section">
                <h3><i class="fas fa-cog"></i> 运费分摊规则配置</h3>
                <form method="post" action="expense_allocation.asp" <%= IIf(canEdit, "", "class='readonly-mask'") %>>
                    <%= GetCSRFTokenField() %>
                    <input type="hidden" name="action" value="save_config">
                    <input type="hidden" name="configScope" value="shipping">
                    
                    <div class="form-grid">
                        <div class="form-group">
                            <label>分摊方式</label>
                            <select name="shippingMethod">
                                <option value="weight" <%= IIf(GetConfigWithDefault("ShippingAllocationMethod", "weight")="weight", "selected", "") %>>按重量分摊</option>
                                <option value="volume" <%= IIf(GetConfigWithDefault("ShippingAllocationMethod", "weight")="volume", "selected", "") %>>按体积分摊</option>
                                <option value="equal" <%= IIf(GetConfigWithDefault("ShippingAllocationMethod", "weight")="equal", "selected", "") %>>平均分摊</option>
                            </select>
                        </div>
                        <div class="form-group">
                            <label>首重 (kg)</label>
                            <input type="number" name="firstWeight" value="<%= GetConfigWithDefault("ShippingFirstWeight", "1") %>" step="0.1" min="0">
                        </div>
                        <div class="form-group">
                            <label>首重价格 (元)</label>
                            <input type="number" name="firstPrice" value="<%= GetConfigWithDefault("ShippingFirstPrice", "10") %>" step="0.01" min="0">
                        </div>
                        <div class="form-group">
                            <label>续重 (kg)</label>
                            <input type="number" name="continueWeight" value="<%= GetConfigWithDefault("ShippingContinueWeight", "1") %>" step="0.1" min="0">
                        </div>
                        <div class="form-group">
                            <label>续重价格 (元)</label>
                            <input type="number" name="continuePrice" value="<%= GetConfigWithDefault("ShippingContinuePrice", "5") %>" step="0.01" min="0">
                        </div>
                        <div class="form-group">
                            <label>体积重系数</label>
                            <input type="number" name="volumeFactor" value="<%= GetConfigWithDefault("ShippingVolumeFactor", "5000") %>" step="100" min="1000">
                            <div class="help-text">体积重 = 长x宽x高 / 系数</div>
                        </div>
                    </div>
                    
                    <% If canEdit Then %>
                    <div style="margin-top: 20px; text-align: right;">
                        <button type="submit" class="btn btn-primary">
                            <i class="fas fa-save"></i> 保存配置
                        </button>
                    </div>
                    <% End If %>
                </form>
            </div>
            
            <div class="card">
                <div class="card-header">
                    <h3 class="card-title"><i class="fas fa-truck-loading"></i> 执行运费分摊</h3>
                </div>
                <div class="card-body">
                    <form method="post" action="expense_allocation.asp" onsubmit="showProgress('正在分摊运费...')" <%= IIf(canEdit, "", "class='readonly-mask'") %>>
                        <%= GetCSRFTokenField() %>
                        <input type="hidden" name="action" value="allocate_shipping">
                        
                        <div class="form-grid">
                            <div class="form-group">
                                <label>订单号（可选）</label>
                                <input type="text" name="orderId" placeholder="输入订单ID，留空则按时间范围">
                            </div>
                            <div class="form-group">
                                <label>开始日期</label>
                                <input type="date" name="startDate">
                            </div>
                            <div class="form-group">
                                <label>结束日期</label>
                                <input type="date" name="endDate">
                            </div>
                        </div>
                        
                        <% If canEdit Then %>
                        <div style="margin-top: 20px;">
                            <button type="submit" class="btn btn-success">
                                <i class="fas fa-play"></i> 开始分摊
                            </button>
                        </div>
                        <% End If %>
                    </form>
                </div>
            </div>
        </div>
        
        <!-- Tab 2: 平台扣点 -->
        <div id="tab-platform" class="tab-content<%= tabPlatformClass %>">
            <div class="config-section">
                <h3><i class="fas fa-cog"></i> 平台费率配置</h3>
                <form method="post" action="expense_allocation.asp" <%= IIf(canEdit, "", "class='readonly-mask'") %>>
                    <%= GetCSRFTokenField() %>
                    <input type="hidden" name="action" value="save_config">
                    <input type="hidden" name="configScope" value="platform">
                    
                    <div class="fee-config-grid">
                        <div class="fee-config-item">
                            <label><i class="fab fa-alipay" style="color: #1677ff;"></i> 支付宝费率 (%)</label>
                            <input type="number" name="platformFeeAlipay" value="<%= GetConfigWithDefault("PlatformFeeAlipay", "0.6") %>" step="0.01" min="0">
                        </div>
                        <div class="fee-config-item">
                            <label><i class="fab fa-weixin" style="color: #07c160;"></i> 微信费率 (%)</label>
                            <input type="number" name="platformFeeWechat" value="<%= GetConfigWithDefault("PlatformFeeWechat", "0.6") %>" step="0.01" min="0">
                        </div>
                        <div class="fee-config-item">
                            <label><i class="fab fa-stripe" style="color: #635bff;"></i> Stripe费率 (%)</label>
                            <input type="number" name="platformFeeStripe" value="<%= GetConfigWithDefault("PlatformFeeStripe", "2.9") %>" step="0.01" min="0">
                        </div>
                        <div class="fee-config-item">
                            <label><i class="fab fa-paypal" style="color: #003087;"></i> PayPal费率 (%)</label>
                            <input type="number" name="platformFeePayPal" value="<%= GetConfigWithDefault("PlatformFeePayPal", "4.4") %>" step="0.01" min="0">
                        </div>
                        <div class="fee-config-item">
                            <label><i class="fas fa-credit-card" style="color: #c00;"></i> 银联费率 (%)</label>
                            <input type="number" name="platformFeeUnionPay" value="<%= GetConfigWithDefault("PlatformFeeUnionPay", "0.6") %>" step="0.01" min="0">
                        </div>
                        <div class="fee-config-item">
                            <label><i class="fas fa-dollar-sign"></i> 固定技术费 (元)</label>
                            <input type="number" name="platformFixedFee" value="<%= GetConfigWithDefault("PlatformFixedFee", "0") %>" step="0.01" min="0">
                        </div>
                    </div>
                    
                    <% If canEdit Then %>
                    <div style="margin-top: 20px; text-align: right;">
                        <button type="submit" class="btn btn-primary">
                            <i class="fas fa-save"></i> 保存配置
                        </button>
                    </div>
                    <% End If %>
                </form>
            </div>
            
            <div class="card">
                <div class="card-header">
                    <h3 class="card-title"><i class="fas fa-calculator"></i> 执行平台费用分摊</h3>
                </div>
                <div class="card-body">
                    <form method="post" action="expense_allocation.asp" onsubmit="showProgress('正在计算平台费用...')" <%= IIf(canEdit, "", "class='readonly-mask'") %>>
                        <%= GetCSRFTokenField() %>
                        <input type="hidden" name="action" value="allocate_platform_fee">
                        
                        <div class="form-grid">
                            <div class="form-group">
                                <label>开始日期</label>
                                <input type="date" name="startDate">
                            </div>
                            <div class="form-group">
                                <label>结束日期</label>
                                <input type="date" name="endDate">
                            </div>
                        </div>
                        
                        <div class="alert alert-warning" style="margin-top: 15px;">
                            <i class="fas fa-info-circle"></i> 
                            系统将自动根据订单的支付方式匹配对应费率，按公式计算：单笔承担费用 = 实付金额 x 费率 + 固定技术费
                        </div>
                        
                        <% If canEdit Then %>
                        <div style="margin-top: 20px;">
                            <button type="submit" class="btn btn-success">
                                <i class="fas fa-play"></i> 开始计算
                            </button>
                        </div>
                        <% End If %>
                    </form>
                </div>
            </div>
        </div>
        
        <!-- Tab 3: 推广费分摊 -->
        <div id="tab-promotion" class="tab-content<%= tabPromoClass %>">
            <div class="card">
                <div class="card-header">
                    <h3 class="card-title"><i class="fas fa-bullhorn"></i> 推广费归因分摊</h3>
                </div>
                <div class="card-body">
                    <form method="post" action="expense_allocation.asp" onsubmit="showProgress('正在分摊推广费...')" <%= IIf(canEdit, "", "class='readonly-mask'") %>>
                        <%= GetCSRFTokenField() %>
                        <input type="hidden" name="action" value="allocate_promotion">
                        
                        <div class="form-grid">
                            <div class="form-group">
                                <label>推广时间段 - 开始</label>
                                <input type="date" name="startDate" required>
                            </div>
                            <div class="form-group">
                                <label>推广时间段 - 结束</label>
                                <input type="date" name="endDate" required>
                            </div>
                            <div class="form-group">
                                <label>推广总费用 (元)</label>
                                <input type="number" name="totalAmount" step="0.01" min="0.01" required placeholder="输入推广总消耗">
                            </div>
                            <div class="form-group">
                                <label>有效成交额 GMV (元)</label>
                                <input type="number" name="gmvAmount" step="0.01" min="0.01" required placeholder="输入该时段总成交额">
                            </div>
                        </div>
                        
                        <div class="alert alert-warning" style="margin-top: 15px;">
                            <i class="fas fa-info-circle"></i> 
                            分摊公式：单笔承担推广费 = (总消耗 / 有效成交额) x 订单金额。系统将按各SKU销售额占比自动分摊。
                        </div>
                        
                        <% If canEdit Then %>
                        <div style="margin-top: 20px;">
                            <button type="submit" class="btn btn-success">
                                <i class="fas fa-play"></i> 开始分摊
                            </button>
                        </div>
                        <% End If %>
                    </form>
                </div>
            </div>
        </div>
        
        <!-- Tab 4: 分摊结果 -->
        <div id="tab-results" class="tab-content<%= tabResultsClass %>">
            <!-- 统计汇总 -->
            <div class="stats-grid">
                <div class="stat-card">
                    <div class="stat-label"><i class="fas fa-truck"></i> 运费总计</div>
                    <div class="stat-value shipping">
                        ¥<%= FormatNumber(CDbl("0" & GetScalar("SELECT IIF(SUM(Amount) IS NULL, 0, SUM(Amount)) FROM ExpenseRecords WHERE ExpenseType = 'Shipping'")), 2) %>
                    </div>
                </div>
                <div class="stat-card">
                    <div class="stat-label"><i class="fas fa-percentage"></i> 平台费总计</div>
                    <div class="stat-value platform">
                        ¥<%= FormatNumber(CDbl("0" & GetScalar("SELECT IIF(SUM(Amount) IS NULL, 0, SUM(Amount)) FROM ExpenseRecords WHERE ExpenseType = 'PlatformFee'")), 2) %>
                    </div>
                </div>
                <div class="stat-card">
                    <div class="stat-label"><i class="fas fa-bullhorn"></i> 推广费总计</div>
                    <div class="stat-value promotion">
                        ¥<%= FormatNumber(CDbl("0" & GetScalar("SELECT IIF(SUM(Amount) IS NULL, 0, SUM(Amount)) FROM ExpenseRecords WHERE ExpenseType = 'Promotion'")), 2) %>
                    </div>
                </div>
                <div class="stat-card">
                    <div class="stat-label"><i class="fas fa-calculator"></i> 费用合计</div>
                    <div class="stat-value total">
                        ¥<%= FormatNumber(CDbl("0" & GetScalar("SELECT IIF(SUM(Amount) IS NULL, 0, SUM(Amount)) FROM ExpenseRecords")), 2) %>
                    </div>
                </div>
            </div>
            
            <div class="card">
                <div class="card-header">
                    <h3 class="card-title"><i class="fas fa-list"></i> 分摊明细</h3>
                </div>
                <div class="card-body">
                    <!-- 筛选条件 -->
                    <form method="get" action="expense_allocation.asp" style="margin-bottom: 20px;">
                        <input type="hidden" name="tab" value="results">
                        <div class="form-grid">
                            <div class="form-group">
                                <label>费用类型</label>
                                <select name="filterType" onchange="this.form.submit()">
                                    <option value="">全部</option>
                                    <option value="Shipping" <%= IIf(Request.QueryString("filterType")="Shipping", "selected", "") %>>运费</option>
                                    <option value="PlatformFee" <%= IIf(Request.QueryString("filterType")="PlatformFee", "selected", "") %>>平台费</option>
                                    <option value="Promotion" <%= IIf(Request.QueryString("filterType")="Promotion", "selected", "") %>>推广费</option>
                                </select>
                            </div>
                            <div class="form-group">
                                <label>月份</label>
                                <select name="filterPeriod" onchange="this.form.submit()">
                                    <option value="">全部</option>
                                    <% 
                                    Dim rsPeriods
                                    Set rsPeriods = ExecuteQuery("SELECT DISTINCT Period FROM ExpenseRecords WHERE Period IS NOT NULL ORDER BY Period DESC")
                                    If Not rsPeriods Is Nothing Then
                                        Do While Not rsPeriods.EOF
                                            Dim periodVal
                                            periodVal = rsPeriods("Period").Value
                                    %>
                                    <option value="<%= periodVal %>" <%= IIf(Request.QueryString("filterPeriod")=periodVal, "selected", "") %>><%= periodVal %></option>
                                    <%
                                            rsPeriods.MoveNext
                                        Loop
                                        rsPeriods.Close
                                        Set rsPeriods = Nothing
                                    End If
                                    %>
                                </select>
                            </div>
                        </div>
                    </form>
                    
                    <!-- 分摊明细表 -->
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th>ID</th>
                                <th>订单ID</th>
                                <th>商品</th>
                                <th>费用类型</th>
                                <th>费用名称</th>
                                <th>分摊金额</th>
                                <th>分摊方式</th>
                                <th>月份</th>
                                <th>创建时间</th>
                                <% If canEdit Then %>
                                <th>操作</th>
                                <% End If %>
                            </tr>
                        </thead>
                        <tbody>
                            <% 
                            Dim filterSQL, filterWhere
                            filterWhere = "1=1"
                            If Request.QueryString("filterType") <> "" Then
                                filterWhere = filterWhere & " AND er.ExpenseType = '" & SafeSQL(Request.QueryString("filterType")) & "'"
                            End If
                            If Request.QueryString("filterPeriod") <> "" Then
                                filterWhere = filterWhere & " AND er.Period = '" & SafeSQL(Request.QueryString("filterPeriod")) & "'"
                            End If
                            
                            filterSQL = "SELECT TOP 50 er.*, p.ProductName FROM (ExpenseRecords AS er LEFT JOIN Products AS p ON er.ProductID = p.ProductID) WHERE " & filterWhere & " ORDER BY er.ExpenseID DESC"
                            
                            Dim rsExpenses
                            Set rsExpenses = ExecuteQuery(filterSQL)
                            
                            If Not rsExpenses Is Nothing Then
                                Do While Not rsExpenses.EOF
                                    Dim expenseTypeClass
                                    Select Case rsExpenses("ExpenseType").Value
                                        Case "Shipping": expenseTypeClass = "style='color: #4CAF50;'"
                                        Case "PlatformFee": expenseTypeClass = "style='color: #2196F3;'"
                                        Case "Promotion": expenseTypeClass = "style='color: #ffa726;'"
                                        Case Else: expenseTypeClass = ""
                                    End Select
                            %>
                            <tr>
                                <td><%= rsExpenses("ExpenseID").Value %></td>
                                <td><%= rsExpenses("OrderID").Value %></td>
                                <td><%= Server.HTMLEncode(rsExpenses("ProductName").Value & "") %></td>
                                <td <%= expenseTypeClass %>><%= rsExpenses("ExpenseType").Value %></td>
                                <td><%= Server.HTMLEncode(rsExpenses("ExpenseName").Value & "") %></td>
                                <td>¥<%= FormatNumber(CDbl("0" & rsExpenses("Amount").Value), 2) %></td>
                                <td><%= rsExpenses("AllocationMethod").Value %></td>
                                <td><%= rsExpenses("Period").Value %></td>
                                <td><%= FormatDateField(rsExpenses("CreatedAt").Value) %></td>
                                <% If canEdit Then %>
                                <td>
                                    <button class="btn btn-secondary" onclick="adjustExpense(<%= rsExpenses("ExpenseID").Value %>, <%= CDbl("0" & rsExpenses("Amount").Value) %>)">
                                        <i class="fas fa-edit"></i> 调整
                                    </button>
                                </td>
                                <% End If %>
                            </tr>
                            <%
                                    rsExpenses.MoveNext
                                Loop
                                rsExpenses.Close
                                Set rsExpenses = Nothing
                            End If
                            %>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    
    <!-- 进度提示遮罩 -->
    <div id="progressOverlay" class="progress-overlay">
        <div class="progress-box">
            <div class="progress-spinner"></div>
            <div id="progressText">正在处理...</div>
        </div>
    </div>
    
    <!-- 调整金额弹窗 -->
    <div id="adjustModal" style="display: none; position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.8); z-index: 9998; align-items: center; justify-content: center;">
        <div style="background: var(--bg-card); padding: 30px; border-radius: 16px; width: 400px; border: 1px solid var(--border-color);">
            <h3 style="margin: 0 0 20px 0;"><i class="fas fa-edit"></i> 调整分摊金额</h3>
            <form method="post" action="expense_allocation.asp">
                <%= GetCSRFTokenField() %>
                <input type="hidden" name="action" value="adjust_expense">
                <input type="hidden" name="expenseId" id="adjustExpenseId">
                
                <div class="form-group">
                    <label>新金额</label>
                    <input type="number" name="newAmount" id="adjustNewAmount" step="0.01" min="0" required>
                </div>
                <div class="form-group">
                    <label>调整原因</label>
                    <input type="text" name="adjustReason" placeholder="输入调整原因" required>
                </div>
                
                <div style="display: flex; gap: 10px; margin-top: 20px;">
                    <button type="button" class="btn btn-secondary" onclick="closeAdjustModal()" style="flex: 1;">取消</button>
                    <button type="submit" class="btn btn-primary" style="flex: 1;">确认调整</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
        // Tab 切换
        function switchTab(tabName, evt) {
            // 隐藏所有 Tab 内容
            document.querySelectorAll('.tab-content').forEach(function(el) {
                el.classList.remove('active');
            });
            
            // 移除所有按钮激活状态
            document.querySelectorAll('.tab-btn').forEach(function(el) {
                el.classList.remove('active');
            });
            
            // 显示当前 Tab
            document.getElementById('tab-' + tabName).classList.add('active');
            
            // 激活当前按钮
            if (evt && evt.target) {
                evt.target.closest('.tab-btn').classList.add('active');
            }
            
            // 更新 URL hash
            window.location.hash = tabName;
        }
        
        // 显示进度提示
        function showProgress(text) {
            document.getElementById('progressText').textContent = text || '正在处理...';
            document.getElementById('progressOverlay').classList.add('active');
        }
        
        // 调整金额弹窗
        function adjustExpense(id, currentAmount) {
            document.getElementById('adjustExpenseId').value = id;
            document.getElementById('adjustNewAmount').value = currentAmount;
            document.getElementById('adjustModal').style.display = 'flex';
        }
        
        function closeAdjustModal() {
            document.getElementById('adjustModal').style.display = 'none';
        }
        
        // 页面加载时根据 hash 切换 Tab
        window.addEventListener('load', function() {
            var hash = window.location.hash.replace('#', '');
            if (hash && document.getElementById('tab-' + hash)) {
                var tabBtn = document.querySelector('.tab-btn[onclick*="' + hash + '"]');
                if (tabBtn) {
                    tabBtn.click();
                }
            }
        });
        
        // 点击弹窗外部关闭
        document.getElementById('adjustModal').addEventListener('click', function(e) {
            if (e.target === this) {
                closeAdjustModal();
            }
        });
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
