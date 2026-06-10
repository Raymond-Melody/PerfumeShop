<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' 统一业务规则模块
' 供应链全局常量、验证函数、工具函数
' ============================================

' ========== 状态常量定义 ==========
' 采购订单状态
Const PO_STATUS_DRAFT = "Draft"
Const PO_STATUS_SUBMITTED = "Submitted"
Const PO_STATUS_ORDERED = "Ordered"
Const PO_STATUS_PARTIAL_RECEIVED = "PartialReceived"
Const PO_STATUS_RECEIVED = "Received"
Const PO_STATUS_CANCELLED = "Cancelled"

' 收货单状态
Const RECEIPT_STATUS_PENDING = "Pending"
Const RECEIPT_STATUS_PARTIAL = "Partial"
Const RECEIPT_STATUS_COMPLETE = "Complete"

' 出库类型
Const OUTBOUND_TYPE_PRODUCTION = "Production"
Const OUTBOUND_TYPE_RETURN = "Return"
Const OUTBOUND_TYPE_SCRAP = "Scrap"

' 生产单状态
Const PROD_STATUS_DRAFT = "Draft"
Const PROD_STATUS_SCHEDULED = "Scheduled"
Const PROD_STATUS_INPROGRESS = "InProgress"
Const PROD_STATUS_COMPLETED = "Completed"
Const PROD_STATUS_QC = "QC"
Const PROD_STATUS_QC_FAIL = "QC_Fail"

' 配方发布状态
Const PUBLISH_STATUS_DRAFT = "Draft"
Const PUBLISH_STATUS_PUBLISHED = "Published"
Const PUBLISH_STATUS_DEPRECATED = "Deprecated"

' 调拨状态
Const TRANSFER_STATUS_REQUESTED = "Requested"
Const TRANSFER_STATUS_ACCEPTED = "Accepted"
Const TRANSFER_STATUS_FULFILLED = "Fulfilled"

' 库存变动方向
Const TX_DIRECTION_IN = "IN"
Const TX_DIRECTION_OUT = "OUT"

' 库存变动类型
Const TX_TYPE_PURCHASE_IN = "采购入库"
Const TX_TYPE_MATERIAL_OUT = "原材料出库"
Const TX_TYPE_ACCORD_PRODUCE = "香调生产"
Const TX_TYPE_PRODUCT_MANUFACTURE = "产品制造"
Const TX_TYPE_SALES_OUT = "销售出库"
Const TX_TYPE_TRANSFER = "库存调拨"

' 车间标识
Const WORKSHOP_ACCORD = "ACCORD"
Const WORKSHOP_MANUFACTURING = "MANUFACTURING"

' 配方发布类型
Const PUBLISH_TYPE_ACCORD = "Accord"
Const PUBLISH_TYPE_PRODUCT = "Product"

' ========== 角色常量 ==========
Const ROLE_SUPER_ADMIN = "SUPER_ADMIN"
Const ROLE_PROD_MANAGER = "PROD_MANAGER"
Const ROLE_ACCORD_PRODUCER = "ACCORD_PRODUCER"
Const ROLE_PROD_MANUFACTURER = "PROD_MANUFACTURER"
Const ROLE_PROD_SCHEDULER = "PROD_SCHEDULER"
Const ROLE_PROD_OPERATOR = "PROD_OPERATOR"
Const ROLE_PROD_QC = "PROD_QC"
Const ROLE_PROD_WAREHOUSE = "PROD_WAREHOUSE"
Const ROLE_PROD_LOGISTICS = "PROD_LOGISTICS"
Const ROLE_TECH_MANAGER = "TECH_MANAGER"
Const ROLE_TECH_STAFF = "TECH_STAFF"

' ========== 安全验证函数 ==========

' 安全数字转换
Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then
        SafeNum = 0
    Else
        On Error Resume Next
        SafeNum = CDbl(val)
        If Err.Number <> 0 Then SafeNum = 0 : Err.Clear
        On Error GoTo 0
    End If
End Function

' 安全SQL字符串
Function SafeSQL(str)
    If IsNull(str) Or str = "" Then
        SafeSQL = ""
    Else
        SafeSQL = Replace(str, "'", "''")
    End If
End Function

' 安全整数转换
Function SafeInt(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then
        SafeInt = 0
    Else
        On Error Resume Next
        SafeInt = CLng(val)
        If Err.Number <> 0 Then SafeInt = 0 : Err.Clear
        On Error GoTo 0
    End If
End Function

' ========== 库存验证函数 ==========

' 验证原材料库存可用量
Function ValidateRawMaterialStock(materialID, requestQty)
    Dim stock
    stock = 0
    On Error Resume Next
    Dim rs : Set rs = conn.Execute("SELECT StockQty FROM RawMaterialInventory WHERE MaterialID=" & CLng(materialID))
    If Not rs Is Nothing Then
        If Not rs.EOF Then stock = SafeNum(rs("StockQty"))
        rs.Close
    End If
    Set rs = Nothing
    On Error GoTo 0
    ValidateRawMaterialStock = (stock >= requestQty)
End Function

' 验证香调库存可用量
Function ValidateNoteStock(noteID, requestQty)
    Dim stock
    stock = 0
    On Error Resume Next
    Dim rs : Set rs = conn.Execute("SELECT StockQuantity FROM NoteInventory WHERE NoteID=" & CLng(noteID))
    If Not rs Is Nothing Then
        If Not rs.EOF Then stock = SafeNum(rs("StockQuantity"))
        rs.Close
    End If
    Set rs = Nothing
    On Error GoTo 0
    ValidateNoteStock = (stock >= requestQty)
End Function

' 验证成品库存可用量
Function ValidateProductStock(productID, requestQty)
    Dim stock
    stock = 0
    On Error Resume Next
    Dim rs : Set rs = conn.Execute("SELECT StockQty FROM ProductInventory WHERE ProductID=" & CLng(productID) & " AND StockType='Product'")
    If Not rs Is Nothing Then
        If Not rs.EOF Then stock = SafeNum(rs("StockQty"))
        rs.Close
    End If
    Set rs = Nothing
    On Error GoTo 0
    ValidateProductStock = (stock >= requestQty)
End Function

' ========== 业务计算函数 ==========

' 生成批次号
Function GenerateBatchNo(prefix, optionalDate)
    Dim dt
    If IsEmpty(optionalDate) Or IsNull(optionalDate) Then
        dt = Now()
    Else
        dt = optionalDate
    End If
    GenerateBatchNo = prefix & Year(dt) & Right("0" & Month(dt), 2) & Right("0" & Day(dt), 2) & Right("0" & Hour(dt), 2) & Right("0" & Minute(dt), 2)
End Function

' 计算配方材料需求量
Function CalculateMaterialRequirement(percentage, batchSize, plannedQty)
    If batchSize <= 0 Then batchSize = 100
    CalculateMaterialRequirement = (percentage / batchSize) * plannedQty
End Function

' ========== 角色权限验证 ==========

' 检查是否为管理角色
Function IsAdminRole(roleCode)
    IsAdminRole = (roleCode = ROLE_SUPER_ADMIN Or roleCode = ROLE_PROD_MANAGER Or roleCode = ROLE_TECH_MANAGER)
End Function

' 检查是否可访问原材料库存
Function CanAccessRawMaterials(roleCode)
    CanAccessRawMaterials = (roleCode = ROLE_SUPER_ADMIN Or roleCode = ROLE_PROD_MANAGER Or roleCode = ROLE_ACCORD_PRODUCER Or roleCode = ROLE_TECH_MANAGER)
End Function

' 检查是否可访问成品库存
Function CanAccessProducts(roleCode)
    CanAccessProducts = (roleCode = ROLE_SUPER_ADMIN Or roleCode = ROLE_PROD_MANAGER Or roleCode = ROLE_PROD_MANUFACTURER Or roleCode = ROLE_PROD_WAREHOUSE Or roleCode = ROLE_TECH_MANAGER)
End Function

' 检查是否可访问香调配方
Function CanAccessAccordRecipes(roleCode)
    CanAccessAccordRecipes = (roleCode = ROLE_SUPER_ADMIN Or roleCode = ROLE_PROD_MANAGER Or roleCode = ROLE_ACCORD_PRODUCER Or roleCode = ROLE_TECH_MANAGER)
End Function

' 检查是否可访问产品配方
Function CanAccessProductRecipes(roleCode)
    CanAccessProductRecipes = (roleCode = ROLE_SUPER_ADMIN Or roleCode = ROLE_PROD_MANAGER Or roleCode = ROLE_PROD_MANUFACTURER Or roleCode = ROLE_TECH_MANAGER)
End Function

' ========== 状态流转验证 ==========

' 验证采购订单状态流转是否合法
Function IsValidPOStatusChange(currentStatus, newStatus)
    IsValidPOStatusChange = False
    Select Case currentStatus
        Case PO_STATUS_DRAFT
            If newStatus = PO_STATUS_SUBMITTED Or newStatus = PO_STATUS_CANCELLED Then IsValidPOStatusChange = True
        Case PO_STATUS_SUBMITTED
            If newStatus = PO_STATUS_ORDERED Or newStatus = PO_STATUS_CANCELLED Then IsValidPOStatusChange = True
        Case PO_STATUS_ORDERED
            If newStatus = PO_STATUS_PARTIAL_RECEIVED Or newStatus = PO_STATUS_RECEIVED Or newStatus = PO_STATUS_CANCELLED Then IsValidPOStatusChange = True
        Case PO_STATUS_PARTIAL_RECEIVED
            If newStatus = PO_STATUS_RECEIVED Then IsValidPOStatusChange = True
    End Select
End Function

' 验证生产单状态流转是否合法
Function IsValidProductionStatusChange(currentStatus, newStatus)
    IsValidProductionStatusChange = False
    Select Case currentStatus
        Case PROD_STATUS_DRAFT
            If newStatus = PROD_STATUS_SCHEDULED Then IsValidProductionStatusChange = True
        Case PROD_STATUS_SCHEDULED
            If newStatus = PROD_STATUS_INPROGRESS Then IsValidProductionStatusChange = True
        Case PROD_STATUS_INPROGRESS
            If newStatus = PROD_STATUS_COMPLETED Then IsValidProductionStatusChange = True
        Case PROD_STATUS_COMPLETED
            If newStatus = PROD_STATUS_QC Or newStatus = PROD_STATUS_QC_FAIL Then IsValidProductionStatusChange = True
        Case PROD_STATUS_QC_FAIL
            If newStatus = PROD_STATUS_INPROGRESS Then IsValidProductionStatusChange = True  ' 返工
    End Select
End Function

' ========== 库存变动记录函数 ==========

' 记录库存变动流水
Function RecordInventoryTransaction(itemType, itemID, qty, txType, txDirection, notes, createdBy)
    On Error Resume Next
    Dim sql
    sql = "INSERT INTO InventoryTransactions ("
    If itemType = "Material" Then
        sql = sql & "MaterialID, "
    ElseIf itemType = "Note" Then
        sql = sql & "NoteID, "
    ElseIf itemType = "Product" Then
        sql = sql & "ProductID, "
    End If
    sql = sql & "Quantity, TransactionType, TransactionDirection, Notes, CreatedBy, CreatedAt) VALUES ("
    sql = sql & itemID & ", " & qty & ", '" & SafeSQL(txType) & "', '" & SafeSQL(txDirection) & "', '" & SafeSQL(notes) & "', '" & SafeSQL(createdBy) & "', GETDATE())"
    
    conn.Execute sql
    If Err.Number <> 0 Then
        RecordInventoryTransaction = False
        Err.Clear
    Else
        RecordInventoryTransaction = True
    End If
    On Error GoTo 0
End Function

' ========== 辅助显示函数 ==========

' 获取状态显示名称
Function GetStatusDisplayName(statusCode)
    Select Case statusCode
        Case PO_STATUS_DRAFT : GetStatusDisplayName = "草稿"
        Case PO_STATUS_SUBMITTED : GetStatusDisplayName = "已提交"
        Case PO_STATUS_ORDERED : GetStatusDisplayName = "已下单"
        Case PO_STATUS_PARTIAL_RECEIVED : GetStatusDisplayName = "部分收货"
        Case PO_STATUS_RECEIVED : GetStatusDisplayName = "已收货"
        Case PO_STATUS_CANCELLED : GetStatusDisplayName = "已取消"
        Case PROD_STATUS_DRAFT : GetStatusDisplayName = "草稿"
        Case PROD_STATUS_SCHEDULED : GetStatusDisplayName = "已排产"
        Case PROD_STATUS_INPROGRESS : GetStatusDisplayName = "生产中"
        Case PROD_STATUS_COMPLETED : GetStatusDisplayName = "已完成"
        Case PROD_STATUS_QC : GetStatusDisplayName = "质检通过"
        Case PROD_STATUS_QC_FAIL : GetStatusDisplayName = "质检未通过"
        Case TRANSFER_STATUS_REQUESTED : GetStatusDisplayName = "待处理"
        Case TRANSFER_STATUS_ACCEPTED : GetStatusDisplayName = "已接受"
        Case TRANSFER_STATUS_FULFILLED : GetStatusDisplayName = "已完成"
        Case Else : GetStatusDisplayName = statusCode
    End Select
End Function

' 获取库存状态（充足/低库存/缺货）
Function GetStockStatus(current, minimum)
    If current <= 0 Then
        GetStockStatus = "缺货"
    ElseIf current <= minimum Then
        GetStockStatus = "低库存"
    Else
        GetStockStatus = "充足"
    End If
End Function

' 三目运算符
Function IIF(cond, tVal, fVal)
    If cond Then IIF = tVal Else IIF = fVal
End Function

' ========== V8 跨模块业务规则 ==========

' 验证订单-生产-物流全链路数据一致性
' 检查: Orders → ProductionOrders → LogisticsShipments 状态链
Function ValidateOrderProductionChain(orderID)
    Dim issues : issues = ""
    On Error Resume Next
    
    ' 检查生产工单是否存在
    Dim prodCount : prodCount = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE OrderID=" & CLng(orderID)))
    
    ' 有订单但无生产工单（跳过取消/退款订单）
    Dim orderStatus : orderStatus = GetScalar("SELECT Status FROM Orders WHERE OrderID=" & CLng(orderID))
    If orderStatus <> "cancelled" And orderStatus <> "refunding" Then
        If prodCount = 0 Then
            issues = issues & IIf(issues<>"","; ","") & "缺失生产工单"
        End If
    End If
    
    ' 检查生产-物流状态联动
    Dim shipStatus : shipStatus = GetScalar("SELECT ShippingStatus FROM Orders WHERE OrderID=" & CLng(orderID))
    If shipStatus = "shipped" Or shipStatus = "delivered" Then
        Dim logisticsExists : logisticsExists = SafeNum(GetScalar("SELECT COUNT(*) FROM LogisticsShipments WHERE OrderID=" & CLng(orderID)))
        If logisticsExists = 0 Then
            issues = issues & IIf(issues<>"","; ","") & "已发货但缺失物流记录"
        End If
    End If
    
    On Error GoTo 0
    ValidateOrderProductionChain = issues
End Function

' 验证采购入库-库存一致性
Function ValidatePurchaseInventoryMatch(poID)
    Dim issues : issues = ""
    On Error Resume Next
    
    Dim orderType : orderType = GetScalar("SELECT OrderType FROM PurchaseOrders WHERE PurchaseOrderID=" & CLng(poID))
    Dim receivedQty : receivedQty = SafeNum(GetScalar("SELECT SUM(ReceivedQty) FROM ReceivingRecords WHERE PurchaseOrderID=" & CLng(poID)))
    Dim inventoryQty : inventoryQty = 0
    
    Select Case orderType
        Case "RawMaterial"
            inventoryQty = SafeNum(GetScalar("SELECT SUM(StockQty) FROM RawMaterialInventory"))
        Case "Packaging"
            inventoryQty = SafeNum(GetScalar("SELECT SUM(StockQty) FROM PackagingInventory"))
        Case "Bottle"
            inventoryQty = SafeNum(GetScalar("SELECT SUM(StockQty) FROM BottleStyles"))
    End Select
    
    If receivedQty > 0 And inventoryQty = 0 Then
        issues = "已收货但库存未同步更新"
    End If
    
    On Error GoTo 0
    ValidatePurchaseInventoryMatch = issues
End Function

' 检查应收账款-订单支付一致性
Function ValidateReceivableOrderMatch(orderID)
    Dim issues : issues = ""
    On Error Resume Next
    
    Dim orderAmount : orderAmount = SafeNum(GetScalar("SELECT TotalAmount FROM Orders WHERE OrderID=" & CLng(orderID)))
    Dim receivedAmount : receivedAmount = SafeNum(GetScalar("SELECT ISNULL(SUM(PaidAmount),0) FROM AccountsReceivable WHERE OrderID=" & CLng(orderID)))
    
    If orderAmount > 0 And receivedAmount > 0 And Abs(orderAmount - receivedAmount) > 0.01 Then
        issues = "应收金额(" & FormatNumber(orderAmount,2) & ")与实收(" & FormatNumber(receivedAmount,2) & ")不一致"
    End If
    
    On Error GoTo 0
    ValidateReceivableOrderMatch = issues
End Function

' 跨模块健康检查（供仪表盘调用）
Function GetCrossModuleHealthCheck()
    Dim results(3, 2)
    On Error Resume Next
    
    ' 1. 无主采购单
    Dim orphanPO : orphanPO = SafeNum(GetScalar("SELECT COUNT(*) FROM PurchaseOrders WHERE SupplierID NOT IN (SELECT SupplierID FROM Suppliers)"))
    results(0, 0) = "无主采购单"
    results(0, 1) = IIf(orphanPO>0,"warn","ok")
    results(0, 2) = orphanPO & " 条"
    
    ' 2. 生产-库存联动
    Dim prodStockGap : prodStockGap = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE Status='Completed' AND OrderID NOT IN (SELECT OrderID FROM ProductInventory WHERE StockType='Product')"))
    results(1, 0) = "生产-成品库存断链"
    results(1, 1) = IIf(prodStockGap>0,"warn","ok")
    results(1, 2) = prodStockGap & " 条"
    
    ' 3. 应付-付款匹配
    Dim apPayGap : apPayGap = SafeNum(GetScalar("SELECT COUNT(*) FROM AccountsPayable WHERE Amount > 0 AND PaidAmount = 0 AND DATEDIFF(DAY,CreatedAt,GETDATE()) > 30"))
    results(2, 0) = "超30天未付应付"
    results(2, 1) = IIf(apPayGap>0,"danger","ok")
    results(2, 2) = apPayGap & " 条"
    
    ' 4. 半成品-成品流转
    Dim semifToProd : semifToProd = SafeNum(GetScalar("SELECT COUNT(*) FROM WorkshopTransfers WHERE Status='Requested' AND DATEDIFF(DAY,CreatedAt,GETDATE()) > 7"))
    results(3, 0) = "超7天未处理调拨"
    results(3, 1) = IIf(semifToProd>0,"warn","ok")
    results(3, 2) = semifToProd & " 条"
    
    On Error GoTo 0
    GetCrossModuleHealthCheck = results
End Function
%>
%>