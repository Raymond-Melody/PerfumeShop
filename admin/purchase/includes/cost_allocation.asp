<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' cost_allocation.asp — 采购成本分摊函数库 (V10)
' 功能：从库存批次记录中按FIFO加权平均方式分摊成本到订单
' ============================================

' 外部依赖：需要 conn 已通过 OpenConnection() 打开
' 调用方式：<!--#include file="includes/cost_allocation.asp"-->

' ========== SafeNum — 安全数值转换 ==========
Function CASafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then
        CASafeNum = 0
    Else
        On Error Resume Next
        CASafeNum = CDbl(val)
        If Err.Number <> 0 Then CASafeNum = 0
        On Error GoTo 0
    End If
End Function

' ========== SafeSQL — SQL注入防护 ==========
Function CASafeSQL(str)
    If IsNull(str) Or str = "" Then
        CASafeSQL = ""
    Else
        CASafeSQL = Replace(str, "'", "''")
    End If
End Function

' ========== 获取库存物料的当前加权平均成本 ==========
' 参数: itemID - 库存记录ID, orderType - 品类类型
' 返回: 加权平均成本值
Function GetWeightedCost(itemID, orderType)
    Dim sql, rs, cost
    cost = 0
    
    On Error Resume Next
    If orderType = "RawMaterial" Then
        sql = "SELECT ISNULL(WeightedUnitCost, UnitPrice) FROM RawMaterialInventory WHERE MaterialID=" & CASafeNum(itemID)
    ElseIf orderType = "Packaging" Then
        sql = "SELECT ISNULL(WeightedUnitCost, UnitPrice) FROM PackagingInventory WHERE PackagingID=" & CASafeNum(itemID)
    ElseIf orderType = "Bottle" Then
        sql = "SELECT ISNULL(WeightedUnitCost, UnitPrice) FROM BottleStyles WHERE BottleID=" & CASafeNum(itemID)
    ElseIf orderType = "Printing" Then
        sql = "SELECT ISNULL(WeightedUnitCost, UnitPrice) FROM PrintingInventory WHERE PrintingID=" & CASafeNum(itemID)
    ElseIf orderType = "SprayHead" Then
        sql = "SELECT ISNULL(WeightedUnitCost, UnitPrice) FROM SprayHeadInventory WHERE SprayHeadID=" & CASafeNum(itemID)
    Else
        GetWeightedCost = 0
        Exit Function
    End If
    
    Set rs = conn.Execute(sql)
    If Not rs Is Nothing And Not rs.EOF Then
        cost = CASafeNum(rs(0))
        rs.Close
    End If
    Set rs = Nothing
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    
    GetWeightedCost = cost
End Function

' ========== 从InventoryBatches中按FIFO分配成本 ==========
' 参数: itemID - 库存记录ID, orderType - 品类类型, qtyNeeded - 需要分配的数量
' 返回: 总成本（加权）
Function AllocateInventoryCostFromBatches(itemID, orderType, qtyNeeded)
    Dim totalCost, remaining, rsBatches, batchQty, batchCost, allocQty
    totalCost = 0
    remaining = CASafeNum(qtyNeeded)
    
    If remaining <= 0 Then
        AllocateInventoryCostFromBatches = 0
        Exit Function
    End If
    
    On Error Resume Next
    ' 查询该库存项的未完全消耗批次(按收货日期先后 FIFO)
    Dim sqlBatches
    sqlBatches = "SELECT InvBatchID, UnitCost, StockQty FROM InventoryBatches " & _
                 "WHERE ItemType='" & CASafeSQL(orderType) & "' AND ItemID=" & CASafeNum(itemID) & " " & _
                 "AND StockQty > 0 ORDER BY CreatedAt ASC"
    
    Set rsBatches = conn.Execute(sqlBatches)
    If Not rsBatches Is Nothing Then
        Do While Not rsBatches.EOF And remaining > 0
            batchQty = CASafeNum(rsBatches("StockQty"))
            batchCost = CASafeNum(rsBatches("UnitCost"))
            
            If batchQty > 0 Then
                ' 确定从该批次分配的数量
                If batchQty >= remaining Then
                    allocQty = remaining
                Else
                    allocQty = batchQty
                End If
                
                ' 累加成本
                totalCost = totalCost + (allocQty * batchCost)
                
                ' 更新批次剩余数量
                Dim newRemaining
                newRemaining = batchQty - allocQty
                conn.Execute "UPDATE InventoryBatches SET StockQty = " & newRemaining & " WHERE InvBatchID=" & rsBatches("InvBatchID")
                
                remaining = remaining - allocQty
            End If
            
            rsBatches.MoveNext
        Loop
        rsBatches.Close
    End If
    Set rsBatches = Nothing
    
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    
    AllocateInventoryCostFromBatches = totalCost
End Function

' ========== 获取库存物料ID（按名称查找）==========
' 参数: itemName - 物料名称, orderType - 品类类型
' 返回: 库存记录ID (0表示未找到)
Function GetInventoryItemID(itemName, orderType)
    Dim sql, rs, itemID
    itemID = 0
    
    On Error Resume Next
    If orderType = "RawMaterial" Then
        sql = "SELECT MaterialID FROM RawMaterialInventory WHERE ItemName='" & CASafeSQL(itemName) & "'"
    ElseIf orderType = "Packaging" Then
        sql = "SELECT PackagingID FROM PackagingInventory WHERE ItemName='" & CASafeSQL(itemName) & "'"
    ElseIf orderType = "Bottle" Then
        sql = "SELECT BottleID FROM BottleStyles WHERE BottleName='" & CASafeSQL(itemName) & "'"
    ElseIf orderType = "Printing" Then
        sql = "SELECT PrintingID FROM PrintingInventory WHERE ItemName='" & CASafeSQL(itemName) & "'"
    ElseIf orderType = "SprayHead" Then
        sql = "SELECT SprayHeadID FROM SprayHeadInventory WHERE ItemName='" & CASafeSQL(itemName) & "'"
    Else
        GetInventoryItemID = 0
        Exit Function
    End If
    
    Set rs = conn.Execute(sql)
    If Not rs Is Nothing And Not rs.EOF Then
        itemID = CLng(rs(0))
        rs.Close
    End If
    Set rs = Nothing
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    
    GetInventoryItemID = itemID
End Function

' ========== 记录订单成本分摊 ==========
' 参数: orderID, orderNo, costType, itemCode, itemName, unitCost, qty, batchID(可选)
' 返回: 新AllocationID
Function RecordOrderCostAllocation(orderID, orderNo, costType, itemCode, itemName, unitCost, qty, batchID)
    Dim sql, allocID, rs
    allocID = 0
    
    On Error Resume Next
    sql = "INSERT INTO OrderCostAllocation (OrderID, OrderNo, CostType, ItemCode, ItemName, UnitCost, Quantity, TotalCost, BatchID, AllocatedAt, CreatedAt) VALUES (" & _
        CASafeNum(orderID) & ", '" & CASafeSQL(orderNo) & "', '" & CASafeSQL(costType) & "', '" & CASafeSQL(itemCode) & "', '" & CASafeSQL(itemName) & "', " & _
        CASafeNum(unitCost) & ", " & CASafeNum(qty) & ", " & CASafeNum(unitCost * qty) & ", " & IIf(CASafeNum(batchID) > 0, CStr(CASafeNum(batchID)), "Null") & ", GETDATE(), GETDATE())"
    
    conn.Execute sql
    If Err.Number = 0 Then
        Set rs = conn.Execute("SELECT SCOPE_IDENTITY()")
        If Not rs Is Nothing And Not rs.EOF Then
            allocID = CLng(rs(0))
            rs.Close
        End If
        Set rs = Nothing
    Else
        Err.Clear
    End If
    On Error GoTo 0
    
    RecordOrderCostAllocation = allocID
End Function

' ========== 综合函数：为订单物料分摊成本 ==========
' 参数: orderID, orderNo, itemName, itemCode, orderType, qtyUsed
' 返回: 分摊的总成本
Function AllocateOrderMaterialCost(orderID, orderNo, itemName, itemCode, orderType, qtyUsed)
    Dim itemID, totalCost, weightedCost, batchCost
    
    itemID = GetInventoryItemID(itemName, orderType)
    qtyUsed = CASafeNum(qtyUsed)
    
    If itemID <= 0 Or qtyUsed <= 0 Then
        AllocateOrderMaterialCost = 0
        Exit Function
    End If
    
    ' 1. 从InventoryBatches按FIFO分摊
    batchCost = AllocateInventoryCostFromBatches(itemID, orderType, qtyUsed)
    
    ' 2. 如果批次分摊为0（无批次记录），回退到加权平均成本
    If batchCost <= 0 Then
        weightedCost = GetWeightedCost(itemID, orderType)
        totalCost = qtyUsed * weightedCost
    Else
        totalCost = batchCost
    End If
    
    ' 3. 计算单位成本并记录分摊
    Dim unitCost
    If qtyUsed > 0 Then
        unitCost = totalCost / qtyUsed
    Else
        unitCost = 0
    End If
    
    ' 4. 写入OrderCostAllocation
    RecordOrderCostAllocation orderID, orderNo, orderType, itemCode, itemName, unitCost, qtyUsed, 0
    
    AllocateOrderMaterialCost = totalCost
End Function

' ========== 获取品类对应的库存表名 ==========
Function GetInventoryTable(orderType)
    If orderType = "RawMaterial" Then
        GetInventoryTable = "RawMaterialInventory"
    ElseIf orderType = "Packaging" Then
        GetInventoryTable = "PackagingInventory"
    ElseIf orderType = "Bottle" Then
        GetInventoryTable = "BottleStyles"
    ElseIf orderType = "Printing" Then
        GetInventoryTable = "PrintingInventory"
    ElseIf orderType = "SprayHead" Then
        GetInventoryTable = "SprayHeadInventory"
    Else
        GetInventoryTable = ""
    End If
End Function

' ========== 获取品类对应的ID字段名 ==========
Function GetInventoryIDField(orderType)
    If orderType = "RawMaterial" Then
        GetInventoryIDField = "MaterialID"
    ElseIf orderType = "Packaging" Then
        GetInventoryIDField = "PackagingID"
    ElseIf orderType = "Bottle" Then
        GetInventoryIDField = "BottleID"
    ElseIf orderType = "Printing" Then
        GetInventoryIDField = "PrintingID"
    ElseIf orderType = "SprayHead" Then
        GetInventoryIDField = "SprayHeadID"
    Else
        GetInventoryIDField = ""
    End If
End Function
%>
