<%
' ============================================
' V17.0 DAL - 库存数据访问层
' 依赖: dal.asp, connection.asp
' 用法: <!--#include file="dal_inventory.asp"-->
' 涵盖：ProductInventory, BottleInventory, PackagingInventory,
'       RawMaterialInventory, NoteInventory, StockMovements
' ============================================

' ============================================
' 成品库存 - 获取库存预警列表
' ============================================
Function DAL_Inv_GetProductAlerts()
    Dim sql
    sql = "SELECT pi.*, p.ProductName FROM ProductInventory pi " & _
          "LEFT JOIN Products p ON pi.ProductID=p.ProductID " & _
          "WHERE pi.StockQty <= pi.SafetyStock AND pi.SafetyStock > 0 " & _
          "ORDER BY CASE WHEN pi.StockQty<=0 THEN 0 ELSE 1 END, pi.StockQty ASC"
    Set DAL_Inv_GetProductAlerts = DAL_GetList(sql, Null)
End Function

' ============================================
' 瓶子库存 - 获取预警列表
' ============================================
Function DAL_Inv_GetBottleAlerts()
    Dim sql
    sql = "SELECT * FROM BottleInventory " & _
          "WHERE StockQty <= SafetyStock AND SafetyStock > 0 " & _
          "ORDER BY CASE WHEN StockQty<=0 THEN 0 ELSE 1 END, StockQty ASC"
    Set DAL_Inv_GetBottleAlerts = DAL_GetList(sql, Null)
End Function

' ============================================
' 包装物库存 - 获取预警列表
' ============================================
Function DAL_Inv_GetPackagingAlerts()
    Dim sql
    sql = "SELECT * FROM PackagingInventory " & _
          "WHERE StockQty <= SafetyStock AND SafetyStock > 0 " & _
          "ORDER BY CASE WHEN StockQty<=0 THEN 0 ELSE 1 END, StockQty ASC"
    Set DAL_Inv_GetPackagingAlerts = DAL_GetList(sql, Null)
End Function

' ============================================
' 原料库存 - 获取预警列表
' ============================================
Function DAL_Inv_GetRawMaterialAlerts()
    Dim sql
    sql = "SELECT * FROM RawMaterialInventory " & _
          "WHERE StockQty <= SafetyStock AND SafetyStock > 0 " & _
          "ORDER BY CASE WHEN StockQty<=0 THEN 0 ELSE 1 END, StockQty ASC"
    Set DAL_Inv_GetRawMaterialAlerts = DAL_GetList(sql, Null)
End Function

' ============================================
' 香调库存 - 获取预警列表（含名称和类型）
' ============================================
Function DAL_Inv_GetNoteAlerts()
    Dim sql
    sql = "SELECT ni.*, fn.NoteName, fn.NoteType FROM NoteInventory ni " & _
          "INNER JOIN FragranceNotes fn ON ni.NoteID=fn.NoteID " & _
          "WHERE ni.StockQuantity <= ni.MinStockLevel AND ni.MinStockLevel > 0 " & _
          "ORDER BY CASE WHEN ni.StockQuantity<=0 THEN 0 ELSE 1 END, ni.StockQuantity ASC"
    Set DAL_Inv_GetNoteAlerts = DAL_GetList(sql, Null)
End Function

' ============================================
' 成品库存 - 按ID获取
' ============================================
Function DAL_Inv_GetProductByID(productId)
    Dim sql, params(0)
    sql = "SELECT pi.*, p.ProductName FROM ProductInventory pi " & _
          "LEFT JOIN Products p ON pi.ProductID=p.ProductID WHERE pi.ProductID=@ProductID"
    params(0) = Array("@ProductID", DAL_adInteger, 0, CLng(productId))
    Set DAL_Inv_GetProductByID = DAL_GetRow(sql, params)
End Function

' ============================================
' 瓶子库存 - 按ID获取
' ============================================
Function DAL_Inv_GetBottleByID(bottleId)
    Dim sql, params(0)
    sql = "SELECT * FROM BottleInventory WHERE BottleID=@BottleID"
    params(0) = Array("@BottleID", DAL_adInteger, 0, CLng(bottleId))
    Set DAL_Inv_GetBottleByID = DAL_GetRow(sql, params)
End Function

' ============================================
' 成品库存 - 更新库存
' ============================================
Function DAL_Inv_UpdateProductStock(productId, stockQty)
    Dim sql, params(1)
    sql = "UPDATE ProductInventory SET StockQty=@StockQty, UpdatedAt=GETDATE() WHERE ProductID=@ProductID"
    params(0) = Array("@StockQty", DAL_adDouble, 0, CDbl(stockQty))
    params(1) = Array("@ProductID", DAL_adInteger, 0, CLng(productId))
    DAL_Inv_UpdateProductStock = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' 瓶子库存 - 更新库存
' ============================================
Function DAL_Inv_UpdateBottleStock(bottleId, stockQty)
    Dim sql, params(1)
    sql = "UPDATE BottleInventory SET StockQty=@StockQty, UpdatedAt=GETDATE() WHERE BottleID=@BottleID"
    params(0) = Array("@StockQty", DAL_adDouble, 0, CDbl(stockQty))
    params(1) = Array("@BottleID", DAL_adInteger, 0, CLng(bottleId))
    DAL_Inv_UpdateBottleStock = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' 成品库存 - 获取总数
' ============================================
Function DAL_Inv_CountProductAlerts()
    DAL_Inv_CountProductAlerts = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM ProductInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0", Null, 0))
End Function

' ============================================
' 成品库存 - 获取严重预警数（零库存）
' ============================================
Function DAL_Inv_CountProductCritical()
    DAL_Inv_CountProductCritical = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM ProductInventory WHERE StockQty <= 0 AND SafetyStock > 0", Null, 0))
End Function

' ============================================
' 瓶子库存 - 获取总数
' ============================================
Function DAL_Inv_CountBottleAlerts()
    DAL_Inv_CountBottleAlerts = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM BottleInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0", Null, 0))
End Function

' ============================================
' 包装物库存 - 获取总数
' ============================================
Function DAL_Inv_CountPackagingAlerts()
    DAL_Inv_CountPackagingAlerts = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM PackagingInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0", Null, 0))
End Function

' ============================================
' 原料库存 - 获取总数
' ============================================
Function DAL_Inv_CountRawMaterialAlerts()
    DAL_Inv_CountRawMaterialAlerts = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM RawMaterialInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0", Null, 0))
End Function

' ============================================
' 香调库存 - 获取总数
' ============================================
Function DAL_Inv_CountNoteAlerts()
    DAL_Inv_CountNoteAlerts = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM NoteInventory WHERE StockQuantity <= MinStockLevel AND MinStockLevel > 0", Null, 0))
End Function

' ============================================
' 库存流水 - 获取列表（分页+筛选）
' ============================================
Function DAL_Inv_GetMovements(movementType, startDate, endDate, search, page, pageSize, ByRef pageInfo)
    Dim sql, params(), paramCount
    
    sql = "SELECT * FROM StockMovements WHERE 1=1"
    paramCount = -1
    ReDim params(0)
    
    If movementType <> "" Then
        sql = sql & " AND MovementType=@MovementType"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@MovementType", DAL_adVarChar, 30, movementType)
    End If
    
    If startDate <> "" Then
        sql = sql & " AND CreatedAt >= @StartDate"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@StartDate", DAL_adVarChar, 20, startDate)
    End If
    
    If endDate <> "" Then
        sql = sql & " AND CreatedAt < @EndDate"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@EndDate", DAL_adVarChar, 30, endDate & " 23:59:59")
    End If
    
    If search <> "" Then
        sql = sql & " AND (ItemName LIKE '%' + @Search + '%' OR ItemCode LIKE '%' + @Search + '%' OR ReferenceNo LIKE '%' + @Search + '%')"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@Search", DAL_adVarChar, 100, search)
    End If
    
    sql = sql & " ORDER BY CreatedAt DESC"
    
    If paramCount >= 0 Then
        Set DAL_Inv_GetMovements = DAL_GetListPaged(sql, params, page, pageSize, pageInfo)
    Else
        Set DAL_Inv_GetMovements = DAL_GetListPaged(sql, Null, page, pageSize, pageInfo)
    End If
End Function

' ============================================
' 库存流水 - 统计各类型数量
' ============================================
Function DAL_Inv_CountMovementsByType(movementType)
    Dim sql, params(0)
    sql = "SELECT COUNT(*) FROM StockMovements WHERE MovementType=@MovementType"
    params(0) = Array("@MovementType", DAL_adVarChar, 30, movementType)
    DAL_Inv_CountMovementsByType = CLng(DAL_GetScalar(sql, params, 0))
End Function

' ============================================
' 库存流水 - 本月入库数
' ============================================
Function DAL_Inv_CountMonthIn()
    DAL_Inv_CountMonthIn = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM StockMovements WHERE MovementType='In' AND CreatedAt >= DATEADD(month,-1,GETDATE())", Null, 0))
End Function

' ============================================
' 库存流水 - 本月出库数
' ============================================
Function DAL_Inv_CountMonthOut()
    DAL_Inv_CountMonthOut = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM StockMovements WHERE MovementType='Out' AND CreatedAt >= DATEADD(month,-1,GETDATE())", Null, 0))
End Function

' ============================================
' 库存流水 - 按ID获取
' ============================================
Function DAL_Inv_GetMovementByID(movementId)
    Dim sql, params(0)
    sql = "SELECT * FROM StockMovements WHERE MovementID=@MovementID"
    params(0) = Array("@MovementID", DAL_adInteger, 0, CLng(movementId))
    Set DAL_Inv_GetMovementByID = DAL_GetRow(sql, params)
End Function

' ============================================
' 库存流水 - 新增记录
' ============================================
Function DAL_Inv_CreateMovement(itemType, itemId, itemName, itemCode, movementType, quantity, beforeQty, afterQty, unit, referenceNo, notes, createdBy)
    Dim sql, params(9), newId
    
    sql = "INSERT INTO StockMovements (ItemType, ItemID, ItemName, ItemCode, MovementType, Quantity, BeforeQty, AfterQty, Unit, ReferenceNo, Notes, CreatedBy) " & _
          "VALUES (@ItemType, @ItemID, @ItemName, @ItemCode, @MovementType, @Quantity, @BeforeQty, @AfterQty, @Unit, @ReferenceNo, @Notes, @CreatedBy); " & _
          "SELECT SCOPE_IDENTITY()"
    
    params(0) = Array("@ItemType", DAL_adVarChar, 30, Left(itemType, 30))
    params(1) = Array("@ItemID", DAL_adInteger, 0, CLng(itemId))
    params(2) = Array("@ItemName", DAL_adVarChar, 200, Left(itemName, 200))
    params(3) = Array("@ItemCode", DAL_adVarChar, 100, Left(itemCode, 100))
    params(4) = Array("@MovementType", DAL_adVarChar, 20, Left(movementType, 20))
    params(5) = Array("@Quantity", DAL_adDouble, 0, CDbl(quantity))
    params(6) = Array("@BeforeQty", DAL_adDouble, 0, CDbl(beforeQty))
    params(7) = Array("@AfterQty", DAL_adDouble, 0, CDbl(afterQty))
    params(8) = Array("@Unit", DAL_adVarChar, 20, Left(unit, 20))
    params(9) = Array("@ReferenceNo", DAL_adVarChar, 100, Left(referenceNo, 100))
    ReDim Preserve params(11)
    params(10) = Array("@Notes", DAL_adVarChar, 500, Left(notes, 500))
    params(11) = Array("@CreatedBy", DAL_adVarChar, 50, Left(createdBy, 50))
    
    newId = CLng(DAL_GetScalar(sql, params, 0))
    DAL_Inv_CreateMovement = newId
End Function

' ============================================
' 库存设置 - 获取预警开关状态
' ============================================
Function DAL_Inv_GetAlertEnabled()
    Dim result
    result = DAL_GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey='EnableLowStockAlert'", Null, "1")
    If IsNull(result) Or result = "" Then result = "1"
    DAL_Inv_GetAlertEnabled = result
End Function

' ============================================
' 库存总览 - 获取所有预警总计
' ============================================
Sub DAL_Inv_GetAlertTotals(ByRef totalAlerts, ByRef totalCritical, ByRef totalWarning)
    Dim pi, bt, pk, rm, nt
    Dim piC, btC, pkC, rmC, ntC
    
    pi = DAL_Inv_CountProductAlerts()
    bt = DAL_Inv_CountBottleAlerts()
    pk = DAL_Inv_CountPackagingAlerts()
    rm = DAL_Inv_CountRawMaterialAlerts()
    nt = DAL_Inv_CountNoteAlerts()
    
    piC = CLng(DAL_GetScalar("SELECT COUNT(*) FROM ProductInventory WHERE StockQty <= 0 AND SafetyStock > 0", Null, 0))
    btC = CLng(DAL_GetScalar("SELECT COUNT(*) FROM BottleInventory WHERE StockQty <= 0 AND SafetyStock > 0", Null, 0))
    pkC = CLng(DAL_GetScalar("SELECT COUNT(*) FROM PackagingInventory WHERE StockQty <= 0 AND SafetyStock > 0", Null, 0))
    rmC = CLng(DAL_GetScalar("SELECT COUNT(*) FROM RawMaterialInventory WHERE StockQty <= 0 AND SafetyStock > 0", Null, 0))
    ntC = CLng(DAL_GetScalar("SELECT COUNT(*) FROM NoteInventory WHERE StockQuantity <= 0 AND MinStockLevel > 0", Null, 0))
    
    totalAlerts = pi + bt + pk + rm + nt
    totalCritical = piC + btC + pkC + rmC + ntC
    totalWarning = totalAlerts - totalCritical
End Sub
%>
