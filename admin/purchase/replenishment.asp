<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<%
Call OpenConnection()
Server.ScriptTimeout = 60
conn.CommandTimeout = 15

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Function SafeSQL(str)
    If IsNull(str) Or str = "" Then SafeSQL = "" Else SafeSQL = Replace(str, "'", "''")
End Function

Function GetScalar(sql)
    Dim rs, val : val = 0
    On Error Resume Next
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        If Not rs Is Nothing Then
            If Not rs.EOF Then val = rs(0)
            If IsNull(val) Then val = 0
            rs.Close
        End If
    Else : Err.Clear
    End If
    Set rs = Nothing
    GetScalar = val
End Function

' ========== V11: 智能补货字段迁移（优化版）==========
' 策略：Session缓存 → 单条canary SQL → 批量IF NOT EXISTS迁移
' 常规请求：0次SQL；首次Session：1次SQL（canary通过）或2次SQL（canary+批量迁移）
If Session("ReplenishSchemaReady") <> "1" Then
    Dim needsMigrate
    needsMigrate = False
    On Error Resume Next
    conn.Execute "SELECT TOP 1 StatID FROM PurchaseHistoryStats WHERE 1=0"
    If Err.Number <> 0 Then needsMigrate = True
    Err.Clear
    On Error GoTo 0

    If needsMigrate Then
        ' 需要迁移 - 用单条批量SQL完成全部schema变更
        ' 使用 IF NOT EXISTS 模式，幂等安全，无错误级联风险
        Dim migSQL
        migSQL = ""
        ' --- RawMaterialInventory 字段 ---
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('RawMaterialInventory') AND name='AvgDailyUsage') ALTER TABLE RawMaterialInventory ADD AvgDailyUsage DECIMAL(19,6) DEFAULT 0; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('RawMaterialInventory') AND name='LeadTimeDays') ALTER TABLE RawMaterialInventory ADD LeadTimeDays INT DEFAULT 7; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('RawMaterialInventory') AND name='LastReplenishDate') ALTER TABLE RawMaterialInventory ADD LastReplenishDate DATETIME; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('RawMaterialInventory') AND name='ReorderPoint') ALTER TABLE RawMaterialInventory ADD ReorderPoint DECIMAL(19,4) DEFAULT 0; "
        ' --- PackagingInventory 字段 ---
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('PackagingInventory') AND name='AvgDailyUsage') ALTER TABLE PackagingInventory ADD AvgDailyUsage DECIMAL(19,6) DEFAULT 0; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('PackagingInventory') AND name='LeadTimeDays') ALTER TABLE PackagingInventory ADD LeadTimeDays INT DEFAULT 7; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('PackagingInventory') AND name='LastReplenishDate') ALTER TABLE PackagingInventory ADD LastReplenishDate DATETIME; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('PackagingInventory') AND name='ReorderPoint') ALTER TABLE PackagingInventory ADD ReorderPoint DECIMAL(19,4) DEFAULT 0; "
        ' --- BottleStyles 字段 ---
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('BottleStyles') AND name='AvgDailyUsage') ALTER TABLE BottleStyles ADD AvgDailyUsage DECIMAL(19,6) DEFAULT 0; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('BottleStyles') AND name='LeadTimeDays') ALTER TABLE BottleStyles ADD LeadTimeDays INT DEFAULT 7; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('BottleStyles') AND name='LastReplenishDate') ALTER TABLE BottleStyles ADD LastReplenishDate DATETIME; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('BottleStyles') AND name='ReorderPoint') ALTER TABLE BottleStyles ADD ReorderPoint DECIMAL(19,4) DEFAULT 0; "
        ' --- PrintingInventory 建表+字段 ---
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name='PrintingInventory') CREATE TABLE PrintingInventory (PrintingID INT IDENTITY(1,1) PRIMARY KEY, ItemName NVARCHAR(100), ItemCode NVARCHAR(50), StockQty DECIMAL(10,2) DEFAULT 0, SafetyStock DECIMAL(10,2) DEFAULT 0, Unit NVARCHAR(20) DEFAULT N'张', UnitPrice DECIMAL(10,2) DEFAULT 0, UpdatedAt DATETIME DEFAULT GETDATE()); "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('PrintingInventory') AND name='AvgDailyUsage') ALTER TABLE PrintingInventory ADD AvgDailyUsage DECIMAL(19,6) DEFAULT 0; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('PrintingInventory') AND name='LeadTimeDays') ALTER TABLE PrintingInventory ADD LeadTimeDays INT DEFAULT 7; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('PrintingInventory') AND name='LastReplenishDate') ALTER TABLE PrintingInventory ADD LastReplenishDate DATETIME; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('PrintingInventory') AND name='ReorderPoint') ALTER TABLE PrintingInventory ADD ReorderPoint DECIMAL(19,4) DEFAULT 0; "
        ' --- SprayHeadInventory 建表+字段 ---
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name='SprayHeadInventory') CREATE TABLE SprayHeadInventory (SprayHeadID INT IDENTITY(1,1) PRIMARY KEY, ItemName NVARCHAR(100), ItemCode NVARCHAR(50), StockQty DECIMAL(10,2) DEFAULT 0, SafetyStock DECIMAL(10,2) DEFAULT 0, Unit NVARCHAR(20) DEFAULT N'个', UnitPrice DECIMAL(10,2) DEFAULT 0, UpdatedAt DATETIME DEFAULT GETDATE()); "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('SprayHeadInventory') AND name='AvgDailyUsage') ALTER TABLE SprayHeadInventory ADD AvgDailyUsage DECIMAL(19,6) DEFAULT 0; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('SprayHeadInventory') AND name='LeadTimeDays') ALTER TABLE SprayHeadInventory ADD LeadTimeDays INT DEFAULT 7; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('SprayHeadInventory') AND name='LastReplenishDate') ALTER TABLE SprayHeadInventory ADD LastReplenishDate DATETIME; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('SprayHeadInventory') AND name='ReorderPoint') ALTER TABLE SprayHeadInventory ADD ReorderPoint DECIMAL(19,4) DEFAULT 0; "
        ' --- PurchaseHistoryStats 建表 ---
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name='PurchaseHistoryStats') CREATE TABLE PurchaseHistoryStats (StatID INT IDENTITY(1,1) PRIMARY KEY, ItemType NVARCHAR(30) NOT NULL, ItemCode NVARCHAR(100), ItemName NVARCHAR(200), Avg30DayUsage DECIMAL(19,6) DEFAULT 0, Avg90DayUsage DECIMAL(19,6) DEFAULT 0, LastOrderDate DATETIME, TotalOrders90Days INT DEFAULT 0, PreferredSupplierID INT, PreferredUnitPrice DECIMAL(19,4) DEFAULT 0, UpdatedAt DATETIME DEFAULT GETDATE()); "

        On Error Resume Next
        conn.CommandTimeout = 30  ' 给批量迁移足够时间
        conn.Execute migSQL
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        conn.CommandTimeout = 15  ' 恢复正常超时
    End If
    Session("ReplenishSchemaReady") = "1"
End If

' ========== 生成补货采购单号（与 purchase_orders.asp 格式统一）==========
Function GenerateReplenishmentNo()
    Dim today, prefix, sql, countNum, suffix
    On Error Resume Next
    today = Date()
    prefix = "PO-" & Year(today) & Right("0" & Month(today), 2) & Right("0" & Day(today), 2) & "-"
    sql = "SELECT COUNT(*) FROM PurchaseOrders WHERE PurchaseNo LIKE '" & prefix & "%'"
    countNum = SafeNum(GetScalar(sql))
    If Err.Number = 0 And IsNumeric(countNum) Then
        suffix = Right("000" & (countNum + 1), 3)
    Else
        suffix = "001"
        Err.Clear
    End If
    On Error GoTo 0
    GenerateReplenishmentNo = prefix & suffix
End Function

' ========== Category tab parameter ==========
Dim activeTab
activeTab = Trim(Request.QueryString("tab"))
If activeTab = "" Then activeTab = "RawMaterial"

' ========== POST: Generate replenishment orders ==========
Dim genMsg, genMsgType
genMsg = ""
genMsgType = "success"

If Request.ServerVariables("REQUEST_METHOD") = "POST" And Trim(Request.Form("action")) = "generate" Then
    If Not ValidateCSRFToken() Then
        genMsg = "安全令牌验证失败，请刷新页面后重试"
        genMsgType = "error"
    Else
        Dim genDetailCount, genResults
        genDetailCount = SafeNum(Request.Form("gen_count"))
        
        If genDetailCount > 0 Then
            On Error Resume Next
            Err.Clear
            Call BeginTransaction()
        
        Dim gi, gSupplierID, gOrderType, gItemName, gItemCode, gQty, gPrice, gDeliveryDate, gNotes
        Dim currentSupplier, currentOrderType, currentOrderDate, insertOrderSQL
        Dim genOrderIDs, genOrderNos
        
        Set genOrderIDs = Server.CreateObject("Scripting.Dictionary")
        Set genOrderNos = Server.CreateObject("Scripting.Dictionary")
        
        Dim genAnyError : genAnyError = False
        
        For gi = 1 To genDetailCount
            gSupplierID = SafeNum(Request.Form("gen_supplier_" & gi))
            gOrderType = Trim(Request.Form("gen_type_" & gi))
            gItemName = Trim(Request.Form("gen_name_" & gi))
            gItemCode = Trim(Request.Form("gen_code_" & gi))
            gQty = SafeNum(Request.Form("gen_qty_" & gi))
            gPrice = SafeNum(Request.Form("gen_price_" & gi))
            gDeliveryDate = Trim(Request.Form("gen_delivery_" & gi))
            gNotes = Trim(Request.Form("gen_notes_" & gi))
            
            If gQty <= 0 Then gQty = 1
            
            ' Generate PurchaseNo
            Dim newPurchaseNo
            newPurchaseNo = GenerateReplenishmentNo()
            
            ' Determine category code from order type
            Dim catCode
            Select Case gOrderType
                Case "RawMaterial" : catCode = "RAW"
                Case "Packaging"   : catCode = "PKG"
                Case "Bottle"      : catCode = "BTL"
                Case "Printing"    : catCode = "PRT"
                Case "SprayHead"   : catCode = "SPR"
                Case Else          : catCode = "GEN"
            End Select
            
            ' Group by supplier: use same purchase order for same supplier within this batch
            Dim orderKey : orderKey = CStr(gSupplierID) & "_" & gOrderType
            Dim existingOrderID : existingOrderID = 0
            
            If genOrderIDs.Exists(orderKey) Then
                existingOrderID = CLng(genOrderIDs(orderKey))
            End If
            
            If existingOrderID = 0 Then
                ' Create new PurchaseOrder
                Dim totalAmount : totalAmount = gQty * gPrice
                insertOrderSQL = "INSERT INTO PurchaseOrders (PurchaseNo, SupplierID, CategoryCode, OrderType, OrderDate, ExpectedDate, TotalAmount, Status, Remarks, CreatedBy, CreatedAt) VALUES ('" & _
                    newPurchaseNo & "', " & gSupplierID & ", '" & catCode & "', '" & gOrderType & "', GETDATE(), " & _
                    IIf(gDeliveryDate <> "", "'" & gDeliveryDate & "'", "DATEADD(DAY, 14, GETDATE())") & ", " & totalAmount & ", 'Draft', '" & SafeSQL(gNotes) & "', '" & SafeSQL(Session("AdminUsername")) & "', GETDATE())"
                
                conn.Execute insertOrderSQL
                If Err.Number <> 0 Then
                    genAnyError = True : Err.Clear
                    Exit For
                End If
                
                Dim rsNewPO : Set rsNewPO = conn.Execute("SELECT SCOPE_IDENTITY()")
                existingOrderID = 0
                If Not rsNewPO Is Nothing And Not rsNewPO.EOF Then existingOrderID = CLng(rsNewPO(0))
                rsNewPO.Close : Set rsNewPO = Nothing
                
                If existingOrderID > 0 Then
                    genOrderIDs.Add orderKey, existingOrderID
                    genOrderNos.Add orderKey, newPurchaseNo
                End If
            End If
            
            ' Insert PurchaseOrderDetail
            If existingOrderID > 0 Then
                conn.Execute "INSERT INTO PurchaseOrderDetails (PurchaseID, ItemName, ItemCode, Specification, Unit, Quantity, UnitPrice, TotalPrice, ReceivedQty) VALUES (" & _
                    existingOrderID & ", '" & SafeSQL(gItemName) & "', '" & SafeSQL(gItemCode) & "', '', '', " & gQty & ", " & gPrice & ", " & (gQty * gPrice) & ", 0)"
                If Err.Number <> 0 Then
                    genAnyError = True : Err.Clear
                    Exit For
                End If
                
                ' Update total amount
                conn.Execute "UPDATE PurchaseOrders SET TotalAmount = (SELECT ISNULL(SUM(TotalPrice),0) FROM PurchaseOrderDetails WHERE PurchaseID=" & existingOrderID & ") WHERE PurchaseID=" & existingOrderID
                If Err.Number <> 0 Then Err.Clear
            End If
        Next
        
        If Not genAnyError Then
            Call CommitTransaction()
            Dim genCount : genCount = 0
            Dim orderListStr : orderListStr = ""
            Dim k
            For Each k In genOrderNos.Keys
                genCount = genCount + 1
                orderListStr = orderListStr & genOrderNos(k) & "; "
            Next
            genMsg = "成功生成 " & genCount & " 个补货采购订单：" & orderListStr
            genMsgType = "success"
        Else
            Call RollbackTransaction()
            genMsg = "生成补货订单失败，数据已回滚"
            genMsgType = "error"
        End If
        On Error GoTo 0
    Else
        genMsg = "请至少选择一项补货物料"
        genMsgType = "error"
    End If
    End If
End If

' ========== Helper: Get latest supplier price for item ==========
Function GetLatestSupplierInfo(itemCode, orderType)
    Dim result(2)
    result(0) = 0 : result(1) = 0 : result(2) = ""
    
    If itemCode = "" Then
        GetLatestSupplierInfo = result
        Exit Function
    End If
    
    On Error Resume Next
    Dim rs, sql
    
    ' Query SupplierPrices for latest active price
    sql = "SELECT TOP 1 sp.SupplierID, s.SupplierName, sp.UnitPrice " & _
          "FROM SupplierPrices sp LEFT JOIN Suppliers s ON sp.SupplierID = s.SupplierID " & _
          "WHERE sp.ItemCode = '" & SafeSQL(itemCode) & "' AND sp.IsActive = 1 " & _
          "ORDER BY sp.CreatedAt DESC"
    
    Set rs = conn.Execute(sql)
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            result(0) = CLng(rs("SupplierID"))
            result(1) = CDbl(rs("UnitPrice"))
            result(2) = rs("SupplierName") & ""
        End If
        rs.Close
    End If
    Set rs = Nothing
    
    ' Fallback: query PurchaseOrders history
    If result(0) = 0 Then
        sql = "SELECT TOP 1 po.SupplierID, s.SupplierName, pod.UnitPrice " & _
              "FROM PurchaseOrderDetails pod " & _
              "INNER JOIN PurchaseOrders po ON pod.PurchaseID = po.PurchaseID " & _
              "LEFT JOIN Suppliers s ON po.SupplierID = s.SupplierID " & _
              "WHERE pod.ItemCode = '" & SafeSQL(itemCode) & "' " & _
              "ORDER BY po.CreatedAt DESC"
        
        Set rs = conn.Execute(sql)
        If Not rs Is Nothing Then
            If Not rs.EOF Then
                result(0) = CLng(rs("SupplierID"))
                result(1) = CDbl(rs("UnitPrice"))
                result(2) = rs("SupplierName") & ""
            End If
            rs.Close
        End If
        Set rs = Nothing
    End If
    
    ' Fallback for Bottle: if BottleCode is NULL and itemCode is numeric ID
    If result(0) = 0 And orderType = "Bottle" And IsNumeric(itemCode) Then
        Dim bottleNameFound
        bottleNameFound = ""
        sql = "SELECT BottleName FROM BottleStyles WHERE BottleID=" & CLng(itemCode)
        Set rs = conn.Execute(sql)
        If Not rs Is Nothing Then
            If Not rs.EOF Then bottleNameFound = rs(0) & ""
            rs.Close
        End If
        Set rs = Nothing
        
        If bottleNameFound <> "" Then
            sql = "SELECT TOP 1 po.SupplierID, s.SupplierName, pod.UnitPrice " & _
                  "FROM PurchaseOrderDetails pod " & _
                  "INNER JOIN PurchaseOrders po ON pod.PurchaseID = po.PurchaseID " & _
                  "LEFT JOIN Suppliers s ON po.SupplierID = s.SupplierID " & _
                  "WHERE pod.ItemName = '" & SafeSQL(bottleNameFound) & "' AND po.OrderType = 'Bottle' " & _
                  "ORDER BY po.CreatedAt DESC"
            
            Set rs = conn.Execute(sql)
            If Not rs Is Nothing Then
                If Not rs.EOF Then
                    result(0) = CLng(rs("SupplierID"))
                    result(1) = CDbl(rs("UnitPrice"))
                    result(2) = rs("SupplierName") & ""
                End If
                rs.Close
            End If
            Set rs = Nothing
        End If
    End If
    
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    
    GetLatestSupplierInfo = result
End Function

' ========== Query low stock items by tab ==========
Function GetLowStockItems(orderType)
    Dim items, sql, rs
    items = ""
    
    On Error Resume Next
    Select Case orderType
        Case "RawMaterial"
            sql = "SELECT TOP 50 MaterialID AS ItemID, ItemName, ItemCode, StockQty, SafetyStock, Unit, UnitPrice AS WCost, " & _
                  "ISNULL(AvgDailyUsage,0) AS AvgDailyUsage, ISNULL(LeadTimeDays,7) AS LeadTimeDays, ISNULL(ReorderPoint,0) AS ReorderPoint " & _
                  "FROM RawMaterialInventory WHERE StockQty <= SafetyStock * 1.5 AND SafetyStock > 0 ORDER BY (StockQty - SafetyStock) ASC"
        Case "Packaging"
            sql = "SELECT TOP 50 PackagingID AS ItemID, ItemName, ItemCode, StockQty, SafetyStock, '' AS Unit, UnitPrice AS WCost, " & _
                  "ISNULL(AvgDailyUsage,0) AS AvgDailyUsage, ISNULL(LeadTimeDays,7) AS LeadTimeDays, ISNULL(ReorderPoint,0) AS ReorderPoint " & _
                  "FROM PackagingInventory WHERE StockQty <= SafetyStock * 1.5 AND SafetyStock > 0 ORDER BY (StockQty - SafetyStock) ASC"
        Case "Bottle"
            sql = "SELECT TOP 50 BottleID AS ItemID, BottleName AS ItemName, CAST(BottleID AS NVARCHAR(20)) AS ItemCode, StockQty, SafetyStock, 'pcs' AS Unit, UnitPrice AS WCost, " & _
                  "ISNULL(AvgDailyUsage,0) AS AvgDailyUsage, ISNULL(LeadTimeDays,7) AS LeadTimeDays, ISNULL(ReorderPoint,0) AS ReorderPoint " & _
                  "FROM BottleStyles WHERE StockQty <= SafetyStock * 1.5 AND SafetyStock > 0 ORDER BY (StockQty - SafetyStock) ASC"
        Case "Printing"
            sql = "SELECT TOP 50 PrintingID AS ItemID, ItemName, ItemCode, StockQty, SafetyStock, Unit, UnitPrice AS WCost, " & _
                  "ISNULL(AvgDailyUsage,0) AS AvgDailyUsage, ISNULL(LeadTimeDays,7) AS LeadTimeDays, ISNULL(ReorderPoint,0) AS ReorderPoint " & _
                  "FROM PrintingInventory WHERE StockQty <= SafetyStock * 1.5 AND SafetyStock > 0 ORDER BY (StockQty - SafetyStock) ASC"
        Case "SprayHead"
            sql = "SELECT TOP 50 SprayHeadID AS ItemID, ItemName, ItemCode, StockQty, SafetyStock, Unit, UnitPrice AS WCost, " & _
                  "ISNULL(AvgDailyUsage,0) AS AvgDailyUsage, ISNULL(LeadTimeDays,7) AS LeadTimeDays, ISNULL(ReorderPoint,0) AS ReorderPoint " & _
                  "FROM SprayHeadInventory WHERE StockQty <= SafetyStock * 1.5 AND SafetyStock > 0 ORDER BY (StockQty - SafetyStock) ASC"
        Case Else
            GetLowStockItems = ""
            Exit Function
    End Select
    
    Set rs = conn.Execute(sql)
    If Not rs Is Nothing Then
        Dim recCount : recCount = 0
        Do While Not rs.EOF And recCount < 50
            recCount = recCount + 1
            Dim itemID, itemName, itemCode, stockQty, safetyStock, unit, wCost
            Dim avgUsage, leadDays, reorderPoint
            itemID = SafeNum(rs("ItemID"))
            itemName = rs("ItemName") & ""
            itemCode = rs("ItemCode") & ""
            stockQty = SafeNum(rs("StockQty"))
            safetyStock = SafeNum(rs("SafetyStock"))
            unit = rs("Unit") & ""
            wCost = SafeNum(rs("WCost"))
            avgUsage = SafeNum(rs("AvgDailyUsage"))
            leadDays = SafeNum(rs("LeadTimeDays"))
            reorderPoint = SafeNum(rs("ReorderPoint"))
            
            If items <> "" Then items = items & vbCrLf
            items = items & Join(Array(itemID, itemName, itemCode, stockQty, safetyStock, unit, wCost, avgUsage, leadDays, reorderPoint), Chr(9))
            
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
    If Err.Number <> 0 Then
        items = ""
        Err.Clear
    End If
    On Error GoTo 0
    
    GetLowStockItems = items
End Function

' ========== Load supplier list for dropdown ==========
Dim rsSup, supOptions
supOptions = ""
On Error Resume Next
Set rsSup = conn.Execute("SELECT SupplierID, SupplierName, Category FROM Suppliers WHERE IsActive=1 ORDER BY SupplierName")
If Err.Number <> 0 Then Err.Clear : Set rsSup = Nothing
On Error GoTo 0
If Not rsSup Is Nothing Then
    Do While Not rsSup.EOF
        supOptions = supOptions & "<option value=""" & rsSup("SupplierID") & """ data-cat=""" & rsSup("Category") & """>" & Server.HTMLEncode(rsSup("SupplierName") & "") & "</option>"
        rsSup.MoveNext
    Loop
    rsSup.Close
End If
Set rsSup = Nothing

' ========== Load low stock data ==========
Dim lowStockData : lowStockData = GetLowStockItems(activeTab)
' Flush headers early so browser can start loading CSS/JS in parallel
Response.Flush

Function GetTypeLabel(ct)
    Select Case ct
        Case "RawMaterial" : GetTypeLabel = "原料"
        Case "Packaging"   : GetTypeLabel = "包装"
        Case "Bottle"      : GetTypeLabel = "瓶子"
        Case "Printing"    : GetTypeLabel = "印刷品"
        Case "SprayHead"   : GetTypeLabel = "喷头"
        Case Else          : GetTypeLabel = ct
    End Select
End Function
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>智能补货管理 - 采购中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --accent: #FF9800; --input-bg: #2d2d44; --card-bg: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: var(--bg); color: var(--text); min-height: 100vh; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-header { margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px; }
        .page-title { font-size: 24px; color: var(--text); display: flex; align-items: center; gap: 10px; }
        .page-title i { color: var(--accent); }
        .breadcrumb { font-size: 13px; color: #888; margin-bottom: 5px; }
        .breadcrumb a { color: var(--accent); text-decoration: none; }
        
        .card { background: var(--card-bg); border-radius: 12px; margin-bottom: 25px; border: 1px solid rgba(255,255,255,0.06); }
        .card-header { padding: 14px 20px; background: rgba(255,152,0,0.08); border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 15px; color: var(--text); display: flex; justify-content: space-between; align-items: center; }
        .card-body { padding: 16px 20px; overflow-x: auto; }
        
        /* Category Tabs */
        .cat-tabs { display: flex; gap: 4px; margin-bottom: 20px; background: rgba(255,255,255,0.03); padding: 4px; border-radius: 10px; }
        .cat-tab { padding: 10px 20px; border-radius: 8px; cursor: pointer; font-size: 14px; color: #888; transition: 0.2s; text-decoration: none; display: flex; align-items: center; gap: 6px; }
        .cat-tab:hover { color: #ccc; background: rgba(255,255,255,0.05); }
        .cat-tab.active { background: var(--accent); color: #fff; }
        .cat-tab i { font-size: 14px; }
        
        /* Item Table */
        table { width: 100%; border-collapse: collapse; font-size: 14px; }
        th { background: rgba(255,255,255,0.03); padding: 10px 12px; text-align: left; border-bottom: 2px solid rgba(255,255,255,0.08); color: #aaa; font-weight: 600; white-space: nowrap; }
        td { padding: 10px 12px; border-bottom: 1px solid rgba(255,255,255,0.04); }
        tr:hover { background: rgba(255,152,0,0.03); }
        tr.row-selected { background: rgba(255,152,0,0.08) !important; }
        
        .text-muted { color: #888; }
        .text-center { text-align: center; }
        .text-right { text-align: right; }
        
        .stock-low { color: #e74c3c; font-weight: 600; }
        .stock-warn { color: #FF9800; }
        .stock-ok { color: #4CAF50; }
        
        /* Replenishment Cart */
        .cart-section { margin-top: 20px; }
        .cart-item { display: flex; align-items: center; gap: 10px; padding: 12px; margin-bottom: 8px; background: rgba(255,255,255,0.03); border-radius: 8px; border: 1px solid rgba(255,255,255,0.06); flex-wrap: wrap; }
        .cart-item input, .cart-item select { background: var(--input-bg); border: 1px solid rgba(255,255,255,0.1); color: var(--text); padding: 7px 10px; border-radius: 6px; font-size: 13px; }
        .cart-item input:focus, .cart-item select:focus { border-color: var(--accent); outline: none; }
        .cart-item .item-info { flex: 2; min-width: 180px; }
        .cart-item .item-info .name { font-weight: 600; font-size: 14px; }
        .cart-item .item-info .code { font-size: 11px; color: #888; }
        .cart-item .qty-input { width: 80px; text-align: center; }
        .cart-item .price-input { width: 100px; text-align: right; }
        .cart-item .date-input { width: 130px; }
        .cart-item .notes-input { width: 150px; }
        .cart-remove { background: none; border: none; color: #e74c3c; cursor: pointer; font-size: 16px; padding: 4px 8px; }
        
        .alert { padding: 12px 18px; border-radius: 8px; margin-bottom: 15px; font-size: 14px; }
        .alert-success { background: rgba(76,175,80,0.15); color: #4CAF50; border-left: 3px solid #4CAF50; }
        .alert-error { background: rgba(231,76,60,0.15); color: #e74c3c; border-left: 3px solid #e74c3c; }
        
        .empty-state { text-align: center; padding: 40px; color: #888; }
        .empty-state i { font-size: 40px; margin-bottom: 10px; display: block; }
        
        .badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600; }
        .badge-danger { background: rgba(231,76,60,0.2); color: #e74c3c; }
        .badge-warn { background: rgba(255,152,0,0.2); color: #FF9800; }
        
        .totals-bar { display: flex; justify-content: space-between; align-items: center; padding: 14px 20px; background: rgba(255,152,0,0.08); border-radius: 8px; margin-top: 12px; }
    </style>
</head>
<body data-theme="purchase-dark">
    <!--#include file="includes/nav.asp"-->
    <div class="main-content">
        <div class="breadcrumb">
            <a href="index.asp"><i class="fas fa-home"></i> 采购中心</a> / 智能补货管理
        </div>
        <div class="page-header">
            <div class="page-title"><i class="fas fa-robot"></i> 智能补货管理</div>
            <div>
                <a href="purchase_orders.asp" class="btn btn-outline"><i class="fas fa-list"></i> 采购订单列表</a>
            </div>
        </div>
        
        <% If genMsg <> "" Then %>
        <div class="alert alert-<%=genMsgType%>">
            <i class="fas fa-<%=IIF(genMsgType="success","check-circle","exclamation-circle")%>"></i> <%=Server.HTMLEncode(genMsg)%>
        </div>
        <% End If %>
        
        <!-- Category Tabs -->
        <div class="cat-tabs" id="catTabs">
            <a href="javascript:switchTab('RawMaterial')" class="cat-tab <%=IIF(activeTab="RawMaterial","active","")%>" data-tab="RawMaterial"><i class="fas fa-flask"></i> 原料</a>
            <a href="javascript:switchTab('Packaging')" class="cat-tab <%=IIF(activeTab="Packaging","active","")%>" data-tab="Packaging"><i class="fas fa-box"></i> 包装</a>
            <a href="javascript:switchTab('Bottle')" class="cat-tab <%=IIF(activeTab="Bottle","active","")%>" data-tab="Bottle"><i class="fas fa-wine-bottle"></i> 瓶子</a>
            <a href="javascript:switchTab('Printing')" class="cat-tab <%=IIF(activeTab="Printing","active","")%>" data-tab="Printing"><i class="fas fa-print"></i> 印刷品</a>
            <a href="javascript:switchTab('SprayHead')" class="cat-tab <%=IIF(activeTab="SprayHead","active","")%>" data-tab="SprayHead"><i class="fas fa-spray-can"></i> 喷头</a>
        </div>
        
        <!-- Low Stock Items Table -->
        <div class="card">
            <div class="card-header">
                <span><i class="fas fa-exclamation-triangle" style="color:#FF9800;"></i> <%=GetTypeLabel(activeTab)%> 低库存预警</span>
                <span class="text-muted" style="font-weight:400;font-size:12px;">勾选需补货项目 → 自动填充历史供应商和价格</span>
            </div>
            <div class="card-body">
                <table id="stockTable">
                    <thead>
                        <tr>
                            <th style="width:40px;"><input type="checkbox" id="selectAll" onchange="toggleSelectAll(this)" title="全选"></th>
                            <th>物料名称</th>
                            <th>编码</th>
                            <th>当前库存</th>
                            <th>安全库存</th>
                            <th>日均消耗</th>
                            <th>交货周期</th>
                            <th>加权成本</th>
                            <th>历史供应商</th>
                            <th>历史单价</th>
                            <th>补货紧迫度</th>
                            <th>补货建议</th>
                        </tr>
                    </thead>
                    <tbody id="stockTableBody">
                    <%
                    If lowStockData <> "" Then
                        Dim rows, ri, rowData
                        rows = Split(lowStockData, vbCrLf)
                        For ri = 0 To UBound(rows)
                            If Trim(rows(ri)) <> "" Then
                                rowData = Split(rows(ri), Chr(9))
                                If UBound(rowData) >= 9 Then
                                    Dim rItemID, rItemName, rItemCode, rStock, rSafety, rUnit, rWCost, rSupID, rSupPrice, rSupName
                                    Dim rAvgUsage, rLeadDays, rReorderPoint
                                    rItemID = rowData(0)
                                    rItemName = rowData(1)
                                    rItemCode = rowData(2)
                                    rStock = CDbl(rowData(3))
                                    rSafety = CDbl(rowData(4))
                                    rUnit = rowData(5)
                                    rWCost = CDbl(rowData(6))
                                    rAvgUsage = CDbl(rowData(7))
                                    rLeadDays = CDbl(rowData(8))
                                    rReorderPoint = CDbl(rowData(9))
                                    If rLeadDays <= 0 Then rLeadDays = 7
                                    ' Supplier info loaded asynchronously via AJAX
                                    rSupID = 0
                                    rSupPrice = 0
                                    rSupName = ""
                                    
                                    ' V11: 升级补货算法 - 基于日均消耗和交货周期
                                    Dim stockClass, suggestQty, urgencyLevel
                                    Dim leadDemand : leadDemand = rAvgUsage * rLeadDays * 1.2
                                    Dim demandBased : demandBased = leadDemand + rSafety - rStock
                                    Dim safetyBased : safetyBased = rSafety * 1.5 - rStock
                                    suggestQty = demandBased
                                    If safetyBased > suggestQty Then suggestQty = safetyBased
                                    suggestQty = Round(suggestQty, 0)
                                    If suggestQty < 1 Then suggestQty = 1
                                    
                                    ' 紧急程度分级
                                    If rStock <= 0 Then
                                        stockClass = "stock-low" : urgencyLevel = "<span class='badge badge-danger'><i class='fas fa-exclamation-circle'></i> 紧急</span>"
                                    ElseIf rStock < rSafety * 0.5 Then
                                        stockClass = "stock-low" : urgencyLevel = "<span class='badge badge-danger'>严重不足</span>"
                                    ElseIf rStock < rSafety Then
                                        stockClass = "stock-low" : urgencyLevel = "<span class='badge badge-warn'>低于安全库存</span>"
                                    ElseIf rAvgUsage > 0 And rReorderPoint > 0 And rStock < rReorderPoint Then
                                        stockClass = "stock-warn" : urgencyLevel = "<span class='badge badge-warn'>低于再订货点</span>"
                                    Else
                                        stockClass = "stock-warn" : urgencyLevel = "<span class='badge badge-warn'>库存偏低</span>"
                                    End If
                    %>
                        <tr data-item-id="<%=rItemID%>" data-item-name="<%=Server.HTMLEncode(rItemName)%>" data-item-code="<%=Server.HTMLEncode(rItemCode)%>" 
                            data-stock="<%=rStock%>" data-safety="<%=rSafety%>" data-unit="<%=Server.HTMLEncode(rUnit)%>" 
                            data-wcost="<%=rWCost%>" data-avgusage="<%=rAvgUsage%>" data-leaddays="<%=rLeadDays%>"
                            data-supid="<%=rSupID%>" data-supname="<%=Server.HTMLEncode(rSupName)%>" data-supprice="<%=rSupPrice%>">
                            <td><input type="checkbox" class="row-check" onchange="onRowCheck(this)"></td>
                            <td><strong><%=Server.HTMLEncode(rItemName)%></strong></td>
                            <td style="color:#888;"><%=Server.HTMLEncode(rItemCode)%></td>
                            <td class="<%=stockClass%>"><%=FormatNumber(rStock, 1)%></td>
                            <td><%=FormatNumber(rSafety, 1)%></td>
                            <td><% If rAvgUsage > 0 Then %><%= FormatNumber(rAvgUsage, 2) %><% Else %><span style='color:#888;'>--</span><% End If %></td>
                            <td><%=CLng(rLeadDays)%> 天</td>
                            <td>¥<%=FormatNumber(rWCost, 4)%></td>
                            <td>
                                <span class="supplier-loading" style="color:#999;"><i class="fas fa-spinner fa-pulse"></i> 加载中…</span>
                            </td>
                            <td class="supplier-price">-</td>
                            <td><%=urgencyLevel%></td>
                            <td><span class="badge <%=IIF(rStock<=0,"badge-danger","badge-warn")%>">建议补 <%=suggestQty%></span></td>
                        </tr>
                    <%
                                End If
                            End If
                        Next
                    Else
                    %>
                        <tr><td colspan="12" class="empty-state">
                            <i class="fas fa-check-circle" style="color:#4CAF50;"></i>
                            该品类库存充足，无需补货
                        </td></tr>
                    <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
        
        <!-- V11: 非预警产品补充 -->
        <div class="card" style="margin-top:20px;">
            <div class="card-header">
                <span><i class="fas fa-search" style="color:#4CAF50;"></i> 非预警产品补充 - 搜索库存充足但需要采购的产品</span>
            </div>
            <div class="card-body">
                <div style="display:flex;gap:10px;margin-bottom:12px;">
                    <input type="text" id="productSearch" placeholder="输入产品名称或编码搜索..." style="flex:2;padding:10px 15px;border:1px solid rgba(255,255,255,0.1);border-radius:8px;background:#1e1e32;color:#e0e0e0;font-size:14px;" onkeyup="if(event.key==='Enter')searchProducts()">
                    <button type="button" class="btn btn-sm" style="background:linear-gradient(135deg,#4CAF50 0%,#388E3C 100%);color:#fff;padding:10px 20px;border:none;border-radius:8px;cursor:pointer;" onclick="searchProducts()"><i class="fas fa-search"></i> 搜索</button>
                </div>
                <div id="productSearchResults" style="max-height:300px;overflow-y:auto;">
                    <div style="text-align:center;color:#666;padding:20px;">输入关键词并点击搜索，查找当前<%=GetTypeLabel(activeTab)%>品类下的产品</div>
                </div>
            </div>
        </div>
        
        <!-- Replenishment Cart -->
        <div class="card cart-section" id="cartSection" style="display:none;">
            <div class="card-header">
                <span><i class="fas fa-shopping-cart"></i> 补货清单 (<span id="cartCount">0</span> 项)</span>
                <button class="btn btn-sm btn--neutral" onclick="clearCart()"><i class="fas fa-trash"></i> 清空</button>
            </div>
            <div class="card-body">
                <form method="post" id="genForm">
                    <input type="hidden" name="action" value="generate">
                    <%=GetCSRFTokenField()%>
                    <input type="hidden" name="gen_count" id="genCount" value="0">
                    <div id="cartItems"></div>
                    
                    <div class="totals-bar" id="cartTotals" style="display:none;">
                        <div>
                            <span class="text-muted">合计数量: </span><strong id="totalQty" style="color:#4CAF50;">0</strong>
                            <span class="text-muted" style="margin-left:20px;">合计金额: </span><strong id="totalAmount" style="color:#FF9800;">¥0.00</strong>
                        </div>
                        <button type="submit" class="btn btn-success" onclick="return confirm('确认生成补货采购订单？系统将按供应商自动分组合并。')">
                            <i class="fas fa-magic"></i> 生成补货订单
                        </button>
                    </div>
                </form>
            </div>
        </div>
    </div>
    
    <script>
    var cartItems = [];
    var currentTab = '<%=activeTab%>';
    var activeXhr = null;  // Track active AJAX request for abort on tab switch
    
    // Switch tab via AJAX - no page reload, no session lock
    function switchTab(tabName) {
        if (tabName === currentTab) return;
        
        // Abort any in-flight AJAX
        if (activeXhr) { try { activeXhr.abort(); } catch(e) {} }
        
        // Update tab UI
        document.querySelectorAll('.cat-tab').forEach(function(el) {
            el.classList.toggle('active', el.getAttribute('data-tab') === tabName);
        });
        
        // Update header
        var labelMap = {'RawMaterial':'原料','Packaging':'包装','Bottle':'瓶子','Printing':'印刷品','SprayHead':'喷头'};
        var label = labelMap[tabName] || tabName;
        var headerEl = document.querySelector('.card-header span');
        if (headerEl) headerEl.innerHTML = '<i class="fas fa-exclamation-triangle" style="color:#FF9800;"></i> ' + escapeHtml(label) + ' 低库存预警';
        
        currentTab = tabName;
        clearCart();
        document.getElementById('stockTableBody').innerHTML = '<tr><td colspan="10" class="empty-state"><i class="fas fa-spinner fa-pulse"></i> 加载中…</td></tr>';
        
        // Step 1: Load low-stock data
        var xhr = new XMLHttpRequest();
        xhr.open('GET', 'ajax_lowstock_data.asp?tab=' + encodeURIComponent(tabName), true);
        xhr.onload = function() {
            if (xhr.status !== 200) return;
            try {
                var items = JSON.parse(xhr.responseText);
                renderTable(items);
                // Step 2: Load supplier info (async, after table renders)
                setTimeout(function() { loadSupplierInfoForRows(); }, 50);
            } catch(e) {}
        };
        xhr.send();
    }
    
    function renderTable(items) {
        var tbody = document.getElementById('stockTableBody');
        if (!items || items.length === 0) {
            tbody.innerHTML = '<tr><td colspan="12" class="empty-state"><i class="fas fa-check-circle" style="color:#4CAF50;"></i>该品类库存充足，无需补货</td></tr>';
            return;
        }
        
        var html = '';
        items.forEach(function(item) {
            var stockClass = item.stock <= 0 ? 'stock-low' : (item.stock < item.safety ? 'stock-low' : 'stock-warn');
            var badgeClass = item.stock <= 0 ? 'badge-danger' : 'badge-warn';
            var avgUsageHtml = item.avgusage > 0 ? item.avgusage.toFixed(2) : '<span style="color:#888;">--</span>';
            var leadDays = item.leaddays > 0 ? item.leaddays : 7;
            
            // urgency level
            var urgencyHtml = '';
            if (item.stock <= 0) {
                urgencyHtml = '<span class="badge badge-danger"><i class="fas fa-exclamation-circle"></i> 紧急</span>';
            } else if (item.stock < item.safety * 0.5) {
                urgencyHtml = '<span class="badge badge-danger">严重不足</span>';
            } else if (item.stock < item.safety) {
                urgencyHtml = '<span class="badge badge-warn">低于安全库存</span>';
            } else {
                urgencyHtml = '<span class="badge badge-warn">库存偏低</span>';
            }
            
            html += '<tr data-item-id="' + item.itemid + '" data-item-name="' + escapeHtml(item.itemname) + '" data-item-code="' + escapeHtml(item.itemcode) + '" ' +
                'data-stock="' + item.stock + '" data-safety="' + item.safety + '" data-unit="' + escapeHtml(item.unit) + '" ' +
                'data-wcost="' + item.wcost + '" data-avgusage="' + (item.avgusage || 0) + '" data-leaddays="' + leadDays + '" ' +
                'data-supid="0" data-supname="" data-supprice="0">' +
                '<td><input type="checkbox" class="row-check" onchange="onRowCheck(this)"></td>' +
                '<td><strong>' + escapeHtml(item.itemname) + '</strong></td>' +
                '<td style="color:#888;">' + escapeHtml(item.itemcode) + '</td>' +
                '<td class="' + stockClass + '">' + item.stock + '</td>' +
                '<td>' + item.safety + '</td>' +
                '<td>' + avgUsageHtml + '</td>' +
                '<td>' + leadDays + ' 天</td>' +
                '<td>¥' + parseFloat(item.wcost).toFixed(4) + '</td>' +
                '<td><span class="supplier-loading" style="color:#999;"><i class="fas fa-spinner fa-pulse"></i> 加载中…</span></td>' +
                '<td class="supplier-price">-</td>' +
                '<td>' + urgencyHtml + '</td>' +
                '<td><span class="badge ' + badgeClass + '">建议补 ' + item.suggest + '</span></td>' +
                '</tr>';
        });
        tbody.innerHTML = html;
        document.getElementById('selectAll').checked = false;
    }
    
    function loadSupplierInfoForRows() {
        loadSupplierInfo();
    }
    
    function toggleSelectAll(checkbox) {
        var checks = document.querySelectorAll('.row-check');
        checks.forEach(function(cb) {
            cb.checked = checkbox.checked;
            onRowCheck(cb);
        });
    }
    
    function onRowCheck(checkbox) {
        var row = checkbox.closest('tr');
        var itemId = row.getAttribute('data-item-id');
        row.classList.toggle('row-selected', checkbox.checked);
        
        if (checkbox.checked) {
            addToCart(row);
        } else {
            removeFromCart(itemId);
        }
        updateCart();
    }
    
    function addToCart(row) {
        var itemId = row.getAttribute('data-item-id');
        if (cartItems.find(function(i) { return i.id === itemId; })) return;
        
        var itemName = row.getAttribute('data-item-name');
        var itemCode = row.getAttribute('data-item-code');
        var stock = parseFloat(row.getAttribute('data-stock'));
        var safety = parseFloat(row.getAttribute('data-safety'));
        var unit = row.getAttribute('data-unit');
        var wcost = parseFloat(row.getAttribute('data-wcost'));
        var supId = row.getAttribute('data-supid');
        var supName = row.getAttribute('data-supname');
        var supPrice = parseFloat(row.getAttribute('data-supprice'));
        
        var suggestQty = safety > stock ? Math.round(safety - stock + safety * 0.3) : Math.round(safety * 1.5 - stock);
        if (suggestQty < 1) suggestQty = 1;
        var price = supPrice > 0 ? supPrice : (wcost > 0 ? wcost : 0);
        
        cartItems.push({
            id: itemId, name: itemName, code: itemCode, stock: stock, safety: safety,
            unit: unit, wcost: wcost, supId: supId, supName: supName, supPrice: supPrice,
            qty: suggestQty, price: price,
            delivery: '', notes: ''
        });
    }
    
    function removeFromCart(itemId) {
        cartItems = cartItems.filter(function(i) { return i.id !== itemId; });
    }
    
    function clearCart() {
        cartItems = [];
        document.querySelectorAll('.row-check').forEach(function(cb) { cb.checked = false; });
        document.querySelectorAll('tr.row-selected').forEach(function(r) { r.classList.remove('row-selected'); });
        document.getElementById('selectAll').checked = false;
        updateCart();
    }
    
    function updateCart() {
        var cartSection = document.getElementById('cartSection');
        var cartItemsDiv = document.getElementById('cartItems');
        var cartCount = document.getElementById('cartCount');
        var cartTotals = document.getElementById('cartTotals');
        var genCount = document.getElementById('genCount');
        
        cartCount.textContent = cartItems.length;
        genCount.value = cartItems.length;
        
        if (cartItems.length === 0) {
            cartSection.style.display = 'none';
            return;
        }
        
        cartSection.style.display = 'block';
        cartTotals.style.display = 'flex';
        
        var html = '';
        var totalQty = 0, totalAmount = 0;
        var activeTab = currentTab;  // Use JS variable to support dynamic tab switching
        
        cartItems.forEach(function(item, idx) {
            var itemTotal = item.qty * item.price;
            totalQty += item.qty;
            totalAmount += itemTotal;
            var idx1 = idx + 1;
            
            html += '<div class="cart-item">' +
                '<input type="hidden" name="gen_type_' + idx1 + '" value="' + activeTab + '">' +
                '<input type="hidden" name="gen_name_' + idx1 + '" value="' + escapeHtml(item.name) + '">' +
                '<input type="hidden" name="gen_code_' + idx1 + '" value="' + escapeHtml(item.code) + '">' +
                '<div class="item-info">' +
                    '<div class="name">' + escapeHtml(item.name) + '</div>' +
                    '<div class="code">' + escapeHtml(item.code) + ' | 库存:' + item.stock.toFixed(1) + item.unit + '</div>' +
                '</div>' +
                '<span class="text-muted" style="font-size:12px;">数量</span>' +
                '<input type="number" class="qty-input" name="gen_qty_' + idx1 + '" value="' + item.qty + '" min="1" step="1" onchange="updateCartItem(' + idx + ', \'qty\', this.value)">' +
                '<span class="text-muted" style="font-size:12px;">单价</span>' +
                '<input type="number" class="price-input" name="gen_price_' + idx1 + '" value="' + item.price.toFixed(2) + '" step="0.01" onchange="updateCartItem(' + idx + ', \'price\', this.value)">' +
                '<span class="text-muted" style="font-size:12px;">交货日</span>' +
                '<input type="date" class="date-input" name="gen_delivery_' + idx1 + '" value="' + item.delivery + '" onchange="updateCartItem(' + idx + ', \'delivery\', this.value)">' +
                '<select name="gen_supplier_' + idx1 + '" onchange="updateCartItem(' + idx + ', \'supId\', this.value)" style="width:140px;">' +
                    '<option value="">选择供应商</option>' +
                    '<%=supOptions%>' +
                '</select>' +
                '<input type="text" class="notes-input" name="gen_notes_' + idx1 + '" placeholder="备注" value="' + escapeHtml(item.notes) + '" onchange="updateCartItem(' + idx + ', \'notes\', this.value)">' +
                '<button type="button" class="cart-remove" onclick="removeCartItem(' + idx + ')" title="移除"><i class="fas fa-times"></i></button>' +
                '</div>';
        });
        
        cartItemsDiv.innerHTML = html;
        document.getElementById('totalQty').textContent = totalQty;
        document.getElementById('totalAmount').textContent = '¥' + totalAmount.toFixed(2);
        
        // Set supplier dropdown values
        cartItems.forEach(function(item, idx) {
            var supSelect = document.querySelector('[name="gen_supplier_' + (idx+1) + '"]');
            if (supSelect && item.supId) {
                supSelect.value = item.supId;
            }
        });
    }
    
    function updateCartItem(idx, field, value) {
        if (idx >= 0 && idx < cartItems.length) {
            if (field === 'qty' || field === 'price') {
                cartItems[idx][field] = parseFloat(value) || 0;
            } else if (field === 'supId') {
                cartItems[idx].supId = value;
            } else {
                cartItems[idx][field] = value;
            }
            updateCart();
        }
    }
    
    function removeCartItem(idx) {
        if (idx >= 0 && idx < cartItems.length) {
            var itemId = cartItems[idx].id;
            cartItems.splice(idx, 1);
            // Uncheck the corresponding row
            var rows = document.querySelectorAll('#stockTableBody tr');
            rows.forEach(function(row) {
                if (row.getAttribute('data-item-id') === itemId) {
                    var cb = row.querySelector('.row-check');
                    if (cb) cb.checked = false;
                    row.classList.remove('row-selected');
                }
            });
            updateCart();
        }
    }
    
    function escapeHtml(str) {
        if (!str) return '';
        return str.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/'/g, '&#39;');
    }
    
    // Load supplier info via AJAX on page load (batch request to avoid session lock)
    function loadSupplierInfo() {
        var rows = document.querySelectorAll('#stockTableBody tr');
        var tab = currentTab;  // Use JS variable instead of ASP to support dynamic tab switching
        var itemCodes = [];
        var rowMap = {};  // itemCode -> row reference
        
        rows.forEach(function(row) {
            var itemCode = row.getAttribute('data-item-code');
            if (!itemCode) return;
            itemCodes.push(itemCode);
            rowMap[itemCode] = row;
        });
        
        if (itemCodes.length === 0) return;
        
        // Single batch request instead of N individual requests
        var xhr = new XMLHttpRequest();
        xhr.open('POST', 'ajax_supplier_info_batch.asp', true);
        xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        
        xhr.onload = function() {
            if (xhr.status === 200) {
                try {
                    var results = JSON.parse(xhr.responseText);
                    results.forEach(function(data) {
                        var row = rowMap[data.itemcode];
                        if (!row) return;
                        
                        var cells = row.querySelectorAll('td');
                        
                        if (data.supplier_id > 0) {
                            row.setAttribute('data-supid', data.supplier_id);
                            row.setAttribute('data-supname', data.supplier_name);
                            row.setAttribute('data-supprice', data.unit_price);
                            
                            if (cells.length >= 9) {
                                cells[7].innerHTML = '<span style="color:#4CAF50;"><i class="fas fa-check-circle"></i> ' + escapeHtml(data.supplier_name) + '</span>';
                                cells[8].innerHTML = '¥' + parseFloat(data.unit_price).toFixed(2);
                            }
                            
                            // If row is already checked, update cart item supplier info
                            var itemId = row.getAttribute('data-item-id');
                            for (var ci = 0; ci < cartItems.length; ci++) {
                                if (cartItems[ci].id === itemId) {
                                    cartItems[ci].supId = data.supplier_id;
                                    cartItems[ci].supName = data.supplier_name;
                                    cartItems[ci].supPrice = data.unit_price;
                                    if (parseFloat(cartItems[ci].price) <= 0) {
                                        cartItems[ci].price = data.unit_price;
                                    }
                                    break;
                                }
                            }
                        } else {
                            // No supplier found
                            if (cells.length >= 9) {
                                cells[7].innerHTML = '<span class="text-muted">无记录</span>';
                                cells[8].innerHTML = '-';
                            }
                        }
                    });
                    updateCart();
                } catch(e) {}
            }
        };
        
        xhr.send('itemcodes=' + encodeURIComponent(itemCodes.join(',')) + '&ordertype=' + encodeURIComponent(tab));
    }
    
    // Run on page load
    document.addEventListener('DOMContentLoaded', function() {
        loadSupplierInfo();
    });
    
    // ========== V11: 非预警产品搜索 ==========
    function searchProducts() {
        var searchText = document.getElementById('productSearch').value.trim();
        var resultsDiv = document.getElementById('productSearchResults');
        if (!searchText) {
            resultsDiv.innerHTML = '<div style="text-align:center;color:#666;padding:20px;">请至少输入一个关键词</div>';
            return;
        }
        
        resultsDiv.innerHTML = '<div style="text-align:center;color:#666;padding:20px;"><i class="fas fa-spinner fa-pulse"></i> 搜索中...</div>';
        
        var xhr = new XMLHttpRequest();
        xhr.open('GET', 'ajax_search_products.asp?ordertype=' + encodeURIComponent(currentTab) + '&search=' + encodeURIComponent(searchText), true);
        xhr.onload = function() {
            if (xhr.status !== 200) {
                resultsDiv.innerHTML = '<div style="text-align:center;color:#F44336;padding:20px;">搜索失败</div>';
                return;
            }
            try {
                var items = JSON.parse(xhr.responseText);
                renderSearchResults(items);
            } catch(e) {
                resultsDiv.innerHTML = '<div style="text-align:center;color:#F44336;padding:20px;">数据解析失败</div>';
            }
        };
        xhr.onerror = function() {
            resultsDiv.innerHTML = '<div style="text-align:center;color:#F44336;padding:20px;">网络错误</div>';
        };
        xhr.send();
    }
    
    function renderSearchResults(items) {
        var resultsDiv = document.getElementById('productSearchResults');
        if (!items || items.length === 0) {
            resultsDiv.innerHTML = '<div style="text-align:center;color:#666;padding:20px;"><i class="fas fa-search"></i> 暂无匹配产品</div>';
            return;
        }
        
        var html = '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
        html += '<thead><tr style="background:rgba(255,255,255,0.03);"><th style="padding:8px;text-align:left;color:#aaa;">产品名称</th><th style="padding:8px;text-align:left;color:#aaa;">编码</th><th style="padding:8px;text-align:right;color:#aaa;">库存</th><th style="padding:8px;text-align:right;color:#aaa;">安全库存</th><th style="padding:8px;text-align:right;color:#aaa;">成本价</th><th style="padding:8px;text-align:center;color:#aaa;">操作</th></tr></thead><tbody>';
        
        items.forEach(function(item) {
            var stockClass = item.stock <= item.safety ? 'stock-low' : 'stock-ok';
            html += '<tr style="border-bottom:1px solid rgba(255,255,255,0.04);">';
            html += '<td style="padding:8px;"><strong>' + escapeHtml(item.itemname) + '</strong></td>';
            html += '<td style="padding:8px;color:#888;">' + escapeHtml(item.itemcode) + '</td>';
            html += '<td style="padding:8px;text-align:right;" class="' + stockClass + '">' + item.stock.toFixed(1) + '</td>';
            html += '<td style="padding:8px;text-align:right;">' + item.safety.toFixed(1) + '</td>';
            html += '<td style="padding:8px;text-align:right;">¥' + parseFloat(item.wcost).toFixed(4) + '</td>';
            html += '<td style="padding:8px;text-align:center;"><button type="button" class="btn btn-sm" style="background:linear-gradient(135deg,#FF9800 0%,#F57C00 100%);color:#fff;border:none;border-radius:6px;padding:6px 12px;cursor:pointer;" onclick="addToCartFromSearch(\'' + escapeHtml(item.itemid) + '\',\'' + escapeHtml(item.itemname).replace(/'/g, "\\'") + '\',\'' + escapeHtml(item.itemcode).replace(/'/g, "\\'") + '\',' + item.stock + ',' + item.safety + ',\'' + escapeHtml(item.unit).replace(/'/g, "\\'") + '\',' + item.wcost + ')">+ 加入补货</button></td>';
            html += '</tr>';
        });
        
        html += '</tbody></table>';
        html += '<div style="text-align:center;color:#888;padding:8px;font-size:12px;">找到 ' + items.length + ' 个结果</div>';
        resultsDiv.innerHTML = html;
    }
    
    function addToCartFromSearch(itemId, itemName, itemCode, stock, safety, unit, wcost) {
        // Check if already in cart
        if (cartItems.find(function(i) { return i.id === itemId; })) {
            alert('该产品已在补货清单中');
            return;
        }
        
        var suggestQty = safety > stock ? Math.round(safety - stock + safety * 0.3) : Math.round(safety * 1.5 - stock);
        if (suggestQty < 1) suggestQty = 1;
        
        cartItems.push({
            id: itemId, name: itemName, code: itemCode, stock: stock, safety: safety,
            unit: unit, wcost: wcost, supId: 0, supName: '', supPrice: 0,
            qty: suggestQty, price: wcost > 0 ? wcost : 0,
            delivery: '', notes: ''
        });
        
        updateCart();
        
        // Flash feedback
        var resultsDiv = document.getElementById('productSearchResults');
        var oldHtml = resultsDiv.innerHTML;
        var bar = document.createElement('div');
        bar.style.cssText = 'background:rgba(76,175,80,0.2);color:#4CAF50;padding:10px;border-radius:6px;margin-bottom:10px;text-align:center;';
        bar.innerHTML = '<i class="fas fa-check-circle"></i> "' + escapeHtml(itemName) + '" 已加入补货清单';
        resultsDiv.insertBefore(bar, resultsDiv.firstChild);
        setTimeout(function() { bar.remove(); }, 2000);
        
        // Scroll cart into view
        document.getElementById('cartSection').scrollIntoView({behavior: 'smooth'});
    }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
