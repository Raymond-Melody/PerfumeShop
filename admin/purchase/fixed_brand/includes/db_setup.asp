<%
' ============================================
' 品牌定香采购模块 - 数据库自动迁移
' 在所有页面调用 OpenConnection() 后 include 此文件
' ============================================
Dim dbSetupOK : dbSetupOK = True

' ----- FixedBrandProducts -----
If Not TableExists("FixedBrandProducts") Then
    dbSetupOK = CreateFixedBrandTable("FixedBrandProducts", _
        "FixedProductID INT IDENTITY(1,1) PRIMARY KEY, " & _
        "ProductID INT, ProductCode NVARCHAR(50), ProductName NVARCHAR(200) NOT NULL, " & _
        "Specification NVARCHAR(100), UnitPrice DECIMAL(19,4) DEFAULT 0, SalePrice DECIMAL(19,4) DEFAULT 0, " & _
        "SupplierID INT, SupplierName NVARCHAR(200), MinOrderQty INT DEFAULT 1, LeadTimeDays INT DEFAULT 7, " & _
        "ImageURL NVARCHAR(500), Status NVARCHAR(20) DEFAULT 'Active', " & _
        "CreatedAt DATETIME DEFAULT GETDATE(), UpdatedAt DATETIME DEFAULT GETDATE()") AND dbSetupOK
End If

' ----- FixedBrandPurchaseOrders -----
If Not TableExists("FixedBrandPurchaseOrders") Then
    dbSetupOK = CreateFixedBrandTable("FixedBrandPurchaseOrders", _
        "PurchaseID INT IDENTITY(1,1) PRIMARY KEY, PurchaseNo NVARCHAR(50), " & _
        "SupplierID INT, SupplierName NVARCHAR(200), TotalAmount DECIMAL(19,4) DEFAULT 0, " & _
        "Status NVARCHAR(30) DEFAULT 'Draft', OrderDate DATETIME DEFAULT GETDATE(), " & _
        "ExpectedDate DATETIME, ApprovedBy NVARCHAR(100), ApprovedAt DATETIME, " & _
        "Remarks NVARCHAR(500), CreatedBy NVARCHAR(100), " & _
        "CreatedAt DATETIME DEFAULT GETDATE(), UpdatedAt DATETIME DEFAULT GETDATE()") AND dbSetupOK
End If

' ----- FixedBrandPurchaseDetails -----
If Not TableExists("FixedBrandPurchaseDetails") Then
    dbSetupOK = CreateFixedBrandTable("FixedBrandPurchaseDetails", _
        "DetailID INT IDENTITY(1,1) PRIMARY KEY, PurchaseID INT, FixedProductID INT, " & _
        "ProductName NVARCHAR(200), Specification NVARCHAR(100), Quantity INT DEFAULT 0, " & _
        "ReceivedQty INT DEFAULT 0, UnitPrice DECIMAL(19,4) DEFAULT 0, SubTotal DECIMAL(19,4) DEFAULT 0, " & _
        "ExpectedDate DATETIME, Remarks NVARCHAR(500)") AND dbSetupOK
End If

' ----- FixedBrandInventory -----
If Not TableExists("FixedBrandInventory") Then
    dbSetupOK = CreateFixedBrandTable("FixedBrandInventory", _
        "InventoryID INT IDENTITY(1,1) PRIMARY KEY, FixedProductID INT, " & _
        "ProductCode NVARCHAR(50), ProductName NVARCHAR(200), Specification NVARCHAR(100), " & _
        "StockQty INT DEFAULT 0, SafetyStock INT DEFAULT 10, MinOrderQty INT DEFAULT 1, " & _
        "AvgUnitCost DECIMAL(19,4) DEFAULT 0, LastPurchasePrice DECIMAL(19,4) DEFAULT 0, " & _
        "LastPurchaseDate DATETIME, LastPurchaseID INT, TotalPurchased INT DEFAULT 0, " & _
        "TotalSold INT DEFAULT 0, ParamMode NVARCHAR(20) DEFAULT 'Manual', " & _
        "DailySalesAvg DECIMAL(10,2) DEFAULT 0, ConsecutiveDataMonths INT DEFAULT 0, " & _
        "LastAutoCalcDate DATETIME, UpdatedAt DATETIME DEFAULT GETDATE()") AND dbSetupOK
End If

' ----- FixedBrandReceipts -----
If Not TableExists("FixedBrandReceipts") Then
    dbSetupOK = CreateFixedBrandTable("FixedBrandReceipts", _
        "ReceiptID INT IDENTITY(1,1) PRIMARY KEY, PurchaseID INT, " & _
        "ReceiptNo NVARCHAR(50), SupplierID INT, ReceivedBy NVARCHAR(100), " & _
        "ReceiptDate DATETIME DEFAULT GETDATE(), TotalReceivedQty INT DEFAULT 0, " & _
        "Notes NVARCHAR(500), CreatedAt DATETIME DEFAULT GETDATE()") AND dbSetupOK
End If

' ----- FixedBrandReceiptDetails -----
If Not TableExists("FixedBrandReceiptDetails") Then
    dbSetupOK = CreateFixedBrandTable("FixedBrandReceiptDetails", _
        "ReceiptDetailID INT IDENTITY(1,1) PRIMARY KEY, ReceiptID INT, " & _
        "DetailID INT, FixedProductID INT, AcceptedQty INT DEFAULT 0, " & _
        "RejectedQty INT DEFAULT 0, RejectReason NVARCHAR(500), UnitPrice DECIMAL(19,4) DEFAULT 0") AND dbSetupOK
End If

' ----- FixedBrandCostAllocation -----
If Not TableExists("FixedBrandCostAllocation") Then
    dbSetupOK = CreateFixedBrandTable("FixedBrandCostAllocation", _
        "AllocationID INT IDENTITY(1,1) PRIMARY KEY, OrderID INT, OrderNo NVARCHAR(100), " & _
        "PurchaseID INT, PurchaseNo NVARCHAR(50), FixedProductID INT, ProductName NVARCHAR(200), " & _
        "CostPerUnit DECIMAL(19,4) DEFAULT 0, Quantity INT DEFAULT 0, TotalCost DECIMAL(19,4) DEFAULT 0, " & _
        "SalePrice DECIMAL(19,4) DEFAULT 0, ProfitAmount DECIMAL(19,4) DEFAULT 0, " & _
        "ProfitRate DECIMAL(10,4) DEFAULT 0, AllocatedAt DATETIME DEFAULT GETDATE()") AND dbSetupOK
End If

' ============================================
' 双模式参数体系 - 列迁移（v2.0）
' ============================================
If dbSetupOK Then
    Call EnsureColumn("FixedBrandInventory", "ParamMode", "NVARCHAR(20) DEFAULT 'Manual'")
    Call EnsureColumn("FixedBrandInventory", "DailySalesAvg", "DECIMAL(10,2) DEFAULT 0")
    Call EnsureColumn("FixedBrandInventory", "ConsecutiveDataMonths", "INT DEFAULT 0")
    Call EnsureColumn("FixedBrandInventory", "LastAutoCalcDate", "DATETIME")
    Call EnsureColumn("FixedBrandProducts", "LeadTimeDaysManual", "INT DEFAULT 7")
    ' 将现有 LeadTimeDays 值同步到 Manual 字段
    On Error Resume Next
    conn.Execute "UPDATE FixedBrandProducts SET LeadTimeDaysManual = LeadTimeDays WHERE LeadTimeDaysManual = 7 AND LeadTimeDays <> 7"
    On Error GoTo 0
    Call EnsureColumn("FixedBrandProducts", "SafetyStockManual", "INT DEFAULT 10")
End If

' ============================================
' 辅助函数
' ============================================

' 检查表是否存在（使用 sys.tables，比 SELECT TOP 1 更可靠）
Function TableExists(tblName)
    Dim rs, exists
    exists = False
    On Error Resume Next
    Set rs = conn.Execute("SELECT 1 FROM sys.tables WHERE name=N'" & tblName & "'")
    If Err.Number = 0 Then
        If Not rs Is Nothing Then
            If Not rs.EOF Then exists = True
            rs.Close
            Set rs = Nothing
        End If
    Else
        Err.Clear
        ' 后备方案：直接 SELECT 查询
        conn.Execute "SELECT TOP 1 1 FROM [" & tblName & "]"
        If Err.Number = 0 Then
            exists = True
        Else
            Err.Clear
        End If
    End If
    On Error GoTo 0
    TableExists = exists
End Function

' 创建 FixedBrand 表（带错误报告）
Function CreateFixedBrandTable(tblName, colDefs)
    On Error Resume Next
    conn.Execute "CREATE TABLE " & tblName & " (" & colDefs & ")"
    If Err.Number <> 0 Then
        Dim errMsg
        errMsg = "[FixedBrand DB Setup] Failed to create table " & tblName & ": " & Err.Description & " (Err#" & Err.Number & ")"
        ' 记录到 Session 供调试，不中断页面
        Session("DBSetupError") = Session("DBSetupError") & errMsg & "<br>"
        Err.Clear
        CreateFixedBrandTable = False
    Else
        CreateFixedBrandTable = True
    End If
    On Error GoTo 0
End Function

' 确保列存在（如不存在则 ALTER TABLE ADD）
Sub EnsureColumn(tblName, colName, colDef)
    On Error Resume Next
    conn.Execute "SELECT TOP 1 " & colName & " FROM " & tblName
    If Err.Number <> 0 Then
        Err.Clear
        conn.Execute "ALTER TABLE " & tblName & " ADD " & colName & " " & colDef
        If Err.Number <> 0 Then
            Session("DBSetupError") = Session("DBSetupError") & "[FixedBrand DB Setup] Failed to add column " & colName & " to " & tblName & ": " & Err.Description & "<br>"
            Err.Clear
        End If
    End If
    On Error GoTo 0
End Sub
%>
