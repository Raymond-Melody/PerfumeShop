<%
' ============================================
' 品牌定香采购模块 - 数据库自动迁移
' 在所有页面调用 OpenConnection() 后 include 此文件
' ============================================
On Error Resume Next
Dim dbSetupErrFlag : dbSetupErrFlag = False

' ----- FixedBrandProducts -----
conn.Execute "SELECT TOP 1 1 FROM FixedBrandProducts"
If Err.Number <> 0 Then
    Err.Clear
    dbSetupErrFlag = False  ' Reset flag for CREATE attempt
    conn.Execute "CREATE TABLE FixedBrandProducts (" & _
        "FixedProductID INT IDENTITY(1,1) PRIMARY KEY, " & _
        "ProductID INT, " & _
        "ProductCode NVARCHAR(50), " & _
        "ProductName NVARCHAR(200) NOT NULL, " & _
        "Specification NVARCHAR(100), " & _
        "UnitPrice DECIMAL(19,4) DEFAULT 0, " & _
        "SalePrice DECIMAL(19,4) DEFAULT 0, " & _
        "SupplierID INT, " & _
        "SupplierName NVARCHAR(200), " & _
        "MinOrderQty INT DEFAULT 1, " & _
        "LeadTimeDays INT DEFAULT 7, " & _
        "ImageURL NVARCHAR(500), " & _
        "Status NVARCHAR(20) DEFAULT 'Active', " & _
        "CreatedAt DATETIME DEFAULT GETDATE(), " & _
        "UpdatedAt DATETIME DEFAULT GETDATE()" & _
        ")"
End If

' ----- FixedBrandPurchaseOrders -----
conn.Execute "SELECT TOP 1 1 FROM FixedBrandPurchaseOrders"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE FixedBrandPurchaseOrders (" & _
        "PurchaseID INT IDENTITY(1,1) PRIMARY KEY, " & _
        "PurchaseNo NVARCHAR(50), " & _
        "SupplierID INT, " & _
        "SupplierName NVARCHAR(200), " & _
        "TotalAmount DECIMAL(19,4) DEFAULT 0, " & _
        "Status NVARCHAR(30) DEFAULT 'Draft', " & _
        "OrderDate DATETIME DEFAULT GETDATE(), " & _
        "ExpectedDate DATETIME, " & _
        "ApprovedBy NVARCHAR(100), " & _
        "ApprovedAt DATETIME, " & _
        "Remarks NVARCHAR(500), " & _
        "CreatedBy NVARCHAR(100), " & _
        "CreatedAt DATETIME DEFAULT GETDATE(), " & _
        "UpdatedAt DATETIME DEFAULT GETDATE()" & _
        ")"
End If

' ----- FixedBrandPurchaseDetails -----
conn.Execute "SELECT TOP 1 1 FROM FixedBrandPurchaseDetails"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE FixedBrandPurchaseDetails (" & _
        "DetailID INT IDENTITY(1,1) PRIMARY KEY, " & _
        "PurchaseID INT, " & _
        "FixedProductID INT, " & _
        "ProductName NVARCHAR(200), " & _
        "Specification NVARCHAR(100), " & _
        "Quantity INT DEFAULT 0, " & _
        "ReceivedQty INT DEFAULT 0, " & _
        "UnitPrice DECIMAL(19,4) DEFAULT 0, " & _
        "SubTotal DECIMAL(19,4) DEFAULT 0, " & _
        "ExpectedDate DATETIME, " & _
        "Remarks NVARCHAR(500)" & _
        ")"
End If

' ----- FixedBrandInventory -----
conn.Execute "SELECT TOP 1 1 FROM FixedBrandInventory"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE FixedBrandInventory (" & _
        "InventoryID INT IDENTITY(1,1) PRIMARY KEY, " & _
        "FixedProductID INT, " & _
        "ProductCode NVARCHAR(50), " & _
        "ProductName NVARCHAR(200), " & _
        "Specification NVARCHAR(100), " & _
        "StockQty INT DEFAULT 0, " & _
        "SafetyStock INT DEFAULT 10, " & _
        "MinOrderQty INT DEFAULT 1, " & _
        "AvgUnitCost DECIMAL(19,4) DEFAULT 0, " & _
        "LastPurchasePrice DECIMAL(19,4) DEFAULT 0, " & _
        "LastPurchaseDate DATETIME, " & _
        "LastPurchaseID INT, " & _
        "TotalPurchased INT DEFAULT 0, " & _
        "TotalSold INT DEFAULT 0, " & _
        "UpdatedAt DATETIME DEFAULT GETDATE()" & _
        ")"
End If

' ----- FixedBrandReceipts -----
conn.Execute "SELECT TOP 1 1 FROM FixedBrandReceipts"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE FixedBrandReceipts (" & _
        "ReceiptID INT IDENTITY(1,1) PRIMARY KEY, " & _
        "PurchaseID INT, " & _
        "ReceiptNo NVARCHAR(50), " & _
        "SupplierID INT, " & _
        "ReceivedBy NVARCHAR(100), " & _
        "ReceiptDate DATETIME DEFAULT GETDATE(), " & _
        "TotalReceivedQty INT DEFAULT 0, " & _
        "Notes NVARCHAR(500), " & _
        "CreatedAt DATETIME DEFAULT GETDATE()" & _
        ")"
End If

' ----- FixedBrandReceiptDetails -----
conn.Execute "SELECT TOP 1 1 FROM FixedBrandReceiptDetails"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE FixedBrandReceiptDetails (" & _
        "ReceiptDetailID INT IDENTITY(1,1) PRIMARY KEY, " & _
        "ReceiptID INT, " & _
        "DetailID INT, " & _
        "FixedProductID INT, " & _
        "AcceptedQty INT DEFAULT 0, " & _
        "RejectedQty INT DEFAULT 0, " & _
        "RejectReason NVARCHAR(500), " & _
        "UnitPrice DECIMAL(19,4) DEFAULT 0" & _
        ")"
End If

' ----- FixedBrandCostAllocation -----
conn.Execute "SELECT TOP 1 1 FROM FixedBrandCostAllocation"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE FixedBrandCostAllocation (" & _
        "AllocationID INT IDENTITY(1,1) PRIMARY KEY, " & _
        "OrderID INT, " & _
        "OrderNo NVARCHAR(100), " & _
        "PurchaseID INT, " & _
        "PurchaseNo NVARCHAR(50), " & _
        "FixedProductID INT, " & _
        "ProductName NVARCHAR(200), " & _
        "CostPerUnit DECIMAL(19,4) DEFAULT 0, " & _
        "Quantity INT DEFAULT 0, " & _
        "TotalCost DECIMAL(19,4) DEFAULT 0, " & _
        "SalePrice DECIMAL(19,4) DEFAULT 0, " & _
        "ProfitAmount DECIMAL(19,4) DEFAULT 0, " & _
        "ProfitRate DECIMAL(10,4) DEFAULT 0, " & _
        "AllocatedAt DATETIME DEFAULT GETDATE()" & _
        ")"
End If

' ============================================
' 双模式参数体系 - 列迁移（v2.0）
' ============================================
On Error Resume Next

' 1. FixedBrandInventory 添加 ParamMode
conn.Execute "SELECT TOP 1 ParamMode FROM FixedBrandInventory"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "ALTER TABLE FixedBrandInventory ADD ParamMode NVARCHAR(20) DEFAULT 'Manual'"
End If

' 2. FixedBrandInventory 添加 DailySalesAvg（统计推算的日均销量）
conn.Execute "SELECT TOP 1 DailySalesAvg FROM FixedBrandInventory"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "ALTER TABLE FixedBrandInventory ADD DailySalesAvg DECIMAL(10,2) DEFAULT 0"
End If

' 3. FixedBrandInventory 添加 ConsecutiveDataMonths
conn.Execute "SELECT TOP 1 ConsecutiveDataMonths FROM FixedBrandInventory"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "ALTER TABLE FixedBrandInventory ADD ConsecutiveDataMonths INT DEFAULT 0"
End If

' 4. FixedBrandInventory 添加 LastAutoCalcDate
conn.Execute "SELECT TOP 1 LastAutoCalcDate FROM FixedBrandInventory"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "ALTER TABLE FixedBrandInventory ADD LastAutoCalcDate DATETIME"
End If

' 5. FixedBrandProducts 添加 LeadTimeDaysManual（人工设定值备份）
conn.Execute "SELECT TOP 1 LeadTimeDaysManual FROM FixedBrandProducts"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "ALTER TABLE FixedBrandProducts ADD LeadTimeDaysManual INT DEFAULT 7"
    ' 将现有 LeadTimeDays 值同步到 Manual 字段
    conn.Execute "UPDATE FixedBrandProducts SET LeadTimeDaysManual = LeadTimeDays"
End If

' 6. FixedBrandProducts 添加 SafetyStockManual（人工设定安全库存）
conn.Execute "SELECT TOP 1 SafetyStockManual FROM FixedBrandProducts"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "ALTER TABLE FixedBrandProducts ADD SafetyStockManual INT DEFAULT 10"
End If

On Error GoTo 0
%>
