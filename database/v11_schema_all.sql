-- V11 Schema Init (ASCII safe)
PRINT 'START Schema Init';

-- 1. RawMaterialInventory columns
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('RawMaterialInventory') AND name='AvgDailyUsage')
BEGIN ALTER TABLE RawMaterialInventory ADD AvgDailyUsage DECIMAL(19,6) DEFAULT 0; PRINT 'OK: RawMaterialInventory.AvgDailyUsage'; END

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('RawMaterialInventory') AND name='LeadTimeDays')
BEGIN ALTER TABLE RawMaterialInventory ADD LeadTimeDays INT DEFAULT 7; PRINT 'OK: RawMaterialInventory.LeadTimeDays'; END

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('RawMaterialInventory') AND name='LastReplenishDate')
BEGIN ALTER TABLE RawMaterialInventory ADD LastReplenishDate DATETIME; PRINT 'OK: RawMaterialInventory.LastReplenishDate'; END

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('RawMaterialInventory') AND name='ReorderPoint')
BEGIN ALTER TABLE RawMaterialInventory ADD ReorderPoint DECIMAL(19,4) DEFAULT 0; PRINT 'OK: RawMaterialInventory.ReorderPoint'; END

-- 2. PackagingInventory table
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name='PackagingInventory')
BEGIN
    CREATE TABLE PackagingInventory (
        PackagingID INT IDENTITY(1,1) PRIMARY KEY, ItemName NVARCHAR(100), ItemCode NVARCHAR(50),
        StockQty DECIMAL(10,2) DEFAULT 0, SafetyStock DECIMAL(10,2) DEFAULT 0,
        Unit NVARCHAR(20) DEFAULT 'pcs', UnitPrice DECIMAL(10,2) DEFAULT 0,
        AvgDailyUsage DECIMAL(19,6) DEFAULT 0, LeadTimeDays INT DEFAULT 7,
        LastReplenishDate DATETIME, ReorderPoint DECIMAL(19,4) DEFAULT 0, UpdatedAt DATETIME DEFAULT GETDATE()
    );
    PRINT 'OK: PackagingInventory created';
END

-- 3. BottleStyles columns
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('BottleStyles') AND name='StockQty')
BEGIN ALTER TABLE BottleStyles ADD StockQty DECIMAL(10,2) DEFAULT 0; PRINT 'OK: BottleStyles.StockQty'; END

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('BottleStyles') AND name='SafetyStock')
BEGIN ALTER TABLE BottleStyles ADD SafetyStock DECIMAL(10,2) DEFAULT 0; PRINT 'OK: BottleStyles.SafetyStock'; END

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('BottleStyles') AND name='AvgDailyUsage')
BEGIN ALTER TABLE BottleStyles ADD AvgDailyUsage DECIMAL(19,6) DEFAULT 0; PRINT 'OK: BottleStyles.AvgDailyUsage'; END

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('BottleStyles') AND name='LeadTimeDays')
BEGIN ALTER TABLE BottleStyles ADD LeadTimeDays INT DEFAULT 7; PRINT 'OK: BottleStyles.LeadTimeDays'; END

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('BottleStyles') AND name='LastReplenishDate')
BEGIN ALTER TABLE BottleStyles ADD LastReplenishDate DATETIME; PRINT 'OK: BottleStyles.LastReplenishDate'; END

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('BottleStyles') AND name='ReorderPoint')
BEGIN ALTER TABLE BottleStyles ADD ReorderPoint DECIMAL(19,4) DEFAULT 0; PRINT 'OK: BottleStyles.ReorderPoint'; END

-- 4. PrintingInventory table
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name='PrintingInventory')
BEGIN
    CREATE TABLE PrintingInventory (
        PrintingID INT IDENTITY(1,1) PRIMARY KEY, ItemName NVARCHAR(100), ItemCode NVARCHAR(50),
        StockQty DECIMAL(10,2) DEFAULT 0, SafetyStock DECIMAL(10,2) DEFAULT 0,
        Unit NVARCHAR(20) DEFAULT 'sheets', UnitPrice DECIMAL(10,2) DEFAULT 0,
        AvgDailyUsage DECIMAL(19,6) DEFAULT 0, LeadTimeDays INT DEFAULT 7,
        LastReplenishDate DATETIME, ReorderPoint DECIMAL(19,4) DEFAULT 0, UpdatedAt DATETIME DEFAULT GETDATE()
    );
    PRINT 'OK: PrintingInventory created';
END

-- 5. SprayHeadInventory table
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name='SprayHeadInventory')
BEGIN
    CREATE TABLE SprayHeadInventory (
        SprayHeadID INT IDENTITY(1,1) PRIMARY KEY, ItemName NVARCHAR(100), ItemCode NVARCHAR(50),
        StockQty DECIMAL(10,2) DEFAULT 0, SafetyStock DECIMAL(10,2) DEFAULT 0,
        Unit NVARCHAR(20) DEFAULT 'pcs', UnitPrice DECIMAL(10,2) DEFAULT 0,
        AvgDailyUsage DECIMAL(19,6) DEFAULT 0, LeadTimeDays INT DEFAULT 7,
        LastReplenishDate DATETIME, ReorderPoint DECIMAL(19,4) DEFAULT 0, UpdatedAt DATETIME DEFAULT GETDATE()
    );
    PRINT 'OK: SprayHeadInventory created';
END

-- 6. PurchaseHistoryStats table
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name='PurchaseHistoryStats')
BEGIN
    CREATE TABLE PurchaseHistoryStats (
        StatID INT IDENTITY(1,1) PRIMARY KEY, ItemType NVARCHAR(30) NOT NULL,
        ItemCode NVARCHAR(100), ItemName NVARCHAR(200),
        Avg30DayUsage DECIMAL(19,6) DEFAULT 0, Avg90DayUsage DECIMAL(19,6) DEFAULT 0,
        LastOrderDate DATETIME, TotalOrders90Days INT DEFAULT 0,
        PreferredSupplierID INT, PreferredUnitPrice DECIMAL(19,4) DEFAULT 0,
        UpdatedAt DATETIME DEFAULT GETDATE()
    );
    PRINT 'OK: PurchaseHistoryStats created';
END

PRINT 'DONE Schema Init';
GO
