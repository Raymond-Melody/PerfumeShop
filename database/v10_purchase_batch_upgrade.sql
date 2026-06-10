-- ============================================
-- V10 采购中心全面迭代：批次与成本管理 — 数据库变更
-- 执行日期：2026-05-20
-- ============================================

-- 1. 创建 PurchaseBatches 表（采购收货批次记录）
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'PurchaseBatches')
BEGIN
    CREATE TABLE [PurchaseBatches] (
        [BatchID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [PurchaseDetailID] INT NULL,
        [PurchaseID] INT NULL,
        [BatchNo] NVARCHAR(50) NULL,
        [ItemType] NVARCHAR(30) DEFAULT 'RawMaterial',
        [ItemCode] NVARCHAR(50) NULL,
        [ItemName] NVARCHAR(200) NULL,
        [UnitPrice] DECIMAL(19,4) DEFAULT 0,
        [Quantity] FLOAT DEFAULT 0,
        [ReceivedQty] FLOAT DEFAULT 0,
        [RemainingQty] FLOAT DEFAULT 0,
        [ReceivedDate] DATETIME NULL,
        [SupplierID] INT NULL,
        [CostAllocated] BIT DEFAULT 0,
        [CreatedAt] DATETIME DEFAULT GETDATE()
    );
END
GO

-- 2. 创建 InventoryBatches 表（库存批次加权成本记录）
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'InventoryBatches')
BEGIN
    CREATE TABLE [InventoryBatches] (
        [InvBatchID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [ItemType] NVARCHAR(30) DEFAULT 'RawMaterial',
        [ItemID] INT NULL,
        [ItemCode] NVARCHAR(50) NULL,
        [ItemName] NVARCHAR(200) NULL,
        [BatchNo] NVARCHAR(50) NULL,
        [PurchaseBatchID] INT NULL,
        [UnitCost] DECIMAL(19,4) DEFAULT 0,
        [StockQty] FLOAT DEFAULT 0,
        [UpdatedAt] DATETIME DEFAULT GETDATE(),
        [CreatedAt] DATETIME DEFAULT GETDATE()
    );
END
GO

-- 3. 创建 OrderCostAllocation 表（订单成本分摊记录）
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'OrderCostAllocation')
BEGIN
    CREATE TABLE [OrderCostAllocation] (
        [AllocationID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [OrderID] INT NULL,
        [OrderNo] NVARCHAR(50) NULL,
        [CostType] NVARCHAR(30) DEFAULT 'Material',
        [ItemCode] NVARCHAR(50) NULL,
        [ItemName] NVARCHAR(200) NULL,
        [BatchID] INT NULL,
        [InvBatchID] INT NULL,
        [UnitCost] DECIMAL(19,4) DEFAULT 0,
        [Quantity] FLOAT DEFAULT 0,
        [TotalCost] DECIMAL(19,4) DEFAULT 0,
        [AllocatedAt] DATETIME DEFAULT GETDATE(),
        [CreatedAt] DATETIME DEFAULT GETDATE()
    );
END
GO

-- 4. 创建 PrintingInventory 表（印刷品库存）
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'PrintingInventory')
BEGIN
    CREATE TABLE [PrintingInventory] (
        [PrintingID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [ItemName] NVARCHAR(200) NOT NULL,
        [ItemCode] NVARCHAR(50) NULL,
        [PrintingType] NVARCHAR(30) DEFAULT 'Manual',
        [Specification] NVARCHAR(200) NULL,
        [Unit] NVARCHAR(20) DEFAULT 'pcs',
        [StockQty] FLOAT DEFAULT 0,
        [SafetyStock] FLOAT DEFAULT 0,
        [UnitPrice] DECIMAL(19,4) DEFAULT 0,
        [WeightedUnitCost] DECIMAL(19,4) DEFAULT 0,
        [SupplierID] INT NULL,
        [LastPurchaseDate] DATETIME NULL,
        [Notes] NVARCHAR(500) NULL,
        [IsActive] BIT DEFAULT 1,
        [CreatedAt] DATETIME DEFAULT GETDATE(),
        [UpdatedAt] DATETIME NULL
    );
END
GO

-- 5. 创建 SprayHeadInventory 表（喷头库存）
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'SprayHeadInventory')
BEGIN
    CREATE TABLE [SprayHeadInventory] (
        [SprayHeadID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [ItemName] NVARCHAR(200) NOT NULL,
        [ItemCode] NVARCHAR(50) NULL,
        [SprayType] NVARCHAR(30) DEFAULT 'Mist',
        [Material] NVARCHAR(100) NULL,
        [Specification] NVARCHAR(200) NULL,
        [Unit] NVARCHAR(20) DEFAULT 'pcs',
        [StockQty] FLOAT DEFAULT 0,
        [SafetyStock] FLOAT DEFAULT 0,
        [UnitPrice] DECIMAL(19,4) DEFAULT 0,
        [WeightedUnitCost] DECIMAL(19,4) DEFAULT 0,
        [SupplierID] INT NULL,
        [LastPurchaseDate] DATETIME NULL,
        [Notes] NVARCHAR(500) NULL,
        [IsActive] BIT DEFAULT 1,
        [CreatedAt] DATETIME DEFAULT GETDATE(),
        [UpdatedAt] DATETIME NULL
    );
END
GO

-- 6. 扩展 PurchaseOrderDetails 表
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('PurchaseOrderDetails') AND name = 'ItemType')
BEGIN
    ALTER TABLE PurchaseOrderDetails ADD ItemType NVARCHAR(30) DEFAULT 'RawMaterial';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('PurchaseOrderDetails') AND name = 'InventoryItemID')
BEGIN
    ALTER TABLE PurchaseOrderDetails ADD InventoryItemID INT NULL;
END
GO

-- 7. 扩展 PurchaseOrders 表：增加 CreatedByRealName
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('PurchaseOrders') AND name = 'CreatedByRealName')
BEGIN
    ALTER TABLE PurchaseOrders ADD CreatedByRealName NVARCHAR(50) NULL;
END
GO

-- 8. 扩展 RawMaterialInventory 表：增加加权成本字段
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('RawMaterialInventory') AND name = 'WeightedUnitCost')
BEGIN
    ALTER TABLE RawMaterialInventory ADD WeightedUnitCost DECIMAL(19,4) DEFAULT 0;
    -- 初始化：设置为当前UnitPrice
    UPDATE RawMaterialInventory SET WeightedUnitCost = ISNULL(UnitPrice, 0);
END
GO

-- 9. 扩展 PackagingInventory 表：增加加权成本字段
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('PackagingInventory') AND name = 'WeightedUnitCost')
BEGIN
    ALTER TABLE PackagingInventory ADD WeightedUnitCost DECIMAL(19,4) DEFAULT 0;
    UPDATE PackagingInventory SET WeightedUnitCost = ISNULL(UnitPrice, 0);
END
GO

-- 10. 扩展 BottleStyles 表：增加加权成本字段
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('BottleStyles') AND name = 'WeightedUnitCost')
BEGIN
    ALTER TABLE BottleStyles ADD WeightedUnitCost DECIMAL(19,4) DEFAULT 0;
    UPDATE BottleStyles SET WeightedUnitCost = ISNULL(UnitPrice, 0);
END
GO

-- 11. SupplierPrices 表增加 PriceType 示例数据兼容
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('SupplierPrices') AND name = 'Unit')
BEGIN
    ALTER TABLE SupplierPrices ADD Unit NVARCHAR(20) DEFAULT 'kg';
END
GO

PRINT 'V10 采购中心迭代数据库变更完成';
GO
