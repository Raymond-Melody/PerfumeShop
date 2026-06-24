-- ============================================
-- V11 智能补货性能优化：关键索引创建
-- 执行日期：2026-06-10
-- 目的：解决智能补货页面超时问题
-- ============================================

PRINT '===== V11 智能补货性能索引优化开始 =====';
GO

-- 1. RawMaterialInventory: 低库存查询核心索引
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_RawMaterial_StockSafety' AND object_id = OBJECT_ID('RawMaterialInventory'))
BEGIN
    CREATE INDEX IX_RawMaterial_StockSafety ON RawMaterialInventory (StockQty, SafetyStock)
        INCLUDE (ItemName, ItemCode, Unit, UnitPrice, AvgDailyUsage, LeadTimeDays, ReorderPoint);
    PRINT '  IX_RawMaterial_StockSafety 创建成功';
END
ELSE
    PRINT '  IX_RawMaterial_StockSafety 已存在，跳过';
GO

-- 2. RawMaterialInventory: 搜索用索引（Name+Code）
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_RawMaterial_Name_Code' AND object_id = OBJECT_ID('RawMaterialInventory'))
BEGIN
    CREATE INDEX IX_RawMaterial_Name_Code ON RawMaterialInventory (ItemName, ItemCode);
    PRINT '  IX_RawMaterial_Name_Code 创建成功';
END
GO

-- 3. PackagingInventory: 低库存查询核心索引
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Packaging_StockSafety' AND object_id = OBJECT_ID('PackagingInventory'))
BEGIN
    CREATE INDEX IX_Packaging_StockSafety ON PackagingInventory (StockQty, SafetyStock)
        INCLUDE (ItemName, ItemCode, UnitPrice, AvgDailyUsage, LeadTimeDays, ReorderPoint);
    PRINT '  IX_Packaging_StockSafety 创建成功';
END
GO

-- 4. PackagingInventory: 搜索用索引
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Packaging_Name_Code' AND object_id = OBJECT_ID('PackagingInventory'))
BEGIN
    CREATE INDEX IX_Packaging_Name_Code ON PackagingInventory (ItemName, ItemCode);
    PRINT '  IX_Packaging_Name_Code 创建成功';
END
GO

-- 5. BottleStyles: 低库存查询核心索引
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_BottleStyles_StockSafety' AND object_id = OBJECT_ID('BottleStyles'))
BEGIN
    CREATE INDEX IX_BottleStyles_StockSafety ON BottleStyles (StockQty, SafetyStock)
        INCLUDE (BottleName, UnitPrice, AvgDailyUsage, LeadTimeDays, ReorderPoint);
    PRINT '  IX_BottleStyles_StockSafety 创建成功';
END
GO

-- 6. BottleStyles: 搜索用索引
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_BottleStyles_Name' AND object_id = OBJECT_ID('BottleStyles'))
BEGIN
    CREATE INDEX IX_BottleStyles_Name ON BottleStyles (BottleName);
    PRINT '  IX_BottleStyles_Name 创建成功';
END
GO

-- 7. PrintingInventory: 低库存查询核心索引
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Printing_StockSafety' AND object_id = OBJECT_ID('PrintingInventory'))
BEGIN
    CREATE INDEX IX_Printing_StockSafety ON PrintingInventory (StockQty, SafetyStock)
        INCLUDE (ItemName, ItemCode, Unit, UnitPrice, AvgDailyUsage, LeadTimeDays, ReorderPoint);
    PRINT '  IX_Printing_StockSafety 创建成功';
END
GO

-- 8. PrintingInventory: 搜索用索引
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Printing_Name_Code' AND object_id = OBJECT_ID('PrintingInventory'))
BEGIN
    CREATE INDEX IX_Printing_Name_Code ON PrintingInventory (ItemName, ItemCode);
    PRINT '  IX_Printing_Name_Code 创建成功';
END
GO

-- 9. SprayHeadInventory: 低库存查询核心索引
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_SprayHead_StockSafety' AND object_id = OBJECT_ID('SprayHeadInventory'))
BEGIN
    CREATE INDEX IX_SprayHead_StockSafety ON SprayHeadInventory (StockQty, SafetyStock)
        INCLUDE (ItemName, ItemCode, Unit, UnitPrice, AvgDailyUsage, LeadTimeDays, ReorderPoint);
    PRINT '  IX_SprayHead_StockSafety 创建成功';
END
GO

-- 10. SprayHeadInventory: 搜索用索引
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_SprayHead_Name_Code' AND object_id = OBJECT_ID('SprayHeadInventory'))
BEGIN
    CREATE INDEX IX_SprayHead_Name_Code ON SprayHeadInventory (ItemName, ItemCode);
    PRINT '  IX_SprayHead_Name_Code 创建成功';
END
GO

-- 11. SupplierPrices: 批量查询核心索引（ajax_supplier_info_batch 使用）
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_SupplierPrices_ItemCode_Active' AND object_id = OBJECT_ID('SupplierPrices'))
BEGIN
    CREATE INDEX IX_SupplierPrices_ItemCode_Active ON SupplierPrices (ItemCode, IsActive)
        INCLUDE (SupplierID, UnitPrice, CreatedAt);
    PRINT '  IX_SupplierPrices_ItemCode_Active 创建成功';
END
GO

-- 12. PurchaseOrders: OrderType + ExpectedDate 索引（采购列表查询优化）
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_PurchaseOrders_Type_Date' AND object_id = OBJECT_ID('PurchaseOrders'))
BEGIN
    CREATE INDEX IX_PurchaseOrders_Type_Date ON PurchaseOrders (OrderType, ExpectedDate);
    PRINT '  IX_PurchaseOrders_Type_Date 创建成功';
END
GO

-- 13. PurchaseOrderDetails: ItemCode 查询索引
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_PODetails_ItemCode' AND object_id = OBJECT_ID('PurchaseOrderDetails'))
BEGIN
    CREATE INDEX IX_PODetails_ItemCode ON PurchaseOrderDetails (ItemCode)
        INCLUDE (ItemName, UnitPrice);
    PRINT '  IX_PODetails_ItemCode 创建成功';
END
GO

-- 14. 更新统计信息，确保查询优化器使用新索引
UPDATE STATISTICS RawMaterialInventory;
UPDATE STATISTICS PackagingInventory;
UPDATE STATISTICS BottleStyles;
UPDATE STATISTICS SupplierPrices;
UPDATE STATISTICS PurchaseOrders;
UPDATE STATISTICS PurchaseOrderDetails;
GO

-- 15. 对可能动态创建的表做安全检查
IF EXISTS (SELECT 1 FROM sys.tables WHERE name = 'PrintingInventory')
    UPDATE STATISTICS PrintingInventory;
IF EXISTS (SELECT 1 FROM sys.tables WHERE name = 'SprayHeadInventory')
    UPDATE STATISTICS SprayHeadInventory;
GO

PRINT '===== V11 智能补货性能索引优化完成 =====';
GO
