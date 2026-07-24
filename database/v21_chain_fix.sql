-- ============================================
-- V21 全链条修复 - 数据库结构迁移脚本
-- 用途：为"生产库存/成本闭环 + 费用并入利润 + 权限/基香映射"补齐所需列
-- 日期：2026-07-24
-- 环境：SQL Server 2017+
-- 执行：sqlcmd -S .\YOURPERFUME -d PerfumeShop -E -i database\v21_chain_fix.sql
-- 说明：全部幂等，可重复执行；各相关 ASP 页顶部亦保留 ALTER TABLE 兜底
-- ============================================

USE PerfumeShop;
GO

-- 1. NoteInventory 增加香调加权成本列（Accord 生产完成时结转）
IF COL_LENGTH('NoteInventory','WeightedUnitCost') IS NULL
BEGIN
    ALTER TABLE NoteInventory ADD WeightedUnitCost DECIMAL(19,4) DEFAULT 0;
    PRINT '[OK] NoteInventory.WeightedUnitCost added';
END
ELSE
    PRINT '[SKIP] NoteInventory.WeightedUnitCost already exists';
GO

-- 2. Orders 增加费用分摊金额列（运费/平台/推广分摊后回写，纳入利润）
IF COL_LENGTH('Orders','ExpenseAmount') IS NULL
BEGIN
    ALTER TABLE Orders ADD ExpenseAmount DECIMAL(19,4) DEFAULT 0;
    PRINT '[OK] Orders.ExpenseAmount added';
END
ELSE
    PRINT '[SKIP] Orders.ExpenseAmount already exists';
GO

-- 3. ProductInventory 确保存在（成品库存，品牌定香发货扣减）
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ProductInventory')
BEGIN
    CREATE TABLE ProductInventory (
        InventoryID INT IDENTITY(1,1) PRIMARY KEY,
        ProductID INT NULL,
        NoteID INT NULL,
        StockType NVARCHAR(20) NULL,
        StockQty INT NULL DEFAULT 0,
        SafetyStock INT NULL DEFAULT 0,
        UnitCost DECIMAL(19,4) NULL DEFAULT 0,
        UpdatedAt DATETIME2(7) NULL DEFAULT GETDATE()
    );
    PRINT '[OK] ProductInventory created';
END
ELSE
    PRINT '[SKIP] ProductInventory already exists';
GO

-- 4. InventoryTransactions 确保关键列存在（生产消耗/领料/发货统一写流水）
IF COL_LENGTH('InventoryTransactions','UnitCost') IS NULL
    ALTER TABLE InventoryTransactions ADD UnitCost DECIMAL(19,4) NULL;
GO
IF COL_LENGTH('InventoryTransactions','ReferenceType') IS NULL
    ALTER TABLE InventoryTransactions ADD ReferenceType NVARCHAR(50) NULL;
GO
IF COL_LENGTH('InventoryTransactions','ReferenceOrderID') IS NULL
    ALTER TABLE InventoryTransactions ADD ReferenceOrderID INT NULL;
GO

-- 5. RolePermissions 确保存在（操作级权限校验）
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'RolePermissions')
BEGIN
    CREATE TABLE RolePermissions (
        PermID INT IDENTITY(1,1) PRIMARY KEY,
        RoleID INT NOT NULL,
        ModuleCode NVARCHAR(50) NOT NULL,
        CanView BIT DEFAULT 0,
        CanCreate BIT DEFAULT 0,
        CanEdit BIT DEFAULT 0,
        CanDelete BIT DEFAULT 0,
        CanExport BIT DEFAULT 0,
        CanApprove BIT DEFAULT 0
    );
    PRINT '[OK] RolePermissions created';
END
ELSE
    PRINT '[SKIP] RolePermissions already exists';
GO

-- 6. ModulePermissions 确保存在（P3 权限表驱动）
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ModulePermissions')
BEGIN
    CREATE TABLE ModulePermissions (
        PermissionID INT IDENTITY(1,1) PRIMARY KEY,
        ModuleCode NVARCHAR(50) NOT NULL,
        ModuleName NVARCHAR(100) NOT NULL,
        ParentModule NVARCHAR(50) NULL,
        PermissionLevel INT NULL,
        RequiredRole NVARCHAR(20) NULL,
        URLPattern NVARCHAR(200) NULL,
        IsActive BIT NULL DEFAULT 1
    );
    PRINT '[OK] ModulePermissions created';
END
ELSE
    PRINT '[SKIP] ModulePermissions already exists';
GO

-- 7. BaseNotes 增加原料显式映射列（P3 替代按名称模糊匹配）
IF COL_LENGTH('BaseNotes','MaterialID') IS NULL
BEGIN
    ALTER TABLE BaseNotes ADD MaterialID INT NULL;
    PRINT '[OK] BaseNotes.MaterialID added';
END
ELSE
    PRINT '[SKIP] BaseNotes.MaterialID already exists';
GO

-- 8. RefundRecords 确保成本回冲标识列存在
IF COL_LENGTH('RefundRecords','CostWriteBack') IS NULL
    ALTER TABLE RefundRecords ADD CostWriteBack BIT NULL DEFAULT 0;
GO

-- 9. 库存预警开关默认值（若缺失则补）
IF NOT EXISTS (SELECT 1 FROM SiteSettings WHERE SettingKey = 'EnableLowStockAlert')
    INSERT INTO SiteSettings (SettingKey, SettingValue) VALUES ('EnableLowStockAlert', '1');
GO

-- 10. 索引优化：生产消耗与流水查询
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_InvTrans_Ref')
    CREATE NONCLUSTERED INDEX IX_InvTrans_Ref ON InventoryTransactions (ReferenceOrderID, ReferenceType);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_RecipeProducts_Product')
    CREATE NONCLUSTERED INDEX IX_RecipeProducts_Product ON RecipeProducts (ProductID, Status) INCLUDE (ProductRecipeID, PublishedAt);
GO

PRINT 'V21 全链条修复 - 数据库迁移完成';
GO
