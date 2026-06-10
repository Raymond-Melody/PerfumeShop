-- ============================================
-- V8 迭代2：采购管理中心完善 - 数据库变更
-- 执行日期：2026-05-10
-- ============================================

-- 1. 扩展 PurchaseOrders：增加 OrderType
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('PurchaseOrders') AND name = 'OrderType')
BEGIN
    ALTER TABLE PurchaseOrders ADD OrderType NVARCHAR(20) DEFAULT 'RawMaterial';
    -- 可选值: RawMaterial / Packaging / Bottle
END
GO

-- 2. 扩展 PurchaseOrders：确保 ExpectedDeliveryDate 字段存在
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('PurchaseOrders') AND name = 'ExpectedDeliveryDate')
BEGIN
    ALTER TABLE PurchaseOrders ADD ExpectedDeliveryDate DATETIME2(7);
END
GO

-- 3. 扩展 SupplierPrices：增加 PriceType 和 Unit
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('SupplierPrices') AND name = 'PriceType')
BEGIN
    ALTER TABLE SupplierPrices ADD PriceType NVARCHAR(30) DEFAULT 'Standard';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('SupplierPrices') AND name = 'Unit')
BEGIN
    ALTER TABLE SupplierPrices ADD Unit NVARCHAR(20) DEFAULT 'kg';
END
GO

-- 4. 新增 SupplierContracts 表
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'SupplierContracts')
BEGIN
    CREATE TABLE [SupplierContracts] (
        [ContractID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [SupplierID] INT NOT NULL,
        [ContractNo] NVARCHAR(50),
        [ContractName] NVARCHAR(200),
        [ContractType] NVARCHAR(30) DEFAULT 'Supply',
        [StartDate] DATETIME2(7),
        [EndDate] DATETIME2(7),
        [TotalAmount] DECIMAL(19,4),
        [PaymentTerms] NVARCHAR(200),
        [TermsSummary] NVARCHAR(MAX),
        [AttachmentURL] NVARCHAR(500),
        [Status] NVARCHAR(20) DEFAULT 'Active',
        [SignedAt] DATETIME2(7),
        [CreatedAt] DATETIME2(7) DEFAULT GETDATE(),
        [UpdatedAt] DATETIME2(7)
    );
END
GO

-- 5. 新增 SupplierEvaluations 表
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'SupplierEvaluations')
BEGIN
    CREATE TABLE [SupplierEvaluations] (
        [EvaluationID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [SupplierID] INT NOT NULL,
        [EvaluatedBy] NVARCHAR(50),
        [EvaluationDate] DATETIME2(7) DEFAULT GETDATE(),
        [QualityScore] INT DEFAULT 0,
        [DeliveryScore] INT DEFAULT 0,
        [PriceScore] INT DEFAULT 0,
        [ServiceScore] INT DEFAULT 0,
        [OverallScore] INT DEFAULT 0,
        [Rating] NVARCHAR(10) DEFAULT 'C',
        [Comments] NVARCHAR(MAX),
        [Recommendations] NVARCHAR(MAX),
        [Period] NVARCHAR(20),
        [CreatedAt] DATETIME2(7) DEFAULT GETDATE()
    );
END
GO

-- 6. 历史数据：将已有 CategoryCode 映射到 OrderType
UPDATE PurchaseOrders SET OrderType = 'RawMaterial' WHERE OrderType IS NULL AND CategoryCode IN ('RAW','BASE');
UPDATE PurchaseOrders SET OrderType = 'Packaging' WHERE OrderType IS NULL AND CategoryCode = 'PACK';
UPDATE PurchaseOrders SET OrderType = 'RawMaterial' WHERE OrderType IS NULL;

PRINT 'V8 迭代2数据库变更完成';
GO
