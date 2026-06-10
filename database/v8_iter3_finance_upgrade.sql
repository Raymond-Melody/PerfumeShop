-- ============================================
-- V8 迭代3：财务管理升级 数据库变更脚本
-- 日期：2026-05-10
-- ============================================

-- 1. 应付账款表
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'AccountsPayable')
BEGIN
    CREATE TABLE [AccountsPayable] (
        [PayableID] INT IDENTITY(1,1) PRIMARY KEY,
        [PurchaseID] INT NULL,
        [SupplierID] INT NOT NULL,
        [SupplierName] NVARCHAR(200),
        [PayableNo] NVARCHAR(50),
        [Amount] DECIMAL(19,4) NOT NULL DEFAULT 0,
        [PaidAmount] DECIMAL(19,4) NOT NULL DEFAULT 0,
        [Balance] AS (Amount - PaidAmount),
        [Status] NVARCHAR(20) DEFAULT 'Pending',
        [DueDate] DATE,
        [InvoiceNo] NVARCHAR(100),
        [Notes] NVARCHAR(MAX),
        [CreatedAt] DATETIME2 DEFAULT GETDATE(),
        [UpdatedAt] DATETIME2 DEFAULT GETDATE()
    );
END
GO

-- 2. 应收账款表
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'AccountsReceivable')
BEGIN
    CREATE TABLE [AccountsReceivable] (
        [ReceivableID] INT IDENTITY(1,1) PRIMARY KEY,
        [OrderID] INT NULL,
        [UserID] INT NOT NULL,
        [CustomerName] NVARCHAR(200),
        [ReceivableNo] NVARCHAR(50),
        [Amount] DECIMAL(19,4) NOT NULL DEFAULT 0,
        [ReceivedAmount] DECIMAL(19,4) NOT NULL DEFAULT 0,
        [Balance] AS (Amount - ReceivedAmount),
        [Status] NVARCHAR(20) DEFAULT 'Pending',
        [DueDate] DATE,
        [Notes] NVARCHAR(MAX),
        [CreatedAt] DATETIME2 DEFAULT GETDATE(),
        [UpdatedAt] DATETIME2 DEFAULT GETDATE()
    );
END
GO

-- 3. 成本中心表
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'CostCenters')
BEGIN
    CREATE TABLE [CostCenters] (
        [CenterID] INT IDENTITY(1,1) PRIMARY KEY,
        [CenterCode] NVARCHAR(50) NOT NULL UNIQUE,
        [CenterName] NVARCHAR(200) NOT NULL,
        [CenterType] NVARCHAR(50) DEFAULT 'Department',
        [ParentID] INT NULL,
        [BudgetAmount] DECIMAL(19,4) DEFAULT 0,
        [IsActive] BIT DEFAULT 1,
        [Notes] NVARCHAR(MAX),
        [CreatedAt] DATETIME2 DEFAULT GETDATE(),
        [UpdatedAt] DATETIME2 DEFAULT GETDATE()
    );
    
    INSERT INTO CostCenters (CenterCode, CenterName, CenterType, BudgetAmount) VALUES
    ('RAW_MAT', '原料采购', 'Procurement', 0),
    ('PACKAGING', '包装物采购', 'Procurement', 0),
    ('BOTTLE', '瓶子采购', 'Procurement', 0),
    ('PRODUCTION', '生产制造', 'Production', 0),
    ('LOGISTICS', '物流运输', 'Logistics', 0),
    ('MARKETING', '市场营销', 'Marketing', 0),
    ('ADMIN', '行政管理', 'Admin', 0),
    ('RND', '研发设计', 'R&D', 0);
END
GO

-- 4. 总账流水表
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'GLTransactions')
BEGIN
    CREATE TABLE [GLTransactions] (
        [GLID] INT IDENTITY(1,1) PRIMARY KEY,
        [GLNo] NVARCHAR(50),
        [TransactionDate] DATETIME2 NOT NULL DEFAULT GETDATE(),
        [AccountCode] NVARCHAR(50),
        [AccountName] NVARCHAR(200),
        [DebitAmount] DECIMAL(19,4) DEFAULT 0,
        [CreditAmount] DECIMAL(19,4) DEFAULT 0,
        [Balance] DECIMAL(19,4) DEFAULT 0,
        [CenterID] INT NULL,
        [RefType] NVARCHAR(30),
        [RefID] INT NULL,
        [RefNo] NVARCHAR(100),
        [Description] NVARCHAR(500),
        [CreatedBy] NVARCHAR(50),
        [CreatedAt] DATETIME2 DEFAULT GETDATE()
    );
END
GO

-- 5. 扩展 PaymentRecords 增加付款类型
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('PaymentRecords') AND name = 'PaymentType')
BEGIN
    ALTER TABLE PaymentRecords ADD PaymentType NVARCHAR(30) DEFAULT 'Receipt';
    ALTER TABLE PaymentRecords ADD PayableID INT NULL;
    ALTER TABLE PaymentRecords ADD ReceivableID INT NULL;
    ALTER TABLE PaymentRecords ADD CenterID INT NULL;
    ALTER TABLE PaymentRecords ADD VoucherNo NVARCHAR(50);
END
GO

-- 6. 扩展 ExpenseRecords 关联成本中心
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('ExpenseRecords') AND name = 'CenterID')
BEGIN
    ALTER TABLE ExpenseRecords ADD CenterID INT NULL;
END
GO
