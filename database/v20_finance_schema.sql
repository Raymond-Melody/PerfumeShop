-- ============================================
-- V20 财务中心数据库Schema迁移脚本
-- 用途：确保所有财务模块所需的表和列存在
-- 日期：2026-07-16
-- 环境：SQL Server 2017+
-- ============================================

-- 1. 确保 SiteSettings 表存在（成本计价方式等配置）
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'SiteSettings')
BEGIN
    CREATE TABLE SiteSettings (
        SettingId INT IDENTITY(1,1) PRIMARY KEY,
        SettingKey NVARCHAR(100) NOT NULL UNIQUE,
        SettingValue NVARCHAR(MAX),
        UpdatedAt DATETIME2 DEFAULT GETDATE()
    );
    INSERT INTO SiteSettings (SettingKey, SettingValue) VALUES ('CostCalculationMethod', 'FIFO');
END
ELSE IF NOT EXISTS (SELECT 1 FROM SiteSettings WHERE SettingKey = 'CostCalculationMethod')
    INSERT INTO SiteSettings (SettingKey, SettingValue) VALUES ('CostCalculationMethod', 'FIFO');

-- 2. 确保 FundAccounts 表存在
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'FundAccounts')
BEGIN
    CREATE TABLE FundAccounts (
        AccountId INT IDENTITY(1,1) PRIMARY KEY,
        AccountName NVARCHAR(200) NOT NULL,
        AccountCode NVARCHAR(50),
        TotalBalance DECIMAL(19,4) DEFAULT 0,
        AlertThreshold DECIMAL(19,4) DEFAULT 0,
        IsActive BIT DEFAULT 1,
        UpdatedAt DATETIME2
    );
END

-- 3. 确保 AccountsPayable 列完整
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('AccountsPayable') AND name = 'PayableNo')
    ALTER TABLE AccountsPayable ADD PayableNo NVARCHAR(50);
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('AccountsPayable') AND name = 'SupplierName')
    ALTER TABLE AccountsPayable ADD SupplierName NVARCHAR(200);
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('AccountsPayable') AND name = 'DueDate')
    ALTER TABLE AccountsPayable ADD DueDate DATE;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('AccountsPayable') AND name = 'PaidAmount')
    ALTER TABLE AccountsPayable ADD PaidAmount DECIMAL(19,4) DEFAULT 0;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('AccountsPayable') AND name = 'UpdatedAt')
    ALTER TABLE AccountsPayable ADD UpdatedAt DATETIME2;

-- 4. 确保 AccountsReceivable 列完整
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('AccountsReceivable') AND name = 'CustomerName')
    ALTER TABLE AccountsReceivable ADD CustomerName NVARCHAR(200);
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('AccountsReceivable') AND name = 'ReceivedAmount')
    ALTER TABLE AccountsReceivable ADD ReceivedAmount DECIMAL(19,4) DEFAULT 0;

-- 5. 确保 ReconciliationLogs 列完整
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('ReconciliationLogs') AND name = 'OrderNo')
    ALTER TABLE ReconciliationLogs ADD OrderNo NVARCHAR(50);
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('ReconciliationLogs') AND name = 'OrderAmount')
    ALTER TABLE ReconciliationLogs ADD OrderAmount DECIMAL(19,4);
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('ReconciliationLogs') AND name = 'PaymentAmount')
    ALTER TABLE ReconciliationLogs ADD PaymentAmount DECIMAL(19,4);
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('ReconciliationLogs') AND name = 'Difference')
    ALTER TABLE ReconciliationLogs ADD [Difference] DECIMAL(19,4);
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('ReconciliationLogs') AND name = 'ReconcileDate')
    ALTER TABLE ReconciliationLogs ADD ReconcileDate DATE;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('ReconciliationLogs') AND name = 'Resolution')
    ALTER TABLE ReconciliationLogs ADD Resolution NVARCHAR(MAX);

-- 6. 确保 RefundRecords 列完整
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('RefundRecords') AND name = 'RefundAmount')
    ALTER TABLE RefundRecords ADD RefundAmount DECIMAL(19,4) DEFAULT 0;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('RefundRecords') AND name = 'Status')
    ALTER TABLE RefundRecords ADD Status NVARCHAR(30) DEFAULT 'Pending';

-- 7. 确保 PaymentRecords 列完整
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PaymentRecords') AND name = 'PaymentType')
    ALTER TABLE PaymentRecords ADD PaymentType NVARCHAR(30);
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PaymentRecords') AND name = 'Amount')
    ALTER TABLE PaymentRecords ADD Amount DECIMAL(19,4);
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PaymentRecords') AND name = 'Status')
    ALTER TABLE PaymentRecords ADD Status NVARCHAR(30);
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PaymentRecords') AND name = 'Remark')
    ALTER TABLE PaymentRecords ADD Remark NVARCHAR(MAX);

-- 8. 确保 Orders 成本字段
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('Orders') AND name = 'CostAmount')
    ALTER TABLE Orders ADD CostAmount DECIMAL(19,4) DEFAULT 0;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('Orders') AND name = 'ProfitAmount')
    ALTER TABLE Orders ADD ProfitAmount DECIMAL(19,4) DEFAULT 0;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('Orders') AND name = 'ShippingStatus')
    ALTER TABLE Orders ADD ShippingStatus NVARCHAR(30);

-- 9. 确保 Products 成本字段
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('Products') AND name = 'Bomcost')
    ALTER TABLE Products ADD Bomcost DECIMAL(19,4) DEFAULT 0;

-- 10. 确保 PurchaseOrders 审核字段
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PurchaseOrders') AND name = 'ApprovedAt')
    ALTER TABLE PurchaseOrders ADD ApprovedAt DATETIME2;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PurchaseOrders') AND name = 'ApprovedBy')
    ALTER TABLE PurchaseOrders ADD ApprovedBy INT;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PurchaseOrders') AND name = 'SupplierId')
    ALTER TABLE PurchaseOrders ADD SupplierId INT;

-- 11. 确保 BudgetPlans 列完整
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('BudgetPlans') AND name = 'BudgetName')
    ALTER TABLE BudgetPlans ADD BudgetName NVARCHAR(200);
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('BudgetPlans') AND name = 'ActualAmount')
    ALTER TABLE BudgetPlans ADD ActualAmount DECIMAL(19,4) DEFAULT 0;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('BudgetPlans') AND name = 'Status')
    ALTER TABLE BudgetPlans ADD Status NVARCHAR(30) DEFAULT 'Active';

-- 12. 确保 ExpenseRecords 列完整
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('ExpenseRecords') AND name = 'AllocationMethod')
    ALTER TABLE ExpenseRecords ADD AllocationMethod NVARCHAR(30);
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('ExpenseRecords') AND name = 'AllocationRatio')
    ALTER TABLE ExpenseRecords ADD AllocationRatio FLOAT;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('ExpenseRecords') AND name = 'CenterId')
    ALTER TABLE ExpenseRecords ADD CenterId INT;

-- 13. 索引优化（V19.4 性能增强）
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Orders_Status_Created')
    CREATE INDEX IX_Orders_Status_Created ON Orders (Status, CreatedAt) WHERE Status IN ('Paid','Processing','Shipped','Completed');
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_ProductCosts_Product_Created')
    CREATE INDEX IX_ProductCosts_Product_Created ON ProductCosts (ProductId, CreatedAt);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_PaymentRecords_Created')
    CREATE INDEX IX_PaymentRecords_Created ON PaymentRecords (CreatedAt) WHERE Status = 'Completed';
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_ReconciliationLogs_Status')
    CREATE INDEX IX_ReconciliationLogs_Status ON ReconciliationLogs (Status);

PRINT 'V20 财务中心数据库迁移完成';
