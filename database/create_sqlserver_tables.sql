-- Auto-generated T-SQL CREATE TABLE script from Access
-- Generated: 2026-05-02 13:30:40
USE PerfumeShop;
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'AccordProductionDetails')
BEGIN
    CREATE TABLE [AccordProductionDetails] (
        [ActualQty] FLOAT NULL,
        [DetailID] INT IDENTITY(1,1) NOT NULL,
        [MaterialID] INT NULL,
        [MaterialName] NVARCHAR(100) NULL,
        [PlannedQty] FLOAT NULL,
        [ProductionID] INT NULL,
        [TotalCost] DECIMAL(19,4) NULL,
        [UnitCost] DECIMAL(19,4) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'AccordProductions')
BEGIN
    CREATE TABLE [AccordProductions] (
        [AccordRecipeID] INT NULL,
        [ActualQty] FLOAT NULL,
        [ApprovedBy] NVARCHAR(50) NULL,
        [BatchNo] NVARCHAR(30) NULL,
        [CompletedAt] DATETIME2(7) NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [NoteID] INT NULL,
        [NoteName] NVARCHAR(100) NULL,
        [Notes] NVARCHAR(MAX) NULL,
        [PlannedQty] FLOAT NULL,
        [ProductionID] INT IDENTITY(1,1) NOT NULL,
        [StartedAt] DATETIME2(7) NULL,
        [Status] NVARCHAR(20) NULL,
        [UpdatedAt] DATETIME2(7) NULL,
        [WorkCenter] NVARCHAR(20) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'AccordQCReports')
BEGIN
    CREATE TABLE [AccordQCReports] (
        [BatchNo] NVARCHAR(30) NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [Notes] NVARCHAR(MAX) NULL,
        [ProductionID] INT NULL,
        [QCReportID] INT IDENTITY(1,1) NOT NULL,
        [QCResult] NVARCHAR(20) NULL,
        [TestDate] DATETIME2(7) NULL,
        [TesterID] INT NULL,
        [TesterName] NVARCHAR(50) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'AdminLogs')
BEGIN
    CREATE TABLE [AdminLogs] (
        [ActionType] NVARCHAR(100) NULL,
        [AdminID] INT NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [LogID] INT IDENTITY(1,1) NOT NULL,
        [ModuleCode] NVARCHAR(50) NULL,
        [Notes] NVARCHAR(MAX) NULL,
        [RecordID] NVARCHAR(50) NULL,
        [TableName] NVARCHAR(50) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'AdminRoles')
BEGIN
    CREATE TABLE [AdminRoles] (
        [CreatedAt] DATETIME2(7) NULL,
        [Description] NVARCHAR(MAX) NULL,
        [Permissions] NVARCHAR(MAX) NULL,
        [RoleCode] NVARCHAR(20) NOT NULL,
        [RoleID] INT IDENTITY(1,1) NOT NULL,
        [RoleName] NVARCHAR(50) NOT NULL,
        [UpdatedAt] DATETIME2(7) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'AdminUsers')
BEGIN
    CREATE TABLE [AdminUsers] (
        [AdminID] INT IDENTITY(1,1) NOT NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [Department] NVARCHAR(50) NULL,
        [Email] NVARCHAR(100) NOT NULL,
        [FullName] NVARCHAR(100) NULL,
        [IsActive] BIT NULL,
        [IsLocked] BIT NULL,
        [LastLogin] DATETIME2(7) NULL,
        [PasswordHash] NVARCHAR(255) NOT NULL,
        [ResetToken] NVARCHAR(255) NULL,
        [ResetTokenExpiry] DATETIME2(7) NULL,
        [RoleID] INT NULL,
        [Username] NVARCHAR(50) NOT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'BaseNotes')
BEGIN
    CREATE TABLE [BaseNotes] (
        [BaseNoteID] INT IDENTITY(1,1) NOT NULL,
        [BaseNoteName] NVARCHAR(100) NOT NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [Description] NVARCHAR(MAX) NULL,
        [Ingredients] NVARCHAR(MAX) NULL,
        [IsActive] BIT NULL,
        [UnitPrice] DECIMAL(19,4) NULL
    );
END
ELSE
BEGIN
    -- V9: 自动添加 UnitPrice 字段
    IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('BaseNotes') AND name = 'UnitPrice')
        ALTER TABLE [BaseNotes] ADD [UnitPrice] DECIMAL(19,4) NULL;
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'BottleStyles')
BEGIN
    CREATE TABLE [BottleStyles] (
        [BottleID] INT IDENTITY(1,1) NOT NULL,
        [BottleName] NVARCHAR(50) NOT NULL,
        [Description] NVARCHAR(MAX) NULL,
        [ImageURL] NVARCHAR(200) NULL,
        [IsActive] BIT NULL,
        [PriceAddition] DECIMAL(19,4) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'BudgetPlans')
BEGIN
    CREATE TABLE [BudgetPlans] (
        [ActualAmount] DECIMAL(19,4) NULL,
        [AlertPercent] FLOAT NULL,
        [AlertROI] FLOAT NULL,
        [BudgetAmount] DECIMAL(19,4) NULL,
        [BudgetID] INT IDENTITY(1,1) NOT NULL,
        [BudgetName] NVARCHAR(100) NULL,
        [Category] NVARCHAR(50) NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [CreatedBy] NVARCHAR(50) NULL,
        [GMVAmount] DECIMAL(19,4) NULL,
        [Period] NVARCHAR(10) NULL,
        [ROI] FLOAT NULL,
        [Status] NVARCHAR(20) NULL,
        [UpdatedAt] DATETIME2(7) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Cart')
BEGIN
    CREATE TABLE [Cart] (
        [BaseNoteID] INT NULL,
        [BottleID] INT NULL,
        [CartID] INT IDENTITY(1,1) NOT NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [CustomLabel] NVARCHAR(200) NULL,
        [MiddleNoteID] INT NULL,
        [ProductID] INT NOT NULL,
        [Quantity] INT NULL,
        [SessionID] NVARCHAR(100) NULL,
        [TopNoteID] INT NULL,
        [UnitPrice] DECIMAL(19,4) NOT NULL,
        [UserID] INT NULL,
        [VolumeID] INT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'CartNoteSelections')
BEGIN
    CREATE TABLE [CartNoteSelections] (
        [CartID] INT NOT NULL,
        [NoteID] INT NOT NULL,
        [NoteType] NVARCHAR(20) NULL,
        [Percentage] INT NOT NULL,
        [SelectionID] INT IDENTITY(1,1) NOT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Categories')
BEGIN
    CREATE TABLE [Categories] (
        [CategoryID] INT IDENTITY(1,1) NOT NULL,
        [CategoryName] NVARCHAR(100) NOT NULL,
        [IsActive] BIT NULL,
        [SortOrder] INT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Coupons')
BEGIN
    CREATE TABLE [Coupons] (
        [CouponCode] NVARCHAR(50) NULL,
        [CouponID] INT IDENTITY(1,1) NOT NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [DiscountType] NVARCHAR(20) NULL,
        [DiscountValue] DECIMAL(19,4) NULL,
        [EndDate] DATETIME2(7) NULL,
        [IsActive] BIT NULL,
        [MinPurchase] DECIMAL(19,4) NULL,
        [StartDate] DATETIME2(7) NULL,
        [UsageLimit] INT NULL,
        [UsedCount] INT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'DailyStatistics')
BEGIN
    CREATE TABLE [DailyStatistics] (
        [CreatedAt] DATETIME2(7) NULL,
        [DataJSON] NVARCHAR(MAX) NULL,
        [NewUsers] INT NULL,
        [StatDate] DATETIME2(7) NOT NULL,
        [StatID] INT IDENTITY(1,1) NOT NULL,
        [TopNoteID] INT NULL,
        [TopProductID] INT NULL,
        [TotalOrders] INT NULL,
        [TotalRevenue] DECIMAL(19,4) NULL,
        [TotalUsers] INT NULL,
        [UpdatedAt] DATETIME2(7) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ExpenseRecords')
BEGIN
    CREATE TABLE [ExpenseRecords] (
        [AllocationMethod] NVARCHAR(20) NULL,
        [AllocationRatio] FLOAT NULL,
        [Amount] DECIMAL(19,4) NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [ExpenseID] INT IDENTITY(1,1) NOT NULL,
        [ExpenseName] NVARCHAR(100) NULL,
        [ExpenseType] NVARCHAR(30) NULL,
        [OrderID] INT NULL,
        [Period] NVARCHAR(10) NULL,
        [ProductID] INT NULL,
        [SourceOrderID] INT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'FormulaNotes')
BEGIN
    CREATE TABLE [FormulaNotes] (
        [FormulaID] INT NOT NULL,
        [ID] INT IDENTITY(1,1) NOT NULL,
        [NoteID] INT NOT NULL,
        [Percentage] INT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Formulas')
BEGIN
    CREATE TABLE [Formulas] (
        [CreatedAt] DATETIME2(7) NULL,
        [Description] NVARCHAR(MAX) NULL,
        [FormulaID] INT IDENTITY(1,1) NOT NULL,
        [FormulaName] NVARCHAR(100) NOT NULL,
        [IsActive] BIT NULL,
        [UpdatedAt] DATETIME2(7) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'FragranceIngredients')
BEGIN
    CREATE TABLE [FragranceIngredients] (
        [CreatedAt] DATETIME2(7) NULL,
        [FragranceIngredientID] INT IDENTITY(1,1) NOT NULL,
        [IngredientID] INT NOT NULL,
        [NoteID] INT NOT NULL,
        [Percentage] REAL NOT NULL,
        [SortOrder] INT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'FragranceNotes')
BEGIN
    CREATE TABLE [FragranceNotes] (
        [BaseNoteID] INT NULL,
        [Description] NVARCHAR(MAX) NULL,
        [ImageURL] NVARCHAR(200) NULL,
        [Ingredients] NVARCHAR(MAX) NULL,
        [IsActive] BIT NULL,
        [IsBaseNote] INT NULL,
        [NoteID] INT IDENTITY(1,1) NOT NULL,
        [NoteName] NVARCHAR(50) NOT NULL,
        [NoteType] NVARCHAR(20) NOT NULL,
        [PriceAddition] DECIMAL(19,4) NULL,
        [RecommendedPercentage] INT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'FundAccounts')
BEGIN
    CREATE TABLE [FundAccounts] (
        [AccountID] INT IDENTITY(1,1) NOT NULL,
        [AccountName] NVARCHAR(100) NULL,
        [AccountType] NVARCHAR(30) NULL,
        [AlertThreshold] DECIMAL(19,4) NULL,
        [AvailableBalance] DECIMAL(19,4) NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [FrozenAmount] DECIMAL(19,4) NULL,
        [IsActive] BIT NULL,
        [LastSyncAt] DATETIME2(7) NULL,
        [PendingSettlement] DECIMAL(19,4) NULL,
        [TotalBalance] DECIMAL(19,4) NULL,
        [UpdatedAt] DATETIME2(7) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Ingredients')
BEGIN
    CREATE TABLE [Ingredients] (
        [CASNumber] NVARCHAR(50) NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [Description] NVARCHAR(255) NULL,
        [IngredientID] INT IDENTITY(1,1) NOT NULL,
        [IngredientName] NVARCHAR(100) NOT NULL,
        [IsActive] BIT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'InventoryTransactions')
BEGIN
    CREATE TABLE [InventoryTransactions] (
        [CreatedAt] DATETIME2(7) NULL,
        [CreatedBy] NVARCHAR(50) NULL,
        [MaterialID] INT NULL,
        [NoteID] INT NOT NULL,
        [Notes] NVARCHAR(MAX) NULL,
        [ProductID] INT NULL,
        [Quantity] INT NOT NULL,
        [ReferenceOrderID] INT NULL,
        [ReferenceType] NVARCHAR(50) NULL,
        [TransactionDirection] NVARCHAR(10) NULL,
        [TransactionID] INT IDENTITY(1,1) NOT NULL,
        [TransactionType] NVARCHAR(20) NULL,
        [UnitCost] DECIMAL(19,4) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'MarketingCampaigns')
BEGIN
    CREATE TABLE [MarketingCampaigns] (
        [CampaignID] INT IDENTITY(1,1) NOT NULL,
        [CampaignName] NVARCHAR(200) NULL,
        [CampaignType] NVARCHAR(50) NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [Description] NVARCHAR(MAX) NULL,
        [DiscountValue] DECIMAL(19,4) NULL,
        [EndDate] DATETIME2(7) NULL,
        [IsActive] BIT NULL,
        [MinPurchase] DECIMAL(19,4) NULL,
        [ParticipantCount] INT NULL,
        [StartDate] DATETIME2(7) NULL,
        [TotalSales] DECIMAL(19,4) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'MaterialOutbound')
BEGIN
    CREATE TABLE [MaterialOutbound] (
        [ApprovedBy] NVARCHAR(50) NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [Notes] NVARCHAR(MAX) NULL,
        [OutboundDate] DATETIME2(7) NULL,
        [OutboundID] INT IDENTITY(1,1) NOT NULL,
        [OutboundNo] NVARCHAR(50) NULL,
        [OutboundType] NVARCHAR(20) NULL,
        [ReferenceID] INT NULL,
        [ReferenceType] NVARCHAR(50) NULL,
        [RequestedBy] NVARCHAR(50) NULL,
        [Status] NVARCHAR(20) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'MaterialOutboundDetails')
BEGIN
    CREATE TABLE [MaterialOutboundDetails] (
        [ActualQty] FLOAT NULL,
        [MaterialID] INT NULL,
        [OutboundDetailID] INT IDENTITY(1,1) NOT NULL,
        [OutboundID] INT NULL,
        [ProductionOrderRef] INT NULL,
        [RequestedQty] FLOAT NULL,
        [TotalAmount] DECIMAL(19,4) NULL,
        [UnitPrice] DECIMAL(19,4) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ModulePermissions')
BEGIN
    CREATE TABLE [ModulePermissions] (
        [IsActive] BIT NULL,
        [ModuleCode] NVARCHAR(50) NOT NULL,
        [ModuleName] NVARCHAR(100) NOT NULL,
        [ParentModule] NVARCHAR(50) NULL,
        [PermissionID] INT IDENTITY(1,1) NOT NULL,
        [PermissionLevel] INT NULL,
        [RequiredRole] NVARCHAR(20) NULL,
        [URLPattern] NVARCHAR(200) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'NoteIngredients')
BEGIN
    CREATE TABLE [NoteIngredients] (
        [BaseNoteID] INT NOT NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [ID] INT IDENTITY(1,1) NOT NULL,
        [NoteID] INT NOT NULL,
        [Percentage] FLOAT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'NoteInventory')
BEGIN
    CREATE TABLE [NoteInventory] (
        [InventoryID] INT IDENTITY(1,1) NOT NULL,
        [LastRestockDate] DATETIME2(7) NULL,
        [MinStockLevel] INT NULL,
        [NoteID] INT NOT NULL,
        [StockQuantity] INT NULL,
        [UpdatedAt] DATETIME2(7) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'OrderDetailNoteSelections')
BEGIN
    CREATE TABLE [OrderDetailNoteSelections] (
        [DetailID] INT NOT NULL,
        [NoteID] INT NOT NULL,
        [NoteType] NVARCHAR(20) NULL,
        [Percentage] INT NOT NULL,
        [SelectionID] INT IDENTITY(1,1) NOT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'OrderDetails')
BEGIN
    CREATE TABLE [OrderDetails] (
        [BaseNoteName] NVARCHAR(100) NULL,
        [BottleName] NVARCHAR(100) NULL,
        [CustomLabel] NVARCHAR(200) NULL,
        [DetailID] INT IDENTITY(1,1) NOT NULL,
        [MiddleNoteName] NVARCHAR(100) NULL,
        [OrderID] INT NOT NULL,
        [ProductID] INT NOT NULL,
        [ProductName] NVARCHAR(200) NULL,
        [Quantity] INT NOT NULL,
        [Subtotal] DECIMAL(19,4) NOT NULL,
        [TopNoteName] NVARCHAR(100) NULL,
        [UnitPrice] DECIMAL(19,4) NOT NULL,
        [VolumeML] INT NULL,
        [VolumeName] NVARCHAR(50) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'OrderIngredients')
BEGIN
    CREATE TABLE [OrderIngredients] (
        [CreatedAt] DATETIME2(7) NULL,
        [DetailID] INT NULL,
        [IngredientID] INT IDENTITY(1,1) NOT NULL,
        [IngredientName] NVARCHAR(100) NOT NULL,
        [OrderID] INT NOT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Orders')
BEGIN
    CREATE TABLE [Orders] (
        [ChannelSource] NVARCHAR(50) NULL,
        [CostAmount] DECIMAL(19,4) NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [DeliveredAt] DATETIME2(7) NULL,
        [Notes] NVARCHAR(MAX) NULL,
        [OrderID] INT IDENTITY(1,1) NOT NULL,
        [OrderNo] NVARCHAR(50) NOT NULL,
        [PaymentMethod] NVARCHAR(50) NULL,
        [ProfitAmount] DECIMAL(19,4) NULL,
        [RefundAmount] DECIMAL(19,4) NULL,
        [ShippedAt] DATETIME2(7) NULL,
        [ShippingAddress] NVARCHAR(200) NULL,
        [ShippingCity] NVARCHAR(50) NULL,
        [ShippingCompany] NVARCHAR(50) NULL,
        [ShippingFee] DECIMAL(19,4) NULL,
        [ShippingName] NVARCHAR(100) NULL,
        [ShippingNotes] NVARCHAR(MAX) NULL,
        [ShippingPhone] NVARCHAR(20) NULL,
        [ShippingPostalCode] NVARCHAR(20) NULL,
        [ShippingStatus] NVARCHAR(20) NULL,
        [Status] NVARCHAR(20) NULL,
        [TotalAmount] DECIMAL(19,4) NOT NULL,
        [TrackingNumber] NVARCHAR(100) NULL,
        [UpdatedAt] DATETIME2(7) NULL,
        [UserID] INT NOT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'PaymentRecords')
BEGIN
    CREATE TABLE [PaymentRecords] (
        [Amount] DECIMAL(19,4) NULL,
        [Category] NVARCHAR(50) NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [Fee] DECIMAL(19,4) NULL,
        [NetAmount] DECIMAL(19,4) NULL,
        [OrderID] INT NULL,
        [OrderNo] NVARCHAR(50) NULL,
        [PaymentMethod] NVARCHAR(50) NULL,
        [ReconcileStatus] NVARCHAR(20) NULL,
        [RecordID] INT IDENTITY(1,1) NOT NULL,
        [Remark] NVARCHAR(200) NULL,
        [Status] NVARCHAR(20) NULL,
        [TransactionNo] NVARCHAR(100) NULL,
        [TransactionType] NVARCHAR(20) NULL,
        [UpdatedAt] DATETIME2(7) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'PointTransactions')
BEGIN
    CREATE TABLE [PointTransactions] (
        [CreatedAt] DATETIME2(7) NULL,
        [CreatedBy] NVARCHAR(50) NULL,
        [Description] NVARCHAR(255) NULL,
        [OrderID] INT NULL,
        [Points] INT NOT NULL,
        [PointsChange] INT NULL,
        [Reason] NVARCHAR(200) NULL,
        [TransactionID] INT IDENTITY(1,1) NOT NULL,
        [TransactionType] NVARCHAR(20) NULL,
        [UserID] INT NOT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ProductBottleStyles')
BEGIN
    CREATE TABLE [ProductBottleStyles] (
        [BottleID] INT NOT NULL,
        [CustomPrice] DECIMAL(19,4) NULL,
        [ID] INT IDENTITY(1,1) NOT NULL,
        [ProductID] INT NOT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ProductCosts')
BEGIN
    CREATE TABLE [ProductCosts] (
        [CostID] INT IDENTITY(1,1) NOT NULL,
        [CostName] NVARCHAR(100) NULL,
        [CostType] NVARCHAR(20) NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [CreatedBy] NVARCHAR(50) NULL,
        [EffectiveDate] DATETIME2(7) NULL,
        [ExpiryDate] DATETIME2(7) NULL,
        [ProductID] INT NOT NULL,
        [Quantity] FLOAT NULL,
        [TotalCost] DECIMAL(19,4) NULL,
        [UnitCost] DECIMAL(19,4) NULL,
        [UpdatedAt] DATETIME2(7) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ProductInventory')
BEGIN
    CREATE TABLE [ProductInventory] (
        [InventoryID] INT IDENTITY(1,1) NOT NULL,
        [NoteID] INT NULL,
        [ProductID] INT NULL,
        [SafetyStock] INT NULL,
        [StockQty] INT NULL,
        [StockType] NVARCHAR(20) NULL,
        [UnitCost] DECIMAL(19,4) NULL,
        [UpdatedAt] DATETIME2(7) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ProductionLogs')
BEGIN
    CREATE TABLE [ProductionLogs] (
        [CreatedAt] DATETIME2(7) NULL,
        [CreatedBy] NVARCHAR(100) NULL,
        [LogID] INT IDENTITY(1,1) NOT NULL,
        [Notes] NVARCHAR(MAX) NULL,
        [ProductionID] INT NOT NULL,
        [Status] NVARCHAR(20) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ProductionOrders')
BEGIN
    CREATE TABLE [ProductionOrders] (
        [AssignedTo] NVARCHAR(100) NULL,
        [BottleIndex] INT NULL,
        [CompletedAt] DATETIME2(7) NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [DetailID] INT NULL,
        [EstimatedDate] DATETIME2(7) NULL,
        [Notes] NVARCHAR(MAX) NULL,
        [OrderID] INT NOT NULL,
        [Priority] INT NULL,
        [PriorityText] NVARCHAR(10) NULL,
        [ProductionID] INT IDENTITY(1,1) NOT NULL,
        [QCNotes] NVARCHAR(MAX) NULL,
        [QCPassedAt] DATETIME2(7) NULL,
        [RecipeID] INT NULL,
        [RecipeName] NVARCHAR(100) NULL,
        [ShippedOutAt] DATETIME2(7) NULL,
        [StartedAt] DATETIME2(7) NULL,
        [Status] NVARCHAR(20) NULL,
        [TotalBottles] INT NULL,
        [UpdatedAt] DATETIME2(7) NULL,
        [WarehouseInAt] DATETIME2(7) NULL,
        [WorkOrderNo] NVARCHAR(50) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ProductManufacturing')
BEGIN
    CREATE TABLE [ProductManufacturing] (
        [ActualQty] FLOAT NULL,
        [BatchNo] NVARCHAR(30) NULL,
        [CompletedAt] DATETIME2(7) NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [ManufacturingID] INT IDENTITY(1,1) NOT NULL,
        [Notes] NVARCHAR(MAX) NULL,
        [PlannedQty] FLOAT NULL,
        [ProductID] INT NULL,
        [ProductName] NVARCHAR(100) NULL,
        [ProductRecipeID] INT NULL,
        [StartedAt] DATETIME2(7) NULL,
        [Status] NVARCHAR(20) NULL,
        [TransferRequestID] INT NULL,
        [UpdatedAt] DATETIME2(7) NULL,
        [WorkCenter] NVARCHAR(20) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ProductManufacturingDetails')
BEGIN
    CREATE TABLE [ProductManufacturingDetails] (
        [ActualQty] FLOAT NULL,
        [DetailID] INT IDENTITY(1,1) NOT NULL,
        [ManufacturingID] INT NULL,
        [NoteID] INT NULL,
        [NoteName] NVARCHAR(100) NULL,
        [PlannedQty] FLOAT NULL,
        [TotalCost] DECIMAL(19,4) NULL,
        [UnitCost] DECIMAL(19,4) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ProductNoteRatios')
BEGIN
    CREATE TABLE [ProductNoteRatios] (
        [NoteID] INT NOT NULL,
        [NoteType] NVARCHAR(20) NULL,
        [Percentage] INT NOT NULL,
        [ProductID] INT NOT NULL,
        [RatioID] INT IDENTITY(1,1) NOT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ProductNotes')
BEGIN
    CREATE TABLE [ProductNotes] (
        [NoteID] INT NOT NULL,
        [ProductID] INT NOT NULL,
        [ProductNoteID] INT IDENTITY(1,1) NOT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ProductReviews')
BEGIN
    CREATE TABLE [ProductReviews] (
        [Comment] NVARCHAR(MAX) NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [OrderID] INT NOT NULL,
        [ProductID] INT NULL,
        [Rating] INT NULL,
        [ReviewID] INT IDENTITY(1,1) NOT NULL,
        [Status] NVARCHAR(20) NULL,
        [UpdatedAt] DATETIME2(7) NULL,
        [UserID] INT NOT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Products')
BEGIN
    CREATE TABLE [Products] (
        [BaseIngredients] NVARCHAR(MAX) NULL,
        [BasePrice] DECIMAL(19,4) NOT NULL,
        [BOMCost] DECIMAL(19,4) NULL,
        [Category] NVARCHAR(50) NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [Description] NVARCHAR(MAX) NULL,
        [Engravable] BIT NULL,
        [EngravingPrice] DECIMAL(19,4) NULL,
        [ImageURL] NVARCHAR(200) NULL,
        [IsActive] BIT NULL,
        [KOLID] INT NULL,
        [ProductID] INT IDENTITY(1,1) NOT NULL,
        [ProductName] NVARCHAR(100) NOT NULL,
        [ProductType] NVARCHAR(50) NULL,
        [RecipeID] INT NULL,
        [ReviewStatus] NVARCHAR(20) NULL,
        [UnitCost] DECIMAL(19,4) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ProductTypeConfig')
BEGIN
    CREATE TABLE [ProductTypeConfig] (
        [ConfigID] INT IDENTITY(1,1) NOT NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [Description] NVARCHAR(MAX) NULL,
        [DisplayName] NVARCHAR(50) NULL,
        [DisplayOrder] INT NULL,
        [Icon] NVARCHAR(100) NULL,
        [IsActive] BIT NULL,
        [NavName] NVARCHAR(50) NULL,
        [RequiresRatio] BIT NULL,
        [RequiresReview] BIT NULL,
        [TypeCode] NVARCHAR(20) NOT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ProductVolumePrices')
BEGIN
    CREATE TABLE [ProductVolumePrices] (
        [Price] DECIMAL(19,4) NOT NULL,
        [ProductID] INT NOT NULL,
        [PVPriceID] INT IDENTITY(1,1) NOT NULL,
        [VolumeID] INT NOT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'PurchaseCategories')
BEGIN
    CREATE TABLE [PurchaseCategories] (
        [CategoryCode] NVARCHAR(20) NULL,
        [CategoryID] INT IDENTITY(1,1) NOT NULL,
        [CategoryName] NVARCHAR(100) NULL,
        [Description] NVARCHAR(MAX) NULL,
        [DisplayOrder] INT NULL,
        [IsActive] BIT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'PurchaseCostReview')
BEGIN
    CREATE TABLE [PurchaseCostReview] (
        [CostAllocation] NVARCHAR(20) NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [PurchaseID] INT NULL,
        [ReviewAmount] DECIMAL(19,4) NULL,
        [ReviewComments] NVARCHAR(MAX) NULL,
        [ReviewedAt] DATETIME2(7) NULL,
        [ReviewerID] INT NULL,
        [ReviewID] INT IDENTITY(1,1) NOT NULL,
        [ReviewStatus] NVARCHAR(20) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'PurchaseOrderDetails')
BEGIN
    CREATE TABLE [PurchaseOrderDetails] (
        [DetailID] INT IDENTITY(1,1) NOT NULL,
        [ItemCode] NVARCHAR(50) NULL,
        [ItemName] NVARCHAR(200) NULL,
        [PurchaseID] INT NULL,
        [Quantity] FLOAT NULL,
        [ReceivedQty] FLOAT NULL,
        [Remarks] NVARCHAR(MAX) NULL,
        [Specification] NVARCHAR(200) NULL,
        [TotalPrice] DECIMAL(19,4) NULL,
        [Unit] NVARCHAR(20) NULL,
        [UnitPrice] DECIMAL(19,4) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'PurchaseOrders')
BEGIN
    CREATE TABLE [PurchaseOrders] (
        [ApprovedAt] DATETIME2(7) NULL,
        [ApprovedBy] INT NULL,
        [CategoryCode] NVARCHAR(20) NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [CreatedBy] INT NULL,
        [ExpectedDate] DATETIME2(7) NULL,
        [OrderDate] DATETIME2(7) NULL,
        [PurchaseID] INT IDENTITY(1,1) NOT NULL,
        [PurchaseNo] NVARCHAR(50) NULL,
        [Remarks] NVARCHAR(MAX) NULL,
        [Status] NVARCHAR(20) NULL,
        [SupplierID] INT NULL,
        [TotalAmount] DECIMAL(19,4) NULL,
        [UpdatedAt] DATETIME2(7) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'PurchaseReceiptDetails')
BEGIN
    CREATE TABLE [PurchaseReceiptDetails] (
        [AcceptedQty] FLOAT NULL,
        [MaterialID] INT NULL,
        [PurchaseDetailID] INT NULL,
        [ReceiptDetailID] INT IDENTITY(1,1) NOT NULL,
        [ReceiptID] INT NULL,
        [ReceivedQty] FLOAT NULL,
        [RejectedQty] FLOAT NULL,
        [RejectReason] NVARCHAR(200) NULL,
        [UnitPrice] DECIMAL(19,4) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'PurchaseReceipts')
BEGIN
    CREATE TABLE [PurchaseReceipts] (
        [CreatedAt] DATETIME2(7) NULL,
        [Notes] NVARCHAR(MAX) NULL,
        [PurchaseID] INT NULL,
        [ReceiptDate] DATETIME2(7) NULL,
        [ReceiptID] INT IDENTITY(1,1) NOT NULL,
        [ReceiptNo] NVARCHAR(50) NULL,
        [ReceivedBy] NVARCHAR(50) NULL,
        [Status] NVARCHAR(20) NULL,
        [SupplierID] INT NULL,
        [TotalReceivedQty] FLOAT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'RawMaterialInventory')
BEGIN
    CREATE TABLE [RawMaterialInventory] (
        [CategoryCode] NVARCHAR(20) NULL,
        [ItemCode] NVARCHAR(50) NULL,
        [ItemName] NVARCHAR(200) NULL,
        [LastPurchaseDate] DATETIME2(7) NULL,
        [MaterialID] INT IDENTITY(1,1) NOT NULL,
        [SafetyStock] FLOAT NULL,
        [StockQty] FLOAT NULL,
        [SupplierID] INT NULL,
        [Unit] NVARCHAR(20) NULL,
        [UnitPrice] DECIMAL(19,4) NULL,
        [UpdatedAt] DATETIME2(7) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'RecipeAccordMaterials')
BEGIN
    CREATE TABLE [RecipeAccordMaterials] (
        [AccordRecipeID] INT NULL,
        [DetailID] INT IDENTITY(1,1) NOT NULL,
        [MaterialID] INT NULL,
        [MaterialName] NVARCHAR(100) NULL,
        [Notes] NVARCHAR(MAX) NULL,
        [Percentage] FLOAT NULL,
        [PlannedQty] FLOAT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'RecipeAccords')
BEGIN
    CREATE TABLE [RecipeAccords] (
        [AccordRecipeID] INT IDENTITY(1,1) NOT NULL,
        [BatchSize] FLOAT NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [NoteID] INT NULL,
        [PublishedAt] DATETIME2(7) NULL,
        [PublishedBy] NVARCHAR(50) NULL,
        [RecipeID] INT NULL,
        [RecipeName] NVARCHAR(100) NULL,
        [Status] NVARCHAR(20) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'RecipeIngredients')
BEGIN
    CREATE TABLE [RecipeIngredients] (
        [ID] INT IDENTITY(1,1) NOT NULL,
        [IngredientName] NVARCHAR(100) NULL,
        [NoteID] INT NULL,
        [Percentage] FLOAT NULL,
        [RecipeID] INT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'RecipeNotes')
BEGIN
    CREATE TABLE [RecipeNotes] (
        [ID] INT IDENTITY(1,1) NOT NULL,
        [NoteID] INT NULL,
        [NoteType] NVARCHAR(20) NULL,
        [Percentage] INT NULL,
        [RecipeID] INT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'RecipePopularity')
BEGIN
    CREATE TABLE [RecipePopularity] (
        [FavoriteCount] INT NULL,
        [LastCalculatedAt] DATETIME2(7) NULL,
        [PopularityID] INT IDENTITY(1,1) NOT NULL,
        [ProductID] INT NOT NULL,
        [PurchaseCount] INT NULL,
        [ViewCount] INT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'RecipeProductNotes')
BEGIN
    CREATE TABLE [RecipeProductNotes] (
        [DetailID] INT IDENTITY(1,1) NOT NULL,
        [NoteID] INT NULL,
        [NoteName] NVARCHAR(100) NULL,
        [Notes] NVARCHAR(MAX) NULL,
        [Percentage] FLOAT NULL,
        [PlannedQty] FLOAT NULL,
        [ProductRecipeID] INT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'RecipeProducts')
BEGIN
    CREATE TABLE [RecipeProducts] (
        [BatchSize] FLOAT NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [ProductID] INT NULL,
        [ProductRecipeID] INT IDENTITY(1,1) NOT NULL,
        [PublishedAt] DATETIME2(7) NULL,
        [PublishedBy] NVARCHAR(50) NULL,
        [RecipeID] INT NULL,
        [Status] NVARCHAR(20) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'RecipePublishLog')
BEGIN
    CREATE TABLE [RecipePublishLog] (
        [IPAddress] NVARCHAR(50) NULL,
        [LogID] INT IDENTITY(1,1) NOT NULL,
        [PublishedAt] DATETIME2(7) NULL,
        [PublishedBy] NVARCHAR(50) NULL,
        [PublishType] NVARCHAR(20) NULL,
        [RecipeID] INT NULL,
        [TargetRecipeID] INT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Recipes')
BEGIN
    CREATE TABLE [Recipes] (
        [CreatedAt] DATETIME2(7) NULL,
        [CreatedBy] NVARCHAR(100) NULL,
        [Description] NVARCHAR(MAX) NULL,
        [IsActive] BIT NULL,
        [ProductType] NVARCHAR(20) NULL,
        [RecipeCode] NVARCHAR(50) NULL,
        [RecipeID] INT IDENTITY(1,1) NOT NULL,
        [RecipeName] NVARCHAR(100) NULL,
        [ReviewStatus] NVARCHAR(20) NULL,
        [UpdatedAt] DATETIME2(7) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'RecommendedRecipes')
BEGIN
    CREATE TABLE [RecommendedRecipes] (
        [CreatedAt] DATETIME2(7) NULL,
        [Description] NVARCHAR(MAX) NULL,
        [IsActive] BIT NULL,
        [ProductID] INT NULL,
        [RecipeID] INT IDENTITY(1,1) NOT NULL,
        [RecipeName] NVARCHAR(200) NULL,
        [SortOrder] INT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ReconciliationLogs')
BEGIN
    CREATE TABLE [ReconciliationLogs] (
        [CreatedAt] DATETIME2(7) NULL,
        [Difference] DECIMAL(19,4) NULL,
        [LogID] INT IDENTITY(1,1) NOT NULL,
        [OrderAmount] DECIMAL(19,4) NULL,
        [OrderID] INT NULL,
        [OrderNo] NVARCHAR(50) NULL,
        [PaymentAmount] DECIMAL(19,4) NULL,
        [ReconcileDate] DATETIME2(7) NULL,
        [Resolution] NVARCHAR(MAX) NULL,
        [ResolvedAt] DATETIME2(7) NULL,
        [ResolvedBy] NVARCHAR(50) NULL,
        [Status] NVARCHAR(20) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'RefundRecords')
BEGIN
    CREATE TABLE [RefundRecords] (
        [ApprovedAt] DATETIME2(7) NULL,
        [ApprovedBy] NVARCHAR(50) NULL,
        [CompletedAt] DATETIME2(7) NULL,
        [CostWriteBack] BIT NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [OrderID] INT NOT NULL,
        [OrderNo] NVARCHAR(50) NULL,
        [RefundAmount] DECIMAL(19,4) NOT NULL,
        [RefundID] INT IDENTITY(1,1) NOT NULL,
        [RefundNo] NVARCHAR(50) NULL,
        [RefundReason] NVARCHAR(MAX) NULL,
        [Status] NVARCHAR(20) NULL,
        [UpdatedAt] DATETIME2(7) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'SiteSettings')
BEGIN
    CREATE TABLE [SiteSettings] (
        [Description] NVARCHAR(255) NULL,
        [SettingKey] NVARCHAR(50) NULL,
        [SettingName] NVARCHAR(100) NULL,
        [SettingValue] NVARCHAR(255) NULL,
        [UpdatedAt] DATETIME2(7) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'SupplierPrices')
BEGIN
    CREATE TABLE [SupplierPrices] (
        [CreatedAt] DATETIME2(7) NULL,
        [EffectiveDate] DATETIME2(7) NULL,
        [ExpiryDate] DATETIME2(7) NULL,
        [IsActive] BIT NULL,
        [ItemCode] NVARCHAR(50) NULL,
        [ItemName] NVARCHAR(200) NULL,
        [MinOrderQty] FLOAT NULL,
        [PriceID] INT IDENTITY(1,1) NOT NULL,
        [SupplierID] INT NULL,
        [UnitPrice] DECIMAL(19,4) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Suppliers')
BEGIN
    CREATE TABLE [Suppliers] (
        [Address] NVARCHAR(255) NULL,
        [Category] NVARCHAR(50) NULL,
        [ContactPerson] NVARCHAR(50) NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [Email] NVARCHAR(100) NULL,
        [IsActive] BIT NULL,
        [Notes] NVARCHAR(MAX) NULL,
        [Phone] NVARCHAR(30) NULL,
        [SupplierID] INT IDENTITY(1,1) NOT NULL,
        [SupplierName] NVARCHAR(100) NOT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'UserAddresses')
BEGIN
    CREATE TABLE [UserAddresses] (
        [Address] NVARCHAR(200) NOT NULL,
        [AddressID] INT IDENTITY(1,1) NOT NULL,
        [City] NVARCHAR(50) NULL,
        [Consignee] NVARCHAR(50) NOT NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [District] NVARCHAR(50) NULL,
        [IsDefault] BIT NULL,
        [Phone] NVARCHAR(20) NOT NULL,
        [Province] NVARCHAR(50) NULL,
        [UserID] INT NOT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'UserFavorites')
BEGIN
    CREATE TABLE [UserFavorites] (
        [CreatedTime] DATETIME2(7) NULL,
        [FavoriteID] INT IDENTITY(1,1) NOT NULL,
        [ProductID] INT NOT NULL,
        [UserID] INT NOT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'UserPoints')
BEGIN
    CREATE TABLE [UserPoints] (
        [AvailablePoints] INT NULL,
        [ExpiredPoints] INT NULL,
        [LastUpdatedAt] DATETIME2(7) NULL,
        [PointID] INT IDENTITY(1,1) NOT NULL,
        [TotalPoints] INT NULL,
        [UsedPoints] INT NULL,
        [UserID] INT NOT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'UserPreferences')
BEGIN
    CREATE TABLE [UserPreferences] (
        [CreatedAt] DATETIME2(7) NULL,
        [PreferenceID] INT IDENTITY(1,1) NOT NULL,
        [PreferredBaseNotes] NVARCHAR(255) NULL,
        [PreferredCategories] NVARCHAR(255) NULL,
        [PreferredMiddleNotes] NVARCHAR(255) NULL,
        [PreferredTopNotes] NVARCHAR(255) NULL,
        [UpdatedAt] DATETIME2(7) NULL,
        [UserID] INT NOT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Users')
BEGIN
    CREATE TABLE [Users] (
        [Address] NVARCHAR(200) NULL,
        [City] NVARCHAR(50) NULL,
        [CreatedAt] DATETIME2(7) NULL,
        [Email] NVARCHAR(100) NOT NULL,
        [FullName] NVARCHAR(100) NULL,
        [IsActive] BIT NULL,
        [IsVIP] BIT NULL,
        [Password] NVARCHAR(255) NOT NULL,
        [Phone] NVARCHAR(20) NULL,
        [Points] INT NULL,
        [PostalCode] NVARCHAR(20) NULL,
        [UserID] INT IDENTITY(1,1) NOT NULL,
        [Username] NVARCHAR(50) NOT NULL,
        [UserRole] NVARCHAR(20) NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Volumes')
BEGIN
    CREATE TABLE [Volumes] (
        [IsActive] BIT NULL,
        [PriceMultiplier] FLOAT NULL,
        [VolumeID] INT IDENTITY(1,1) NOT NULL,
        [VolumeML] INT NOT NULL,
        [VolumeName] NVARCHAR(50) NOT NULL
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'WorkshopTransfer')
BEGIN
    CREATE TABLE [WorkshopTransfer] (
        [CreatedAt] DATETIME2(7) NULL,
        [FromWorkshop] NVARCHAR(20) NULL,
        [FulfilledAt] DATETIME2(7) NULL,
        [NoteID] INT NULL,
        [Notes] NVARCHAR(MAX) NULL,
        [RequestedAt] DATETIME2(7) NULL,
        [RequestedBy] NVARCHAR(50) NULL,
        [RequestQty] FLOAT NULL,
        [Status] NVARCHAR(20) NULL,
        [ToWorkshop] NVARCHAR(20) NULL,
        [TransferID] INT IDENTITY(1,1) NOT NULL,
        [TransferNo] NVARCHAR(30) NULL
    );
END
GO

-- ============================================
-- 性能优化索引
-- ============================================

-- Orders 表索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Orders_UserID')
    CREATE NONCLUSTERED INDEX [IX_Orders_UserID] ON [Orders]([UserID]) INCLUDE ([OrderID],[TotalAmount],[Status],[CreatedAt]);
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Orders_Status')
    CREATE NONCLUSTERED INDEX [IX_Orders_Status] ON [Orders]([Status]) INCLUDE ([OrderID],[TotalAmount],[CreatedAt]);
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Orders_CreatedAt')
    CREATE NONCLUSTERED INDEX [IX_Orders_CreatedAt] ON [Orders]([CreatedAt]) INCLUDE ([OrderID],[TotalAmount],[Status]);
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Orders_OrderNo')
    CREATE NONCLUSTERED INDEX [IX_Orders_OrderNo] ON [Orders]([OrderNo]);
GO

-- OrderDetails 表索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_OrderDetails_OrderID')
    CREATE NONCLUSTERED INDEX [IX_OrderDetails_OrderID] ON [OrderDetails]([OrderID]) INCLUDE ([ProductID],[Quantity],[Subtotal]);
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_OrderDetails_ProductID')
    CREATE NONCLUSTERED INDEX [IX_OrderDetails_ProductID] ON [OrderDetails]([ProductID]) INCLUDE ([OrderID],[Quantity],[Subtotal]);
GO

-- Products 表索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Products_ProductType')
    CREATE NONCLUSTERED INDEX [IX_Products_ProductType] ON [Products]([ProductType]) INCLUDE ([ProductID],[ProductName],[BasePrice],[UnitCost],[IsActive]);
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Products_IsActive')
    CREATE NONCLUSTERED INDEX [IX_Products_IsActive] ON [Products]([IsActive]) INCLUDE ([ProductID],[ProductName],[ProductType],[BasePrice]);
GO

-- Users 表索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Users_Username')
    CREATE NONCLUSTERED INDEX [IX_Users_Username] ON [Users]([Username]);
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Users_Email')
    CREATE NONCLUSTERED INDEX [IX_Users_Email] ON [Users]([Email]);
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Users_CreatedAt')
    CREATE NONCLUSTERED INDEX [IX_Users_CreatedAt] ON [Users]([CreatedAt]) INCLUDE ([UserID]);
GO

-- ProductReviews 表索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_ProductReviews_ProductID')
    CREATE NONCLUSTERED INDEX [IX_ProductReviews_ProductID] ON [ProductReviews]([ProductID],[Status]) INCLUDE ([Rating],[UserID]);
GO

-- UserFavorites 表索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_UserFavorites_ProductID')
    CREATE NONCLUSTERED INDEX [IX_UserFavorites_ProductID] ON [UserFavorites]([ProductID]) INCLUDE ([UserID]);
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_UserFavorites_UserID')
    CREATE NONCLUSTERED INDEX [IX_UserFavorites_UserID] ON [UserFavorites]([UserID]) INCLUDE ([ProductID]);
GO

-- AdminLogs 表索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_AdminLogs_CreatedAt')
    CREATE NONCLUSTERED INDEX [IX_AdminLogs_CreatedAt] ON [AdminLogs]([CreatedAt]) INCLUDE ([AdminID],[ActionType],[ModuleCode]);
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_AdminLogs_ModuleCode')
    CREATE NONCLUSTERED INDEX [IX_AdminLogs_ModuleCode] ON [AdminLogs]([ModuleCode],[CreatedAt]);
GO

-- RawMaterialInventory 表索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_RawMaterialInventory_ItemCode')
    CREATE NONCLUSTERED INDEX [IX_RawMaterialInventory_ItemCode] ON [RawMaterialInventory]([ItemCode]) INCLUDE ([MaterialID],[ItemName],[StockQty],[UnitPrice],[SafetyStock]);
GO

-- NoteInventory 表索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_NoteInventory_NoteID')
    CREATE NONCLUSTERED INDEX [IX_NoteInventory_NoteID] ON [NoteInventory]([NoteID]) INCLUDE ([StockQuantity]);
GO

-- FragranceNotes 表索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_FragranceNotes_NoteType')
    CREATE NONCLUSTERED INDEX [IX_FragranceNotes_NoteType] ON [FragranceNotes]([NoteType]) INCLUDE ([NoteID],[NoteName],[PriceAddition]);
GO

-- ProductInventory 表索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_ProductInventory_ProductID')
    CREATE NONCLUSTERED INDEX [IX_ProductInventory_ProductID] ON [ProductInventory]([ProductID]) INCLUDE ([StockQty],[UnitCost]);
GO

-- SupplierPrices 表索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_SupplierPrices_ItemCode')
    CREATE NONCLUSTERED INDEX [IX_SupplierPrices_ItemCode] ON [SupplierPrices]([ItemCode],[IsActive]) INCLUDE ([UnitPrice],[CreatedAt]);
GO

-- ProductionOrders 表索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_ProductionOrders_Status')
    CREATE NONCLUSTERED INDEX [IX_ProductionOrders_Status] ON [ProductionOrders]([Status]) INCLUDE ([ProductionID],[CreatedAt]);
GO

-- PurchaseOrders 表索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_PurchaseOrders_Status')
    CREATE NONCLUSTERED INDEX [IX_PurchaseOrders_Status] ON [PurchaseOrders]([Status]) INCLUDE ([PurchaseID],[TotalAmount]);
GO

-- RecipeAccordMaterials 表索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_RecipeAccordMaterials_AccordRecipeID')
    CREATE NONCLUSTERED INDEX [IX_RecipeAccordMaterials_AccordRecipeID] ON [RecipeAccordMaterials]([AccordRecipeID]) INCLUDE ([MaterialID],[PlannedQty]);
GO

-- RecipeAccords 表索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_RecipeAccords_NoteID')
    CREATE NONCLUSTERED INDEX [IX_RecipeAccords_NoteID] ON [RecipeAccords]([NoteID]) INCLUDE ([AccordRecipeID],[RecipeID]);
GO

-- ProductNoteRatios 表索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_ProductNoteRatios_ProductID')
    CREATE NONCLUSTERED INDEX [IX_ProductNoteRatios_ProductID] ON [ProductNoteRatios]([ProductID]) INCLUDE ([NoteID],[Percentage]);
GO

PRINT '数据库索引创建完成';
GO