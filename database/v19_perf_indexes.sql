-- ============================================
-- V19 新增性能索引 (EF Core DbContext 中新增但数据库中可能缺失)
-- 基于 v18_perf_indexes.sql 的复合索引策略
-- 使用 IF NOT EXISTS 确保幂等安全
-- ============================================

PRINT '=== V19 Performance Index Migration ==='
PRINT '开始时间: ' + CONVERT(VARCHAR, GETDATE(), 120)
GO

-- ============================================
-- 1. Products 新增复合索引 (V19 DbContext Fluent API)
-- ============================================

-- 1.1 产品类型 + 活跃状态
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Products_ProductType_IsActive')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Products_ProductType_IsActive]
    ON [dbo].[Products] ([ProductType], [IsActive])
    INCLUDE ([ProductID], [ProductName], [Category], [BasePrice], [ImageURL])
    WITH (FILLFACTOR = 90);
    PRINT '  + IX_Products_ProductType_IsActive 创建成功'
END
ELSE
    PRINT '  o IX_Products_ProductType_IsActive 已存在'

-- 1.2 分类 + 活跃状态
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Products_Category_IsActive')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Products_Category_IsActive]
    ON [dbo].[Products] ([Category], [IsActive])
    INCLUDE ([ProductID], [ProductName], [ProductType], [BasePrice], [ImageURL])
    WITH (FILLFACTOR = 90);
    PRINT '  + IX_Products_Category_IsActive 创建成功'
END
ELSE
    PRINT '  o IX_Products_Category_IsActive 已存在'

-- 1.3 活跃 + 创建时间（降序，覆盖排序）
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Products_CreatedAt_Active')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Products_CreatedAt_Active]
    ON [dbo].[Products] ([IsActive], [CreatedAt] DESC)
    INCLUDE ([ProductID], [ProductName], [Category], [ProductType], [BasePrice])
    WITH (FILLFACTOR = 90);
    PRINT '  + IX_Products_CreatedAt_Active 创建成功'
END
ELSE
    PRINT '  o IX_Products_CreatedAt_Active 已存在'

-- 1.4 KOL 商品过滤索引
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Products_KOL_Active')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Products_KOL_Active]
    ON [dbo].[Products] ([ProductType], [IsActive], [KOLID])
    WHERE [ProductType] = 'KOL' AND [IsActive] <> 0
    WITH (FILLFACTOR = 90);
    PRINT '  + IX_Products_KOL_Active (Filtered) 创建成功'
END
ELSE
    PRINT '  o IX_Products_KOL_Active 已存在'

-- 1.5 产品名称搜索
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Products_ProductName')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Products_ProductName]
    ON [dbo].[Products] ([ProductName])
    INCLUDE ([ProductID], [Category], [ProductType], [IsActive])
    WITH (FILLFACTOR = 90);
    PRINT '  + IX_Products_ProductName 创建成功'
END
ELSE
    PRINT '  o IX_Products_ProductName 已存在'

GO

-- ============================================
-- 2. Orders 新增复合索引
-- ============================================

-- 2.1 用户ID + 创建时间（用户订单列表）
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Orders_UserID_CreatedAt')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Orders_UserID_CreatedAt]
    ON [dbo].[Orders] ([UserID], [CreatedAt] DESC)
    INCLUDE ([OrderID], [Status], [TotalAmount], [PaymentMethod])
    WITH (FILLFACTOR = 90);
    PRINT '  + IX_Orders_UserID_CreatedAt 创建成功'
END
ELSE
    PRINT '  o IX_Orders_UserID_CreatedAt 已存在'

-- 2.2 状态 + 创建时间（后台管理排序）
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Orders_Status_CreatedAt')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Orders_Status_CreatedAt]
    ON [dbo].[Orders] ([Status], [CreatedAt] DESC)
    INCLUDE ([OrderID], [UserID], [TotalAmount], [PaymentMethod], [ShippingName])
    WITH (FILLFACTOR = 90);
    PRINT '  + IX_Orders_Status_CreatedAt 创建成功'
END
ELSE
    PRINT '  o IX_Orders_Status_CreatedAt 已存在'

GO

-- ============================================
-- 3. OrderItems 新增复合索引
-- ============================================

-- 3.1 订单ID + 产品ID
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_OrderItems_OrderID_ProductID')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_OrderItems_OrderID_ProductID]
    ON [dbo].[OrderItems] ([OrderID], [ProductID])
    INCLUDE ([Quantity], [UnitPrice])
    WITH (FILLFACTOR = 90);
    PRINT '  + IX_OrderItems_OrderID_ProductID 创建成功'
END
ELSE
    PRINT '  o IX_OrderItems_OrderID_ProductID 已存在'

GO

-- ============================================
-- 4. AppLogs 新增过滤索引
-- ============================================

-- 4.1 日志级别过滤索引（仅ERROR/WARN）
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_AppLogs_Level_Date')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_AppLogs_Level_Date]
    ON [dbo].[AppLogs] ([LogLevel], [CreatedAt] DESC)
    WHERE [LogLevel] IN ('ERROR', 'WARN')
    WITH (FILLFACTOR = 90);
    PRINT '  + IX_AppLogs_Level_Date (Filtered) 创建成功'
END
ELSE
    PRINT '  o IX_AppLogs_Level_Date 已存在'

GO

-- ============================================
-- 5. 统计信息更新
-- ============================================

PRINT '更新表统计信息...'

UPDATE STATISTICS [dbo].[Products] WITH FULLSCAN;
PRINT '  + Products 统计信息已更新'

UPDATE STATISTICS [dbo].[Orders] WITH FULLSCAN;
PRINT '  + Orders 统计信息已更新'

UPDATE STATISTICS [dbo].[OrderItems] WITH FULLSCAN;
PRINT '  + OrderItems 统计信息已更新'

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = 'AppLogs')
BEGIN
    UPDATE STATISTICS [dbo].[AppLogs] WITH FULLSCAN;
    PRINT '  + AppLogs 统计信息已更新'
END

GO

-- ============================================
-- 6. 验证索引状态
-- ============================================

PRINT ''
PRINT '=== V19 新增索引状态 ==='
SELECT 
    t.name AS TableName,
    i.name AS IndexName,
    i.type_desc AS IndexType,
    i.has_filter AS IsFiltered,
    STRING_AGG(c.name, ', ') WITHIN GROUP (ORDER BY ic.key_ordinal) AS KeyColumns
FROM sys.indexes i
JOIN sys.tables t ON i.object_id = t.object_id
JOIN sys.index_columns ic ON i.index_id = ic.index_id AND i.object_id = ic.object_id
JOIN sys.columns c ON ic.column_id = c.column_id AND ic.object_id = c.object_id
WHERE i.name IN (
    'IX_Products_ProductType_IsActive', 'IX_Products_Category_IsActive',
    'IX_Products_CreatedAt_Active', 'IX_Products_KOL_Active', 'IX_Products_ProductName',
    'IX_Orders_UserID_CreatedAt', 'IX_Orders_Status_CreatedAt',
    'IX_OrderItems_OrderID_ProductID',
    'IX_AppLogs_Level_Date'
) AND i.is_primary_key = 0
GROUP BY t.name, i.name, i.type_desc, i.has_filter
ORDER BY t.name, i.name;

PRINT ''
PRINT '=== V19 性能索引迁移完成 ==='
PRINT '结束时间: ' + CONVERT(VARCHAR, GETDATE(), 120)
GO
