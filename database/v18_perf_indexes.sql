-- ============================================
-- V18.0 数据库性能索引 (Performance Indexes)
-- 适用于 SQL Server
-- 执行方式: sqlcmd -S <server> -d <database> -i v18_perf_indexes.sql
-- ============================================

PRINT '=== V18.0 Performance Index Migration ==='
PRINT '开始时间: ' + CONVERT(VARCHAR, GETDATE(), 120)
GO

-- ============================================
-- 1. Products 表索引优化
-- ============================================

-- 1.1 产品类型 + 活跃状态复合索引
-- 用于: 首页栏目查询 (WHERE ProductType='xxx' AND IsActive<>0)
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Products_ProductType_IsActive')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Products_ProductType_IsActive]
    ON [dbo].[Products] ([ProductType], [IsActive])
    INCLUDE ([ProductID], [ProductName], [Category], [BasePrice], [ImageURL], [Description], [CreatedAt])
    WITH (ONLINE = ON, FILLFACTOR = 90);
    PRINT '  ✓ IX_Products_ProductType_IsActive 创建成功'
END
ELSE
    PRINT '  → IX_Products_ProductType_IsActive 已存在，跳过'

-- 1.2 分类 + 活跃状态复合索引
-- 用于: 分类筛选查询 (WHERE Category='花香调' AND IsActive<>0)
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Products_Category_IsActive')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Products_Category_IsActive]
    ON [dbo].[Products] ([Category], [IsActive])
    INCLUDE ([ProductID], [ProductName], [ProductType], [BasePrice], [ImageURL])
    WITH (ONLINE = ON, FILLFACTOR = 90);
    PRINT '  ✓ IX_Products_Category_IsActive 创建成功'
END
ELSE
    PRINT '  → IX_Products_Category_IsActive 已存在，跳过'

-- 1.3 产品创建时间索引（覆盖排序）
-- 用于: ORDER BY CreatedAt DESC 查询
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Products_CreatedAt_Active')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Products_CreatedAt_Active]
    ON [dbo].[Products] ([IsActive], [CreatedAt] DESC)
    INCLUDE ([ProductID], [ProductName], [Category], [ProductType], [BasePrice])
    WITH (ONLINE = ON, FILLFACTOR = 90);
    PRINT '  ✓ IX_Products_CreatedAt_Active 创建成功'
END
ELSE
    PRINT '  → IX_Products_CreatedAt_Active 已存在，跳过'

-- 1.4 KOL 商品索引
-- 用于: KOL 类型商品快速检索
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Products_KOL_Active')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Products_KOL_Active]
    ON [dbo].[Products] ([ProductType], [IsActive], [KOLID])
    WHERE [ProductType] = 'KOL' AND [IsActive] <> 0
    WITH (FILLFACTOR = 90);
    PRINT '  ✓ IX_Products_KOL_Active (Filtered) 创建成功'
END
ELSE
    PRINT '  → IX_Products_KOL_Active 已存在，跳过'

-- 1.5 产品名称搜索索引
-- 用于: LIKE '%keyword%' 搜索（注: 中文全文搜索建议使用全文索引）
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Products_ProductName')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Products_ProductName]
    ON [dbo].[Products] ([ProductName])
    INCLUDE ([ProductID], [Category], [ProductType], [IsActive])
    WITH (ONLINE = ON, FILLFACTOR = 90);
    PRINT '  ✓ IX_Products_ProductName 创建成功'
END
ELSE
    PRINT '  → IX_Products_ProductName 已存在，跳过'

GO

-- ============================================
-- 2. Orders 表索引优化
-- ============================================

-- 2.1 用户ID + 创建时间复合索引
-- 用于: 用户订单列表 (WHERE UserID=xxx ORDER BY CreatedAt DESC)
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Orders_UserID_CreatedAt')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Orders_UserID_CreatedAt]
    ON [dbo].[Orders] ([UserID], [CreatedAt] DESC)
    INCLUDE ([OrderID], [Status], [TotalAmount], [PaymentMethod])
    WITH (ONLINE = ON, FILLFACTOR = 90);
    PRINT '  ✓ IX_Orders_UserID_CreatedAt 创建成功'
END
ELSE
    PRINT '  → IX_Orders_UserID_CreatedAt 已存在，跳过'

-- 2.2 状态 + 创建时间复合索引
-- 用于: 后台订单管理 (WHERE Status='pending' ORDER BY CreatedAt)
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Orders_Status_CreatedAt')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_Orders_Status_CreatedAt]
    ON [dbo].[Orders] ([Status], [CreatedAt] DESC)
    INCLUDE ([OrderID], [UserID], [TotalAmount], [PaymentMethod], [ShippingName])
    WITH (ONLINE = ON, FILLFACTOR = 90);
    PRINT '  ✓ IX_Orders_Status_CreatedAt 创建成功'
END
ELSE
    PRINT '  → IX_Orders_Status_CreatedAt 已存在，跳过'

-- 2.3 订单号唯一索引（确保已存在）
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Orders_OrderNumber')
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX [IX_Orders_OrderNumber]
    ON [dbo].[Orders] ([OrderNumber])
    WITH (ONLINE = ON);
    PRINT '  ✓ IX_Orders_OrderNumber (Unique) 创建成功'
END
ELSE
    PRINT '  → IX_Orders_OrderNumber 已存在，跳过'

GO

-- ============================================
-- 3. OrderItems 索引优化
-- ============================================

-- 3.1 订单ID + 产品ID 复合索引
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_OrderItems_OrderID_ProductID')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_OrderItems_OrderID_ProductID]
    ON [dbo].[OrderItems] ([OrderID], [ProductID])
    INCLUDE ([Quantity], [Price], [ProductName])
    WITH (ONLINE = ON, FILLFACTOR = 90);
    PRINT '  ✓ IX_OrderItems_OrderID_ProductID 创建成功'
END
ELSE
    PRINT '  → IX_OrderItems_OrderID_ProductID 已存在，跳过'

GO

-- ============================================
-- 4. AppLogs 索引优化
-- ============================================

-- 4.1 创建时间 + 日志类型复合索引
-- 用于: 后台日志查询 (WHERE LogType='xxx' ORDER BY CreatedAt DESC)
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_AppLogs_CreatedAt_LogType')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_AppLogs_CreatedAt_LogType]
    ON [dbo].[AppLogs] ([CreatedAt] DESC, [LogType])
    INCLUDE ([UserID], [Message], [IPAddress])
    WITH (ONLINE = ON, FILLFACTOR = 90);
    PRINT '  ✓ IX_AppLogs_CreatedAt_LogType 创建成功'
END
ELSE
    PRINT '  → IX_AppLogs_CreatedAt_LogType 已存在，跳过'

-- 4.2 日志级别索引
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_AppLogs_Level_Date')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_AppLogs_Level_Date]
    ON [dbo].[AppLogs] ([LogLevel], [CreatedAt] DESC)
    WHERE [LogLevel] IN ('ERROR', 'WARN')
    WITH (FILLFACTOR = 90);
    PRINT '  ✓ IX_AppLogs_Level_Date (Filtered) 创建成功'
END
ELSE
    PRINT '  → IX_AppLogs_Level_Date 已存在，跳过'

GO

-- ============================================
-- 5. UserBehavior 表索引
-- ============================================

-- 5.1 用户行为记录索引
IF EXISTS (SELECT 1 FROM sys.tables WHERE name = 'UserBehavior')
BEGIN
    IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_UserBehavior_UserID_Time')
    BEGIN
        CREATE NONCLUSTERED INDEX [IX_UserBehavior_UserID_Time]
        ON [dbo].[UserBehavior] ([UserID], [CreatedAt] DESC)
        INCLUDE ([BehaviorType], [TargetID], [TargetType])
        WITH (ONLINE = ON, FILLFACTOR = 90);
        PRINT '  ✓ IX_UserBehavior_UserID_Time 创建成功'
    END
    ELSE
        PRINT '  → IX_UserBehavior_UserID_Time 已存在，跳过'
END
ELSE
    PRINT '  → UserBehavior 表不存在，跳过其索引'

GO

-- ============================================
-- 6. 统计信息更新
-- ============================================

PRINT '更新表统计信息...'

UPDATE STATISTICS [dbo].[Products] WITH FULLSCAN;
PRINT '  ✓ Products 统计信息已更新'

UPDATE STATISTICS [dbo].[Orders] WITH FULLSCAN;
PRINT '  ✓ Orders 统计信息已更新'

UPDATE STATISTICS [dbo].[OrderItems] WITH FULLSCAN;
PRINT '  ✓ OrderItems 统计信息已更新'

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = 'AppLogs')
BEGIN
    UPDATE STATISTICS [dbo].[AppLogs] WITH FULLSCAN;
    PRINT '  ✓ AppLogs 统计信息已更新'
END

IF EXISTS (SELECT 1 FROM sys.tables WHERE name = 'UserBehavior')
BEGIN
    UPDATE STATISTICS [dbo].[UserBehavior] WITH FULLSCAN;
    PRINT '  ✓ UserBehavior 统计信息已更新'
END

GO

-- ============================================
-- 7. 索引使用建议（查询）
-- ============================================

PRINT ''
PRINT '=== 当前数据库索引状态 ==='
SELECT 
    t.name AS TableName,
    i.name AS IndexName,
    i.type_desc AS IndexType,
    STRING_AGG(c.name, ', ') WITHIN GROUP (ORDER BY ic.key_ordinal) AS KeyColumns
FROM sys.indexes i
JOIN sys.tables t ON i.object_id = t.object_id
JOIN sys.index_columns ic ON i.index_id = ic.index_id AND i.object_id = ic.object_id
JOIN sys.columns c ON ic.column_id = c.column_id AND ic.object_id = c.object_id
WHERE i.name LIKE 'IX_%' AND i.is_primary_key = 0
GROUP BY t.name, i.name, i.type_desc
ORDER BY t.name, i.name;

PRINT ''
PRINT '=== V18.0 性能索引迁移完成 ==='
PRINT '结束时间: ' + CONVERT(VARCHAR, GETDATE(), 120)
GO
