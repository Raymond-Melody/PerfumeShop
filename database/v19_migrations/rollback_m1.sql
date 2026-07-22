-- =============================================================================
-- PerfumeShop V19 - M1 里程碑回滚脚本
-- 文件: database/v19_migrations/rollback_m1.sql
-- 编码: UTF-8 (无 BOM，兼容 sqlcmd -i 执行)
-- =============================================================================
-- 警告: 回滚将永久删除 M1 迁移产生的所有数据库对象和数据！
--       - ReviewImages 表及其中的所有图片记录
--       - Orders 表的优惠券/积分字段及其中存储的数据
--       - ProductReviews 的扩展字段及其中存储的数据
--       请在 M1 功能验证失败时才能执行此脚本。
-- =============================================================================
-- 前置条件:
--   1. 停止所有应用程序（避免回滚期间有并发写入）
--   2. 确认已有 PreM1 备份文件（PerfumeShop_PreM1_*.bak）
-- =============================================================================
-- 执行方式:
--   sqlcmd -S "localhost\YOURPERFUME" -d PerfumeShop -E -C -i "database\v19_migrations\rollback_m1.sql"
-- =============================================================================
-- 幂等: 所有语句使用 IF EXISTS 包裹，可重复执行不报错
-- =============================================================================

SET NOCOUNT ON;
GO

-- =============================================================================
-- [R1-05] 删除 ReviewImages.ReviewId 索引
-- 回滚 M1-05
-- =============================================================================
IF EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_ReviewImages_ReviewId' AND object_id = OBJECT_ID(N'[dbo].[ReviewImages]'))
BEGIN
    DROP INDEX [IX_ReviewImages_ReviewId] ON [dbo].[ReviewImages];
    PRINT '[R1-05] IX_ReviewImages_ReviewId 索引已删除';
END
ELSE
BEGIN
    PRINT '[R1-05] IX_ReviewImages_ReviewId 索引不存在，跳过';
END
GO

-- =============================================================================
-- [R1-04] 删除 ReviewImages 表
-- 回滚 M1-04。外键约束会随表一起删除。
-- 注意: 表中所有图片记录将被永久丢失。
-- =============================================================================
IF EXISTS (SELECT 1 FROM sys.tables WHERE name = N'ReviewImages' AND schema_id = SCHEMA_ID(N'dbo'))
BEGIN
    DROP TABLE [dbo].[ReviewImages];
    PRINT '[R1-04] ReviewImages 表已删除';
END
ELSE
BEGIN
    PRINT '[R1-04] ReviewImages 表不存在，跳过';
END
GO

-- =============================================================================
-- [R1-03] 删除 ProductReviews 扩展字段
-- 回滚 M1-03: Title, IsVerifiedPurchase, AIFeelingSummary, LikeCount
-- 注意: 带 DEFAULT 约束的列需先删除约束再删除列，数据将永久丢失。
-- =============================================================================

-- LikeCount (有 DEFAULT 约束)
IF EXISTS (
    SELECT 1 FROM sys.default_constraints dc
    JOIN sys.columns c ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
    WHERE dc.parent_object_id = OBJECT_ID(N'[dbo].[ProductReviews]') AND c.name = N'LikeCount'
)
BEGIN
    DECLARE @dcName NVARCHAR(256);
    SELECT @dcName = dc.name FROM sys.default_constraints dc
    JOIN sys.columns c ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
    WHERE dc.parent_object_id = OBJECT_ID(N'[dbo].[ProductReviews]') AND c.name = N'LikeCount';
    EXEC('ALTER TABLE [dbo].[ProductReviews] DROP CONSTRAINT [' + @dcName + ']');
    PRINT '[R1-03d] DEFAULT 约束 ' + @dcName + ' 已删除';
END
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ProductReviews]') AND name = N'LikeCount')
BEGIN
    ALTER TABLE [dbo].[ProductReviews] DROP COLUMN [LikeCount];
    PRINT '[R1-03d] ProductReviews.LikeCount 列已删除';
END
ELSE
    PRINT '[R1-03d] ProductReviews.LikeCount 列不存在，跳过';
GO

-- IsVerifiedPurchase (有 DEFAULT 约束)
IF EXISTS (
    SELECT 1 FROM sys.default_constraints dc
    JOIN sys.columns c ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
    WHERE dc.parent_object_id = OBJECT_ID(N'[dbo].[ProductReviews]') AND c.name = N'IsVerifiedPurchase'
)
BEGIN
    DECLARE @dcName2 NVARCHAR(256);
    SELECT @dcName2 = dc.name FROM sys.default_constraints dc
    JOIN sys.columns c ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
    WHERE dc.parent_object_id = OBJECT_ID(N'[dbo].[ProductReviews]') AND c.name = N'IsVerifiedPurchase';
    EXEC('ALTER TABLE [dbo].[ProductReviews] DROP CONSTRAINT [' + @dcName2 + ']');
    PRINT '[R1-03b] DEFAULT 约束 ' + @dcName2 + ' 已删除';
END
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ProductReviews]') AND name = N'IsVerifiedPurchase')
BEGIN
    ALTER TABLE [dbo].[ProductReviews] DROP COLUMN [IsVerifiedPurchase];
    PRINT '[R1-03b] ProductReviews.IsVerifiedPurchase 列已删除';
END
ELSE
    PRINT '[R1-03b] ProductReviews.IsVerifiedPurchase 列不存在，跳过';
GO

-- AIFeelingSummary (无 DEFAULT 约束)
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ProductReviews]') AND name = N'AIFeelingSummary')
BEGIN
    ALTER TABLE [dbo].[ProductReviews] DROP COLUMN [AIFeelingSummary];
    PRINT '[R1-03c] ProductReviews.AIFeelingSummary 列已删除';
END
ELSE
    PRINT '[R1-03c] ProductReviews.AIFeelingSummary 列不存在，跳过';
GO

-- Title (无 DEFAULT 约束)
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ProductReviews]') AND name = N'Title')
BEGIN
    ALTER TABLE [dbo].[ProductReviews] DROP COLUMN [Title];
    PRINT '[R1-03a] ProductReviews.Title 列已删除';
END
ELSE
    PRINT '[R1-03a] ProductReviews.Title 列不存在，跳过';
GO

-- =============================================================================
-- [R1-03e] 删除 ProductReviews 主键约束
-- 回滚 M1-03e: PK_ProductReviews
-- 注意: 必须在删除 ReviewImages 表之后执行（外键依赖已解除）
-- =============================================================================
IF EXISTS (
    SELECT 1 FROM sys.key_constraints
    WHERE parent_object_id = OBJECT_ID(N'[dbo].[ProductReviews]')
      AND name = N'PK_ProductReviews' AND type = 'PK'
)
BEGIN
    ALTER TABLE [dbo].[ProductReviews] DROP CONSTRAINT [PK_ProductReviews];
    PRINT '[R1-03e] PK_ProductReviews 主键约束已删除';
END
ELSE
    PRINT '[R1-03e] PK_ProductReviews 不存在，跳过';
GO

-- =============================================================================
-- [R1-02] 删除 Orders 积分相关字段
-- 回滚 M1-02: PointsEarned, PointsRedeemed, PointsDiscount
-- 注意: 带 DEFAULT 约束的列需先删除约束再删除列，数据将永久丢失。
-- =============================================================================

-- PointsDiscount (有 DEFAULT 约束)
IF EXISTS (
    SELECT 1 FROM sys.default_constraints dc
    JOIN sys.columns c ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
    WHERE dc.parent_object_id = OBJECT_ID(N'[dbo].[Orders]') AND c.name = N'PointsDiscount'
)
BEGIN
    DECLARE @dcName3 NVARCHAR(256);
    SELECT @dcName3 = dc.name FROM sys.default_constraints dc
    JOIN sys.columns c ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
    WHERE dc.parent_object_id = OBJECT_ID(N'[dbo].[Orders]') AND c.name = N'PointsDiscount';
    EXEC('ALTER TABLE [dbo].[Orders] DROP CONSTRAINT [' + @dcName3 + ']');
    PRINT '[R1-02c] DEFAULT 约束 ' + @dcName3 + ' 已删除';
END
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Orders]') AND name = N'PointsDiscount')
BEGIN
    ALTER TABLE [dbo].[Orders] DROP COLUMN [PointsDiscount];
    PRINT '[R1-02c] Orders.PointsDiscount 列已删除';
END
ELSE
    PRINT '[R1-02c] Orders.PointsDiscount 列不存在，跳过';
GO

-- PointsRedeemed (有 DEFAULT 约束)
IF EXISTS (
    SELECT 1 FROM sys.default_constraints dc
    JOIN sys.columns c ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
    WHERE dc.parent_object_id = OBJECT_ID(N'[dbo].[Orders]') AND c.name = N'PointsRedeemed'
)
BEGIN
    DECLARE @dcName4 NVARCHAR(256);
    SELECT @dcName4 = dc.name FROM sys.default_constraints dc
    JOIN sys.columns c ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
    WHERE dc.parent_object_id = OBJECT_ID(N'[dbo].[Orders]') AND c.name = N'PointsRedeemed';
    EXEC('ALTER TABLE [dbo].[Orders] DROP CONSTRAINT [' + @dcName4 + ']');
    PRINT '[R1-02b] DEFAULT 约束 ' + @dcName4 + ' 已删除';
END
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Orders]') AND name = N'PointsRedeemed')
BEGIN
    ALTER TABLE [dbo].[Orders] DROP COLUMN [PointsRedeemed];
    PRINT '[R1-02b] Orders.PointsRedeemed 列已删除';
END
ELSE
    PRINT '[R1-02b] Orders.PointsRedeemed 列不存在，跳过';
GO

-- PointsEarned (有 DEFAULT 约束)
IF EXISTS (
    SELECT 1 FROM sys.default_constraints dc
    JOIN sys.columns c ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
    WHERE dc.parent_object_id = OBJECT_ID(N'[dbo].[Orders]') AND c.name = N'PointsEarned'
)
BEGIN
    DECLARE @dcName5 NVARCHAR(256);
    SELECT @dcName5 = dc.name FROM sys.default_constraints dc
    JOIN sys.columns c ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
    WHERE dc.parent_object_id = OBJECT_ID(N'[dbo].[Orders]') AND c.name = N'PointsEarned';
    EXEC('ALTER TABLE [dbo].[Orders] DROP CONSTRAINT [' + @dcName5 + ']');
    PRINT '[R1-02a] DEFAULT 约束 ' + @dcName5 + ' 已删除';
END
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Orders]') AND name = N'PointsEarned')
BEGIN
    ALTER TABLE [dbo].[Orders] DROP COLUMN [PointsEarned];
    PRINT '[R1-02a] Orders.PointsEarned 列已删除';
END
ELSE
    PRINT '[R1-02a] Orders.PointsEarned 列不存在，跳过';
GO

-- =============================================================================
-- [R1-01] 删除 Orders 优惠券相关字段
-- 回滚 M1-01: CouponCode, CouponDiscount
-- 注意: 带 DEFAULT 约束的列需先删除约束再删除列，数据将永久丢失。
-- =============================================================================

-- CouponDiscount (有 DEFAULT 约束)
IF EXISTS (
    SELECT 1 FROM sys.default_constraints dc
    JOIN sys.columns c ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
    WHERE dc.parent_object_id = OBJECT_ID(N'[dbo].[Orders]') AND c.name = N'CouponDiscount'
)
BEGIN
    DECLARE @dcName6 NVARCHAR(256);
    SELECT @dcName6 = dc.name FROM sys.default_constraints dc
    JOIN sys.columns c ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
    WHERE dc.parent_object_id = OBJECT_ID(N'[dbo].[Orders]') AND c.name = N'CouponDiscount';
    EXEC('ALTER TABLE [dbo].[Orders] DROP CONSTRAINT [' + @dcName6 + ']');
    PRINT '[R1-01b] DEFAULT 约束 ' + @dcName6 + ' 已删除';
END
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Orders]') AND name = N'CouponDiscount')
BEGIN
    ALTER TABLE [dbo].[Orders] DROP COLUMN [CouponDiscount];
    PRINT '[R1-01b] Orders.CouponDiscount 列已删除';
END
ELSE
    PRINT '[R1-01b] Orders.CouponDiscount 列不存在，跳过';
GO

-- CouponCode (无 DEFAULT 约束)
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Orders]') AND name = N'CouponCode')
BEGIN
    ALTER TABLE [dbo].[Orders] DROP COLUMN [CouponCode];
    PRINT '[R1-01a] Orders.CouponCode 列已删除';
END
ELSE
    PRINT '[R1-01a] Orders.CouponCode 列不存在，跳过';
GO

-- =============================================================================
-- 回滚完成摘要
-- =============================================================================
PRINT '';
PRINT '============================================';
PRINT '  M1 里程碑回滚完成';
PRINT '============================================';
PRINT '  [R1-01] Orders: CouponCode, CouponDiscount 已移除';
PRINT '  [R1-02] Orders: PointsEarned, PointsRedeemed, PointsDiscount 已移除';
PRINT '  [R1-03] ProductReviews: Title, IsVerifiedPurchase, AIFeelingSummary, LikeCount 已移除';
PRINT '  [R1-03e] PK_ProductReviews 主键约束已移除';
PRINT '  [R1-04] ReviewImages 表已移除';
PRINT '  [R1-05] IX_ReviewImages_ReviewId 索引已移除';
PRINT '============================================';
GO
