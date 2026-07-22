-- =============================================================================
-- PerfumeShop V19 - M1 里程碑迁移脚本
-- 文件: setup/v19_migrate_m1.sql
-- 编码: UTF-8 (无 BOM，兼容 sqlcmd -i 执行)
-- =============================================================================
-- 目标: 为 Orders 表增加优惠券/积分字段，为 ProductReviews 增加扩展字段，
--       创建 ReviewImages 表及相关索引
-- 依赖: Orders 表、ProductReviews 表必须已存在
-- 回滚: 执行 database/v19_migrations/rollback_m1.sql
-- 幂等: 所有语句使用 IF NOT EXISTS 包裹，可重复执行不报错
-- =============================================================================
-- 执行方式:
--   sqlcmd -S "localhost\YOURPERFUME" -d PerfumeShop -E -C -i "setup\v19_migrate_m1.sql"
-- =============================================================================

SET NOCOUNT ON;
GO

-- =============================================================================
-- [M1-01] Orders 表: 增加优惠券相关字段
-- 目的: 支持订单使用优惠券码及折扣金额记录
-- 依赖: Orders 表已存在
-- 回滚: ALTER TABLE Orders DROP COLUMN CouponCode, CouponDiscount
-- =============================================================================

-- CouponCode: 优惠券编码（关联优惠券系统的 CouponCode）
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[Orders]') AND name = N'CouponCode'
)
BEGIN
    ALTER TABLE [dbo].[Orders] ADD [CouponCode] NVARCHAR(50) NULL;
    PRINT '[M1-01a] Orders.CouponCode 列已添加';
END
ELSE
BEGIN
    PRINT '[M1-01a] Orders.CouponCode 列已存在，跳过';
END
GO

-- CouponDiscount: 优惠券折扣金额
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[Orders]') AND name = N'CouponDiscount'
)
BEGIN
    ALTER TABLE [dbo].[Orders] ADD [CouponDiscount] DECIMAL(10,2) NULL DEFAULT 0;
    PRINT '[M1-01b] Orders.CouponDiscount 列已添加';
END
ELSE
BEGIN
    PRINT '[M1-01b] Orders.CouponDiscount 列已存在，跳过';
END
GO

-- =============================================================================
-- [M1-02] Orders 表: 增加积分相关字段
-- 目的: 记录订单赚取/兑换的积分数量及积分折扣金额
-- 依赖: Orders 表已存在
-- 回滚: ALTER TABLE Orders DROP COLUMN PointsEarned, PointsRedeemed, PointsDiscount
-- =============================================================================

-- PointsEarned: 本单赚取的积分
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[Orders]') AND name = N'PointsEarned'
)
BEGIN
    ALTER TABLE [dbo].[Orders] ADD [PointsEarned] INT NULL DEFAULT 0;
    PRINT '[M1-02a] Orders.PointsEarned 列已添加';
END
ELSE
BEGIN
    PRINT '[M1-02a] Orders.PointsEarned 列已存在，跳过';
END
GO

-- PointsRedeemed: 本单兑换消耗的积分
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[Orders]') AND name = N'PointsRedeemed'
)
BEGIN
    ALTER TABLE [dbo].[Orders] ADD [PointsRedeemed] INT NULL DEFAULT 0;
    PRINT '[M1-02b] Orders.PointsRedeemed 列已添加';
END
ELSE
BEGIN
    PRINT '[M1-02b] Orders.PointsRedeemed 列已存在，跳过';
END
GO

-- PointsDiscount: 积分兑换产生的折扣金额
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[Orders]') AND name = N'PointsDiscount'
)
BEGIN
    ALTER TABLE [dbo].[Orders] ADD [PointsDiscount] DECIMAL(10,2) NULL DEFAULT 0;
    PRINT '[M1-02c] Orders.PointsDiscount 列已添加';
END
ELSE
BEGIN
    PRINT '[M1-02c] Orders.PointsDiscount 列已存在，跳过';
END
GO

-- =============================================================================
-- [M1-03] ProductReviews 表: 增加评论扩展字段
-- 目的: 支持评论标题、验证购买标记、AI 情感摘要、点赞数
-- 依赖: ProductReviews 表已存在
-- 回滚: ALTER TABLE ProductReviews DROP COLUMN Title, IsVerifiedPurchase,
--       AIFeelingSummary, LikeCount
-- =============================================================================

-- Title: 评论标题
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[ProductReviews]') AND name = N'Title'
)
BEGIN
    ALTER TABLE [dbo].[ProductReviews] ADD [Title] NVARCHAR(200) NULL;
    PRINT '[M1-03a] ProductReviews.Title 列已添加';
END
ELSE
BEGIN
    PRINT '[M1-03a] ProductReviews.Title 列已存在，跳过';
END
GO

-- IsVerifiedPurchase: 是否为验证购买用户的评论
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[ProductReviews]') AND name = N'IsVerifiedPurchase'
)
BEGIN
    ALTER TABLE [dbo].[ProductReviews] ADD [IsVerifiedPurchase] BIT NOT NULL DEFAULT 0;
    PRINT '[M1-03b] ProductReviews.IsVerifiedPurchase 列已添加';
END
ELSE
BEGIN
    PRINT '[M1-03b] ProductReviews.IsVerifiedPurchase 列已存在，跳过';
END
GO

-- AIFeelingSummary: AI 生成的情感分析摘要
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[ProductReviews]') AND name = N'AIFeelingSummary'
)
BEGIN
    ALTER TABLE [dbo].[ProductReviews] ADD [AIFeelingSummary] NVARCHAR(500) NULL;
    PRINT '[M1-03c] ProductReviews.AIFeelingSummary 列已添加';
END
ELSE
BEGIN
    PRINT '[M1-03c] ProductReviews.AIFeelingSummary 列已存在，跳过';
END
GO

-- LikeCount: 评论点赞数
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[ProductReviews]') AND name = N'LikeCount'
)
BEGIN
    ALTER TABLE [dbo].[ProductReviews] ADD [LikeCount] INT NOT NULL DEFAULT 0;
    PRINT '[M1-03d] ProductReviews.LikeCount 列已添加';
END
ELSE
BEGIN
    PRINT '[M1-03d] ProductReviews.LikeCount 列已存在，跳过';
END
GO

-- =============================================================================
-- [M1-03e] 确保 ProductReviews 表有主键（ReviewID）
-- 目的: ReviewImages 表外键需要引用 ProductReviews 的主键
-- 回滚: ALTER TABLE ProductReviews DROP CONSTRAINT PK_ProductReviews
-- =============================================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.key_constraints
    WHERE parent_object_id = OBJECT_ID(N'[dbo].[ProductReviews]') AND type = 'PK'
)
BEGIN
    ALTER TABLE [dbo].[ProductReviews]
        ADD CONSTRAINT [PK_ProductReviews] PRIMARY KEY CLUSTERED ([ReviewID] ASC);
    PRINT '[M1-03e] ProductReviews 主键 PK_ProductReviews 已创建';
END
ELSE
BEGIN
    PRINT '[M1-03e] ProductReviews 主键已存在，跳过';
END
GO

-- =============================================================================
-- [M1-04] 创建 ReviewImages 表
-- 目的: 存储评论关联的图片（支持每条评论多图上传）
-- 依赖: ProductReviews 表已存在且有主键 PK_ProductReviews（外键引用 ReviewID）
-- 回滚: DROP TABLE IF EXISTS ReviewImages
-- =============================================================================

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = N'ReviewImages' AND schema_id = SCHEMA_ID(N'dbo'))
BEGIN
    CREATE TABLE [dbo].[ReviewImages] (
        [Id]          INT IDENTITY(1,1) NOT NULL,
        [ReviewId]    INT NOT NULL,
        [ImageUrl]    NVARCHAR(500) NOT NULL,
        [SortOrder]   INT NOT NULL DEFAULT 0,
        [CreatedAt]   DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME(),
        CONSTRAINT [PK_ReviewImages] PRIMARY KEY CLUSTERED ([Id] ASC),
        CONSTRAINT [FK_ReviewImages_ProductReviews] FOREIGN KEY ([ReviewId])
            REFERENCES [dbo].[ProductReviews]([ReviewID]) ON DELETE CASCADE
    );
    PRINT '[M1-04] ReviewImages 表已创建';
END
ELSE
BEGIN
    PRINT '[M1-04] ReviewImages 表已存在，跳过';
END
GO

-- =============================================================================
-- [M1-05] 创建 ReviewImages.ReviewId 索引
-- 目的: 加速按评论 ID 查询图片的性能
-- 依赖: ReviewImages 表已存在
-- 回滚: DROP INDEX IF EXISTS IX_ReviewImages_ReviewId ON ReviewImages
-- =============================================================================

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_ReviewImages_ReviewId' AND object_id = OBJECT_ID(N'[dbo].[ReviewImages]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_ReviewImages_ReviewId]
        ON [dbo].[ReviewImages]([ReviewId] ASC);
    PRINT '[M1-05] IX_ReviewImages_ReviewId 索引已创建';
END
ELSE
BEGIN
    PRINT '[M1-05] IX_ReviewImages_ReviewId 索引已存在，跳过';
END
GO

-- =============================================================================
-- 迁移完成摘要
-- =============================================================================
PRINT '';
PRINT '============================================';
PRINT '  M1 里程碑迁移完成';
PRINT '============================================';
PRINT '  [M1-01] Orders: CouponCode, CouponDiscount';
PRINT '  [M1-02] Orders: PointsEarned, PointsRedeemed, PointsDiscount';
PRINT '  [M1-03] ProductReviews: Title, IsVerifiedPurchase, AIFeelingSummary, LikeCount';
PRINT '  [M1-03e] PK_ProductReviews 主键约束';
PRINT '  [M1-04] ReviewImages 表';
PRINT '  [M1-05] IX_ReviewImages_ReviewId 索引';
PRINT '============================================';
GO
