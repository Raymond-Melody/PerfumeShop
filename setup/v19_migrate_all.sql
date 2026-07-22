-- =============================================================================
-- PerfumeShop V19 SQL 迁移脚本汇总
-- 文件: setup/v19_migrate_all.sql
-- 编码: UTF-8 (无 BOM，兼容 sqlcmd -i 执行)
-- =============================================================================
-- 用途: 合并 M1-M5 所有增量迁移脚本，按里程碑顺序排列
-- 执行: sqlcmd -S "localhost\YOURPERFUME" -d PerfumeShop -E -C -i "setup\v19_migrate_all.sql"
-- 设计: 全部幂等（IF NOT EXISTS / IF EXISTS 保护），可重复执行
-- =============================================================================

SET NOCOUNT ON;
PRINT '============================================';
PRINT '  PerfumeShop V19 全量 SQL 迁移';
PRINT '  开始时间: ' + CONVERT(NVARCHAR(20), GETDATE(), 120);
PRINT '============================================';
GO

-- *****************************************************************************
-- M1 里程碑: 订单/评价扩展字段 + ReviewImages 表
-- 参考: setup/v19_migrate_m1.sql
-- *****************************************************************************

PRINT '';
PRINT '--- [M1] 订单/评价扩展字段 + ReviewImages ---';

-- [M1-01] Orders: 优惠券字段
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Orders]') AND name = N'CouponCode')
BEGIN
    ALTER TABLE [dbo].[Orders] ADD [CouponCode] NVARCHAR(50) NULL;
    PRINT '[M1-01a] Orders.CouponCode 已添加';
END
ELSE
    PRINT '[M1-01a] Orders.CouponCode 已存在，跳过';
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Orders]') AND name = N'CouponDiscount')
BEGIN
    ALTER TABLE [dbo].[Orders] ADD [CouponDiscount] DECIMAL(10,2) NULL DEFAULT 0;
    PRINT '[M1-01b] Orders.CouponDiscount 已添加';
END
ELSE
    PRINT '[M1-01b] Orders.CouponDiscount 已存在，跳过';
GO

-- [M1-02] Orders: 积分字段
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Orders]') AND name = N'PointsEarned')
BEGIN
    ALTER TABLE [dbo].[Orders] ADD [PointsEarned] INT NULL DEFAULT 0;
    PRINT '[M1-02a] Orders.PointsEarned 已添加';
END
ELSE
    PRINT '[M1-02a] Orders.PointsEarned 已存在，跳过';
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Orders]') AND name = N'PointsRedeemed')
BEGIN
    ALTER TABLE [dbo].[Orders] ADD [PointsRedeemed] INT NULL DEFAULT 0;
    PRINT '[M1-02b] Orders.PointsRedeemed 已添加';
END
ELSE
    PRINT '[M1-02b] Orders.PointsRedeemed 已存在，跳过';
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Orders]') AND name = N'PointsDiscount')
BEGIN
    ALTER TABLE [dbo].[Orders] ADD [PointsDiscount] DECIMAL(10,2) NULL DEFAULT 0;
    PRINT '[M1-02c] Orders.PointsDiscount 已添加';
END
ELSE
    PRINT '[M1-02c] Orders.PointsDiscount 已存在，跳过';
GO

-- [M1-03] ProductReviews: 扩展字段
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ProductReviews]') AND name = N'Title')
BEGIN
    ALTER TABLE [dbo].[ProductReviews] ADD [Title] NVARCHAR(200) NULL;
    PRINT '[M1-03a] ProductReviews.Title 已添加';
END
ELSE
    PRINT '[M1-03a] ProductReviews.Title 已存在，跳过';
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ProductReviews]') AND name = N'IsVerifiedPurchase')
BEGIN
    ALTER TABLE [dbo].[ProductReviews] ADD [IsVerifiedPurchase] BIT NOT NULL DEFAULT 0;
    PRINT '[M1-03b] ProductReviews.IsVerifiedPurchase 已添加';
END
ELSE
    PRINT '[M1-03b] ProductReviews.IsVerifiedPurchase 已存在，跳过';
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ProductReviews]') AND name = N'AIFeelingSummary')
BEGIN
    ALTER TABLE [dbo].[ProductReviews] ADD [AIFeelingSummary] NVARCHAR(500) NULL;
    PRINT '[M1-03c] ProductReviews.AIFeelingSummary 已添加';
END
ELSE
    PRINT '[M1-03c] ProductReviews.AIFeelingSummary 已存在，跳过';
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ProductReviews]') AND name = N'LikeCount')
BEGIN
    ALTER TABLE [dbo].[ProductReviews] ADD [LikeCount] INT NOT NULL DEFAULT 0;
    PRINT '[M1-03d] ProductReviews.LikeCount 已添加';
END
ELSE
    PRINT '[M1-03d] ProductReviews.LikeCount 已存在，跳过';
GO

-- [M1-03e] ProductReviews 主键
IF NOT EXISTS (SELECT 1 FROM sys.key_constraints WHERE parent_object_id = OBJECT_ID(N'[dbo].[ProductReviews]') AND type = 'PK')
BEGIN
    ALTER TABLE [dbo].[ProductReviews]
        ADD CONSTRAINT [PK_ProductReviews] PRIMARY KEY CLUSTERED ([ReviewID] ASC);
    PRINT '[M1-03e] PK_ProductReviews 已创建';
END
ELSE
    PRINT '[M1-03e] PK_ProductReviews 已存在，跳过';
GO

-- [M1-04] ReviewImages 表
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
    PRINT '[M1-04] ReviewImages 表已存在，跳过';
GO

-- [M1-05] ReviewImages 索引
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_ReviewImages_ReviewId' AND object_id = OBJECT_ID(N'[dbo].[ReviewImages]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_ReviewImages_ReviewId]
        ON [dbo].[ReviewImages]([ReviewId] ASC);
    PRINT '[M1-05] IX_ReviewImages_ReviewId 索引已创建';
END
ELSE
    PRINT '[M1-05] IX_ReviewImages_ReviewId 索引已存在，跳过';
GO

-- *****************************************************************************
-- M2 里程碑: AuthTokens 表（ASP/ASP.NET Core 双系统 Session 互通）
-- 参考: setup/v19_migrate_m2_auth.sql
-- *****************************************************************************

PRINT '';
PRINT '--- [M2] AuthTokens 表 ---';

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'AuthTokens')
BEGIN
    CREATE TABLE [dbo].[AuthTokens] (
        [TokenID]    INT           IDENTITY(1,1) NOT NULL,
        [UserID]     INT           NOT NULL,
        [Token]      NVARCHAR(128) NOT NULL,
        [CreatedAt]  DATETIME2     NOT NULL CONSTRAINT DF_AuthTokens_CreatedAt DEFAULT (GETUTCDATE()),
        [ExpiresAt]  DATETIME2     NOT NULL,
        [Source]     NVARCHAR(20)  NOT NULL,
        [IsActive]   BIT           NOT NULL CONSTRAINT DF_AuthTokens_IsActive DEFAULT (1),
        [IpAddress]  NVARCHAR(50)  NULL,

        CONSTRAINT PK_AuthTokens PRIMARY KEY CLUSTERED ([TokenID] ASC),
        CONSTRAINT FK_AuthTokens_Users FOREIGN KEY ([UserID]) REFERENCES [dbo].[Users]([UserID])
    );

    CREATE NONCLUSTERED INDEX [IX_AuthTokens_Token_ExpiresAt]
        ON [dbo].[AuthTokens] ([Token] ASC, [ExpiresAt] ASC)
        INCLUDE ([IsActive], [UserID], [Source]);

    CREATE NONCLUSTERED INDEX [IX_AuthTokens_UserId]
        ON [dbo].[AuthTokens] ([UserID] ASC)
        INCLUDE ([IsActive], [ExpiresAt]);

    PRINT '[M2] AuthTokens 表已创建';
END
ELSE
    PRINT '[M2] AuthTokens 表已存在，跳过';
GO

-- *****************************************************************************
-- M3 里程碑: PasswordResetTokens 表（密码重置令牌独立化）
-- 参考: setup/v19_migrate_m3_auth.sql
-- *****************************************************************************

PRINT '';
PRINT '--- [M3] PasswordResetTokens 表 ---';

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'PasswordResetTokens')
BEGIN
    CREATE TABLE PasswordResetTokens (
        TokenId       INT IDENTITY(1,1) PRIMARY KEY,
        UserId        INT NOT NULL,
        Token         NVARCHAR(128) NOT NULL,
        TokenHash     NVARCHAR(256) NOT NULL,
        ExpiresAt     DATETIME2 NOT NULL,
        IsUsed        BIT NOT NULL DEFAULT 0,
        CreatedAt     DATETIME2 NOT NULL DEFAULT GETDATE(),
        UsedAt        DATETIME2 NULL,
        CONSTRAINT FK_PasswordResetTokens_Users FOREIGN KEY (UserId) REFERENCES Users(UserId)
    );

    CREATE INDEX IX_PasswordResetTokens_TokenHash ON PasswordResetTokens(TokenHash);
    CREATE INDEX IX_PasswordResetTokens_UserId    ON PasswordResetTokens(UserId);
    PRINT '[M3] PasswordResetTokens 表已创建';
END
ELSE
    PRINT '[M3] PasswordResetTokens 表已存在，跳过';
GO

-- *****************************************************************************
-- M4-M5 里程碑: 预留位置（如无增量 Schema 变更则跳过）
-- *****************************************************************************

PRINT '';
PRINT '--- [M4-M5] 预留（当前无增量 Schema 变更） ---';
PRINT '[M4-M5] 无操作';
GO

-- =============================================================================
-- 迁移完成汇总
-- =============================================================================

PRINT '';
PRINT '============================================';
PRINT '  PerfumeShop V19 全量 SQL 迁移完成';
PRINT '  结束时间: ' + CONVERT(NVARCHAR(20), GETDATE(), 120);
PRINT '============================================';
PRINT '  [M1] Orders 优惠券/积分字段';
PRINT '  [M1] ProductReviews 扩展字段 + PK';
PRINT '  [M1] ReviewImages 表 + 索引';
PRINT '  [M2] AuthTokens 表';
PRINT '  [M3] PasswordResetTokens 表';
PRINT '  [M4-M5] (预留)';
PRINT '============================================';
GO
