-- ============================================
-- LoginAlerts & IPBlacklist 表创建脚本
-- 执行方式：在 SQL Server Management Studio 中打开此文件，连接到 PerfumeShop 数据库后执行
-- ============================================
USE [PerfumeShop]
GO

-- LoginAlerts 表
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'LoginAlerts')
BEGIN
    CREATE TABLE [LoginAlerts] (
        [AlertID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [AlertType] NVARCHAR(50) NULL,
        [AlertLevel] NVARCHAR(20) NULL DEFAULT 'info',
        [AlertMessage] NVARCHAR(500) NULL,
        [IPAddress] NVARCHAR(50) NULL,
        [AdminID] INT NULL,
        [IsRead] BIT NULL DEFAULT 0,
        [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()
    )
    PRINT 'LoginAlerts 表创建成功'
END
ELSE
BEGIN
    PRINT 'LoginAlerts 表已存在，跳过'
END
GO

-- LoginAlerts 索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_LoginAlerts_CreatedAt')
    CREATE NONCLUSTERED INDEX [IX_LoginAlerts_CreatedAt] ON [LoginAlerts]([CreatedAt]) INCLUDE ([AlertType],[AlertLevel],[IsRead])
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_LoginAlerts_IsRead')
    CREATE NONCLUSTERED INDEX [IX_LoginAlerts_IsRead] ON [LoginAlerts]([IsRead],[CreatedAt])
GO

-- IPBlacklist 表
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'IPBlacklist')
BEGIN
    CREATE TABLE [IPBlacklist] (
        [IPID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [IPAddress] NVARCHAR(50) NOT NULL,
        [Reason] NVARCHAR(255) NULL,
        [BlockedAt] DATETIME2(7) NULL DEFAULT GETDATE(),
        [BlockedBy] INT NULL,
        [IsActive] BIT NULL DEFAULT 1,
        [ExpiresAt] DATETIME2(7) NULL,
        [HitCount] INT NULL DEFAULT 0,
        [LastHitAt] DATETIME2(7) NULL
    )
    PRINT 'IPBlacklist 表创建成功'
END
ELSE
BEGIN
    PRINT 'IPBlacklist 表已存在，跳过'
END
GO

-- IPBlacklist 索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_IPBlacklist_IPAddress')
    CREATE NONCLUSTERED INDEX [IX_IPBlacklist_IPAddress] ON [IPBlacklist]([IPAddress],[IsActive])
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_IPBlacklist_BlockedAt')
    CREATE NONCLUSTERED INDEX [IX_IPBlacklist_BlockedAt] ON [IPBlacklist]([BlockedAt]) INCLUDE ([IsActive])
GO

PRINT '========================================='
PRINT '所有表及索引创建完成！'
PRINT '========================================='
