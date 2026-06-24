-- ============================================
-- V14 会员推荐制数据库升级脚本
-- 执行方式：在 SQL Server Management Studio 中打开此文件，连接到 PerfumeShop 数据库后执行
-- 或通过 db_setup.asp 工具自动执行
-- ============================================
USE [PerfumeShop]
GO

PRINT '===== V14 会员推荐制 Schema 升级开始 ====='
GO

-- 1. ReferralTokens 推荐链接Token表
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ReferralTokens')
BEGIN
    CREATE TABLE [ReferralTokens] (
        [TokenID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [TokenHash] NVARCHAR(64) NOT NULL,
        [OriginalToken] NVARCHAR(1000) NULL,
        [ReferrerUserID] INT NOT NULL,
        [ReferrerType] NVARCHAR(10) NOT NULL DEFAULT 'user',
        [ExpiresAt] DATETIME2(7) NOT NULL,
        [MaxUses] INT NOT NULL DEFAULT 1,
        [UsedCount] INT NOT NULL DEFAULT 0,
        [IsActive] BIT NOT NULL DEFAULT 1,
        [CreatedAt] DATETIME2(7) NOT NULL DEFAULT GETDATE()
    )
    PRINT 'ReferralTokens 表创建成功'
END
ELSE
BEGIN
    PRINT 'ReferralTokens 表已存在，跳过'
END
GO

-- ReferralTokens 索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_ReferralTokens_TokenHash')
    CREATE NONCLUSTERED INDEX [IX_ReferralTokens_TokenHash] ON [ReferralTokens]([TokenHash]) INCLUDE ([ReferrerUserID],[ExpiresAt],[MaxUses],[UsedCount],[IsActive])
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_ReferralTokens_ReferrerUserID')
    CREATE NONCLUSTERED INDEX [IX_ReferralTokens_ReferrerUserID] ON [ReferralTokens]([ReferrerUserID],[CreatedAt]) INCLUDE ([TokenHash],[ExpiresAt],[UsedCount],[IsActive])
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_ReferralTokens_ExpiresAt')
    CREATE NONCLUSTERED INDEX [IX_ReferralTokens_ExpiresAt] ON [ReferralTokens]([ExpiresAt]) INCLUDE ([IsActive])
GO

-- 2. ReferralRelations 推荐关系祖先链条表
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ReferralRelations')
BEGIN
    CREATE TABLE [ReferralRelations] (
        [RelationID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [AncestorUserID] INT NOT NULL,
        [DescendantUserID] INT NOT NULL,
        [Depth] INT NOT NULL,
        [CreatedAt] DATETIME2(7) NOT NULL DEFAULT GETDATE()
    )
    PRINT 'ReferralRelations 表创建成功'
END
ELSE
BEGIN
    PRINT 'ReferralRelations 表已存在，跳过'
END
GO

-- ReferralRelations 索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_ReferralRelations_Descendant')
    CREATE NONCLUSTERED INDEX [IX_ReferralRelations_Descendant] ON [ReferralRelations]([DescendantUserID]) INCLUDE ([AncestorUserID],[Depth])
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_ReferralRelations_Ancestor')
    CREATE NONCLUSTERED INDEX [IX_ReferralRelations_Ancestor] ON [ReferralRelations]([AncestorUserID]) INCLUDE ([DescendantUserID],[Depth])
GO

-- 3. RegistrationAttempts 注册尝试速率限制表
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'RegistrationAttempts')
BEGIN
    CREATE TABLE [RegistrationAttempts] (
        [AttemptID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [IPAddress] NVARCHAR(50) NOT NULL,
        [DeviceFingerprint] NVARCHAR(100) NULL,
        [Success] BIT NOT NULL DEFAULT 0,
        [TokenHash] NVARCHAR(64) NULL,
        [AttemptedAt] DATETIME2(7) NOT NULL DEFAULT GETDATE()
    )
    PRINT 'RegistrationAttempts 表创建成功'
END
ELSE
BEGIN
    PRINT 'RegistrationAttempts 表已存在，跳过'
END
GO

-- RegistrationAttempts 索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_RegistrationAttempts_IP')
    CREATE NONCLUSTERED INDEX [IX_RegistrationAttempts_IP] ON [RegistrationAttempts]([IPAddress],[AttemptedAt]) INCLUDE ([Success])
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_RegistrationAttempts_Fingerprint')
    CREATE NONCLUSTERED INDEX [IX_RegistrationAttempts_Fingerprint] ON [RegistrationAttempts]([DeviceFingerprint],[AttemptedAt]) INCLUDE ([Success])
GO

-- 4. Users 表添加 ReferrerUserID 字段
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('Users') AND name = 'ReferrerUserID')
BEGIN
    ALTER TABLE [Users] ADD [ReferrerUserID] INT NULL
    PRINT 'Users.ReferrerUserID 字段添加成功'
END
ELSE
BEGIN
    PRINT 'Users.ReferrerUserID 字段已存在，跳过'
END
GO

-- 5. Users 表添加 DeviceFingerprint 字段
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('Users') AND name = 'DeviceFingerprint')
BEGIN
    ALTER TABLE [Users] ADD [DeviceFingerprint] NVARCHAR(100) NULL
    PRINT 'Users.DeviceFingerprint 字段添加成功'
END
ELSE
BEGIN
    PRINT 'Users.DeviceFingerprint 字段已存在，跳过'
END
GO

-- 6. Users 表添加 UserRole 字段支持 member 角色（如果不存在）
-- UserRole 字段已存在于原表，无需额外添加，仅确保默认值正确

-- 7. ReferralRelations 唯一约束（防止重复写入同一对关系）
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'UQ_ReferralRelations_Pair')
    CREATE UNIQUE NONCLUSTERED INDEX [UQ_ReferralRelations_Pair] ON [ReferralRelations]([AncestorUserID],[DescendantUserID])
GO

PRINT '===== V14 会员推荐制 Schema 升级完成 ====='
GO

-- 8. V14.1: ReferralTokens 添加 OriginalToken 列（升级已有数据库）
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[ReferralTokens]') AND name = 'OriginalToken')
BEGIN
    ALTER TABLE [ReferralTokens] ADD [OriginalToken] NVARCHAR(1000) NULL
    PRINT 'ReferralTokens.OriginalToken 字段添加成功'
END
ELSE
BEGIN
    PRINT 'ReferralTokens.OriginalToken 字段已存在，跳过'
END
GO
