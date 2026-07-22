-- ============================================================
-- V19 M2: AuthTokens 表建表脚本（幂等）
-- 支持 ASP 与 ASP.NET Core 双系统 Session 互通
-- 执行: sqlcmd -S <server> -d <db> -i v19_migrate_m2_auth.sql
-- ============================================================

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'AuthTokens')
BEGIN
    CREATE TABLE [dbo].[AuthTokens] (
        [TokenID]    INT           IDENTITY(1,1) NOT NULL,
        [UserID]     INT           NOT NULL,
        [Token]      NVARCHAR(128) NOT NULL,   -- SHA-256 哈希（64 位十六进制）
        [CreatedAt]  DATETIME2     NOT NULL CONSTRAINT DF_AuthTokens_CreatedAt DEFAULT (GETUTCDATE()),
        [ExpiresAt]  DATETIME2     NOT NULL,
        [Source]     NVARCHAR(20)  NOT NULL,   -- UserLogin / AdminLogin
        [IsActive]   BIT           NOT NULL CONSTRAINT DF_AuthTokens_IsActive DEFAULT (1),
        [IpAddress]  NVARCHAR(50)  NULL,

        CONSTRAINT PK_AuthTokens PRIMARY KEY CLUSTERED ([TokenID] ASC),
        CONSTRAINT FK_AuthTokens_Users FOREIGN KEY ([UserID]) REFERENCES [dbo].[Users]([UserID])
    );

    -- Token + ExpiresAt 复合索引（AuthBridgeMiddleware 查询优化）
    CREATE NONCLUSTERED INDEX [IX_AuthTokens_Token_ExpiresAt]
        ON [dbo].[AuthTokens] ([Token] ASC, [ExpiresAt] ASC)
        INCLUDE ([IsActive], [UserID], [Source]);

    -- UserId 索引（按用户查询/撤销 Token）
    CREATE NONCLUSTERED INDEX [IX_AuthTokens_UserId]
        ON [dbo].[AuthTokens] ([UserID] ASC)
        INCLUDE ([IsActive], [ExpiresAt]);

    PRINT 'AuthTokens table created successfully.';
END
ELSE
BEGIN
    PRINT 'AuthTokens table already exists. Skipped.';
END
GO
