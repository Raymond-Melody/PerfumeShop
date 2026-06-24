-- V15.0 结构化日志表
-- 用于 logger.asp 的 SQL 日志存储（ERROR/FATAL 级别）
-- 执行方式: sqlcmd -S localhost\YOURPERFUME -d PerfumeShop -i v15_app_logs.sql

USE PerfumeShop;
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'AppLogs')
BEGIN
    CREATE TABLE [AppLogs] (
        [LogID] BIGINT IDENTITY(1,1) NOT NULL,
        [LogLevel] NVARCHAR(10) NOT NULL,
        [LogMessage] NVARCHAR(500) NULL,
        [LogSource] NVARCHAR(100) NULL,
        [LineNumber] INT NULL,
        [UserName] NVARCHAR(100) NULL,
        [IPAddress] NVARCHAR(50) NULL,
        [PageURL] NVARCHAR(200) NULL,
        [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(),
        CONSTRAINT [PK_AppLogs] PRIMARY KEY CLUSTERED ([LogID] ASC)
    );
    
    -- 按时间索引（日志查询最常用）
    CREATE NONCLUSTERED INDEX [IX_AppLogs_CreatedAt] 
        ON [AppLogs]([CreatedAt] DESC) 
        INCLUDE ([LogLevel], [LogMessage], [LogSource], [UserName]);
    
    -- 按级别索引（过滤ERROR/FATAL）
    CREATE NONCLUSTERED INDEX [IX_AppLogs_Level] 
        ON [AppLogs]([LogLevel], [CreatedAt] DESC) 
        INCLUDE ([LogMessage], [LogSource]);
    
    -- 自动清理30天前的日志
    PRINT 'AppLogs table created successfully.';
END
GO

-- 日志清理存储过程（可选：计划任务定期执行）
IF NOT EXISTS (SELECT * FROM sys.objects WHERE type = 'P' AND name = 'usp_CleanOldLogs')
BEGIN
    EXEC('CREATE PROCEDURE usp_CleanOldLogs
        @RetentionDays INT = 30
    AS
    BEGIN
        DELETE FROM AppLogs 
        WHERE CreatedAt < DATEADD(DAY, -@RetentionDays, GETDATE())
        AND LogLevel NOT IN (''FATAL'');
        
        SELECT @@ROWCOUNT AS DeletedRows;
    END');
    PRINT 'usp_CleanOldLogs created.';
END
GO
