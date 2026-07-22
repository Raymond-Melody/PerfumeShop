-- ============================================================
-- V19 M2: AuthTokens 回滚脚本（幂等）
-- 删除 AuthTokens 表及相关索引
-- 执行: sqlcmd -S <server> -d <db> -i rollback_m2.sql
-- ============================================================

IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'AuthTokens')
BEGIN
    -- 先删除外键约束（如果存在）
    DECLARE @fkName NVARCHAR(128);
    SELECT @fkName = CONSTRAINT_NAME
    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS
    WHERE TABLE_NAME = 'AuthTokens' AND CONSTRAINT_TYPE = 'FOREIGN KEY';

    IF @fkName IS NOT NULL
    BEGIN
        EXEC('ALTER TABLE [dbo].[AuthTokens] DROP CONSTRAINT [' + @fkName + ']');
        PRINT 'Dropped foreign key: ' + @fkName;
    END

    DROP TABLE [dbo].[AuthTokens];
    PRINT 'AuthTokens table dropped successfully.';
END
ELSE
BEGIN
    PRINT 'AuthTokens table does not exist. Skipped.';
END
GO
