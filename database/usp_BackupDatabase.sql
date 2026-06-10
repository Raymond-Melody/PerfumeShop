-- ============================================
-- usp_BackupDatabase - 以 dbo 权限执行备份
-- 使用 EXECUTE AS OWNER 绕过用户权限限制
-- ============================================
USE [PerfumeShop]
GO

IF EXISTS (SELECT * FROM sys.procedures WHERE name = 'usp_BackupDatabase')
    DROP PROCEDURE usp_BackupDatabase
GO

CREATE PROCEDURE usp_BackupDatabase
    @backupPath    NVARCHAR(500),
    @backupFile    NVARCHAR(200),
    @dbName        NVARCHAR(100) = 'PerfumeShop'
WITH EXECUTE AS OWNER  -- 以数据库所有者(dbo)身份运行，拥有完整备份权限
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @sql NVARCHAR(2000)
    DECLARE @fullPath NVARCHAR(700)
    
    -- 构造完整路径
    SET @fullPath = @backupPath + '\' + @backupFile
    
    -- 执行 BACKUP DATABASE (作为 dbo 拥有充分权限)
    SET @sql = 'BACKUP DATABASE [' + @dbName + '] TO DISK = N''' + @fullPath + ''' WITH INIT, NAME = N''自动备份-' + @backupFile + ''''
    
    EXEC sp_executesql @sql
    
    IF @@ERROR <> 0
    BEGIN
        RAISERROR('备份执行失败，错误代码: %d', 16, 1, @@ERROR)
        RETURN 1
    END
    
    RETURN 0
END
GO

-- 授予 public 角色执行权限（所有用户均可调用）
GRANT EXECUTE ON usp_BackupDatabase TO public
GO

-- 验证存储过程已创建
SELECT '存储过程 usp_BackupDatabase 创建成功' AS Status
GO
