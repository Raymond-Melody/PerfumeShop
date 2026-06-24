-- ============================================
-- PerfumeShop 完整权限管理脚本
-- ============================================
-- 功能：为 IIS 应用程序池身份授予完整的数据库权限
-- 使用方式：在 SQL Server Management Studio (SSMS) 中执行
-- 执行环境：需要 sysadmin 或 securityadmin 角色
-- ============================================

USE [master]
GO

PRINT '========================================='
PRINT 'PerfumeShop 权限配置脚本'
PRINT '开始时间: ' + CONVERT(NVARCHAR, GETDATE(), 120)
PRINT '========================================='
GO

-- ============================================
-- 第1步：识别当前 Windows 用户
-- ============================================
PRINT ''
PRINT '[步骤 1] 识别当前登录用户...'
DECLARE @CurrentUser NVARCHAR(255)
SELECT @CurrentUser = SUSER_SNAME()
PRINT '当前登录用户: ' + @CurrentUser
GO

-- ============================================
-- 第2步：确认 PerfumeShop 数据库存在
-- ============================================
PRINT ''
PRINT '[步骤 2] 检查 PerfumeShop 数据库...'
IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = 'PerfumeShop')
BEGIN
    PRINT '❌ 错误：PerfumeShop 数据库不存在！'
    PRINT '请先运行部署工具创建数据库：http://localhost/setup/deploy.asp'
    RETURN
END
PRINT '✓ PerfumeShop 数据库存在'
GO

-- ============================================
-- 第3步：定义需要授权的用户列表
-- ============================================
USE [PerfumeShop]
GO

PRINT ''
PRINT '[步骤 3] 准备授权用户列表...'

-- 常见的 IIS 应用池身份
DECLARE @Users TABLE (
    LoginName NVARCHAR(255),
    Description NVARCHAR(100)
)

INSERT INTO @Users VALUES 
    ('NT AUTHORITY\IUSR', 'IIS 默认匿名用户'),
    ('IIS APPPOOL\DefaultAppPool', 'IIS 默认应用池'),
    ('NT AUTHORITY\NETWORK SERVICE', 'Network Service 账户')

-- 显示将要处理的用户
SELECT 
    LoginName AS '用户登录名',
    Description AS '说明'
FROM @Users

PRINT ''
PRINT '共 ' + CAST((SELECT COUNT(*) FROM @Users) AS NVARCHAR) + ' 个用户需要配置'
GO

-- ============================================
-- 第4步：为每个用户创建数据库用户并授予角色
-- ============================================
PRINT ''
PRINT '[步骤 4] 开始授予权限...'

DECLARE @LoginName NVARCHAR(255)
DECLARE @SQL NVARCHAR(MAX)
DECLARE @UserCount INT = 0
DECLARE @SuccessCount INT = 0
DECLARE @ErrorCount INT = 0

DECLARE user_cursor CURSOR FOR 
SELECT LoginName FROM @Users

OPEN user_cursor
FETCH NEXT FROM user_cursor INTO @LoginName

WHILE @@FETCH_STATUS = 0
BEGIN
    SET @UserCount = @UserCount + 1
    PRINT ''
    PRINT '--- 处理用户: ' + @LoginName + ' ---'
    
    BEGIN TRY
        -- 4.1 创建登录名（如果不存在）
        -- 注意：内置账户通常已存在服务器级登录名，只需创建数据库用户
        
        -- 4.2 创建数据库用户（如果不存在）
        SET @SQL = '
        IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = ''' + @LoginName + ''')
        BEGIN
            CREATE USER [' + @LoginName + '] FOR LOGIN [' + @LoginName + ']
            PRINT ''  ✓ 数据库用户创建成功''
        END
        ELSE
        BEGIN
            PRINT ''  ℹ 数据库用户已存在''
        END'
        EXEC sp_executesql @SQL
        
        -- 4.3 授予 db_ddladmin 角色（DDL 操作权限）
        SET @SQL = '
        IF IS_ROLEMEMBER(''db_ddladmin'', ''' + @LoginName + ''') = 0
        BEGIN
            ALTER ROLE db_ddladmin ADD MEMBER [' + @LoginName + ']
            PRINT ''  ✓ db_ddladmin 角色授予成功（CREATE/ALTER/DROP 表、索引等）''
        END
        ELSE
        BEGIN
            PRINT ''  ℹ db_ddladmin 角色已授予''
        END'
        EXEC sp_executesql @SQL
        
        -- 4.4 授予 db_datareader 角色（读取数据权限）
        SET @SQL = '
        IF IS_ROLEMEMBER(''db_datareader'', ''' + @LoginName + ''') = 0
        BEGIN
            ALTER ROLE db_datareader ADD MEMBER [' + @LoginName + ']
            PRINT ''  ✓ db_datareader 角色授予成功（SELECT 所有用户表）''
        END
        ELSE
        BEGIN
            PRINT ''  ℹ db_datareader 角色已授予''
        END'
        EXEC sp_executesql @SQL
        
        -- 4.5 授予 db_datawriter 角色（写入数据权限）
        SET @SQL = '
        IF IS_ROLEMEMBER(''db_datawriter'', ''' + @LoginName + ''') = 0
        BEGIN
            ALTER ROLE db_datawriter ADD MEMBER [' + @LoginName + ']
            PRINT ''  ✓ db_datawriter 角色授予成功（INSERT/UPDATE/DELETE 所有用户表）''
        END
        ELSE
        BEGIN
            PRINT ''  ℹ db_datawriter 角色已授予''
        END'
        EXEC sp_executesql @SQL
        
        -- 4.6 授予 BACKUP DATABASE 权限
        SET @SQL = '
        BEGIN TRY
            -- 检查是否已有权限
            IF NOT EXISTS (
                SELECT * FROM sys.database_permissions p
                JOIN sys.database_principals dp ON p.grantee_principal_id = dp.principal_id
                WHERE dp.name = ''' + @LoginName + ''' 
                AND p.permission_name = ''BACKUP DATABASE''
            )
            BEGIN
                GRANT BACKUP DATABASE TO [' + @LoginName + ']
                PRINT ''  ✓ BACKUP DATABASE 权限授予成功''
            END
            ELSE
            BEGIN
                PRINT ''  ℹ BACKUP DATABASE 权限已授予''
            END
        END TRY
        BEGIN CATCH
            PRINT ''  ⚠ BACKUP DATABASE 权限授予失败: '' + ERROR_MESSAGE()
            PRINT ''    （备份权限可选，不影响核心功能）''
        END CATCH'
        EXEC sp_executesql @SQL
        
        -- 4.7 授予 BACKUP LOG 权限
        SET @SQL = '
        BEGIN TRY
            IF NOT EXISTS (
                SELECT * FROM sys.database_permissions p
                JOIN sys.database_principals dp ON p.grantee_principal_id = dp.principal_id
                WHERE dp.name = ''' + @LoginName + ''' 
                AND p.permission_name = ''BACKUP LOG''
            )
            BEGIN
                GRANT BACKUP LOG TO [' + @LoginName + ']
                PRINT ''  ✓ BACKUP LOG 权限授予成功''
            END
            ELSE
            BEGIN
                PRINT ''  ℹ BACKUP LOG 权限已授予''
            END
        END TRY
        BEGIN CATCH
            PRINT ''  ⚠ BACKUP LOG 权限授予失败: '' + ERROR_MESSAGE()
            PRINT ''    （备份权限可选，不影响核心功能）''
        END CATCH'
        EXEC sp_executesql @SQL
        
        SET @SuccessCount = @SuccessCount + 1
        PRINT '  ✅ 用户 ' + @LoginName + ' 权限配置完成'
        
    END TRY
    BEGIN CATCH
        SET @ErrorCount = @ErrorCount + 1
        PRINT '  ❌ 用户 ' + @LoginName + ' 权限配置失败: ' + ERROR_MESSAGE()
    END CATCH
    
    FETCH NEXT FROM user_cursor INTO @LoginName
END

CLOSE user_cursor
DEALLOCATE user_cursor

-- ============================================
-- 第5步：权限验证
-- ============================================
PRINT ''
PRINT '========================================='
PRINT '[步骤 5] 权限验证报告'
PRINT '========================================='
PRINT ''
PRINT '处理用户总数: ' + CAST(@UserCount AS NVARCHAR)
PRINT '成功: ' + CAST(@SuccessCount AS NVARCHAR)
PRINT '失败: ' + CAST(@ErrorCount AS NVARCHAR)
PRINT ''

-- 显示每个用户的权限详情
PRINT '--- 权限详情 ---'
SELECT 
    dp.name AS '用户',
    dp.type_desc AS '用户类型',
    STRING_AGG(
        CASE 
            WHEN dp2.name = 'db_ddladmin' THEN 'DDL管理'
            WHEN dp2.name = 'db_datareader' THEN '数据读取'
            WHEN dp2.name = 'db_datawriter' THEN '数据写入'
            ELSE dp2.name
        END, 
        ', '
    ) AS '数据库角色'
FROM sys.database_principals dp
LEFT JOIN sys.database_role_members drm ON dp.principal_id = drm.member_principal_id
LEFT JOIN sys.database_principals dp2 ON drm.role_principal_id = dp2.principal_id
WHERE dp.name IN ('NT AUTHORITY\IUSR', 'IIS APPPOOL\DefaultAppPool', 'NT AUTHORITY\NETWORK SERVICE')
GROUP BY dp.name, dp.type_desc

-- 显示 BACKUP 权限
PRINT ''
PRINT '--- BACKUP 权限详情 ---'
SELECT 
    dp.name AS '用户',
    p.permission_name AS '权限名称',
    p.state_desc AS '权限状态'
FROM sys.database_permissions p
JOIN sys.database_principals dp ON p.grantee_principal_id = dp.principal_id
WHERE p.permission_name IN ('BACKUP DATABASE', 'BACKUP LOG')
AND dp.name IN ('NT AUTHORITY\IUSR', 'IIS APPPOOL\DefaultAppPool', 'NT AUTHORITY\NETWORK SERVICE')

-- ============================================
-- 第6步：完成总结
-- ============================================
PRINT ''
PRINT '========================================='
PRINT '权限配置完成！'
PRINT '完成时间: ' + CONVERT(NVARCHAR, GETDATE(), 120)
PRINT '========================================='
PRINT ''
PRINT '已授予的权限说明：'
PRINT '  • db_ddladmin    - 创建/修改/删除表、索引、视图等数据库对象'
PRINT '  • db_datareader  - 读取所有用户表中的数据'
PRINT '  • db_datawriter  - 插入/更新/删除所有用户表中的数据'
PRINT '  • BACKUP DATABASE - 执行数据库完整备份'
PRINT '  • BACKUP LOG     - 执行事务日志备份'
PRINT ''
PRINT '下一步：'
PRINT '  1. 刷新 IIS 应用程序池：iisreset'
PRINT '  2. 访问部署工具验证：http://localhost/setup/deploy.asp'
PRINT '  3. 运行权限验证：http://localhost/setup/verify_permissions.asp'
PRINT ''

GO
