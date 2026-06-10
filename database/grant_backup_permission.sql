-- ============================================
-- 授予 BACKUP DATABASE 权限给 ASP 应用程序用户
-- ============================================
-- 使用方式: 在 SQL Server Management Studio (SSMS) 中打开并执行
-- 
-- 第一步：先在 ASP 页面运行时执行以下查询
-- 能反应 ASP 页面实际的连接用户身份
-- ============================================

-- ============================================
-- 使用方法（两步）：
-- 
-- 第1步：在备份中心页面刷新时，任意页面底部或单独
--       新建一个 ASP 文件运行：
--       <%= "当前登录: " & SUSER_SNAME() %>
--       或在 SSMS 中先执行下面的 SELECT
--
-- 第2步：根据查到的登录名，执行对应的 GRANT
-- ============================================

-- === 第1步：查询当前登录账号 ===
SELECT SUSER_SNAME() AS CurrentLogin

-- === 第2步：根据结果授权 ===
-- 常见结果和对应的授权语句：

-- 情况 A：NT AUTHORITY\NETWORK SERVICE
USE [PerfumeShop]
GO
GRANT BACKUP DATABASE TO [NT AUTHORITY\NETWORK SERVICE]
GO

-- 情况 B：IIS APPPOOL\DefaultAppPool
USE [PerfumeShop]
GO
GRANT BACKUP DATABASE TO [IIS APPPOOL\DefaultAppPool]
GO

-- 情况 C：NT AUTHORITY\SYSTEM
USE [PerfumeShop]
GO
GRANT BACKUP DATABASE TO [NT AUTHORITY\SYSTEM]
GO

-- === 第3步：验证权限 ===
USE [PerfumeShop]
GO
-- 查看当前登录是否已有 backup 权限
SELECT DP.name AS UserName, 
       DP.type_desc AS UserType,
       PERMISSION_NAME = 'BACKUP DATABASE',
       STATE_DESC = '存在'
FROM sys.database_permissions P
JOIN sys.database_principals DP ON P.grantee_principal_id = DP.principal_id
WHERE P.class = 0 AND P.permission_name = 'BACKUP DATABASE'
GO

-- ============================================
-- 备选方案：使用 db_backupoperator 角色
-- 如果 GRANT BACKUP DATABASE 不可用，用此方案
--
-- 注意：[LOGIN_NAME] 替换为第1步查到的值
-- ============================================
USE [PerfumeShop]
GO
EXEC sp_addrolemember 'db_backupoperator', [LOGIN_NAME]
GO

PRINT '完成！请刷新备份中心页面重试。'
