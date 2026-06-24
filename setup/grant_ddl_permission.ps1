# ============================================
# PerfumeShop - 一键授予完整数据库权限脚本
# 使用方式：右键"以管理员身份运行" PowerShell 执行此脚本
# 功能：自动检测 IIS 应用池身份并授予完整权限
# ============================================

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "PerfumeShop 数据库权限自动配置工具" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 检测当前 Windows 用户
$currentWindowsUser = [Environment]::UserName
Write-Host "[*] 当前 Windows 用户: $currentWindowsUser" -ForegroundColor Yellow

# 定义需要配置的 IIS 应用池身份
$iisUsers = @(
    "NT AUTHORITY\IUSR",
    "IIS APPPOOL\DefaultAppPool",
    "NT AUTHORITY\NETWORK SERVICE"
)

Write-Host "[*] 将配置以下 IIS 身份:" -ForegroundColor Yellow
foreach ($user in $iisUsers) {
    Write-Host "    - $user" -ForegroundColor Gray
}
Write-Host ""

try {
    # 构建完整的 SQL 命令
    $sqlCommands = @"
USE [PerfumeShop]
GO

PRINT '========================================='
PRINT '开始授予权限...'
PRINT '========================================='
GO
"@

    foreach ($user in $iisUsers) {
        $escapedUser = $user -replace "'", "''"
        $sqlCommands += @"

-- 配置用户: $user
PRINT ''
PRINT '--- 处理: $user ---'

-- 创建数据库用户（如果不存在）
IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = '$escapedUser')
BEGIN
    CREATE USER [$escapedUser] FOR LOGIN [$escapedUser]
    PRINT '✓ 数据库用户创建成功'
END
ELSE
    PRINT 'ℹ 数据库用户已存在'
GO

-- 授予 db_ddladmin 角色
IF IS_ROLEMEMBER('db_ddladmin', '$escapedUser') = 0
BEGIN
    ALTER ROLE db_ddladmin ADD MEMBER [$escapedUser]
    PRINT '✓ db_ddladmin 授予成功（DDL 操作权限）'
END
ELSE
    PRINT 'ℹ db_ddladmin 已授予'
GO

-- 授予 db_datareader 角色
IF IS_ROLEMEMBER('db_datareader', '$escapedUser') = 0
BEGIN
    ALTER ROLE db_datareader ADD MEMBER [$escapedUser]
    PRINT '✓ db_datareader 授予成功（数据读取权限）'
END
ELSE
    PRINT 'ℹ db_datareader 已授予'
GO

-- 授予 db_datawriter 角色
IF IS_ROLEMEMBER('db_datawriter', '$escapedUser') = 0
BEGIN
    ALTER ROLE db_datawriter ADD MEMBER [$escapedUser]
    PRINT '✓ db_datawriter 授予成功（数据写入权限）'
END
ELSE
    PRINT 'ℹ db_datawriter 已授予'
GO

-- 授予 BACKUP DATABASE 权限
BEGIN TRY
    IF NOT EXISTS (
        SELECT * FROM sys.database_permissions p
        JOIN sys.database_principals dp ON p.grantee_principal_id = dp.principal_id
        WHERE dp.name = '$escapedUser' AND p.permission_name = 'BACKUP DATABASE'
    )
    BEGIN
        GRANT BACKUP DATABASE TO [$escapedUser]
        PRINT '✓ BACKUP DATABASE 权限授予成功'
    END
    ELSE
        PRINT 'ℹ BACKUP DATABASE 已授予'
END TRY
BEGIN CATCH
    PRINT '⚠ BACKUP DATABASE 授予失败（可选权限）'
END CATCH
GO

-- 授予 BACKUP LOG 权限
BEGIN TRY
    IF NOT EXISTS (
        SELECT * FROM sys.database_permissions p
        JOIN sys.database_principals dp ON p.grantee_principal_id = dp.principal_id
        WHERE dp.name = '$escapedUser' AND p.permission_name = 'BACKUP LOG'
    )
    BEGIN
        GRANT BACKUP LOG TO [$escapedUser]
        PRINT '✓ BACKUP LOG 权限授予成功'
    END
    ELSE
        PRINT 'ℹ BACKUP LOG 已授予'
END TRY
BEGIN CATCH
    PRINT '⚠ BACKUP LOG 授予失败（可选权限）'
END CATCH
GO

PRINT '✅ $user 权限配置完成'
GO
"@
    }

    $sqlCommands += @"

PRINT ''
PRINT '========================================='
PRINT '所有权限配置完成！'
PRINT '========================================='
GO
"@

    # 使用 sqlcmd 执行（Windows 集成认证）
    $sqlCmdPath = "sqlcmd"
    $server = "localhost\YOURPERFUME"
    
    Write-Host "[*] 正在连接 SQL Server: $server" -ForegroundColor Yellow
    Write-Host ""
    
    $output = & $sqlCmdPath -S $server -E -Q $sqlCommands 2>&1
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "" -ForegroundColor Green
        Write-Host "✅ 权限授予成功！" -ForegroundColor Green
        Write-Host ""
        Write-Host "详细输出:" -ForegroundColor Cyan
        Write-Host "========================================"
        $output | ForEach-Object { Write-Host "  $_" }
        Write-Host "========================================"
        Write-Host ""
        Write-Host "已授予的权限:" -ForegroundColor Green
        Write-Host "  • db_ddladmin    - DDL 操作（CREATE/ALTER/DROP 表、索引等）" -ForegroundColor White
        Write-Host "  • db_datareader  - 数据读取（SELECT 所有用户表）" -ForegroundColor White
        Write-Host "  • db_datawriter  - 数据写入（INSERT/UPDATE/DELETE）" -ForegroundColor White
        Write-Host "  • BACKUP DATABASE - 数据库完整备份" -ForegroundColor White
        Write-Host "  • BACKUP LOG     - 事务日志备份" -ForegroundColor White
        Write-Host ""
        Write-Host "下一步操作:" -ForegroundColor Yellow
        Write-Host "  1. 重启 IIS 应用程序池: iisreset" -ForegroundColor White
        Write-Host "  2. 访问部署工具: http://localhost/setup/deploy.asp?action=run" -ForegroundColor White
        Write-Host "  3. 运行权限验证: http://localhost/setup/verify_permissions.asp" -ForegroundColor White
    } else {
        Write-Host ""
        Write-Host "❌ 执行失败" -ForegroundColor Red
        Write-Host ""
        Write-Host "错误输出:" -ForegroundColor Red
        $output | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
        Write-Host ""
        Write-Host "故障排除步骤:" -ForegroundColor Yellow
        Write-Host "  1. 确认 SQL Server 服务正在运行" -ForegroundColor White
        Write-Host "  2. 确认当前 Windows 用户有 SQL Server sysadmin 权限" -ForegroundColor White
        Write-Host "  3. 检查数据库是否存在: http://localhost/setup/deploy.asp" -ForegroundColor White
        Write-Host "  4. 或在 SSMS 中手动执行 setup/grant_full_permissions.sql" -ForegroundColor White
    }
} catch {
    Write-Host ""
    Write-Host "❌ 发生异常: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "请手动执行以下步骤：" -ForegroundColor Yellow
    Write-Host "  1. 打开 SQL Server Management Studio" -ForegroundColor White
    Write-Host "  2. 连接到 localhost\YOURPERFUME" -ForegroundColor White
    Write-Host "  3. 打开 setup/grant_full_permissions.sql 并执行" -ForegroundColor White
}

Write-Host ""
Write-Host "按任意键退出..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
