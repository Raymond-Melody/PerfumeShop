@echo off
chcp 65001 >nul
title PerfumeShop 数据库备份权限自动修复工具
cd /d "%~dp0"

echo ============================================
echo   PerfumeShop 数据库备份权限自动修复工具
echo   请以管理员身份运行（右键 -> 以管理员身份运行）
echo ============================================
echo.

:: 检查 sqlcmd 是否可用
where sqlcmd >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [失败] 未找到 sqlcmd 命令！
    echo 请安装 SQL Server Management Studio 或 SQL Server 命令行工具。
    echo.
    pause
    exit /b 1
)

set "INSTANCE=localhost\SQLEXPRESS"
echo [信息] SQL Server 实例: %INSTANCE%
echo.

:: 获取当前 Windows 身份
echo [信息] 当前 Windows 用户: %USERDOMAIN%\%USERNAME%
echo.

:: SQL 脚本路径
set "SQL_FILE=%~dp0usp_BackupDatabase.sql"

:: ============================================
:: 步骤 1：创建存储过程
:: ============================================
echo [步骤 1/2] 正在创建存储过程 usp_BackupDatabase...
echo        （该存储过程以 dbo 身份运行，拥有完整备份权限）
echo.
sqlcmd -S %INSTANCE% -E -C -i "%SQL_FILE%" 2>&1
if %ERRORLEVEL% EQU 0 (
    echo.
    echo [成功] ✓ 存储过程创建成功！
) else (
    echo.
    echo [失败] ✗ 存储过程创建失败
    echo   可能原因：当前用户没有 CREATE PROCEDURE 权限
    echo   解决方法：请以管理员身份运行此脚本（右键 -> 以管理员身份运行）
    echo.
    pause
    exit /b 1
)

:: ============================================
:: 步骤 2：验证存储过程
:: ============================================
echo.
echo [步骤 2/2] 正在验证存储过程...
echo.
sqlcmd -S %INSTANCE% -E -C -Q "USE [PerfumeShop]; IF EXISTS(SELECT * FROM sys.procedures WHERE name='usp_BackupDatabase') PRINT '验证成功：usp_BackupDatabase 已存在' ELSE PRINT '验证失败：usp_BackupDatabase 未创建'" 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [成功] ✓ 验证通过
) else (
    echo [失败] ✗ 验证失败
)

:: ============================================
echo.
echo ============================================
echo   操作完成！
echo.
echo   接下来请：
echo   1. 关闭此窗口
echo   2. 刷新备份中心页面（localhost/admin/system/backup_center.asp）
echo   3. 点击「立即备份」按钮
echo ============================================
echo.
pause
