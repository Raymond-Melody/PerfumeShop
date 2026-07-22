@echo off
REM ============================================================
REM V19 IIS 子应用程序一键配置脚本
REM 请以管理员身份运行此脚本
REM ============================================================
cd /d "%~dp0"

set SITE_NAME="Default Web Site"
set API_DIR="%~dp0v19-api"
set ADMIN_DIR="%~dp0v19-admin"

echo ============================================
echo  V19 IIS 子应用程序配置
echo ============================================
echo.
echo 此脚本将:
echo  1. 检查并安装 ASP.NET Core Hosting Bundle
echo  2. 创建 IIS 子应用程序
echo  3. 重启 IIS
echo.

REM --- 步骤 1: 安装 Hosting Bundle ---
if exist "%TEMP%\dotnet-hosting-8.0.28-win.exe" (
    echo [1/3] 安装 ASP.NET Core Hosting Bundle...
    start /wait "" "%TEMP%\dotnet-hosting-8.0.28-win.exe" /quiet /norestart OPT_NO_RUNTIME=1
    echo   Hosting Bundle 安装完成
) else (
    echo [1/3] Hosting Bundle 安装包不存在，跳过
)
echo.

REM --- 步骤 2: 创建 IIS 子应用程序 ---
echo [2/3] 创建 IIS 子应用程序...

REM 删除旧的子应用程序（如果存在）
%windir%\system32\inetsrv\appcmd.exe delete app /app.name:Default Web Site\api 2>nul
%windir%\system32\inetsrv\appcmd.exe delete app /app.name:Default Web Site\admin\v2 2>nul

REM 创建 API 子应用程序 (/api)
%windir%\system32\inetsrv\appcmd.exe add app /site.name:%SITE_NAME% /path:/api /physicalPath:%API_DIR% /applicationPool:DefaultAppPool
if %errorlevel% equ 0 (
    echo   [OK] /api -^> %API_DIR%
) else (
    echo   [失败] /api 创建失败，尝试备用方法...
    REM 备用：设置虚拟目录
    %windir%\system32\inetsrv\appcmd.exe add vdir /app.name:Default Web Site\ /path:/api /physicalPath:%API_DIR%
)

REM 创建 Admin 子应用程序 (/admin/v2)
%windir%\system32\inetsrv\appcmd.exe add app /site.name:%SITE_NAME% /path:/admin/v2 /physicalPath:%ADMIN_DIR% /applicationPool:DefaultAppPool
if %errorlevel% equ 0 (
    echo   [OK] /admin/v2 -^> %ADMIN_DIR%
) else (
    echo   [失败] /admin/v2 创建失败，尝试备用方法...
    %windir%\system32\inetsrv\appcmd.exe add vdir /app.name:Default Web Site\admin /path:/v2 /physicalPath:%ADMIN_DIR%
)

echo.

REM --- 步骤 3: 重启 IIS ---
echo [3/3] 重启 IIS...
iisreset
echo.

echo ============================================
echo  配置完成！
echo ============================================
echo  访问测试:
echo    V19 API:    http://localhost/api/health
echo    V19 Admin:  http://localhost/admin/v2/
echo    V18 Classic: http://localhost/index.asp
echo ============================================
pause
