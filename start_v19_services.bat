@echo off
REM =====================================================
REM V19 ASP.NET Core 服务启动引导脚本
REM 委托给 start_v19_services.ps1 (读取 v19_service_choice.txt 配置)
REM 运行 .\select_v19_service.ps1 可更改开机启动的服务
REM =====================================================
set "ROOT=%~dp0"
echo [V19] 正在启动，请稍候...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%ROOT%start_v19_services.ps1"
echo [V19] 启动完成（查看日志: logs\ 目录）
@echo off
REM =====================================================
REM V19 ASP.NET Core 服务启动脚本
REM 自动启动 API (端口5000) 和 Admin (端口5207)
REM =====================================================
set "ROOT=%~dp0"

REM 检查端口5000是否已在监听
netstat -ano | findstr ":5000 " >nul 2>&1
if %errorlevel% equ 0 (
    echo [V19] API 已在端口 5000 运行
) else (
    echo [V19] 启动 API (端口 5000)...
    start "V19_API" cmd /c "cd /d "%ROOT%src\PerfumeShop.Api" && set ASPNETCORE_URLS=http://localhost:5000 && set ASPNETCORE_ENVIRONMENT=Development && dotnet run --no-launch-profile"
)

REM 检查端口5207是否已在监听
netstat -ano | findstr ":5207 " >nul 2>&1
if %errorlevel% equ 0 (
    echo [V19] Admin 已在端口 5207 运行
) else (
    echo [V19] 启动 Admin (端口 5207)...
    start "V19_Admin" cmd /c "cd /d "%ROOT%src\PerfumeShop.Admin" && set ASPNETCORE_URLS=http://localhost:5207 && set ASPNETCORE_ENVIRONMENT=Development && dotnet run --no-launch-profile"
)

echo [V19] 服务启动完成
