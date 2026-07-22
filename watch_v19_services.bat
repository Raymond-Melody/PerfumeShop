@echo off
REM =====================================================
REM V19 dotnet watch 热重载启动脚本
REM 自动检测源码变化：.cs / .razor / .cshtml 热重载，
REM Program.cs 等"重大编辑"自动重启进程，
REM wwwroot 静态文件变化自动刷新浏览器。
REM   API   -> http://localhost:5000  (Razor Pages + Swagger)
REM   Admin -> http://localhost:5207  (Blazor Server)
REM 每个服务在独立窗口运行，关闭窗口即可停止。
REM =====================================================
set "ROOT=%~dp0"

REM --- V19 API (端口 5000) 热重载 ---
netstat -ano | findstr ":5000 " >nul 2>&1
if %errorlevel% equ 0 (
    echo [V19] API 已在端口 5000 运行，跳过
) else (
    echo [V19] 启动 API 热重载 (端口 5000)...
    start "V19_API (watch)" cmd /c "cd /d "%ROOT%src\PerfumeShop.Api" && set ASPNETCORE_URLS=http://localhost:5000 && set ASPNETCORE_ENVIRONMENT=Development && dotnet watch --non-interactive --no-launch-profile"
)

REM --- V19 Admin (端口 5207) 热重载 ---
netstat -ano | findstr ":5207 " >nul 2>&1
if %errorlevel% equ 0 (
    echo [V19] Admin 已在端口 5207 运行，跳过
) else (
    echo [V19] 启动 Admin 热重载 (端口 5207)...
    start "V19_Admin (watch)" cmd /c "cd /d "%ROOT%src\PerfumeShop.Admin" && set ASPNETCORE_URLS=http://localhost:5207 && set ASPNETCORE_ENVIRONMENT=Development && dotnet watch --non-interactive --no-launch-profile"
)

echo [V19] dotnet watch 已启动（独立窗口）。修改源码将自动热重载/重启。
