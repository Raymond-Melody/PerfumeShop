# V19 Admin dotnet watch 启动器（开机自启）
# 以 dotnet watch 热重载模式启动管理后台
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$adminDir = Join-Path $root "src\PerfumeShop.Admin"
$logDir = Join-Path $root "logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$logFile = Join-Path $logDir "admin_watch_$(Get-Date -Format 'yyyyMMdd').log"

Set-Location $adminDir
$env:ASPNETCORE_URLS = "http://localhost:5207"
$env:ASPNETCORE_ENVIRONMENT = "Development"
# rude edit（方法签名/字段结构变更）时自动重启而非交互式等待
$env:DOTNET_WATCH_RESTART_ON_RUDE_EDIT = "1"
$ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
"[$ts] Starting V19 Admin with dotnet watch..." | Out-File -FilePath $logFile -Append -Encoding utf8

# dotnet watch 启动，输出到日志
dotnet watch run --no-launch-profile 2>&1 | Out-File -FilePath $logFile -Append -Encoding utf8
