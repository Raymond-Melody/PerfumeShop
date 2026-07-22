# V19 ASP.NET Core 服务启动脚本
# 支持选择启动 API (端口5000) 和/或 Admin (端口5207)
# 用法: .\start_v19_services.ps1 -Service All|Api|Admin
#       不加参数时：交互模式弹菜单，隐藏/非交互模式读配置文件

param(
    [ValidateSet("All","Api","Admin","")]
    [string]$Service = ""
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$logDir = Join-Path $scriptDir "logs"
$configFile = Join-Path $scriptDir "v19_service_choice.txt"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

# ---- 配置文件读写 ----
function Read-Choice {
    if (Test-Path $configFile) {
        $c = (Get-Content $configFile -Raw).Trim()
        if ($c -in "All","Api","Admin") { return $c }
    }
    return ""
}
function Save-Choice($c) { $c | Out-File -FilePath $configFile -Encoding utf8 }

# ---- 确定要启动的服务 ----
if ([string]::IsNullOrEmpty($Service)) {
    # 优先读配置文件（开机自启场景）
    $Service = Read-Choice
    if ([string]::IsNullOrEmpty($Service)) {
        # 无配置文件时判断是否可交互
        $canPrompt = $true
        try { $null = [Console]::KeyAvailable } catch { $canPrompt = $false }
        if ($canPrompt -and [Environment]::UserInteractive) {
            # 交互模式：显示菜单
            Write-Host ""
            Write-Host "=======================================" -ForegroundColor Cyan
            Write-Host "   V19 服务启动选择器" -ForegroundColor Cyan
            Write-Host "=======================================" -ForegroundColor Cyan
            Write-Host "  1) 启动全部 (API + Admin)" -ForegroundColor White
            Write-Host "  2) 仅启动 API (端口 5000)" -ForegroundColor White
            Write-Host "  3) 仅启动 Admin (端口 5207)" -ForegroundColor White
            Write-Host "  0) 退出" -ForegroundColor White
            Write-Host "=======================================" -ForegroundColor Cyan
            $key = Read-Host "请选择 (1/2/3/0)"
            switch ($key) {
                "1" { $Service = "All" }
                "2" { $Service = "Api" }
                "3" { $Service = "Admin" }
                "0" { Write-Host "已退出"; exit }
                default { Write-Host "无效选择，退出"; exit }
            }
            Save-Choice $Service
            Write-Host "选择已保存到: $configFile" -ForegroundColor Green
            Write-Host "下次开机自启将使用此选择。重新选择请运行: .\select_v19_service.ps1" -ForegroundColor Green
        }
    }
    # 仍然为空则默认全部
    if ([string]::IsNullOrEmpty($Service)) { $Service = "All" }
}

Write-Host "[V19] 启动模式: $Service" -ForegroundColor Cyan

# ---- 通用启动函数 ----
function Start-ServiceJob {
    param([string]$Name, [string]$Port, [string]$Dir, [string]$Label)
    $running = $false
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $tcp.Connect("127.0.0.1", $Port)
        $tcp.Close()
        $running = $true
    } catch {}
    if ($running) {
        Write-Host "[V19] $Label 已在端口 $Port 运行，跳过" -ForegroundColor Yellow
    } else {
        $log = Join-Path $logDir "$($Name.ToLower())_stdout.log"
        Write-Host "[V19] 启动 $Label (端口 $Port)..." -ForegroundColor Green
        Start-Job -Name "V19_$Name" -ScriptBlock {
            param($d, $l, $p)
            Set-Location $d
            $env:ASPNETCORE_URLS = "http://localhost:$p"
            $env:ASPNETCORE_ENVIRONMENT = "Development"
            dotnet run --no-launch-profile 2>&1 | Out-File -FilePath $l -Append -Encoding utf8
        } -ArgumentList $Dir, $log, $Port | Out-Null
    }
}

if ($Service -in "All","Api")   { Start-ServiceJob -Name "Api" -Port 5000 -Dir (Join-Path $scriptDir "src\PerfumeShop.Api") -Label "V19 API" }
if ($Service -in "All","Admin") { Start-ServiceJob -Name "Admin" -Port 5207 -Dir (Join-Path $scriptDir "src\PerfumeShop.Admin") -Label "V19 Admin" }

Write-Host "[V19] 服务启动完成" -ForegroundColor Green
