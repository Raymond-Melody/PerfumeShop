# =============================================================================
# PerfumeShop V19 Windows Service 注册脚本
# 文件: setup/register_v19_services.ps1
# 编码: UTF-8
# =============================================================================
# 用途: 将 V19 Admin + API 注册为 Windows 服务，实现开机自启 + 崩溃自动恢复
# 使用: 以管理员身份运行 PowerShell，执行此脚本
#   .\setup\register_v19_services.ps1
#   .\setup\register_v19_services.ps1 -Remove    (仅移除现有服务)
#   .\setup\register_v19_services.ps1 -Restart   (重启服务)
# =============================================================================

[CmdletBinding()]
param(
    [switch]$Remove,
    [switch]$Restart,
    [switch]$Confirm
)

# ---- 配置 ----
$projectRoot = Split-Path -Parent $PSScriptRoot
$publishDir  = Join-Path $projectRoot "publish"

$services = @(
    @{
        Name        = "PerfumeShopV19_API"
        Display     = "PerfumeShop V19 API Service (port 5000)"
        Description = "香氛电商 V19 ASP.NET Core API 服务，处理前台页面请求与 REST API"
        ExePath     = Join-Path $publishDir "PerfumeShop.Api\PerfumeShop.Api.exe"
        WorkDir     = Join-Path $publishDir "PerfumeShop.Api"
        Port        = 5000
    },
    @{
        Name        = "PerfumeShopV19_Admin"
        Display     = "PerfumeShop V19 Admin Service (port 5207)"
        Description = "香氛电商 V19 Blazor Server 管理后台，处理管理员登录与运营管理"
        ExePath     = Join-Path $publishDir "PerfumeShop.Admin\PerfumeShop.Admin.exe"
        WorkDir     = Join-Path $publishDir "PerfumeShop.Admin"
        Port        = 5207
    }
)

$ErrorActionPreference = 'Stop'

# ---- 颜色输出 ----
function Write-Step { Write-Host ">>> $args" -ForegroundColor Cyan }
function Write-Ok   { Write-Host "  ✓ $args" -ForegroundColor Green }
function Write-Warn { Write-Host "  ⚠ $args" -ForegroundColor Yellow }
function Write-Err  { Write-Host "  ✗ $args" -ForegroundColor Red }

# ---- 检查管理员权限 ----
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Host "此脚本需要管理员权限。请以管理员身份运行 PowerShell。" -ForegroundColor Red
    Write-Host "右键 PowerShell → 以管理员身份运行，然后重新执行此脚本。" -ForegroundColor Yellow
    exit 1
}

Write-Host ''
Write-Host '============================================' -ForegroundColor Cyan
Write-Host '  PerfumeShop V19 Windows Service 注册工具' -ForegroundColor Cyan
Write-Host '============================================' -ForegroundColor Cyan
Write-Host ''

# ---- 模式 1: 仅移除 ----
if ($Remove) {
    Write-Host '将移除所有 V19 Windows 服务...' -ForegroundColor Yellow
    foreach ($svc in $services) {
        $existing = Get-Service -Name $svc.Name -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Step "停止并删除服务: $($svc.Name)"
            try { sc.exe stop $svc.Name 2>&1 | Out-Null } catch {}
            Start-Sleep -Seconds 2
            sc.exe delete $svc.Name 2>&1 | Out-Null
            Write-Ok "已删除: $($svc.Name)"
        } else {
            Write-Warn "服务不存在: $($svc.Name)"
        }
    }
    Write-Host ''
    Write-Host '✓ 所有 V19 服务已移除。' -ForegroundColor Green
    exit 0
}

# ---- 模式 2: 重启 ----
if ($Restart) {
    foreach ($svc in $services) {
        $existing = Get-Service -Name $svc.Name -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Step "重启服务: $($svc.Name)"
            Restart-Service -Name $svc.Name -Force
            Start-Sleep -Seconds 3
            $status = (Get-Service -Name $svc.Name).Status
            if ($status -eq 'Running') { Write-Ok "已重启 (状态: $status)" }
            else { Write-Err "重启后状态异常: $status" }
        } else {
            Write-Err "服务不存在: $($svc.Name) (请先运行安装)" 
        }
    }
    Write-Host ''
    exit
}

# ---- 确认操作 ----
Write-Host '此操作将:' -ForegroundColor Yellow
Write-Host '  1. 停止并删除旧版 V19 服务'
Write-Host '  2. 注册 Windows 服务（开机自启）'
Write-Host '  3. 配置崩溃自动恢复策略'
Write-Host '  4. 立即启动服务'
Write-Host ''

if (-not $Confirm) {
    $answer = Read-Host '确认执行？(yes/no)'
    if ($answer -ne 'yes') {
        Write-Host '已取消。' -ForegroundColor Yellow
        exit 0
    }
}

# =============================================================================
# 步骤 1: 检查发布文件
# =============================================================================
Write-Step "检查发布文件..."

foreach ($svc in $services) {
    if (-not (Test-Path $svc.ExePath)) {
        Write-Err "找不到发布文件: $($svc.ExePath)"
        Write-Host "  请先执行: dotnet publish --configuration Release" -ForegroundColor Yellow
        exit 1
    }
    if (-not (Test-Path $svc.WorkDir)) {
        Write-Err "找不到工作目录: $($svc.WorkDir)"
        exit 1
    }
    Write-Ok "找到: $($svc.ExePath)"
}

# =============================================================================
# 步骤 2: 停止并删除旧服务
# =============================================================================
Write-Step "清理旧服务..."

foreach ($svc in $services) {
    $existing = Get-Service -Name $svc.Name -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  正在停止服务: $($svc.Name)..."
        try {
            sc.exe stop $svc.Name 2>&1 | Out-Null
            Start-Sleep -Seconds 3
        } catch { }

        Write-Host "  正在删除服务: $($svc.Name)..."
        sc.exe delete $svc.Name 2>&1 | Out-Null
        Start-Sleep -Seconds 1
        Write-Ok "已删除旧服务: $($svc.Name)"
    } else {
        Write-Host "  服务不存在，跳过: $($svc.Name)"
    }
}

# =============================================================================
# 步骤 3: 注册新服务
# =============================================================================
Write-Step "注册 Windows Service..."

foreach ($svc in $services) {
    # 构建 binPath: exe 路径 + 参数
    # 使用 --urls 指定端口，--contentRoot 指定内容根目录
    $binPath = "`"$($svc.ExePath)`" --urls http://localhost:$($svc.Port) --contentRoot `"$($svc.WorkDir)`""

    Write-Host "  创建服务: $($svc.Name)"
    Write-Host "    路径: $($svc.ExePath)"
    Write-Host "    端口: $($svc.Port)"

    # 创建服务 (start=auto 表示开机自启)
    $result = sc.exe create $svc.Name `
        binPath= $binPath `
        start= auto `
        DisplayName= $svc.Display `
        obj= LocalSystem 2>&1

    if ($LASTEXITCODE -ne 0) {
        Write-Err "创建服务失败: $result"
        exit 1
    }

    # 设置服务描述
    sc.exe description $svc.Name $svc.Description 2>&1 | Out-Null

    # ===================== 配置失败恢复策略 =====================
    # 失败后操作: restart/30000 (30秒后重启) → restart/30000 → restart/30000
    # 重置失败计数: 86400 秒 (24小时)
    # 重启服务延迟: 30000 毫秒 (30秒)
    sc.exe failure $svc.Name `
        reset= 86400 `
        reboot= "" `
        command= "" `
        actions= restart/30000/restart/30000/restart/30000 2>&1 | Out-Null

    Write-Ok "服务已创建: $($svc.Name) (开机自启 + 3次崩溃自动重启)"
}

# =============================================================================
# 步骤 4: 启动服务
# =============================================================================
Write-Step "启动服务..."

foreach ($svc in $services) {
    Write-Host "  启动: $($svc.Name)..."
    sc.exe start $svc.Name 2>&1 | Out-Null
    Start-Sleep -Seconds 5

    $status = (Get-Service -Name $svc.Name -ErrorAction SilentlyContinue).Status
    if ($status -eq 'Running') {
        Write-Ok "$($svc.Name) — 运行中 (端口 $($svc.Port))"
    } else {
        Write-Err "$($svc.Name) — 状态异常: $status"
        Write-Host "    查看日志: Get-EventLog -LogName Application -Source '$($svc.Name)' -Newest 10" -ForegroundColor Yellow
    }
}

# =============================================================================
# 步骤 5: 健康检查
# =============================================================================
Write-Host ''
Write-Step "健康检查..."

Start-Sleep -Seconds 3

foreach ($svc in $services) {
    try {
        $url = "http://localhost:$($svc.Port)/"
        $resp = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 15
        if ($resp.StatusCode -eq 200) {
            Write-Ok "$($svc.Name) HTTP 200 OK → $url"
        } else {
            Write-Warn "$($svc.Name) HTTP $($resp.StatusCode) → $url"
        }
    } catch {
        Write-Warn "$($svc.Name) 健康检查未通过: $($_.Exception.Message)"
        Write-Host "    服务可能仍在启动中，请等待 15-30 秒后手动访问验证。" -ForegroundColor Yellow
    }
}

# =============================================================================
# 结果汇总
# =============================================================================
Write-Host ''
Write-Host '============================================' -ForegroundColor Green
Write-Host '  V19 服务注册完成！' -ForegroundColor Green
Write-Host '============================================' -ForegroundColor Green
Write-Host ''
Write-Host '  访问地址:' -ForegroundColor Cyan
Write-Host '    V19 Admin:  http://localhost:5207/login' -ForegroundColor White
Write-Host '    V19 API:    http://localhost:5000/swagger' -ForegroundColor White
Write-Host '    V18 ASP:    http://localhost/admin/login.asp' -ForegroundColor White
Write-Host ''
Write-Host '  管理命令:' -ForegroundColor Cyan
Write-Host '    Get-Service PerfumeShop* ' -ForegroundColor Gray -NoNewline
Write-Host '            # 查看服务状态'
Write-Host '    Restart-Service PerfumeShopV19_Admin ' -ForegroundColor Gray -NoNewline
Write-Host ' # 重启管理后台'
Write-Host '    Remove-Service PerfumeShopV19_Admin ' -ForegroundColor Gray -NoNewline
Write-Host '  # 或运行: .\setup\register_v19_services.ps1 -Remove'
Write-Host ''
Write-Host '  注意: 服务崩溃后将在 30 秒内自动重启（最多连续重启 3 次）。' -ForegroundColor Yellow
Write-Host ''

