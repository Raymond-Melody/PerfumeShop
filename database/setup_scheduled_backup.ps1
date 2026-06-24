# PerfumeShop Windows Task Scheduler 备份任务注册脚本
# 用法: .\setup_scheduled_backup.ps1 [-BackupDir <路径>] [-RetentionDays <天数>] [-NotificationEmail <邮箱>]
#       管理员身份运行: 右键 → 以管理员身份运行 PowerShell → 执行此脚本

param(
    [string]$BackupDir = "",
    [int]$RetentionDays = 30,
    [string]$NotificationEmail = "",
    [string]$TaskName = "PerfumeShop_每日数据库备份"
)

$ErrorActionPreference = "Continue"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# 默认备份目录为脚本所在目录的 backups 子目录
if ($BackupDir -eq "") {
    $BackupDir = Join-Path $scriptDir "backups"
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  PerfumeShop 定时备份任务注册工具" -ForegroundColor Cyan
Write-Host "  Version: V10.4" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 检查管理员权限
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "[错误] 此脚本需要管理员权限才能注册计划任务。" -ForegroundColor Red
    Write-Host "请右键 PowerShell → 以管理员身份运行，然后重新执行此脚本。" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "按任意键退出..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# 1. 验证 backup_database.ps1 存在
$backupScript = Join-Path $scriptDir "backup_database.ps1"
if (-not (Test-Path $backupScript)) {
    Write-Host "[错误] 找不到备份脚本: $backupScript" -ForegroundColor Red
    exit 1
}
Write-Host "[OK] 备份脚本: $backupScript" -ForegroundColor Green

# 2. 创建备份目录
if (-not (Test-Path $BackupDir)) {
    New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null
    Write-Host "[OK] 已创建备份目录: $BackupDir" -ForegroundColor Green
} else {
    Write-Host "[OK] 备份目录已存在: $BackupDir" -ForegroundColor Green
}

# 3. 构建 PowerShell 执行的参数字符串
$arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$backupScript`" -BackupDir `"$BackupDir`" -RetentionDays $RetentionDays"
if ($NotificationEmail -ne "") {
    $arguments += " -NotificationEmail `"$NotificationEmail`""
}

# 4. 删除已有的同名任务（如果存在）
$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "[信息] 已存在同名任务，正在更新..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

# 5. 创建计划任务触发器 — 每日凌晨 2:00
$trigger = New-ScheduledTaskTrigger -Daily -At 2:00AM
Write-Host "[OK] 触发器: 每日 02:00 AM" -ForegroundColor Green

# 6. 创建计划任务操作
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $arguments
Write-Host "[OK] 操作: powershell.exe $arguments" -ForegroundColor Green

# 7. 设置任务配置
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -MultipleInstances IgnoreNew `
    -WakeToRun

# 8. 注册计划任务（以 SYSTEM 账户运行，确保有足够权限执行备份）
$principal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest

try {
    Register-ScheduledTask -TaskName $TaskName `
        -Trigger $trigger `
        -Action $action `
        -Settings $settings `
        -Principal $principal `
        -Description "PerfumeShop 数据库每日完整备份任务 (V10.4)，每日凌晨 2:00 自动执行" `
        -Force | Out-Null
    
    Write-Host ""
    Write-Host "[OK] 计划任务注册成功!" -ForegroundColor Green
} catch {
    Write-Host ""
    Write-Host "[错误] 计划任务注册失败: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "备选方案: 手动创建计划任务" -ForegroundColor Yellow
    Write-Host "1. 打开「任务计划程序」(taskschd.msc)" -ForegroundColor Yellow
    Write-Host "2. 创建基本任务 → 名称: $TaskName" -ForegroundColor Yellow
    Write-Host "3. 触发器: 每天 → 02:00" -ForegroundColor Yellow
    Write-Host "4. 操作: 启动程序 → 程序: powershell.exe" -ForegroundColor Yellow
    Write-Host "   参数: $arguments" -ForegroundColor Yellow
    exit 1
}

# 9. 输出最终配置摘要
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  定时备份任务配置摘要" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  任务名称:     $TaskName"
Write-Host "  执行时间:     每日 02:00 AM"
Write-Host "  备份目录:     $BackupDir"
Write-Host "  保留天数:     $RetentionDays"
if ($NotificationEmail -ne "") {
    Write-Host "  通知邮箱:     $NotificationEmail"
}
Write-Host "  运行账户:     NT AUTHORITY\SYSTEM"
Write-Host "  备份脚本:     $backupScript"
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 10. 记录注册日志
$logFile = Join-Path $scriptDir "schedule_setup_log.txt"
$logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | 任务已注册/更新 | 任务名称: $TaskName | 备份目录: $BackupDir | 保留天数: $RetentionDays"
Add-Content -Path $logFile -Value $logEntry -Encoding UTF8
Write-Host "[OK] 注册日志已保存至: $logFile" -ForegroundColor Green
Write-Host ""

Write-Host "下一步: 可在「任务计划程序」(taskschd.msc) 中查看和管理此任务。" -ForegroundColor Cyan
Write-Host "如需手动测试: 在任务计划程序中右键该任务 → 运行" -ForegroundColor Cyan
Write-Host ""

Write-Host "按任意键退出..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
