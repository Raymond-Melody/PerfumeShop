#Requires -Version 5.1
# =============================================================================
# PerfumeShop V19 里程碑备份脚本
# 用途：在执行数据库迁移前，创建带里程碑版本号的完整备份
# 依赖：sqlcmd 命令行工具、SQL Server 实例可访问
# 兼容：PowerShell 5.1+
# =============================================================================
# 用法示例：
#   .\backup_before_migration.ps1 -Milestone M1
#   .\backup_before_migration.ps1 -Milestone M2 -RetentionDays 60
# =============================================================================

param(
    [Parameter(Mandatory = $true, HelpMessage = "里程碑版本号，如 M1、M2")]
    [ValidatePattern('^M\d+$')]
    [string]$Milestone,

    [int]$RetentionDays = 30,

    [string]$ServerInstance = "localhost\YOURPERFUME",

    [string]$DatabaseName = "PerfumeShop"
)

$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# 路径计算：基于脚本所在目录向上推算项目根目录
# ---------------------------------------------------------------------------
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$projectRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
$backupDir = Join-Path $projectRoot "database\backups"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  PerfumeShop V19 里程碑备份工具" -ForegroundColor Cyan
Write-Host "  里程碑: $Milestone" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ---------------------------------------------------------------------------
# 1. 创建备份目录
# ---------------------------------------------------------------------------
if (-not (Test-Path $backupDir)) {
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    Write-Host "[OK] 已创建备份目录: $backupDir" -ForegroundColor Green
} else {
    Write-Host "[OK] 备份目录已存在: $backupDir" -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# 2. 生成带里程碑版本号与时间戳的备份文件名
#    格式: PerfumeShop_PreM1_20260710.bak
# ---------------------------------------------------------------------------
$dateStamp = Get-Date -Format "yyyyMMdd"
$backupFileName = "${DatabaseName}_Pre${Milestone}_${dateStamp}.bak"
$backupFile = Join-Path $backupDir $backupFileName

# 如果同一天同一里程碑的备份已存在，追加序号避免覆盖
if (Test-Path $backupFile) {
    $seq = 2
    do {
        $backupFileName = "${DatabaseName}_Pre${Milestone}_${dateStamp}_${seq}.bak"
        $backupFile = Join-Path $backupDir $backupFileName
        $seq++
    } while (Test-Path $backupFile)
    Write-Host "[提示] 检测到已有同名备份，使用序号: $($seq - 1)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "备份参数:" -ForegroundColor Yellow
Write-Host "  服务器实例: $ServerInstance"
Write-Host "  数据库名称: $DatabaseName"
Write-Host "  里程碑版本: $Milestone"
Write-Host "  备份文件:   $backupFile"
Write-Host "  保留天数:   $RetentionDays"
Write-Host ""

# ---------------------------------------------------------------------------
# 3. 执行完整数据库备份（FORMAT + INIT 覆盖写入）
# ---------------------------------------------------------------------------
Write-Host "正在执行数据库备份..." -ForegroundColor Yellow
$startTime = Get-Date

$backupSql = "BACKUP DATABASE [$DatabaseName] TO DISK = N'$backupFile' WITH FORMAT, INIT, NAME = N'${DatabaseName}-Pre${Milestone}-${dateStamp}', STATS = 10"
$backupResult = & sqlcmd -S $ServerInstance -E -C -Q $backupSql 2>&1

if ($LASTEXITCODE -ne 0) {
    Write-Host "[失败] 数据库备份失败!" -ForegroundColor Red
    Write-Host ($backupResult -join "`n")
    exit 1
}

Write-Host ($backupResult -join "`n")
Write-Host "[OK] 数据库备份完成" -ForegroundColor Green
Write-Host ""

# ---------------------------------------------------------------------------
# 4. 验证备份完整性（RESTORE VERIFYONLY）
# ---------------------------------------------------------------------------
Write-Host "正在验证备份完整性..." -ForegroundColor Yellow
$verifySql = "RESTORE VERIFYONLY FROM DISK = N'$backupFile'"
$verifyResult = & sqlcmd -S $ServerInstance -E -C -Q $verifySql 2>&1

if ($LASTEXITCODE -ne 0) {
    Write-Host "[失败] 备份验证失败!" -ForegroundColor Red
    Write-Host ($verifyResult -join "`n")
    exit 1
}

Write-Host "[OK] 备份验证通过 - 文件完整有效" -ForegroundColor Green
Write-Host ""

# ---------------------------------------------------------------------------
# 5. 清理超过 RetentionDays 天的旧备份（*.bak）
# ---------------------------------------------------------------------------
$cutoffDate = (Get-Date).AddDays(-$RetentionDays)
$oldBackups = Get-ChildItem $backupDir -Filter "*.bak" | Where-Object { $_.LastWriteTime -lt $cutoffDate }

if ($oldBackups.Count -gt 0) {
    Write-Host "正在清理 $($oldBackups.Count) 个超过 $RetentionDays 天的旧备份..." -ForegroundColor Yellow
    foreach ($old in $oldBackups) {
        Remove-Item $old.FullName -Force
        Write-Host "  已删除: $($old.Name)" -ForegroundColor DarkGray
    }
    Write-Host "[OK] 旧备份清理完成" -ForegroundColor Green
} else {
    Write-Host "[OK] 无需清理旧备份" -ForegroundColor Green
}

Write-Host ""

# ---------------------------------------------------------------------------
# 6. 输出备份结果摘要
# ---------------------------------------------------------------------------
$endTime = Get-Date
$fileInfo = Get-Item $backupFile
$duration = ($endTime - $startTime).TotalSeconds
$sizeMB = [math]::Round($fileInfo.Length / 1MB, 2)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  备份结果摘要" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  里程碑: $Milestone"
Write-Host "  文件名: $($fileInfo.Name)"
Write-Host "  大  小: $sizeMB MB ($($fileInfo.Length) bytes)"
Write-Host "  耗  时: $([math]::Round($duration, 1)) 秒"
Write-Host "  时  间: $($fileInfo.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "  验  证: 通过" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan

# 列出当前所有备份
$allBackups = Get-ChildItem $backupDir -Filter "*.bak" | Sort-Object LastWriteTime -Descending
Write-Host ""
Write-Host "当前备份文件列表 ($($allBackups.Count) 个):" -ForegroundColor Yellow
foreach ($bak in $allBackups) {
    $bakSize = [math]::Round($bak.Length / 1MB, 2)
    Write-Host "  $($bak.Name)  ($bakSize MB)  $($bak.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))"
}

