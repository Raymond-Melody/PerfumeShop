# PerfumeShop 数据库完整备份脚本
# 用法: .\backup_database.ps1 [-BackupDir <路径>] [-RetentionDays <天数>] [-ServerInstance <实例>] [-DatabaseName <库名>]

param(
    [string]$BackupDir = (Join-Path (Get-Location) "database\backups"),
    [int]$RetentionDays = 30,
    [string]$ServerInstance = "localhost\YOURPERFUME",
    [string]$DatabaseName = "PerfumeShop"
)

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  PerfumeShop 数据库备份工具" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 1. 创建备份目录
if (-not (Test-Path $BackupDir)) {
    New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null
    Write-Host "[OK] 已创建备份目录: $BackupDir" -ForegroundColor Green
} else {
    Write-Host "[OK] 备份目录已存在: $BackupDir" -ForegroundColor Green
}

# 2. 生成带时间戳的备份文件名
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupFile = Join-Path $BackupDir "${DatabaseName}_${timestamp}.bak"

Write-Host ""
Write-Host "备份参数:" -ForegroundColor Yellow
Write-Host "  服务器实例: $ServerInstance"
Write-Host "  数据库名称: $DatabaseName"
Write-Host "  备份文件:   $backupFile"
Write-Host "  保留天数:   $RetentionDays"
Write-Host ""

# 3. 执行完整数据库备份
Write-Host "正在执行数据库备份..." -ForegroundColor Yellow
$startTime = Get-Date

$backupSql = "BACKUP DATABASE [$DatabaseName] TO DISK = N'$backupFile' WITH FORMAT, INIT, NAME = N'${DatabaseName}-Full Backup ${timestamp}', STATS = 10"
$backupResult = & sqlcmd -S $ServerInstance -E -C -Q $backupSql 2>&1

if ($LASTEXITCODE -ne 0) {
    Write-Host "[失败] 数据库备份失败!" -ForegroundColor Red
    Write-Host $backupResult
    exit 1
}

Write-Host $backupResult
Write-Host "[OK] 数据库备份完成" -ForegroundColor Green
Write-Host ""

# 4. 验证备份完整性
Write-Host "正在验证备份完整性..." -ForegroundColor Yellow
$verifySql = "RESTORE VERIFYONLY FROM DISK = N'$backupFile'"
$verifyResult = & sqlcmd -S $ServerInstance -E -C -Q $verifySql 2>&1

if ($LASTEXITCODE -ne 0) {
    Write-Host "[失败] 备份验证失败!" -ForegroundColor Red
    Write-Host $verifyResult
    exit 1
}

Write-Host "[OK] 备份验证通过 - 文件完整有效" -ForegroundColor Green
Write-Host ""

# 5. 清理超期旧备份
$cutoffDate = (Get-Date).AddDays(-$RetentionDays)
$oldBackups = Get-ChildItem $BackupDir -Filter "*.bak" | Where-Object { $_.LastWriteTime -lt $cutoffDate }

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

# 6. 输出备份结果摘要
$endTime = Get-Date
$fileInfo = Get-Item $backupFile
$duration = ($endTime - $startTime).TotalSeconds
$sizeMB = [math]::Round($fileInfo.Length / 1MB, 2)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  备份结果摘要" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  文件名: $($fileInfo.Name)"
Write-Host "  大  小: $sizeMB MB ($($fileInfo.Length) bytes)"
Write-Host "  耗  时: $([math]::Round($duration, 1)) 秒"
Write-Host "  时  间: $($fileInfo.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "  验  证: 通过" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan

# 列出当前所有备份
$allBackups = Get-ChildItem $BackupDir -Filter "*.bak" | Sort-Object LastWriteTime -Descending
Write-Host ""
Write-Host "当前备份文件列表 ($($allBackups.Count) 个):" -ForegroundColor Yellow
foreach ($bak in $allBackups) {
    $bakSize = [math]::Round($bak.Length / 1MB, 2)
    Write-Host "  $($bak.Name)  ($bakSize MB)  $($bak.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))"
}

