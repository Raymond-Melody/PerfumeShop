# =============================================================================
# PerfumeShop V19 一键回滚脚本
# 文件: setup/v19_rollback.ps1
# 编码: UTF-8
# =============================================================================
# 用途: 将 V19 回滚到 V18 Classic ASP + V19 双路由灰度状态
# 操作:
#   1. 恢复 web.config 为灰度版（V18/V19 双路由）
#   2. 恢复数据库备份
#   3. 重启 IIS 应用池
#   4. 验证回滚成功
# =============================================================================

[CmdletBinding()]
param(
    [string]$ServerInstance   = 'localhost\YOURPERFUME',
    [string]$DatabaseName     = 'PerfumeShop',
    [string]$BackupDir        = 'F:\archive\V18_ClassicASP\db',
    [string]$WebRoot          = (Join-Path $PSScriptRoot '..'),
    [string]$IISAppPool       = 'DefaultAppPool',
    [string]$LogDir           = (Join-Path $PSScriptRoot '..\logs'),
    [switch]$Confirm
)

$ErrorActionPreference = 'Stop'

# ---- 初始化日志 ----
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}
$logTimestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile = Join-Path $LogDir "v19_rollback_${logTimestamp}.log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] $Message"
    $line | Out-File -FilePath $logFile -Append -Encoding utf8
    switch ($Level) {
        'ERROR'   { Write-Host $line -ForegroundColor Red }
        'WARN'    { Write-Host $line -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $line -ForegroundColor Green }
        default   { Write-Host $line }
    }
}

# =============================================================================
# 确认操作
# =============================================================================

Write-Host ''
Write-Host '============================================' -ForegroundColor Red
Write-Host '  V19 一键回滚 -> V18 灰度共存' -ForegroundColor Red
Write-Host '============================================' -ForegroundColor Red
Write-Host ''
Write-Host '此操作将:' -ForegroundColor Yellow
Write-Host '  1. 恢复 web.config 为 V18/V19 双路由灰度版'
Write-Host '  2. 从最近备份恢复数据库'
Write-Host '  3. 重启 IIS 应用池'
Write-Host '  4. 验证 V18 页面可访问'
Write-Host ''

if (-not $Confirm) {
    $answer = Read-Host '确认执行回滚？(yes/no)'
    if ($answer -ne 'yes') {
        Write-Host '已取消回滚操作。' -ForegroundColor Yellow
        exit 0
    }
}

$startTime = Get-Date
Write-Log '回滚操作开始'

# =============================================================================
# 步骤 1: 恢复 web.config 为灰度版
# =============================================================================

Write-Log '===== 步骤 1: 恢复 web.config 灰度版 ====='

$webConfigPath   = Join-Path $WebRoot 'web.config'
$webConfigBackup = Join-Path $WebRoot 'web.config.v19-final'

# 备份当前 V19 最终版 web.config
if (Test-Path $webConfigPath) {
    Copy-Item -Path $webConfigPath -Destination $webConfigBackup -Force
    Write-Log '已备份当前 web.config -> web.config.v19-final'
}

# 生成灰度版 web.config（V18 Classic ASP + V19 双路由）
function Write-GrayWebConfig {
    param([string]$OutputPath)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
    [void]$sb.AppendLine('<configuration>')
    [void]$sb.AppendLine('    <system.web>')
    [void]$sb.AppendLine('        <customErrors mode="RemoteOnly" defaultRedirect="/error.asp" />')
    [void]$sb.AppendLine('        <httpRuntime requestValidationMode="2.0" executionTimeout="300" maxRequestLength="102400" />')
    [void]$sb.AppendLine('    </system.web>')
    [void]$sb.AppendLine('    <location path="database">')
    [void]$sb.AppendLine('        <system.webServer>')
    [void]$sb.AppendLine('            <security>')
    [void]$sb.AppendLine('                <requestFiltering>')
    [void]$sb.AppendLine('                    <fileExtensions allowUnlisted="true">')
    [void]$sb.AppendLine('                        <add fileExtension=".mdb" allowed="false" />')
    [void]$sb.AppendLine('                        <add fileExtension=".sql" allowed="false" />')
    [void]$sb.AppendLine('                    </fileExtensions>')
    [void]$sb.AppendLine('                </requestFiltering>')
    [void]$sb.AppendLine('            </security>')
    [void]$sb.AppendLine('        </system.webServer>')
    [void]$sb.AppendLine('    </location>')
    [void]$sb.AppendLine('    <location path="includes">')
    [void]$sb.AppendLine('        <system.webServer>')
    [void]$sb.AppendLine('            <security>')
    [void]$sb.AppendLine('                <requestFiltering>')
    [void]$sb.AppendLine('                    <fileExtensions allowUnlisted="true">')
    [void]$sb.AppendLine('                        <add fileExtension=".asp" allowed="false" />')
    [void]$sb.AppendLine('                    </fileExtensions>')
    [void]$sb.AppendLine('                </requestFiltering>')
    [void]$sb.AppendLine('            </security>')
    [void]$sb.AppendLine('        </system.webServer>')
    [void]$sb.AppendLine('    </location>')
    [void]$sb.AppendLine('    <system.webServer>')
    [void]$sb.AppendLine('        <urlCompression doStaticCompression="true" doDynamicCompression="true" />')
    [void]$sb.AppendLine('        <staticContent>')
    [void]$sb.AppendLine('            <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="7.00:00:00" />')
    [void]$sb.AppendLine('        </staticContent>')
    [void]$sb.AppendLine('        <httpProtocol>')
    [void]$sb.AppendLine('            <customHeaders>')
    [void]$sb.AppendLine('                <add name="X-Frame-Options" value="SAMEORIGIN" />')
    [void]$sb.AppendLine('                <add name="X-Content-Type-Options" value="nosniff" />')
    [void]$sb.AppendLine('                <add name="Strict-Transport-Security" value="max-age=31536000; includeSubDomains" />')
    [void]$sb.AppendLine('            </customHeaders>')
    [void]$sb.AppendLine('        </httpProtocol>')
    [void]$sb.AppendLine('        <security>')
    [void]$sb.AppendLine('            <requestFiltering>')
    [void]$sb.AppendLine('                <requestLimits maxAllowedContentLength="10485760" />')
    [void]$sb.AppendLine('            </requestFiltering>')
    [void]$sb.AppendLine('        </security>')
    [void]$sb.AppendLine('        <defaultDocument>')
    [void]$sb.AppendLine('            <files>')
    [void]$sb.AppendLine('                <clear />')
    [void]$sb.AppendLine('                <add value="index.asp" />')
    [void]$sb.AppendLine('                <add value="Default.asp" />')
    [void]$sb.AppendLine('                <add value="index.html" />')
    [void]$sb.AppendLine('            </files>')
    [void]$sb.AppendLine('        </defaultDocument>')
    [void]$sb.AppendLine('        <asp enableParentPaths="true" scriptErrorSentToBrowser="true">')
    [void]$sb.AppendLine('            <limits scriptTimeout="00:05:00" />')
    [void]$sb.AppendLine('        </asp>')
    [void]$sb.AppendLine('        <rewrite>')
    [void]$sb.AppendLine('            <rules>')
    [void]$sb.AppendLine('                <rule name="ApiV2" stopProcessing="true">')
    [void]$sb.AppendLine('                    <match url="^api/v2/(.*)" />')
    [void]$sb.AppendLine('                    <action type="Rewrite" url="http://localhost:5000/api/{R:1}" />')
    [void]$sb.AppendLine('                </rule>')
    [void]$sb.AppendLine('                <rule name="AdminV2" stopProcessing="true">')
    [void]$sb.AppendLine('                    <match url="^admin/v2/(.*)" />')
    [void]$sb.AppendLine('                    <action type="Rewrite" url="http://localhost:5207/admin/{R:1}" />')
    [void]$sb.AppendLine('                </rule>')
    [void]$sb.AppendLine('            </rules>')
    [void]$sb.AppendLine('        </rewrite>')
    [void]$sb.AppendLine('    </system.webServer>')
    [void]$sb.AppendLine('</configuration>')

    [System.IO.File]::WriteAllText($OutputPath, $sb.ToString(), [System.Text.Encoding]::UTF8)
}

Write-GrayWebConfig -OutputPath $webConfigPath
Write-Log '已恢复灰度版 web.config (V18 + V19 双路由)' -Level 'SUCCESS'

# =============================================================================
# 步骤 2: 恢复数据库备份
# =============================================================================

Write-Log '===== 步骤 2: 恢复数据库备份 ====='

# 查找最近的备份文件
$latestBackup = Get-ChildItem -Path $BackupDir -Filter '*.bak' -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if (-not $latestBackup) {
    Write-Log "未找到数据库备份文件: $BackupDir\*.bak" -Level 'ERROR'
    Write-Log '请手动恢复数据库' -Level 'ERROR'
} else {
    $bakSizeMB = [math]::Round($latestBackup.Length / 1MB, 2)
    Write-Log "找到最近备份: $($latestBackup.Name) ($bakSizeMB MB)"

    $bakPath = $latestBackup.FullName
    $restoreSql = 'USE [master];' + "`n" + `
        "ALTER DATABASE [$DatabaseName] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;" + "`n" + `
        "RESTORE DATABASE [$DatabaseName] FROM DISK = N'$bakPath' WITH REPLACE, STATS = 10;" + "`n" + `
        "ALTER DATABASE [$DatabaseName] SET MULTI_USER;"

    Write-Log '正在恢复数据库（可能需要几分钟）...'
    $result = & sqlcmd -S $ServerInstance -E -C -Q $restoreSql 2>&1

    if ($LASTEXITCODE -ne 0) {
        Write-Log "数据库恢复失败: $result" -Level 'ERROR'
    } else {
        Write-Log '数据库恢复完成' -Level 'SUCCESS'
    }
}

# =============================================================================
# 步骤 3: 重启 IIS 应用池
# =============================================================================

Write-Log '===== 步骤 3: 重启 IIS 应用池 ====='

try {
    & iisreset /stop 2>&1 | Out-Null
    Start-Sleep -Seconds 2
    & iisreset /start 2>&1 | Out-Null
    Start-Sleep -Seconds 3
    Write-Log 'IIS 已重启' -Level 'SUCCESS'
} catch {
    Write-Log "IIS 重启失败: $($_.Exception.Message)" -Level 'ERROR'
    Write-Log '请手动执行: iisreset' -Level 'WARN'
}

# =============================================================================
# 步骤 4: 验证回滚成功
# =============================================================================

Write-Log '===== 步骤 4: 验证回滚 ====='

$verifyUrls = @(
    @{ Url = 'http://localhost/index.asp';    Desc = 'V18 首页' },
    @{ Url = 'http://localhost/login.asp';    Desc = 'V18 登录' },
    @{ Url = 'http://localhost/products.asp'; Desc = 'V18 商品列表' }
)

$allPass = $true
foreach ($item in $verifyUrls) {
    try {
        $resp = Invoke-WebRequest -Uri $item.Url -UseBasicParsing -TimeoutSec 10
        if ($resp.StatusCode -eq 200) {
            Write-Log "$($item.Desc) [$($item.Url)] -> 200 OK" -Level 'SUCCESS'
        } else {
            Write-Log "$($item.Desc) [$($item.Url)] -> $($resp.StatusCode)" -Level 'WARN'
            $allPass = $false
        }
    } catch {
        Write-Log "$($item.Desc) [$($item.Url)] -> 访问失败: $($_.Exception.Message)" -Level 'ERROR'
        $allPass = $false
    }
}

# 验证 V19 API 仍可访问（灰度路由）
try {
    $resp = Invoke-WebRequest -Uri 'http://localhost/api/v2/health' -UseBasicParsing -TimeoutSec 10
    Write-Log "V19 API 灰度路由 [/api/v2/health] -> $($resp.StatusCode)" -Level 'SUCCESS'
} catch {
    Write-Log 'V19 API 灰度路由不可用（预期行为，如 V19 进程未启动）' -Level 'WARN'
}

# =============================================================================
# 结果汇总
# =============================================================================

$endTime  = Get-Date
$duration = ($endTime - $startTime).TotalSeconds
$durationRound = [math]::Round($duration, 1)

Write-Host ''
if ($allPass) {
    Write-Host '============================================' -ForegroundColor Green
    Write-Host '  回滚完成 - V18 页面已恢复' -ForegroundColor Green
} else {
    Write-Host '============================================' -ForegroundColor Yellow
    Write-Host '  回滚已完成 - 部分验证未通过，请检查日志' -ForegroundColor Yellow
}
Write-Host "  耗时: $durationRound 秒" -ForegroundColor Cyan
Write-Host "  日志: $logFile" -ForegroundColor Cyan
Write-Host '============================================' -ForegroundColor Cyan

if ($allPass) {
    Write-Log "回滚完成, 耗时 ${durationRound}s, 验证全部通过"
} else {
    Write-Log "回滚完成, 耗时 ${durationRound}s, 部分验证未通过" -Level 'WARN'
}
