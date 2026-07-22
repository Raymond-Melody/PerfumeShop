# =============================================================================
# PerfumeShop V19 最终数据迁移脚本
# 文件: setup/v19_final_migrate.ps1
# 编码: UTF-8
# =============================================================================
# 用途: 从 V18 生产数据库全量迁移核心历史数据到 V19
# 迁移范围: Orders, Users, ProductReviews, Addresses (4 张表)
# 设计原则:
#   - 参数化连接字符串
#   - 幂等执行（SET IDENTITY_INSERT + MERGE 避免重复）
#   - 每步骤前后 COUNT 校验
#   - 完整日志输出到文件
# =============================================================================
# 执行方式:
#   .\setup\v19_final_migrate.ps1 `
#       -SourceServer "localhost\YOURPERFUME" `
#       -SourceDatabase "PerfumeShop" `
#       -TargetServer "localhost\YOURPERFUME" `
#       -TargetDatabase "PerfumeShopV19"
# =============================================================================

[CmdletBinding()]
param(
    [string]$SourceServer    = "localhost\YOURPERFUME",
    [string]$SourceDatabase  = "PerfumeShop",
    [string]$TargetServer    = "localhost\YOURPERFUME",
    [string]$TargetDatabase  = "PerfumeShopV19",
    [string]$LogDir          = (Join-Path $PSScriptRoot "..\logs")
)

$ErrorActionPreference = "Stop"

# ---- 初始化日志 ----
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}
$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile    = Join-Path $LogDir "v19_final_migrate_${timestamp}.log"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    $line | Out-File -FilePath $logFile -Append -Encoding utf8
    switch ($Level) {
        "ERROR"   { Write-Host $line -ForegroundColor Red }
        "WARN"    { Write-Host $line -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $line -ForegroundColor Green }
        default   { Write-Host $line }
    }
}

function Invoke-Sql {
    param(
        [string]$Server,
        [string]$Database,
        [string]$Query,
        [switch]$Scalar
    )
    $connStr = "Server=$Server;Database=$Database;Integrated Security=True;TrustServerCertificate=True;"
    $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
    $conn.Open()
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $Query
    $cmd.CommandTimeout = 300
    if ($Scalar) {
        $result = $cmd.ExecuteScalar()
        $conn.Close()
        return $result
    } else {
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dt = New-Object System.Data.DataTable
        [void]$adapter.Fill($dt)
        $conn.Close()
        return $dt
    }
}

function Invoke-SqlNonQuery {
    param(
        [string]$Server,
        [string]$Database,
        [string]$Query
    )
    $connStr = "Server=$Server;Database=$Database;Integrated Security=True;TrustServerCertificate=True;"
    $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
    $conn.Open()
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $Query
    $cmd.CommandTimeout = 600
    $rows = $cmd.ExecuteNonQuery()
    $conn.Close()
    return $rows
}

function Migrate-Table {
    param(
        [string]$TableName,
        [string]$KeyColumn,
        [string]$ColumnList,
        [string]$InsertColumnList,
        [string]$SelectList
    )

    Write-Log "========== 开始迁移: $TableName =========="

    # 源表计数
    $srcCount = Invoke-Sql -Server $SourceServer -Database $SourceDatabase `
        -Query "SELECT COUNT(*) FROM [$TableName]" -Scalar
    Write-Log "源表 [$TableName] 记录数: $srcCount"

    if ([int]$srcCount -eq 0) {
        Write-Log "源表为空，跳过迁移" -Level "WARN"
        return
    }

    # 目标表计数
    $tgtCount = Invoke-Sql -Server $TargetServer -Database $TargetDatabase `
        -Query "SELECT COUNT(*) FROM [$TableName]" -Scalar
    Write-Log "目标表 [$TableName] 现有记录数: $tgtCount"

    if ([int]$tgtCount -ge [int]$srcCount) {
        Write-Log "目标表数据量 ($tgtCount) >= 源表 ($srcCount)，数据已迁移，跳过" -Level "WARN"
        return
    }

    # 使用 MERGE + IDENTITY_INSERT 进行幂等迁移
    $mergeSql = @"
SET IDENTITY_INSERT [$TableName] ON;

MERGE INTO [$TargetDatabase].dbo.[$TableName] AS tgt
USING (
    SELECT $SelectList
    FROM [$SourceDatabase].dbo.[$TableName]
) AS src
ON tgt.[$KeyColumn] = src.[$KeyColumn]
WHEN NOT MATCHED THEN
    INSERT ($InsertColumnList)
    VALUES ($ColumnList);

SET IDENTITY_INSERT [$TableName] OFF;
"@

    Write-Log "正在执行 MERGE 迁移..."
    $affected = Invoke-SqlNonQuery -Server $TargetServer -Database $TargetDatabase -Query $mergeSql
    Write-Log "MERGE 影响行数: $affected"

    # 迁移后校验
    $tgtAfter = Invoke-Sql -Server $TargetServer -Database $TargetDatabase `
        -Query "SELECT COUNT(*) FROM [$TableName]" -Scalar
    Write-Log "迁移后目标表 [$TableName] 记录数: $tgtAfter"

    if ([int]$tgtAfter -ge [int]$srcCount) {
        Write-Log "$TableName 迁移完成，COUNT 校验通过 ($srcCount -> $tgtAfter)" -Level "SUCCESS"
    } else {
        Write-Log "$TableName COUNT 校验失败！源=$srcCount 目标=$tgtAfter" -Level "ERROR"
        throw "迁移 $TableName 数据校验失败"
    }
}

# =============================================================================
# 主流程
# =============================================================================

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  PerfumeShop V19 最终数据迁移" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Log "迁移开始"
Write-Log "源数据库: $SourceServer / $SourceDatabase"
Write-Log "目标数据库: $TargetServer / $TargetDatabase"

$startTime = Get-Date

try {
    # ---- 1. 迁移 Users ----
    Migrate-Table `
        -TableName "Users" `
        -KeyColumn "UserID" `
        -SelectList "UserID, Username, PasswordHash, Email, Phone, FullName, Gender, Birthday, PasswordVersion, CreatedAt, UpdatedAt" `
        -InsertColumnList "[UserID], [Username], [PasswordHash], [Email], [Phone], [FullName], [Gender], [Birthday], [PasswordVersion], [CreatedAt], [UpdatedAt]" `
        -ColumnList "src.[UserID], src.[Username], src.[PasswordHash], src.[Email], src.[Phone], src.[FullName], src.[Gender], src.[Birthday], src.[PasswordVersion], src.[CreatedAt], src.[UpdatedAt]"

    # ---- 2. 迁移 Addresses ----
    Migrate-Table `
        -TableName "Addresses" `
        -KeyColumn "AddressID" `
        -SelectList "AddressID, UserID, RecipientName, Phone, Province, City, District, DetailAddress, IsDefault, CreatedAt" `
        -InsertColumnList "[AddressID], [UserID], [RecipientName], [Phone], [Province], [City], [District], [DetailAddress], [IsDefault], [CreatedAt]" `
        -ColumnList "src.[AddressID], src.[UserID], src.[RecipientName], src.[Phone], src.[Province], src.[City], src.[District], src.[DetailAddress], src.[IsDefault], src.[CreatedAt]"

    # ---- 3. 迁移 Orders ----
    Migrate-Table `
        -TableName "Orders" `
        -KeyColumn "OrderID" `
        -SelectList "OrderID, UserID, OrderNumber, TotalAmount, Status, ShippingAddress, PaymentMethod, CouponCode, CouponDiscount, PointsEarned, PointsRedeemed, PointsDiscount, CreatedAt, UpdatedAt" `
        -InsertColumnList "[OrderID], [UserID], [OrderNumber], [TotalAmount], [Status], [ShippingAddress], [PaymentMethod], [CouponCode], [CouponDiscount], [PointsEarned], [PointsRedeemed], [PointsDiscount], [CreatedAt], [UpdatedAt]" `
        -ColumnList "src.[OrderID], src.[UserID], src.[OrderNumber], src.[TotalAmount], src.[Status], src.[ShippingAddress], src.[PaymentMethod], src.[CouponCode], src.[CouponDiscount], src.[PointsEarned], src.[PointsRedeemed], src.[PointsDiscount], src.[CreatedAt], src.[UpdatedAt]"

    # ---- 4. 迁移 ProductReviews ----
    Migrate-Table `
        -TableName "ProductReviews" `
        -KeyColumn "ReviewID" `
        -SelectList "ReviewID, ProductID, UserID, Rating, Content, Title, IsVerifiedPurchase, AIFeelingSummary, LikeCount, CreatedAt" `
        -InsertColumnList "[ReviewID], [ProductID], [UserID], [Rating], [Content], [Title], [IsVerifiedPurchase], [AIFeelingSummary], [LikeCount], [CreatedAt]" `
        -ColumnList "src.[ReviewID], src.[ProductID], src.[UserID], src.[Rating], src.[Content], src.[Title], src.[IsVerifiedPurchase], src.[AIFeelingSummary], src.[LikeCount], src.[CreatedAt]"

    Write-Log "============================================"
    Write-Log "  全部 4 张表迁移完成" -Level "SUCCESS"
    Write-Log "============================================"
}
catch {
    Write-Log "迁移失败: $($_.Exception.Message)" -Level "ERROR"
    Write-Log $_.ScriptStackTrace -Level "ERROR"
    exit 1
}

$endTime  = Get-Date
$duration = ($endTime - $startTime).TotalSeconds
Write-Log "总耗时: $([math]::Round($duration, 1)) 秒"
Write-Log "日志文件: $logFile"

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  迁移完成" -ForegroundColor Green
Write-Host "  耗时: $([math]::Round($duration, 1)) 秒" -ForegroundColor Cyan
Write-Host "  日志: $logFile" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
