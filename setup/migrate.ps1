# PowerShell Migration Script: Access -> SQL Server
$ErrorActionPreference = "Continue"

$accessPath = Join-Path (Get-Location) "database\PerfumeShop.mdb"
$sqlServer = "localhost\YOURPERFUME"
$database = "PerfumeShop"

# Migration order - base tables first
$migrationOrder = @(
    "Users", "AdminRoles", "AdminUsers", "Categories", "ProductTypeConfig",
    "Volumes", "FragranceNotes", "BaseNotes", "BottleStyles", "Ingredients",
    "Products", "ProductVolumePrices", "ProductNotes", "ProductNoteRatios",
    "ProductBottleStyles", "FragranceIngredients", "FormulaNotes", "Formulas",
    "NoteIngredients", "NoteInventory", "ProductInventory",
    "Recipes", "RecipeNotes", "RecipeIngredients", "RecipeProducts", "RecipeProductNotes",
    "RecipeAccords", "RecipeAccordMaterials", "RecipePublishLog", "RecommendedRecipes",
    "RecipePopularity",
    "Suppliers", "PurchaseCategories", "SupplierPrices",
    "PurchaseOrders", "PurchaseOrderDetails", "PurchaseReceipts", "PurchaseReceiptDetails",
    "PurchaseCostReview", "RawMaterialInventory",
    "MaterialOutbound", "MaterialOutboundDetails",
    "ProductionOrders", "ProductionLogs", "ProductManufacturing", "ProductManufacturingDetails",
    "AccordProductions", "AccordProductionDetails", "AccordQCReports",
    "WorkshopTransfer",
    "Orders", "OrderDetails", "OrderDetailNoteSelections", "OrderIngredients",
    "Cart", "CartNoteSelections",
    "UserAddresses", "UserFavorites", "UserPoints", "PointTransactions",
    "UserPreferences", "ProductReviews",
    "PaymentRecords", "RefundRecords", "ReconciliationLogs",
    "ExpenseRecords", "BudgetPlans", "FundAccounts", "ProductCosts",
    "Coupons", "MarketingCampaigns",
    "ModulePermissions", "AdminLogs",
    "DailyStatistics", "SiteSettings"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Data Migration: Access -> SQL Server" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Connect to Access
Write-Host "[INFO] Connecting to Access..." -ForegroundColor Gray
$connAcc = New-Object -ComObject ADODB.Connection
try {
    $connAcc.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$accessPath;")
    Write-Host "[OK] Access connected" -ForegroundColor Green
} catch {
    try {
        $connAcc.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$accessPath;")
        Write-Host "[OK] Access connected (Jet)" -ForegroundColor Green
    } catch {
        Write-Host "[FATAL] Cannot connect Access: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

# Connect to SQL Server
Write-Host "[INFO] Connecting to SQL Server..." -ForegroundColor Gray
$connSQL = New-Object -ComObject ADODB.Connection
try {
    $connSQL.Open("Provider=SQLOLEDB;Server=$sqlServer;Database=$database;Integrated Security=SSPI;")
    Write-Host "[OK] SQL Server connected" -ForegroundColor Green
} catch {
    Write-Host "[FATAL] Cannot connect SQL Server: $($_.Exception.Message)" -ForegroundColor Red
    $connAcc.Close()
    exit 1
}

# Get SQL tables
$rsSQLTables = $connSQL.Execute("SELECT name FROM sys.tables ORDER BY name")
$sqlTableSet = @{}
while (-not $rsSQLTables.EOF) {
    $sqlTableSet[$rsSQLTables.Fields(0).Value] = $true
    $rsSQLTables.MoveNext()
}
$rsSQLTables.Close()

# Get Access tables
$rsAccSchema = $connAcc.OpenSchema(20)  # adSchemaTables
$accTableSet = @{}
while (-not $rsAccSchema.EOF) {
    $tableName = $rsAccSchema.Fields("TABLE_NAME").Value
    if ($rsAccSchema.Fields("TABLE_TYPE").Value -eq "TABLE" -and 
        -not $tableName.StartsWith("MSys") -and -not $tableName.StartsWith("~")) {
        $accTableSet[$tableName] = $true
    }
    $rsAccSchema.MoveNext()
}
$rsAccSchema.Close()

Write-Host "[INFO] Access tables: $($accTableSet.Count) | SQL tables: $($sqlTableSet.Count)" -ForegroundColor Gray

$totalRows = 0
$totalSuccess = 0
$totalSkipped = 0

foreach ($tableName in $migrationOrder) {
    if (-not $sqlTableSet.ContainsKey($tableName)) {
        Write-Host "[SKIP] [$tableName] - not in SQL Server" -ForegroundColor Yellow
        continue
    }
    if (-not $accTableSet.ContainsKey($tableName)) {
        Write-Host "[SKIP] [$tableName] - not in Access" -ForegroundColor Yellow
        continue
    }
    
    # Check if SQL table already has data
    $rsCount = $connSQL.Execute("SELECT COUNT(*) FROM [$tableName]")
    $existingCount = $rsCount.Fields(0).Value
    $rsCount.Close()
    
    if ($existingCount -gt 0) {
        Write-Host "[SKIP] [$tableName] - already has $existingCount rows" -ForegroundColor Gray
        $totalSkipped++
        continue
    }
    
    # Read Access data
    $rsAcc = $connAcc.Execute("SELECT * FROM [$tableName]")
    if ($rsAcc.EOF) {
        Write-Host "[SKIP] [$tableName] - no data in Access" -ForegroundColor Gray
        $rsAcc.Close()
        continue
    }
    
    # Get column names
    $colCount = $rsAcc.Fields.Count - 1
    $accCols = @()
    for ($j = 0; $j -le $colCount; $j++) {
        $accCols += $rsAcc.Fields($j).Name
    }
    
    # Try IDENTITY_INSERT
    $hasIdentity = $true
    try {
        $connSQL.Execute("SET IDENTITY_INSERT [$tableName] ON")
    } catch {
        $hasIdentity = $false
    }
    
    $startCol = 0
    if (-not $hasIdentity) {
        $startCol = 1
    }
    
    $rowCount = 0
    $errorCount = 0
    
    while (-not $rsAcc.EOF) {
        $colList = @()
        $valList = @()
        
        for ($j = $startCol; $j -le $colCount; $j++) {
            $colList += "[$($accCols[$j])]"
            $val = $rsAcc.Fields($j).Value
            
            if ($val -eq [DBNull]::Value -or $val -eq $null) {
                $valList += "NULL"
            } else {
                $type = $val.GetType().Name
                switch ($type) {
                    "Int16" { $valList += $val.ToString() }
                    "Int32" { $valList += $val.ToString() }
                    "Int64" { $valList += $val.ToString() }
                    "Double" { $valList += $val.ToString().Replace(",", ".") }
                    "Decimal" { $valList += $val.ToString().Replace(",", ".") }
                    "Single" { $valList += $val.ToString().Replace(",", ".") }
                    "Boolean" { $valList += if($val) { "1" } else { "0" } }
                    "DateTime" { $valList += "'" + $val.ToString("yyyy-MM-dd HH:mm:ss") + "'" }
                    default {
                        $escaped = $val.ToString().Replace("'", "''")
                        $valList += "N'$escaped'"
                    }
                }
            }
        }
        
        $colStr = $colList -join ", "
        $valStr = $valList -join ", "
        $insertSQL = "INSERT INTO [$tableName] ($colStr) VALUES ($valStr)"
        
        try {
            $connSQL.Execute($insertSQL)
            $rowCount++
        } catch {
            $errorCount++
            if ($errorCount -le 3) {
                Write-Host "[WARN] [$tableName] row insert error: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        
        $rsAcc.MoveNext()
    }
    
    if ($hasIdentity) {
        try { $connSQL.Execute("SET IDENTITY_INSERT [$tableName] OFF") } catch { }
    }
    
    $totalRows += $rowCount
    $totalSuccess++
    
    $color = if($errorCount -eq 0) { "Green" } else { "Yellow" }
    $msg = "[OK] [$tableName]: $rowCount rows"
    if ($errorCount -gt 0) { $msg += " ($errorCount errors)" }
    Write-Host $msg -ForegroundColor $color
    
    $rsAcc.Close()
}

$connAcc.Close()
$connSQL.Close()

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Migration Complete!" -ForegroundColor Green
Write-Host "Tables migrated: $totalSuccess | Skipped: $totalSkipped | Total rows: $totalRows" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
