$conn = New-Object -ComObject ADODB.Connection
$conn.Open("Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;")

function Test-ColumnExists($table, $column) {
    try {
        $rs = $conn.Execute("SELECT TOP 1 [$column] FROM [$table]")
        $rs.Close()
        return $true
    } catch {
        return $false
    }
}

function Test-TableExists($table) {
    try {
        $rs = $conn.Execute("SELECT TOP 1 1 FROM [$table]")
        $rs.Close()
        return $true
    } catch {
        return $false
    }
}

# 1. PurchaseOrders.OrderType
Write-Host "1. PurchaseOrders.OrderType..."
if (Test-ColumnExists "PurchaseOrders" "OrderType") {
    Write-Host "   Already exists"
} else {
    $conn.Execute("ALTER TABLE PurchaseOrders ADD OrderType NVARCHAR(20) DEFAULT 'RawMaterial'")
    Write-Host "   Added"
    $conn.Execute("UPDATE PurchaseOrders SET OrderType = 'RawMaterial' WHERE OrderType IS NULL")
    Write-Host "   Default values set"
}

# 2. PurchaseOrders.ExpectedDeliveryDate
Write-Host "2. PurchaseOrders.ExpectedDeliveryDate..."
if (Test-ColumnExists "PurchaseOrders" "ExpectedDeliveryDate") {
    Write-Host "   Already exists"
} else {
    $conn.Execute("ALTER TABLE PurchaseOrders ADD ExpectedDeliveryDate DATETIME2(7)")
    Write-Host "   Added"
}

# 3. SupplierPrices.PriceType
Write-Host "3. SupplierPrices.PriceType..."
if (Test-ColumnExists "SupplierPrices" "PriceType") {
    Write-Host "   Already exists"
} else {
    $conn.Execute("ALTER TABLE SupplierPrices ADD PriceType NVARCHAR(30) DEFAULT 'Standard'")
    Write-Host "   Added"
}

# 4. SupplierPrices.Unit
Write-Host "4. SupplierPrices.Unit..."
if (Test-ColumnExists "SupplierPrices" "Unit") {
    Write-Host "   Already exists"
} else {
    $conn.Execute("ALTER TABLE SupplierPrices ADD Unit NVARCHAR(20) DEFAULT 'kg'")
    Write-Host "   Added"
}

# 5. SupplierContracts table
Write-Host "5. SupplierContracts table..."
if (Test-TableExists "SupplierContracts") {
    Write-Host "   Already exists"
} else {
    $conn.Execute(@"
CREATE TABLE [SupplierContracts] (
    [ContractID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    [SupplierID] INT NOT NULL,
    [ContractNo] NVARCHAR(50),
    [ContractName] NVARCHAR(200),
    [ContractType] NVARCHAR(30) DEFAULT 'Supply',
    [StartDate] DATETIME2(7),
    [EndDate] DATETIME2(7),
    [TotalAmount] DECIMAL(19,4),
    [PaymentTerms] NVARCHAR(200),
    [TermsSummary] NVARCHAR(MAX),
    [AttachmentURL] NVARCHAR(500),
    [Status] NVARCHAR(20) DEFAULT 'Active',
    [SignedAt] DATETIME2(7),
    [CreatedAt] DATETIME2(7) DEFAULT GETDATE(),
    [UpdatedAt] DATETIME2(7)
)
"@)
    Write-Host "   Created"
}

# 6. SupplierEvaluations table
Write-Host "6. SupplierEvaluations table..."
if (Test-TableExists "SupplierEvaluations") {
    Write-Host "   Already exists"
} else {
    $conn.Execute(@"
CREATE TABLE [SupplierEvaluations] (
    [EvaluationID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    [SupplierID] INT NOT NULL,
    [EvaluatedBy] NVARCHAR(50),
    [EvaluationDate] DATETIME2(7) DEFAULT GETDATE(),
    [QualityScore] INT DEFAULT 0,
    [DeliveryScore] INT DEFAULT 0,
    [PriceScore] INT DEFAULT 0,
    [ServiceScore] INT DEFAULT 0,
    [OverallScore] INT DEFAULT 0,
    [Rating] NVARCHAR(10) DEFAULT 'C',
    [Comments] NVARCHAR(MAX),
    [Recommendations] NVARCHAR(MAX),
    [Period] NVARCHAR(20),
    [CreatedAt] DATETIME2(7) DEFAULT GETDATE()
)
"@)
    Write-Host "   Created"
}

$conn.Close()
Write-Host ""
Write-Host "=== V8 Iter2 purchase upgrade complete ==="
