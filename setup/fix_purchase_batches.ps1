$conn = New-Object -ComObject ADODB.Connection
$conn.Open("Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;")

Write-Host "=== Checking PurchaseBatches ==="
$exists = $false
try {
    $rs = $conn.Execute("SELECT TOP 1 1 FROM PurchaseBatches")
    $rs.Close()
    $exists = $true
    Write-Host "PurchaseBatches already exists"
} catch {
    Write-Host "PurchaseBatches missing, creating..."
}

if (-not $exists) {
    $sql = @"
CREATE TABLE [PurchaseBatches] (
    [BatchID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    [PurchaseDetailID] INT NULL,
    [PurchaseID] INT NULL,
    [BatchNo] NVARCHAR(50) NULL,
    [ItemType] NVARCHAR(30) DEFAULT 'RawMaterial',
    [ItemCode] NVARCHAR(50) NULL,
    [ItemName] NVARCHAR(200) NULL,
    [UnitPrice] DECIMAL(19,4) DEFAULT 0,
    [Quantity] FLOAT DEFAULT 0,
    [ReceivedQty] FLOAT DEFAULT 0,
    [RemainingQty] FLOAT DEFAULT 0,
    [ReceivedDate] DATETIME NULL,
    [SupplierID] INT NULL,
    [CostAllocated] BIT DEFAULT 0,
    [CreatedAt] DATETIME DEFAULT GETDATE()
)
"@
    $conn.Execute($sql)
    Write-Host "PurchaseBatches created successfully"
}

Write-Host ""
Write-Host "=== Checking ShippingCompanies ==="
$scExists = $false
try {
    $rs = $conn.Execute("SELECT TOP 1 1 FROM ShippingCompanies")
    $rs.Close()
    $scExists = $true
    Write-Host "ShippingCompanies already exists"
} catch {
    Write-Host "ShippingCompanies missing, creating..."
}

if (-not $scExists) {
    $sql2 = @"
CREATE TABLE [ShippingCompanies] (
    [CompanyID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    [CompanyName] NVARCHAR(100) NOT NULL,
    [ContactPerson] NVARCHAR(50),
    [ContactPhone] NVARCHAR(20),
    [Website] NVARCHAR(200),
    [IsActive] BIT DEFAULT 1,
    [Notes] NVARCHAR(MAX),
    [CreatedAt] DATETIME DEFAULT GETDATE(),
    [UpdatedAt] DATETIME
)
"@
    $conn.Execute($sql2)
    Write-Host "ShippingCompanies created successfully"
}

Write-Host ""
Write-Host "=== Fixing AdminRoles ==="
$rs = $conn.Execute("SELECT RoleID, RoleName, RoleCode FROM AdminRoles ORDER BY RoleID")
while ($rs.EOF -eq $false) {
    $rid = $rs.Fields.Item("RoleID").Value
    $rname = $rs.Fields.Item("RoleName").Value
    $rcode = $rs.Fields.Item("RoleCode").Value
    Write-Host ("RoleID=" + $rid + " | " + $rname + " | " + $rcode)
    $rs.MoveNext()
}
$rs.Close()

$conn.Execute("UPDATE AdminRoles SET RoleName = N'Finance Admin', RoleCode = 'FINANCE' WHERE RoleID = 4")
Write-Host "Updated RoleID=4 to Finance Admin / FINANCE"

Write-Host ""
Write-Host "=== Verify ==="
$rs = $conn.Execute("SELECT RoleID, RoleName, RoleCode FROM AdminRoles ORDER BY RoleID")
while ($rs.EOF -eq $false) {
    $rid = $rs.Fields.Item("RoleID").Value
    $rname = $rs.Fields.Item("RoleName").Value
    $rcode = $rs.Fields.Item("RoleCode").Value
    Write-Host ("AFTER: RoleID=" + $rid + " | " + $rname + " | " + $rcode)
    $rs.MoveNext()
}
$rs.Close()

$conn.Close()
Write-Host ""
Write-Host "=== All fixes applied successfully ==="
