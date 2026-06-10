$conn = New-Object -ComObject ADODB.Connection
$conn.Open("Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;")

Write-Host "=== ProductionOrders columns ==="
$rs = $conn.Execute("SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='ProductionOrders' ORDER BY ORDINAL_POSITION")
while (-not $rs.EOF) {
    Write-Host ("{0} ({1})" -f $rs.Fields(0).Value, $rs.Fields(1).Value)
    $rs.MoveNext()
}
$rs.Close()

# Add missing columns if needed
Write-Host "`nAdding missing columns..."
try { $conn.Execute("ALTER TABLE ProductionOrders ADD PlannedQty INT DEFAULT 0"); Write-Host "Added PlannedQty" } catch { Write-Host "PlannedQty already exists" }
try { $conn.Execute("ALTER TABLE ProductionOrders ADD AssignedTo NVARCHAR(100)"); Write-Host "Added AssignedTo" } catch { Write-Host "AssignedTo already exists" }
try { $conn.Execute("ALTER TABLE ProductionOrders ADD RecipeName NVARCHAR(200)"); Write-Host "Added RecipeName" } catch { Write-Host "RecipeName already exists" }
try { $conn.Execute("ALTER TABLE ProductionOrders ADD BatchNo NVARCHAR(50)"); Write-Host "Added BatchNo" } catch { Write-Host "BatchNo already exists" }

$conn.Close()
Write-Host "Done!"
