$conn = New-Object -ComObject ADODB.Connection
$conn.Open("Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;")

Write-Host "=== Before fix ==="
$rs = $conn.Execute("SELECT ProductType, ReviewStatus, COUNT(*) as cnt FROM Products GROUP BY ProductType, ReviewStatus")
while (-not $rs.EOF) {
    Write-Host ("Type={0}, Review={1}, Count={2}" -f $rs.Fields(0).Value, $rs.Fields(1).Value, $rs.Fields(2).Value)
    $rs.MoveNext()
}
$rs.Close()

# Fix 1: Rename Fixed -> standard
Write-Host "`nFixing ProductType 'Fixed' -> 'standard'..."
$conn.Execute("UPDATE Products SET ProductType='standard' WHERE ProductType='Fixed'")
Write-Host "Done."

# Fix 2: Approve pending Custom products (they need review)
# Actually, let me check if standard type requires review...
$rs = $conn.Execute("SELECT TypeCode, RequiresReview FROM ProductTypeConfig WHERE TypeCode='standard'")
if (-not $rs.EOF) {
    $reqReview = $rs.Fields(1).Value
    Write-Host "standard RequiresReview: $reqReview"
    if ($reqReview) {
        Write-Host "Approving standard products..."
        $conn.Execute("UPDATE Products SET ReviewStatus='Approved' WHERE ProductType='standard' AND ReviewStatus='Pending'")
        Write-Host "Done."
    }
}
$rs.Close()

# Fix 3: Approve Custom products
$rs = $conn.Execute("SELECT TypeCode, RequiresReview FROM ProductTypeConfig WHERE TypeCode='custom'")
if (-not $rs.EOF) {
    $reqReview = $rs.Fields(1).Value
    Write-Host "custom RequiresReview: $reqReview"
    if ($reqReview) {
        Write-Host "Approving custom products..."
        $conn.Execute("UPDATE Products SET ReviewStatus='Approved' WHERE ProductType='custom' AND ReviewStatus='Pending'")
        Write-Host "Done."
    }
}
$rs.Close()

Write-Host "`n=== After fix ==="
$rs = $conn.Execute("SELECT ProductType, ReviewStatus, COUNT(*) as cnt FROM Products GROUP BY ProductType, ReviewStatus")
while (-not $rs.EOF) {
    Write-Host ("Type={0}, Review={1}, Count={2}" -f $rs.Fields(0).Value, $rs.Fields(1).Value, $rs.Fields(2).Value)
    $rs.MoveNext()
}
$rs.Close()

$conn.Close()
Write-Host "`nAll fixes applied!"
