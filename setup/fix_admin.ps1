$conn = New-Object -ComObject ADODB.Connection
$conn.Open("Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;")

# Fix AdminRoles - clean insert
Write-Host "=== Fixing AdminRoles ==="
$conn.Execute("DELETE FROM AdminRoles")
$conn.Execute("DBCC CHECKIDENT ('AdminRoles', RESEED, 0)")

$conn.Execute("SET IDENTITY_INSERT AdminRoles ON")

# Use plain ASCII to avoid encoding issues
$sql = @"
INSERT INTO AdminRoles (RoleID, RoleName, RoleCode, Permissions) VALUES (1, 'Super Admin', 'super_admin', 'all')
"@
$conn.Execute($sql)
Write-Host "Inserted RoleID=1"

$sql = @"
INSERT INTO AdminRoles (RoleID, RoleName, RoleCode, Permissions) VALUES (2, 'Operations', 'operation', 'operation,orders,customers')
"@
$conn.Execute($sql)
Write-Host "Inserted RoleID=2"

$sql = @"
INSERT INTO AdminRoles (RoleID, RoleName, RoleCode, Permissions) VALUES (3, 'Production', 'production', 'production,purchase,logistics')
"@
$conn.Execute($sql)
Write-Host "Inserted RoleID=3"

$sql = @"
INSERT INTO AdminRoles (RoleID, RoleName, RoleCode, Permissions) VALUES (4, 'Read Only', 'readonly', 'read')
"@
$conn.Execute($sql)
Write-Host "Inserted RoleID=4"

$conn.Execute("SET IDENTITY_INSERT AdminRoles OFF")

# Fix AdminUsers - create Ray88
Write-Host "`n=== Fixing AdminUsers ==="
$conn.Execute("DELETE FROM AdminUsers WHERE Username='Ray88'")

# Plain ASCII
$conn.Execute("INSERT INTO AdminUsers (Username, PasswordHash, Email, RoleID, IsActive) VALUES ('Ray88', '526179323334353621', 'ray88@perfumeshop.com', 1, 1)")
Write-Host "Created Ray88"

# Also make sure admin has correct V1 hash
$conn.Execute("UPDATE AdminUsers SET PasswordHash='61646d696e31323321', RoleID=1 WHERE Username='admin'")
Write-Host "Updated admin password"

# Verify
Write-Host "`n=== AdminUsers ==="
$rs = $conn.Execute("SELECT AdminID, Username, Email, RoleID FROM AdminUsers")
while (-not $rs.EOF) {
    Write-Host ("ID={0}, User={1}, Email={2}, Role={3}" -f $rs.Fields(0).Value, $rs.Fields(1).Value, $rs.Fields(2).Value, $rs.Fields(3).Value)
    $rs.MoveNext()
}
$rs.Close()

Write-Host "`n=== AdminRoles ==="
$rs = $conn.Execute("SELECT RoleID, RoleName, RoleCode, Permissions FROM AdminRoles")
while (-not $rs.EOF) {
    Write-Host ("RoleID={0}, Name={1}, Code={2}, Perms={3}" -f $rs.Fields(0).Value, $rs.Fields(1).Value, $rs.Fields(2).Value, $rs.Fields(3).Value)
    $rs.MoveNext()
}
$rs.Close()

$conn.Close()
Write-Host "`nDone!"
