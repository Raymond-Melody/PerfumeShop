$conn = New-Object -ComObject ADODB.Connection
$conn.Open("Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;")

# Fix RoleCode for Super Admin - must be uppercase "SUPER_ADMIN"
$conn.Execute("UPDATE AdminRoles SET RoleCode='SUPER_ADMIN' WHERE RoleID=1")
Write-Host "Updated RoleID=1 RoleCode to SUPER_ADMIN"

# Verify
$rs = $conn.Execute("SELECT RoleID, RoleName, RoleCode FROM AdminRoles")
while (-not $rs.EOF) {
    Write-Host ("RoleID={0}, Name={1}, Code={2}" -f $rs.Fields(0).Value, $rs.Fields(1).Value, $rs.Fields(2).Value)
    $rs.MoveNext()
}
$rs.Close()
$conn.Close()
Write-Host "Done!"
