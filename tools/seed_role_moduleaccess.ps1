<#
  Seed AdminRole.ModuleAccess for non-super roles (idempotent).
  Module gate (PermissionService.CanAccessAsync) reads this comma-separated,
  case-insensitive list. SUPER_ADMIN bypasses in code and is left untouched.
  Re-runnable: simply sets the target value each time.
  Usage: powershell -NoProfile -ExecutionPolicy Bypass -File tools/seed_role_moduleaccess.ps1
#>
param(
    [string]$ConnStr = 'Server=localhost\YOURPERFUME;Database=PerfumeShop;Trusted_Connection=True;TrustServerCertificate=True;'
)
$ErrorActionPreference = 'Stop'

# RoleId -> ModuleAccess (system module reserved for SUPER_ADMIN only)
$map = @{
    2 = 'operation,logistics,inventory,analytics'                                      # Operations
    3 = 'prodcenter,semifinished,inventory'                                            # Production (no techcenter: full-formula isolation)
    4 = 'finance,purchase,analytics'                                                   # Finance
    5 = 'operation,finance,purchase,logistics,inventory,prodcenter,semifinished,techcenter,analytics'  # general admin (all except system)
    6 = 'operation'                                                                    # editor
}

$c = New-Object System.Data.SqlClient.SqlConnection($ConnStr)
$c.Open()
try {
    Write-Output '--- BEFORE ---'
    $b = $c.CreateCommand(); $b.CommandText = 'SELECT RoleId,RoleCode,ISNULL(ModuleAccess,'''') AS M FROM AdminRoles ORDER BY RoleId'
    $rd = $b.ExecuteReader(); while ($rd.Read()) { Write-Output ("id=" + $rd['RoleId'] + " code=" + $rd['RoleCode'] + " modules=[" + $rd['M'] + "]") }; $rd.Close()

    foreach ($id in ($map.Keys | Sort-Object)) {
        $cmd = $c.CreateCommand()
        $cmd.CommandText = 'UPDATE AdminRoles SET ModuleAccess=@m WHERE RoleId=@id AND RoleCode<>''SUPER_ADMIN'''
        [void]$cmd.Parameters.AddWithValue('@m', $map[$id])
        [void]$cmd.Parameters.AddWithValue('@id', $id)
        $n = $cmd.ExecuteNonQuery()
        Write-Output ("UPDATE id=" + $id + " rows=" + $n)
    }

    Write-Output '--- AFTER ---'
    $a = $c.CreateCommand(); $a.CommandText = 'SELECT RoleId,RoleCode,ISNULL(ModuleAccess,'''') AS M FROM AdminRoles ORDER BY RoleId'
    $rd = $a.ExecuteReader(); while ($rd.Read()) { Write-Output ("id=" + $rd['RoleId'] + " code=" + $rd['RoleCode'] + " modules=[" + $rd['M'] + "]") }; $rd.Close()
    Write-Output 'SEED: DONE'
} finally { $c.Close() }
