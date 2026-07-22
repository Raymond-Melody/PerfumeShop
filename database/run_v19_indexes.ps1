# V19: 执行性能索引SQL脚本
$server = "localhost\YOURPERFUME"
$database = "PerfumeShop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$sqlFile = Join-Path $scriptDir "v19_perf_indexes.sql"

$connStr = "Server=$server;Database=$database;Trusted_Connection=True;TrustServerCertificate=True;"

Write-Host "Connecting to $server/$database..."

try {
    $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
    $conn.Open()
    Write-Host "Connected. Executing v19_perf_indexes.sql..."
    
    $sql = Get-Content $sqlFile -Raw -Encoding UTF8
    $batches = $sql -split "\bGO\b"
    
    foreach ($batch in $batches) {
        $trimmed = $batch.Trim()
        if ($trimmed.Length -gt 0) {
            $cmd = $conn.CreateCommand()
            $cmd.CommandText = $trimmed
            $cmd.CommandTimeout = 60
            try {
                $reader = $cmd.ExecuteReader()
                if ($reader.HasRows) {
                    $dt = New-Object System.Data.DataTable
                    $dt.Load($reader)
                    $dt | Format-Table -AutoSize
                }
                $reader.Close()
            } catch {
                Write-Host "  SKIP: $_" -ForegroundColor Yellow
            }
            $cmd.Dispose()
        }
    }
    
    $conn.Close()
    Write-Host "V19 indexes applied successfully." -ForegroundColor Green
}
catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
}
finally {
    if ($conn.State -eq 'Open') { $conn.Close() }
}
