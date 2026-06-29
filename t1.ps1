try {
    $r = Invoke-WebRequest -Uri 'http://localhost/debug_iso.asp' -UseBasicParsing -TimeoutSec 10
    Write-Host ('OK: ' + $r.StatusCode)
    Write-Host $r.Content
} catch {
    Write-Host ('FAIL: ' + $_.Exception.Message)
}
