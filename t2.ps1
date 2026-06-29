try {
    # Login
    $s = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $r1 = Invoke-WebRequest -Uri 'http://localhost/user/login.asp' -WebSession $s -UseBasicParsing -TimeoutSec 15
    $csrf = [regex]::Match($r1.Content, 'name="csrf_token"\s+value="([^"]+)"')
    $token = if ($csrf.Success) { $csrf.Groups[1].Value } else { '' }
    $body = @{ username = 'raymond'; password = 'raymond@2026' }
    if ($token) { $body.csrf_token = $token }
    $r2 = Invoke-WebRequest -Uri 'http://localhost/user/login.asp' -WebSession $s -Method POST -Body $body -UseBasicParsing -TimeoutSec 15
    Write-Host ('Login: ' + $r2.StatusCode)

    Start-Sleep -Milliseconds 500

    # Test debug_iso.asp
    try {
        $r = Invoke-WebRequest -Uri 'http://localhost/debug_iso.asp' -WebSession $s -UseBasicParsing -TimeoutSec 15
        Write-Host ('debug_iso: ' + $r.StatusCode)
        Write-Host $r.Content
    } catch {
        Write-Host ('debug_iso FAIL: ' + $_.Exception.Message)
    }

    Start-Sleep -Milliseconds 500

    # Test product.asp
    try {
        $r = Invoke-WebRequest -Uri 'http://localhost/product.asp?id=1' -WebSession $s -UseBasicParsing -TimeoutSec 15
        Write-Host ('product.asp: ' + $r.StatusCode)
    } catch {
        Write-Host ('product.asp FAIL: ' + $_.Exception.Message)
    }
} catch {
    Write-Host ('Login FAIL: ' + $_.Exception.Message)
}
