$s = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$r1 = Invoke-WebRequest -Uri 'http://localhost/user/login.asp' -WebSession $s -UseBasicParsing -TimeoutSec 15
$csrf = [regex]::Match($r1.Content, 'name="csrf_token"\s+value="([^"]+)"')
$token = if ($csrf.Success) { $csrf.Groups[1].Value } else { '' }
$body = @{ username = 'raymond'; password = 'raymond@2026' }
if ($token) { $body.csrf_token = $token }
$r2 = Invoke-WebRequest -Uri 'http://localhost/user/login.asp' -WebSession $s -Method POST -Body $body -UseBasicParsing -TimeoutSec 15
Write-Host ('Login: ' + $r2.StatusCode)
Start-Sleep -Milliseconds 500

try {
    $r = Invoke-WebRequest -Uri 'http://localhost/debug_iso2.asp' -WebSession $s -UseBasicParsing -TimeoutSec 15
    Write-Host ('Status: ' + $r.StatusCode)
    $c = $r.Content -replace '<[^>]+>',''
    Write-Host $c
} catch {
    Write-Host ('FAIL: ' + $_.Exception.Message)
}
