$s = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$r1 = Invoke-WebRequest -Uri 'http://localhost/user/login.asp' -WebSession $s -UseBasicParsing -TimeoutSec 15
$csrf = [regex]::Match($r1.Content, 'name="csrf_token"\s+value="([^"]+)"')
$token = if ($csrf.Success) { $csrf.Groups[1].Value } else { '' }
$body = @{ username = 'raymond'; password = 'raymond@2026' }
if ($token) { $body.csrf_token = $token }
$r2 = Invoke-WebRequest -Uri 'http://localhost/user/login.asp' -WebSession $s -Method POST -Body $body -UseBasicParsing -TimeoutSec 15
"Login status: " + $r2.StatusCode
"Cookies: " + $s.Cookies.Count
foreach ($c in $s.Cookies.GetCookies('http://localhost/')) {
    "  Cookie: " + $c.Name + "=" + $c.Value
}
"Response body preview:"
$r2.Content.Substring(0, [Math]::Min(300, $r2.Content.Length))

Start-Sleep -Milliseconds 500

try {
    $r3 = Invoke-WebRequest -Uri 'http://localhost/debug_iso.asp' -WebSession $s -UseBasicParsing -TimeoutSec 15
    "debug_iso status: " + $r3.StatusCode
    $r3.Content -replace '<[^>]+>',''
} catch {
    "debug_iso FAIL: " + $_.Exception.Message
}
