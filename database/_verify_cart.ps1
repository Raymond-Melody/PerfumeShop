$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'
$base = 'http://localhost:5000'
$s = New-Object Microsoft.PowerShell.Commands.WebRequestSession

# 登录
$g = Invoke-WebRequest -Uri "$base/login" -WebSession $s -UseBasicParsing -TimeoutSec 20
$token = ''
if ($g.Content -match 'name="__RequestVerificationToken"[^>]*value="([^"]+)"') { $token = $Matches[1] }
$body = @{ email='raymond@example.com'; password='Ray123456'; '__RequestVerificationToken'=$token }
try { Invoke-WebRequest -Uri "$base/login" -Method POST -Body $body -WebSession $s -UseBasicParsing -TimeoutSec 20 -MaximumRedirection 0 -ErrorAction Stop | Out-Null } catch {}
$auth = ($s.Cookies.GetCookies($base) | Where-Object { $_.Name -eq 'V19_AUTH' })
Write-Output ("LOGIN auth cookie=" + ($auth -ne $null))

# 先清空 Raymond 旧购物车（避免干扰），再加购三件产品
foreach ($prodId in @(50,58,56)) {
  try {
    $r = Invoke-WebRequest -Uri "$base/cart?add=$prodId" -WebSession $s -UseBasicParsing -TimeoutSec 20 -ErrorAction Stop
    Write-Output ("add=$prodId => " + $r.StatusCode)
  } catch { Write-Output ("add=$prodId => ERR " + $_.Exception.Message) }
}

# 查看购物车内容
$cart = Invoke-WebRequest -Uri "$base/cart" -WebSession $s -UseBasicParsing -TimeoutSec 20
$names = [regex]::Matches($cart.Content, '<span>([^<]+)</span>') | ForEach-Object { $_.Groups[1].Value }
Write-Output "=== cart page spans (product names) ==="
$names | Where-Object { $_ -match '香|定制|联名' } | Select-Object -Unique | ForEach-Object { Write-Output ("  " + $_) }
if ($cart.Content -match '合计:\s*<span[^>]*>¥([\d\.]+)') { Write-Output ("合计=" + $Matches[1]) }
elseif ($cart.Content -match '¥([\d\.]+)</span></h4>') { Write-Output ("合计(alt)=" + $Matches[1]) }
