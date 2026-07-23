<#
  V19 Admin health check (parameterized, no hardcoded secrets).
  - Verifies /login serves 200 (catches "port up but site 500").
  - Logs in via form POST, then sweeps ALL @page routes and asserts HTTP 200
    (allowing known intentional redirects).
  - -Deep also checks the AdminAuditLog sink is reachable.

  Usage:
    $env:V19_ADMIN_PWD = '<password>'; powershell -NoProfile -ExecutionPolicy Bypass -File tools/smoke_healthcheck.ps1
    powershell -NoProfile -ExecutionPolicy Bypass -File tools/smoke_healthcheck.ps1 -Password '<pwd>' -Deep

  Exit codes: 0 = PASS, 1 = FAIL (broken route/login), 2 = usage error.
#>
param(
    [string]$BaseUrl  = 'http://localhost:5207',
    [string]$Username = 'Ray88',
    [string]$Password = $env:V19_ADMIN_PWD,
    [switch]$Deep,
    [string]$ConnStr  = 'Server=localhost\YOURPERFUME;Database=PerfumeShop;Trusted_Connection=True;TrustServerCertificate=True;'
)
$ErrorActionPreference = 'SilentlyContinue'
$pagesDir = Join-Path $PSScriptRoot '..\src\PerfumeShop.Admin\Components\Pages'
# Routes that intentionally 302 (redirect/alias) - not failures
$expectedRedirect = @('/login', '/admin/Operation/OrderReviews', '/admin/Operation/BaseNotes')
$fail = 0

if ([string]::IsNullOrEmpty($Password)) {
    Write-Output 'ERROR: password not provided. Set $env:V19_ADMIN_PWD or pass -Password.'
    exit 2
}

function Get-Status([string]$url, $session) {
    try {
        $r = Invoke-WebRequest -UseBasicParsing -Uri $url -WebSession $session -TimeoutSec 30 -MaximumRedirection 0
        return [int]$r.StatusCode
    } catch {
        $resp = $_.Exception.Response
        if ($resp) { return [int]$resp.StatusCode }
        return -1
    }
}

# 1) Anonymous: login page must be 200 (catches global 500)
try { $lp = Invoke-WebRequest -UseBasicParsing "$BaseUrl/login" -SessionVariable ses -TimeoutSec 30 }
catch { Write-Output ("FAIL: /login unreachable: " + $_.Exception.Message); exit 1 }
if ([int]$lp.StatusCode -ne 200) { Write-Output ("FAIL: /login status " + [int]$lp.StatusCode); exit 1 }
Write-Output 'OK: /login serves 200'

# 2) Login via form POST (hidden fields + credentials)
$body = @{}
foreach ($m in [regex]::Matches($lp.Content, '<input[^>]*type="hidden"[^>]*>')) {
    $n = [regex]::Match($m.Value, 'name="([^"]+)"').Groups[1].Value
    $v = [regex]::Match($m.Value, 'value="([^"]*)"').Groups[1].Value
    if ($n) { $body[$n] = $v }
}
$body['username'] = $Username
$body['password'] = $Password
$post = Invoke-WebRequest -UseBasicParsing "$BaseUrl/login" -Method POST -Body $body -WebSession $ses -TimeoutSec 30 -MaximumRedirection 5
if ($post.Content -match 'name="password"') { Write-Output 'FAIL: login rejected (still on login page)'; exit 1 }
Write-Output ('OK: login succeeded as ' + $Username)

# 3) Route sweep
$routes = New-Object System.Collections.Generic.List[string]
Get-ChildItem -Path $pagesDir -Recurse -Filter *.razor | ForEach-Object {
    foreach ($line in (Select-String -Path $_.FullName -Pattern '@page\s+"(/[^"]+)"' -AllMatches)) {
        foreach ($mm in $line.Matches) {
            $rt = [regex]::Replace($mm.Groups[1].Value, '\{[^}]+\}', '1')
            if (-not $routes.Contains($rt)) { $routes.Add($rt) }
        }
    }
}
$bad = New-Object System.Collections.Generic.List[string]
foreach ($rt in ($routes | Sort-Object)) {
    $code = Get-Status ("$BaseUrl" + $rt) $ses
    if ($code -eq 200) { continue }
    if (($expectedRedirect -contains $rt) -and ($code -eq 302)) { continue }
    $bad.Add("$code $rt")
}
Write-Output ("ROUTES: total=" + $routes.Count + " bad=" + $bad.Count)
foreach ($b in $bad) { Write-Output ("  BAD " + $b); $fail = 1 }

# 4) Deep: audit sink reachability
if ($Deep) {
    $c = New-Object System.Data.SqlClient.SqlConnection($ConnStr)
    try {
        $c.Open(); $m = $c.CreateCommand()
        $m.CommandText = 'SELECT COUNT(*), ISNULL(MAX(LogId),0) FROM AdminAuditLog'
        $rd = $m.ExecuteReader()
        if ($rd.Read()) { Write-Output ("AUDIT: reachable rows=" + $rd[0] + " maxLogId=" + $rd[1]) }
        $rd.Close()
    } catch { Write-Output ("FAIL: audit DB check: " + $_.Exception.Message); $fail = 1 }
    finally { $c.Close() }
}

if ($fail -eq 0) { Write-Output 'HEALTHCHECK: PASS'; exit 0 } else { Write-Output 'HEALTHCHECK: FAIL'; exit 1 }
