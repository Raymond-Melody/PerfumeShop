$f = 'f:\web_index.html'
$b = [System.IO.File]::ReadAllBytes($f)

# Build hex string of full file (it's ~24KB)
$hex = ''
$maxCheck = [Math]::Min($b.Length, 50000)
for ($i = 0; $i -lt $maxCheck; $i++) {
    $hex += '{0:X2}' -f $b[$i]
}

Write-Host "File size: $($b.Length) bytes"
Write-Host "Hex size: $($hex.Length) chars"

# Search for fa-home
$idx = $hex.IndexOf('66612D686F6D65')
Write-Host "fa-home hex index: $idx"

# Also search for "666173" (fas) 
$idxFas = $hex.IndexOf('666173')
Write-Host "fas hex index: $idxFas"

# Search for common nav patterns
$patterns = @(
    @('6E6176', 'nav'),
    @('6D61696E2D6E6176', 'main-nav'),
    @('6E61762D6C697374', 'nav-list'),
    @('6661', 'fa-'),
    @('696E646578', 'index'),
    @('686F6D65', 'home'),
    @('C3A9', 'e-acute (garbled marker)')
)

foreach ($p in $patterns) {
    $pos = $hex.IndexOf($p[0])
    if ($pos -ge 0) {
        Write-Host "FOUND '$($p[1])' at hex pos $pos"
    } else {
        Write-Host "MISSING '$($p[1])'"
    }
}

# If fa-home not found, search for nav structure
Write-Host ""
Write-Host "=== NAV section analysis ==="

# Look for "main-nav" area
$mainNav = $hex.IndexOf('6D61696E2D6E6176')
if ($mainNav -ge 0) {
    $navStart = [int]($mainNav / 2)
    Write-Host "Nav section at byte $navStart, content:"
    for ($i = $navStart; $i -lt [Math]::Min($navStart + 300, $b.Length); $i++) {
        $c = $b[$i]
        if ($c -ge 0x20 -and $c -le 0x7E) { 
            Write-Host -NoNewline ([char]$c)
        } elseif ($c -eq 0x0A) {
            Write-Host ""
        }
    }
    Write-Host ""
}
