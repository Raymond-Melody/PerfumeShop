<#
  Consolidate the universal page-header trio (.page-header / .page-title / .breadcrumb)
  out of per-page <style> blocks. These are byte-identical across pages and now defined
  globally in admin-v19.css, so removal is visually identical.
  Dry-run by default (prints per-file counts + DISTINCT matched CSS for verification).
  Pass -Apply to write changes (UTF-8 no BOM).
#>
param([switch]$Apply)
$ErrorActionPreference = 'Stop'
$pagesDir = Join-Path $PSScriptRoot '..\src\PerfumeShop.Admin\Components\Pages'
# Match a full CSS rule by selector; [^{}]* stays within one rule block (CSS has no nested braces).
$patterns = @(
  '\.page-header\s*\{[^{}]*\}',
  '\.page-title\s*\{[^{}]*\}',
  '\.breadcrumb\s*\{[^{}]*\}'
)
$files = Get-ChildItem -Path $pagesDir -Recurse -Filter *.razor
$distinct = New-Object System.Collections.Generic.HashSet[string]
$totalFiles = 0; $totalRules = 0
foreach ($f in $files) {
    $orig = [System.IO.File]::ReadAllText($f.FullName)
    $new = $orig; $cnt = 0
    foreach ($p in $patterns) {
        foreach ($m in [regex]::Matches($new, $p)) { [void]$distinct.Add($m.Value.Trim()); $cnt++ }
        $new = [regex]::Replace($new, $p, '')
    }
    if ($cnt -gt 0) {
        $totalFiles++; $totalRules += $cnt
        Write-Output ("  " + $f.Directory.Name + "/" + $f.Name + " : removed " + $cnt + " rule(s)")
        if ($Apply) {
            [System.IO.File]::WriteAllText($f.FullName, $new, (New-Object System.Text.UTF8Encoding($false)))
        }
    }
}
Write-Output ("---- files affected: " + $totalFiles + "  rules removed: " + $totalRules + "  (Apply=" + $Apply + ") ----")
Write-Output "==== DISTINCT matched CSS (verify these are ONLY the header trio) ===="
foreach ($d in ($distinct | Sort-Object)) { Write-Output ("   [" + $d + "]") }
