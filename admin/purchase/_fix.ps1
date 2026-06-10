$file = Join-Path $PSScriptRoot "purchase_orders.asp"
$bytes = [System.IO.File]::ReadAllBytes($file)
$content = [System.Text.Encoding]::UTF8.GetString($bytes)

# Fix 1: Fix garbled confirm message using context
$idx = $content.IndexOf("if (!confirm('")
if ($idx -ge 0) {
    $endIdx = $content.IndexOf("')) return;", $idx)
    if ($endIdx -ge 0) {
        $before = $content.Substring(0, $idx)
        $after = $content.Substring($endIdx + 12)
        $correct1 = [char]0x786e + [char]0x5b9a + [char]0x8981 + [char]0x590d + [char]0x5236 + [char]0x6b64 + [char]0x8ba2 + [char]0x5355 + [char]0x5417 + [char]0xff1f + [char]0x5c06 + [char]0x521b + [char]0x5efa + [char]0x4e00 + [char]0x4e2a + [char]0x65b0 + [char]0x7684 + [char]0x8349 + [char]0x7a3f + [char]0x8ba2 + [char]0x5355 + [char]0x3002
        $content = $before + "if (!confirm('" + $correct1 + "')) return;" + $after
        Write-Host "Fix 1 applied via context search"
    }
} else {
    Write-Host "Fix 1: context not found"
}

# Fix 2: Fix garbled button text using context
$idx = $content.IndexOf('<i class="fas fa-copy"></i> ')
if ($idx -ge 0) {
    $endIdx = $content.IndexOf("</button>", $idx)
    if ($endIdx -ge 0) {
        $before = $content.Substring(0, $idx + 32)
        $after = $content.Substring($endIdx)
        $correct2 = [char]0x590d + [char]0x5236 + [char]0x8ba2 + [char]0x5355
        $content = $before + $correct2 + [Environment]::NewLine + "                    " + $after
        Write-Host "Fix 2 applied via context search"
    }
} else {
    Write-Host "Fix 2: context not found"
}

[System.IO.File]::WriteAllText($file, $content, [System.Text.UTF8Encoding]::new($true))
Write-Host "Fixes complete"
