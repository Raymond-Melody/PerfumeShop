$wc = New-Object System.Net.WebClient
$wc.Encoding = [System.Text.Encoding]::UTF8

# First get products page snippet
$result = $wc.DownloadString("http://localhost/products.asp")
Write-Host "=== Products page (first 2000 chars) ==="
$bodyStart = $result.IndexOf("<body")
if ($bodyStart -gt 0) {
    Write-Host $result.Substring($bodyStart, [Math]::Min(2000, $result.Length - $bodyStart))
}

Write-Host ""
Write-Host "=== POSTing to migration ==="
$wc2 = New-Object System.Net.WebClient
$wc2.Encoding = [System.Text.Encoding]::UTF8
$wc2.Headers.Add("Content-Type", "application/x-www-form-urlencoded")
try {
    $responseText = $wc2.UploadString("http://localhost/setup/data_migrate.asp", "action=migrate")
    Write-Host "Response length: $($responseText.Length)"
    
    # Show all step messages
    $lines = $responseText -split '<'
    foreach ($line in $lines) {
        if ($line -match '^div class=.step (success|error|warning|info)') {
            # Extract content between > and <
            if ($line -match '>(.+)$') {
                $content = $matches[1]
                # Remove any remaining HTML
                $content = $content -replace '<[^>]+>', ''
                Write-Host $content.Trim()
            }
        }
    }
    
    if ($responseText.Length -lt 3000) {
        Write-Host ""
        Write-Host "Full response (short):"
        Write-Host $responseText
    }
} catch {
    Write-Host "POST Error: $($_.Exception.Message)"
    Write-Host "Stack: $($_.ScriptStackTrace)"
}

Write-Host ""
Write-Host "=== Checking products page again ==="
$result2 = $wc.DownloadString("http://localhost/products.asp")
Write-Host "Page length before: $($result.Length), after: $($result2.Length)"
