$wc = New-Object System.Net.WebClient
$wc.Encoding = [System.Text.Encoding]::UTF8

# Test all type filters
$types = @("", "custom", "standard", "kol")
foreach ($type in $types) {
    $url = "http://localhost/products.asp"
    if ($type -ne "") { $url += "?type=$type" }
    $result = $wc.DownloadString($url)
    
    # Extract product count
    if ($result -match 'product-card|product-item') {
        # Count product cards
        $matches = [regex]::Matches($result, 'product-card')
        $count = $matches.Count
        Write-Host "Type='$type': $count product cards found"
    }
    
    # Also look for the count text
    if ($result -match '(\d+)\s*\u4ef6\u5546\u54c1') {
        Write-Host "Type='$type': Count text = $($matches[1])"
    }
    
    Write-Host "---"
}
