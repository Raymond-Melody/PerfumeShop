# POST to data_migrate.asp with action=migrate
try {
    $wc = New-Object System.Net.WebClient
    $wc.Encoding = [System.Text.Encoding]::UTF8
    $wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded")
    
    Write-Host "Starting migration via HTTP POST..."
    $result = $wc.UploadString("http://localhost/setup/data_migrate.asp", "action=migrate")
    
    # Save result
    $result | Out-File -FilePath (Join-Path (Get-Location) "setup\migrate_result.html") -Encoding UTF8
    Write-Host "Migration request completed. Result length: $($result.Length)"
    
    # Show summary lines
    $lines = $result -split "`n"
    foreach ($line in $lines) {
        if ($line -match "success|error|warning|complete|Migrated|rows") {
            # Strip HTML tags for display
            $clean = $line -replace '<[^>]+>', ''
            if ($clean.Trim() -ne '') {
                Write-Host $clean.Trim()
            }
        }
    }
} catch {
    Write-Host "Error: $($_.Exception.Message)"
    if ($_.Exception.InnerException) {
        Write-Host "Inner: $($_.Exception.InnerException.Message)"
    }
}
