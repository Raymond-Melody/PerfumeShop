[System.IO.File]::WriteAllText(
  'f:\网站制作\网站\网站二\product.asp',
  ([System.IO.File]::ReadAllText('f:\网站制作\网站\网站二\product.asp') -replace
    '<noscript>\r\n						<div class="star-selector"',
    '<noscript><div class="star-selector"'
  ),
  (New-Object System.Text.UTF8Encoding $true)
)
Write-Host 'step1 done'
