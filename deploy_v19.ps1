# V19部署脚本 - 以管理员身份运行
$ErrorActionPreference = "Stop"
$newDir = Join-Path $PSScriptRoot "publish\PerfumeShop.Admin.new"
$targetDir = Join-Path $PSScriptRoot "publish\PerfumeShop.Admin"
$serviceName = "PerfumeShopV19_Admin"

Write-Host "=== V19 Deployment ===" -ForegroundColor Cyan

Write-Host "[1/4] Stopping service..." -ForegroundColor Yellow
Stop-Service $serviceName -Force
Start-Sleep -Seconds 3
Write-Host "  Service stopped" -ForegroundColor Green

Write-Host "[2/4] Copying new files..." -ForegroundColor Yellow
Copy-Item (Join-Path $newDir "*") $targetDir -Recurse -Force
Write-Host "  Files copied" -ForegroundColor Green

Write-Host "[3/4] Starting service..." -ForegroundColor Yellow
Start-Service $serviceName
Start-Sleep -Seconds 5
Write-Host "  Service started" -ForegroundColor Green

Write-Host "[4/4] Cleaning up..." -ForegroundColor Yellow
Remove-Item $newDir -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "  Cleanup done" -ForegroundColor Green

Write-Host ""
Write-Host "=== Deployment Complete! ===" -ForegroundColor Cyan
Write-Host "V19: http://localhost:5207/admin" -ForegroundColor White
