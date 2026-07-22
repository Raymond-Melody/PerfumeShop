<#
.SYNOPSIS
  V19 服务选择器 - 选择开机自启动的服务
.DESCRIPTION
  运行此菜单选择开机时启动哪些服务，选择保存到 v19_service_choice.txt
#>

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$configFile = Join-Path $scriptDir "v19_service_choice.txt"
$current = ""
if (Test-Path $configFile) { $current = (Get-Content $configFile -Raw).Trim() }

Clear-Host
Write-Host ""
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "     V19 服务选择器" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "  选择开机自启动时运行的服务" -ForegroundColor White
Write-Host "  (当前: $(if($current -eq 'All' -or !$current){'全部'}elseif($current -eq 'Api'){'仅 API'}elseif($current -eq 'Admin'){'仅 Admin'}else{'全部'}))" -ForegroundColor Yellow
Write-Host ""
Write-Host "  1) 启动全部 (API + Admin) *默认" -ForegroundColor White
Write-Host "  2) 仅启动 API (端口 5000)" -ForegroundColor White
Write-Host "  3) 仅启动 Admin (端口 5207)" -ForegroundColor White
Write-Host "  0) 退出" -ForegroundColor White
Write-Host "=======================================" -ForegroundColor Cyan
$key = Read-Host "请选择 (1/2/3/0)"

switch ($key) {
    "1" { $c = "All"; $msg = "开机将启动全部服务 (API + Admin)" }
    "2" { $c = "Api"; $msg = "开机将仅启动 API (端口 5000)" }
    "3" { $c = "Admin"; $msg = "开机将仅启动 Admin (端口 5207)" }
    "0" { Write-Host "已退出"; exit }
    default { Write-Host "无效选择"; exit }
}

$c | Out-File -FilePath $configFile -Encoding utf8
Write-Host ""
Write-Host "================================" -ForegroundColor Green
Write-Host " $msg" -ForegroundColor Green
Write-Host " 选择已保存到: $configFile" -ForegroundColor Green
Write-Host " 立即启动请运行: .\start_v19_services.ps1" -ForegroundColor Green
Write-Host "================================" -ForegroundColor Green
