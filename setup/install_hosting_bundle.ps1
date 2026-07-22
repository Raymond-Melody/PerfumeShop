# ASP.NET Core 8.0 Hosting Bundle 静默安装脚本
# 请以管理员身份运行此脚本

$installer = "$env:TEMP\dotnet-hosting-8.0.28-win.exe"
$logFile = "$PSScriptRoot\install_hosting_bundle.log"

Write-Host "正在安装 ASP.NET Core 8.0 Hosting Bundle..." -ForegroundColor Cyan
Write-Host "安装程序: $installer" -ForegroundColor Cyan
Write-Host "日志: $logFile" -ForegroundColor Cyan

if (-not (Test-Path $installer)) {
    Write-Host "错误: 安装程序未找到，请先下载" -ForegroundColor Red
    exit 1
}

# 执行静默安装
Start-Process -FilePath $installer -ArgumentList "/quiet /norestart OPT_NO_RUNTIME=0" -Wait -NoNewWindow

# 检查安装结果
$check = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\IIS Extensions\AspNetCore Module\Module" -Name "FileVersion" -ErrorAction SilentlyContinue
if ($check) {
    Write-Host "Hosting Bundle 安装成功! 版本: $($check.FileVersion)" -ForegroundColor Green
} else {
    Write-Host "安装可能未完成，请手动检查" -ForegroundColor Yellow
}

Read-Host "按 Enter 键退出"
