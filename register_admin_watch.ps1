# V19 Admin dotnet watch 开机自启注册脚本
# 以管理员身份运行此脚本
# 作用：在注册表 Run 键中添加 V19 Admin (dotnet watch) 开机自启

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$vbsPath = Join-Path $scriptDir "watch_v19_admin.vbs"
$regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"

# 移除旧的 V19Services 启动项
$old = Get-ItemProperty -Path $regPath -Name "V19Services" -ErrorAction SilentlyContinue
if ($old) {
    Remove-ItemProperty -Path $regPath -Name "V19Services"
    Write-Host "[OK] 已移除旧启动项 V19Services" -ForegroundColor Green
}

# 移除旧的 V19_Watch 启动项（如果存在）
$old3 = Get-ItemProperty -Path $regPath -Name "V19_Watch" -ErrorAction SilentlyContinue
if ($old3) {
    Remove-ItemProperty -Path $regPath -Name "V19_Watch"
    Write-Host "[OK] 已移除旧启动项 V19_Watch" -ForegroundColor Green
}

# 移除旧的 V19AdminWatch 启动项（如果存在）
$old2 = Get-ItemProperty -Path $regPath -Name "V19AdminWatch" -ErrorAction SilentlyContinue
if ($old2) {
    Remove-ItemProperty -Path $regPath -Name "V19AdminWatch"
    Write-Host "[OK] 已移除旧启动项 V19AdminWatch" -ForegroundColor Green
}

# 注册新的开机自启：Admin dotnet watch
$value = 'wscript.exe "' + $vbsPath + '"'
Set-ItemProperty -Path $regPath -Name "V19AdminWatch" -Value $value

Write-Host ""
Write-Host "========== 已注册开机自启 ==========" -ForegroundColor Cyan
Write-Host "  启动项: V19AdminWatch" -ForegroundColor White
Write-Host "  启动内容: Admin (端口 5207) + dotnet watch" -ForegroundColor White
Write-Host "  脚本位置: $vbsPath" -ForegroundColor White
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "开机自启已配置完成！重启生效。" -ForegroundColor Green
