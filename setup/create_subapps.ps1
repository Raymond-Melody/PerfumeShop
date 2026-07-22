# V19 IIS 子应用程序配置脚本
# 请以管理员身份运行

$webroot = "F:\网站制作\网站\网站二"
$appcmd = "$env:windir\system32\inetsrv\appcmd.exe"
$site = "Default Web Site"

Write-Host "=== 创建 IIS 子应用程序 ===" -ForegroundColor Cyan

# 删除旧的子应用
& $appcmd delete app /app.name:"$site\api" 2>$null
& $appcmd delete app /app.name:"$site\admin\v2" 2>$null

# 创建 /api 子应用程序
Write-Host "创建 /api → $webroot\v19-api" -ForegroundColor Yellow
& $appcmd add app /site.name:"$site" /path:"/api" /physicalPath:"$webroot\v19-api" /applicationPool:"DefaultAppPool"

# 创建 /admin/v2 子应用程序
Write-Host "创建 /admin/v2 → $webroot\v19-admin" -ForegroundColor Yellow
& $appcmd add app /site.name:"$site" /path:"/admin/v2" /physicalPath:"$webroot\v19-admin" /applicationPool:"DefaultAppPool"

Write-Host "=== 完成 ===" -ForegroundColor Green
