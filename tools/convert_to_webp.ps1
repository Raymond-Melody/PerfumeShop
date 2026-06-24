# ============================================
# PerfumeShop V12.0 - WebP图片转换工具
# 批量转换images目录下的JPEG/PNG为WebP格式
# ============================================

param(
    [string]$SourceDir = ".\images",
    [string]$Quality = "80"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "PerfumeShop V12.0 - WebP图片转换工具" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 检查cwebp是否安装
$cwebp = Get-Command cwebp -ErrorAction SilentlyContinue
if (-not $cwebp) {
    Write-Host "[错误] 未找到cwebp工具，请先安装WebP：" -ForegroundColor Red
    Write-Host "  下载地址: https://developers.google.com/speed/webp/download" -ForegroundColor Yellow
    Write-Host "  或将cwebp.exe添加到PATH环境变量" -ForegroundColor Yellow
    exit 1
}

Write-Host "[信息] cwebp工具已找到: $($cwebp.Source)" -ForegroundColor Green
Write-Host "[信息] 源目录: $SourceDir" -ForegroundColor Green
Write-Host "[信息] 质量设置: $Quality" -ForegroundColor Green
Write-Host ""

# 统计信息
$totalCount = 0
$successCount = 0
$skipCount = 0
$errorCount = 0

# 查找所有JPEG和PNG图片
$imageFiles = Get-ChildItem -Path $SourceDir -Recurse -Include "*.jpg","*.jpeg","*.png" -File

Write-Host "[信息] 找到 $($imageFiles.Count) 个图片文件" -ForegroundColor Green
Write-Host ""

foreach ($file in $imageFiles) {
    $totalCount++
    $webpPath = $file.FullName -replace '\.(jpg|jpeg|png)$', '.webp'
    
    # 检查WebP是否已存在且比源文件新
    if (Test-Path $webpPath) {
        $webpFile = Get-Item $webpPath
        if ($webpFile.LastWriteTime -ge $file.LastWriteTime) {
            Write-Host "[跳过] $($file.Name) - WebP已存在且为最新" -ForegroundColor Gray
            $skipCount++
            continue
        }
    }
    
    Write-Host "[转换] $($file.Name) -> $($file.BaseName).webp" -ForegroundColor Yellow
    
    # 执行转换
    $arguments = "-q", $Quality, "`"$($file.FullName)`"", "-o", "`"$webpPath`""
    $result = Start-Process -FilePath $cwebp.Source -ArgumentList $arguments -Wait -NoNewWindow -PassThru
    
    if ($result.ExitCode -eq 0) {
        $successCount++
        
        # 计算压缩率
        $originalSize = $file.Length
        $webpSize = (Get-Item $webpPath).Length
        $compressionRatio = [math]::Round((1 - ($webpSize / $originalSize)) * 100, 2)
        
        Write-Host "  [成功] 原始: $([math]::Round($originalSize/1024, 2))KB -> WebP: $([math]::Round($webpSize/1024, 2))KB (压缩率: ${compressionRatio}%)" -ForegroundColor Green
    } else {
        $errorCount++
        Write-Host "  [失败] 转换出错" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "转换完成！" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "总计: $totalCount 个文件" -ForegroundColor White
Write-Host "成功: $successCount 个" -ForegroundColor Green
Write-Host "跳过: $skipCount 个" -ForegroundColor Gray
Write-Host "失败: $errorCount 个" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Green" })
Write-Host ""

if ($successCount -gt 0) {
    Write-Host "[提示] 现在需要在产品页面中使用<picture>标签支持WebP" -ForegroundColor Yellow
    Write-Host "[提示] 示例代码已生成到 docs/webp_usage_example.html" -ForegroundColor Yellow
}
