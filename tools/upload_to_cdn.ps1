<#
.SYNOPSIS
    PerfumeShop V13.3 - CDN资源上传脚本
.DESCRIPTION
    将静态资源（图片/CSS/JS）批量上传到CDN
    支持阿里云OSS和腾讯云COS
.PARAMETER Provider
    CDN提供商：Aliyun 或 Tencent（默认Aliyun）
.PARAMETER DryRun
    预览模式，不实际执行上传
.EXAMPLE
    .\upload_to_cdn.ps1 -DryRun
    .\upload_to_cdn.ps1 -Provider Tencent
#>

param(
    [ValidateSet("Aliyun", "Tencent")]
    [string]$Provider = "Aliyun",
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
$projectRoot = "F:\网站制作\网站\网站二"

# ============ CDN配置 ============
$cdnConfig = @{
    Aliyun = @{
        Bucket    = "perfumeshop-cdn"
        Region    = "oss-cn-hangzhou"
        Endpoint  = "oss-cn-hangzhou.aliyuncs.com"
        Tool      = "ossutil64"
        ToolCheck = { Get-Command ossutil64 -ErrorAction SilentlyContinue }
    }
    Tencent = @{
        Bucket    = "perfumeshop-1250000000"
        Region    = "ap-guangzhou"
        Tool      = "coscmd"
        ToolCheck = { Get-Command coscmd -ErrorAction SilentlyContinue }
    }
}

# ============ 检查CDN工具 ============
$cfg = $cdnConfig[$Provider]
$tool = & $cfg.ToolCheck
if (-not $tool) {
    Write-Host "[错误] 未找到 $($cfg.Tool) 工具，请先安装" -ForegroundColor Red
    Write-Host ""
    Write-Host "阿里云OSS安装：" -ForegroundColor Yellow
    Write-Host "  winget install aliyun.ossutil" -ForegroundColor Gray
    Write-Host "腾讯云COS安装：" -ForegroundColor Yellow
    Write-Host "  pip install coscmd" -ForegroundColor Gray
    exit 1
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " PerfumeShop CDN 资源上传工具 V13.3" -ForegroundColor Cyan
Write-Host " 提供商: $Provider" -ForegroundColor Cyan
if ($DryRun) { Write-Host " 模式: 预览 (不会实际执行)" -ForegroundColor Yellow }
Write-Host "========================================" -ForegroundColor Cyan

# ============ 上传目录列表 ============
$uploadDirs = @(
    @{ Local = "$projectRoot\images"; Remote = "/images/"; Desc = "产品图片" },
    @{ Local = "$projectRoot\css";    Remote = "/css/";    Desc = "样式文件", Exclude = "*.map" },
    @{ Local = "$projectRoot\js";     Remote = "/js/";     Desc = "脚本文件", Exclude = "*.map" }
)

foreach ($dir in $uploadDirs) {
    $localPath = $dir.Local
    $remotePath = $dir.Remote
    $desc = $dir.Desc
    
    if (-not (Test-Path $localPath)) {
        Write-Host "[跳过] $desc - 目录不存在: $localPath" -ForegroundColor DarkGray
        continue
    }
    
    # 统计文件数量
    $fileCount = (Get-ChildItem -Path $localPath -Recurse -File | Measure-Object).Count
    Write-Host "[上传] $desc ($fileCount 个文件)" -ForegroundColor Green
    
    if ($DryRun) {
        Write-Host "  [预览] $localPath -> $remotePath" -ForegroundColor Gray
        continue
    }
    
    # 执行上传
    try {
        switch ($Provider) {
            "Aliyun" {
                $cmd = "ossutil64 cp -r -u `"$localPath`" `"oss://$($cfg.Bucket)$remotePath`" --region=$($cfg.Region)"
            }
            "Tencent" {
                $cmd = "coscmd upload -r `"$localPath`" `"$remotePath`""
            }
        }
        Write-Host "  执行: $cmd" -ForegroundColor DarkGray
        Invoke-Expression $cmd
    }
    catch {
        Write-Host "  [失败] $desc 上传出错: $_" -ForegroundColor Red
    }
}

# ============ 设置缓存头（阿里云OSS） ============
if (-not $DryRun -and $Provider -eq "Aliyun") {
    Write-Host ""
    Write-Host "[缓存] 设置CDN缓存策略..." -ForegroundColor Green
    
    # 图片：30天缓存
    $imgCmd = "ossutil64 set-meta oss://$($cfg.Bucket)/images/ Cache-Control:max-age=2592000 --region=$($cfg.Region) --update --recursive"
    Write-Host "  图片: max-age=2592000 (30天)" -ForegroundColor DarkGray
    Invoke-Expression $imgCmd
    
    # CSS/JS：7天缓存
    $staticCmd = "ossutil64 set-meta oss://$($cfg.Bucket)/css/ Cache-Control:max-age=604800 --region=$($cfg.Region) --update --recursive"
    Invoke-Expression $staticCmd
    $staticCmd = "ossutil64 set-meta oss://$($cfg.Bucket)/js/ Cache-Control:max-age=604800 --region=$($cfg.Region) --update --recursive"
    Invoke-Expression $staticCmd
    Write-Host "  CSS/JS: max-age=604800 (7天)" -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " CDN上传完成！" -ForegroundColor Green
Write-Host ""
Write-Host " 下一步：修改 includes\cdn_config.asp" -ForegroundColor Yellow
Write-Host "   CDN_ENABLED = True" -ForegroundColor White
Write-Host "   CDN_DOMAIN = `"https://cdn.perfumeshop.com`"" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Cyan
