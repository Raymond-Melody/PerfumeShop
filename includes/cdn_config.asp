<%
' ============================================
' PerfumeShop V14.6 - 图片CDN加速配置
' ============================================

' CDN域名配置
Dim CDN_DOMAIN
CDN_DOMAIN = "https://cdn.perfumeshop.com" ' 替换为实际CDN域名

' CDN开关（生产环境设为True）
Dim CDN_ENABLED
CDN_ENABLED = False ' 开发环境关闭，生产环境开启

' 获取图片CDN URL
Function GetImageCDNURL(localPath)
    If Not CDN_ENABLED Then
        GetImageCDNURL = localPath
        Exit Function
    End If
    
    ' 如果已经是完整URL，直接返回
    If InStr(localPath, "http") = 1 Or InStr(localPath, "//") = 1 Then
        GetImageCDNURL = localPath
        Exit Function
    End If
    
    ' 转换为CDN URL
    If Left(localPath, 1) <> "/" Then
        localPath = "/" & localPath
    End If
    
    GetImageCDNURL = CDN_DOMAIN & localPath
End Function

' 获取产品图片CDN URL
Function GetProductImageCDN(imageURL)
    If imageURL = "" Or IsNull(imageURL) Then
        GetProductImageCDN = GetImageCDNURL("/images/default-product.svg")
    Else
        GetProductImageCDN = GetImageCDNURL(imageURL)
    End If
End Function

' 批量替换HTML中的图片URL为CDN URL
Function ConvertImagesToCDN(htmlContent)
    If Not CDN_ENABLED Then
        ConvertImagesToCDN = htmlContent
        Exit Function
    End If
    
    Dim regex, matches, match
    Set regex = New RegExp
    regex.Pattern = "src=""(/images/[^""]*?)"""
    regex.Global = True
    regex.IgnoreCase = True
    
    Set matches = regex.Execute(htmlContent)
    
    For Each match In matches
        Dim oldSrc, newSrc
        oldSrc = match.SubMatches(0)
        newSrc = CDN_DOMAIN & oldSrc
        htmlContent = Replace(htmlContent, oldSrc, newSrc)
    Next
    
    ConvertImagesToCDN = htmlContent
End Function

' ============================================
' web.config CDN路由配置（已集成到 web.config）
' 启用方法：取消 web.config 中 V13.3 CDN URL重写规则的注释
' ============================================

' ============================================
' CDN上传脚本（已保存为 tools/upload_to_cdn.ps1）
' 运行: .\tools\upload_to_cdn.ps1 -Provider Aliyun -DryRun
' ============================================

' ============================================
' CDN缓存配置建议
' ============================================
' 1. 图片缓存策略：
'    - 产品图片：30天（Cache-Control: max-age=2592000）
'    - 用户上传：7天（Cache-Control: max-age=604800）
'    - 默认图片：永久（Cache-Control: max-age=31536000）
' 2. CSS/JS缓存策略：
'    - 带版本号：1年（?v=13.3）
'    - 不带版本号：1天
' 3. 压缩配置：
'    - 启用GZIP压缩 / Brotli压缩 / WebP自动转换
' 4. HTTPS配置：
'    - 强制HTTPS / HSTS头部 / TLS 1.2+
' ============================================
%>
