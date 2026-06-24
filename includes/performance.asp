<%
' ============================================
' V15.0 性能优化工具集 (Performance Optimization)
' 依赖: config.asp
' 用法: <!--#include file="performance.asp"-->
' ============================================

' 性能开关
Const PERF_ENABLE_GZIP = True           ' 启用GZIP压缩
Const PERF_ENABLE_CACHE_HEADERS = True  ' 启用Cache-Control优化
Const PERF_ENABLE_LAZY_LOAD = True      ' 启用原生懒加载 (loading="lazy")
Const PERF_ENABLE_PRELOAD = True        ' 启用关键资源预加载

' 缓存时长配置（秒）
Const PERF_CACHE_STATIC = 31536000      ' 静态资源：1年
Const PERF_CACHE_IMAGES = 2592000       ' 图片：30天
Const PERF_CACHE_PAGES = 3600           ' 页面：1小时
Const PERF_CACHE_API = 0                ' API：不缓存

' ============================================
' 设置静态资源缓存头
' 用法: Call PERF_SetStaticCacheHeaders("css")
' ============================================
Sub PERF_SetStaticCacheHeaders(resourceType)
    If Not PERF_ENABLE_CACHE_HEADERS Then Exit Sub
    
    Dim maxAge
    Select Case LCase(resourceType)
        Case "css", "js", "font", "woff", "woff2"
            maxAge = PERF_CACHE_STATIC
        Case "image", "img", "png", "jpg", "svg", "webp", "ico"
            maxAge = PERF_CACHE_IMAGES
        Case "api", "json"
            maxAge = PERF_CACHE_API
        Case Else
            maxAge = PERF_CACHE_PAGES
    End Select
    
    If maxAge > 0 Then
        Response.AddHeader "Cache-Control", "public, max-age=" & maxAge & ", immutable"
        Response.AddHeader "Expires", DateAdd("s", maxAge, Now())
    Else
        Response.AddHeader "Cache-Control", "no-cache, no-store, must-revalidate"
    End If
End Sub

' ============================================
' 输出预加载链接头
' 用法: Call PERF_PreloadResource("/css/style.css", "style")
' ============================================
Sub PERF_PreloadResource(url, resourceType)
    If Not PERF_ENABLE_PRELOAD Then Exit Sub
    Response.AddHeader "Link", "<" & url & ">; rel=preload; as=" & resourceType
End Sub

' ============================================
' 生成版本化资源URL
' 用法: Response.Write PERF_AssetURL("/css/style.css")
' 输出: /css/style.css?v=15.0
' ============================================
Function PERF_AssetURL(path)
    If InStr(path, "?") > 0 Then
        PERF_AssetURL = path & "&v=" & SYS_VERSION
    Else
        PERF_AssetURL = path & "?v=" & SYS_VERSION
    End If
End Function

' ============================================
' 图片懒加载属性生成
' 用法: Response.Write PERF_LazyImage("/images/product.jpg", "产品图片")
' ============================================
Function PERF_LazyImage(src, alt)
    Dim attrs
    attrs = "src=""" & src & """ alt=""" & alt & """"
    If PERF_ENABLE_LAZY_LOAD Then
        attrs = attrs & " loading=""lazy"" decoding=""async"""
    End If
    PERF_LazyImage = "<img " & attrs & ">"
End Function

' ============================================
' 启用GZIP压缩（如果未开启）
' 注：IIS通常已启用，此处为双重保险
' ============================================
Sub PERF_EnableGzip()
    If Not PERF_ENABLE_GZIP Then Exit Sub
    ' IIS通过web.config配置，此处无需额外操作
    ' ASP层不做重复压缩，仅做标记
End Sub

' ============================================
' 数据库索引优化查询建议（手动运行）
' 用法: PERF_GetIndexSuggestions() 返回建议SQL数组
' ============================================
Function PERF_GetIndexSuggestions()
    Dim suggestions(9)
    
    suggestions(0) = "-- 检查Producs表缺少ProductName搜索索引"
    suggestions(1) = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_Products_ProductName')" & vbCrLf & _
                     "CREATE INDEX IX_Products_ProductName ON Products(ProductName) INCLUDE (BasePrice,Category,ImageURL);"
    
    suggestions(2) = "-- 检查Orders复合查询索引"
    suggestions(3) = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_Orders_UserID_Status')" & vbCrLf & _
                     "CREATE INDEX IX_Orders_UserID_Status ON Orders(UserID,Status) INCLUDE (TotalAmount,CreatedAt);"
    
    suggestions(4) = "-- 检查Cart表SessionID索引"
    suggestions(5) = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_Cart_SessionID')" & vbCrLf & _
                     "CREATE INDEX IX_Cart_SessionID ON Cart(SessionID) INCLUDE (ProductID,Quantity);"
    
    suggestions(6) = "-- 检查OrderDetails聚合索引"
    suggestions(7) = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_OrderDetails_ProductID_Subtotal')" & vbCrLf & _
                     "CREATE INDEX IX_OrderDetails_ProductID_Subtotal ON OrderDetails(ProductID) INCLUDE (Quantity,Subtotal);"
    
    suggestions(8) = "-- 慢查询日志分析"
    suggestions(9) = "SELECT TOP 20 qs.total_elapsed_time/qs.execution_count AS avg_time_ms," & vbCrLf & _
                     "qs.execution_count, SUBSTRING(st.text, qs.statement_start_offset/2+1," & vbCrLf & _
                     "CASE WHEN qs.statement_end_offset=-1 THEN LEN(st.text) ELSE (qs.statement_end_offset-qs.statement_start_offset)/2+1 END) AS query_text" & vbCrLf & _
                     "FROM sys.dm_exec_query_stats qs CROSS APPLY sys.dm_exec_sql_text(qs.sql_handle) st" & vbCrLf & _
                     "ORDER BY avg_time_ms DESC;"
    
    PERF_GetIndexSuggestions = suggestions
End Function

' ============================================
' 页面执行计时器
' ============================================
Dim PERF_PageStartTime
PERF_PageStartTime = Timer()

Function PERF_GetElapsedMs()
    PERF_GetElapsedMs = Round((Timer() - PERF_PageStartTime) * 1000, 1)
End Function

' ============================================
' 输出调试信息（HTML注释）
' ============================================
Sub PERF_OutputDebugInfo()
    Response.Write vbCrLf & "<!-- V15.0 Page rendered in " & PERF_GetElapsedMs() & "ms | " & Now() & " -->" & vbCrLf
End Sub
%>