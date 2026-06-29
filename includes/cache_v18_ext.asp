<%
' ============================================
' V18.0 缓存管理器扩展 (Cache Manager Extensions)
' 由 cache_manager.asp 自动引入
' ============================================

' ============================================
' CM_GetOrEmpty - 获取缓存值或返回 Empty
' 简化调用：无需额外判断 IsEmpty
' ============================================
Function CM_GetOrEmpty(key)
    Dim val : val = CM_Get(key)
    If IsEmpty(val) Then
        CM_GetOrEmpty = Empty
    Else
        CM_GetOrEmpty = val
    End If
End Function

' ============================================
' CM_CacheProductQuery - 缓存产品查询结果包装
' 用于首页栏目、产品列表等高频查询
' 返回: 缓存的数据，未命中时返回 Empty
' 用法:
'   Dim html : html = CM_CacheProductQuery("home_section_standard", 600)
'   If IsEmpty(html) Then
'       ' 从数据库查询并构建 HTML
'       html = BuildSectionHTML(...)
'       Call CM_Set("home_section_standard", html, 600)
'   End If
'   Response.Write html
' ============================================
Function CM_CacheProductQuery(cacheKey, ttlSeconds)
    If Not FEATURE_CACHE_MANAGER Then
        CM_CacheProductQuery = Empty
        Exit Function
    End If
    If IsNull(ttlSeconds) Or ttlSeconds <= 0 Then ttlSeconds = CM_DEFAULT_TTL
    CM_CacheProductQuery = CM_Get(cacheKey)
End Function

' ============================================
' CM_ClearProductCaches - 清除所有产品相关缓存
' 在产品新增/修改/删除后调用
' ============================================
Sub CM_ClearProductCaches()
    If Not FEATURE_CACHE_MANAGER Then Exit Sub
    Call CM_Clear("home_section_")
    Call CM_Clear("products_list_")
    Call CM_Clear("product_detail_")
    Call CM_Clear("recommend_")
    Call CM_Clear("search_suggest_")
End Sub

' ============================================
' CM_RecordCacheHeaders - 添加缓存状态响应头
' 用于调试和性能监控
' ============================================
Sub CM_RecordCacheHeaders(isHit)
    If isHit Then
        Response.AddHeader "X-Cache", "HIT"
        Call CM_RecordHit()
    Else
        Response.AddHeader "X-Cache", "MISS"
        Call CM_RecordMiss()
    End If
End Sub

' ============================================
' CM_GetCacheStatsJSON - 获取缓存统计 JSON
' 用于 /api/health_check.asp 监控端点
' ============================================
Function CM_GetCacheStatsJSON()
    Dim stats, json
    Set stats = CM_GetStats()
    json = "{"
    json = json & """hits"":" & stats("hits") & ","
    json = json & """misses"":" & stats("misses") & ","
    json = json & """hitRate"":" & stats("hitRate") & ","
    json = json & """activeEntries"":" & stats("activeEntries") & ","
    json = json & """enabled"":" & LCase(stats("enabled"))
    json = json & "}"
    Set stats = Nothing
    CM_GetCacheStatsJSON = json
End Function
%>
