<%
' ============================================
' V15.0 缓存管理器 (Cache Manager)
' 依赖: config.asp (CacheGet/CacheSet 兼容)
' 用法: <!--#include file="cache_manager.asp"-->
' 调用: CM_Get "products_list"
'        CM_Set "products_list", jsonString, 600
'        CM_Clear "products_"  ' 清除所有以 "products_" 开头的缓存
' ============================================

' 缓存配置
Const CM_DEFAULT_TTL = 600              ' 默认过期时间（秒），10分钟
Const CM_CLEANUP_THRESHOLD = 100        ' 每N个Session触发一次清理
Const CM_FILE_CACHE_DIR = "/cache/"      ' 文件缓存目录
Const CM_FILE_CACHE_TTL = 3600          ' 文件缓存默认过期（秒），1小时

' ============================================
' 缓存统计对象（存储在Application中）
' ============================================
Sub CM_InitStats()
    If Not IsObject(Application("CM_Stats")) Then
        Application.Lock
        If Not IsObject(Application("CM_Stats")) Then
            Dim stats
            Set stats = Server.CreateObject("Scripting.Dictionary")
            stats.Add "hits", 0
            stats.Add "misses", 0
            stats.Add "sets", 0
            stats.Add "deletes", 0
            stats.Add "lastCleanup", CDbl(Now())
            Set Application("CM_Stats") = stats
        End If
        Application.UnLock
    End If
End Sub

' ============================================
' 获取缓存（增强版：自动过期检查 + 统计 + 延迟清理）
' ============================================
Function CM_Get(key)
    Dim cacheKey : cacheKey = "CACHE_" & key
    Dim cacheEntry, parts, cacheTime, cacheTTL, nowTs
    
    Call CM_InitStats()
    
    cacheEntry = Application(cacheKey)
    If IsEmpty(cacheEntry) Or cacheEntry = "" Then
        ' 尝试文件缓存
        CM_Get = CM_FileGet(key)
        If IsEmpty(CM_Get) Then
            CM_RecordMiss()
        Else
            CM_RecordHit()
        End If
        Exit Function
    End If
    
    parts = Split(cacheEntry, Chr(1))
    If UBound(parts) < 2 Then
        CM_RecordMiss()
        CM_Get = Empty
        Exit Function
    End If
    
    cacheTime = CDbl(parts(0))
    cacheTTL = CInt(parts(1))
    nowTs = CDbl(Now())
    
    ' 检查过期
    If (nowTs - cacheTime) > (cacheTTL / 86400.0) Then
        ' 过期 — 延迟清理（不立即删除，减少锁竞争）
        Dim stats
        Set stats = Application("CM_Stats")
        Application.Lock
        stats("misses") = stats("misses") + 1
        Application.UnLock
        Set stats = Nothing
        CM_Get = Empty
        Exit Function
    End If
    
    CM_RecordHit()
    CM_Get = parts(2)
End Function

' ============================================
' 设置缓存
' ============================================
Sub CM_Set(key, value, ttlSeconds)
    Dim cacheKey : cacheKey = "CACHE_" & key
    Dim cacheEntry
    
    If IsNull(ttlSeconds) Or ttlSeconds <= 0 Then ttlSeconds = CM_DEFAULT_TTL
    
    cacheEntry = CDbl(Now()) & Chr(1) & ttlSeconds & Chr(1) & value
    
    Application.Lock
    Application(cacheKey) = cacheEntry
    Application.UnLock
    
    ' 记录统计
    Call CM_InitStats()
    Dim stats
    Set stats = Application("CM_Stats")
    Application.Lock
    stats("sets") = stats("sets") + 1
    Application.UnLock
    Set stats = Nothing
    
    ' 同时写入文件缓存（异步优化）
    If ttlSeconds >= CM_FILE_CACHE_TTL Then
        Call CM_FileSet(key, value, ttlSeconds)
    End If
End Sub

' ============================================
' 删除单个缓存
' ============================================
Sub CM_Delete(key)
    Dim cacheKey : cacheKey = "CACHE_" & key
    
    Application.Lock
    Application.Contents.Remove(cacheKey)
    Application.UnLock
    
    Call CM_InitStats()
    Dim stats
    Set stats = Application("CM_Stats")
    Application.Lock
    stats("deletes") = stats("deletes") + 1
    Application.UnLock
    Set stats = Nothing
    
    ' 同时删除文件缓存
    Call CM_FileDelete(key)
End Sub

' ============================================
' 按前缀批量清除缓存
' ============================================
Sub CM_Clear(prefix)
    Dim contents, i, key, prefixUpper, fullPrefix
    fullPrefix = "CACHE_" & prefix
    prefixUpper = UCase(fullPrefix)
    
    Application.Lock
    Set contents = Application.Contents
    ' 收集要删除的键（不能在遍历时直接删除）
    Dim keysToDelete(), deleteCount
    deleteCount = 0
    ReDim keysToDelete(0)
    
    For i = 0 To contents.Count - 1
        key = contents.Keys()(i)
        If UCase(Left(key, Len(fullPrefix))) = prefixUpper Then
            ReDim Preserve keysToDelete(deleteCount)
            keysToDelete(deleteCount) = key
            deleteCount = deleteCount + 1
        End If
    Next
    
    ' 执行删除
    For i = 0 To UBound(keysToDelete)
        Application.Contents.Remove(keysToDelete(i))
    Next
    Application.UnLock
    Set contents = Nothing
    
    ' 更新统计
    If deleteCount > 0 Then
        Call CM_InitStats()
        Dim stats
        Set stats = Application("CM_Stats")
        Application.Lock
        stats("deletes") = stats("deletes") + deleteCount
        Application.UnLock
        Set stats = Nothing
    End If
End Sub

' ============================================
' 主动清理所有过期缓存
' ============================================
Sub CM_Cleanup()
    Dim contents, i, key, cacheEntry, parts, cacheTime, cacheTTL, nowTs, fullPrefix
    fullPrefix = "CACHE_"
    nowTs = CDbl(Now())
    
    Application.Lock
    Set contents = Application.Contents
    
    Dim keysToDelete(), deleteCount
    deleteCount = 0
    ReDim keysToDelete(0)
    
    For i = 0 To contents.Count - 1
        key = contents.Keys()(i)
        If UCase(Left(key, Len(fullPrefix))) = "CACHE_" Then
            cacheEntry = contents.Item(key)
            If Not IsEmpty(cacheEntry) And cacheEntry <> "" Then
                parts = Split(cacheEntry, Chr(1))
                If UBound(parts) >= 2 Then
                    cacheTime = CDbl(parts(0))
                    cacheTTL = CInt(parts(1))
                    If (nowTs - cacheTime) > (cacheTTL / 86400.0) Then
                        ReDim Preserve keysToDelete(deleteCount)
                        keysToDelete(deleteCount) = key
                        deleteCount = deleteCount + 1
                    End If
                End If
            End If
        End If
    Next
    
    For i = 0 To UBound(keysToDelete)
        Application.Contents.Remove(keysToDelete(i))
    Next
    
    ' 更新最后清理时间
    If IsObject(Application("CM_Stats")) Then
        Application("CM_Stats")("lastCleanup") = nowTs
    End If
    
    Application.UnLock
    Set contents = Nothing
End Sub

' ============================================
' 获取缓存统计信息
' ============================================
Function CM_GetStats()
    Call CM_InitStats()
    
    Dim stats, result
    Set stats = Application("CM_Stats")
    Set result = Server.CreateObject("Scripting.Dictionary")
    
    Dim hits, misses, sets, deletes, total
    hits = CLng(stats("hits"))
    misses = CLng(stats("misses"))
    total = hits + misses
    
    result.Add "hits", hits
    result.Add "misses", misses
    result.Add "totalRequests", total
    If total > 0 Then
        result.Add "hitRate", Round((hits / total) * 100, 1)
    Else
        result.Add "hitRate", 0
    End If
    result.Add "sets", CLng(stats("sets"))
    result.Add "deletes", CLng(stats("deletes"))
    result.Add "lastCleanup", stats("lastCleanup")
    
    ' 统计Application中缓存项数量
    Dim contents, i, key, cacheCount
    cacheCount = 0
    Set contents = Application.Contents
    For i = 0 To contents.Count - 1
        key = contents.Keys()(i)
        If UCase(Left(key, 6)) = "CACHE_" Then
            cacheCount = cacheCount + 1
        End If
    Next
    Set contents = Nothing
    result.Add "activeEntries", cacheCount
    result.Add "enabled", FEATURE_CACHE_MANAGER
    
    Set CM_GetStats = result
End Sub

' ============================================
' 记录缓存命中
' ============================================
Sub CM_RecordHit()
    Dim stats
    Set stats = Application("CM_Stats")
    If IsObject(stats) Then
        Application.Lock
        stats("hits") = stats("hits") + 1
        Application.UnLock
    End If
    Set stats = Nothing
End Sub

' ============================================
' 记录缓存未命中
' ============================================
Sub CM_RecordMiss()
    Dim stats
    Set stats = Application("CM_Stats")
    If IsObject(stats) Then
        Application.Lock
        stats("misses") = stats("misses") + 1
        Application.UnLock
    End If
    Set stats = Nothing
End Sub

' ============================================
' Session启动时触发清理（随机概率触发）
' ============================================
Sub CM_MaybeCleanup()
    Dim sessionCounter
    
    ' 使用Session计数器减少清理频率
    sessionCounter = CLng(Application("CM_SessionCounter"))
    sessionCounter = sessionCounter + 1
    Application.Lock
    Application("CM_SessionCounter") = sessionCounter
    Application.UnLock
    
    ' 每 CM_CLEANUP_THRESHOLD 个Session触发一次清理
    If sessionCounter Mod CM_CLEANUP_THRESHOLD = 0 Then
        Call CM_Cleanup()
    End If
End Sub

' ============================================
' 文件缓存：获取（二级缓存）
' ============================================
Function CM_FileGet(key)
    Dim fso, filePath, ts, ttl, value, nowTs
    On Error Resume Next
    
    Dim safeKey : safeKey = Replace(Replace(key, "\", "_"), "/", "_")
    filePath = Server.MapPath(CM_FILE_CACHE_DIR & safeKey & ".cache")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then
        Err.Clear
        Set fso = Nothing
        CM_FileGet = Empty
        Exit Function
    End If
    
    If Not fso.FileExists(filePath) Then
        Set fso = Nothing
        CM_FileGet = Empty
        Exit Function
    End If
    
    Dim file
    Set file = fso.OpenTextFile(filePath, 1) ' ForReading
    ts = CDbl(file.ReadLine())
    ttl = CInt(file.ReadLine())
    value = file.ReadAll()
    file.Close
    Set file = Nothing
    Set fso = Nothing
    
    nowTs = CDbl(Now())
    If (nowTs - ts) > (ttl / 86400.0) Then
        CM_FileGet = Empty
        Exit Function
    End If
    
    CM_FileGet = value
    Err.Clear
End Function

' ============================================
' 文件缓存：设置
' ============================================
Sub CM_FileSet(key, value, ttlSeconds)
    Dim fso, filePath, safeKey, cacheDir
    On Error Resume Next
    
    safeKey = Replace(Replace(key, "\", "_"), "/", "_")
    cacheDir = Server.MapPath(CM_FILE_CACHE_DIR)
    filePath = cacheDir & "\" & safeKey & ".cache"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then
        Err.Clear
        Set fso = Nothing
        Exit Sub
    End If
    
    ' 确保目录存在
    If Not fso.FolderExists(cacheDir) Then
        fso.CreateFolder(cacheDir)
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        Set fso = Nothing
        Exit Sub
    End If
    
    Dim file
    Set file = fso.OpenTextFile(filePath, 2, True) ' ForWriting, Create
    file.WriteLine CDbl(Now())
    file.WriteLine ttlSeconds
    file.Write value
    file.Close
    Set file = Nothing
    Set fso = Nothing
    Err.Clear
End Sub

' ============================================
' 文件缓存：删除
' ============================================
Sub CM_FileDelete(key)
    Dim fso, filePath, safeKey
    On Error Resume Next
    
    safeKey = Replace(Replace(key, "\", "_"), "/", "_")
    filePath = Server.MapPath(CM_FILE_CACHE_DIR & safeKey & ".cache")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Err.Number = 0 Then
        If fso.FileExists(filePath) Then
            fso.DeleteFile filePath, True
        End If
    End If
    Err.Clear
    Set fso = Nothing
End Sub

' ============================================
' 缓存预热：应用启动时预加载热门数据
' ============================================
Sub CM_Warmup()
    ' 检查是否已预热
    If Application("CM_WarmupDone") Then Exit Sub
    
    Application.Lock
    If Application("CM_WarmupDone") Then
        Application.UnLock
        Exit Sub
    End If
    
    ' 预热标志
    Application("CM_WarmupDone") = True
    Application("CM_Initialized") = True
    Application("CM_SessionCounter") = 0
    Application.UnLock
    
    ' 初始化统计
    Call CM_InitStats()
End Sub

' ============================================
' V18: CM_GetOrEmpty - 获取缓存值或返回 Empty
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
' V18: CM_CacheProductQuery - 缓存产品查询结果
' 用法: 先尝试缓存，未命中时由调用方查询并 CM_Set
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
' V18: CM_ClearProductCaches - 清除所有产品相关缓存
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
' V18: CM_RecordCacheHeaders - 添加缓存状态响应头
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
' 初始化
' ============================================
If Not Application("CM_Initialized") Then
    Call CM_Warmup()
End If
%>