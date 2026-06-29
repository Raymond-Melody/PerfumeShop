<% 
' ============================================
' V18.0 速率限制器 (Rate Limiter)
' 基于 Session + IP 的令牌桶算法
' 使用 Application 字符串变量（兼容 Classic ASP）
' 支持 429 Too Many Requests 响应
' 依赖: config.asp
' 用法: <!--#include file="rate_limiter.asp"-->
'        If Not RL_Check("api_cart", 30, 60) Then Response.End
' ============================================

' 速率限制配置
Const RL_ENABLED = True                    ' 启用速率限制
Const RL_DEFAULT_MAX = 60                  ' 默认每窗口最大请求数
Const RL_DEFAULT_WINDOW = 60              ' 默认时间窗口（秒）
Const RL_CLEANUP_INTERVAL = 300            ' 清理间隔（秒）

' ============================================
' 内部辅助：生成 Application 存储键
' ============================================
Function RL_MakeAppKey(clientKey)
    RL_MakeAppKey = "RLB_" & clientKey
End Function

' ============================================
' 内部辅助：注册/注销客户端键到键列表
' ============================================
Sub RL_RegisterKey(clientKey)
    Dim keyList
    keyList = Application("RL_KeyList")
    If IsEmpty(keyList) Or keyList = "" Then keyList = ""
    ' 避免重复注册
    If InStr(1, "|" & keyList & "|", "|" & clientKey & "|", vbTextCompare) = 0 Then
        If keyList = "" Then
            keyList = clientKey
        Else
            keyList = keyList & "|" & clientKey
        End If
        Application("RL_KeyList") = keyList
    End If
End Sub

Sub RL_UnregisterKey(clientKey)
    Dim keyList, newList, keys, i
    keyList = Application("RL_KeyList")
    If IsEmpty(keyList) Or keyList = "" Then Exit Sub
    keys = Split(keyList, "|")
    newList = ""
    For i = 0 To UBound(keys)
        If keys(i) <> clientKey And keys(i) <> "" Then
            If newList = "" Then
                newList = keys(i)
            Else
                newList = newList & "|" & keys(i)
            End If
        End If
    Next
    Application("RL_KeyList") = newList
End Sub

' ============================================
' RL_Check(key, maxRequests, windowSec)
' 检查是否超过速率限制
' 参数:
'   key         - 限流键（如 "api_cart"、"api_login"）
'   maxRequests - 时间窗口内最大请求数
'   windowSec   - 时间窗口（秒）
' 返回: True=允许, False=拒绝（应返回 429）
' ============================================
Function RL_Check(key, maxRequests, windowSec)
    If Not RL_ENABLED Then
        RL_Check = True
        Exit Function
    End If
    
    If IsNull(maxRequests) Or maxRequests <= 0 Then maxRequests = RL_DEFAULT_MAX
    If IsNull(windowSec) Or windowSec <= 0 Then windowSec = RL_DEFAULT_WINDOW
    
    Dim clientKey, nowTs, bucket, appKey
    
    ' 构建客户端唯一键：IP + Session(可选) + key
    clientKey = "RL_" & key & "_" & Request.ServerVariables("REMOTE_ADDR")
    If Session("UserID") <> "" Then
        clientKey = clientKey & "_U" & Session("UserID")
    End If
    appKey = RL_MakeAppKey(clientKey)
    
    nowTs = CDbl(DateDiff("s", CDate("1970-01-01 00:00:00"), Now()))
    
    ' 从 Application 获取或创建令牌桶（使用字符串变量）
    Application.Lock
    
    bucket = Application(appKey)
    If IsEmpty(bucket) Or bucket = "" Then
        bucket = nowTs & "," & maxRequests
    End If
    
    Dim parts, windowStart, tokens
    parts = Split(bucket, ",")
    windowStart = CDbl(parts(0))
    tokens = CInt(parts(1))
    
    ' 检查是否需要重置窗口
    If nowTs - windowStart > windowSec Then
        windowStart = nowTs
        tokens = maxRequests
    End If
    
    ' 扣减令牌
    If tokens > 0 Then
        tokens = tokens - 1
        Application(appKey) = windowStart & "," & tokens
        Call RL_RegisterKey(clientKey)
        Application.UnLock
        RL_Check = True
        Exit Function
    End If
    
    ' 令牌耗尽
    Application(appKey) = windowStart & "," & 0
    Application.UnLock
    
    RL_Check = False
End Function

' ============================================
' RL_GetRemaining(key): 获取剩余令牌数
' 用于返回 X-RateLimit-Remaining 头
' ============================================
Function RL_GetRemaining(key)
    If Not RL_ENABLED Then
        RL_GetRemaining = -1
        Exit Function
    End If
    
    Dim clientKey, bucket, appKey
    clientKey = "RL_" & key & "_" & Request.ServerVariables("REMOTE_ADDR")
    appKey = RL_MakeAppKey(clientKey)
    
    bucket = Application(appKey)
    If IsEmpty(bucket) Or bucket = "" Then
        RL_GetRemaining = RL_DEFAULT_MAX
        Exit Function
    End If
    
    Dim parts
    parts = Split(bucket, ",")
    RL_GetRemaining = CInt(parts(1))
End Function

' ============================================
' RL_Cleanup(): 清理过期桶（定期调用）
' ============================================
Sub RL_Cleanup()
    Dim keyList, keys, i, key, bucket, parts, windowStart, appKey
    Dim nowTs : nowTs = CDbl(DateDiff("s", CDate("1970-01-01 00:00:00"), Now()))
    
    keyList = Application("RL_KeyList")
    If IsEmpty(keyList) Or keyList = "" Then Exit Sub
    
    Application.Lock
    
    ' 重新读取（防止并发修改）
    keyList = Application("RL_KeyList")
    If IsEmpty(keyList) Or keyList = "" Then
        Application.UnLock
        Exit Sub
    End If
    
    keys = Split(keyList, "|")
    For i = 0 To UBound(keys)
        key = keys(i)
        If key <> "" Then
            appKey = RL_MakeAppKey(key)
            bucket = Application(appKey)
            If Not IsEmpty(bucket) And bucket <> "" Then
                parts = Split(bucket, ",")
                If UBound(parts) >= 1 Then
                    windowStart = CDbl(parts(0))
                    If nowTs - windowStart > RL_CLEANUP_INTERVAL Then
                        Application.Contents.Remove appKey
                        Call RL_UnregisterKey(key)
                    End If
                End If
            End If
        End If
    Next
    
    Application.UnLock
End Sub

' ============================================
' RateLimitCheck(bucketName): api_guard.asp 兼容包装
' 使用默认 60 req/60s 窗口
' ============================================
Function RateLimitCheck(bucketName)
    RateLimitCheck = RL_Check(bucketName, RL_DEFAULT_MAX, RL_DEFAULT_WINDOW)
End Function

' ============================================
' RateLimitSend429(bucketName): api_guard.asp 兼容包装
' 发送标准 429 Too Many Requests 响应
' ============================================
Sub RateLimitSend429(bucketName)
    Dim remaining
    remaining = RL_GetRemaining(bucketName)
    If remaining < 0 Then remaining = 0
    Response.Status = "429 Too Many Requests"
    Response.AddHeader "Retry-After", RL_DEFAULT_WINDOW
    Response.ContentType = "application/json"
    Response.Write "{""error"":""Rate limit exceeded"",""code"":""RATE_LIMITED"",""bucket"":""" & bucketName & """,""retry_after"":" & RL_DEFAULT_WINDOW & "}"
End Sub

' ============================================
' 内部：定时清理（每 RL_CLEANUP_INTERVAL 秒触发一次）
' ============================================
Dim lastCleanup
lastCleanup = Application("RL_LastCleanup")
If IsEmpty(lastCleanup) Or Not IsNumeric(lastCleanup) Then lastCleanup = 0

Dim nowCleanupTs
nowCleanupTs = CDbl(DateDiff("s", CDate("1970-01-01 00:00:00"), Now()))
If nowCleanupTs - CDbl(lastCleanup) > RL_CLEANUP_INTERVAL Then
    Call RL_Cleanup()
    Application("RL_LastCleanup") = nowCleanupTs
End If
%>
