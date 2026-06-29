$fpath = 'f:\网站制作\网站\网站二\api\health_check.asp'
$enc = [System.Text.UTF8Encoding]::new($true)
$txt = [System.IO.File]::ReadAllText($fpath, $enc)

$insert = @'

' V18: 内存使用率
Dim memUsageMB, memStatus
memUsageMB = 0
memStatus = "unknown"
On Error Resume Next
Dim wmi, memCol, memItem
Set wmi = GetObject("winmgmts:\\.\root\cimv2")
If Err.Number = 0 Then
    Set memCol = wmi.ExecQuery("SELECT TotalVisibleMemorySize,FreePhysicalMemory FROM Win32_OperatingSystem")
    For Each memItem In memCol
        Dim totalMem, freeMem
        totalMem = CDbl(memItem.TotalVisibleMemorySize) / 1024
        freeMem = CDbl(memItem.FreePhysicalMemory) / 1024
        memUsageMB = Round(totalMem - freeMem, 1)
        If freeMem / totalMem < 0.1 Then memStatus = "critical" Else If freeMem / totalMem < 0.2 Then memStatus = "warn" Else memStatus = "ok"
    Next
    Set memCol = Nothing
End If
Set wmi = Nothing
On Error GoTo 0

' V18: 慢查询统计
Dim slowQueryCount, slowQueryAvgMs
slowQueryCount = 0
slowQueryAvgMs = 0
On Error Resume Next
slowQueryCount = CLng(Application("DAL_SlowQueryCount"))
If slowQueryCount > 0 Then
    slowQueryAvgMs = Round(CDbl(Application("DAL_SlowQueryTotalMs")) / slowQueryCount, 1)
End If
On Error GoTo 0

' V18: 缓存统计
Dim cacheHits, cacheMisses, cacheHitRate
cacheHits = 0
cacheMisses = 0
cacheHitRate = 0
On Error Resume Next
If IsObject(Application("CM_Stats")) Then
    cacheHits = CLng(Application("CM_Stats")("hits"))
    cacheMisses = CLng(Application("CM_Stats")("misses"))
    If cacheHits + cacheMisses > 0 Then
        cacheHitRate = Round(cacheHits / (cacheHits + cacheMisses) * 100, 1)
    End If
End If
On Error GoTo 0

' V18: AppLog 错误计数（最近24小时）
Dim errorCount24h
errorCount24h = 0
On Error Resume Next
If dbStatus <> "fail" Then
    Dim diagConn2, diagRs2
    Set diagConn2 = Server.CreateObject("ADODB.Connection")
    diagConn2.Open GetConnectionString()
    If Err.Number = 0 Then
        Set diagRs2 = diagConn2.Execute("SELECT COUNT(*) FROM AppLogs WHERE LogType='ERROR' AND CreatedAt > DATEADD(HOUR, -24, GETDATE())")
        If Not diagRs2 Is Nothing And Not diagRs2.EOF Then
            errorCount24h = CLng(diagRs2(0))
            diagRs2.Close
        End If
        Set diagRs2 = Nothing
        diagConn2.Close
    End If
    Set diagConn2 = Nothing
End If
Err.Clear
On Error GoTo 0

'@

$oldEnd = 'Response.Write """disk"": {"
Response.Write """freeMB"": " & diskFreeMB & ","
Response.Write """status"": """ & diskStatus & """"
Response.Write "},"'
$newEnd = 'Response.Write """disk"": {"
Response.Write """freeMB"": " & diskFreeMB & ","
Response.Write """status"": """ & diskStatus & """"
Response.Write "},"
Response.Write """memory"": {"
Response.Write """usageMB"": " & memUsageMB & ","
Response.Write """status"": """ & memStatus & """"
Response.Write "},"
Response.Write """slowQueries"": {"
Response.Write """count"": " & slowQueryCount & ","
Response.Write """avgMs"": " & slowQueryAvgMs
Response.Write "},"
Response.Write """cache"": {"
Response.Write """hits"": " & cacheHits & ","
Response.Write """misses"": " & cacheMisses & ","
Response.Write """hitRate"": " & cacheHitRate
Response.Write "},"
Response.Write """errors24h"": " & errorCount24h & ","'

$txt = $txt.Replace($oldEnd, $insert + $newEnd)

[System.IO.File]::WriteAllText($fpath, $txt, $enc)
Write-Host 'health_check.asp updated OK'
