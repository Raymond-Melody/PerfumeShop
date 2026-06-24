<%
' ============================================
' V15.0 结构化日志系统 (Structured Logging)
' 依赖: connection.asp (可选，SQL日志需要conn)
' 用法: <!--#include file="logger.asp"-->
' 调用: LOG_INFO "用户登录成功"
'        LOG_ERROR "数据库连接失败", "connection.asp", 42
'        LOG_WARN "库存低于安全阈值", "inventory.asp", 0, "ProductID=123,StockQty=5"
' ============================================

' 日志级别常量
Const LOG_LEVEL_DEBUG = 0
Const LOG_LEVEL_INFO  = 1
Const LOG_LEVEL_WARN  = 2
Const LOG_LEVEL_ERROR = 3
Const LOG_LEVEL_FATAL = 4

' 日志配置
Const LOG_FILE_DIR = "/logs/"           ' 日志文件目录
Const LOG_MAX_DAYS = 30                 ' 日志保留天数
Const LOG_BUFFER_SIZE = 100             ' 内存缓冲区最大条数
Const LOG_DEFAULT_LEVEL = LOG_LEVEL_INFO ' 默认最低日志级别（低于此级别不记录）

' ============================================
' 内部函数：获取当前用户标识
' ============================================
Function LOG_GetCurrentUser()
    If Session("AdminID") <> "" And Not IsEmpty(Session("AdminID")) Then
        LOG_GetCurrentUser = "Admin:" & Session("AdminID")
    ElseIf Session("UserID") <> "" And Not IsEmpty(Session("UserID")) Then
        LOG_GetCurrentUser = "User:" & Session("UserID")
    Else
        LOG_GetCurrentUser = "Guest"
    End If
End Function

' ============================================
' 内部函数：获取当前页面名称
' ============================================
Function LOG_GetPageName()
    Dim path, parts
    path = Request.ServerVariables("SCRIPT_NAME")
    If path <> "" Then
        parts = Split(path, "/")
        If UBound(parts) >= 0 Then
            LOG_GetPageName = parts(UBound(parts))
        Else
            LOG_GetPageName = "unknown.asp"
        End If
    Else
        LOG_GetPageName = "unknown.asp"
    End If
End Function

' ============================================
' 内部函数：获取客户端IP
' ============================================
Function LOG_GetClientIP()
    Dim ip
    ip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
    If ip = "" Then ip = Request.ServerVariables("REMOTE_ADDR")
    If ip = "" Then ip = "0.0.0.0"
    LOG_GetClientIP = ip
End Function

' ============================================
' 内部函数：获取日志级别名称
' ============================================
Function LOG_GetLevelName(level)
    Select Case level
        Case LOG_LEVEL_DEBUG: LOG_GetLevelName = "DEBUG"
        Case LOG_LEVEL_INFO:  LOG_GetLevelName = "INFO"
        Case LOG_LEVEL_WARN:  LOG_GetLevelName = "WARN"
        Case LOG_LEVEL_ERROR: LOG_GetLevelName = "ERROR"
        Case LOG_LEVEL_FATAL: LOG_GetLevelName = "FATAL"
        Case Else:            LOG_GetLevelName = "UNKNOWN"
    End Select
End Function

' ============================================
' 内部函数：格式化时间戳
' ============================================
Function LOG_FormatTimestamp()
    Dim nowVal
    nowVal = Now()
    LOG_FormatTimestamp = Year(nowVal) & "-" & Right("0" & Month(nowVal), 2) & "-" & Right("0" & Day(nowVal), 2) & " " & _
                          Right("0" & Hour(nowVal), 2) & ":" & Right("0" & Minute(nowVal), 2) & ":" & Right("0" & Second(nowVal), 2)
End Function

' ============================================
' 内部函数：写入Application内存缓冲区
' ============================================
Sub LOG_WriteToBuffer(logEntry)
    On Error Resume Next
    Dim buffer
    
    If Not IsObject(Application("Logger_Buffer")) Then
        Application.Lock
        If Not IsObject(Application("Logger_Buffer")) Then
            Set Application("Logger_Buffer") = Server.CreateObject("Scripting.Dictionary")
        End If
        Application.Unlock
    End If
    
    Set buffer = Application("Logger_Buffer")
    If IsObject(buffer) Then
        Application.Lock
        ' 超过最大容量时移除最旧条目
        Do While buffer.Count >= LOG_BUFFER_SIZE
            Dim firstKey
            firstKey = buffer.Keys()(0)
            buffer.Remove firstKey
        Loop
        ' 使用时间戳+随机数作为唯一Key
        buffer.Add Timer() & "_" & buffer.Count, logEntry
        Application.Unlock
    End If
End Sub

' ============================================
' 内部函数：写入日志文件（按日轮转）
' ============================================
Sub LOG_WriteToFile(logEntry)
    Dim fso, logFile, logPath, logDir, fileName, dateStr
    On Error Resume Next
    
    dateStr = Year(Now()) & "-" & Right("0" & Month(Now()), 2) & "-" & Right("0" & Day(Now()), 2)
    fileName = "app_" & dateStr & ".log"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    
    logDir = Server.MapPath(LOG_FILE_DIR)
    
    ' 确保日志目录存在
    If Not fso.FolderExists(logDir) Then
        fso.CreateFolder(logDir)
        If Err.Number <> 0 Then
            Err.Clear
            Set fso = Nothing
            Exit Sub
        End If
    End If
    
    logPath = logDir & "\" & fileName
    
    ' 追加写入
    Set logFile = fso.OpenTextFile(logPath, 8, True) ' 8=ForAppending, True=Create
    If Err.Number = 0 Then
        logFile.WriteLine logEntry
        logFile.Close
    End If
    Err.Clear
    Set logFile = Nothing
    
    ' 后台清理过期日志（随机触发，约1/20概率）
    Dim rndVal
    Randomize
    rndVal = Int(Rnd * 20)
    If rndVal = 0 Then
        Call LOG_CleanOldFiles(fso, logDir)
    End If
    
    Set fso = Nothing
End Sub

' ============================================
' 内部函数：清理过期日志文件
' ============================================
Sub LOG_CleanOldFiles(fso, logDir)
    Dim folder, file, cutoffDate, fileName, fileDate, fileDateStr
    On Error Resume Next
    
    cutoffDate = DateAdd("d", -LOG_MAX_DAYS, Date())
    
    If fso.FolderExists(logDir) Then
        Set folder = fso.GetFolder(logDir)
        For Each file In folder.Files
            ' 文件名格式: app_YYYY-MM-DD.log
            fileName = file.Name
            If Left(fileName, 4) = "app_" And Right(fileName, 4) = ".log" Then
                fileDateStr = Mid(fileName, 5, 10) ' YYYY-MM-DD
                If IsDate(fileDateStr) Then
                    If CDate(fileDateStr) < cutoffDate Then
                        fso.DeleteFile file.Path, True
                    End If
                End If
            End If
        Next
        Set folder = Nothing
    End If
    Err.Clear
End Sub

' ============================================
' 内部函数：写入SQL日志表（可选，需conn可用）
' ============================================
Sub LOG_WriteToDB(level, message, source, lineNum)
    Dim sql, params(4)
    On Error Resume Next
    
    ' 检查conn是否可用
    If Not IsObject(conn) Then Exit Sub
    If conn.State <> 1 Then Exit Sub ' adStateOpen = 1
    
    sql = "INSERT INTO AppLogs (LogLevel, LogMessage, LogSource, LineNumber, " & _
          "UserName, IPAddress, PageURL, CreatedAt) " & _
          "VALUES (@Level, @Message, @Source, @LineNum, " & _
          "@UserName, @IP, @PageURL, GETDATE())"
    
    params(0) = Array("@Level", DAL_adVarChar, 10, LOG_GetLevelName(level))
    params(1) = Array("@Message", DAL_adVarChar, 500, Left(message, 500))
    params(2) = Array("@Source", DAL_adVarChar, 100, source)
    params(3) = Array("@LineNum", DAL_adInteger, 0, CLng(lineNum))
    params(4) = Array("@UserName", DAL_adVarChar, 100, LOG_GetCurrentUser())
    params(5) = Array("@IP", DAL_adVarChar, 50, LOG_GetClientIP())
    params(6) = Array("@PageURL", DAL_adVarChar, 200, Request.ServerVariables("SCRIPT_NAME"))
    
    DAL_Execute sql, params
    Err.Clear
End Sub

' ============================================
' 核心日志写入函数
' 参数:
'   level    - 日志级别 (LOG_LEVEL_DEBUG/INFO/WARN/ERROR/FATAL)
'   message  - 日志消息
'   source   - 来源文件名（可选，自动检测）
'   lineNum  - 行号（可选）
'   extra    - 额外信息（可选）
' ============================================
Sub LOG_Write(level, message, source, lineNum, extra)
    Dim logEntry, levelName, user, page, ip
    
    ' Feature Flag 检查
    If Not FEATURE_STRUCTURED_LOGGING And level >= LOG_LEVEL_INFO Then
        ' 即使关闭结构化日志，ERROR和FATAL也始终记录
        If level < LOG_LEVEL_ERROR Then Exit Sub
    End If
    
    ' 级别过滤
    If level < LOG_DEFAULT_LEVEL And level < LOG_LEVEL_ERROR Then Exit Sub
    
    ' 自动检测来源
    If IsNull(source) Or IsEmpty(source) Or source = "" Then
        source = LOG_GetPageName()
    End If
    If IsNull(lineNum) Or IsEmpty(lineNum) Then lineNum = 0
    
    levelName = LOG_GetLevelName(level)
    user = LOG_GetCurrentUser()
    ip = LOG_GetClientIP()
    
    ' 构建日志条目
    logEntry = "[" & LOG_FormatTimestamp() & "] [" & levelName & "] [" & user & "] [" & source
    If CLng(lineNum) > 0 Then logEntry = logEntry & ":" & lineNum
    logEntry = logEntry & "] " & message
    If Not IsNull(extra) And extra <> "" Then logEntry = logEntry & " | " & extra
    logEntry = logEntry & " | IP=" & ip
    
    ' 写入文件
    Call LOG_WriteToFile(logEntry)
    
    ' 写入内存缓冲区
    Call LOG_WriteToBuffer(logEntry)
    
    ' 写入SQL（ERROR和FATAL级别）
    If level >= LOG_LEVEL_ERROR And FEATURE_DAL_ENABLED Then
        Call LOG_WriteToDB(level, message, source, lineNum)
    End If
    
    ' FATAL特殊处理：写入Session
    If level >= LOG_LEVEL_ERROR Then
        Session("LastLogError") = logEntry
    End If
End Sub

' ============================================
' 便捷日志函数
' ============================================
Sub LOG_DEBUG(message)
    Call LOG_Write(LOG_LEVEL_DEBUG, message, "", 0, "")
End Sub

Sub LOG_INFO(message)
    Call LOG_Write(LOG_LEVEL_INFO, message, "", 0, "")
End Sub

Sub LOG_WARN(message)
    Call LOG_Write(LOG_LEVEL_WARN, message, "", 0, "")
End Sub

Sub LOG_ERROR(message)
    Call LOG_Write(LOG_LEVEL_ERROR, message, "", 0, "")
End Sub

Sub LOG_FATAL(message)
    Call LOG_Write(LOG_LEVEL_FATAL, message, "", 0, "")
End Sub

' ============================================
' 带来源的日志函数（推荐使用）
' ============================================
Sub LOG_DEBUG_EX(message, source, lineNum)
    Call LOG_Write(LOG_LEVEL_DEBUG, message, source, lineNum, "")
End Sub

Sub LOG_INFO_EX(message, source, lineNum)
    Call LOG_Write(LOG_LEVEL_INFO, message, source, lineNum, "")
End Sub

Sub LOG_WARN_EX(message, source, lineNum)
    Call LOG_Write(LOG_LEVEL_WARN, message, source, lineNum, "")
End Sub

Sub LOG_ERROR_EX(message, source, lineNum)
    Call LOG_Write(LOG_LEVEL_ERROR, message, source, lineNum, "")
End Sub

' ============================================
' 获取内存缓冲区最近N条日志
' ============================================
Function LOG_GetRecent(count)
    Dim buffer, result, keys, i, startIdx
    
    If IsNull(count) Or count < 1 Then count = 50
    
    Set result = Server.CreateObject("Scripting.Dictionary")
    
    If IsObject(Application("Logger_Buffer")) Then
        Set buffer = Application("Logger_Buffer")
        If buffer.Count > 0 Then
            keys = buffer.Keys()
            ' 从尾部开始取
            startIdx = buffer.Count - count
            If startIdx < 0 Then startIdx = 0
            
            For i = startIdx To buffer.Count - 1
                result.Add i, buffer.Item(keys(i))
            Next
        End If
    End If
    
    Set LOG_GetRecent = result
End Function

' ============================================
' 清空内存缓冲区
' ============================================
Sub LOG_ClearBuffer()
    Application.Lock
    If IsObject(Application("Logger_Buffer")) Then
        Application("Logger_Buffer").RemoveAll
    End If
    Application.Unlock
End Sub

' ============================================
' 获取日志统计信息
' ============================================
Function LOG_GetStats()
    Dim stats, buffer
    Set stats = Server.CreateObject("Scripting.Dictionary")
    
    stats.Add "bufferSize", 0
    stats.Add "logDir", Server.MapPath(LOG_FILE_DIR)
    stats.Add "maxDays", LOG_MAX_DAYS
    stats.Add "enabled", FEATURE_STRUCTURED_LOGGING
    
    If IsObject(Application("Logger_Buffer")) Then
        Set buffer = Application("Logger_Buffer")
        stats("bufferSize") = buffer.Count
    End If
    
    Set LOG_GetStats = stats
End Function

' ============================================
' 初始化：记录应用启动日志
' ============================================
If Not Application("Logger_Initialized") Then
    Application.Lock
    If Not Application("Logger_Initialized") Then
        Call LOG_WriteToBuffer("[STARTUP] Logger initialized at " & LOG_FormatTimestamp())
        Application("Logger_Initialized") = True
    End If
    Application.Unlock
End If
%>