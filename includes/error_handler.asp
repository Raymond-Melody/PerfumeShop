<%
' ============================================
' 统一错误处理模块 - Error Handler
' 功能：全局错误捕获、日志记录、友好错误展示
' 用法：在页面开端包含此文件，使用 EH_HandleError 替换 On Error Resume Next
' ============================================

' 错误日志文件路径
Const EH_LOG_PATH = "/logs/error_log.txt"
Const EH_DEBUG_MODE = True  ' 开发模式显示详细错误，生产环境设为 False

' ============================================
' 安全获取当前管理员（用于日志记录）
' ============================================
Function EH_GetCurrentUser()
    Dim user
    user = ""
    If Session("AdminID") <> "" And Not IsEmpty(Session("AdminID")) Then
        user = "Admin:" & Session("AdminID")
    ElseIf Session("UserID") <> "" And Not IsEmpty(Session("UserID")) Then
        user = "User:" & Session("UserID")
    Else
        user = "Guest"
    End If
    EH_GetCurrentUser = user
End Function

' ============================================
' 写入错误日志文件
' ============================================
Sub EH_WriteLog(errSource, errDesc, errNum, errLine, additionalInfo)
    Dim fso, logFile, logPath, logEntry
    On Error Resume Next
    
    ' 构建日志条目
    logEntry = "[" & Now() & "] User=" & EH_GetCurrentUser() & " IP=" & Request.ServerVariables("REMOTE_ADDR")
    logEntry = logEntry & " | Source=" & errSource
    logEntry = logEntry & " | Err#" & errNum & ": " & errDesc
    If errLine <> "" Then logEntry = logEntry & " | Line=" & errLine
    If additionalInfo <> "" Then logEntry = logEntry & " | Info=" & additionalInfo
    
    ' 同时写入Session供后续调试
    Session("EH_LastError") = logEntry
    
    ' 写入日志文件
    logPath = Server.MapPath(EH_LOG_PATH)
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 确保日志目录存在
    Dim logDir
    logDir = Left(logPath, InStrRev(logPath, "\") - 1)
    If Not fso.FolderExists(logDir) Then
        fso.CreateFolder(logDir)
    End If
    
    ' 追加写入（日志文件最大1MB，超过则截断）
    Dim maxSize
    maxSize = 1048576 ' 1MB
    If fso.FileExists(logPath) Then
        If fso.GetFile(logPath).Size > maxSize Then
            ' 重命名为备份
            fso.CopyFile logPath, logPath & "." & Replace(Replace(Replace(Now(), ":", ""), "/", ""), " ", "_") & ".bak"
            fso.DeleteFile logPath, True
        End If
    End If
    
    Set logFile = fso.OpenTextFile(logPath, 8, True) ' 8=ForAppending, True=Create
    logFile.WriteLine logEntry
    logFile.Close
    Set logFile = Nothing
    Set fso = Nothing
End Sub

' ============================================
' 统一错误处理 - 替换 On Error Resume Next 模式
' ============================================
Sub EH_HandleError(errSource, additionalInfo)
    If Err.Number <> 0 Then
        Dim errDesc, errNum
        errDesc = Err.Description
        errNum = Err.Number
        
        ' 记录日志
        Call EH_WriteLog(errSource, errDesc, errNum, "", additionalInfo)
        
        ' 保存到 Session 供 EH_DisplayError 使用
        Session("EH_LastErrorSource") = errSource
        Session("EH_LastErrorDesc") = errDesc
        Session("EH_LastErrorNum") = errNum
        
        ' 清理
        Err.Clear
    End If
End Sub

' ============================================
' 安全执行数据库查询（带错误处理）
' ============================================
Function EH_ExecuteQuery(sql, context)
    Dim rs
    On Error Resume Next
    Set rs = conn.Execute(sql)
    If Err.Number <> 0 Then
        Call EH_WriteLog("EH_ExecuteQuery(" & context & ")", Err.Description, Err.Number, "", "SQL: " & Left(sql, 500))
        Set EH_ExecuteQuery = Nothing
        Err.Clear
    Else
        Set EH_ExecuteQuery = rs
    End If
End Function

' ============================================
' 安全获取标量值（带错误处理）
' ============================================
Function EH_GetScalar(sql, context, defaultValue)
    Dim rs, val
    val = defaultValue
    On Error Resume Next
    Set rs = conn.Execute(sql)
    If Err.Number <> 0 Then
        Call EH_WriteLog("EH_GetScalar(" & context & ")", Err.Description, Err.Number, "", "SQL: " & Left(sql, 300))
        Err.Clear
    ElseIf Not rs Is Nothing Then
        If Not rs.EOF Then
            val = rs(0)
            If IsNull(val) Then val = defaultValue
        End If
        rs.Close
    End If
    Set rs = Nothing
    EH_GetScalar = val
End Function

' ============================================
' 展示用户友好的错误信息
' ============================================
Sub EH_DisplayError(title, details)
    If EH_DEBUG_MODE Then
        ' 开发模式：显示详细错误
        Response.Write "<div style='background:#2d2d44;border:1px solid #e57373;border-radius:8px;padding:20px;margin:20px;color:#e0e0e0;'>"
        Response.Write "<h3 style='color:#e57373;margin:0 0 10px;'><i class='fas fa-exclamation-triangle'></i> " & Server.HTMLEncode(title) & "</h3>"
        Response.Write "<p style='color:#aaa;font-size:13px;'>" & Server.HTMLEncode(details) & "</p>"
        
        ' 显示最近错误
        If Session("EH_LastErrorDesc") <> "" And Not IsEmpty(Session("EH_LastErrorDesc")) Then
            Response.Write "<div style='background:#1a1a2e;border-radius:6px;padding:10px;margin-top:10px;font-size:12px;'>"
            Response.Write "<strong style='color:#e57373;'>错误详情:</strong><br>"
            Response.Write "来源: " & Server.HTMLEncode(Session("EH_LastErrorSource")) & "<br>"
            Response.Write "描述: " & Server.HTMLEncode(Session("EH_LastErrorDesc")) & "<br>"
            Response.Write "编号: " & Server.HTMLEncode(Session("EH_LastErrorNum"))
            Response.Write "</div>"
        End If
        
        Response.Write "</div>"
    Else
        ' 生产模式：显示通用信息
        Response.Write "<div style='background:#fff3cd;border:1px solid #ffc107;border-radius:8px;padding:20px;margin:20px;'>"
        Response.Write "<h3 style='color:#856404;margin:0 0 10px;'>系统提示</h3>"
        Response.Write "<p style='color:#856404;'>" & Server.HTMLEncode(title) & "，请稍后重试或联系管理员。</p>"
        Response.Write "</div>"
    End If
End Sub

' ============================================
' 页面启动初始化 - 放于页面顶部
' ============================================
Sub EH_PageInit(pageName)
    ' 重置错误Session（每次页面加载时重置旧错误记录）
    Session("EH_LastError") = ""
    
    ' 慢查询监控开始时间
    Session("EH_PageStartTime") = Timer()
    Session("EH_PageName") = pageName
    
    ' 设置脚本超时（大报表查询可能需要更长时间）
    Server.ScriptTimeout = 90
End Sub

' ============================================
' 页面结束统计 - 放于页面底部
' ============================================
Sub EH_PageComplete()
    ' 计算页面执行时间（仅调试模式）
    If EH_DEBUG_MODE Then
        Dim startTime, elapsed
        startTime = Session("EH_PageStartTime")
        If startTime <> "" And IsNumeric(startTime) Then
            elapsed = Round((Timer() - CDbl(startTime)) * 1000, 1)
            If elapsed > 2000 Then ' 超过2秒记录慢查询
                Call EH_WriteLog("SlowPage", "Page execution time > 2s", 0, "", Session("EH_PageName") & " took " & elapsed & "ms")
            End If
            ' 页面底部显示（可选调试信息）
            ' Response.Write "<!-- Page loaded in " & elapsed & "ms -->"
        End If
    End If
End Sub

' ============================================
' 数据库操作事务包装器
' ============================================
Sub EH_BeginTransaction(context)
    On Error Resume Next
    conn.BeginTrans
    If Err.Number <> 0 Then
        Call EH_WriteLog("EH_BeginTransaction(" & context & ")", Err.Description, Err.Number, "", "")
        Err.Clear
    End If
End Sub

Sub EH_CommitTransaction(context)
    On Error Resume Next
    conn.CommitTrans
    If Err.Number <> 0 Then
        Call EH_WriteLog("EH_CommitTransaction(" & context & ")", Err.Description, Err.Number, "", "")
        Err.Clear
    End If
End Sub

Sub EH_RollbackTransaction(context)
    On Error Resume Next
    conn.RollbackTrans
    If Err.Number <> 0 Then
        Call EH_WriteLog("EH_RollbackTransaction(" & context & ")", Err.Description, Err.Number, "", "")
        Err.Clear
    End If
End Sub

' ============================================
' 日志浏览页面（管理员调试用）
' 访问 /logs/view_errors.asp 查看
' ============================================
' 此函数不在 include 中直接输出，通过独立页面调用
Sub EH_GetRecentLogs(maxLines)
    Dim fso, logFile, logPath, lines, i
    lines = Array()
    On Error Resume Next
    
    logPath = Server.MapPath(EH_LOG_PATH)
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(logPath) Then
        Set logFile = fso.OpenTextFile(logPath, 1) ' 1=ForReading
        Do While Not logFile.AtEndOfStream
            ReDim Preserve lines(UBound(lines) + 1)
            lines(UBound(lines)) = logFile.ReadLine
            If UBound(lines) > maxLines * 2 Then
                ' 只保留最后 maxLines 行
                Dim shifted
                shifted = Array()
                For i = UBound(lines) - maxLines To UBound(lines)
                    ReDim Preserve shifted(UBound(shifted) + 1)
                    shifted(UBound(shifted)) = lines(i)
                Next
                lines = shifted
            End If
        Loop
        logFile.Close
    End If
    
    Set logFile = Nothing
    Set fso = Nothing
    EH_GetRecentLogs = lines
End Function
%>
