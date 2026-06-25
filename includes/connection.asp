<% 
' ============================================ 
' 数据库连接模块 - SQL Server版本 (V17.0) 
' 支持 SQLOLEDB (旧) 和 MSOLEDBSQL (新) 双Provider 
' V17: 添加连接健康检查、自动重试机制、连接池自动重置 
' 通过 config.asp 中 FEATURE_MSOLEDBSQL 开关切换 
' ============================================ 

Dim conn
Const CONN_MAX_RETRIES = 3          ' V17: 最大连接重试次数
Const CONN_RETRY_DELAY_MS = 500     ' V17: 重试间隔毫秒

' ============================================ 
' V15: 构建数据库连接字符串 
' 支持双Provider，通过Feature Flag切换 
' ============================================ 
Function BuildConnectionString() 
    Dim serverName, dbName 
    serverName = "localhost\YOURPERFUME" 
    dbName = "PerfumeShop" 
    
    If FEATURE_MSOLEDBSQL Then 
        ' V17: Microsoft OLE DB Driver for SQL Server (推荐) 
        ' 需安装: https://aka.ms/downloadmsoledbsql 
        ' 优势: TLS 1.2加密、更好的UTF-8支持、持续维护 
        BuildConnectionString = "Provider=MSOLEDBSQL;Server=" & serverName & ";Database=" & dbName & ";Integrated Security=SSPI;TrustServerCertificate=yes;" 
    Else 
        ' V14.6兼容: SQLOLEDB (已弃用但仍可用) 
        BuildConnectionString = "Provider=SQLOLEDB;Server=" & serverName & ";Database=" & dbName & ";Integrated Security=SSPI;" 
    End If 
End Function 

' V17: 增强版打开数据库连接（含健康检查和自动重试） 
Sub OpenConnection() 
    Dim connStr, retryCount, connected 
    retryCount = 0 
    connected = False 
    
    Do While Not connected And retryCount <= CONN_MAX_RETRIES 
        On Error Resume Next 
        
        If retryCount > 0 Then 
            ' 重试前等待 
            Dim waitUntil : waitUntil = Timer() + (CONN_RETRY_DELAY_MS / 1000) 
            Do While Timer() < waitUntil : Loop 
        End If 
        
        Err.Clear 
        
        ' 每次重试重新创建连接对象 
        If Not conn Is Nothing Then 
            If conn.State = 1 Then conn.Close 
            Set conn = Nothing 
        End If 
        Set conn = Server.CreateObject("ADODB.Connection") 
        
        If Err.Number = 0 And Not conn Is Nothing Then 
            ' 根据Feature Flag选择连接字符串 
            connStr = BuildConnectionString() 
            conn.Open connStr 
            
            If Err.Number = 0 Then 
                ' 连接成功，执行健康检查 
                Dim verifyRs 
                Set verifyRs = conn.Execute("SELECT 1 AS health_check") 
                If Err.Number = 0 And Not verifyRs Is Nothing Then 
                    verifyRs.Close : Set verifyRs = Nothing 
                    connected = True 
                    If retryCount > 0 Then 
                        Session("DBConnectInfo") = "数据库连接成功 (重试" & retryCount & "次)" 
                    End If 
                Else 
                    Err.Clear 
                End If 
                If Not verifyRs Is Nothing Then 
                    verifyRs.Close : Set verifyRs = Nothing 
                End If 
            ElseIf FEATURE_MSOLEDBSQL Then 
                ' MSOLEDBSQL失败时尝试回退到SQLOLEDB 
                Dim fallbackErr : fallbackErr = Err.Description 
                Err.Clear 
                Set conn = Nothing 
                Set conn = Server.CreateObject("ADODB.Connection") 
                conn.Open "Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;" 
                If Err.Number = 0 Then 
                    Dim fbRs : Set fbRs = conn.Execute("SELECT 1 AS health_check") 
                    If Err.Number = 0 Then 
                        fbRs.Close : Set fbRs = Nothing 
                        connected = True 
                        Session("DBFallbackNotice") = "MSOLEDBSQL不可用，已回退到SQLOLEDB。错误: " & Left(fallbackErr, 200) 
                    Else 
                        Err.Clear 
                        If Not fbRs Is Nothing Then fbRs.Close : Set fbRs = Nothing 
                    End If 
                Else 
                    Err.Clear 
                End If 
                If conn Is Nothing Or conn.State <> 1 Then 
                    Set conn = Nothing 
                End If 
            Else 
                Err.Clear 
                Set conn = Nothing 
            End If 
        Else 
            Err.Clear 
            Set conn = Nothing 
        End If 
        
        On Error GoTo 0 
        retryCount = retryCount + 1 
    Loop 
    
    ' V17: 连接成功时按概率记录心跳日志
    If connected Then
        Randomize
        If Int(Rnd * 50) = 0 Then
            Call LogHeartbeat()
        End If
    End If

    ' 如果所有重试都失败 
    If Not connected Then 
        Session("DBLastError") = "数据库连接失败 (已重试" & retryCount & "次)" 
        Response.Write "<div class='error'>数据库连接失败，请稍后重试。" & _ 
                      "错误代码: DB-CONN-001</div>" 
        Response.End 
    End If 
End Sub

' V17: 增强版关闭数据库连接 
Sub CloseConnection() 
    On Error Resume Next 
    If IsObject(conn) Then 
        If Not conn Is Nothing Then 
            If conn.State = 1 Then 
                conn.Close 
            End If 
        End If 
        Set conn = Nothing 
    End If 
    Err.Clear 
End Sub

' 执行查询并返回记录集 - 修复参数绑定问题
Function ExecuteQuery(sql)
    Dim rs
    On Error Resume Next
    Set rs = Server.CreateObject("ADODB.Recordset")
    
    If Err.Number <> 0 Or rs Is Nothing Then
        Session("LastDBError") = "无法创建记录集对象: " & Err.Description & " 错误号: " & Err.Number
        Set ExecuteQuery = Nothing
        Exit Function
    End If
    
    ' 直接使用conn和SQL字符串，避免可能的参数绑定问题
    rs.CursorLocation = 3  ' adUseClient - 支持 MoveLast/RecordCount
    rs.CursorType = 1  ' adOpenKeyset
    rs.LockType = 1    ' adLockOptimistic
    rs.Open sql, conn  ' 简化调用方式
    
    If Err.Number <> 0 Then
        ' 记录错误到Session而不是输出到响应
        Session("LastDBError") = "查询错误: " & Err.Description & " 错误号: " & Err.Number & " SQL: " & sql
        If Not rs Is Nothing Then
            If rs.State = 1 Then rs.Close
            Set rs = Nothing
        End If
        Set ExecuteQuery = Nothing
    Else
        Set ExecuteQuery = rs
    End If
End Function

' 执行非查询SQL（INSERT, UPDATE, DELETE）
Function ExecuteNonQuery(sql)
    On Error Resume Next
    conn.Execute sql
    If Err.Number <> 0 Then
        ' 保存错误信息供调试
        Session("LastDBError") = "ExecuteNonQuery错误: " & Err.Description & " 错误号: " & Err.Number & " | SQL: " & sql
        ExecuteNonQuery = False
        Err.Clear  ' 清除错误状态，防止残留影响后续操作
    Else
        Session("LastDBError") = ""
        ExecuteNonQuery = True
    End If
End Function

' ============================================
' V10.1: SafeSelect - 安全查询包装器
' 自动检查表是否存在，防止无效查询导致运行时错误
' 用法: Set rs = SafeSelect("Products", "ProductID, ProductName", "IsActive=1", "ProductName ASC")
' ============================================
Function SafeSelect(tableName, columns, whereClause, orderBy)
    Dim rs, sql, tableCheck
    
    ' 默认参数
    If IsNull(columns) Or columns = "" Then columns = "*"
    If IsNull(whereClause) Or whereClause = "" Then whereClause = "1=1"
    If IsNull(orderBy) Or orderBy = "" Then orderBy = ""
    
    On Error Resume Next
    
    ' 检查表是否存在
    Set tableCheck = conn.Execute("SELECT 1 FROM sys.tables WHERE name='" & Replace(tableName,"'","''") & "'")
    If Err.Number <> 0 Then
        Err.Clear
        Set SafeSelect = Nothing
        Exit Function
    End If
    
    If tableCheck Is Nothing Or tableCheck.EOF Then
        Session("LastDBError") = "SafeSelect: 表 '" & Server.HTMLEncode(tableName) & "' 不存在"
        If Not tableCheck Is Nothing Then tableCheck.Close : Set tableCheck = Nothing
        Set SafeSelect = Nothing
        Exit Function
    End If
    tableCheck.Close : Set tableCheck = Nothing
    
    ' 构建SQL
    sql = "SELECT " & columns & " FROM [" & tableName & "] WHERE " & whereClause
    If orderBy <> "" Then sql = sql & " ORDER BY " & orderBy
    
    ' 执行查询
    Set rs = Server.CreateObject("ADODB.Recordset")
    If Err.Number <> 0 Or rs Is Nothing Then
        Err.Clear
        Set SafeSelect = Nothing
        Exit Function
    End If
    
    rs.CursorLocation = 3
    rs.Open sql, conn, 1, 1
    
    If Err.Number <> 0 Then
        Session("LastDBError") = "SafeSelect错误: " & Err.Description & " | SQL: " & sql
        If Not rs Is Nothing Then
            If rs.State = 1 Then rs.Close
            Set rs = Nothing
        End If
        Err.Clear
        Set SafeSelect = Nothing
    Else
        Set SafeSelect = rs
    End If
    
    On Error GoTo 0
End Function

' ============================================
' 事务控制函数
' ============================================

' 开始事务
Sub BeginTransaction()
    On Error Resume Next
    conn.BeginTrans
End Sub

' 提交事务
Sub CommitTransaction()
    On Error Resume Next
    conn.CommitTrans
End Sub

' 回滚事务
Sub RollbackTransaction()
    On Error Resume Next
    conn.RollbackTrans
End Sub

' 获取单个值
Function GetScalar(sql)
    Dim rs, result
    On Error Resume Next
    Set rs = ExecuteQuery(sql)
    If Not rs Is Nothing And IsObject(rs) Then
        If Not rs.EOF And Not rs.BOF Then
            result = rs.Fields(0).Value
            If IsNull(result) Then result = "0"
        Else
            result = "0"
        End If
        rs.Close
        Set rs = Nothing
    Else
        result = "0"
    End If
    On Error GoTo 0
    GetScalar = result
End Function

' 获取最后插入的ID (SQL Server版本)
Function GetLastInsertID(tableName)
    Dim result
    result = GetScalar("SELECT SCOPE_IDENTITY()")
    If IsNull(result) Or result = "" Or result = "0" Then
        result = 0
    End If
    GetLastInsertID = CLng(result)
End Function

' 安全处理SQL字符串，防止SQL注入
Function SafeSQL(str)
    If IsNull(str) Or str = "" Then
        SafeSQL = ""
    Else
        SafeSQL = Replace(str, "'", "''")
    End If
End Function

' 安全转换为数字
Function SafeNum(val)
    If IsNull(val) Or val = "" Then
        SafeNum = 0
    Else
        On Error Resume Next
        SafeNum = CDbl(val)
        If Err.Number <> 0 Then
            Err.Clear
            SafeNum = 0
        End If
    End If
End Function

' HTML编码，防止XSS
Function HTMLEncode(str)
    If IsNull(str) Or str = "" Then
        HTMLEncode = ""
    Else
        HTMLEncode = Server.HTMLEncode(str)
    End If
End Function

' 安全输出字符串到HTML，处理引号等特殊字符
Function SafeOutput(str)
    If IsNull(str) Or str = "" Then
        SafeOutput = ""
    Else
        Dim tempStr
        tempStr = CStr(str)
        ' 转义换行符为JS安全转义序列（必须在HTMLEncode之前）
        tempStr = Replace(tempStr, Chr(13) & Chr(10), "\n")
        tempStr = Replace(tempStr, Chr(10), "\n")
        tempStr = Replace(tempStr, Chr(13), "\r")
        tempStr = Server.HTMLEncode(tempStr)
        ' 转换单引号为HTML实体
        tempStr = Replace(tempStr, "'", "&#39;")
        ' 转换双引号为HTML实体
        tempStr = Replace(tempStr, """", "&quot;")
        SafeOutput = tempStr
    End If
End Function

' 生成唯一订单号
Function GenerateOrderNo()
    Dim orderNo
    Randomize
    orderNo = ORDER_PREFIX & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2)
    orderNo = orderNo & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
    orderNo = orderNo & Right("000" & Int(Rnd * 1000), 3)
    GenerateOrderNo = orderNo
End Function

' 格式化货币（兼容SQL Server Decimal类型）
Function FormatMoney(amount)
    On Error Resume Next
    Dim dblAmount
    If IsNull(amount) Or IsEmpty(amount) Then
        FormatMoney = "¥0.00"
        Exit Function
    End If
    dblAmount = CDbl(amount)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        FormatMoney = "¥0.00"
        Exit Function
    End If
    On Error GoTo 0
    FormatMoney = "¥" & FormatNumber(dblAmount, 2)
End Function

' ============================================
' V17: 数据库心跳检测日志
' 写入 AppLogs 表，用于监控数据库连接可用性
' ============================================
Sub LogHeartbeat()
    Dim sql, params()
    On Error Resume Next
    
    ' 确保 conn 有效
    If conn Is Nothing Then Exit Sub
    If conn.State <> 1 Then Exit Sub
    
    sql = "INSERT INTO AppLogs (LogLevel, LogMessage, LogSource, IPAddress, PageURL) " & _
          "VALUES (@LogLevel, @LogMessage, @LogSource, @IPAddress, @PageURL)"
    
    ReDim params(4)
    params(0) = Array("@LogLevel", DAL_adVarChar, 10, "INFO")
    params(1) = Array("@LogMessage", DAL_adVarChar, 500, "DB-Heartbeat: 连接正常")
    params(2) = Array("@LogSource", DAL_adVarChar, 100, "connection.asp")
    params(3) = Array("@IPAddress", DAL_adVarChar, 50, Left(Request.ServerVariables("REMOTE_ADDR"), 50))
    params(4) = Array("@PageURL", DAL_adVarChar, 200, Left(Request.ServerVariables("SCRIPT_NAME"), 200))
    
    ' 尝试写入心跳日志（忽略失败，不影响主流程）
    Dim cmd
    Set cmd = DAL_CreateCommand(sql, params)
    If Not cmd Is Nothing Then
        cmd.Execute
    End If
    Set cmd = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

' 获取当前日期时间
Function GetNow()
    GetNow = Now()
End Function

' IIF函数（VBScript兼容）
Function IIF(condition, trueValue, falseValue)
    On Error Resume Next
    If condition Then
        IIF = trueValue
    Else
        IIF = falseValue
    End If
End Function

' 日期字段格式化函数（处理NULL值）
Function FormatDateField(dateValue)
    On Error Resume Next
    
    If IsNull(dateValue) Or dateValue = "" Then
        FormatDateField = "-"
    Else
        Dim formattedDate
        formattedDate = CStr(dateValue)
        ' 简格式：只取前10个字符（成为 YYYY-MM-DD 格式）
        If Len(formattedDate) >= 10 Then
            FormatDateField = Left(formattedDate, 10)
        Else
            FormatDateField = formattedDate
        End If
    End If
End Function

' 安全格式化日期时间（处理NULL值）
Function SafeFormatDateTime(dateValue, formatType)
    On Error Resume Next
    
    If IsNull(dateValue) Or dateValue = "" Or IsEmpty(dateValue) Then
        SafeFormatDateTime = "-"
        Exit Function
    End If
    
    ' Handle DATETIME2(7) - strip fractional seconds for VBScript IsDate/CDate
    Dim dtStr
    dtStr = CStr(dateValue & "")
    Dim dotPos
    dotPos = InStr(dtStr, ".")
    If dotPos > 0 Then
        dtStr = Left(dtStr, dotPos - 1)
    End If
    
    If Not IsDate(dtStr) Then
        SafeFormatDateTime = "-"
    Else
        SafeFormatDateTime = FormatDateTime(dtStr, formatType)
    End If
    
    If Err.Number <> 0 Then
        SafeFormatDateTime = "-"
        Err.Clear
    End If
End Function

' ============================================
' 安全工具函数模块
' ============================================

' Cookie签名密钥 (V10 增强长度) - 已移至 config.asp
' Const COOKIE_SECRET = "PerfumeShop_SecKey_2026_X9K3m"
' Const COOKIE_SECRET_V10 = "PF_V10_SHA256_Salt_7kM2xP9qR4vN8wL3jH6fD1sA5gK0"

' ============================================
' V10: SHA-256 纯VBScript实现
' 用于安全令牌签名，替换旧的DJB2 SimpleHash
' ============================================

' 32位无符号右旋转
Function ROTR32(n, k)
    Dim mask : mask = (2 ^ k) - 1
    ROTR32 = ((n And &H7FFFFFFF) \ (2 ^ (32 - k))) Or _
             (((n And mask) * (2 ^ (32 - k))) And &H7FFFFFFF)
    If (n And &H80000000) <> 0 Then
        If k > 0 Then
            ROTR32 = ROTR32 Or (&H80000000 \ (2 ^ (k - 1)))
        End If
    End If
End Function

' 32位无符号加法
Function AddU32(a, b)
    AddU32 = CLng((CDbl(a And &H7FFFFFFF) + CDbl(b And &H7FFFFFFF)) Xor _
                   ((a And &H80000000) Xor (b And &H80000000)))
End Function

' SHA-256 常量 K[0..63]
Function SHA256_K(i)
    Dim k(63)
    k(0)  = &H428A2F98 : k(1)  = &H71374491 : k(2)  = &HB5C0FBCF : k(3)  = &HE9B5DBA5
    k(4)  = &H3956C25B : k(5)  = &H59F111F1 : k(6)  = &H923F82A4 : k(7)  = &HAB1C5ED5
    k(8)  = &HD807AA98 : k(9)  = &H12835B01 : k(10) = &H243185BE : k(11) = &H550C7DC3
    k(12) = &H72BE5D74 : k(13) = &H80DEB1FE : k(14) = &H9BDC06A7 : k(15) = &HC19BF174
    k(16) = &HE49B69C1 : k(17) = &HEFBE4786 : k(18) = &H0FC19DC6 : k(19) = &H240CA1CC
    k(20) = &H2DE92C6F : k(21) = &H4A7484AA : k(22) = &H5CB0A9DC : k(23) = &H76F988DA
    k(24) = &H983E5152 : k(25) = &HA831C66D : k(26) = &HB00327C8 : k(27) = &HBF597FC7
    k(28) = &HC6E00BF3 : k(29) = &HD5A79147 : k(30) = &H06CA6351 : k(31) = &H14292967
    k(32) = &H27B70A85 : k(33) = &H2E1B2138 : k(34) = &H4D2C6DFC : k(35) = &H53380D13
    k(36) = &H650A7354 : k(37) = &H766A0ABB : k(38) = &H81C2C92E : k(39) = &H92722C85
    k(40) = &HA2BFE8A1 : k(41) = &HA81A664B : k(42) = &HC24B8B70 : k(43) = &HC76C51A3
    k(44) = &HD192E819 : k(45) = &HD6990624 : k(46) = &HF40E3585 : k(47) = &H106AA070
    k(48) = &H19A4C116 : k(49) = &H1E376C08 : k(50) = &H2748774C : k(51) = &H34B0BCB5
    k(52) = &H391C0CB3 : k(53) = &H4ED8AA4A : k(54) = &H5B9CCA4F : k(55) = &H682E6FF3
    k(56) = &H748F82EE : k(57) = &H78A5636F : k(58) = &H84C87814 : k(59) = &H8CC70208
    k(60) = &H90BEFFFA : k(61) = &HA4506CEB : k(62) = &HBEF9A3F7 : k(63) = &HC67178F2
    SHA256_K = k(i)
End Function

' SHA-256 主函数
Function SHA256Hash(str)
    Dim msg(), padded(), i, j, msgLen, paddedLen, numBlocks
    Dim H0, H1, H2, H3, H4, H5, H6, H7
    Dim a, b, c, d, e, f, g, h
    Dim W(63), t1, t2, S0, S1, ch, maj, blockIdx, bOffset
    
    If IsNull(str) Or str = "" Then
        SHA256Hash = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
        Exit Function
    End If
    
    ' 1. 准备消息字节数组
    msgLen = Len(str)
    ' 计算填充后的长度（512位块对齐）
    paddedLen = ((msgLen + 8) \ 64 + 1) * 64
    If (msgLen + 8) Mod 64 = 0 Then paddedLen = paddedLen + 64
    ReDim padded(paddedLen - 1)
    
    ' 复制消息字节
    For i = 0 To msgLen - 1
        padded(i) = Asc(Mid(str, i + 1, 1))
    Next
    
    ' 添加 0x80
    padded(msgLen) = &H80
    
    ' 填充零
    For i = msgLen + 1 To paddedLen - 9
        padded(i) = 0
    Next
    
    ' 添加64位消息长度（big-endian）
    Dim bitLen : bitLen = CDbl(msgLen) * 8
    For i = 0 To 7
        padded(paddedLen - 1 - i) = CByte(bitLen And &HFF)
        bitLen = Int(bitLen / 256)
    Next
    
    ' 2. 初始化哈希值
    H0 = &H6A09E667 : H1 = &HBB67AE85 : H2 = &H3C6EF372 : H3 = &HA54FF53A
    H4 = &H510E527F : H5 = &H9B05688C : H6 = &H1F83D9AB : H7 = &H5BE0CD19
    
    ' 3. 处理每个512位块
    numBlocks = paddedLen \ 64
    For blockIdx = 0 To numBlocks - 1
        bOffset = blockIdx * 64
        
        ' 准备消息调度 W[0..15]
        For i = 0 To 15
            W(i) = CLng(padded(bOffset + i*4)) * &H1000000 Or _
                   CLng(padded(bOffset + i*4 + 1)) * &H10000 Or _
                   CLng(padded(bOffset + i*4 + 2)) * &H100 Or _
                   CLng(padded(bOffset + i*4 + 3))
        Next
        
        ' 扩展 W[16..63]
        For i = 16 To 63
            S0 = ROTR32(W(i-15), 7) Xor ROTR32(W(i-15), 18) Xor (W(i-15) \ 8 And &H1FFFFFF)
            S1 = ROTR32(W(i-2), 17) Xor ROTR32(W(i-2), 19) Xor (W(i-2) \ 32 And &H7FFFFFF)
            W(i) = AddU32(AddU32(AddU32(W(i-16), S0), W(i-7)), S1)
        Next
        
        ' 初始化工作变量
        a = H0 : b = H1 : c = H2 : d = H3
        e = H4 : f = H5 : g = H6 : h = H7
        
        ' 64轮压缩
        For i = 0 To 63
            S1 = ROTR32(e, 6) Xor ROTR32(e, 11) Xor ROTR32(e, 25)
            ch = (e And f) Xor ((Not e) And g)
            t1 = AddU32(AddU32(AddU32(AddU32(h, S1), ch), SHA256_K(i)), W(i))
            S0 = ROTR32(a, 2) Xor ROTR32(a, 13) Xor ROTR32(a, 22)
            maj = (a And b) Xor (a And c) Xor (b And c)
            t2 = AddU32(S0, maj)
            
            h = g : g = f : f = e
            e = AddU32(d, t1)
            d = c : c = b : b = a
            a = AddU32(t1, t2)
        Next
        
        ' 累加哈希值
        H0 = AddU32(H0, a) : H1 = AddU32(H1, b)
        H2 = AddU32(H2, c) : H3 = AddU32(H3, d)
        H4 = AddU32(H4, e) : H5 = AddU32(H5, f)
        H6 = AddU32(H6, g) : H7 = AddU32(H7, h)
    Next
    
    ' 4. 输出十六进制
    SHA256Hash = Right("0000000" & Hex(H0), 8) & Right("0000000" & Hex(H1), 8) & _
                 Right("0000000" & Hex(H2), 8) & Right("0000000" & Hex(H3), 8) & _
                 Right("0000000" & Hex(H4), 8) & Right("0000000" & Hex(H5), 8) & _
                 Right("0000000" & Hex(H6), 8) & Right("0000000" & Hex(H7), 8)
End Function

' ============================================
' V9 兼容: 保留旧SimpleHash供参考，所有新代码应使用 SHA256Hash
' ============================================
' 简单哈希函数 - 已废弃，保留向后兼容
Function SimpleHash(str)
    SimpleHash = SafeSHA256Hash(str)
End Function

' ============================================
' 安全SHA-256包装：防止VBScript整数溢出
' 当SHA256Hash溢出时回退到DJB2简单哈希
' ============================================
Function SafeSHA256Hash(str)
    On Error Resume Next
    Dim result
    result = SHA256Hash(str)
    If Err.Number <> 0 Or result = "" Then
        Err.Clear
        ' 回退到DJB2简单哈希（不会溢出）
        result = DJB2Hash(str)
    End If
    On Error GoTo 0
    SafeSHA256Hash = result
End Function

' DJB2哈希（使用Double避免VBScript整数溢出）
Function DJB2Hash(str)
    Dim i, charCode, dblHash, quotient
    dblHash = CDbl(5381)
    If IsNull(str) Or str = "" Then
        DJB2Hash = "000000001505"
        Exit Function
    End If
    For i = 1 To Len(str)
        charCode = Asc(Mid(str, i, 1))
        dblHash = CDbl(dblHash * 33 + CDbl(charCode))
        ' Manual modulo: VBScript Mod converts to Long which can overflow.
        ' Use division-based approach safe for Double values.
        If dblHash > 2147483647 Then
            quotient = CDbl(dblHash / 2147483647)
            dblHash = CDbl(dblHash - CDbl(Fix(quotient)) * 2147483647)
        End If
    Next
    DJB2Hash = Right("0000000" & Hex(CLng(dblHash)), 8) & Right("0000000" & Hex(CLng(dblHash Xor 1431655765)), 8)
End Function

' 生成安全令牌 - V10: 使用SHA-256 + 增强密钥
Function GenerateSecureToken(adminId)
    Dim timestamp, dataToSign, signature
    
    If IsNull(adminId) Or adminId = "" Then
        GenerateSecureToken = ""
        Exit Function
    End If
    
    timestamp = CLng(DateDiff("s", "1970-01-01 00:00:00", Now()))
    dataToSign = adminId & "|" & timestamp & "|" & COOKIE_SECRET_V10
    signature = SafeSHA256Hash(dataToSign)
    
    GenerateSecureToken = adminId & "|" & timestamp & "|" & signature
End Function

' V9 DJB2 兼容哈希（仅用于向后兼容验证旧令牌）
Function LegacyDJB2(str)
    Dim i, hash1, hash2, charCode, tempDbl
    hash1 = 5381 : hash2 = 0
    If IsNull(str) Or str = "" Then LegacyDJB2 = "0" : Exit Function
    For i = 1 To Len(str)
        charCode = Asc(Mid(str, i, 1))
        tempDbl = (CDbl(hash1) * 33 + charCode)
        hash1 = CLng(tempDbl - Int(tempDbl / 2147483647) * 2147483647)
        tempDbl = (CDbl(hash2) * 31 + charCode)
        hash2 = CLng(tempDbl - Int(tempDbl / 2147483647) * 2147483647)
    Next
    If hash1 < 0 Then hash1 = hash1 + 2147483647
    If hash2 < 0 Then hash2 = hash2 + 2147483647
    LegacyDJB2 = Hex(hash1) & Hex(hash2)
End Function

' 验证安全令牌 - V10: 优先SHA-256验证，兼容旧DJB2令牌
Function ValidateSecureToken(token)
    Dim parts, adminId, timestamp, signature
    Dim expectedSignature, dataToSign
    Dim tokenTime, currentTime, daysDiff
    
    ValidateSecureToken = ""
    
    If IsNull(token) Or token = "" Then
        Exit Function
    End If
    
    ' 分割令牌
    parts = Split(token, "|")
    If UBound(parts) <> 2 Then
        Exit Function
    End If
    
    adminId = parts(0)
    timestamp = parts(1)
    signature = parts(2)
    
    ' 验证AdminID是否为数字
    If Not IsNumeric(adminId) Then
        Exit Function
    End If
    
    ' 验证时间戳是否为数字
    If Not IsNumeric(timestamp) Then
        Exit Function
    End If
    
    ' 检查令牌是否过期（30天）
    currentTime = CLng(DateDiff("s", "1970-01-01 00:00:00", Now()))
    daysDiff = (currentTime - CLng(timestamp)) / 86400
    If daysDiff > 30 Or daysDiff < 0 Then
        Exit Function
    End If
    
    ' V10: 优先验证SHA-256签名
    dataToSign = adminId & "|" & timestamp & "|" & COOKIE_SECRET_V10
    expectedSignature = SafeSHA256Hash(dataToSign)
    If signature = expectedSignature Then
        ValidateSecureToken = adminId
        Exit Function
    End If
    
    ' V9兼容: 尝试旧DJB2签名（过渡期）
    dataToSign = adminId & "|" & timestamp & "|" & COOKIE_SECRET
    expectedSignature = LegacyDJB2(dataToSign)
    If signature = expectedSignature Then
        ValidateSecureToken = adminId
    End If
End Function

' ============================================
' CSRF防护函数模块
' ============================================

' 确保CSRF令牌存在
Sub EnsureCSRFToken()
    If Session("CSRFToken") = "" Or IsEmpty(Session("CSRFToken")) Then
        Call GenerateCSRFToken()
    End If
End Sub

' 生成CSRF令牌
' V17: 生成新token时保留旧token到历史池，防止并发请求失效
Function GenerateCSRFToken()
    Dim token, i, oldToken, tokenArr, maxHistory
    
    ' 保存旧token到历史池（最多保留5个）
    maxHistory = 4
    oldToken = Session("CSRFToken")
    If oldToken <> "" And Not IsEmpty(oldToken) Then
        If IsArray(Session("CSRFTokenHistory")) Then
            tokenArr = Session("CSRFTokenHistory")
            ' 如果历史池未满，追加
            If UBound(tokenArr) < maxHistory Then
                ReDim Preserve tokenArr(UBound(tokenArr) + 1)
                tokenArr(UBound(tokenArr)) = oldToken
            Else
                ' 历史池已满，移除最旧的，添加最新的
                Dim j
                For j = 0 To UBound(tokenArr) - 1
                    tokenArr(j) = tokenArr(j + 1)
                Next
                tokenArr(UBound(tokenArr)) = oldToken
            End If
        Else
            tokenArr = Array(oldToken)
        End If
        Session("CSRFTokenHistory") = tokenArr
    End If
    
    Randomize
    token = ""
    
    For i = 1 To 32
        Dim charType, charCode
        charType = Int(Rnd * 3)
        Select Case charType
            Case 0
                charCode = Int(Rnd * 26) + 65  ' A-Z
            Case 1
                charCode = Int(Rnd * 26) + 97  ' a-z
            Case 2
                charCode = Int(Rnd * 10) + 48  ' 0-9
        End Select
        token = token & Chr(charCode)
    Next
    
    Session("CSRFToken") = token
    GenerateCSRFToken = token
End Function

' 获取CSRF令牌隐藏表单字段HTML
Function GetCSRFTokenField()
    Call EnsureCSRFToken()
    GetCSRFTokenField = "<input type=""hidden"" name=""csrf_token"" value=""" & Session("CSRFToken") & """>"
End Function

' 验证CSRF令牌
' V17: 增强版 - 支持Token池验证，防止并发请求失效
Function ValidateCSRFToken()
    Dim sessionToken, requestToken, i, tokenArr
    
    sessionToken = Session("CSRFToken")
    
    ' 先检查POST表单
    requestToken = Request.Form("csrf_token")
    
    ' 如果表单中没有，检查QueryString（支持GET请求）
    If requestToken = "" Then
        requestToken = Request.QueryString("csrf_token")
    End If
    
    ' 如果QueryString中没有，检查请求头（支持AJAX）
    If requestToken = "" Then
        requestToken = Request.ServerVariables("HTTP_X_CSRF_TOKEN")
    End If
    
    ' 验证主token
    If requestToken <> "" And requestToken = sessionToken Then
        ValidateCSRFToken = True
        Exit Function
    End If
    
    ' V17: 检查历史token池（防止并发AJAX请求失效）
    If requestToken <> "" And IsArray(Session("CSRFTokenHistory")) Then
        tokenArr = Session("CSRFTokenHistory")
        For i = 0 To UBound(tokenArr)
            If requestToken = tokenArr(i) Then
                ValidateCSRFToken = True
                Exit Function
            End If
        Next
    End If
    
    ValidateCSRFToken = False
End Function

' 获取CSRF令牌URL参数
Function GetCSRFTokenParam()
    Call EnsureCSRFToken()
    GetCSRFTokenParam = "csrf_token=" & Session("CSRFToken")
End Function

' ============================================
' 登录速率限制函数模块
' ============================================

' 检查登录是否被锁定
Function IsLoginLocked(prefix)
    Dim failCount, lockTime, minutesPassed
    
    ' 初始化Session变量
    If Session(prefix & "LoginFailCount") = "" Or IsEmpty(Session(prefix & "LoginFailCount")) Then
        Session(prefix & "LoginFailCount") = 0
    End If
    
    failCount = CInt(Session(prefix & "LoginFailCount"))
    
    If failCount >= 5 Then
        If Not IsEmpty(Session(prefix & "LoginLockTime")) And Session(prefix & "LoginLockTime") <> "" Then
            lockTime = Session(prefix & "LoginLockTime")
            minutesPassed = DateDiff("n", lockTime, Now())
            
            If minutesPassed < 15 Then
                IsLoginLocked = True
            Else
                ' 冷却期过，重置计数
                Session(prefix & "LoginFailCount") = 0
                IsLoginLocked = False
            End If
        Else
            IsLoginLocked = False
        End If
    Else
        IsLoginLocked = False
    End If
End Function

' 记录登录失败
Sub RecordLoginFailure(prefix)
    If Session(prefix & "LoginFailCount") = "" Or IsEmpty(Session(prefix & "LoginFailCount")) Then
        Session(prefix & "LoginFailCount") = 0
    End If
    
    Session(prefix & "LoginFailCount") = CInt(Session(prefix & "LoginFailCount")) + 1
    
    If CInt(Session(prefix & "LoginFailCount")) >= 5 Then
        Session(prefix & "LoginLockTime") = Now()
    End If
End Sub

' 重置登录失败计数
Sub ResetLoginFailure(prefix)
    Session(prefix & "LoginFailCount") = 0
    Session(prefix & "LoginLockTime") = ""
End Sub

' ============================================
' V14.6: 参数化查询函数
' 使用 ADODB.Command 实现安全的参数化查询，防止 SQL 注入
' 用法:
'   Dim params(1)
'   params(0) = Array("@UserID", adInteger, 4, userId)
'   params(1) = Array("@Status", adVarChar, 50, "Active")
'   Set rs = ExecuteParameterizedQuery("SELECT * FROM Users WHERE UserID=@UserID AND Status=@Status", params)
' ============================================
Function ExecuteParameterizedQuery(sql, params)
    Dim cmd, rs, i, param
    
    On Error Resume Next
    
    Set cmd = Server.CreateObject("ADODB.Command")
    If Err.Number <> 0 Or cmd Is Nothing Then
        Session("LastDBError") = "无法创建 ADODB.Command 对象: " & Err.Description
        Set ExecuteParameterizedQuery = Nothing
        Exit Function
    End If
    
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = 1  ' adCmdText
    cmd.CommandTimeout = 30
    
    ' 添加参数
    If IsArray(params) Then
        For i = 0 To UBound(params)
            If IsArray(params(i)) Then
                ' params(i) = Array(name, type, size, value)
                If UBound(params(i)) >= 3 Then
                    Set param = cmd.CreateParameter(params(i)(0), params(i)(1), 1, params(i)(2), params(i)(3))  ' 1 = adParamInput
                    cmd.Parameters.Append param
                End If
            End If
        Next
    End If
    
    ' 执行查询
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3  ' adUseClient
    rs.CursorType = 1      ' adOpenKeyset
    rs.LockType = 1        ' adLockOptimistic
    rs.Open cmd
    
    If Err.Number <> 0 Then
        Session("LastDBError") = "ExecuteParameterizedQuery错误: " & Err.Description & " 错误号: " & Err.Number & " SQL: " & sql
        If Not rs Is Nothing Then
            If rs.State = 1 Then rs.Close
            Set rs = Nothing
        End If
        Set ExecuteParameterizedQuery = Nothing
    Else
        Set ExecuteParameterizedQuery = rs
    End If
    
    Set cmd = Nothing
End Function

' ============================================
' V14.6: LIKE 通配符转义函数
' 对用户输入进行转义，防止 LIKE 注入
' 用法: sql = "SELECT * FROM Products WHERE ProductName LIKE '%" & SafeLike(userInput) & "%'"
' ============================================
Function SafeLike(str)
    If IsNull(str) Or str = "" Then
        SafeLike = ""
    Else
        ' 先进行基本 SQL 注入防护
        str = Replace(str, "'", "''")
        ' 转义 LIKE 通配符
        str = Replace(str, "[", "[[]")
        str = Replace(str, "%", "[%]")
        str = Replace(str, "_", "[_]")
        SafeLike = str
    End If
End Function
%>