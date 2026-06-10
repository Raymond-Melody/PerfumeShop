<%
' ============================================
' 数据库连接模块 - SQL Server版本
' ============================================

Dim conn

' 打开数据库连接
Sub OpenConnection()
    On Error Resume Next
    Set conn = Server.CreateObject("ADODB.Connection")
    
    If Err.Number <> 0 Then
        Response.Write "<div class='error'>无法创建数据库连接对象: " & Replace(Server.HTMLEncode(Err.Description), "'", "&#39;") & " 错误号: " & Err.Number & "</div>"
        Response.End
    End If
    
    ' 检查对象是否成功创建
    If conn Is Nothing Then
        Response.Write "<div class='error'>数据库连接对象创建失败</div>"
        Response.End
    End If
    
    ' SQL Server数据库连接字符串 (默认实例 MSSQLSERVER, SQL Server 2017)
    conn.Open "Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;"
    
    If Err.Number <> 0 Then
        Response.Write "<div class='error'>数据库连接失败: " & Replace(Server.HTMLEncode(Err.Description), "'", "&#39;") & " 错误号: " & Err.Number & "</div>"
        Response.End
    End If
End Sub

' 关闭数据库连接
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
    orderNo = "PF" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2)
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
    ElseIf Not IsDate(dateValue) Then
        SafeFormatDateTime = "-"
    Else
        SafeFormatDateTime = FormatDateTime(dateValue, formatType)
    End If
    
    If Err.Number <> 0 Then
        SafeFormatDateTime = "-"
        Err.Clear
    End If
End Function

' ============================================
' 安全工具函数模块
' ============================================

' Cookie签名密钥
Const COOKIE_SECRET = "PerfumeShop_SecKey_2026_X9K3m"

' 简单哈希函数 - 用于Cookie签名
Function SimpleHash(str)
    Dim i, hash1, hash2, charCode
    Dim tempDbl
    hash1 = 5381
    hash2 = 0
    
    If IsNull(str) Or str = "" Then
        SimpleHash = "0"
        Exit Function
    End If
    
    For i = 1 To Len(str)
        charCode = Asc(Mid(str, i, 1))
        ' 使用DJB2变种算法 - 使用双精度进行计算，避免溢出
        tempDbl = (CDbl(hash1) * 33 + charCode)
        ' 使用 Mod 前先确保数值在合理范围内
        hash1 = CLng(tempDbl - Int(tempDbl / 2147483647) * 2147483647)
        
        tempDbl = (CDbl(hash2) * 31 + charCode)
        hash2 = CLng(tempDbl - Int(tempDbl / 2147483647) * 2147483647)
    Next
    
    ' 确保哈希值为正数
    If hash1 < 0 Then hash1 = hash1 + 2147483647
    If hash2 < 0 Then hash2 = hash2 + 2147483647
    
    ' 组合两个哈希值并转换为16进制字符串
    SimpleHash = Hex(hash1) & Hex(hash2)
End Function

' 生成安全令牌 - 用于Cookie加密
Function GenerateSecureToken(adminId)
    Dim timestamp, dataToSign, signature
    
    If IsNull(adminId) Or adminId = "" Then
        GenerateSecureToken = ""
        Exit Function
    End If
    
    timestamp = CLng(DateDiff("s", "1970-01-01 00:00:00", Now()))
    dataToSign = adminId & "|" & timestamp & "|" & COOKIE_SECRET
    signature = SimpleHash(dataToSign)
    
    GenerateSecureToken = adminId & "|" & timestamp & "|" & signature
End Function

' 验证安全令牌 - 返回AdminID或空字符串
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
    
    ' 验证签名
    dataToSign = adminId & "|" & timestamp & "|" & COOKIE_SECRET
    expectedSignature = SimpleHash(dataToSign)
    
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
Function GenerateCSRFToken()
    Dim token, i
    Randomize
    token = ""
    
    For i = 1 To 32
        ' 随机生成A-Z, a-z, 0-9字符
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
Function ValidateCSRFToken()
    Dim sessionToken, requestToken
    
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
    
    ' 验证
    If sessionToken = "" Or requestToken = "" Then
        ValidateCSRFToken = False
    ElseIf sessionToken = requestToken Then
        ValidateCSRFToken = True
    Else
        ValidateCSRFToken = False
    End If
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
%>