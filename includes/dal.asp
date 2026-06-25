<%
' ============================================
' V15.0 统一数据访问层 (Data Access Layer)
' 提供参数化查询接口，杜绝SQL注入风险
' 依赖: connection.asp (conn, ExecuteParameterizedQuery)
' 用法: <!--#include file="dal.asp"-->
' ============================================

' ADO 常量定义
Const DAL_adInteger = 3      ' adInteger
Const DAL_adVarChar = 200    ' adVarChar
Const DAL_adBoolean = 11     ' adBoolean
Const DAL_adCurrency = 6     ' adCurrency
Const DAL_adDate = 7         ' adDate
Const DAL_adDBTimeStamp = 135 ' adDBTimeStamp
Const DAL_adDouble = 5       ' adDouble
Const DAL_adParamInput = 1   ' adParamInput

' ============================================
' 内部函数：创建 ADODB.Command 并绑定参数
' V17.1: 修复 MSOLEDBSQL/SQLOLEDB 不支持命名参数的问题
' 将 @ParamName 替换为 ? 位置占位符，使用匿名参数绑定
' 重要：调用者必须保证 arrParams 顺序与 SQL 中 @Param 出现的顺序一致
' ============================================
Function DAL_CreateCommand(sql, arrParams)
    Dim cmd, i, pName, pType, pSize, pValue, posSql
    
    On Error Resume Next
    Set cmd = Server.CreateObject("ADODB.Command")
    If Err.Number <> 0 Or cmd Is Nothing Then
        Set DAL_CreateCommand = Nothing
        Exit Function
    End If
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = 1  ' adCmdText
    cmd.CommandTimeout = DAL_QUERY_TIMEOUT
    
    ' V17.1: 转换命名参数为位置参数
    ' MSOLEDBSQL/SQLOLEDB 不支持 ADO Command 的命名参数绑定
    ' 必须使用 ? 占位符 + 匿名参数
    ' 注意：同一个 @ParamName 可能在 SQL 中出现多次，每次都需要一个 ? 占位符
    posSql = sql
    
    ' 绑定参数
    If IsArray(arrParams) Then
        For i = 0 To UBound(arrParams)
            If IsArray(arrParams(i)) Then
                If UBound(arrParams(i)) >= 2 Then
                    pName = arrParams(i)(0)
                    pType = arrParams(i)(1)
                    If UBound(arrParams(i)) >= 3 Then
                        pSize = arrParams(i)(2)
                    Else
                        pSize = 0
                    End If
                    If UBound(arrParams(i)) >= 3 Then
                        pValue = arrParams(i)(UBound(arrParams(i)))
                    Else
                        pValue = arrParams(i)(2)
                    End If
                    ' 处理NULL值
                    If IsNull(pValue) Or (VarType(pValue) = 0 And pValue = "") Then
                        pValue = Null
                    End If
                    
                    ' V17.1: 将 SQL 中所有出现的 @ParamName 替换为 ?
                    ' 同时统计出现次数，为每个出现创建一个匿名参数
                    Dim occurCount, j
                    occurCount = CountParamOccurrences(posSql, pName)
                    posSql = ReplaceAllParams(posSql, pName)
                    
                    For j = 1 To occurCount
                        Dim paramObj
                        If pSize > 0 Then
                            Set paramObj = cmd.CreateParameter("", pType, DAL_adParamInput, pSize, pValue)
                        Else
                            Set paramObj = cmd.CreateParameter("", pType, DAL_adParamInput, 0, pValue)
                        End If
                        cmd.Parameters.Append paramObj
                    Next
                End If
            End If
        Next
    End If
    
    ' V17.1: 设置转换后的 SQL（? 占位符版本）
    cmd.CommandText = posSql
    
    Set DAL_CreateCommand = cmd
End Function

' ============================================
' V17.1: 辅助函数 - 统计 @paramName 在 SQL 中出现的次数
' ============================================
Function CountParamOccurrences(sqlStr, paramName)
    Dim count, pos
    count = 0
    pos = InStr(1, sqlStr, paramName, 1)  ' vbTextCompare = 1
    Do While pos > 0
        count = count + 1
        pos = InStr(pos + Len(paramName), sqlStr, paramName, 1)
    Loop
    CountParamOccurrences = count
End Function

' ============================================
' V17.1: 辅助函数 - 将 SQL 中所有匹配的 @paramName 替换为 ?
' ============================================
Function ReplaceAllParams(sqlStr, paramName)
    Dim result, pos, paramLen
    result = sqlStr
    paramLen = Len(paramName)
    pos = InStr(1, result, paramName, 1)  ' vbTextCompare = 1，大小写不敏感
    Do While pos > 0
        result = Left(result, pos - 1) & "?" & Mid(result, pos + paramLen)
        pos = InStr(pos + 1, result, paramName, 1)
    Loop
    ReplaceAllParams = result
End Function

' ============================================
' DAL_Execute: 执行非查询语句 (INSERT/UPDATE/DELETE)
' 返回: 受影响行数（成功）或 -1（失败）
' 用法: DAL_Execute "UPDATE Users SET Name=@Name WHERE ID=@ID", _
'         Array(Array("@Name", DAL_adVarChar, 50, "张三"), _
'               Array("@ID", DAL_adInteger, 0, 123))
' ============================================
Function DAL_Execute(sql, arrParams)
    Dim cmd, rowsAffected, startTime
    
    startTime = Timer()
    rowsAffected = -1
    
    On Error Resume Next
    Set cmd = DAL_CreateCommand(sql, arrParams)
    If cmd Is Nothing Then
        DAL_Execute = -1
        Exit Function
    End If
    
    cmd.Execute rowsAffected
    
    If Err.Number <> 0 Then
        Session("DAL_LastError") = "DAL_Execute错误 (" & Err.Number & "): " & Err.Description & " | SQL: " & Left(sql, 200)
        rowsAffected = -1
        Err.Clear
    End If
    
    Set cmd = Nothing
    
    ' 慢查询记录
    If DAL_LOG_SLOW_QUERIES Then
        Dim elapsed : elapsed = Round((Timer() - startTime) * 1000, 1)
        If elapsed > DAL_SLOW_QUERY_THRESHOLD Then
            Session("DAL_SlowQuery") = "DAL_Execute: " & elapsed & "ms | " & Left(sql, 150)
        End If
    End If
    
    DAL_Execute = rowsAffected
End Function

' ============================================
' DAL_GetScalar: 获取单个标量值
' 返回: 查询结果的第一列第一行值，无结果返回 defaultVal
' ============================================
Function DAL_GetScalar(sql, arrParams, defaultVal)
    Dim cmd, rs, result, startTime
    
    If IsNull(defaultVal) Or IsEmpty(defaultVal) Then defaultVal = ""
    result = defaultVal
    startTime = Timer()
    
    On Error Resume Next
    Set cmd = DAL_CreateCommand(sql, arrParams)
    If cmd Is Nothing Then
        DAL_GetScalar = defaultVal
        Exit Function
    End If
    
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3  ' adUseClient
    rs.Open cmd, , 1, 1    ' adOpenForwardOnly, adLockReadOnly
    
    If Err.Number = 0 Then
        If Not rs.EOF Then
            result = rs.Fields(0).Value
            If IsNull(result) Then result = defaultVal
        End If
    Else
        Session("DAL_LastError") = "DAL_GetScalar错误 (" & Err.Number & "): " & Err.Description & " | SQL: " & Left(sql, 200)
        Err.Clear
    End If
    
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Set cmd = Nothing
    
    If DAL_LOG_SLOW_QUERIES Then
        Dim elapsed : elapsed = Round((Timer() - startTime) * 1000, 1)
        If elapsed > DAL_SLOW_QUERY_THRESHOLD Then
            Session("DAL_SlowQuery") = "DAL_GetScalar: " & elapsed & "ms | " & Left(sql, 150)
        End If
    End If
    
    DAL_GetScalar = result
End Function

' ============================================
' DAL_GetRow: 获取单行记录（返回Dictionary）
' 返回: Dictionary对象（字段名→值），无结果返回 Nothing
' ============================================
Function DAL_GetRow(sql, arrParams)
    Dim cmd, rs, dict, fld, startTime
    
    startTime = Timer()
    
    On Error Resume Next
    Set cmd = DAL_CreateCommand(sql, arrParams)
    If cmd Is Nothing Then
        Set DAL_GetRow = Nothing
        Exit Function
    End If
    
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3
    rs.Open cmd, , 1, 1
    
    If Err.Number = 0 Then
        If Not rs.EOF Then
            Set dict = Server.CreateObject("Scripting.Dictionary")
            For Each fld In rs.Fields
                If IsNull(fld.Value) Then
                    dict.Add fld.Name, ""
                Else
                    dict.Add fld.Name, fld.Value
                End If
            Next
            Set DAL_GetRow = dict
        End If
    Else
        Session("DAL_LastError") = "DAL_GetRow错误 (" & Err.Number & "): " & Err.Description & " | SQL: " & Left(sql, 200)
        Err.Clear
        Set DAL_GetRow = Nothing
    End If
    
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Set cmd = Nothing
    
    If DAL_LOG_SLOW_QUERIES Then
        Dim elapsed : elapsed = Round((Timer() - startTime) * 1000, 1)
        If elapsed > DAL_SLOW_QUERY_THRESHOLD Then
            Session("DAL_SlowQuery") = "DAL_GetRow: " & elapsed & "ms | " & Left(sql, 150)
        End If
    End If
End Function

' ============================================
' DAL_GetList: 获取多行记录（返回Recordset，兼容现有代码）
' 返回: ADODB.Recordset（客户端游标，支持RecordCount），无结果返回 Nothing
' ============================================
Function DAL_GetList(sql, arrParams)
    Dim cmd, rs, startTime
    
    startTime = Timer()
    
    On Error Resume Next
    Set cmd = DAL_CreateCommand(sql, arrParams)
    If cmd Is Nothing Then
        Set DAL_GetList = Nothing
        Exit Function
    End If
    
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3  ' adUseClient - 支持 RecordCount/MoveLast
    rs.CursorType = 1      ' adOpenKeyset
    rs.LockType = 1        ' adLockOptimistic
    rs.Open cmd
    
    If Err.Number = 0 Then
        Set DAL_GetList = rs
    Else
        Session("DAL_LastError") = "DAL_GetList错误 (" & Err.Number & "): " & Err.Description & " | SQL: " & Left(sql, 200)
        Err.Clear
        If Not rs Is Nothing Then
            If rs.State = 1 Then rs.Close
            Set rs = Nothing
        End If
        Set DAL_GetList = Nothing
    End If
    
    Set cmd = Nothing
    
    If DAL_LOG_SLOW_QUERIES Then
        Dim elapsed : elapsed = Round((Timer() - startTime) * 1000, 1)
        If elapsed > DAL_SLOW_QUERY_THRESHOLD Then
            Session("DAL_SlowQuery") = "DAL_GetList: " & elapsed & "ms | " & Left(sql, 150)
        End If
    End If
End Function

' ============================================
' DAL_GetListPaged: 分页查询
' 返回: Recordset + 通过pageInfo Dictionary返回分页信息
' ============================================
Function DAL_GetListPaged(sql, arrParams, page, pageSize, ByRef pageInfo)
    Dim countSql, totalCount, totalPages, rs
    
    ' 计算总数（从SQL中提取COUNT）
    ' 简化：执行两次查询（COUNT + 数据）
    If IsNull(page) Or page < 1 Then page = 1
    If IsNull(pageSize) Or pageSize < 1 Then pageSize = PAGE_SIZE
    
    ' 构建COUNT查询
    countSql = "SELECT COUNT(*) FROM (" & sql & ") AS DAL_CountSub"
    totalCount = CLng(DAL_GetScalar(countSql, arrParams, 0))
    totalPages = Int((totalCount + pageSize - 1) / pageSize)
    If totalPages < 1 Then totalPages = 1
    If page > totalPages Then page = totalPages
    
    ' 构建分页查询 (SQL Server 2012+ OFFSET/FETCH)
    Dim offset : offset = (page - 1) * pageSize
    Dim pagedSql
    pagedSql = sql & " OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY"
    
    Set rs = DAL_GetList(pagedSql, arrParams)
    Set DAL_GetListPaged = rs
    
    ' 填充分页信息
    Set pageInfo = Server.CreateObject("Scripting.Dictionary")
    pageInfo.Add "totalCount", totalCount
    pageInfo.Add "totalPages", totalPages
    pageInfo.Add "currentPage", page
    pageInfo.Add "pageSize", pageSize
    pageInfo.Add "hasNext", (page < totalPages)
    pageInfo.Add "hasPrev", (page > 1)
End Function

' ============================================
' DAL_Exists: 检查记录是否存在
' 返回: Boolean
' ============================================
Function DAL_Exists(tableName, whereClause, arrParams)
    Dim sql, count
    sql = "SELECT COUNT(*) FROM [" & tableName & "] WHERE " & whereClause
    count = CLng(DAL_GetScalar(sql, arrParams, 0))
    DAL_Exists = (count > 0)
End Function

' ============================================
' DAL_Insert: 插入记录并返回自增ID
' 返回: 新插入记录的IDENTITY值，失败返回0
' ============================================
Function DAL_Insert(tableName, fields, arrParams)
    Dim sql, cols, vals, i, insertId
    
    If IsArray(fields) Then
        cols = Join(fields, ",")
        vals = ""
        For i = 0 To UBound(fields)
            If i > 0 Then vals = vals & ","
            vals = vals & "@" & Replace(fields(i), " ", "")
        Next
    Else
        DAL_Insert = 0
        Exit Function
    End If
    
    sql = "INSERT INTO [" & tableName & "] (" & cols & ") VALUES (" & vals & "); SELECT SCOPE_IDENTITY();"
    insertId = CLng(DAL_GetScalar(sql, arrParams, 0))
    DAL_Insert = insertId
End Function

' ============================================
' DAL_GetLastError: 获取最后一次DAL错误
' ============================================
Function DAL_GetLastError()
    If Session("DAL_LastError") <> "" Then
        DAL_GetLastError = Session("DAL_LastError")
        Session("DAL_LastError") = ""
    Else
        DAL_GetLastError = ""
    End If
End Function
%>