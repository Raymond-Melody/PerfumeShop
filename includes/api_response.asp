<%
' ============================================
' V16.0 API 统一响应格式 (API Response Standardization)
' 用法: <!--#include file="api_response.asp"-->
' 调用: API_Success objData, "操作成功"
'        API_Error API_ERR_PARAM_MISSING, "缺少必填参数"
'        API_Response API_ERR_SUCCESS, "success", objData
' ============================================

' ============================================
' 标准错误码定义
' ============================================
' 成功
Const API_ERR_SUCCESS = 0

' 认证/授权错误 (1000-1999)
Const API_ERR_AUTH_REQUIRED = 1001
Const API_ERR_AUTH_EXPIRED  = 1002
Const API_ERR_CSRF_INVALID  = 1003
Const API_ERR_FORBIDDEN     = 1004

' 参数验证错误 (2000-2999)
Const API_ERR_PARAM_MISSING = 2001
Const API_ERR_PARAM_INVALID = 2002
Const API_ERR_PARAM_TYPE    = 2003

' 业务逻辑错误 (3000-3999)
Const API_ERR_NOT_FOUND     = 3001
Const API_ERR_DUPLICATE     = 3002
Const API_ERR_LIMIT_EXCEEDED = 3003
Const API_ERR_BUSINESS_RULE = 3004

' 数据库错误 (4000-4999)
Const API_ERR_DB_ERROR      = 4001
Const API_ERR_DB_TIMEOUT    = 4002
Const API_ERR_DB_DEADLOCK   = 4003

' 文件/上传错误 (5000-5999)
Const API_ERR_FILE_UPLOAD   = 5001
Const API_ERR_FILE_TYPE     = 5002
Const API_ERR_FILE_SIZE     = 5003

' 服务端错误 (6000-6999)
Const API_ERR_SERVER_ERROR  = 6001
Const API_ERR_MAINTENANCE   = 6002

' ============================================
' 生成唯一请求ID
' ============================================
Function API_GetRequestId()
    Dim nowVal, randomPart
    nowVal = Now()
    Randomize
    randomPart = Right("0000" & Hex(Int(Rnd * 65535)), 4)
    API_GetRequestId = Year(nowVal) & Right("0" & Month(nowVal), 2) & Right("0" & Day(nowVal), 2) & _
                       Right("0" & Hour(nowVal), 2) & Right("0" & Minute(nowVal), 2) & Right("0" & Second(nowVal), 2) & _
                       randomPart
End Function

' ============================================
' JSON编码辅助函数 - 安全转义字符串值
' ============================================
Function API_JsonEncode(val)
    If IsNull(val) Or IsEmpty(val) Then
        API_JsonEncode = "null"
        Exit Function
    End If
    
    If VarType(val) = vbBoolean Then
        If val Then
            API_JsonEncode = "true"
        Else
            API_JsonEncode = "false"
        End If
        Exit Function
    End If
    
    If IsNumeric(val) And VarType(val) <> vbString Then
        API_JsonEncode = CStr(val)
        Exit Function
    End If
    
    ' 字符串转义
    Dim s
    s = CStr(val)
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    API_JsonEncode = """" & s & """"
End Function

' ============================================
' 将Dictionary/数组转换为JSON字符串
' ============================================
Function API_DictToJson(dict)
    If Not IsObject(dict) Then
        API_DictToJson = "{}"
        Exit Function
    End If
    
    Dim keys, key, i, parts, val, isArray
    
    ' 检查是否为数组 (通过 Recordset.GetRows 等方式)
    On Error Resume Next
    keys = dict.Keys()
    If Err.Number <> 0 Then
        Err.Clear
        API_DictToJson = "{}"
        Exit Function
    End If
    
    ' 判断是简单值、数组还是对象
    isArray = True
    For Each key In keys
        If Not IsNumeric(key) Then
            isArray = False
            Exit For
        End If
    Next
    
    ReDim parts(0)
    i = 0
    
    If isArray Then
        ' 输出为JSON数组
        For Each key In keys
            val = dict.Item(key)
            If i > 0 Then ReDim Preserve parts(i)
            If IsObject(val) Then
                parts(i) = API_DictToJson(val)
            Else
                parts(i) = API_JsonEncode(val)
            End If
            i = i + 1
        Next
        API_DictToJson = "[" & Join(parts, ",") & "]"
    Else
        ' 输出为JSON对象
        For Each key In keys
            val = dict.Item(key)
            If i > 0 Then ReDim Preserve parts(i)
            If IsObject(val) Then
                parts(i) = API_JsonEncode(key) & ":" & API_DictToJson(val)
            Else
                parts(i) = API_JsonEncode(key) & ":" & API_JsonEncode(val)
            End If
            i = i + 1
        Next
        API_DictToJson = "{" & Join(parts, ",") & "}"
    End If
End Function

' ============================================
' 将Recordset转换为JSON数组字符串
' ============================================
Function API_RecordsetToJson(rs)
    Dim rows, fld, rowCount, colCount, rowParts, colParts
    
    If rs Is Nothing Then
        API_RecordsetToJson = "[]"
        Exit Function
    End If
    
    On Error Resume Next
    If rs.EOF Then
        API_RecordsetToJson = "[]"
        Exit Function
    End If
    
    rowCount = 0
    ReDim rows(0)
    
    ' 获取列名
    colCount = rs.Fields.Count - 1
    ReDim colNames(colCount)
    For Each fld In rs.Fields
        colNames(colCount) = API_JsonEncode(fld.Name)
        colCount = colCount - 1
    Next
    ' 修正索引（因为Fields是倒序遍历的）
    
    ' 简化处理：逐行构建
    Do While Not rs.EOF
        ReDim Preserve rows(rowCount)
        ReDim colParts(rs.Fields.Count - 1)
        Dim j: j = 0
        For Each fld In rs.Fields
            colParts(j) = API_JsonEncode(fld.Name) & ":" & API_JsonEncode(fld.Value)
            j = j + 1
        Next
        rows(rowCount) = "{" & Join(colParts, ",") & "}"
        rowCount = rowCount + 1
        rs.MoveNext
    Loop
    
    API_RecordsetToJson = "[" & Join(rows, ",") & "]"
    Err.Clear
End Function

' ============================================
' 核心响应函数：输出标准化JSON
' ============================================
Sub API_Response(code, message, data)
    Dim requestId, json
    
    requestId = API_GetRequestId()
    
    ' 设置响应头
    Response.ContentType = "application/json"
    Response.Charset = "UTF-8"
    Response.AddHeader "X-API-Version", "v1"
    Response.AddHeader "X-Request-ID", requestId
    
    ' 构建JSON
    json = "{"
    json = json & """code"":" & CLng(code)
    json = json & ",""message"":" & API_JsonEncode(message)
    json = json & ",""timestamp"":" & API_JsonEncode(Now())
    json = json & ",""requestId"":" & API_JsonEncode(requestId)
    
    ' data字段
    If IsNull(data) Or IsEmpty(data) Then
        json = json & ",""data"":null"
    ElseIf IsObject(data) Then
        ' 尝试判断是Dictionary还是Recordset
        On Error Resume Next
        Dim testKeys
        testKeys = data.Keys()
        If Err.Number = 0 Then
            ' 是Dictionary
            json = json & ",""data"":" & API_DictToJson(data)
        Else
            Err.Clear
            ' 尝试Recordset
            json = json & ",""data"":" & API_RecordsetToJson(data)
            Err.Clear
        End If
    ElseIf VarType(data) = vbBoolean Then
        If data Then
            json = json & ",""data"":true"
        Else
            json = json & ",""data"":false"
        End If
    ElseIf IsNumeric(data) Then
        json = json & ",""data"":" & CStr(data)
    Else
        json = json & ",""data"":" & API_JsonEncode(data)
    End If
    
    json = json & "}"
    
    Response.Write json
End Sub

' ============================================
' 便捷函数：成功响应
' ============================================
Sub API_Success(data, message)
    Dim msg
    If IsNull(message) Or IsEmpty(message) Or message = "" Then
        msg = "success"
    Else
        msg = CStr(message)
    End If
    Call API_Response(API_ERR_SUCCESS, msg, data)
End Sub

' ============================================
' 便捷函数：错误响应
' ============================================
Sub API_Error(code, message)
    Call API_Response(code, message, Null)
End Sub

' ============================================
' 便捷函数：仅消息响应（无数据）
' ============================================
Sub API_Message(success, message)
    If success Then
        Call API_Success(Null, message)
    Else
        Call API_Error(API_ERR_BUSINESS_RULE, message)
    End If
End Sub

' ============================================
' 快速CSRF验证
' 返回True表示验证通过或无需验证
' ============================================
Function API_CheckCSRF()
    Dim token
    token = Request.Form("csrf_token")
    If token = "" Then token = Request.QueryString("csrf_token")
    
    ' 如果token为空且Session中也没有存储token，允许通过（兼容旧版API）
    If token = "" And (Session("CSRFToken") = "" Or IsEmpty(Session("CSRFToken"))) Then
        API_CheckCSRF = True
        Exit Function
    End If
    
    If token <> Session("CSRFToken") Then
        API_CheckCSRF = False
        Exit Function
    End If
    
    API_CheckCSRF = True
End Function

' ============================================
' 快速登录验证
' 返回True表示已登录
' ============================================
Function API_RequireLogin()
    If Session("UserID") <> "" And Not IsEmpty(Session("UserID")) Then
        API_RequireLogin = True
    ElseIf Session("AdminID") <> "" And Not IsEmpty(Session("AdminID")) Then
        API_RequireLogin = True
    Else
        API_RequireLogin = False
        Call API_Error(API_ERR_AUTH_REQUIRED, "请先登录")
        Response.End
    End If
End Function

' ============================================
' 简易错误消息映射
' ============================================
Function API_GetErrorMessage(code)
    Select Case code
        Case API_ERR_AUTH_REQUIRED: API_GetErrorMessage = "请先登录"
        Case API_ERR_AUTH_EXPIRED:  API_GetErrorMessage = "登录已过期，请重新登录"
        Case API_ERR_CSRF_INVALID:  API_GetErrorMessage = "安全验证失败，请刷新页面重试"
        Case API_ERR_FORBIDDEN:     API_GetErrorMessage = "没有操作权限"
        Case API_ERR_PARAM_MISSING: API_GetErrorMessage = "缺少必填参数"
        Case API_ERR_PARAM_INVALID: API_GetErrorMessage = "参数格式不正确"
        Case API_ERR_NOT_FOUND:     API_GetErrorMessage = "请求的资源不存在"
        Case API_ERR_DUPLICATE:     API_GetErrorMessage = "数据已存在，请勿重复操作"
        Case API_ERR_DB_ERROR:      API_GetErrorMessage = "数据库操作失败"
        Case API_ERR_SERVER_ERROR:  API_GetErrorMessage = "服务器内部错误"
        Case Else:                  API_GetErrorMessage = "未知错误"
    End Select
End Function
%>