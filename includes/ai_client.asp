<%
' ============================================
' V18.0 AI 微服务客户端 (ASP HTTP Wrapper)
' 提供 Classic ASP 到 Python AI 微服务的桥梁
' 用法: <!--#include file="ai_client.asp"-->
' 调用: AI_CallService("recommend/personalized", payload)
' ============================================

' AI 服务配置
Const AI_SERVICE_BASE_URL = "http://127.0.0.1:5000"
Const AI_SERVICE_TIMEOUT = 10      ' 请求超时（秒）
Const AI_SERVICE_FALLBACK = True   ' 服务不可用时返回空数据（不中断页面）

' ============================================
' AI_CallService: 调用 AI 微服务端点
' 参数:
'   endpoint - API 路径（不含 /api/ 前缀），如 "recommend/personalized"
'   payload  - 请求体（Dictionary 对象，自动转 JSON）
'   method   - HTTP 方法，默认 POST
' 返回:
'   Dictionary 对象（解析后的 JSON 响应），失败返回 Empty
' ============================================
Function AI_CallService(endpoint, payload, method)
    If IsEmpty(method) Or IsNull(method) Then method = "POST"
    
    Dim http, url, jsonBody, responseText, result
    Set result = Nothing
    
    On Error Resume Next
    
    ' 构建 URL
    url = AI_SERVICE_BASE_URL & "/api/" & endpoint
    
    ' 将 payload Dictionary 转为 JSON 字符串
    jsonBody = AI_DictToJsonString(payload)
    
    ' 创建 HTTP 请求对象
    Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If http Is Nothing Then
        Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
    End If
    If http Is Nothing Then
        Set http = Server.CreateObject("Microsoft.XMLHTTP")
    End If
    
    If http Is Nothing Then
        If Not AI_SERVICE_FALLBACK Then
            Session("AI_Error") = "无法创建 HTTP 客户端"
        End If
        AI_CallService = Empty
        Exit Function
    End If
    
    ' 设置超时
    http.SetTimeouts AI_SERVICE_TIMEOUT * 1000, AI_SERVICE_TIMEOUT * 1000, _
                     AI_SERVICE_TIMEOUT * 1000, AI_SERVICE_TIMEOUT * 1000
    
    ' 发送请求
    http.Open method, url, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Accept", "application/json"
    
    If jsonBody <> "" And jsonBody <> "{}" Then
        http.Send jsonBody
    Else
        http.Send
    End If
    
    ' 检查响应
    If http.Status = 200 Then
        responseText = http.ResponseText
        If responseText <> "" Then
            Set result = AI_ParseJsonResponse(responseText)
        End If
    Else
        If Not AI_SERVICE_FALLBACK Then
            Session("AI_Error") = "AI 服务返回错误: HTTP " & http.Status
        End If
    End If
    
    Set http = Nothing
    On Error GoTo 0
    
    If result Is Nothing Then
        AI_CallService = Empty
    Else
        Set AI_CallService = result
    End If
End Function

' ============================================
' AI_CallServiceGET: GET 请求简化版
' ============================================
Function AI_CallServiceGET(endpoint)
    AI_CallServiceGET = AI_CallService(endpoint, Empty, "GET")
End Function

' ============================================
' AI_GetPersonalized: 获取个性化推荐（便捷函数）
' 返回: Array of product_id
' ============================================
Function AI_GetPersonalized(userId, limit, excludeIds)
    If IsEmpty(limit) Or Not IsNumeric(limit) Then limit = 10
    
    Dim payload
    Set payload = Server.CreateObject("Scripting.Dictionary")
    payload.Add "user_id", CLng(userId)
    payload.Add "limit", CInt(limit)
    If Not IsEmpty(excludeIds) Then
        If IsArray(excludeIds) Then
            payload.Add "exclude_ids", excludeIds
        End If
    End If
    
    Dim result, data, i, productIds()
    Set result = AI_CallService("recommend/personalized", payload, "POST")
    
    If Not IsEmpty(result) And IsObject(result) Then
        If result.Exists("data") Then
            Set data = result("data")
            If IsObject(data) Then
                ReDim productIds(data.Count - 1)
                For i = 0 To data.Count - 1
                    If IsObject(data(i)) Then
                        If data(i).Exists("product_id") Then
                            productIds(i) = data(i)("product_id")
                        End If
                    End If
                Next
                AI_GetPersonalized = productIds
                Exit Function
            End If
        End If
    End If
    
    AI_GetPersonalized = Array()
End Function

' ============================================
' AI_GetSimilarProducts: 获取相似产品（便捷函数）
' ============================================
Function AI_GetSimilarProducts(productId, limit)
    If IsEmpty(limit) Or Not IsNumeric(limit) Then limit = 6
    
    Dim payload
    Set payload = Server.CreateObject("Scripting.Dictionary")
    payload.Add "product_id", CLng(productId)
    payload.Add "limit", CInt(limit)
    
    Dim result, data, i, productIds()
    Set result = AI_CallService("recommend/similar", payload, "POST")
    
    If Not IsEmpty(result) And IsObject(result) Then
        If result.Exists("data") Then
            Set data = result("data")
            If IsObject(data) Then
                ReDim productIds(data.Count - 1)
                For i = 0 To data.Count - 1
                    If IsObject(data(i)) Then
                        If data(i).Exists("product_id") Then
                            productIds(i) = data(i)("product_id")
                        End If
                    End If
                Next
                AI_GetSimilarProducts = productIds
                Exit Function
            End If
        End If
    End If
    
    AI_GetSimilarProducts = Array()
End Function

' ============================================
' AI_MatchFragrance: 香氛匹配（便捷函数）
' ============================================
Function AI_MatchFragrance(answers)
    Dim payload
    Set payload = Server.CreateObject("Scripting.Dictionary")
    
    If IsObject(answers) Then
        Set payload("answers") = answers
    End If
    
    Dim result
    Set result = AI_CallService("fragrance/match", payload, "POST")
    
    If Not IsEmpty(result) And IsObject(result) Then
        If result.Exists("data") Then
            Set AI_MatchFragrance = result("data")
            Exit Function
        End If
    End If
    
    AI_MatchFragrance = Empty
End Function

' ============================================
' AI_AnalyzeSentiment: 情感分析（便捷函数）
' ============================================
Function AI_AnalyzeSentiment(text)
    Dim payload
    Set payload = Server.CreateObject("Scripting.Dictionary")
    payload.Add "text", text
    
    Dim result
    Set result = AI_CallService("sentiment/analyze", payload, "POST")
    
    If Not IsEmpty(result) And IsObject(result) Then
        If result.Exists("data") Then
            Set AI_AnalyzeSentiment = result("data")
            Exit Function
        End If
    End If
    
    AI_AnalyzeSentiment = Empty
End Function

' ============================================
' AI_ChatbotMessage: 智能客服消息（便捷函数）
' ============================================
Function AI_ChatbotMessage(message, sessionId)
    Dim payload
    Set payload = Server.CreateObject("Scripting.Dictionary")
    payload.Add "message", message
    If sessionId <> "" Then payload.Add "session_id", sessionId
    
    Dim result
    Set result = AI_CallService("chatbot/message", payload, "POST")
    
    If Not IsEmpty(result) And IsObject(result) Then
        If result.Exists("data") Then
            Set AI_ChatbotMessage = result("data")
            Exit Function
        End If
    End If
    
    AI_ChatbotMessage = Empty
End Function

' ============================================
' AI_DictToJsonString: Dictionary → 简单 JSON 字符串
' （minimal implementation，复杂结构请用 api_response.asp）
' ============================================
Function AI_DictToJsonString(dict)
    If Not IsObject(dict) Then
        AI_DictToJsonString = "{}"
        Exit Function
    End If
    
    Dim keys, key, parts(), i, val
    On Error Resume Next
    keys = dict.Keys()
    If Err.Number <> 0 Then
        Err.Clear
        AI_DictToJsonString = "{}"
        Exit Function
    End If
    
    i = 0
    ReDim parts(0)
    For Each key In keys
        If i > 0 Then ReDim Preserve parts(i)
        val = dict.Item(key)
        
        parts(i) = """" & AI_EscapeJson(key) & """:"
        
        If IsNull(val) Or IsEmpty(val) Then
            parts(i) = parts(i) & "null"
        ElseIf VarType(val) = vbBoolean Then
            parts(i) = parts(i) & LCase(CStr(val))
        ElseIf IsNumeric(val) And VarType(val) <> vbString Then
            parts(i) = parts(i) & CStr(val)
        ElseIf IsObject(val) Then
            parts(i) = parts(i) & AI_DictToJsonString(val)
        ElseIf IsArray(val) Then
            parts(i) = parts(i) & AI_ArrayToJsonString(val)
        Else
            parts(i) = parts(i) & """" & AI_EscapeJson(CStr(val)) & """"
        End If
        i = i + 1
    Next
    
    AI_DictToJsonString = "{" & Join(parts, ",") & "}"
End Function

' ============================================
' AI_ArrayToJsonString: VBScript Array → JSON 数组
' ============================================
Function AI_ArrayToJsonString(arr)
    Dim i, parts()
    If Not IsArray(arr) Then
        AI_ArrayToJsonString = "[]"
        Exit Function
    End If
    
    ReDim parts(UBound(arr))
    For i = 0 To UBound(arr)
        If IsNull(arr(i)) Or IsEmpty(arr(i)) Then
            parts(i) = "null"
        ElseIf IsNumeric(arr(i)) Then
            parts(i) = CStr(arr(i))
        Else
            parts(i) = """" & AI_EscapeJson(CStr(arr(i))) & """"
        End If
    Next
    
    AI_ArrayToJsonString = "[" & Join(parts, ",") & "]"
End Function

' ============================================
' AI_EscapeJson: 转义 JSON 字符串
' ============================================
Function AI_EscapeJson(s)
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    AI_EscapeJson = s
End Function

' ============================================
' AI_ParseJsonResponse: 解析 JSON 响应为 Dictionary
' （简化版，生产环境建议使用 JSON 解析 COM 组件）
' ============================================
Function AI_ParseJsonResponse(jsonStr)
    ' 使用 ScriptControl 或 RegExp 解析简单 JSON
    ' 简化实现：提取 "data" 字段为嵌套 Dictionary
    Dim result, dataDict
    Set result = Server.CreateObject("Scripting.Dictionary")
    
    ' 提取 code
    result.Add "code", AI_ExtractJsonInt(jsonStr, "code")
    
    ' 提取 message
    Dim msg
    msg = AI_ExtractJsonString(jsonStr, "message")
    If msg <> "" Then result.Add "message", msg
    
    ' 提取 data 对象（简化：仅支持对象和数组）
    Dim dataObj
    Set dataObj = AI_ExtractJsonObject(jsonStr, "data")
    If Not dataObj Is Nothing Then
        result.Add "data", dataObj
    End If
    
    Set AI_ParseJsonResponse = result
End Function

' ============================================
' AI_ExtractJsonInt: 提取 JSON 整数字段
' ============================================
Function AI_ExtractJsonInt(jsonStr, fieldName)
    Dim regex, matches
    Set regex = New RegExp
    regex.Pattern = """" & fieldName & """:\s*(-?\d+)"
    regex.IgnoreCase = True
    Set matches = regex.Execute(jsonStr)
    If matches.Count > 0 Then
        AI_ExtractJsonInt = CLng(matches(0).SubMatches(0))
    Else
        AI_ExtractJsonInt = 0
    End If
    Set regex = Nothing
End Function

' ============================================
' AI_ExtractJsonString: 提取 JSON 字符串字段
' ============================================
Function AI_ExtractJsonString(jsonStr, fieldName)
    Dim regex, matches
    Set regex = New RegExp
    regex.Pattern = """" & fieldName & """:\s*""([^""]*)"""
    regex.IgnoreCase = True
    Set matches = regex.Execute(jsonStr)
    If matches.Count > 0 Then
        AI_ExtractJsonString = matches(0).SubMatches(0)
    Else
        AI_ExtractJsonString = ""
    End If
    Set regex = Nothing
End Function

' ============================================
' AI_ExtractJsonObject: 提取 JSON 对象/数组字段
' ============================================
Function AI_ExtractJsonObject(jsonStr, fieldName)
    ' 简化实现：使用正则提取嵌套 JSON
    Dim dict, regex, matches, subJson, items(), itemEnd, braceCount, i, ch
    Set dict = Nothing
    
    Set regex = New RegExp
    regex.Pattern = """" & fieldName & """:\s*(\[)"
    regex.IgnoreCase = True
    Set matches = regex.Execute(jsonStr)
    
    If matches.Count = 0 Then
        ' 尝试匹配对象 {}
        regex.Pattern = """" & fieldName & """:\s*(\{)"
        Set matches = regex.Execute(jsonStr)
    End If
    
    If matches.Count > 0 Then
        Dim startPos, openChar, closeChar
        startPos = matches(0).FirstIndex + Len(matches(0).Value) - 1
        openChar = Mid(jsonStr, startPos + 1, 1)
        If openChar = "[" Then closeChar = "]" Else closeChar = "}"
        
        ' 找到匹配的闭合括号
        braceCount = 1
        For i = startPos + 2 To Len(jsonStr)
            ch = Mid(jsonStr, i, 1)
            If ch = openChar Then braceCount = braceCount + 1
            If ch = closeChar Then
                braceCount = braceCount - 1
                If braceCount = 0 Then
                    subJson = Mid(jsonStr, startPos + 1, i - startPos)
                    Exit For
                End If
            End If
        Next
        
        If subJson <> "" Then
            If openChar = "[" Then
                Set dict = AI_ParseJsonArray(subJson)
            Else
                Set dict = AI_ParseJsonObject(subJson)
            End If
        End If
    End If
    
    Set regex = Nothing
    If dict Is Nothing Then
        Set AI_ExtractJsonObject = Nothing
    Else
        Set AI_ExtractJsonObject = dict
    End If
End Function

' ============================================
' AI_ParseJsonArray: 解析 JSON 数组
' ============================================
Function AI_ParseJsonArray(jsonStr)
    Dim dict, regex, matches, item, i
    Set dict = Server.CreateObject("Scripting.Dictionary")
    
    ' 提取数组中的对象
    i = 0
    Set regex = New RegExp
    regex.Pattern = "\{[^}]+\}"
    regex.Global = True
    Set matches = regex.Execute(jsonStr)
    
    For Each item In matches
        Set dict(i) = AI_ParseJsonObject(item.Value)
        i = i + 1
    Next
    
    Set regex = Nothing
    Set AI_ParseJsonArray = dict
End Function

' ============================================
' AI_ParseJsonObject: 解析 JSON 对象
' ============================================
Function AI_ParseJsonObject(jsonStr)
    Dim dict, regex, matches, match, key, value
    Set dict = Server.CreateObject("Scripting.Dictionary")
    
    Set regex = New RegExp
    regex.Pattern = """(\w+)""\s*:\s*(""[^""]*""|-?\d+\.?\d*|true|false|null)"
    regex.Global = True
    Set matches = regex.Execute(jsonStr)
    
    For Each match In matches
        key = match.SubMatches(0)
        value = match.SubMatches(1)
        
        ' 判断类型
        If Left(value, 1) = """" Then
            value = Mid(value, 2, Len(value) - 2)  ' 去掉引号
        ElseIf value = "true" Then
            value = True
        ElseIf value = "false" Then
            value = False
        ElseIf value = "null" Then
            value = Null
        ElseIf IsNumeric(value) Then
            value = CDbl(value)
        End If
        
        dict.Add key, value
    Next
    
    Set regex = Nothing
    Set AI_ParseJsonObject = dict
End Function
%>
