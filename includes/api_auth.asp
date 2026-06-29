<%
' ============================================
' V18.0 API 认证层 (API Authentication)
' API Key + HMAC 签名验证层
' 依赖: config.asp, connection.asp
' 用法: <!--#include file="api_auth.asp"-->
'        If Not API_AuthCheck() Then Response.End
' ============================================

' API 认证配置常量
Const API_AUTH_ENABLED = True
Const API_AUTH_SECRET = "V18_API_HMAC_Secret_9xK4mR7vN2w"
Const API_AUTH_REQUIRE_SIGNATURE = True
Const API_KEY_HEADER = "X-API-Key"
Const API_SIGN_HEADER = "X-API-Signature"
Const API_TS_HEADER = "X-API-Timestamp"
Const API_SIGN_TIMEOUT = 300              ' 签名有效期（秒），5分钟
Const API_INTERNAL_KEY = "internal_service_v18_key"

' ============================================
' API_AuthCheck(): 验证 API 请求的合法性
' 支持两种模式：
'   1. Session 认证（已登录用户的 API 请求）
'   2. API Key + HMAC 签名（第三方/自动化调用）
' 返回: True（通过） / False（拒绝）
' 注意: 不调用 Response.End，由调用方处理
' ============================================
Function API_AuthCheck()
    If Not API_AUTH_ENABLED Then
        API_AuthCheck = True
        Exit Function
    End If
    
    ' 模式1: Session 认证（已登录）
    If Session("UserID") <> "" And CLng(Session("UserID")) > 0 Then
        API_AuthCheck = True
        Exit Function
    End If
    
    ' 模式2: API Key + HMAC 签名
    Dim apiKey, timestamp, signature
    
    ' 从 HTTP 请求头获取认证信息
    apiKey = Request.ServerVariables("HTTP_" & Replace(API_KEY_HEADER, "-", "_"))
    timestamp = Request.ServerVariables("HTTP_" & Replace(API_TS_HEADER, "-", "_"))
    signature = Request.ServerVariables("HTTP_" & Replace(API_SIGN_HEADER, "-", "_"))
    
    ' 如果没有提供任何认证头，拒绝
    If apiKey = "" And signature = "" Then
        API_AuthCheck = False
        Exit Function
    End If
    
    ' 验证时间戳（5分钟内有效，防重放攻击）
    If timestamp = "" Or Not IsNumeric(timestamp) Then
        API_AuthCheck = False
        Exit Function
    End If
    
    Dim nowTimestamp, timeDiff
    nowTimestamp = API_GetUnixTimestamp()
    timeDiff = Abs(CLng(nowTimestamp) - CLng(timestamp))
    If timeDiff > API_SIGN_TIMEOUT Then
        API_AuthCheck = False
        Exit Function
    End If
    
    ' 验证 API Key
    If apiKey = "" Then
        API_AuthCheck = False
        Exit Function
    End If
    
    If Not API_ValidateKey(apiKey) Then
        API_AuthCheck = False
        Exit Function
    End If
    
    ' 验证 HMAC 签名（如果提供了签名头）
    If signature <> "" Then
        Dim expectedSig
        expectedSig = API_GenerateSignature(apiKey, timestamp, Request.ServerVariables("REQUEST_METHOD"))
        If StrComp(signature, expectedSig, 1) <> 0 Then  ' vbTextCompare=1
            API_AuthCheck = False
            Exit Function
        End If
    End If
    
    API_AuthCheck = True
End Function

' ============================================
' API_IsAuthenticated(): 检查是否已认证（非阻塞）
' 返回: Boolean
' ============================================
Function API_IsAuthenticated()
    If Not API_AUTH_ENABLED Then
        API_IsAuthenticated = True
        Exit Function
    End If
    
    ' Session 认证
    If Session("UserID") <> "" And CLng(Session("UserID")) > 0 Then
        API_IsAuthenticated = True
        Exit Function
    End If
    
    ' API Key 认证
    Dim apiKey
    apiKey = Request.ServerVariables("HTTP_" & Replace(API_KEY_HEADER, "-", "_"))
    API_IsAuthenticated = (apiKey <> "" And API_ValidateKey(apiKey))
End Function

' ============================================
' API_ValidateKey(): 验证 API Key 是否有效
' ============================================
Function API_ValidateKey(apiKey)
    ' 内置服务密钥（内部微服务通信）
    If apiKey = API_INTERNAL_KEY Then
        API_ValidateKey = True
        Exit Function
    End If
    
    ' 简单前缀验证（可扩展为数据库查询）
    If InStr(apiKey, "V18_API_") = 1 Then
        API_ValidateKey = True
        Exit Function
    End If
    
    ' 从数据库 ApiKeys 表验证
    On Error Resume Next
    Dim count
    count = CLng(DAL_GetScalar("SELECT COUNT(*) FROM ApiKeys WHERE ApiKey=@Key AND IsActive=1 AND (ExpiresAt IS NULL OR ExpiresAt > GETDATE())", _
        Array(Array("@Key", DAL_adVarChar, 100, apiKey)), 0))
    If Err.Number = 0 And count > 0 Then
        API_ValidateKey = True
    Else
        API_ValidateKey = False
    End If
    On Error GoTo 0
End Function

' ============================================
' API_GenerateSignature(): 生成 HMAC 签名
' ============================================
Function API_GenerateSignature(apiKey, timestamp, method)
    Dim payload
    payload = method & Chr(10) & apiKey & Chr(10) & CStr(timestamp)
    API_GenerateSignature = API_ComputeHMAC(payload, API_AUTH_SECRET)
End Function

' ============================================
' API_CheckCSRF: CSRF Token 验证（V16增强版）
' ============================================
Function API_CheckCSRF()
    Dim clientToken, serverToken
    clientToken = Request.ServerVariables("HTTP_X_CSRF_TOKEN")
    If clientToken = "" Then
        clientToken = Trim(Request.Form("csrf_token"))
    End If
    If clientToken = "" Then
        clientToken = Trim(Request.QueryString("csrf_token"))
    End If
    
    serverToken = Session("CSRFToken")
    If serverToken = "" Or clientToken = "" Then
        API_CheckCSRF = False
        Exit Function
    End If
    
    ' 主 token 验证
    If clientToken = serverToken Then
        API_CheckCSRF = True
        Exit Function
    End If
    
    ' 历史 token 池验证（V17增强）
    Dim historyTokens, i
    historyTokens = Session("CSRFTokenHistory")
    If IsArray(historyTokens) Then
        For i = 0 To UBound(historyTokens)
            If clientToken = historyTokens(i) Then
                API_CheckCSRF = True
                Exit Function
            End If
        Next
    End If
    
    API_CheckCSRF = False
End Function

' ============================================
' API_GetUnixTimestamp: 获取 Unix 时间戳
' ============================================
Function API_GetUnixTimestamp()
    API_GetUnixTimestamp = CLng(DateDiff("s", "1970-01-01 00:00:00", Now()))
End Function

' ============================================
' API_ComputeHMAC: 简化 HMAC 签名计算
' ============================================
Function API_ComputeHMAC(message, secret)
    Dim combined, hash
    combined = message & secret
    hash = MD5Hash(combined)
    API_ComputeHMAC = Left(UCase(hash), 64)
End Function

' ============================================
' API_GenerateKey: 生成 API Key（管理用）
' ============================================
Function API_GenerateKey(seed)
    If seed = "" Then seed = CStr(Timer() & Now())
    API_GenerateKey = "V18_API_" & UCase(Left(MD5Hash(seed & API_AUTH_SECRET), 32))
End Function

' ============================================
' MD5Hash(): 计算 MD5 哈希
' 优先使用 .NET Cryptography，降级为简单哈希
' ============================================
Function MD5Hash(str)
    On Error Resume Next
    Dim md5, bytes, i, result
    Set md5 = Server.CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    If Err.Number <> 0 Then
        Err.Clear
        ' 降级：简单的 FNV-1a 风格哈希
        Dim hashVal, chVal, kIdx
        hashVal = 2166136261
        For kIdx = 1 To Len(str)
            chVal = Asc(Mid(str, kIdx, 1))
            hashVal = hashVal Xor chVal
            hashVal = hashVal * 16777619
            ' 保持在 32-bit 范围内
            If hashVal < 0 Then hashVal = hashVal And &H7FFFFFFF
            If hashVal >= 4294967296 Then hashVal = hashVal - 4294967296
        Next
        MD5Hash = LCase(Right("0000000" & Hex(hashVal), 8))
        Exit Function
    End If
    
    bytes = md5.ComputeHash_2(GetBytes_String(str))
    result = ""
    For i = 0 To UBound(bytes)
        result = result & Right("0" & Hex(bytes(i)), 2)
    Next
    MD5Hash = LCase(result)
    Set md5 = Nothing
    On Error GoTo 0
End Function

' ============================================
' 辅助：字符串 → 字节数组
' ============================================
Function GetBytes_String(str)
    Dim i, bytes()
    ReDim bytes(Len(str) - 1)
    For i = 1 To Len(str)
        bytes(i - 1) = AscB(MidB(str, i, 1))
    Next
    GetBytes_String = bytes
End Function
%>
