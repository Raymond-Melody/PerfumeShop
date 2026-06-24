<%
' ============================================
' 密码处理工具函数 (安全加固版 V3)
' V15.0: 新增 SHA-256 多轮迭代哈希，向后兼容V1/V2
' ============================================
' 全局盐值 - 已移至 config.asp (Const PASSWORD_PEPPER)
' 注意：确保 config.asp 在此文件之前被 include

' ============================================
' V3: SHA-256 密码哈希 (V15推荐)
' 使用 1000轮 SHA-256 + 随机Salt + Pepper
' 格式: V3$<salt>$<hash>
' ============================================
Function HashPasswordV3(password)
    Dim salt, combined, hash, i
    
    ' 生成32字符随机盐
    salt = GenerateSaltV3(32)
    
    ' 第一轮：salt + password + pepper
    combined = salt & password & PASSWORD_PEPPER
    hash = HASH_SHA256(combined)
    
    ' 999轮额外迭代 = 总共1000轮 (key stretching)
    For i = 1 To 999
        hash = HASH_SHA256(hash & salt)
    Next
    
    HashPasswordV3 = "V3$" & salt & "$" & hash
End Function

' ============================================
' V3随机盐生成器 (使用加密安全种子)
' ============================================
Function GenerateSaltV3(length)
    Dim salt, i, r
    Randomize
    salt = ""
    For i = 1 To length
        r = Int(Rnd() * 62)
        If r < 10 Then
            salt = salt & Chr(r + 48)  ' 0-9
        ElseIf r < 36 Then
            salt = salt & Chr(r + 55)  ' A-Z
        Else
            salt = salt & Chr(r + 61)  ' a-z
        End If
    Next
    GenerateSaltV3 = salt
End Function

' ============================================
' SHA-256 哈希包装器
' 优先使用 connection.asp 中的 SafeSHA256Hash (VBScript实现)
' 不可用时回退到 SQL Server HASHBYTES
' ============================================
Function HASH_SHA256(inputStr)
    On Error Resume Next
    ' 途径1: 使用 connection.asp 中的 SafeSHA256Hash
    HASH_SHA256 = SafeSHA256Hash(inputStr)
    If Err.Number <> 0 Or HASH_SHA256 = "" Then
        Err.Clear
        ' 途径2: 使用 SQL Server HASHBYTES
        If IsObject(conn) Then
            Dim rs
            Set rs = conn.Execute("SELECT CONVERT(VARCHAR(64), HASHBYTES('SHA2_256', '" & SafeSQL(inputStr) & "'), 2)")
            If Not rs Is Nothing And Not rs.EOF Then
                HASH_SHA256 = LCase(rs(0))
                rs.Close
            End If
            Set rs = Nothing
        End If
    End If
    On Error GoTo 0
    
    ' 途径3: 完全回退到DJB2 (比V2的模256循环强)
    If HASH_SHA256 = "" Then
        HASH_SHA256 = DJB2Hash(inputStr)
    End If
End Function

' ============================================
' V2: 增强型密码哈希 (保留向后兼容)
' ============================================
Function HashPasswordV2(password)
    Dim salted, i
    salted = password & PASSWORD_PEPPER
    For i = 1 To 10
        salted = InternalHash(salted)
    Next
    HashPasswordV2 = "V2_" & salted
End Function

' 内部哈希函数 (迭代使用) - V2自定义算法
Function InternalHash(inputStr)
    Dim hash, char, i, prev
    hash = ""
    prev = 0
    For i = 1 To Len(inputStr)
        char = Asc(Mid(inputStr, i, 1))
        prev = (prev + char) Mod 256
        hash = hash & Hex((char * 7 + prev) Mod 256)
    Next
    InternalHash = LCase(hash)
End Function

' V1: 原有简单哈希 (保持向后兼容, 已弃用)
Function GenerateSimpleHash(inputStr)
    Dim hash, char, i
    hash = ""
    For i = 1 To Len(inputStr)
        char = Mid(inputStr, i, 1)
        hash = hash & Hex(Asc(char))
    Next
    GenerateSimpleHash = LCase(hash)
End Function

' ============================================
' 生成密码哈希 (自动选型: V3优先)
' ============================================
Function HashPassword(password)
    If FEATURE_PASSWORD_V3 Then
        HashPassword = HashPasswordV3(password)
    Else
        HashPassword = HashPasswordV2(password)
    End If
End Function

' ============================================
' 验证密码 (同时支持V1/V2/V3)
' ============================================
Function VerifyPassword(password, storedHash)
    Dim inputHash
    
    ' V3验证
    If Left(storedHash, 3) = "V3$" Then
        inputHash = HashPasswordV3FromSalt(password, storedHash)
        VerifyPassword = (inputHash = storedHash)
        Exit Function
    End If
    
    ' V2验证
    If Left(storedHash, 3) = "V2_" Then
        inputHash = HashPasswordV2(password)
        VerifyPassword = (inputHash = storedHash)
        Exit Function
    End If
    
    ' V1向后兼容验证
    inputHash = GenerateSimpleHash(password)
    VerifyPassword = (inputHash = storedHash)
End Function

' ============================================
' V3: 用已有hash中的salt重新计算哈希 (用于验证)
' ============================================
Function HashPasswordV3FromSalt(password, storedHash)
    Dim parts, salt, combined, hash, i
    parts = Split(storedHash, "$")
    If UBound(parts) < 2 Then
        HashPasswordV3FromSalt = ""
        Exit Function
    End If
    salt = parts(1)
    combined = salt & password & PASSWORD_PEPPER
    hash = HASH_SHA256(combined)
    For i = 1 To 999
        hash = HASH_SHA256(hash & salt)
    Next
    HashPasswordV3FromSalt = "V3$" & salt & "$" & hash
End Function

' 生成随机重置令牌 (增强版)
Function GenerateResetToken()
    Dim token, i, rndNum
    Randomize
    token = ""
    For i = 1 To 32
        rndNum = Int(Rnd() * 62)
        If rndNum < 10 Then
            token = token & Chr(rndNum + 48)
        ElseIf rndNum < 36 Then
            token = token & Chr(rndNum + 55)
        Else
            token = token & Chr(rndNum + 61)
        End If
    Next
    GenerateResetToken = token
End Function

' 检查重置令牌是否过期
Function IsTokenExpired(expiryDate)
    If IsNull(expiryDate) Or expiryDate = "" Then
        IsTokenExpired = True
        Exit Function
    End If
    IsTokenExpired = (Now() > expiryDate)
End Function

' ============================================
' 检查是否需要升级密码哈希
' V1/V2 → V3 自动升级检测
' ============================================
Function NeedsPasswordUpgrade(storedHash)
    ' V3是最新，无需升级
    If Left(storedHash, 3) = "V3$" Then
        NeedsPasswordUpgrade = False
    ElseIf FEATURE_PASSWORD_V3 Then
        ' V3已启用但不是V3格式，需要升级
        NeedsPasswordUpgrade = True
    Else
        ' V3未启用，维持现有版本
        NeedsPasswordUpgrade = False
    End If
End Function

' 管理员登录兼容验证 (支持V1/V2/V3并自动升级)
Function AdminVerifyAndUpgrade(password, storedHash, adminId)
    Dim inputHash
    
    ' 尝试V3验证
    If Left(storedHash, 3) = "V3$" Then
        AdminVerifyAndUpgrade = (HashPasswordV3FromSalt(password, storedHash) = storedHash)
        Exit Function
    End If
    
    ' 尝试V2验证
    inputHash = HashPasswordV2(password)
    If inputHash = storedHash Then
        ' V2验证成功，检查是否需要升级到V3
        If FEATURE_PASSWORD_V3 Then
            Call UpgradePasswordHash(adminId, password)
        End If
        AdminVerifyAndUpgrade = True
        Exit Function
    End If
    
    ' 尝试V1验证
    If Left(storedHash, 3) <> "V2_" And Left(storedHash, 3) <> "V3$" Then
        inputHash = GenerateSimpleHash(password)
        If inputHash = storedHash Then
            ' V1匹配，升级到当前推荐版本
            On Error Resume Next
            Call UpgradePasswordHash(adminId, password)
            On Error GoTo 0
            AdminVerifyAndUpgrade = True
            Exit Function
        End If
    End If
    
    AdminVerifyAndUpgrade = False
End Function

' 升级密码哈希到当前推荐版本
Sub UpgradePasswordHash(adminId, password)
    Dim newHash, sql
    If FEATURE_PASSWORD_V3 Then
        newHash = HashPasswordV3(password)
    Else
        newHash = HashPasswordV2(password)
    End If
    sql = "UPDATE AdminUsers SET PasswordHash = '" & SafeSQL(newHash) & "' WHERE AdminID = " & CLng(adminId)
    ExecuteNonQuery sql
End Sub
%>