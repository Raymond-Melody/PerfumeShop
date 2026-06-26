<%
' ============================================
' 密码处理工具函数 (安全加固版 V3)
' V15.0: 新增 SHA-256 多轮迭代哈希，向后兼容V1/V2
' ============================================
' 全局盐值 - 已移至 config.asp (Const PASSWORD_PEPPER)
' 注意：确保 config.asp 在此文件之前被 include

' ============================================
' V17.2: 密码输入标准化 - 全角↔半角双向转换
' 解决用户误输入全角/半角标点符号导致的密码验证失败
' Normalize: 全角→半角  (！→!  ＠→@  Ａ→A)
' Expand:    半角→全角  (!→！  @→＠  A→Ａ)
' ============================================
Function NormalizePasswordInput(pwd)
    If IsNull(pwd) Or pwd = "" Then
        NormalizePasswordInput = ""
        Exit Function
    End If
    
    Dim result, i, charCode
    result = ""
    For i = 1 To Len(pwd)
        charCode = AscW(Mid(pwd, i, 1))
        ' 全角ASCII标点 U+FF01~U+FF5E → 半角 U+0021~U+007E
        ' 偏移量: &HFF01 - &H0021 = &HFEE0
        If charCode >= &HFF01 And charCode <= &HFF5E Then
            result = result & ChrW(charCode - &HFEE0)
        ' 全角空格 U+3000 → 半角空格 U+0020
        ElseIf charCode = &H3000 Then
            result = result & ChrW(&H20)
        Else
            result = result & Mid(pwd, i, 1)
        End If
    Next
    NormalizePasswordInput = result
End Function

' 半角→全角转换（反向标准化，用于匹配旧全角哈希）
' 仅转换标点符号，不转换字母和数字
Function ExpandToFullwidth(pwd)
    If IsNull(pwd) Or pwd = "" Then
        ExpandToFullwidth = ""
        Exit Function
    End If
    
    Dim result, i, charCode
    result = ""
    For i = 1 To Len(pwd)
        charCode = AscW(Mid(pwd, i, 1))
        ' 只转换ASCII标点符号（非字母非数字）→ 全角
        ' 范围: 0x21-0x2F, 0x3A-0x40, 0x5B-0x60, 0x7B-0x7E
        If (charCode >= &H21 And charCode <= &H2F) Or _
           (charCode >= &H3A And charCode <= &H40) Or _
           (charCode >= &H5B And charCode <= &H60) Or _
           (charCode >= &H7B And charCode <= &H7E) Then
            result = result & ChrW(charCode + &HFEE0)
        ' 半角空格 U+0020 → 全角空格 U+3000
        ElseIf charCode = &H20 Then
            result = result & ChrW(&H3000)
        Else
            result = result & Mid(pwd, i, 1)
        End If
    Next
    ExpandToFullwidth = result
End Function

' ============================================
' V3: SHA-256 密码哈希 (V15推荐)
' 使用 1000轮 SHA-256 + 随机Salt + Pepper
' 格式: V3$<salt>$<hash>
' ============================================
Function HashPasswordV3(password)
    Dim salt, combined, hash, i
    
    ' V17.2: 密码标准化 (全角→半角)
    password = NormalizePasswordInput(password)
    
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
' V17.1: 优先使用 SQL Server HASHBYTES（可靠、确定性）
'    VBScript SHA-256 因整数溢出不可靠，DJB2 因浮点精度非确定性
'    回退顺序: SQL HASHBYTES → VBScript → DJB2 (仅最后手段)
' ============================================
Function HASH_SHA256(inputStr)
    ' V17.1: 途径1 - SQL Server HASHBYTES（最可靠，真正的 SHA-256）
    On Error Resume Next
    If IsObject(conn) Then
        Dim rs
        Set rs = conn.Execute("SELECT CONVERT(VARCHAR(64), HASHBYTES('SHA2_256', '" & SafeSQL(inputStr) & "'), 2)")
        If Err.Number = 0 And Not rs Is Nothing Then
            If Not rs.EOF Then
                HASH_SHA256 = LCase(rs(0))
                rs.Close : Set rs = Nothing
                On Error GoTo 0
                Exit Function
            End If
            rs.Close
        End If
        Set rs = Nothing
        Err.Clear
    End If
    
    ' 途径2: VBScript SafeSHA256Hash
    HASH_SHA256 = SafeSHA256Hash(inputStr)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    
    ' 途径3: DJB2 非加密回退（仅当上面都失败时）
    If HASH_SHA256 = "" Then
        HASH_SHA256 = DJB2Hash(inputStr)
    End If
End Function

' ============================================
' V2: 增强型密码哈希 (保留向后兼容)
' ============================================
Function HashPasswordV2(password)
    Dim salted, i
    ' V17.2: 密码标准化 (全角→半角)
    password = NormalizePasswordInput(password)
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
    ' V17.2: 密码标准化 (全角→半角)
    password = NormalizePasswordInput(password)
    If FEATURE_PASSWORD_V3 Then
        HashPassword = HashPasswordV3(password)
    Else
        HashPassword = HashPasswordV2(password)
    End If
End Function

' ============================================
' 验证密码 (同时支持V1/V2/V3 + V17.2全角/半角双向兼容)
' 依次尝试: 原始密码 → 全角→半角 → 半角→全角
' ============================================
Function VerifyPassword(password, storedHash)
    Dim inputHash, normalized, expanded
    
    ' V3验证
    If Left(storedHash, 3) = "V3$" Then
        ' 1. 原始密码（向后兼容）
        inputHash = HashPasswordV3FromSalt(password, storedHash)
        If inputHash = storedHash Then VerifyPassword = True : Exit Function
        
        ' 2. 全角→半角标准化
        normalized = NormalizePasswordInput(password)
        If normalized <> password Then
            inputHash = HashPasswordV3FromSalt(normalized, storedHash)
            If inputHash = storedHash Then VerifyPassword = True : Exit Function
        End If
        
        ' 3. 半角→全角扩展（匹配旧全角哈希）
        expanded = ExpandToFullwidth(password)
        If expanded <> password And expanded <> normalized Then
            inputHash = HashPasswordV3FromSalt(expanded, storedHash)
            VerifyPassword = (inputHash = storedHash)
        Else
            VerifyPassword = False
        End If
        Exit Function
    End If
    
    ' V2验证
    If Left(storedHash, 3) = "V2_" Then
        ' 1. 原始密码
        inputHash = HashPasswordV2(password)
        If inputHash = storedHash Then VerifyPassword = True : Exit Function
        
        ' 2. 全角→半角标准化
        normalized = NormalizePasswordInput(password)
        If normalized <> password Then
            inputHash = HashPasswordV2(normalized)
            If inputHash = storedHash Then VerifyPassword = True : Exit Function
        End If
        
        ' 3. 半角→全角扩展
        expanded = ExpandToFullwidth(password)
        If expanded <> password And expanded <> normalized Then
            inputHash = HashPasswordV2(expanded)
            VerifyPassword = (inputHash = storedHash)
        Else
            VerifyPassword = False
        End If
        Exit Function
    End If
    
    ' V1向后兼容验证
    ' 1. 原始密码
    inputHash = GenerateSimpleHash(password)
    If inputHash = storedHash Then VerifyPassword = True : Exit Function
    
    ' 2. 全角→半角标准化
    normalized = NormalizePasswordInput(password)
    If normalized <> password Then
        inputHash = GenerateSimpleHash(normalized)
        If inputHash = storedHash Then VerifyPassword = True : Exit Function
    End If
    
    ' 3. 半角→全角扩展
    expanded = ExpandToFullwidth(password)
    If expanded <> password And expanded <> normalized Then
        inputHash = GenerateSimpleHash(expanded)
        VerifyPassword = (inputHash = storedHash)
    Else
        VerifyPassword = False
    End If
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

' 管理员登录兼容验证 (支持V1/V2/V3并自动升级 + V17.2全角/半角双向兼容)
Function AdminVerifyAndUpgrade(password, storedHash, adminId)
    Dim inputHash, normalized, expanded
    
    ' 尝试V3验证
    If Left(storedHash, 3) = "V3$" Then
        ' 1. 原始密码
        If HashPasswordV3FromSalt(password, storedHash) = storedHash Then
            AdminVerifyAndUpgrade = True : Exit Function
        End If
        
        ' 2. 全角→半角标准化
        normalized = NormalizePasswordInput(password)
        If normalized <> password Then
            If HashPasswordV3FromSalt(normalized, storedHash) = storedHash Then
                Call UpgradePasswordHash(adminId, normalized)
                AdminVerifyAndUpgrade = True : Exit Function
            End If
        End If
        
        ' 3. 半角→全角扩展（匹配旧全角哈希）
        expanded = ExpandToFullwidth(password)
        If expanded <> password And expanded <> normalized Then
            If HashPasswordV3FromSalt(expanded, storedHash) = storedHash Then
                ' 旧全角哈希匹配，升级为标准化（半角）
                Call UpgradePasswordHash(adminId, normalized)
                AdminVerifyAndUpgrade = True : Exit Function
            End If
        End If
        
        AdminVerifyAndUpgrade = False
        Exit Function
    End If
    
    ' 尝试V2验证
    ' 1. 原始密码
    inputHash = HashPasswordV2(password)
    If inputHash = storedHash Then
        If FEATURE_PASSWORD_V3 Then Call UpgradePasswordHash(adminId, password)
        AdminVerifyAndUpgrade = True : Exit Function
    End If
    
    ' 2. 全角→半角标准化
    normalized = NormalizePasswordInput(password)
    If normalized <> password Then
        inputHash = HashPasswordV2(normalized)
        If inputHash = storedHash Then
            If FEATURE_PASSWORD_V3 Then Call UpgradePasswordHash(adminId, normalized)
            AdminVerifyAndUpgrade = True : Exit Function
        End If
    End If
    
    ' 3. 半角→全角扩展
    expanded = ExpandToFullwidth(password)
    If expanded <> password And expanded <> normalized Then
        inputHash = HashPasswordV2(expanded)
        If inputHash = storedHash Then
            If FEATURE_PASSWORD_V3 Then Call UpgradePasswordHash(adminId, normalized)
            AdminVerifyAndUpgrade = True : Exit Function
        End If
    End If
    
    ' 尝试V1验证
    If Left(storedHash, 3) <> "V2_" And Left(storedHash, 3) <> "V3$" Then
        ' 1. 原始密码
        inputHash = GenerateSimpleHash(password)
        If inputHash = storedHash Then
            On Error Resume Next : Call UpgradePasswordHash(adminId, password) : On Error GoTo 0
            AdminVerifyAndUpgrade = True : Exit Function
        End If
        
        ' 2. 全角→半角标准化
        If normalized = "" Then normalized = NormalizePasswordInput(password)
        If normalized <> password Then
            inputHash = GenerateSimpleHash(normalized)
            If inputHash = storedHash Then
                On Error Resume Next : Call UpgradePasswordHash(adminId, normalized) : On Error GoTo 0
                AdminVerifyAndUpgrade = True : Exit Function
            End If
        End If
        
        ' 3. 半角→全角扩展
        If expanded = "" Then expanded = ExpandToFullwidth(password)
        If expanded <> password And expanded <> normalized Then
            inputHash = GenerateSimpleHash(expanded)
            If inputHash = storedHash Then
                On Error Resume Next : Call UpgradePasswordHash(adminId, normalized) : On Error GoTo 0
                AdminVerifyAndUpgrade = True : Exit Function
            End If
        End If
    End If
    
    AdminVerifyAndUpgrade = False
End Function

' 升级密码哈希到当前推荐版本
Sub UpgradePasswordHash(adminId, password)
    Dim newHash, sql, params(1)
    If FEATURE_PASSWORD_V3 Then
        newHash = HashPasswordV3(password)
    Else
        newHash = HashPasswordV2(password)
    End If
    ' V17: 参数化查询防止SQL注入
    sql = "UPDATE AdminUsers SET PasswordHash=@PasswordHash WHERE AdminID=@AdminID"
    params(0) = Array("@PasswordHash", DAL_adVarChar, 255, newHash)
    params(1) = Array("@AdminID", DAL_adInteger, 0, CLng(adminId))
    DAL_Execute sql, params
End Sub
%>