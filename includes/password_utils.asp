<%
' ============================================
' 密码处理工具函数 (安全加固版 V2)
' ============================================
' 全局盐值 - 已移至 config.asp (Const PASSWORD_PEPPER)
' 注意：确保 config.asp 在此文件之前被 include

' ============================================
' V2: 增强型密码哈希 (推荐使用)
' 使用多轮迭代哈希 + 盐值 + Pepper
' ============================================
Function HashPasswordV2(password)
    Dim salted, i
    ' 第一层: 混合密码与Pepper
    salted = password & PASSWORD_PEPPER
    ' 多次迭代哈希增加破解难度
    For i = 1 To 10
        salted = InternalHash(salted)
    Next
    ' 添加版本标识以便未来升级
    HashPasswordV2 = "V2_" & salted
End Function

' 内部哈希函数 (迭代使用)
Function InternalHash(inputStr)
    Dim hash, char, i, prev
    hash = ""
    prev = 0
    For i = 1 To Len(inputStr)
        char = Asc(Mid(inputStr, i, 1))
        ' 混合当前字符与前一结果，增加非线性
        prev = (prev + char) Mod 256
        hash = hash & Hex((char * 7 + prev) Mod 256)
    Next
    InternalHash = LCase(hash)
End Function

' V1: 原有简单哈希 (保持向后兼容)
Function GenerateSimpleHash(inputStr)
    Dim hash, char, i
    hash = ""
    For i = 1 To Len(inputStr)
        char = Mid(inputStr, i, 1)
        hash = hash & Hex(Asc(char))
    Next
    GenerateSimpleHash = LCase(hash)
End Function

' 生成密码哈希 (自动选型)
Function HashPassword(password)
    ' 默认使用V2增强版
    HashPassword = HashPasswordV2(password)
End Function

' 验证密码 (同时支持V1和V2)
Function VerifyPassword(password, storedHash)
    Dim inputHash
    ' 检测哈希版本
    If Left(storedHash, 3) = "V2_" Then
        ' V2验证
        inputHash = HashPasswordV2(password)
        VerifyPassword = (inputHash = storedHash)
    Else
        ' V1向后兼容验证
        inputHash = GenerateSimpleHash(password)
        VerifyPassword = (inputHash = storedHash)
    End If
End Function

' 生成随机重置令牌 (增强版)
Function GenerateResetToken()
    Dim token, i, rndNum
    Randomize
    token = ""
    For i = 1 To 32
        ' 使用更复杂的字符集
        rndNum = Int(Rnd() * 62)
        If rndNum < 10 Then
            token = token & Chr(rndNum + 48) ' 0-9
        ElseIf rndNum < 36 Then
            token = token & Chr(rndNum + 55) ' A-Z
        Else
            token = token & Chr(rndNum + 61) ' a-z
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

' 检查是否需要升级密码哈希 (V1->V2)
Function NeedsPasswordUpgrade(storedHash)
    If Left(storedHash, 3) = "V2_" Then
        NeedsPasswordUpgrade = False
    Else
        NeedsPasswordUpgrade = True
    End If
End Function

' 管理员登录兼容验证 (支持V1/V2并自动升级)
Function AdminVerifyAndUpgrade(password, storedHash, adminId)
    Dim inputHashV2, inputHashV1
    
    ' 先尝试V2验证
    inputHashV2 = HashPasswordV2(password)
    If inputHashV2 = storedHash Then
        AdminVerifyAndUpgrade = True
        Exit Function
    End If
    
    ' 再尝试V1验证 (向后兼容)
    If Left(storedHash, 3) <> "V2_" Then
        inputHashV1 = GenerateSimpleHash(password)
        If inputHashV1 = storedHash Then
            ' V1匹配成功，自动升级到V2
            On Error Resume Next
            Call UpgradePasswordHash(adminId, password)
            On Error GoTo 0
            AdminVerifyAndUpgrade = True
            Exit Function
        End If
    End If
    
    AdminVerifyAndUpgrade = False
End Function

' 升级密码哈希 (V1->V2)
Sub UpgradePasswordHash(adminId, password)
    Dim newHash, sql
    newHash = HashPasswordV2(password)
    sql = "UPDATE AdminUsers SET PasswordHash = '" & SafeSQL(newHash) & "' WHERE AdminID = " & CLng(adminId)
    ExecuteNonQuery sql
End Sub
%>