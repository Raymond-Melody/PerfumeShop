<%
' ============================================
' 会员等级与积分管理 - Member Utils
' 管理会员等级、积分、特权
' ============================================

' ============================================
' 获取会员等级配置
' ============================================
Function MU_GetLevelConfig()
    Dim rs, levels, lv, minSpent
    Set levels = Server.CreateObject("Scripting.Dictionary")
    
    ' 从 SiteSettings 读取等级配置
    On Error Resume Next
    Set rs = conn.Execute("SELECT SettingKey, SettingValue FROM SiteSettings WHERE SettingKey LIKE 'MemberLevel_%' ORDER BY SettingKey")
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            lv = Replace(rs("SettingKey"), "MemberLevel_", "")
            minSpent = rs("SettingValue")
            If IsNumeric(minSpent) Then
                levels.Add lv, CDbl(minSpent)
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
    
    ' 默认等级配置（如果数据库未设置）
    If levels.Count = 0 Then
        levels.Add "L0", 0       ' 普通会员
        levels.Add "L1", 500     ' 银卡会员 (累计消费≥500)
        levels.Add "L2", 2000    ' 金卡会员 (累计消费≥2000)
        levels.Add "L3", 5000    ' 铂金会员 (累计消费≥5000)
        levels.Add "L4", 15000   ' 钻石会员 (累计消费≥15000)
    End If
    
    Set MU_GetLevelConfig = levels
End Function

' ============================================
' 计算用户会员等级
' ============================================
Function MU_CalcUserLevel(userId)
    Dim levels, totalSpent, maxLv, lv, threshold
    
    totalSpent = 0
    maxLv = "L0"
    
    On Error Resume Next
    
    ' 获取用户累计消费金额
    Dim rs
    Set rs = conn.Execute("SELECT ISNULL(SUM(TotalAmount), 0) FROM Orders WHERE UserID = " & userId & " AND Status = 'Paid'")
    If Not rs Is Nothing Then
        If Not rs.EOF Then totalSpent = CDbl(rs(0))
        rs.Close
    End If
    Set rs = Nothing
    
    ' 获取等级配置
    Set levels = MU_GetLevelConfig()
    
    ' 找到最高等级
    Dim keys, i
    keys = levels.Keys
    For i = 0 To levels.Count - 1
        lv = keys(i)
        threshold = levels.Item(lv)
        If totalSpent >= threshold Then
            maxLv = lv
        End If
    Next
    
    Set levels = Nothing
    MU_CalcUserLevel = maxLv
End Function

' ============================================
' 获取会员等级折扣率
' ============================================
Function MU_GetLevelDiscount(levelCode)
    Select Case levelCode
        Case "L0": MU_GetLevelDiscount = 1.0    ' 无折扣
        Case "L1": MU_GetLevelDiscount = 0.95   ' 95折
        Case "L2": MU_GetLevelDiscount = 0.90   ' 9折
        Case "L3": MU_GetLevelDiscount = 0.85   ' 85折
        Case "L4": MU_GetLevelDiscount = 0.80   ' 8折
        Case Else: MU_GetLevelDiscount = 1.0
    End Select
End Function

' ============================================
' V18: 获取有效折扣率（自动适配新旧等级系统）
' ============================================
Function MU_GetEffectiveDiscountRate(userId)
    If FEATURE_MEMBER_TIERS Then
        Dim tier
        Set tier = MU_V18_CalcUserTier(userId)
        If Not tier Is Nothing Then
            MU_GetEffectiveDiscountRate = CDbl(tier("discountRate"))
            Set tier = Nothing
        Else
            MU_GetEffectiveDiscountRate = 1.0
        End If
    Else
        Dim level
        level = MU_CalcUserLevel(userId)
        MU_GetEffectiveDiscountRate = MU_GetLevelDiscount(level)
    End If
End Function

' ============================================
' 获取会员等级名称
' ============================================
Function MU_GetLevelName(levelCode)
    Select Case levelCode
        Case "L0": MU_GetLevelName = "普通会员"
        Case "L1": MU_GetLevelName = "银卡会员"
        Case "L2": MU_GetLevelName = "金卡会员"
        Case "L3": MU_GetLevelName = "铂金会员"
        Case "L4": MU_GetLevelName = "钻石会员"
        Case Else: MU_GetLevelName = "普通会员"
    End Select
End Function

' ============================================
' 获取用户当前积分
' ============================================
Function MU_GetUserPoints(userId)
    Dim rs, pts
    pts = 0
    
    On Error Resume Next
    
    ' 优先从 UserPoints 表获取
    Set rs = conn.Execute("SELECT ISNULL(AvailablePoints, 0) FROM UserPoints WHERE UserID = " & userId)
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            pts = rs(0)
        End If
        rs.Close
    End If
    Set rs = Nothing
    
    ' 如果 UserPoints 无数据，回退到 Users.Points
    If pts = 0 Then
        Set rs = conn.Execute("SELECT ISNULL(Points, 0) FROM Users WHERE UserID = " & userId)
        If Not rs Is Nothing Then
            If Not rs.EOF Then pts = rs(0)
            rs.Close
        End If
        Set rs = Nothing
    End If
    
    MU_GetUserPoints = pts
End Function

' ============================================
' 添加用户积分并记录日志
' ============================================
Sub MU_AddPoints(userId, points, reason, orderId)
    On Error Resume Next
    
    If points <= 0 Then Exit Sub
    
    ' 更新 UserPoints
    Dim checkSQL, updateSQL
    checkSQL = "SELECT COUNT(*) FROM UserPoints WHERE UserID = " & userId
    Dim rsChk
    Set rsChk = conn.Execute(checkSQL)
    If Not rsChk Is Nothing Then
        If Not rsChk.EOF Then
            If CLng(rsChk(0)) > 0 Then
                ' 已存在记录，更新
                updateSQL = "UPDATE UserPoints SET AvailablePoints = ISNULL(AvailablePoints, 0) + " & points & _
                          ", TotalPoints = ISNULL(TotalPoints, 0) + " & points & _
                          ", LastUpdatedAt = GETDATE() WHERE UserID = " & userId
                conn.Execute updateSQL
            Else
                ' 新增记录
                updateSQL = "INSERT INTO UserPoints (UserID, AvailablePoints, TotalPoints, LastUpdatedAt) VALUES (" & userId & ", " & points & ", " & points & ", GETDATE())"
                conn.Execute updateSQL
            End If
        End If
        rsChk.Close
    End If
    Set rsChk = Nothing
    
    ' 同步更新 Users.Points
    conn.Execute "UPDATE Users SET Points = ISNULL(Points, 0) + " & points & " WHERE UserID = " & userId
    
    ' 记录积分日志（使用 SiteSettings 辅助记录）
    Dim logKey
    logKey = "PointsLog_" & userId & "_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Hour(Now) & Minute(Now) & Second(Now)
    Dim logVal
    logVal = "UserID:" & userId & "|Points:" & points & "|Reason:" & reason & "|OrderID:" & orderId
    conn.Execute "INSERT INTO SiteSettings (SettingKey, SettingValue) VALUES ('" & SafeSQL(logKey) & "', '" & SafeSQL(logVal) & "')"
End Sub

' ============================================
' 使用积分抵扣金额
' ============================================
Function MU_UsePoints(userId, pointsToUse, orderId)
    Dim available, usedPoints
    usedPoints = 0
    
    available = MU_GetUserPoints(userId)
    If pointsToUse > available Then pointsToUse = available
    
    On Error Resume Next
    
    ' 扣除积分
    Dim updateSQL
    updateSQL = "UPDATE UserPoints SET AvailablePoints = ISNULL(AvailablePoints, 0) - " & pointsToUse & _
              ", UsedPoints = ISNULL(UsedPoints, 0) + " & pointsToUse & _
              ", LastUpdatedAt = GETDATE() WHERE UserID = " & userId
    conn.Execute updateSQL
    
    ' 同步 Users.Points
    conn.Execute "UPDATE Users SET Points = ISNULL(Points, 0) - " & pointsToUse & " WHERE UserID = " & userId
    
    ' 记录使用日志
    Call MU_AddPoints(userId, 0, "Used " & pointsToUse & " points for order #" & orderId, orderId)
    
    MU_UsePoints = pointsToUse * 0.01  ' 1积分=0.01元
End Function

' ============================================
' 渲染会员等级徽章（在用户中心使用）
' ============================================
Sub MU_RenderLevelBadge(userId)
    Dim level, levelName, discount, badgeClass, icon
    
    ' V18: 使用新等级系统
    If FEATURE_MEMBER_TIERS Then
        Dim tier
        Set tier = MU_V18_CalcUserTier(userId)
        If Not tier Is Nothing Then
            levelName = tier("tierName")
            discount = CDbl(tier("discountRate"))
            icon = tier("iconClass")
            badgeClass = "level-" & tier("tierCode")
            Set tier = Nothing
        Else
            levelName = "普通会员"
            discount = 1.0
            icon = "fa-user"
            badgeClass = "level-0"
        End If
    Else
        ' 旧等级系统
        level = MU_CalcUserLevel(userId)
        levelName = MU_GetLevelName(level)
        discount = MU_GetLevelDiscount(level)
        Select Case level
            Case "L0": badgeClass = "level-0": icon = "fa-user"
            Case "L1": badgeClass = "level-1": icon = "fa-star"
            Case "L2": badgeClass = "level-2": icon = "fa-gem"
            Case "L3": badgeClass = "level-3": icon = "fa-crown"
            Case "L4": badgeClass = "level-4": icon = "fa-crown"
        End Select
    End If
%>
<span class="level-badge <%= badgeClass %>">
    <i class="fas <%= icon %>"></i> <%= levelName %>
    <% If discount < 1 Then %><small>(<%= FormatPercent(1 - discount, 0) %> OFF)</small><% End If %>
</span>
<style>
.level-badge { display:inline-flex; align-items:center; gap:5px; padding:4px 12px; border-radius:12px; font-size:13px; font-weight:600; }
.level-0 { background:#f0f0f0; color:#666; }
.level-L0 { background:#f0f0f0; color:#666; }
.level-1 { background:#e8f4f8; color:#2196F3; }
.level-L1 { background:#e8f4f8; color:#2196F3; }
.level-2 { background:#fff3e0; color:#FF9800; }
.level-L2 { background:#fff3e0; color:#FF9800; }
.level-3 { background:#fce4ec; color:#e91e63; }
.level-L3 { background:#fce4ec; color:#e91e63; }
.level-4 { background:#e8f5e9; color:#4CAF50; }
.level-L4 { background:#e8f5e9; color:#4CAF50; }
.level-silver { background:#f5f5f5; color:#9E9E9E; }
.level-gold { background:#fff3e0; color:#FF9800; }
.level-diamond { background:#e3f2fd; color:#2196F3; }
.level-black { background:#fafafa; color:#212121; border:1px solid #212121; }
.level-badge small { opacity:0.7; font-weight:400; }
</style>
<%
End Sub

' ============================================
' V14 会员推荐制 - Referral System
' ============================================

' 推荐Token签名密钥 - 已移至 config.asp (Const REFERRAL_SECRET)
' 注意：确保 config.asp 在此文件之前被 include

' ============================================
' 生成推荐Token
' 输入: referrerUserId - 推荐人用户ID
'       daysValid - 有效天数（默认30天）
'       maxUses - 最大使用次数（默认1次）
'       referrerType - 推荐人类型 user/admin（默认user）
' 输出: 加密的Token字符串（格式: userId|timestamp|expiry|maxUses|nonce|signature）
' ============================================
Function MU_GenerateReferralToken(referrerUserId, daysValid, maxUses, referrerType)
    If IsNull(referrerUserId) Or referrerUserId = "" Then
        MU_GenerateReferralToken = ""
        Exit Function
    End If
    
    If IsNull(daysValid) Or daysValid = "" Or daysValid <= 0 Then daysValid = 30
    If IsNull(maxUses) Or maxUses = "" Or maxUses <= 0 Then maxUses = 1
    If IsNull(referrerType) Or referrerType = "" Then referrerType = "user"
    
    ' 生成时间戳（Unix秒）
    Dim timestamp, expiry, nonce, dataToSign, signature
    timestamp = CLng(DateDiff("s", "1970-01-01 00:00:00", Now()))
    expiry = timestamp + (daysValid * 86400)
    
    ' 生成随机nonce（防重放）
    Randomize
    nonce = Right("00000000" & Hex(Int(Rnd * &H7FFFFFFF)), 8) & _
            Right("00000000" & Hex(Int(Rnd * &H7FFFFFFF)), 8)
    
    ' 构建签名数据
    dataToSign = referrerUserId & "|" & timestamp & "|" & expiry & "|" & maxUses & "|" & nonce & "|" & REFERRAL_SECRET
    signature = SafeSHA256Hash(dataToSign)
    
    ' Token格式: userId|timestamp|expiry|maxUses|nonce|signature
    MU_GenerateReferralToken = referrerUserId & "|" & timestamp & "|" & expiry & "|" & maxUses & "|" & nonce & "|" & signature
End Function

' ============================================
' 存储推荐Token到数据库
' ============================================
Function MU_StoreReferralToken(tokenStr, referrerType)
    Dim parts, tokenHash
    parts = Split(tokenStr, "|")
    If UBound(parts) < 5 Then
        MU_StoreReferralToken = False
        Exit Function
    End If
    
    ' Token哈希用于数据库存储（只存签名部分验证，不存完整Token）
    tokenHash = Left(SafeSHA256Hash(tokenStr), 64)
    
    Dim referrerUserId, expiry, maxUses
    referrerUserId = CLng(parts(0))
    expiry = CLng(parts(2))
    maxUses = CLng(parts(3))
    
    ' 将Unix时间戳转换为SQL Server DATETIME2
    Dim expiryDate
    expiryDate = DateAdd("s", expiry, "1970-01-01 00:00:00")
    Dim expiryStr
    expiryStr = Year(expiryDate) & "-" & Right("0" & Month(expiryDate), 2) & "-" & Right("0" & Day(expiryDate), 2) & " " & _
                Right("0" & Hour(expiryDate), 2) & ":" & Right("0" & Minute(expiryDate), 2) & ":" & Right("0" & Second(expiryDate), 2)
    
    Dim sql
    sql = "INSERT INTO ReferralTokens (TokenHash, OriginalToken, ReferrerUserID, ReferrerType, ExpiresAt, MaxUses, UsedCount, IsActive) VALUES (" & _
        "'" & SafeSQL(tokenHash) & "', '" & SafeSQL(tokenStr) & "', " & referrerUserId & ", '" & SafeSQL(referrerType) & "', '" & expiryStr & "', " & maxUses & ", 0, 1)"
    
    MU_StoreReferralToken = ExecuteNonQuery(sql)
End Function

' ============================================
' 验证推荐Token
' 输入: token - URL传递的Token字符串
' 输出: Dictionary对象，包含 success, referrerUserId, message, tokenHash
' ============================================
Function MU_ValidateReferralToken(token)
    Dim result
    Set result = Server.CreateObject("Scripting.Dictionary")
    result.Add "success", False
    result.Add "referrerUserId", 0
    result.Add "referrerName", ""
    result.Add "message", ""
    result.Add "expiryDate", ""
    result.Add "maxUses", 0
    result.Add "tokenHash", ""
    
    If IsNull(token) Or token = "" Then
        result("message") = "缺少推荐Token"
        Set MU_ValidateReferralToken = result
        Exit Function
    End If
    
    Dim parts, userId, timestamp, expiry, maxUses, nonce, signature
    parts = Split(token, "|")
    If UBound(parts) < 5 Then
        result("message") = "Token格式无效"
        Set MU_ValidateReferralToken = result
        Exit Function
    End If
    
    userId = parts(0)
    timestamp = parts(1)
    expiry = parts(2)
    maxUses = parts(3)
    nonce = parts(4)
    signature = parts(5)
    
    ' 验证各部分为数字
    If Not IsNumeric(userId) Or Not IsNumeric(timestamp) Or Not IsNumeric(expiry) Or Not IsNumeric(maxUses) Then
        result("message") = "Token数据格式无效"
        Set MU_ValidateReferralToken = result
        Exit Function
    End If
    
    ' 验证签名
    Dim dataToSign, expectedSignature
    dataToSign = userId & "|" & timestamp & "|" & expiry & "|" & maxUses & "|" & nonce & "|" & REFERRAL_SECRET
    expectedSignature = SafeSHA256Hash(dataToSign)
    
    If signature <> expectedSignature Then
        result("message") = "Token签名验证失败"
        ' 记录可疑尝试
        Dim attemptIP
        attemptIP = Request.ServerVariables("REMOTE_ADDR")
        ExecuteNonQuery "INSERT INTO RegistrationAttempts (IPAddress, DeviceFingerprint, Success, TokenHash) VALUES ('" & SafeSQL(attemptIP) & "', 'invalid_signature', 0, '" & SafeSQL(Left(token, 64)) & "')"
        Set MU_ValidateReferralToken = result
        Exit Function
    End If
    
    ' 验证是否过期
    Dim currentTime
    currentTime = CLng(DateDiff("s", "1970-01-01 00:00:00", Now()))
    If currentTime > CLng(expiry) Then
        result("message") = "推荐链接已过期"
        Set MU_ValidateReferralToken = result
        Exit Function
    End If
    
    If currentTime < CLng(timestamp) Then
        result("message") = "Token时间异常"
        Set MU_ValidateReferralToken = result
        Exit Function
    End If
    
    ' 计算Token哈希用于数据库查询
    Dim tokenHash
    tokenHash = Left(SafeSHA256Hash(token), 64)
    
    ' 检查数据库中Token是否有效
    Dim rsToken
    Set rsToken = ExecuteQuery("SELECT * FROM ReferralTokens WHERE TokenHash = '" & SafeSQL(tokenHash) & "' AND IsActive = 1")
    If rsToken Is Nothing Or rsToken.EOF Then
        result("message") = "推荐链接不存在或已失效"
        If Not rsToken Is Nothing Then rsToken.Close : Set rsToken = Nothing
        Set MU_ValidateReferralToken = result
        Exit Function
    End If
    
    ' 检查是否超过使用次数
    If CLng(rsToken("UsedCount")) >= CLng(rsToken("MaxUses")) Then
        result("message") = "推荐链接已达到使用上限"
        rsToken.Close : Set rsToken = Nothing
        Set MU_ValidateReferralToken = result
        Exit Function
    End If
    
    ' 检查数据库中的过期时间
    If Not IsNull(rsToken("ExpiresAt")) Then
        ' Handle DATETIME2(7) - must strip fractional seconds entirely (VBScript IsDate/CDate rejects .NNN)
        Dim dbExpiryStr, dbDotPos
        dbExpiryStr = CStr(rsToken("ExpiresAt") & "")
        dbDotPos = InStr(dbExpiryStr, ".")
        If dbDotPos > 0 Then dbExpiryStr = Left(dbExpiryStr, dbDotPos - 1)
        If IsDate(dbExpiryStr) Then
            If CDate(dbExpiryStr) < Now() Then
                result("message") = "推荐链接已过期"
                rsToken.Close : Set rsToken = Nothing
                Set MU_ValidateReferralToken = result
                Exit Function
            End If
        End If
    End If
    
    ' 获取推荐人信息
    Dim rsReferrer
    Set rsReferrer = ExecuteQuery("SELECT Username, FullName FROM Users WHERE UserID = " & CLng(userId) & " AND IsActive <> 0")
    Dim referrerName
    referrerName = ""
    If Not rsReferrer Is Nothing And Not rsReferrer.EOF Then
        If Not IsNull(rsReferrer("FullName")) And rsReferrer("FullName") <> "" Then
            referrerName = rsReferrer("FullName")
        Else
            referrerName = rsReferrer("Username")
        End If
        rsReferrer.Close
    End If
    Set rsReferrer = Nothing
    
    ' 格式化过期日期
    Dim expiryDateObj, expiryStr
    expiryDateObj = DateAdd("s", CLng(expiry), "1970-01-01 00:00:00")
    expiryStr = Year(expiryDateObj) & "年" & Month(expiryDateObj) & "月" & Day(expiryDateObj) & "日"
    
    ' 成功
    result("success") = True
    result("referrerUserId") = CLng(userId)
    result("referrerName") = referrerName
    result("expiryDate") = expiryStr
    result("maxUses") = CLng(maxUses)
    result("tokenHash") = tokenHash
    
    rsToken.Close : Set rsToken = Nothing
    Set MU_ValidateReferralToken = result
End Function

' ============================================
' 标记Token已被使用（注册成功后调用）
' ============================================
Sub MU_MarkTokenUsed(tokenHash)
    If tokenHash <> "" Then
        ExecuteNonQuery "UPDATE ReferralTokens SET UsedCount = ISNULL(UsedCount, 0) + 1 WHERE TokenHash = '" & SafeSQL(tokenHash) & "' AND IsActive = 1"
    End If
End Sub

' ============================================
' 一次性写入推荐祖先链条
' 例如: A推荐了B, B推荐了C
' 当C注册时，写入 (A, C, 2), (B, C, 1)
' ============================================
Sub MU_RecordReferralChain(newUserId, referrerUserId)
    On Error Resume Next
    
    If newUserId <= 0 Or referrerUserId <= 0 Then Exit Sub
    
    ' 写入直接推荐关系（Depth=1）
    ExecuteNonQuery "INSERT INTO ReferralRelations (AncestorUserID, DescendantUserID, Depth) VALUES (" & referrerUserId & ", " & newUserId & ", 1)"
    
    ' 查询推荐人的所有祖先
    Dim rsAncestors
    Set rsAncestors = ExecuteQuery("SELECT AncestorUserID, Depth FROM ReferralRelations WHERE DescendantUserID = " & referrerUserId & " ORDER BY Depth ASC")
    
    If Not rsAncestors Is Nothing Then
        Do While Not rsAncestors.EOF
            Dim ancestorId, ancestorDepth
            ancestorId = rsAncestors("AncestorUserID")
            ancestorDepth = rsAncestors("Depth")
            ' 写入祖先关系：祖先 -> 新用户，Depth = 祖先到推荐人的深度 + 1
            ExecuteNonQuery "INSERT INTO ReferralRelations (AncestorUserID, DescendantUserID, Depth) VALUES (" & ancestorId & ", " & newUserId & ", " & (ancestorDepth + 1) & ")"
            rsAncestors.MoveNext
        Loop
        rsAncestors.Close
    End If
    Set rsAncestors = Nothing
    
    On Error GoTo 0
End Sub

' ============================================
' 检查会员每日生成推荐链接数量
' 限制: 会员每日最多生成5个推荐链接
' ============================================
Function MU_CheckDailyReferralLimit(userId)
    MU_CheckDailyReferralLimit = True ' 默认允许
    
    Dim todayCount
    todayCount = GetScalar("SELECT COUNT(*) FROM ReferralTokens WHERE ReferrerUserID = " & userId & " AND ReferrerType = 'user' AND CAST(CreatedAt AS DATE) = CAST(GETDATE() AS DATE)")
    If IsNull(todayCount) Then todayCount = 0
    
    If CLng(todayCount) >= 5 Then
        MU_CheckDailyReferralLimit = False
    End If
End Function

' ============================================
' 检查注册速率限制
' 输入: ip - 客户端IP地址
'       fingerprint - 设备指纹（来自前端JS）
' 输出: Dictionary，包含 allowed, message
' 规则:
'   - 同一IP每24小时最多3次成功注册
'   - 同一设备指纹每24小时最多2次成功注册
'   - 同一IP每小时内最多5次注册尝试（含失败）
' ============================================
Function MU_CheckRegistrationRateLimit(ip, fingerprint)
    Dim result
    Set result = Server.CreateObject("Scripting.Dictionary")
    result.Add "allowed", True
    result.Add "message", ""
    
    If IsNull(ip) Or ip = "" Then ip = Request.ServerVariables("REMOTE_ADDR")
    
    Dim ipSuccessCount, fpSuccessCount, ipHourlyAttempts
    
    On Error Resume Next
    
    ' 检查IP最近24小时成功注册次数
    ipSuccessCount = GetScalar("SELECT COUNT(*) FROM RegistrationAttempts WHERE IPAddress = '" & SafeSQL(ip) & "' AND Success = 1 AND AttemptedAt >= DATEADD(hour, -24, GETDATE())")
    If IsNull(ipSuccessCount) Then ipSuccessCount = 0
    
    ' 检查设备指纹最近24小时成功注册次数
    If fingerprint <> "" Then
        fpSuccessCount = GetScalar("SELECT COUNT(*) FROM RegistrationAttempts WHERE DeviceFingerprint = '" & SafeSQL(fingerprint) & "' AND Success = 1 AND AttemptedAt >= DATEADD(hour, -24, GETDATE())")
        If IsNull(fpSuccessCount) Then fpSuccessCount = 0
    Else
        fpSuccessCount = 0
    End If
    
    ' 检查IP最近1小时尝试次数（含失败）
    ipHourlyAttempts = GetScalar("SELECT COUNT(*) FROM RegistrationAttempts WHERE IPAddress = '" & SafeSQL(ip) & "' AND AttemptedAt >= DATEADD(hour, -1, GETDATE())")
    If IsNull(ipHourlyAttempts) Then ipHourlyAttempts = 0
    
    On Error GoTo 0
    
    If CLng(ipHourlyAttempts) >= 5 Then
        result("allowed") = False
        result("message") = "注册尝试过于频繁，请1小时后再试"
    ElseIf CLng(ipSuccessCount) >= 3 Then
        result("allowed") = False
        result("message") = "该IP已超过注册次数限制，请24小时后再试"
    ElseIf CLng(fpSuccessCount) >= 2 And fingerprint <> "" Then
        result("allowed") = False
        result("message") = "该设备已超过注册次数限制，请24小时后再试"
    End If
    
    Set MU_CheckRegistrationRateLimit = result
End Function

' ============================================
' 记录注册尝试
' ============================================
Sub MU_RecordRegistrationAttempt(ip, fingerprint, success, tokenHash)
    If IsNull(ip) Or ip = "" Then ip = Request.ServerVariables("REMOTE_ADDR")
    If IsNull(fingerprint) Then fingerprint = ""
    If IsNull(tokenHash) Then tokenHash = ""
    
    Dim successVal
    If success Then successVal = 1 Else successVal = 0
    
    ExecuteNonQuery "INSERT INTO RegistrationAttempts (IPAddress, DeviceFingerprint, Success, TokenHash) VALUES ('" & SafeSQL(ip) & "', '" & SafeSQL(fingerprint) & "', " & successVal & ", '" & SafeSQL(tokenHash) & "')"
End Sub

' ============================================
' 获取用户的推荐统计
' ============================================
Function MU_GetReferralStats(userId)
    Dim stats
    Set stats = Server.CreateObject("Scripting.Dictionary")
    stats.Add "todayLinks", 0
    stats.Add "totalLinks", 0
    stats.Add "activeLinks", 0
    stats.Add "totalInvitees", 0
    
    On Error Resume Next
    
    ' 今日生成链接数
    Dim todayLinks
    todayLinks = GetScalar("SELECT COUNT(*) FROM ReferralTokens WHERE ReferrerUserID = " & userId & " AND CAST(CreatedAt AS DATE) = CAST(GETDATE() AS DATE)")
    If Not IsNull(todayLinks) Then stats("todayLinks") = CLng(todayLinks)
    
    ' 累计生成链接数
    Dim totalLinks
    totalLinks = GetScalar("SELECT COUNT(*) FROM ReferralTokens WHERE ReferrerUserID = " & userId)
    If Not IsNull(totalLinks) Then stats("totalLinks") = CLng(totalLinks)
    
    ' 有效链接数
    Dim activeLinks
    activeLinks = GetScalar("SELECT COUNT(*) FROM ReferralTokens WHERE ReferrerUserID = " & userId & " AND IsActive = 1 AND ExpiresAt > GETDATE() AND UsedCount < MaxUses")
    If Not IsNull(activeLinks) Then stats("activeLinks") = CLng(activeLinks)
    
    ' 已邀请人数
    Dim totalInvitees
    totalInvitees = GetScalar("SELECT COUNT(*) FROM ReferralRelations WHERE AncestorUserID = " & userId & " AND Depth = 1")
    If Not IsNull(totalInvitees) Then stats("totalInvitees") = CLng(totalInvitees)
    
    On Error GoTo 0
    
    Set MU_GetReferralStats = stats
End Function

' ============================================
' 获取用户的有效推荐链接列表
' ============================================
Function MU_GetActiveReferralLinks(userId)
    Dim sql
    sql = "SELECT rt.TokenHash, rt.OriginalToken, rt.MaxUses, rt.UsedCount, rt.ExpiresAt, rt.CreatedAt, rt.IsActive, " & _
          "u.Username AS ReferrerName FROM ReferralTokens rt " & _
          "LEFT JOIN Users u ON rt.ReferrerUserID = u.UserID " & _
          "WHERE rt.ReferrerUserID = " & userId & " AND rt.IsActive = 1 AND rt.ExpiresAt > GETDATE() AND rt.OriginalToken IS NOT NULL " & _
          "ORDER BY rt.CreatedAt DESC"
    Set MU_GetActiveReferralLinks = ExecuteQuery(sql)
End Function

' ============================================
' 从Token字符串重建完整Token（用于生成注册URL）
' Token已包含所有信息，直接URL编码即可
' ============================================
Function MU_GetTokenURL(tokenStr)
    If tokenStr = "" Then
        MU_GetTokenURL = ""
        Exit Function
    End If
    MU_GetTokenURL = Server.URLEncode(tokenStr)
End Function

' ============================================
' V18 会员等级体系 (Member Tiers V2)
' 基于 MemberTiers/MemberBenefits 数据库表
' 等级: 银卡(0-3000) → 金卡(3000-10000) → 钻石(10000-30000) → 黑金(30000+)
' ============================================

' 获取 V18 等级配置（从 MemberTiers 表读取）
' 返回: Recordset（按 SortOrder 排序）
Function MU_V18_GetAllTiers()
    If Not FEATURE_MEMBER_TIERS Then
        Set MU_V18_GetAllTiers = Nothing
        Exit Function
    End If
    
    On Error Resume Next
    Dim sql
    sql = "SELECT * FROM MemberTiers WHERE IsActive = 1 ORDER BY SortOrder ASC"
    Set MU_V18_GetAllTiers = conn.Execute(sql)
    On Error GoTo 0
End Function

' 从 MemberTiers 表获取单个等级信息
' 返回: Dictionary {tierCode, tierName, minSpent, maxSpent, discountRate, freeShipping, ...}
Function MU_V18_GetTierByCode(tierCode)
    If Not FEATURE_MEMBER_TIERS Then
        Set MU_V18_GetTierByCode = Nothing
        Exit Function
    End If
    
    On Error Resume Next
    Dim sql, rs, dict
    sql = "SELECT * FROM MemberTiers WHERE TierCode = '" & SafeSQL(tierCode) & "' AND IsActive = 1"
    Set rs = conn.Execute(sql)
    If Not rs Is Nothing And Not rs.EOF Then
        Set dict = Server.CreateObject("Scripting.Dictionary")
        dict.Add "tierCode", CStr(rs("TierCode"))
        dict.Add "tierName", CStr(rs("TierName") & "")
        dict.Add "tierNameEN", CStr(rs("TierNameEN") & "")
        dict.Add "minSpent", CDbl(rs("MinSpent"))
        If Not IsNull(rs("MaxSpent")) Then
            dict.Add "maxSpent", CDbl(rs("MaxSpent"))
        Else
            dict.Add "maxSpent", 99999999
        End If
        dict.Add "discountRate", CDbl(rs("DiscountRate"))
        dict.Add "freeShipping", CBool(rs("FreeShipping"))
        dict.Add "priorityShipping", CBool(rs("PriorityShipping"))
        dict.Add "birthdayGift", CBool(rs("BirthdayGift"))
        dict.Add "dedicatedSupport", CBool(rs("DedicatedSupport"))
        dict.Add "iconClass", CStr(rs("IconClass") & "")
        dict.Add "color", CStr(rs("Color") & "")
        dict.Add "badgeBg", CStr(rs("BadgeBg") & "")
        rs.Close
        Set MU_V18_GetTierByCode = dict
    Else
        Set MU_V18_GetTierByCode = Nothing
    End If
    Set rs = Nothing
    On Error GoTo 0
End Function

' 计算用户当前会员等级（V18新算法）
' 返回: Dictionary {tierCode, tierName, discountRate, totalSpent, ...}
Function MU_V18_CalcUserTier(userId)
    Dim tier, totalSpent, rs, rsTiers, currentTier
    
    Set currentTier = Nothing
    totalSpent = 0
    
    If Not FEATURE_MEMBER_TIERS Then
        ' 回退到旧等级系统
        Dim oldLevel
        oldLevel = MU_CalcUserLevel(userId)
        Set currentTier = Server.CreateObject("Scripting.Dictionary")
        currentTier.Add "tierCode", oldLevel
        currentTier.Add "tierName", MU_GetLevelName(oldLevel)
        currentTier.Add "discountRate", MU_GetLevelDiscount(oldLevel)
        currentTier.Add "totalSpent", 0
        currentTier.Add "iconClass", "fa-user"
        currentTier.Add "color", "#666"
        currentTier.Add "badgeBg", "#f0f0f0"
        currentTier.Add "freeShipping", False
        currentTier.Add "priorityShipping", False
        currentTier.Add "birthdayGift", False
        currentTier.Add "dedicatedSupport", False
        Set MU_V18_CalcUserTier = currentTier
        Exit Function
    End If
    
    On Error Resume Next
    
    ' 获取累计消费金额
    Set rs = conn.Execute("SELECT ISNULL(SUM(TotalAmount), 0) FROM Orders WHERE UserID = " & userId & " AND Status IN ('Paid','Shipped','Completed')")
    If Not rs Is Nothing Then
        If Not rs.EOF Then totalSpent = CDbl(rs(0))
        rs.Close
    End If
    Set rs = Nothing
    
    ' 查找对应等级
    Set rsTiers = conn.Execute("SELECT * FROM MemberTiers WHERE IsActive = 1 AND MinSpent <= " & totalSpent & " ORDER BY SortOrder DESC")
    If Not rsTiers Is Nothing And Not rsTiers.EOF Then
        Set currentTier = Server.CreateObject("Scripting.Dictionary")
        currentTier.Add "tierCode", CStr(rsTiers("TierCode"))
        currentTier.Add "tierName", CStr(rsTiers("TierName") & "")
        currentTier.Add "tierNameEN", CStr(rsTiers("TierNameEN") & "")
        currentTier.Add "discountRate", CDbl(rsTiers("DiscountRate"))
        currentTier.Add "totalSpent", totalSpent
        currentTier.Add "iconClass", CStr(rsTiers("IconClass") & "")
        currentTier.Add "color", CStr(rsTiers("Color") & "")
        currentTier.Add "badgeBg", CStr(rsTiers("BadgeBg") & "")
        currentTier.Add "freeShipping", CBool(rsTiers("FreeShipping"))
        currentTier.Add "priorityShipping", CBool(rsTiers("PriorityShipping"))
        currentTier.Add "birthdayGift", CBool(rsTiers("BirthdayGift"))
        currentTier.Add "dedicatedSupport", CBool(rsTiers("DedicatedSupport"))
        rsTiers.Close
    Else
        ' 无匹配等级，返回默认
        Set currentTier = Server.CreateObject("Scripting.Dictionary")
        currentTier.Add "tierCode", "silver"
        currentTier.Add "tierName", "银卡会员"
        currentTier.Add "tierNameEN", "Silver"
        currentTier.Add "discountRate", 0.95
        currentTier.Add "totalSpent", totalSpent
        currentTier.Add "iconClass", "fa-medal"
        currentTier.Add "color", "#9E9E9E"
        currentTier.Add "badgeBg", "#f5f5f5"
        currentTier.Add "freeShipping", False
        currentTier.Add "priorityShipping", False
        currentTier.Add "birthdayGift", False
        currentTier.Add "dedicatedSupport", False
    End If
    Set rsTiers = Nothing
    
    On Error GoTo 0
    Set MU_V18_CalcUserTier = currentTier
End Function

' 获取下一等级信息（用于进度条）
' 返回: Dictionary {nextTierCode, nextTierName, needSpent, progress, iconClass, color, ...}
Function MU_V18_GetNextTierInfo(userId)
    If Not FEATURE_MEMBER_TIERS Then
        Set MU_V18_GetNextTierInfo = Nothing
        Exit Function
    End If
    
    Dim currentTier, tierCode, totalSpent, rs, nextInfo
    Set currentTier = MU_V18_CalcUserTier(userId)
    tierCode = currentTier("tierCode")
    totalSpent = currentTier("totalSpent")
    
    On Error Resume Next
    
    ' 查找下一个等级
    Dim sql
    sql = "SELECT TOP 1 * FROM MemberTiers WHERE IsActive = 1 AND MinSpent > (SELECT ISNULL(MinSpent, 0) FROM MemberTiers WHERE TierCode = '" & SafeSQL(tierCode) & "') ORDER BY SortOrder ASC"
    Set rs = conn.Execute(sql)
    
    If Not rs Is Nothing And Not rs.EOF Then
        Set nextInfo = Server.CreateObject("Scripting.Dictionary")
        Dim nextMinSpent, currentMinSpent
        nextMinSpent = CDbl(rs("MinSpent"))
        
        ' 获取当前等级门槛
        Dim rsCur
        Set rsCur = conn.Execute("SELECT ISNULL(MinSpent, 0) AS MinSpent FROM MemberTiers WHERE TierCode = '" & SafeSQL(tierCode) & "'")
        If Not rsCur Is Nothing And Not rsCur.EOF Then
            currentMinSpent = CDbl(rsCur("MinSpent"))
            rsCur.Close
        Else
            currentMinSpent = 0
        End If
        Set rsCur = Nothing
        
        Dim needSpent, progressPct
        needSpent = nextMinSpent - totalSpent
        If needSpent < 0 Then needSpent = 0
        
        ' 进度百分比
        Dim rangeTotal
        rangeTotal = nextMinSpent - currentMinSpent
        If rangeTotal > 0 Then
            progressPct = Round((totalSpent - currentMinSpent) / rangeTotal * 100, 1)
            If progressPct < 0 Then progressPct = 0
            If progressPct > 100 Then progressPct = 100
        Else
            progressPct = 100
        End If
        
        nextInfo.Add "nextTierCode", CStr(rs("TierCode"))
        nextInfo.Add "nextTierName", CStr(rs("TierName") & "")
        nextInfo.Add "nextTierNameEN", CStr(rs("TierNameEN") & "")
        nextInfo.Add "nextMinSpent", nextMinSpent
        nextInfo.Add "needSpent", needSpent
        nextInfo.Add "progress", progressPct
        nextInfo.Add "iconClass", CStr(rs("IconClass") & "")
        nextInfo.Add "color", CStr(rs("Color") & "")
        nextInfo.Add "currentMinSpent", currentMinSpent
        rs.Close
    Else
        ' 已是最高等级
        Set nextInfo = Server.CreateObject("Scripting.Dictionary")
        nextInfo.Add "isMax", True
        nextInfo.Add "progress", 100
    End If
    Set rs = Nothing
    
    On Error GoTo 0
    Set MU_V18_GetNextTierInfo = nextInfo
End Function

' 获取某等级的权益列表
' 返回: Recordset
Function MU_V18_GetTierBenefits(tierCode)
    If Not FEATURE_MEMBER_TIERS Then
        Set MU_V18_GetTierBenefits = Nothing
        Exit Function
    End If
    
    On Error Resume Next
    Dim sql
    sql = "SELECT * FROM MemberBenefits WHERE TierCode = '" & SafeSQL(tierCode) & "' AND IsActive = 1 ORDER BY SortOrder ASC"
    Set MU_V18_GetTierBenefits = conn.Execute(sql)
    On Error GoTo 0
End Function

' 获取所有等级列表（用于展示全部等级对比）
' 返回: Recordset
Function MU_V18_GetAllTiersWithBenefits()
    If Not FEATURE_MEMBER_TIERS Then
        Set MU_V18_GetAllTiersWithBenefits = Nothing
        Exit Function
    End If
    
    On Error Resume Next
    Dim sql
    sql = "SELECT t.*, (SELECT COUNT(*) FROM MemberBenefits b WHERE b.TierCode = t.TierCode AND b.IsActive = 1) AS BenefitCount " & _
          "FROM MemberTiers t WHERE t.IsActive = 1 ORDER BY t.SortOrder ASC"
    Set MU_V18_GetAllTiersWithBenefits = conn.Execute(sql)
    On Error GoTo 0
End Function

' 格式化金额为可读字符串
Function MU_V18_FormatSpent(amount)
    If amount >= 10000 Then
        MU_V18_FormatSpent = FormatNumber(amount / 10000, 1) & "万"
    Else
        MU_V18_FormatSpent = FormatNumber(amount, 0)
    End If
End Function

%>