<%
' ============================================
' V18 积分与奖励引擎 (Points & Rewards Engine)
' 功能: 积分获取/消费/兑换/过期/签到
' 兼容: 现有 UserPoints 表 + Users.Points 字段
' 依赖: dal.asp, config.asp
' ============================================

' ============================================
' 内部缓存：避免同一请求中重复查询规则表
' ============================================
Dim PE_RuleCache
Set PE_RuleCache = Nothing

Private Function PE_GetRuleCache()
    If IsObject(PE_RuleCache) And Not PE_RuleCache Is Nothing Then
        Set PE_GetRuleCache = PE_RuleCache
        Exit Function
    End If
    Set PE_RuleCache = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Dim rsRules, code, val
    Set rsRules = conn.Execute("SELECT RuleCode, RuleValue, RuleUnit FROM PointsRules WHERE IsEnabled = 1")
    If Not rsRules Is Nothing Then
        Do While Not rsRules.EOF
            code = LCase(rsRules("RuleCode").Value)
            val = CDbl(rsRules("RuleValue").Value)
            If Err.Number = 0 Then
                PE_RuleCache.Add code, val
            End If
            Err.Clear
            rsRules.MoveNext
        Loop
        rsRules.Close
    End If
    Set rsRules = Nothing
    On Error GoTo 0
    Set PE_GetRuleCache = PE_RuleCache
End Function

' ============================================
' 获取积分规则值
' ============================================
Function PE_GetRule(ruleCode)
    Dim cache, code, rawVal
    code = LCase(ruleCode)
    Set cache = PE_GetRuleCache()
    If cache.Exists(code) Then
        rawVal = cache(code)
        ' 安全类型转换：处理可能的 Decimal 子类型或空值
        If IsNumeric(rawVal) Then
            PE_GetRule = CDbl(rawVal)
        ElseIf IsNull(rawVal) Or IsEmpty(rawVal) Then
            ' 缓存值为空，走回退逻辑
        Else
            ' 非数字值，尝试转换
            On Error Resume Next
            PE_GetRule = CDbl(rawVal)
            If Err.Number <> 0 Then
                Err.Clear
                ' CDbl 失败，走回退
            Else
                On Error GoTo 0
                Exit Function
            End If
            On Error GoTo 0
        End If
        ' 如果 CDbl 成功（上面 Exit Function 会跳出），否则继续走回退
        If PE_GetRule <> 0 Then Exit Function
    End If
    ' 回退默认值
    Select Case code
        Case "purchase_rate":       PE_GetRule = 1
        Case "signin_points":       PE_GetRule = 5
        Case "review_points":       PE_GetRule = 20
        Case "review_with_photo":   PE_GetRule = 10
        Case "share_points":        PE_GetRule = 10
        Case "referral_points":     PE_GetRule = 100
        Case "referral_purchase":   PE_GetRule = 50
        Case "redeem_discount_rate": PE_GetRule = 100
        Case "max_redeem_pct":      PE_GetRule = 30
        Case "points_expire_months": PE_GetRule = 12
        Case Else:                  PE_GetRule = 0
    End Select
End Function

' ============================================
' 获取用户可用积分（汇总过期过滤）
' 优先从 PointsLedger 计算，回退到 UserPoints
' ============================================
Function PE_GetAvailablePoints(userId)
    Dim pts, expireMonths, rs, todayStr
    
    ' 先处理过期
    Call PE_ExpireOutdatedPoints(userId)
    
    pts = 0
    On Error Resume Next
    
    ' 从 PointsLedger 汇总有效积分
    Set rs = conn.Execute("SELECT ISNULL(SUM(Points), 0) FROM PointsLedger WHERE UserID = " & userId & " AND IsExpired = 0")
    If Not rs Is Nothing Then
        If Not rs.EOF Then pts = CLng(rs(0))
        rs.Close
    End If
    Set rs = Nothing
    
    ' 如果 PointsLedger 无数据，回退到 UserPoints
    If pts = 0 Then
        Set rs = conn.Execute("SELECT ISNULL(AvailablePoints, 0) FROM UserPoints WHERE UserID = " & userId)
        If Not rs Is Nothing Then
            If Not rs.EOF Then pts = CLng(rs(0))
            rs.Close
        End If
        Set rs = Nothing
        ' 再回退到 Users.Points
        If pts = 0 Then
            Set rs = conn.Execute("SELECT ISNULL(Points, 0) FROM Users WHERE UserID = " & userId)
            If Not rs Is Nothing Then
                If Not rs.EOF Then pts = CLng(rs(0))
                rs.Close
            End If
            Set rs = Nothing
        End If
    End If
    
    On Error GoTo 0
    PE_GetAvailablePoints = pts
End Function

' ============================================
' 处理过期积分（标记 IsExpired）
' ============================================
Sub PE_ExpireOutdatedPoints(userId)
    On Error Resume Next
    Dim expireMonths
    expireMonths = PE_GetRule("points_expire_months")
    If expireMonths <= 0 Then Exit Sub  ' 0=永不过期
    
    ' 标记过期积分（仅标记，不删除）
    conn.Execute "UPDATE PointsLedger SET IsExpired = 1 WHERE UserID = " & userId & _
               " AND IsExpired = 0 AND PointType = 'earn' AND ExpiresAt IS NOT NULL AND ExpiresAt < GETDATE()"
               
    ' 生成过期记录（汇总）
    Dim rsExpired
    Set rsExpired = conn.Execute("SELECT ISNULL(SUM(Points), 0) AS ExpiredSum, COUNT(*) AS ExpiredCount FROM PointsLedger WHERE UserID = " & userId & " AND IsExpired = 1 AND PointType = 'earn' AND Points > 0")
    On Error GoTo 0
End Sub

' ============================================
' 获取积分（消费/签到/评价/分享等）
' ============================================
Function PE_EarnPoints(userId, points, source, referenceId, description)
    PE_EarnPoints = False
    If points <= 0 Then Exit Function
    
    On Error Resume Next
    Dim expireMonths, expiresAt, expireSQL
    
    ' 计算过期时间
    expireMonths = PE_GetRule("points_expire_months")
    If expireMonths > 0 Then
        expiresAt = DateAdd("m", CInt(expireMonths), Now())
    Else
        expiresAt = Null
    End If
    
    ' 插入 PointsLedger
    Dim insertSQL
    insertSQL = "INSERT INTO PointsLedger (UserID, Points, PointType, Source, ReferenceID, Description, ExpiresAt, IsExpired) VALUES (" & _
              userId & ", " & points & ", 'earn', '" & SafeSQL(source) & "', " & IIf(IsNull(referenceId) Or referenceId = 0, "NULL", referenceId) & _
              ", '" & SafeSQL(description) & "', " & IIf(IsNull(expiresAt), "NULL", "'" & SafeSQL(FormatDateTime(expiresAt, 0)) & "'") & ", 0)"
    
    conn.Execute insertSQL
    
    If Err.Number = 0 Then
        ' 同步更新 UserPoints（保持兼容）
        Dim checkSQL
        checkSQL = "SELECT COUNT(*) FROM UserPoints WHERE UserID = " & userId
        Dim rsChk
        Set rsChk = conn.Execute(checkSQL)
        Dim hasRecord : hasRecord = False
        If Not rsChk Is Nothing Then
            If Not rsChk.EOF Then hasRecord = (CLng(rsChk(0)) > 0)
            rsChk.Close
        End If
        Set rsChk = Nothing
        
        If hasRecord Then
            conn.Execute "UPDATE UserPoints SET AvailablePoints = ISNULL(AvailablePoints, 0) + " & points & _
                       ", TotalPoints = ISNULL(TotalPoints, 0) + " & points & _
                       ", LastUpdatedAt = GETDATE() WHERE UserID = " & userId
        Else
            conn.Execute "INSERT INTO UserPoints (UserID, AvailablePoints, TotalPoints, LastUpdatedAt) VALUES (" & userId & ", " & points & ", " & points & ", GETDATE())"
        End If
        
        ' 同步 Users.Points
        conn.Execute "UPDATE Users SET Points = ISNULL(Points, 0) + " & points & " WHERE UserID = " & userId
        
        PE_EarnPoints = True
    End If
    
    On Error GoTo 0
End Function

' ============================================
' 消费积分（兑换/抵扣）
' ============================================
Function PE_RedeemPoints(userId, points, redemptionType, referenceId)
    PE_RedeemPoints = False
    If points <= 0 Then Exit Function
    
    Dim available
    available = PE_GetAvailablePoints(userId)
    If points > available Then
        points = available
        If points <= 0 Then Exit Function
    End If
    
    On Error Resume Next
    Dim description
    description = redemptionType & " redemption (" & points & " pts)"
    
    ' 插入消耗记录（负值）
    Dim insertSQL
    insertSQL = "INSERT INTO PointsLedger (UserID, Points, PointType, Source, ReferenceID, Description, ExpiresAt, IsExpired) VALUES (" & _
              userId & ", -" & points & ", 'redeem', '" & SafeSQL(redemptionType) & "', " & _
              IIf(IsNull(referenceId) Or referenceId = 0, "NULL", referenceId) & _
              ", '" & SafeSQL(description) & "', NULL, 0)"
    
    conn.Execute insertSQL
    
    If Err.Number = 0 Then
        ' 同步 UserPoints
        conn.Execute "UPDATE UserPoints SET AvailablePoints = ISNULL(AvailablePoints, 0) - " & points & _
                   ", UsedPoints = ISNULL(UsedPoints, 0) + " & points & _
                   ", LastUpdatedAt = GETDATE() WHERE UserID = " & userId
        
        ' 同步 Users.Points
        conn.Execute "UPDATE Users SET Points = ISNULL(Points, 0) - " & points & " WHERE UserID = " & userId
        
        PE_RedeemPoints = True
    End If
    
    On Error GoTo 0
End Function

' ============================================
' 计算积分的货币价值（元）
' ============================================
Function PE_CalcPointsValue(points)
    Dim rate
    rate = PE_GetRule("redeem_discount_rate")
    If rate <= 0 Then rate = 100
    PE_CalcPointsValue = CDbl(points) / CDbl(rate)
End Function

' ============================================
' 计算消费应得积分
' ============================================
Function PE_CalcOrderPoints(orderAmount)
    Dim rate
    rate = PE_GetRule("purchase_rate")
    If rate <= 0 Then rate = 1
    PE_CalcOrderPoints = CInt(CDbl(orderAmount) * rate)
End Function

' ============================================
' 获取最大抵扣百分比
' ============================================
Function PE_GetMaxRedeemPct()
    PE_GetMaxRedeemPct = PE_GetRule("max_redeem_pct")
End Function

' ============================================
' 计算订单可抵扣的最大积分
' ============================================
Function PE_CalcMaxRedeemablePoints(userId, orderAmount)
    Dim maxPct, maxValue, userPoints, rate
    rate = PE_GetRule("redeem_discount_rate")
    If rate <= 0 Then rate = 100
    maxPct = PE_GetRule("max_redeem_pct") / 100
    If maxPct <= 0 Then maxPct = 0.3
    maxValue = orderAmount * maxPct  ' 最大可抵扣金额
    PE_CalcMaxRedeemablePoints = Int(maxValue * rate)  ' 转换为积分数
End Function

' ============================================
' 签到检查
' ============================================
Function PE_CheckSignIn(userId)
    PE_CheckSignIn = False
    On Error Resume Next
    Dim rs, todayStr
    todayStr = FormatDateTime(Date(), 2)  ' yyyy-mm-dd
    Set rs = conn.Execute("SELECT COUNT(*) FROM PointsLedger WHERE UserID = " & userId & _
                        " AND Source = 'signin' AND CONVERT(date, CreatedAt) = CONVERT(date, GETDATE())")
    If Not rs Is Nothing Then
        If Not rs.EOF Then PE_CheckSignIn = (CLng(rs(0)) > 0)
        rs.Close
    End If
    Set rs = Nothing
    On Error GoTo 0
End Function

' ============================================
' 执行签到
' ============================================
Function PE_DoSignIn(userId)
    PE_DoSignIn = 0
    If PE_CheckSignIn(userId) Then
        PE_DoSignIn = -1  ' 已签到
        Exit Function
    End If
    
    Dim points, success
    points = CInt(PE_GetRule("signin_points"))
    If points <= 0 Then points = 5
    
    success = PE_EarnPoints(userId, points, "signin", 0, "Daily sign-in")
    If success Then
        PE_DoSignIn = points
    End If
End Function

' ============================================
' 获取积分账本（分页）
' ============================================
Function PE_GetPointsLedger(userId, pageNum, pageSize)
    If pageNum < 1 Then pageNum = 1
    If pageSize < 1 Then pageSize = 20
    Dim offset
    offset = (pageNum - 1) * pageSize
    
    Dim sql
    sql = "SELECT * FROM PointsLedger WHERE UserID = " & userId & " ORDER BY CreatedAt DESC " & _
          "OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY"
    
    Set PE_GetPointsLedger = conn.Execute(sql)
End Function

' ============================================
' 获取积分账本总数
' ============================================
Function PE_GetPointsLedgerCount(userId)
    Dim rs
    PE_GetPointsLedgerCount = 0
    On Error Resume Next
    Set rs = conn.Execute("SELECT COUNT(*) FROM PointsLedger WHERE UserID = " & userId)
    If Not rs Is Nothing Then
        If Not rs.EOF Then PE_GetPointsLedgerCount = CLng(rs(0))
        rs.Close
    End If
    Set rs = Nothing
    On Error GoTo 0
End Function

' ============================================
' 获取积分汇总（仪表盘用）
' ============================================
Function PE_GetPointsSummary(userId)
    Dim result
    Set result = Server.CreateObject("Scripting.Dictionary")
    
    Dim available, totalEarned, totalRedeemed, todayEarned, expiringSoon
    
    available = PE_GetAvailablePoints(userId)
    
    On Error Resume Next
    Dim rs
    
    ' 累计获得
    Set rs = conn.Execute("SELECT ISNULL(SUM(Points), 0) FROM PointsLedger WHERE UserID = " & userId & " AND PointType = 'earn' AND IsExpired = 0")
    totalEarned = 0
    If Not rs Is Nothing Then
        If Not rs.EOF Then totalEarned = CLng(rs(0))
        rs.Close
    End If
    Set rs = Nothing
    
    ' 累计消费
    Set rs = conn.Execute("SELECT ISNULL(SUM(ABS(Points)), 0) FROM PointsLedger WHERE UserID = " & userId & " AND PointType = 'redeem'")
    totalRedeemed = 0
    If Not rs Is Nothing Then
        If Not rs.EOF Then totalRedeemed = CLng(rs(0))
        rs.Close
    End If
    Set rs = Nothing
    
    ' 今日获得
    Set rs = conn.Execute("SELECT ISNULL(SUM(Points), 0) FROM PointsLedger WHERE UserID = " & userId & " AND PointType = 'earn' AND CONVERT(date, CreatedAt) = CONVERT(date, GETDATE())")
    todayEarned = 0
    If Not rs Is Nothing Then
        If Not rs.EOF Then todayEarned = CLng(rs(0))
        rs.Close
    End If
    Set rs = Nothing
    
    ' 即将过期（30天内）
    Dim expireMonths
    expireMonths = PE_GetRule("points_expire_months")
    expiringSoon = 0
    If expireMonths > 0 Then
        Set rs = conn.Execute("SELECT ISNULL(SUM(Points), 0) FROM PointsLedger WHERE UserID = " & userId & _
                            " AND PointType = 'earn' AND IsExpired = 0 AND ExpiresAt IS NOT NULL AND ExpiresAt <= DATEADD(day, 30, GETDATE())")
        If Not rs Is Nothing Then
            If Not rs.EOF Then expiringSoon = CLng(rs(0))
            rs.Close
        End If
        Set rs = Nothing
    End If
    
    On Error GoTo 0
    
    result.Add "available", available
    result.Add "totalEarned", totalEarned
    result.Add "totalRedeemed", totalRedeemed
    result.Add "todayEarned", todayEarned
    result.Add "expiringSoon", expiringSoon
    
    Set PE_GetPointsSummary = result
End Function

' ============================================
' 获取订单中已获/已用的积分
' ============================================
Function PE_GetOrderPoints(orderId)
    Dim result
    Set result = Server.CreateObject("Scripting.Dictionary")
    result.Add "earned", 0
    result.Add "redeemed", 0
    
    On Error Resume Next
    Dim rs
    ' 从 Orders 表读取
    Set rs = conn.Execute("SELECT ISNULL(PointsEarned, 0), ISNULL(PointsRedeemed, 0), ISNULL(PointsDiscount, 0) FROM Orders WHERE OrderID = " & orderId)
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            result("earned") = CLng(rs(0))
            result("redeemed") = CLng(rs(1))
            If Not IsNull(rs(2)) Then result.Add "discount", CDbl(rs(2))
        End If
        rs.Close
    End If
    Set rs = Nothing
    
    ' 如果 Orders 无数据，从 PointsLedger 查
    If result("earned") = 0 Then
        Set rs = conn.Execute("SELECT ISNULL(SUM(Points), 0) FROM PointsLedger WHERE UserID > 0 AND ReferenceID = " & orderId & " AND Source = 'purchase'")
        If Not rs Is Nothing Then
            If Not rs.EOF Then result("earned") = CLng(rs(0))
            rs.Close
        End If
        Set rs = Nothing
    End If
    On Error GoTo 0
    
    Set PE_GetOrderPoints = result
End Function

' ============================================
' 获取兑换商品列表
' ============================================
Function PE_GetRedemptionItems()
    Dim rs
    On Error Resume Next
    Set rs = conn.Execute("SELECT * FROM PointsRedemption WHERE IsEnabled = 1 AND Stock > 0 ORDER BY SortOrder, PointsCost")
    Set PE_GetRedemptionItems = rs
    On Error GoTo 0
End Function

' ============================================
' 获取单个兑换商品
' ============================================
Function PE_GetRedemptionItem(redemptionId)
    Dim rs
    On Error Resume Next
    Set rs = conn.Execute("SELECT * FROM PointsRedemption WHERE RedemptionID = " & redemptionId)
    If Not rs Is Nothing Then
        If rs.EOF Then
            rs.Close
            Set rs = Nothing
        End If
    End If
    Set PE_GetRedemptionItem = rs
    On Error GoTo 0
End Function

' ============================================
' 兑换处理
' ============================================
Function PE_DoRedeem(userId, redemptionId)
    PE_DoRedeem = ""
    
    Dim rsItem, itemName, itemType, pointsCost, stock, success
    
    Set rsItem = PE_GetRedemptionItem(redemptionId)
    If rsItem Is Nothing Then
        PE_DoRedeem = "兑换项目不存在或已售罄"
        Exit Function
    End If
    
    itemName = rsItem("ItemName")
    itemType = rsItem("ItemType")
    pointsCost = CLng(rsItem("PointsCost"))
    stock = CLng(rsItem("Stock"))
    rsItem.Close
    Set rsItem = Nothing
    
    If stock <= 0 Then
        PE_DoRedeem = "该兑换项目已售罄"
        Exit Function
    End If
    
    ' 检查积分是否足够
    Dim available
    available = PE_GetAvailablePoints(userId)
    If available < pointsCost Then
        PE_DoRedeem = "积分不足（需要" & pointsCost & "积分，当前" & available & "积分）"
        Exit Function
    End If
    
    ' 执行积分扣除
    success = PE_RedeemPoints(userId, pointsCost, itemType, 0)
    If Not success Then
        PE_DoRedeem = "积分扣除失败，请稍后再试"
        Exit Function
    End If
    
    ' 扣减库存
    On Error Resume Next
    conn.Execute "UPDATE PointsRedemption SET Stock = Stock - 1 WHERE RedemptionID = " & redemptionId & " AND Stock > 0"
    
    ' 记录兑换（在 PointsLedger 描述中已记录，这里可扩展优惠券生成等逻辑）
    If itemType = "coupon" Then
        ' 可在此生成优惠券记录
    End If
    
    On Error GoTo 0
    PE_DoRedeem = ""  ' 成功返回空字符串
End Function

' ============================================
' 格式化积分显示
' ============================================
Function PE_FormatPoints(pts)
    If pts >= 10000 Then
        PE_FormatPoints = FormatNumber(pts / 10000, 1) & "万"
    Else
        PE_FormatPoints = FormatNumber(pts, 0)
    End If
End Function

' ============================================
' 获取积分获取渠道名称
' ============================================
Function PE_GetSourceName(source)
    Select Case LCase(source)
        Case "purchase":       PE_GetSourceName = "消费得积分"
        Case "signin":         PE_GetSourceName = "签到"
        Case "review":         PE_GetSourceName = "评价"
        Case "review_photo":   PE_GetSourceName = "带图评价"
        Case "share":          PE_GetSourceName = "分享"
        Case "referral":       PE_GetSourceName = "推荐好友"
        Case "redeem_discount": PE_GetSourceName = "积分抵扣"
        Case "redeem_coupon":  PE_GetSourceName = "兑换优惠券"
        Case "redeem_sample":  PE_GetSourceName = "兑换小样"
        Case "redeem_bottle":  PE_GetSourceName = "兑换瓶身"
        Case "expire":         PE_GetSourceName = "积分过期"
        Case "adjust":         PE_GetSourceName = "积分调整"
        Case "manual":         PE_GetSourceName = "管理员操作"
        Case Else:             PE_GetSourceName = source
    End Select
End Function
%>
