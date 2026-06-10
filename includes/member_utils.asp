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
    Dim level, levelName, discount
    level = MU_CalcUserLevel(userId)
    levelName = MU_GetLevelName(level)
    discount = MU_GetLevelDiscount(level)
    
    Dim badgeClass, icon
    Select Case level
        Case "L0": badgeClass = "level-0": icon = "fa-user"
        Case "L1": badgeClass = "level-1": icon = "fa-star"
        Case "L2": badgeClass = "level-2": icon = "fa-gem"
        Case "L3": badgeClass = "level-3": icon = "fa-crown"
        Case "L4": badgeClass = "level-4": icon = "fa-crown"
    End Select
%>
<span class="level-badge <%= badgeClass %>">
    <i class="fas <%= icon %>"></i> <%= levelName %>
    <% If discount < 1 Then %><small>(<%= FormatPercent(1 - discount, 0) %> OFF)</small><% End If %>
</span>
<style>
.level-badge { display:inline-flex; align-items:center; gap:5px; padding:4px 12px; border-radius:12px; font-size:13px; font-weight:600; }
.level-0 { background:#f0f0f0; color:#666; }
.level-1 { background:#e8f4f8; color:#2196F3; }
.level-2 { background:#fff3e0; color:#FF9800; }
.level-3 { background:#fce4ec; color:#e91e63; }
.level-4 { background:#e8f5e9; color:#4CAF50; }
.level-badge small { opacity:0.7; font-weight:400; }
</style>
<%
End Sub
%>