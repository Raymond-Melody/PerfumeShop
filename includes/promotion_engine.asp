<%
' ============================================
' 促销活动引擎 - Promotion Engine
' 管理优惠券、满减、限时折扣等促销活动
' ============================================

' ============================================
' 获取当前有效的促销活动
' ============================================
Function PE_GetActivePromotions()
    Dim rs, sql
    
    On Error Resume Next
    
    sql = "SELECT * FROM SiteSettings WHERE SettingKey LIKE 'Promotion_%' AND SettingValue LIKE '%Active=1%' AND SettingValue LIKE '%EndDate%'"
    Set rs = conn.Execute(sql)
    Set PE_GetActivePromotions = rs
End Function

' ============================================
' 检查购物车是否满足促销条件
' ============================================
Function PE_CheckPromotionEligibility(cartTotal, userId)
    Dim result, rs, key, val, promoType, minAmount, discount, promoName, canUse
    Set result = Server.CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    
    ' 检查满减活动
    Set rs = conn.Execute("SELECT SettingKey, SettingValue FROM SiteSettings WHERE SettingKey='Promotion_Threshold'")
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            val = rs("SettingValue")
            If InStr(val, "|") > 0 Then
                Dim parts
                parts = Split(val, "|")
                minAmount = CDbl(parts(0))
                discount = CDbl(parts(1))
                promoName = parts(2)
                
                If cartTotal >= minAmount Then
                    result.Add "threshold", "{""name"":""" & promoName & """,""minAmount"":" & minAmount & ",""discount"":" & discount & "}"
                End If
            End If
        End If
        rs.Close
    End If
    Set rs = Nothing
    
    Set PE_CheckPromotionEligibility = result
End Function

' ============================================
' 计算促销折扣金额
' ============================================
Function PE_CalculateDiscount(cartTotal, userId)
    Dim discount, thresholdVal, percentVal, shippingFree, promoThreshold, promoPercent
    discount = 0
    
    On Error Resume Next
    
    ' 满减折扣
    Dim rs
    Set rs = conn.Execute("SELECT SettingValue FROM SiteSettings WHERE SettingKey='Promotion_Threshold'")
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            thresholdVal = rs(0)
            If InStr(thresholdVal, "|") > 0 Then
                Dim parts
                parts = Split(thresholdVal, "|")
                promoThreshold = CDbl(parts(0))
                promoPercent = CDbl(parts(1))
                If cartTotal >= promoThreshold Then
                    discount = discount + promoPercent
                End If
            End If
        End If
        rs.Close
    End If
    Set rs = Nothing
    
    ' 会员等级折扣（V18: 自动适配新旧等级系统）
    Dim level, levelDiscount
    levelDiscount = (1 - MU_GetEffectiveDiscountRate(userId)) * cartTotal
    discount = discount + levelDiscount
    
    ' 首单折扣
    Dim orderCount
    Set rs = conn.Execute("SELECT COUNT(*) FROM Orders WHERE UserID = " & userId)
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            orderCount = CLng(rs(0))
            If orderCount <= 1 Then
                Set rs = conn.Execute("SELECT SettingValue FROM SiteSettings WHERE SettingKey='Promotion_FirstOrder'")
                If Not rs Is Nothing Then
                    If Not rs.EOF Then
                        Dim firstDiscount
                        firstDiscount = CDbl(rs(0))
                        If firstDiscount > 0 Then
                            discount = discount + (cartTotal * firstDiscount / 100)
                        End If
                    End If
                    rs.Close
                End If
                Set rs = Nothing
            End If
        End If
        rs.Close
    End If
    Set rs = Nothing
    
    If discount > cartTotal * 0.5 Then discount = cartTotal * 0.5 ' 最高50%折扣
    
    PE_CalculateDiscount = discount
End Function

' ============================================
' 运费减免检查
' ============================================
Function PE_CheckFreeShipping(cartTotal)
    PE_CheckFreeShipping = (cartTotal >= FREE_SHIPPING_AMOUNT)
End Function

' ============================================
' 渲染促销信息横幅
' ============================================
Sub PE_RenderPromotionBanner(cartTotal)
    Dim rs, val, minAmount, discount, promoName
    
    On Error Resume Next
    
    Set rs = conn.Execute("SELECT SettingValue FROM SiteSettings WHERE SettingKey='Promotion_Threshold'")
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            val = rs(0)
            If InStr(val, "|") > 0 Then
                Dim parts
                parts = Split(val, "|")
                minAmount = CDbl(parts(0))
                discount = CDbl(parts(1))
                promoName = parts(2)
                
                If cartTotal >= minAmount Then
                    Response.Write "<div class=""promo-banner promo-success"">"
                    Response.Write "<i class=""fas fa-gift""></i> 恭喜！已享受 <strong>" & promoName & "</strong>，优惠 ¥" & FormatNumber(discount, 2)
                    Response.Write "</div>"
                Else
                    Dim remaining
                    remaining = minAmount - cartTotal
                    If remaining > 0 Then
                        Response.Write "<div class=""promo-banner promo-hint"">"
                        Response.Write "<i class=""fas fa-info-circle""></i> 再消费 <strong>¥" & FormatNumber(remaining, 2) & "</strong> 即可享受 " & promoName
                        Response.Write "</div>"
                    End If
                End If
            End If
        End If
        rs.Close
    End If
    Set rs = Nothing
    
    ' 免费送货提示
    If cartTotal < FREE_SHIPPING_AMOUNT Then
        Dim shipRemain
        shipRemain = FREE_SHIPPING_AMOUNT - cartTotal
        Response.Write "<div class=""promo-banner promo-shipping"">"
        Response.Write "<i class=""fas fa-truck""></i> 再消费 <strong>¥" & FormatNumber(shipRemain, 2) & "</strong> 即可享受免费配送"
        Response.Write "</div>"
    Else
        Response.Write "<div class=""promo-banner promo-success"">"
        Response.Write "<i class=""fas fa-truck""></i> 已享受免费配送"
        Response.Write "</div>"
    End If
End Sub

' ============================================
' 促销活动管理界面（后台）
' ============================================
Function PE_RenderAdminPanel()
%>
<div class="promo-admin-panel">
    <h3><i class="fas fa-bullhorn"></i> 促销活动管理</h3>
    
    <form method="post" action="../operation/marketing.asp" class="promo-form">
        <div class="form-group">
            <label>满减活动配置</label>
            <div class="input-row">
                <input type="number" name="promoThreshold" placeholder="满额门槛 (元)" step="0.01">
                <input type="number" name="promoDiscount" placeholder="减金额 (元)" step="0.01">
                <input type="text" name="promoName" placeholder="活动名称 (如：满299减50)">
                <button type="submit" class="btn btn-primary">设置</button>
            </div>
        </div>
        
        <div class="form-group">
            <label>首单折扣 (%)</label>
            <div class="input-row">
                <input type="number" name="firstOrderDiscount" placeholder="折扣百分比" min="0" max="50" step="1">
                <button type="submit" class="btn btn-primary">设置</button>
            </div>
        </div>
    </form>
</div>
<style>
.promo-admin-panel { background:#2d2d44; border-radius:12px; padding:25px; margin:20px 0; }
.promo-admin-panel h3 { color:#e0e0e0; margin-bottom:20px; }
.promo-form .form-group { margin-bottom:20px; }
.promo-form label { display:block; color:#b0b0b0; margin-bottom:8px; font-size:14px; }
.promo-form .input-row { display:flex; gap:10px; flex-wrap:wrap; }
.promo-form input { padding:10px 15px; border:2px solid #3a3a4a; border-radius:8px; background:#1a1a2e; color:#e0e0e0; font-size:14px; }
.promo-form input:focus { border-color:#00bcd4; outline:none; }
.promo-banner { padding:12px 18px; border-radius:8px; margin-bottom:12px; display:flex; align-items:center; gap:10px; font-size:14px; }
.promo-success { background:rgba(76,175,80,0.15); color:#81c784; border:1px solid rgba(76,175,80,0.3); }
.promo-hint { background:rgba(33,150,243,0.15); color:#64b5f6; border:1px solid rgba(33,150,243,0.3); }
.promo-shipping { background:rgba(255,152,0,0.15); color:#ffb74d; border:1px solid rgba(255,152,0,0.3); }
</style>
<%
End Function
%>

<%
' ============================================
' V18 优惠券引擎 (Coupon Engine)
' 迁移 SiteSettings → Coupons/UserCoupons 专用表
' 支持: 满减券/折扣券/免邮券/礼品券
' 获取渠道: 新注册/活动领取/积分兑换/会员升级
' ============================================

' ============================================
' 内部辅助: 获取单个优惠券定义
' ============================================
Private Function PE_CouponGetByCodeRaw(code)
    Dim rs, sql
    On Error Resume Next
    sql = "SELECT * FROM Coupons WHERE CouponCode = '" & SafeSQL(code) & "' AND IsActive = 1"
    Set rs = conn.Execute(sql)
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            Set PE_CouponGetByCodeRaw = rs
            Exit Function
        End If
        rs.Close
    End If
    Set rs = Nothing
    Set PE_CouponGetByCodeRaw = Nothing
    On Error GoTo 0
End Function

' ============================================
' 获取优惠券定义（返回 Dictionary）
' ============================================
Function PE_CouponGetDict(rsCoupon)
    Dim d
    Set d = Server.CreateObject("Scripting.Dictionary")
    If rsCoupon Is Nothing Then
        Set PE_CouponGetDict = d
        Exit Function
    End If
    d.Add "id", rsCoupon("CouponID")
    d.Add "code", rsCoupon("CouponCode")
    d.Add "name", rsCoupon("CouponName")
    d.Add "type", rsCoupon("CouponType")
    d.Add "value", CDbl(rsCoupon("DiscountValue"))
    d.Add "minSpend", CDbl(rsCoupon("MinSpend"))
    d.Add "maxDiscount", CDbl(rsCoupon("MaxDiscount"))
    d.Add "firstOrder", CBool(rsCoupon("FirstOrderOnly"))
    d.Add "category", rsCoupon("ApplicableCategory") & ""
    d.Add "productId", rsCoupon("ApplicableProductID") & ""
    d.Add "desc", rsCoupon("Description") & ""
    d.Add "terms", rsCoupon("Terms") & ""
    d.Add "isPublic", CBool(rsCoupon("IsPublic"))
    If Not IsNull(rsCoupon("TotalQty")) Then d.Add "totalQty", CLng(rsCoupon("TotalQty"))
    If Not IsNull(rsCoupon("UsedQty")) Then d.Add "usedQty", CLng(rsCoupon("UsedQty"))
    If Not IsNull(rsCoupon("ValidFrom")) Then d.Add "validFrom", rsCoupon("ValidFrom")
    If Not IsNull(rsCoupon("ValidTo")) Then d.Add "validTo", rsCoupon("ValidTo")
    Set PE_CouponGetDict = d
End Function

' ============================================
' 验证优惠券是否可用于当前购物车
' 返回: Dictionary { valid: bool, message: string, discount: double, type: string }
' ============================================
Function PE_CouponValidate(code, userId, cartTotal)
    Dim result
    Set result = Server.CreateObject("Scripting.Dictionary")
    result.Add "valid", False
    result.Add "message", ""
    result.Add "discount", 0
    result.Add "type", ""
    
    If code = "" Then
        result("message") = "请输入优惠码"
        Set PE_CouponValidate = result
        Exit Function
    End If
    
    Dim rs
    Set rs = PE_CouponGetByCodeRaw(code)
    If rs Is Nothing Then
        result("message") = "优惠码不存在或已失效"
        Set PE_CouponValidate = result
        Exit Function
    End If
    
    Dim couponType, discountValue, minSpend, maxDiscount, totalQty, usedQty, firstOrder, validFrom, validTo, couponId
    couponId = rs("CouponID")
    couponType = rs("CouponType")
    discountValue = CDbl(rs("DiscountValue"))
    minSpend = CDbl(rs("MinSpend"))
    maxDiscount = CDbl(rs("MaxDiscount"))
    totalQty = CLng(rs("TotalQty"))
    usedQty = CLng(rs("UsedQty"))
    firstOrder = CBool(rs("FirstOrderOnly"))
    validFrom = rs("ValidFrom")
    validTo = rs("ValidTo")
    rs.Close
    Set rs = Nothing
    
    ' 有效期检查
    If Now() < validFrom Or Now() > validTo Then
        result("message") = "优惠码不在有效期内"
        Set PE_CouponValidate = result
        Exit Function
    End If
    
    ' 库存检查
    If totalQty > 0 And usedQty >= totalQty Then
        result("message") = "优惠码已被领完"
        Set PE_CouponValidate = result
        Exit Function
    End If
    
    ' 最低消费检查
    If minSpend > 0 And cartTotal < minSpend Then
        result("message") = "未达最低消费 ¥" & FormatNumber(minSpend, 0) & "（当前 ¥" & FormatNumber(cartTotal, 2) & "）"
        Set PE_CouponValidate = result
        Exit Function
    End If
    
    ' 首单检查
    If firstOrder Then
        Dim orderCount
        orderCount = 0
        On Error Resume Next
        Dim rsOC
        Set rsOC = conn.Execute("SELECT COUNT(*) FROM Orders WHERE UserID = " & userId)
        If Not rsOC Is Nothing Then
            If Not rsOC.EOF Then orderCount = CLng(rsOC(0))
            rsOC.Close
        End If
        Set rsOC = Nothing
        On Error GoTo 0
        If orderCount > 0 Then
            result("message") = "此优惠券仅限首单使用"
            Set PE_CouponValidate = result
            Exit Function
        End If
    End If
    
    ' 检查用户是否已使用过此券
    Dim rsUsed
    On Error Resume Next
    Set rsUsed = conn.Execute("SELECT COUNT(*) FROM UserCoupons WHERE UserID = " & userId & " AND CouponCode = '" & SafeSQL(code) & "' AND Status = 'used'")
    If Err.Number <> 0 Then
        Err.Clear
        Set rsUsed = Nothing
    End If
    If Not rsUsed Is Nothing Then
        If Not rsUsed.EOF Then
            If CLng(rsUsed(0)) > 0 Then
                result("message") = "该优惠码您已使用过"
                rsUsed.Close
                Set rsUsed = Nothing
                Set PE_CouponValidate = result
                Exit Function
            End If
        End If
        rsUsed.Close
    End If
    Set rsUsed = Nothing
    On Error GoTo 0
    
    ' 计算折扣金额
    Dim discount
    discount = PE_CouponCalcDiscount(couponType, discountValue, cartTotal, maxDiscount)
    
    result("valid") = True
    result("message") = "优惠码验证通过"
    result("discount") = discount
    result("type") = couponType
    Set PE_CouponValidate = result
End Function

' ============================================
' 计算优惠券折扣金额
' ============================================
Function PE_CouponCalcDiscount(couponType, discountValue, cartTotal, maxDiscount)
    Dim discount
    discount = 0
    Select Case LCase(couponType)
        Case "fixed"
            discount = discountValue
        Case "percentage"
            discount = cartTotal * (discountValue / 100)
            If maxDiscount > 0 And discount > maxDiscount Then
                discount = maxDiscount
            End If
        Case "free_shipping"
            discount = 0  ' 免邮券不产生金额折扣，在结算时另外处理
        Case "gift"
            discount = 0  ' 礼品券不直接折现
    End Select
    If discount > cartTotal Then discount = cartTotal
    PE_CouponCalcDiscount = discount
End Function

' ============================================
' 使用优惠券（在订单创建成功后调用）
' ============================================
Function PE_CouponUse(code, userId, orderId)
    PE_CouponUse = False
    If code = "" Then Exit Function
    
    On Error Resume Next
    
    ' 标记用户券为已使用
    conn.Execute "UPDATE UserCoupons SET Status = 'used', UsedAt = GETDATE(), UsedOrderID = " & orderId & _
               " WHERE UserID = " & userId & " AND CouponCode = '" & SafeSQL(code) & "' AND Status = 'available'"
    
    ' 增加券的已使用计数
    conn.Execute "UPDATE Coupons SET UsedQty = UsedQty + 1 WHERE CouponCode = '" & SafeSQL(code) & "'"
    
    ' 更新 Orders 表
    conn.Execute "UPDATE Orders SET CouponCode = '" & SafeSQL(code) & "' WHERE OrderID = " & orderId
    
    If Err.Number = 0 Then PE_CouponUse = True
    On Error GoTo 0
End Function

' ============================================
' 给用户发放优惠券
' ============================================
Function PE_CouponIssue(userId, code, source)
    PE_CouponIssue = False
    If code = "" Or userId = "" Then Exit Function
    
    On Error Resume Next
    
    ' 检查券是否存在且有效
    Dim rs, couponId, validTo
    Set rs = PE_CouponGetByCodeRaw(code)
    If rs Is Nothing Then Exit Function
    
    couponId = rs("CouponID")
    validTo = rs("ValidTo")
    Dim totalQty, usedQty
    totalQty = CLng(rs("TotalQty"))
    usedQty = CLng(rs("UsedQty"))
    rs.Close
    Set rs = Nothing
    
    ' 库存检查
    If totalQty > 0 And usedQty >= totalQty Then Exit Function
    
    ' 检查用户是否已有此券且未使用
    Dim rsHas
    Set rsHas = conn.Execute("SELECT COUNT(*) FROM UserCoupons WHERE UserID = " & userId & " AND CouponCode = '" & SafeSQL(code) & "' AND Status = 'available'")
    If Not rsHas Is Nothing Then
        If Not rsHas.EOF Then
            If CLng(rsHas(0)) > 0 Then
                rsHas.Close
                Set rsHas = Nothing
                Exit Function  ' 已有可用券，不重复发放
            End If
        End If
        rsHas.Close
    End If
    Set rsHas = Nothing
    
    ' 发放
    conn.Execute "INSERT INTO UserCoupons (UserID, CouponID, CouponCode, Source, Status, ObtainedAt, ExpiresAt) VALUES (" & _
               userId & ", " & couponId & ", '" & SafeSQL(code) & "', '" & SafeSQL(source) & "', 'available', GETDATE(), '" & SafeSQL(FormatDateTime(validTo, 0)) & "')"
    
    ' 更新库存
    conn.Execute "UPDATE Coupons SET UsedQty = UsedQty + 1 WHERE CouponID = " & couponId
    
    If Err.Number = 0 Then PE_CouponIssue = True
    On Error GoTo 0
End Function

' ============================================
' 新人礼包发放
' ============================================
Function PE_CouponIssueWelcome(userId)
    Dim success, count
    success = False
    count = 0
    
    ' 发放新人专属券
    If PE_CouponIssue(userId, "WELCOME10", "new_user") Then count = count + 1
    If PE_CouponIssue(userId, "WELCOME20", "new_user") Then count = count + 1
    
    ' 检查是否有公开可领的免邮券
    Dim rs
    On Error Resume Next
    Set rs = conn.Execute("SELECT CouponCode FROM Coupons WHERE CouponType = 'free_shipping' AND IsActive = 1 AND IsPublic = 1 AND GETDATE() BETWEEN ValidFrom AND ValidTo")
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            If PE_CouponIssue(userId, rs("CouponCode"), "new_user") Then count = count + 1
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
    On Error GoTo 0
    
    PE_CouponIssueWelcome = count
End Function

' ============================================
' 会员升级礼券
' ============================================
Function PE_CouponIssueTierUpgrade(userId, tierCode)
    PE_CouponIssueTierUpgrade = False
    Select Case LCase(tierCode)
        Case "gold":
            PE_CouponIssueTierUpgrade = PE_CouponIssue(userId, "TIER_GOLD", "tier_upgrade")
        Case "diamond", "black":
            Dim result : result = False
            If PE_CouponIssue(userId, "VIP5", "tier_upgrade") Then result = True
            If PE_CouponIssue(userId, "FREESHIP", "tier_upgrade") Then result = True
            PE_CouponIssueTierUpgrade = result
    End Select
End Function

' ============================================
' 获取用户优惠券列表
' ============================================
Function PE_CouponGetUserCoupons(userId, statusFilter)
    If statusFilter = "" Then statusFilter = "available"
    On Error Resume Next
    Dim sql
    sql = "SELECT uc.*, c.CouponName, c.CouponType, c.DiscountValue, c.MinSpend, c.MaxDiscount, " & _
          "c.Description, c.Terms, c.ValidFrom, c.ValidTo " & _
          "FROM UserCoupons uc LEFT JOIN Coupons c ON uc.CouponID = c.CouponID " & _
          "WHERE uc.UserID = " & userId
    If statusFilter <> "all" Then
        sql = sql & " AND uc.Status = '" & SafeSQL(statusFilter) & "'"
    End If
    sql = sql & " ORDER BY uc.ObtainedAt DESC"
    Set PE_CouponGetUserCoupons = conn.Execute(sql)
    If Err.Number <> 0 Then
        Err.Clear
        Set PE_CouponGetUserCoupons = Nothing
    End If
    On Error GoTo 0
End Function

' ============================================
' 获取用户可用优惠券数量
' ============================================
Function PE_CouponGetUserCount(userId)
    PE_CouponGetUserCount = 0
    On Error Resume Next
    Dim rs
    Set rs = conn.Execute("SELECT COUNT(*) FROM UserCoupons WHERE UserID = " & userId & " AND Status = 'available'")
    If Not rs Is Nothing Then
        If Not rs.EOF Then PE_CouponGetUserCount = CLng(rs(0))
        rs.Close
    End If
    Set rs = Nothing
    On Error GoTo 0
End Function

' ============================================
' 获取购物车可用的优惠券
' ============================================
Function PE_CouponGetApplicable(userId, cartTotal)
    On Error Resume Next
    Dim sql
    sql = "SELECT uc.*, c.CouponName, c.CouponType, c.DiscountValue, c.MinSpend, c.MaxDiscount, c.Description " & _
          "FROM UserCoupons uc LEFT JOIN Coupons c ON uc.CouponID = c.CouponID " & _
          "WHERE uc.UserID = " & userId & " AND uc.Status = 'available' " & _
          "AND c.MinSpend <= " & cartTotal & " " & _
          "AND GETDATE() BETWEEN c.ValidFrom AND c.ValidTo " & _
          "ORDER BY c.DiscountValue DESC"
    Set PE_CouponGetApplicable = conn.Execute(sql)
    If Err.Number <> 0 Then
        Err.Clear
        Set PE_CouponGetApplicable = Nothing
    End If
    On Error GoTo 0
End Function

' ============================================
' 获取所有优惠券（管理后台用）
' ============================================
Function PE_CouponGetAll()
    On Error Resume Next
    Set PE_CouponGetAll = conn.Execute("SELECT * FROM Coupons ORDER BY CouponID DESC")
    If Err.Number <> 0 Then
        Set PE_CouponGetAll = Nothing
    End If
    On Error GoTo 0
End Function

' ============================================
' 获取优惠券使用统计
' ============================================
Function PE_CouponGetStats(couponId)
    Dim result
    Set result = Server.CreateObject("Scripting.Dictionary")
    result.Add "totalIssued", 0
    result.Add "totalUsed", 0
    result.Add "totalAmount", 0
    
    On Error Resume Next
    Dim rs
    Set rs = conn.Execute("SELECT COUNT(*) FROM UserCoupons WHERE CouponID = " & couponId)
    If Not rs Is Nothing Then
        If Not rs.EOF Then result("totalIssued") = CLng(rs(0))
        rs.Close
    End If
    Set rs = Nothing
    
    Set rs = conn.Execute("SELECT COUNT(*) FROM UserCoupons WHERE CouponID = " & couponId & " AND Status = 'used'")
    If Not rs Is Nothing Then
        If Not rs.EOF Then result("totalUsed") = CLng(rs(0))
        rs.Close
    End If
    Set rs = Nothing
    
    Set rs = conn.Execute("SELECT ISNULL(SUM(CouponDiscount), 0) FROM Orders WHERE CouponCode IN (SELECT CouponCode FROM Coupons WHERE CouponID = " & couponId & ")")
    If Not rs Is Nothing Then
        If Not rs.EOF Then result("totalAmount") = CDbl(rs(0))
        rs.Close
    End If
    Set rs = Nothing
    On Error GoTo 0
    
    Set PE_CouponGetStats = result
End Function

' ============================================
' 格式化优惠券类型
' ============================================
Function PE_CouponTypeName(couponType)
    Select Case LCase(couponType)
        Case "fixed":         PE_CouponTypeName = "满减券"
        Case "percentage":    PE_CouponTypeName = "折扣券"
        Case "free_shipping": PE_CouponTypeName = "免邮券"
        Case "gift":          PE_CouponTypeName = "礼品券"
        Case Else:            PE_CouponTypeName = couponType
    End Select
End Function

' ============================================
' 格式化优惠券描述文本
' ============================================
Function PE_CouponFormatDesc(couponType, discountValue, minSpend, maxDiscount)
    Select Case LCase(couponType)
        Case "fixed":
            If minSpend > 0 Then
                PE_CouponFormatDesc = "满" & FormatNumber(minSpend, 0) & "减" & FormatNumber(discountValue, 0)
            Else
                PE_CouponFormatDesc = "减" & FormatNumber(discountValue, 0) & "元"
            End If
        Case "percentage":
            Dim str
            str = FormatNumber(discountValue, 0) & "折"
            If maxDiscount > 0 Then str = str & "（最高减" & FormatNumber(maxDiscount, 0) & "）"
            PE_CouponFormatDesc = str
        Case "free_shipping":
            PE_CouponFormatDesc = "免运费"
        Case "gift":
            PE_CouponFormatDesc = "礼品"
        Case Else:
            PE_CouponFormatDesc = "优惠"
    End Select
End Function
%>