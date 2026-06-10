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
    
    ' 会员等级折扣
    Dim level, levelDiscount
    level = MU_CalcUserLevel(userId)
    levelDiscount = (1 - MU_GetLevelDiscount(level)) * cartTotal
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