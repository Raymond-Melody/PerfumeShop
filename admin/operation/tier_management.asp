<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
Call OpenConnection()

Dim errorMsg, successMsg
errorMsg = ""
successMsg = ""

' 辅助函数
Function GetScalar(sql)
    Dim rs, val : val = ""
    On Error Resume Next
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        If Not rs Is Nothing Then
            If Not rs.EOF Then val = rs(0) & ""
            rs.Close
        End If
    Else
        Err.Clear
    End If
    Set rs = Nothing : GetScalar = val
End Function

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Function FormatPercentDisplay(rate)
    If IsNumeric(rate) Then
        FormatPercentDisplay = FormatNumber(CDbl(rate) * 100, 0) & "%"
    Else
        FormatPercentDisplay = "100%"
    End If
End Function

Function DiscountLabel(rate)
    Dim pct
    pct = FormatPercentDisplay(rate)
    If IsNumeric(rate) Then
        Dim d : d = CDbl(rate)
        If d >= 1.0 Then
            DiscountLabel = "原价"
        Else
            DiscountLabel = CLng((1 - d) * 100) & "% OFF"
        End If
    Else
        DiscountLabel = "原价"
    End If
End Function

' 处理 POST 请求
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not ValidateCSRFToken() Then
        errorMsg = "安全验证失败，请刷新页面重试"
    Else
        Dim tierCode, discountRate, minSpent, maxSpent, freeShipping, priorityShipping, birthdayGift, dedicatedSupport, isActive
        tierCode = Request.Form("tier_code")

        If tierCode = "" Then
            errorMsg = "参数错误"
        Else
            ' 读取当前值用于变更日志
            Dim rsOld
            Set rsOld = conn.Execute("SELECT * FROM MemberTiers WHERE TierCode = '" & SafeSQL(tierCode) & "'")
            If Not rsOld Is Nothing And Not rsOld.EOF Then
                Dim adminUser
                adminUser = Session("AdminUsername")
                If adminUser = "" Then adminUser = "admin"

                ' 折扣率
                discountRate = Request.Form("discount_rate")
                If discountRate <> "" Then
                    Dim oldRate, newRate
                    oldRate = FormatNumber(CDbl(rsOld("DiscountRate")), 3)
                    newRate = FormatNumber(CDbl(discountRate), 3)
                    conn.Execute "UPDATE MemberTiers SET DiscountRate = " & newRate & ", UpdatedAt = GETDATE() WHERE TierCode = '" & SafeSQL(tierCode) & "'"
                    If oldRate <> newRate Then
                        conn.Execute "INSERT INTO TierConfigLog (TierCode, FieldName, OldValue, NewValue, ChangedBy) VALUES ('" & SafeSQL(tierCode) & "', N'折扣率', '" & oldRate & "', '" & newRate & "', '" & SafeSQL(adminUser) & "')"
                    End If
                End If

                ' 消费门槛下限
                minSpent = Request.Form("min_spent")
                If minSpent <> "" Then
                    Dim oldMin, newMin
                    oldMin = CStr(rsOld("MinSpent"))
                    newMin = CDbl(minSpent)
                    conn.Execute "UPDATE MemberTiers SET MinSpent = " & newMin & ", UpdatedAt = GETDATE() WHERE TierCode = '" & SafeSQL(tierCode) & "'"
                    If oldMin <> CStr(newMin) Then
                        conn.Execute "INSERT INTO TierConfigLog (TierCode, FieldName, OldValue, NewValue, ChangedBy) VALUES ('" & SafeSQL(tierCode) & "', N'消费门槛', '" & SafeSQL(oldMin) & "', '" & SafeSQL(CStr(newMin)) & "', '" & SafeSQL(adminUser) & "')"
                    End If
                End If

                ' 消费门槛上限
                maxSpent = Request.Form("max_spent")
                If maxSpent <> "" Then
                    Dim oldMax, newMax, oldMaxStr, newMaxStr
                    oldMaxStr = CStr(rsOld("MaxSpent") & "")
                    If oldMaxStr = "" Then oldMaxStr = "无上限"
                    newMax = CDbl(maxSpent)
                    newMaxStr = CStr(newMax)
                    conn.Execute "UPDATE MemberTiers SET MaxSpent = " & newMax & ", UpdatedAt = GETDATE() WHERE TierCode = '" & SafeSQL(tierCode) & "'"
                    If oldMaxStr <> newMaxStr Then
                        conn.Execute "INSERT INTO TierConfigLog (TierCode, FieldName, OldValue, NewValue, ChangedBy) VALUES ('" & SafeSQL(tierCode) & "', N'消费上限', '" & SafeSQL(oldMaxStr) & "', '" & SafeSQL(newMaxStr) & "', '" & SafeSQL(adminUser) & "')"
                    End If
                End If

                ' 免运费
                freeShipping = Request.Form("free_shipping")
                If freeShipping <> "" Then
                    Dim oldFS, newFS
                    oldFS = CStr(rsOld("FreeShipping"))
                    newFS = IIf(freeShipping = "1", "1", "0")
                    conn.Execute "UPDATE MemberTiers SET FreeShipping = " & newFS & ", UpdatedAt = GETDATE() WHERE TierCode = '" & SafeSQL(tierCode) & "'"
                    If oldFS <> newFS Then
                        conn.Execute "INSERT INTO TierConfigLog (TierCode, FieldName, OldValue, NewValue, ChangedBy) VALUES ('" & SafeSQL(tierCode) & "', N'免运费', '" & IIf(oldFS="1","是","否") & "', '" & IIf(newFS="1","是","否") & "', '" & SafeSQL(adminUser) & "')"
                    End If
                End If

                ' 优先发货
                priorityShipping = Request.Form("priority_shipping")
                If priorityShipping <> "" Then
                    If priorityShipping = "1" Then
                        conn.Execute "UPDATE MemberTiers SET PriorityShipping = 1, UpdatedAt = GETDATE() WHERE TierCode = '" & SafeSQL(tierCode) & "'"
                    Else
                        conn.Execute "UPDATE MemberTiers SET PriorityShipping = 0, UpdatedAt = GETDATE() WHERE TierCode = '" & SafeSQL(tierCode) & "'"
                    End If
                End If

                ' 生日礼
                birthdayGift = Request.Form("birthday_gift")
                If birthdayGift <> "" Then
                    If birthdayGift = "1" Then
                        conn.Execute "UPDATE MemberTiers SET BirthdayGift = 1, UpdatedAt = GETDATE() WHERE TierCode = '" & SafeSQL(tierCode) & "'"
                    Else
                        conn.Execute "UPDATE MemberTiers SET BirthdayGift = 0, UpdatedAt = GETDATE() WHERE TierCode = '" & SafeSQL(tierCode) & "'"
                    End If
                End If

                ' 专属客服
                dedicatedSupport = Request.Form("dedicated_support")
                If dedicatedSupport <> "" Then
                    If dedicatedSupport = "1" Then
                        conn.Execute "UPDATE MemberTiers SET DedicatedSupport = 1, UpdatedAt = GETDATE() WHERE TierCode = '" & SafeSQL(tierCode) & "'"
                    Else
                        conn.Execute "UPDATE MemberTiers SET DedicatedSupport = 0, UpdatedAt = GETDATE() WHERE TierCode = '" & SafeSQL(tierCode) & "'"
                    End If
                End If

                ' 启用状态
                isActive = Request.Form("is_active")
                If isActive <> "" Then
                    If isActive = "1" Then
                        conn.Execute "UPDATE MemberTiers SET IsActive = 1, UpdatedAt = GETDATE() WHERE TierCode = '" & SafeSQL(tierCode) & "'"
                    Else
                        conn.Execute "UPDATE MemberTiers SET IsActive = 0, UpdatedAt = GETDATE() WHERE TierCode = '" & SafeSQL(tierCode) & "'"
                    End If
                End If

                successMsg = "等级「" & rsOld("TierName") & "」配置已更新"
                rsOld.Close
            Else
                errorMsg = "未找到等级: " & tierCode
            End If
            Set rsOld = Nothing
        End If
    End If
End If

' 读取所有等级
Dim rsTiers, tierCount
tierCount = 0
Set rsTiers = conn.Execute("SELECT * FROM MemberTiers ORDER BY SortOrder ASC")

' 读取变更日志（最近20条）
Dim rsLog
Set rsLog = conn.Execute("SELECT TOP 20 l.*, t.TierName FROM TierConfigLog l LEFT JOIN MemberTiers t ON l.TierCode = t.TierCode ORDER BY l.ChangedAt DESC")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>会员等级管理 - 运营管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root {
            --tier-silver: #9E9E9E;
            --tier-gold: #FF9800;
            --tier-diamond: #2196F3;
            --tier-black: #212121;
        }
        body { background: #f5f5f5; color: #333; font-family: -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif; margin: 0; }
        .page-wrapper { max-width: 1100px; margin: 0 auto; padding: 20px; }

        .page-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 24px; flex-wrap: wrap; gap: 12px; }
        .page-header h1 { font-size: 24px; margin: 0; display: flex; align-items: center; gap: 10px; }
        .page-header h1 i { color: #ff8f00; }
        .btn-back { display: inline-flex; align-items: center; gap: 6px; padding: 8px 16px; border-radius: 6px; color: #666; text-decoration: none; font-size: 14px; background: #fff; border: 1px solid #ddd; transition: all .2s; }
        .btn-back:hover { background: #f0f0f0; color: #333; }

        .alert { padding: 12px 16px; border-radius: 8px; margin-bottom: 20px; display: flex; align-items: center; gap: 8px; font-size: 14px; }
        .alert-success { background: #e8f5e9; color: #2e7d32; border: 1px solid #a5d6a7; }
        .alert-error { background: #ffebee; color: #c62828; border: 1px solid #ef9a9a; }

        .tier-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(480px, 1fr)); gap: 20px; margin-bottom: 30px; }
        .tier-card { background: #fff; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); overflow: hidden; transition: box-shadow .3s; }
        .tier-card:hover { box-shadow: 0 4px 16px rgba(0,0,0,0.1); }
        .tier-card-header { padding: 16px 20px; display: flex; align-items: center; gap: 12px; border-bottom: 1px solid #f0f0f0; }
        .tier-icon { width: 44px; height: 44px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 20px; color: #fff; flex-shrink: 0; }
        .tier-icon.silver { background: var(--tier-silver); }
        .tier-icon.gold { background: var(--tier-gold); }
        .tier-icon.diamond { background: var(--tier-diamond); }
        .tier-icon.black { background: var(--tier-black); }
        .tier-title { font-size: 18px; font-weight: 700; }
        .tier-code { font-size: 12px; color: #999; margin-top: 2px; }
        .tier-discount-badge { margin-left: auto; padding: 4px 12px; border-radius: 20px; font-size: 14px; font-weight: 700; }
        .discount-high { background: #e8f5e9; color: #2e7d32; }
        .discount-mid { background: #fff3e0; color: #e65100; }
        .discount-low { background: #fce4ec; color: #c62828; }
        .discount-none { background: #f5f5f5; color: #999; }

        .tier-card-body { padding: 16px 20px; }
        .tier-card-body form { display: flex; flex-direction: column; gap: 10px; }
        .form-row { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }
        .form-row label { font-size: 13px; color: #666; font-weight: 600; min-width: 80px; flex-shrink: 0; }
        .form-row input[type="number"],
        .form-row input[type="text"] { padding: 8px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 14px; width: 120px; transition: border-color .2s; }
        .form-row input:focus { outline: none; border-color: #ff8f00; box-shadow: 0 0 0 2px rgba(255,143,0,0.1); }
        .form-row input.wide { width: 180px; }
        .form-row .unit { font-size: 13px; color: #999; }
        .form-row .switch-group { display: flex; gap: 6px; align-items: center; }
        .switch-label { display: inline-flex; align-items: center; gap: 4px; padding: 4px 10px; border-radius: 4px; font-size: 12px; cursor: pointer; border: 1px solid #ddd; background: #fafafa; transition: all .2s; user-select: none; }
        .switch-label input { display: none; }
        .switch-label.active { background: #e8f5e9; border-color: #4CAF50; color: #2e7d32; }
        .switch-label.inactive { background: #fce4ec; border-color: #e53935; color: #c62828; }
        .switch-label.disabled { background: #f5f5f5; border-color: #ddd; color: #999; }
        .btn-save { display: inline-flex; align-items: center; gap: 6px; padding: 8px 20px; background: #ff8f00; color: #fff; border: none; border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; transition: background .2s; margin-top: 4px; align-self: flex-start; }
        .btn-save:hover { background: #e07b00; }

        .spent-range { font-size: 13px; color: #888; margin-top: 2px; }
        .benefit-row { display: flex; gap: 4px; flex-wrap: wrap; padding: 8px 0; border-top: 1px dashed #f0f0f0; margin-top: 6px; }
        .benefit-tag { display: inline-flex; align-items: center; gap: 3px; padding: 2px 8px; border-radius: 4px; font-size: 11px; background: #f5f5f5; color: #666; }
        .benefit-tag.active { background: #e8f5e9; color: #2e7d32; }

        .log-section { margin-top: 30px; }
        .log-section h2 { font-size: 20px; margin-bottom: 16px; display: flex; align-items: center; gap: 8px; }
        .log-section h2 i { color: #ff8f00; }
        .log-table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.04); font-size: 13px; }
        .log-table th { background: #fafafa; padding: 10px 14px; text-align: left; font-size: 12px; color: #888; font-weight: 600; border-bottom: 2px solid #f0f0f0; }
        .log-table td { padding: 10px 14px; border-bottom: 1px solid #f5f5f5; color: #555; }
        .log-table tr:hover td { background: #fafafa; }
        .log-old { color: #c62828; text-decoration: line-through; margin-right: 4px; }
        .log-new { color: #2e7d32; font-weight: 600; }
        .log-arrow { color: #999; margin: 0 4px; }

        @media (max-width: 600px) {
            .tier-grid { grid-template-columns: 1fr; }
            .form-row { flex-direction: column; align-items: flex-start; }
            .form-row label { min-width: auto; }
        }
    </style>
</head>
<body>
<div class="page-wrapper">
    <div class="page-header">
        <h1><i class="fas fa-layer-group"></i> 会员等级折扣管理</h1>
        <a href="index.asp" class="btn-back"><i class="fas fa-arrow-left"></i> 返回运营中心</a>
    </div>

    <% If successMsg <> "" Then %>
    <div class="alert alert-success"><i class="fas fa-check-circle"></i> <%= successMsg %></div>
    <% End If %>
    <% If errorMsg <> "" Then %>
    <div class="alert alert-error"><i class="fas fa-exclamation-circle"></i> <%= errorMsg %></div>
    <% End If %>

    <div class="tier-grid">
        <%
        Dim tierCardCode, tierCardName, tierMinSpent, tierMaxSpent, tierDiscountRate
        Dim tierFreeShip, tierPriorityShip, tierBirthdayGift, tierDedicatedSupport
        Dim tierIconClass, tierThemeColor, tierBadgeBg, tierIsActive
        Dim discountBadgeClass, discountPercent, spentRangeDesc
        If Not rsTiers Is Nothing Then
            Do While Not rsTiers.EOF
                tierCount = tierCount + 1
                tierCardCode = rsTiers("TierCode")
                tierCardName = rsTiers("TierName")
                tierMinSpent = rsTiers("MinSpent")
                tierMaxSpent = rsTiers("MaxSpent") & ""
                tierDiscountRate = CDbl(rsTiers("DiscountRate"))
                tierFreeShip = CBool(rsTiers("FreeShipping"))
                tierPriorityShip = CBool(rsTiers("PriorityShipping"))
                tierBirthdayGift = CBool(rsTiers("BirthdayGift"))
                tierDedicatedSupport = CBool(rsTiers("DedicatedSupport"))
                tierIconClass = rsTiers("IconClass") & ""
                tierThemeColor = rsTiers("Color") & ""
                tierBadgeBg = rsTiers("BadgeBg") & ""
                tierIsActive = CBool(rsTiers("IsActive"))

                discountPercent = (1 - tierDiscountRate) * 100
                If discountPercent >= 20 Then discountBadgeClass = "discount-high"
                ElseIf discountPercent >= 10 Then discountBadgeClass = "discount-mid"
                ElseIf discountPercent > 0 Then discountBadgeClass = "discount-low"
                Else discountBadgeClass = "discount-none"

                If tierMaxSpent <> "" Then
                    spentRangeDesc = "¥" & FormatNumber(tierMinSpent, 0) & " ~ ¥" & FormatNumber(CDbl(tierMaxSpent), 0)
                Else
                    spentRangeDesc = "¥" & FormatNumber(tierMinSpent, 0) & " 以上"
                End If
        %>
        <div class="tier-card">
            <div class="tier-card-header">
                <div class="tier-icon <%= tCode %>" style="background: <%= tColor %>;">
                    <i class="fas <%= tIcon %>"></i>
                </div>
                <div>
                    <div class="tier-title"><%= tName %></div>
                    <div class="tier-code"><%= tCode %> · <%= rangeDesc %></div>
                </div>
                <div class="tier-discount-badge <%= badgeClass %>"><%= DiscountLabel(tRate) %></div>
            </div>
            <div class="tier-card-body">
                <form method="post">
                    <%= GetCSRFTokenField() %>
                    <input type="hidden" name="tier_code" value="<%= tCode %>">

                    <div class="form-row">
                        <label>折扣率</label>
                        <input type="number" name="discount_rate" value="<%= FormatNumber(tRate, 3) %>" step="0.001" min="0.001" max="1.000" class="wide" title="例如: 0.950=95折, 0.850=85折, 0.800=8折">
                        <span class="unit">= <%= FormatPercentDisplay(tRate) %>（<%= DiscountLabel(tRate) %>）</span>
                    </div>

                    <div class="form-row">
                        <label>消费门槛</label>
                        <input type="number" name="min_spent" value="<%= tMin %>" step="100" min="0" title="该等级的最低累计消费金额">
                        <span class="unit">元起</span>
                        <input type="number" name="max_spent" value="<%= tMax %>" step="100" min="0" placeholder="无上限" title="该等级的最高累计消费金额（留空=无上限）" style="margin-left:8px;">
                        <span class="unit">元止（留空=无上限）</span>
                    </div>

                    <div class="form-row">
                        <label>免运费</label>
                        <div class="switch-group">
                            <label class="switch-label <%= IIf(tFS, "active", "inactive") %>">
                                <input type="radio" name="free_shipping" value="1" <%= IIf(tFS, "checked", "") %> onchange="this.parentElement.classList.add('active');this.parentElement.classList.remove('inactive');this.parentElement.nextElementSibling.classList.remove('active');this.parentElement.nextElementSibling.classList.add('inactive');"> 是
                            </label>
                            <label class="switch-label <%= IIf(Not tFS, "active", "inactive") %>">
                                <input type="radio" name="free_shipping" value="0" <%= IIf(Not tFS, "checked", "") %> onchange="this.parentElement.classList.add('active');this.parentElement.classList.remove('inactive');this.parentElement.previousElementSibling.classList.remove('active');this.parentElement.previousElementSibling.classList.add('inactive');"> 否
                            </label>
                        </div>
                        <label style="margin-left:16px;">优先发货</label>
                        <div class="switch-group">
                            <label class="switch-label <%= IIf(tPS, "active", "inactive") %>">
                                <input type="radio" name="priority_shipping" value="1" <%= IIf(tPS, "checked", "") %> onchange="this.parentElement.classList.add('active');this.parentElement.classList.remove('inactive');this.parentElement.nextElementSibling.classList.remove('active');this.parentElement.nextElementSibling.classList.add('inactive');"> 是
                            </label>
                            <label class="switch-label <%= IIf(Not tPS, "active", "inactive") %>">
                                <input type="radio" name="priority_shipping" value="0" <%= IIf(Not tPS, "checked", "") %> onchange="this.parentElement.classList.add('active');this.parentElement.classList.remove('inactive');this.parentElement.previousElementSibling.classList.remove('active');this.parentElement.previousElementSibling.classList.add('inactive');"> 否
                            </label>
                        </div>
                    </div>

                    <div class="form-row">
                        <label>生日礼</label>
                        <div class="switch-group">
                            <label class="switch-label <%= IIf(tBG, "active", "inactive") %>">
                                <input type="radio" name="birthday_gift" value="1" <%= IIf(tBG, "checked", "") %> onchange="this.parentElement.classList.add('active');this.parentElement.classList.remove('inactive');this.parentElement.nextElementSibling.classList.remove('active');this.parentElement.nextElementSibling.classList.add('inactive');"> 是
                            </label>
                            <label class="switch-label <%= IIf(Not tBG, "active", "inactive") %>">
                                <input type="radio" name="birthday_gift" value="0" <%= IIf(Not tBG, "checked", "") %> onchange="this.parentElement.classList.add('active');this.parentElement.classList.remove('inactive');this.parentElement.previousElementSibling.classList.remove('active');this.parentElement.previousElementSibling.classList.add('inactive');"> 否
                            </label>
                        </div>
                        <label style="margin-left:16px;">专属客服</label>
                        <div class="switch-group">
                            <label class="switch-label <%= IIf(tDS, "active", "inactive") %>">
                                <input type="radio" name="dedicated_support" value="1" <%= IIf(tDS, "checked", "") %> onchange="this.parentElement.classList.add('active');this.parentElement.classList.remove('inactive');this.parentElement.nextElementSibling.classList.remove('active');this.parentElement.nextElementSibling.classList.add('inactive');"> 是
                            </label>
                            <label class="switch-label <%= IIf(Not tDS, "active", "inactive") %>">
                                <input type="radio" name="dedicated_support" value="0" <%= IIf(Not tDS, "checked", "") %> onchange="this.parentElement.classList.add('active');this.parentElement.classList.remove('inactive');this.parentElement.previousElementSibling.classList.remove('active');this.parentElement.previousElementSibling.classList.add('inactive');"> 否
                            </label>
                        </div>
                    </div>

                    <div class="form-row">
                        <label>启用状态</label>
                        <div class="switch-group">
                            <label class="switch-label <%= IIf(tActive, "active", "inactive") %>">
                                <input type="radio" name="is_active" value="1" <%= IIf(tActive, "checked", "") %> onchange="this.parentElement.classList.add('active');this.parentElement.classList.remove('inactive');this.parentElement.nextElementSibling.classList.remove('active');this.parentElement.nextElementSibling.classList.add('inactive');"> 启用
                            </label>
                            <label class="switch-label <%= IIf(Not tActive, "active", "inactive") %>">
                                <input type="radio" name="is_active" value="0" <%= IIf(Not tActive, "checked", "") %> onchange="this.parentElement.classList.add('active');this.parentElement.classList.remove('inactive');this.parentElement.previousElementSibling.classList.remove('active');this.parentElement.previousElementSibling.classList.add('inactive');"> 停用
                            </label>
                        </div>
                    </div>

                    <button type="submit" class="btn-save"><i class="fas fa-save"></i> 保存设置</button>
                </form>

                <div class="benefit-row">
                    <span class="benefit-tag <%= IIf(tFS, "active", "") %>"><i class="fas fa-truck"></i> <%= IIf(tFS,"免运费","无免运") %></span>
                    <span class="benefit-tag <%= IIf(tPS, "active", "") %>"><i class="fas fa-rocket"></i> <%= IIf(tPS,"优先发货","普通发货") %></span>
                    <span class="benefit-tag <%= IIf(tBG, "active", "") %>"><i class="fas fa-cake"></i> <%= IIf(tBG,"生日礼","无生日礼") %></span>
                    <span class="benefit-tag <%= IIf(tDS, "active", "") %>"><i class="fas fa-headset"></i> <%= IIf(tDS,"专属客服","标准客服") %></span>
                    <span class="benefit-tag <%= IIf(tActive, "active", "") %>"><i class="fas fa-check-circle"></i> <%= IIf(tActive,"已启用","已停用") %></span>
                </div>
            </div>
        </div>
        <%
                rsTiers.MoveNext
            Loop
            rsTiers.Close
        End If
        %>
    </div>

    <% If tierCount = 0 Then %>
    <div class="alert alert-error"><i class="fas fa-exclamation-triangle"></i> 未找到会员等级配置。请确认 MemberTiers 表已创建并包含数据。</div>
    <% End If %>

    <!-- 变更日志 -->
    <div class="log-section">
        <h2><i class="fas fa-history"></i> 配置变更日志（最近 20 条）</h2>
        <table class="log-table">
            <thead>
                <tr>
                    <th>时间</th>
                    <th>等级</th>
                    <th>字段</th>
                    <th>变更内容</th>
                    <th>操作人</th>
                </tr>
            </thead>
            <tbody>
                <%
                If Not rsLog Is Nothing Then
                    Dim logCount : logCount = 0
                    Do While Not rsLog.EOF And logCount < 20
                        logCount = logCount + 1
                        Dim lTierName : lTierName = rsLog("TierName") & ""
                        If lTierName = "" Then lTierName = rsLog("TierCode")
                %>
                <tr>
                    <td><%= Left(rsLog("ChangedAt"), 19) %></td>
                    <td><strong><%= lTierName %></strong></td>
                    <td><%= rsLog("FieldName") %></td>
                    <td>
                        <% If rsLog("OldValue") <> "" Then %>
                        <span class="log-old"><%= rsLog("OldValue") %></span>
                        <span class="log-arrow">→</span>
                        <% End If %>
                        <span class="log-new"><%= rsLog("NewValue") %></span>
                    </td>
                    <td><%= rsLog("ChangedBy") %></td>
                </tr>
                <%
                        rsLog.MoveNext
                    Loop
                    rsLog.Close
                End If
                %>
                <% If logCount = 0 Then %>
                <tr><td colspan="5" style="text-align:center;padding:30px;color:#999;">暂无变更记录</td></tr>
                <% End If %>
            </tbody>
        </table>
    </div>
</div>
</body>
</html>
<%
Call CloseConnection()
%>
