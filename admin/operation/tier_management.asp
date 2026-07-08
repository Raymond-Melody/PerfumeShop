<%@ Language="VBScript" CodePage="65001" %>
<%
Response.CodePage = 65001
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
' Session.CodePage removed: @CodePage=65001 directive is sufficient for file encoding
' Removing explicit Session.CodePage prevents OLEDB from misinterpreting NVARCHAR data as UTF-8
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

' V18: 修复 OLEDB 驱动对 NVARCHAR 数据的编码转换问题
' 根因: OLEDB 驱动受 Session.CodePage=65001 影响将 NVARCHAR 转为 UTF-8 字节,
' 但 ADO 用系统 ANSI 代码页(GBK/CP936)来解释这些 UTF-8 字节, 产生乱码
' 解决: 将乱码字符串以 GBK 编码还原为原始 UTF-8 字节, 再用 UTF-8 解码
Function FixDbUnicode(val)
    If IsNull(val) Or val = "" Then
        FixDbUnicode = ""
        Exit Function
    End If
    
    Dim stream, bytes, result
    On Error Resume Next
    
    ' 步骤1: 将乱码字符串以 GBK 编码 → 得到原始的 UTF-8 字节序列
    Set stream = Server.CreateObject("ADODB.Stream")
    stream.Type = 2  ' adTypeText
    stream.Charset = "gb2312"
    stream.Open
    stream.WriteText val
    stream.Position = 0
    stream.Type = 1  ' adTypeBinary
    bytes = stream.Read
    stream.Close
    
    ' 步骤2: 将原始字节以 UTF-8 解码 → 还原为正确的中文
    stream.Open
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Write bytes
    stream.Position = 0
    result = stream.ReadText
    stream.Close
    Set stream = Nothing
    
    If Err.Number = 0 And result <> "" Then
        FixDbUnicode = result
    Else
        Err.Clear
        FixDbUnicode = val  ' 回退到原始值
    End If
    On Error GoTo 0
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
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>会员等级管理 - 运营管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; }
        .main-content { padding: 24px; }

        /* 统计概览卡 */
        .tier-stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 24px; }
        .tier-stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); padding: 20px; border-radius: 10px; border: 1px solid rgba(255,255,255,0.05); text-align: center; transition: transform 0.2s; }
        .tier-stat-card:hover { transform: translateY(-3px); border-color: rgba(0,188,212,0.3); }
        .tier-stat-card .stat-icon { font-size: 28px; margin-bottom: 8px; }
        .tier-stat-card .stat-icon.silver { color: #9E9E9E; }
        .tier-stat-card .stat-icon.gold { color: #FF9800; }
        .tier-stat-card .stat-icon.diamond { color: #2196F3; }
        .tier-stat-card .stat-icon.black { color: #B0BEC5; }
        .tier-stat-card .stat-value { font-size: 24px; font-weight: bold; color: #fff; }
        .tier-stat-card .stat-label { font-size: 13px; color: #888; margin-top: 4px; }
        .tier-stat-card .stat-discount { font-size: 12px; color: #00bcd4; margin-top: 2px; font-weight: 600; }

        /* 等级卡片网格 - 响应式自适应排列 */
        .tier-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(380px, 1fr)); gap: 20px; margin-bottom: 30px; }
        .tier-panel { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); overflow: hidden; transition: border-color 0.3s; }
        .tier-panel:hover { border-color: rgba(0,188,212,0.3); }
        .tier-panel-header { padding: 16px 20px; display: flex; align-items: center; gap: 12px; border-bottom: 1px solid rgba(255,255,255,0.08); }
        .tier-icon { width: 44px; height: 44px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 20px; color: #fff; flex-shrink: 0; }
        .tier-title { font-size: 18px; font-weight: 700; color: #fff; }
        .tier-code { font-size: 12px; color: #888; margin-top: 2px; }
        .tier-discount-badge { margin-left: auto; padding: 5px 14px; border-radius: 20px; font-size: 14px; font-weight: 700; }
        .discount-high { background: rgba(76,175,80,0.2); color: #66bb6a; }
        .discount-mid { background: rgba(255,152,0,0.2); color: #ffa726; }
        .discount-low { background: rgba(244,67,54,0.2); color: #ef5350; }
        .discount-none { background: rgba(255,255,255,0.05); color: #888; }

        .tier-panel-body { padding: 16px 20px; }
        .tier-panel-body form { display: flex; flex-direction: column; gap: 10px; }
        .form-row { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }
        .form-row > label { font-size: 13px; color: #b0b0b0; font-weight: 600; min-width: 72px; flex-shrink: 0; }
        .form-row input[type="number"],
        .form-row input[type="text"] {
            padding: 8px 12px; border: 1px solid #3a3a4a; border-radius: 6px; font-size: 14px; width: 120px;
            background: #1a1a2e; color: #e0e0e0; transition: border-color 0.2s;
        }
        .form-row input:focus { outline: none; border-color: #00bcd4; box-shadow: 0 0 0 2px rgba(0,188,212,0.15); }
        .form-row input.wide { width: 180px; }
        .form-row .unit { font-size: 13px; color: #888; }
        .form-row .switch-group { display: flex; gap: 4px; align-items: center; }
        .switch-label { display: inline-flex; align-items: center; gap: 4px; padding: 5px 12px; border-radius: 4px; font-size: 12px; cursor: pointer; border: 1px solid #3a3a4a; background: #1a1a2e; color: #888; transition: all 0.2s; user-select: none; }
        .switch-label input { display: none; }
        .switch-label.active { background: rgba(76,175,80,0.15); border-color: #4CAF50; color: #66bb6a; }
        .switch-label.inactive { background: rgba(244,67,54,0.15); border-color: #e53935; color: #ef5350; }

        .btn-save { display: inline-flex; align-items: center; gap: 6px; padding: 8px 20px; background: linear-gradient(135deg, #ff8f00, #f57c00); color: #fff; border: none; border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; transition: all 0.2s; margin-top: 4px; align-self: flex-start; }
        .btn-save:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(255,143,0,0.3); }

        .benefit-row { display: flex; gap: 6px; flex-wrap: wrap; padding: 10px 0 0; border-top: 1px solid rgba(255,255,255,0.06); margin-top: 8px; }
        .benefit-tag { display: inline-flex; align-items: center; gap: 4px; padding: 3px 10px; border-radius: 4px; font-size: 11px; background: rgba(255,255,255,0.05); color: #888; }
        .benefit-tag.active { background: rgba(76,175,80,0.15); color: #66bb6a; }

        /* 日志面板 */
        .log-panel { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; padding: 24px; border: 1px solid rgba(255,255,255,0.05); margin-top: 10px; }
        .log-panel h3 { color: #fff; margin: 0 0 16px; font-size: 18px; display: flex; align-items: center; gap: 8px; }
        .log-panel h3 i { color: #00bcd4; }
        .log-table { width: 100%; border-collapse: collapse; font-size: 13px; }
        .log-table th { text-align: left; padding: 10px 14px; background: rgba(0,0,0,0.2); color: #888; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; }
        .log-table td { padding: 10px 14px; border-bottom: 1px solid rgba(255,255,255,0.05); color: #ccc; }
        .log-table tr:hover td { background: rgba(255,255,255,0.02); }
        .log-old { color: #ef5350; text-decoration: line-through; margin-right: 4px; }
        .log-new { color: #66bb6a; font-weight: 600; }
        .log-arrow { color: #666; margin: 0 4px; }

        /* 平板端：统计卡2列，等级卡2列 */
        @media (max-width: 1100px) {
            .tier-stats-grid { grid-template-columns: repeat(2, 1fr); }
            .tier-grid { grid-template-columns: repeat(auto-fill, minmax(340px, 1fr)); }
        }
        /* 手机端：全部单列 */
        @media (max-width: 768px) {
            .tier-stats-grid { grid-template-columns: repeat(2, 1fr); }
            .tier-grid { grid-template-columns: 1fr; }
            .form-row { flex-direction: column; align-items: flex-start; }
            .form-row > label { min-width: auto; }
            .main-content { padding: 16px; }
        }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-layer-group"></i> 会员等级折扣管理</h2><!-- V18.7.3c removeFixDbUnicode -->
            <div class="breadcrumb">
                <a href="index.asp">运营中心</a> / <span>会员等级</span>
            </div>
        </div>

        <% If successMsg <> "" Then %>
        <div class="alert alert-success"><i class="fas fa-check-circle"></i> <%= successMsg %></div>
        <% End If %>
        <% If errorMsg <> "" Then %>
        <div class="alert alert-error"><i class="fas fa-exclamation-circle"></i> <%= errorMsg %></div>
        <% End If %>

        <!-- 等级概览统计 -->
        <div class="tier-stats-grid">
            <%
            Dim tierCardCode, tierCardName, tierMinSpent, tierMaxSpent, tierDiscountRate
            Dim tierFreeShip, tierPriorityShip, tierBirthdayGift, tierDedicatedSupport
            Dim tierIconClass, tierThemeColor, tierBadgeBg, tierIsActive
            Dim discountBadgeClass, discountPercent, spentRangeDesc
            If Not rsTiers Is Nothing Then
                Do While Not rsTiers.EOF
                    tierCardCode = rsTiers("TierCode")
                    tierCardName = rsTiers("TierName")
                    tierDiscountRate = CDbl(rsTiers("DiscountRate"))
                    tierIconClass = rsTiers("IconClass") & ""
                    tierThemeColor = rsTiers("Color") & ""
            %>
            <div class="tier-stat-card">
                <div class="stat-icon <%= tierCardCode %>"><i class="fas <%= tierIconClass %>"></i></div>
                <div class="stat-value"><%= tierCardName %></div>
                <div class="stat-discount"><%= DiscountLabel(tierDiscountRate) %></div>
                <div class="stat-label"><%= tierCardCode %></div>
            </div>
            <%
                    rsTiers.MoveNext
                Loop
                rsTiers.MoveFirst
            End If
            %>
        </div>

        <!-- 等级配置卡片 -->
        <div class="tier-grid">
            <%
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
                    If discountPercent >= 20 Then
                        discountBadgeClass = "discount-high"
                    ElseIf discountPercent >= 10 Then
                        discountBadgeClass = "discount-mid"
                    ElseIf discountPercent > 0 Then
                        discountBadgeClass = "discount-low"
                    Else
                        discountBadgeClass = "discount-none"
                    End If

                    If tierMaxSpent <> "" Then
                        spentRangeDesc = "¥" & FormatNumber(tierMinSpent, 0) & " ~ ¥" & FormatNumber(CDbl(tierMaxSpent), 0)
                    Else
                        spentRangeDesc = "¥" & FormatNumber(tierMinSpent, 0) & " 以上"
                    End If
            %>
            <div class="tier-panel">
                <div class="tier-panel-header">
                    <div class="tier-icon" style="background: <%= tierThemeColor %>;">
                        <i class="fas <%= tierIconClass %>"></i>
                    </div>
                    <div>
                        <div class="tier-title"><%= tierCardName %></div>
                        <div class="tier-code"><%= tierCardCode %> · <%= spentRangeDesc %></div>
                    </div>
                    <div class="tier-discount-badge <%= discountBadgeClass %>"><%= DiscountLabel(tierDiscountRate) %></div>
                </div>
                <div class="tier-panel-body">
                    <form method="post">
                        <%= GetCSRFTokenField() %>
                        <input type="hidden" name="tier_code" value="<%= tierCardCode %>">

                        <div class="form-row">
                            <label>折扣率</label>
                            <input type="number" name="discount_rate" value="<%= FormatNumber(tierDiscountRate, 3) %>" step="0.001" min="0.001" max="1.000" class="wide" title="0.950=95折, 0.850=85折, 0.800=8折">
                            <span class="unit">= <%= FormatPercentDisplay(tierDiscountRate) %>（<%= DiscountLabel(tierDiscountRate) %>）</span>
                        </div>

                        <div class="form-row">
                            <label>消费门槛</label>
                            <input type="number" name="min_spent" value="<%= tierMinSpent %>" step="100" min="0" title="最低累计消费金额">
                            <span class="unit">元起</span>
                            <input type="number" name="max_spent" value="<%= tierMaxSpent %>" step="100" min="0" placeholder="无上限" title="最高累计消费金额（留空=无上限）" style="margin-left:8px;">
                            <span class="unit">元止（留空=无上限）</span>
                        </div>

                        <div class="form-row">
                            <label>免运费</label>
                            <div class="switch-group">
                                <label class="switch-label <%= IIf(tierFreeShip, "active", "inactive") %>">
                                    <input type="radio" name="free_shipping" value="1" <%= IIf(tierFreeShip, "checked", "") %> onchange="this.parentElement.classList.add('active');this.parentElement.classList.remove('inactive');this.parentElement.nextElementSibling.classList.remove('active');this.parentElement.nextElementSibling.classList.add('inactive');"> 是
                                </label>
                                <label class="switch-label <%= IIf(Not tierFreeShip, "active", "inactive") %>">
                                    <input type="radio" name="free_shipping" value="0" <%= IIf(Not tierFreeShip, "checked", "") %> onchange="this.parentElement.classList.add('active');this.parentElement.classList.remove('inactive');this.parentElement.previousElementSibling.classList.remove('active');this.parentElement.previousElementSibling.classList.add('inactive');"> 否
                                </label>
                            </div>
                            <label style="margin-left:16px;">优先发货</label>
                            <div class="switch-group">
                                <label class="switch-label <%= IIf(tierPriorityShip, "active", "inactive") %>">
                                    <input type="radio" name="priority_shipping" value="1" <%= IIf(tierPriorityShip, "checked", "") %> onchange="this.parentElement.classList.add('active');this.parentElement.classList.remove('inactive');this.parentElement.nextElementSibling.classList.remove('active');this.parentElement.nextElementSibling.classList.add('inactive');"> 是
                                </label>
                                <label class="switch-label <%= IIf(Not tierPriorityShip, "active", "inactive") %>">
                                    <input type="radio" name="priority_shipping" value="0" <%= IIf(Not tierPriorityShip, "checked", "") %> onchange="this.parentElement.classList.add('active');this.parentElement.classList.remove('inactive');this.parentElement.previousElementSibling.classList.remove('active');this.parentElement.previousElementSibling.classList.add('inactive');"> 否
                                </label>
                            </div>
                        </div>

                        <div class="form-row">
                            <label>生日礼</label>
                            <div class="switch-group">
                                <label class="switch-label <%= IIf(tierBirthdayGift, "active", "inactive") %>">
                                    <input type="radio" name="birthday_gift" value="1" <%= IIf(tierBirthdayGift, "checked", "") %> onchange="this.parentElement.classList.add('active');this.parentElement.classList.remove('inactive');this.parentElement.nextElementSibling.classList.remove('active');this.parentElement.nextElementSibling.classList.add('inactive');"> 是
                                </label>
                                <label class="switch-label <%= IIf(Not tierBirthdayGift, "active", "inactive") %>">
                                    <input type="radio" name="birthday_gift" value="0" <%= IIf(Not tierBirthdayGift, "checked", "") %> onchange="this.parentElement.classList.add('active');this.parentElement.classList.remove('inactive');this.parentElement.previousElementSibling.classList.remove('active');this.parentElement.previousElementSibling.classList.add('inactive');"> 否
                                </label>
                            </div>
                            <label style="margin-left:16px;">专属客服</label>
                            <div class="switch-group">
                                <label class="switch-label <%= IIf(tierDedicatedSupport, "active", "inactive") %>">
                                    <input type="radio" name="dedicated_support" value="1" <%= IIf(tierDedicatedSupport, "checked", "") %> onchange="this.parentElement.classList.add('active');this.parentElement.classList.remove('inactive');this.parentElement.nextElementSibling.classList.remove('active');this.parentElement.nextElementSibling.classList.add('inactive');"> 是
                                </label>
                                <label class="switch-label <%= IIf(Not tierDedicatedSupport, "active", "inactive") %>">
                                    <input type="radio" name="dedicated_support" value="0" <%= IIf(Not tierDedicatedSupport, "checked", "") %> onchange="this.parentElement.classList.add('active');this.parentElement.classList.remove('inactive');this.parentElement.previousElementSibling.classList.remove('active');this.parentElement.previousElementSibling.classList.add('inactive');"> 否
                                </label>
                            </div>
                        </div>

                        <div class="form-row">
                            <label>启用状态</label>
                            <div class="switch-group">
                                <label class="switch-label <%= IIf(tierIsActive, "active", "inactive") %>">
                                    <input type="radio" name="is_active" value="1" <%= IIf(tierIsActive, "checked", "") %> onchange="this.parentElement.classList.add('active');this.parentElement.classList.remove('inactive');this.parentElement.nextElementSibling.classList.remove('active');this.parentElement.nextElementSibling.classList.add('inactive');"> 启用
                                </label>
                                <label class="switch-label <%= IIf(Not tierIsActive, "active", "inactive") %>">
                                    <input type="radio" name="is_active" value="0" <%= IIf(Not tierIsActive, "checked", "") %> onchange="this.parentElement.classList.add('active');this.parentElement.classList.remove('inactive');this.parentElement.previousElementSibling.classList.remove('active');this.parentElement.previousElementSibling.classList.add('inactive');"> 停用
                                </label>
                            </div>
                        </div>

                        <button type="submit" class="btn-save"><i class="fas fa-save"></i> 保存设置</button>
                    </form>

                    <div class="benefit-row">
                        <span class="benefit-tag <%= IIf(tierFreeShip, "active", "") %>"><i class="fas fa-truck"></i> <%= IIf(tierFreeShip,"免运费","无免运") %></span>
                        <span class="benefit-tag <%= IIf(tierPriorityShip, "active", "") %>"><i class="fas fa-rocket"></i> <%= IIf(tierPriorityShip,"优先发货","普通发货") %></span>
                        <span class="benefit-tag <%= IIf(tierBirthdayGift, "active", "") %>"><i class="fas fa-cake"></i> <%= IIf(tierBirthdayGift,"生日礼","无生日礼") %></span>
                        <span class="benefit-tag <%= IIf(tierDedicatedSupport, "active", "") %>"><i class="fas fa-headset"></i> <%= IIf(tierDedicatedSupport,"专属客服","标准客服") %></span>
                        <span class="benefit-tag <%= IIf(tierIsActive, "active", "") %>"><i class="fas fa-check-circle"></i> <%= IIf(tierIsActive,"已启用","已停用") %></span>
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
        <div class="log-panel">
            <h3><i class="fas fa-history"></i> 配置变更日志（最近 20 条）</h3>
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
                        Dim logRowCount : logRowCount = 0
                        Dim logTierName
                        Do While Not rsLog.EOF And logRowCount < 20
                            logRowCount = logRowCount + 1
                            logTierName = rsLog("TierName") & ""
                            If logTierName = "" Then logTierName = rsLog("TierCode")
                    %>
                    <tr>
                        <td><%= Left(rsLog("ChangedAt"), 19) %></td>
                        <td><strong style="color:#00bcd4;"><%= logTierName %></strong></td>
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
                    <% If logRowCount = 0 Then %>
                    <tr><td colspan="5" style="text-align:center;padding:30px;color:#888;">暂无变更记录</td></tr>
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
