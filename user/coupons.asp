<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"

If Session("UserID") = "" Then
    Response.Redirect "/user/login.asp?return=" & Server.URLEncode("/user/coupons.asp")
    Response.End
End If
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/member_utils.asp"-->
<!--#include file="../includes/promotion_engine.asp"-->
<%
Call OpenConnection()

Dim userId, action, actionMsg, actionResult
userId = Session("UserID")
action = Request.QueryString("action")
If action = "" Then action = Request.Form("action")
actionMsg = ""
actionResult = True

' 处理领取操作
If action = "claim" Then
    Dim claimCode
    claimCode = Request.Form("code")
    If claimCode = "" Then claimCode = Request.QueryString("code")
    If claimCode = "" Then
        actionMsg = "请输入优惠码"
        actionResult = False
    Else
        If PE_CouponIssue(userId, claimCode, "activity") Then
            actionMsg = "领取成功！"
        Else
            actionMsg = "领取失败，优惠码可能不存在、已领完或您已有此券"
            actionResult = False
        End If
    End If
End If

' 获取统计
Dim availableCount, usedCount, expiredCount
availableCount = PE_CouponGetUserCount(userId)

Dim rsAll, rsAvailable, rsUsed, rsExpired
' 可用
Set rsAvailable = PE_CouponGetUserCoupons(userId, "available")
' 已用
Set rsUsed = PE_CouponGetUserCoupons(userId, "used")
' 已过期
Set rsExpired = PE_CouponGetUserCoupons(userId, "expired")

' 统计已用/已过期数量
usedCount = 0
If Not rsUsed Is Nothing Then
    Do While Not rsUsed.EOF
        usedCount = usedCount + 1
        rsUsed.MoveNext
    Loop
    If rsUsed.RecordCount >= 0 Then rsUsed.MoveFirst
End If

expiredCount = 0
If Not rsExpired Is Nothing Then
    Do While Not rsExpired.EOF
        expiredCount = expiredCount + 1
        rsExpired.MoveNext
    Loop
    If rsExpired.RecordCount >= 0 Then rsExpired.MoveFirst
End If

' 当前显示的 Tab
Dim tab
tab = Request.QueryString("tab")
If tab = "" Then tab = "available"
%>
<!--#include file="../includes/header.asp"-->

<div class="breadcrumb">
    <div class="container">
        <a href="/index.asp">首页</a>
        <span class="separator">/</span>
        <a href="/user/index.asp">个人中心</a>
        <span class="separator">/</span>
        <span>我的优惠券</span>
    </div>
</div>

<div class="container">
    <div class="user-center">
        <!--#include file="nav.asp"-->

        <div class="user-main">
            <div class="welcome-section">
                <h1><i class="fas fa-ticket-alt" style="color:#2e7d32;"></i> 我的优惠券</h1>
                <p>管理您的优惠券，享受更多优惠</p>
            </div>

            <% If actionMsg <> "" Then %>
            <div class="alert <% If actionResult Then %>alert-success<% Else %>alert-error<% End If %>" style="margin-bottom:20px;">
                <i class="fas fa-<% If actionResult Then %>check-circle<% Else %>exclamation-circle<% End If %>"></i> <%= actionMsg %>
            </div>
            <% End If %>

            <!-- 统计卡片 -->
            <div class="stats-grid" style="grid-template-columns: repeat(3, 1fr);">
                <div class="stat-card" style="border-top:3px solid #4CAF50;">
                    <div class="stat-icon" style="background:#e8f5e9;color:#4CAF50;"><i class="fas fa-ticket-alt"></i></div>
                    <div class="stat-info">
                        <span class="stat-value"><%= availableCount %></span>
                        <span class="stat-label">可用</span>
                    </div>
                </div>
                <div class="stat-card" style="border-top:3px solid #2196F3;">
                    <div class="stat-icon" style="background:#e3f2fd;color:#2196F3;"><i class="fas fa-check-circle"></i></div>
                    <div class="stat-info">
                        <span class="stat-value"><%= usedCount %></span>
                        <span class="stat-label">已使用</span>
                    </div>
                </div>
                <div class="stat-card" style="border-top:3px solid #9e9e9e;">
                    <div class="stat-icon" style="background:#f5f5f5;color:#9e9e9e;"><i class="fas fa-clock"></i></div>
                    <div class="stat-info">
                        <span class="stat-value"><%= expiredCount %></span>
                        <span class="stat-label">已过期</span>
                    </div>
                </div>
            </div>

            <!-- 领取优惠码 -->
            <div class="claim-section" style="background:#e8f5e9;border-radius:12px;padding:20px;margin:20px 0;display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
                <div style="flex:1;min-width:200px;">
                    <h4 style="margin:0;color:#2e7d32;"><i class="fas fa-gift"></i> 兑换优惠码</h4>
                    <p style="margin:4px 0 0;color:#388e3c;font-size:13px;">输入优惠码领取优惠券</p>
                </div>
                <form method="post" style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
                    <input type="hidden" name="action" value="claim">
                    <input type="text" name="code" placeholder="输入优惠码" style="padding:10px 16px;border:2px solid #4CAF50;border-radius:8px;font-size:15px;width:200px;text-transform:uppercase;" required>
                    <button type="submit" style="background:linear-gradient(135deg, #43a047, #2e7d32);color:#fff;border:none;padding:10px 24px;border-radius:8px;font-size:15px;cursor:pointer;font-weight:bold;">领取</button>
                </form>
            </div>

            <!-- Tab 切换 -->
            <div style="display:flex;gap:0;margin-bottom:20px;border-bottom:2px solid #e0e0e0;">
                <a href="?tab=available" style="padding:10px 24px;text-decoration:none;font-weight:500;color:<% If tab = "available" Then %>#4CAF50;border-bottom:2px solid #4CAF50;margin-bottom:-2px;<% Else %>#666;<% End If %>">
                    <i class="fas fa-ticket-alt"></i> 可用 (<%= availableCount %>)
                </a>
                <a href="?tab=used" style="padding:10px 24px;text-decoration:none;font-weight:500;color:<% If tab = "used" Then %>#2196F3;border-bottom:2px solid #2196F3;margin-bottom:-2px;<% Else %>#666;<% End If %>">
                    <i class="fas fa-check-circle"></i> 已使用 (<%= usedCount %>)
                </a>
                <a href="?tab=expired" style="padding:10px 24px;text-decoration:none;font-weight:500;color:<% If tab = "expired" Then %>#9e9e9e;border-bottom:2px solid #9e9e9e;margin-bottom:-2px;<% Else %>#666;<% End If %>">
                    <i class="fas fa-clock"></i> 已过期 (<%= expiredCount %>)
                </a>
            </div>

            <!-- 优惠券列表 -->
            <%
            Dim activeRs
            If tab = "available" Then
                Set activeRs = rsAvailable
            ElseIf tab = "used" Then
                Set activeRs = rsUsed
            Else
                Set activeRs = rsExpired
            End If
            %>

            <% If Not activeRs Is Nothing And Not activeRs.EOF Then %>
            <div class="coupon-list">
                <%
                Do While Not activeRs.EOF
                    Dim cType, cName, cValue, cMin, cMax, cDesc, cValidTo, cSource, cStatus, cCode, cTerms
                    cType = activeRs("CouponType") & ""
                    cName = activeRs("CouponName") & ""
                    cValue = CDbl(activeRs("DiscountValue"))
                    cMin = CDbl(activeRs("MinSpend"))
                    cMax = CDbl(activeRs("MaxDiscount"))
                    cDesc = activeRs("Description") & ""
                    cTerms = activeRs("Terms") & ""
                    cCode = activeRs("CouponCode")
                    cSource = activeRs("Source") & ""
                    cStatus = activeRs("Status") & ""
                    cValidTo = activeRs("ValidTo") & ""
                    
                    Dim typeColor, typeIcon, valueText
                    Select Case LCase(cType)
                        Case "fixed":
                            typeColor = "#FF5722"
                            typeIcon = "yen-sign"
                            valueText = "¥" & FormatNumber(cValue, 0)
                        Case "percentage":
                            typeColor = "#9C27B0"
                            typeIcon = "percent"
                            valueText = FormatNumber(cValue, 0) & "折"
                        Case "free_shipping":
                            typeColor = "#2196F3"
                            typeIcon = "truck"
                            valueText = "免邮"
                        Case "gift":
                            typeColor = "#FF9800"
                            typeIcon = "gift"
                            valueText = "礼品"
                        Case Else:
                            typeColor = "#607D8B"
                            typeIcon = "tag"
                            valueText = "优惠"
                    End Select
                    
                    ' 条件文本
                    Dim condText
                    condText = ""
                    If cMin > 0 Then condText = condText & "满¥" & FormatNumber(cMin, 0)
                    If LCase(cType) = "percentage" And cMax > 0 Then
                        If condText <> "" Then condText = condText & " · "
                        condText = condText & "最高¥" & FormatNumber(cMax, 0)
                    End If
                    If condText = "" Then condText = "无门槛"
                %>
                <div class="coupon-card" style="display:flex;border-radius:12px;overflow:hidden;margin-bottom:16px;box-shadow:0 2px 8px rgba(0,0,0,0.08);<% If tab <> "available" Then %>opacity:0.55;<% End If %>">
                    <!-- 左侧面额 -->
                    <div style="width:110px;background:linear-gradient(135deg, <%= typeColor %>, <%= typeColor %>dd);color:#fff;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:16px 8px;text-align:center;min-height:120px;">
                        <div style="font-size:28px;font-weight:bold;line-height:1;"><%= valueText %></div>
                        <div style="font-size:11px;opacity:0.9;margin-top:4px;"><%= PE_CouponTypeName(cType) %></div>
                    </div>
                    <!-- 右侧信息 -->
                    <div style="flex:1;padding:16px 20px;background:#fff;display:flex;flex-direction:column;justify-content:space-between;">
                        <div>
                            <h4 style="margin:0 0 4px;font-size:16px;"><%= cName %></h4>
                            <% If cDesc <> "" Then %>
                            <p style="margin:0 0 8px;font-size:13px;color:#888;"><%= cDesc %></p>
                            <% End If %>
                            <div style="display:flex;gap:8px;flex-wrap:wrap;">
                                <span style="background:#f5f5f5;padding:2px 8px;border-radius:4px;font-size:11px;color:#666;"><%= condText %></span>
                                <% If cTerms <> "" Then %>
                                <span style="background:#f5f5f5;padding:2px 8px;border-radius:4px;font-size:11px;color:#999;" title="<%= cTerms %>">条款</span>
                                <% End If %>
                            </div>
                        </div>
                        <div style="display:flex;justify-content:space-between;align-items:center;margin-top:8px;">
                            <div style="font-size:11px;color:#bbb;">
                                <% If Not IsNull(cValidTo) And cValidTo <> "" Then %>
                                有效期至 <%= SafeFormatDateTime(cValidTo, 2) %>
                                <% End If %>
                            </div>
                            <div style="font-size:11px;">
                                <% If cStatus = "available" Then %>
                                <span style="color:#4CAF50;"><i class="fas fa-check-circle"></i> 可用</span>
                                <% ElseIf cStatus = "used" Then %>
                                <span style="color:#2196F3;"><i class="fas fa-check"></i> 已使用</span>
                                <% Else %>
                                <span style="color:#9e9e9e;"><i class="fas fa-clock"></i> 已过期</span>
                                <% End If %>
                            </div>
                        </div>
                    </div>
                </div>
                <%
                    activeRs.MoveNext
                Loop
                %>
            </div>
            <% Else %>
            <div style="text-align:center;padding:60px 20px;color:#999;">
                <i class="fas fa-ticket-alt" style="font-size:56px;display:block;margin-bottom:16px;color:#ddd;"></i>
                <p style="font-size:16px;">
                    <% If tab = "available" Then %>
                    暂无可用优惠券
                    <% ElseIf tab = "used" Then %>
                    暂无已使用的优惠券
                    <% Else %>
                    暂无过期优惠券
                    <% End If %>
                </p>
                <a href="/products.asp" style="color:#4CAF50;">去购物 →</a>
            </div>
            <% End If %>
        </div>
    </div>
</div>

<style>
.coupon-card { transition: transform 0.2s; }
.coupon-card:hover { transform: translateY(-2px); }
@media (max-width: 768px) {
    .stats-grid { grid-template-columns: repeat(3, 1fr) !important; }
    .coupon-card { flex-direction: column; }
    .coupon-card > div:first-child { width: 100% !important; min-height: 80px !important; }
}
</style>

<%
If Not rsAvailable Is Nothing Then rsAvailable.Close : Set rsAvailable = Nothing
If Not rsUsed Is Nothing Then rsUsed.Close : Set rsUsed = Nothing
If Not rsExpired Is Nothing Then rsExpired.Close : Set rsExpired = Nothing
%>

<!--#include file="../includes/footer.asp"-->
<%
Call CloseConnection()
%>
