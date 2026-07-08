<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"

' 检查登录
If Session("UserID") = "" Then
    Response.Redirect "/user/login.asp?return=" & Server.URLEncode("/user/points.asp")
    Response.End
End If
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/member_utils.asp"-->
<!--#include file="../includes/points_engine.asp"-->
<%
Call OpenConnection()

Dim userId, action, actionResult, actionMsg
userId = Session("UserID")
action = Request.QueryString("action")
If action = "" Then action = Request.Form("action")
actionResult = True
actionMsg = ""

' 处理签到
If action = "signin" Then
    Dim signinResult
    signinResult = PE_DoSignIn(userId)
    If signinResult > 0 Then
        actionMsg = "签到成功！获得 " & signinResult & " 积分"
    ElseIf signinResult = -1 Then
        actionMsg = "今日已签到"
        actionResult = False
    Else
        actionMsg = "签到失败，请稍后再试"
        actionResult = False
    End If
End If

' 处理兑换
If action = "redeem" Then
    Dim redeemId, redeemErr
    redeemId = Request.Form("redemption_id")
    If redeemId = "" Then redeemId = Request.QueryString("redemption_id")
    If Not IsNumeric(redeemId) Or redeemId = "" Then
        actionMsg = "无效的兑换项目"
        actionResult = False
    Else
        If Not ValidateCSRFToken() Then
            actionMsg = "安全校验失败，请刷新页面重试"
            actionResult = False
        Else
            redeemErr = PE_DoRedeem(userId, CLng(redeemId))
            If redeemErr <> "" Then
                actionMsg = redeemErr
                actionResult = False
            Else
                actionMsg = "兑换成功！"
            End If
        End If
    End If
End If

' 获取积分汇总
Dim ptsSummary, availablePoints, totalEarned, totalRedeemed, todayEarned, expiringSoon
Set ptsSummary = PE_GetPointsSummary(userId)
availablePoints = ptsSummary("available")
totalEarned = ptsSummary("totalEarned")
totalRedeemed = ptsSummary("totalRedeemed")
todayEarned = ptsSummary("todayEarned")
expiringSoon = ptsSummary("expiringSoon")

' 签到状态
Dim hasSignedIn
hasSignedIn = PE_CheckSignIn(userId)

' 该月兑换总数（兑换商品表）
Dim rsRedemptionItems
Set rsRedemptionItems = PE_GetRedemptionItems()

' 分页积分账本
Dim pageNum, pageSize, totalRecords, totalPages
pageNum = 1
If IsNumeric(Request.QueryString("page")) And Request.QueryString("page") <> "" Then
    pageNum = CLng(Request.QueryString("page"))
End If
pageSize = 15
totalRecords = PE_GetPointsLedgerCount(userId)
totalPages = Int((totalRecords + pageSize - 1) / pageSize)
If totalPages < 1 Then totalPages = 1

Dim rsLedger
Set rsLedger = PE_GetPointsLedger(userId, pageNum, pageSize)

' 确保 CSRF
Call EnsureCSRFToken()
%>
<!--#include file="../includes/header.asp"-->

<div class="breadcrumb">
    <div class="container">
        <a href="/index.asp">首页</a>
        <span class="separator">/</span>
        <a href="/user/index.asp">个人中心</a>
        <span class="separator">/</span>
        <span>积分中心</span>
    </div>
</div>

<div class="container">
    <div class="user-center">
        <!--#include file="nav.asp"-->

        <!-- 主内容 -->
        <div class="user-main">
            <div class="welcome-section">
                <h1><i class="fas fa-coins" style="color:#ff8f00;"></i> 积分中心</h1>
                <p>管理您的积分，兑换好礼或抵扣订单</p>
            </div>

            <!-- 操作反馈 -->
            <% If actionMsg <> "" Then %>
            <div class="alert <% If actionResult Then %>alert-success<% Else %>alert-error<% End If %>" style="margin-bottom:20px;">
                <i class="fas fa-<% If actionResult Then %>check-circle<% Else %>exclamation-circle<% End If %>"></i> <%= actionMsg %>
            </div>
            <% End If %>

            <!-- 积分概览卡片 -->
            <div class="stats-grid" style="grid-template-columns: repeat(4, 1fr);">
                <div class="stat-card" style="border-top:3px solid #ff8f00;">
                    <div class="stat-icon" style="background:#fff3e0;color:#ff8f00;"><i class="fas fa-coins"></i></div>
                    <div class="stat-info">
                        <span class="stat-value"><%= PE_FormatPoints(availablePoints) %></span>
                        <span class="stat-label">可用积分</span>
                    </div>
                </div>
                <div class="stat-card" style="border-top:3px solid #4CAF50;">
                    <div class="stat-icon" style="background:#e8f5e9;color:#4CAF50;"><i class="fas fa-plus-circle"></i></div>
                    <div class="stat-info">
                        <span class="stat-value"><%= PE_FormatPoints(totalEarned) %></span>
                        <span class="stat-label">累计获得</span>
                    </div>
                </div>
                <div class="stat-card" style="border-top:3px solid #F44336;">
                    <div class="stat-icon" style="background:#ffebee;color:#F44336;"><i class="fas fa-minus-circle"></i></div>
                    <div class="stat-info">
                        <span class="stat-value"><%= PE_FormatPoints(totalRedeemed) %></span>
                        <span class="stat-label">已使用</span>
                    </div>
                </div>
                <div class="stat-card" style="border-top:3px solid #FF9800;">
                    <div class="stat-icon" style="background:#fff3e0;color:#FF9800;"><i class="fas fa-clock"></i></div>
                    <div class="stat-info">
                        <span class="stat-value"><%= PE_FormatPoints(expiringSoon) %></span>
                        <span class="stat-label">30天内到期</span>
                    </div>
                </div>
            </div>

            <!-- 签到区域 -->
            <div class="signin-section" style="background:linear-gradient(135deg, #fff8e1 0%, #ffe0b2 100%);border-radius:12px;padding:24px;margin:24px 0;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:16px;">
                <div>
                    <h3 style="margin:0;color:#e65100;"><i class="fas fa-calendar-check"></i> 每日签到</h3>
                    <p style="margin:4px 0 0;color:#bf360c;font-size:14px;">每天签到可获 <strong><%= CInt(PE_GetRule("signin_points")) %></strong> 积分，连续签到还有额外奖励</p>
                </div>
                <% If hasSignedIn Then %>
                <div style="background:#c8e6c9;color:#2e7d32;padding:12px 24px;border-radius:8px;font-weight:bold;">
                    <i class="fas fa-check-circle"></i> 今日已签到
                </div>
                <% Else %>
                <form method="post" style="margin:0;">
                    <input type="hidden" name="action" value="signin">
                    <button type="submit" style="background:linear-gradient(135deg, #ff6f00, #ff8f00);color:#fff;border:none;padding:12px 32px;border-radius:8px;font-size:16px;font-weight:bold;cursor:pointer;box-shadow:0 4px 12px rgba(255,111,0,0.3);">
                        <i class="fas fa-gift"></i> 立即签到领积分
                    </button>
                </form>
                <% End If %>
            </div>

            <!-- 积分获得方式 -->
            <div class="earn-ways" style="background:#fafafa;border-radius:12px;padding:20px;margin-bottom:24px;">
                <h3 style="margin:0 0 12px;"><i class="fas fa-lightbulb" style="color:#ff8f00;"></i> 如何获得积分</h3>
                <div style="display:grid;grid-template-columns:repeat(auto-fit, minmax(200px, 1fr));gap:12px;">
                    <div style="background:#fff;padding:12px 16px;border-radius:8px;border:1px solid #eee;">
                        <strong style="color:#333;">🛒 消费购物</strong>
                        <div style="color:#666;font-size:13px;">每消费1元得 <%= CInt(PE_GetRule("purchase_rate")) %> 积分</div>
                    </div>
                    <div style="background:#fff;padding:12px 16px;border-radius:8px;border:1px solid #eee;">
                        <strong style="color:#333;">📝 签到打卡</strong>
                        <div style="color:#666;font-size:13px;">每日签到得 <%= CInt(PE_GetRule("signin_points")) %> 积分</div>
                    </div>
                    <div style="background:#fff;padding:12px 16px;border-radius:8px;border:1px solid #eee;">
                        <strong style="color:#333;">⭐ 发表评价</strong>
                        <div style="color:#666;font-size:13px;">评价得 <%= CInt(PE_GetRule("review_points")) %> 积分，带图+<%= CInt(PE_GetRule("review_with_photo")) %></div>
                    </div>
                    <div style="background:#fff;padding:12px 16px;border-radius:8px;border:1px solid #eee;">
                        <strong style="color:#333;">📤 分享推广</strong>
                        <div style="color:#666;font-size:13px;">分享产品/邀请好友各得积分</div>
                    </div>
                </div>
            </div>

            <!-- 积分账本 -->
            <div class="ledger-section" style="background:#fff;border-radius:12px;padding:24px;box-shadow:0 2px 8px rgba(0,0,0,0.06);margin-bottom:24px;">
                <h3 style="margin:0 0 16px;">积分明细</h3>
                
                <% If Not rsLedger Is Nothing And Not rsLedger.EOF Then %>
                <div class="ledger-list">
                    <%
                    Do While Not rsLedger.EOF
                        Dim lType, lPoints, lSource, lDesc, lTime, lTypeClass, lTypeIcon, lPointStr
                        lType = rsLedger("PointType")
                        lPoints = CLng(rsLedger("Points"))
                        lSource = rsLedger("Source") & ""
                        lDesc = rsLedger("Description") & ""
                        lTime = rsLedger("CreatedAt")
                        
                        If lType = "earn" Then
                            lTypeClass = "earn"
                            lTypeIcon = "plus-circle"
                            lPointStr = "+" & PE_FormatPoints(lPoints)
                        ElseIf lType = "redeem" Then
                            lTypeClass = "redeem"
                            lTypeIcon = "minus-circle"
                            lPointStr = "-" & PE_FormatPoints(Abs(lPoints))
                        Else
                            lTypeClass = "expire"
                            lTypeIcon = "clock"
                            lPointStr = "-" & PE_FormatPoints(Abs(lPoints))
                        End If
                    %>
                    <div class="ledger-item" style="display:flex;align-items:center;gap:12px;padding:12px 0;border-bottom:1px solid #f0f0f0;">
                        <div style="width:36px;height:36px;border-radius:50%;display:flex;align-items:center;justify-content:center;
                            <% If lTypeClass = "earn" Then %>background:#e8f5e9;color:#4CAF50;<% Else %>background:#ffebee;color:#F44336;<% End If %>">
                            <i class="fas fa-<%= lTypeIcon %>"></i>
                        </div>
                        <div style="flex:1;min-width:0;">
                            <div style="font-weight:500;color:#333;"><%= PE_GetSourceName(lSource) %></div>
                            <% If lDesc <> "" Then %>
                            <div style="font-size:12px;color:#999;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;"><%= lDesc %></div>
                            <% End If %>
                            <div style="font-size:11px;color:#bbb;"><%= lTime %></div>
                        </div>
                        <div style="font-weight:bold;font-size:16px;<% If lTypeClass = "earn" Then %>color:#4CAF50;<% Else %>color:#F44336;<% End If %>white-space:nowrap;">
                            <%= lPointStr %>
                        </div>
                    </div>
                    <%
                        rsLedger.MoveNext
                    Loop
                    %>
                </div>
                
                <!-- 分页 -->
                <% If totalPages > 1 Then %>
                <div style="display:flex;justify-content:center;gap:8px;margin-top:16px;">
                    <% Dim pg
                    For pg = 1 To totalPages
                        If pg = pageNum Then %>
                        <span style="padding:6px 14px;background:#ff8f00;color:#fff;border-radius:6px;font-weight:bold;"><%= pg %></span>
                        <% Else %>
                        <a href="?page=<%= pg %>" style="padding:6px 14px;background:#f0f0f0;color:#666;border-radius:6px;text-decoration:none;"><%= pg %></a>
                        <% End If
                    Next %>
                </div>
                <% End If %>
                
                <% Else %>
                <div style="text-align:center;padding:40px;color:#999;">
                    <i class="fas fa-coins" style="font-size:48px;display:block;margin-bottom:12px;color:#ddd;"></i>
                    <p>暂无积分记录</p>
                    <a href="/products.asp" style="color:#ff8f00;">去购物获取积分 →</a>
                </div>
                <% End If %>
            </div>

            <!-- 积分兑换商城 -->
            <div class="redeem-section" style="background:#fff;border-radius:12px;padding:24px;box-shadow:0 2px 8px rgba(0,0,0,0.06);">
                <h3 style="margin:0 0 16px;"><i class="fas fa-gift" style="color:#ff8f00;"></i> 积分兑换好礼</h3>
                
                <% If Not rsRedemptionItems Is Nothing And Not rsRedemptionItems.EOF Then %>
                <div style="display:grid;grid-template-columns:repeat(auto-fill, minmax(260px, 1fr));gap:16px;">
                    <%
                    Do While Not rsRedemptionItems.EOF
                        Dim rId, rName, rType, rCost, rStock, rValue, rImg, rDesc
                        rId = rsRedemptionItems("RedemptionID")
                        rName = rsRedemptionItems("ItemName")
                        rType = rsRedemptionItems("ItemType")
                        rCost = CLng(rsRedemptionItems("PointsCost"))
                        rStock = CLng(rsRedemptionItems("Stock"))
                        rValue = CDbl(rsRedemptionItems("RedemptionValue"))
                        rImg = rsRedemptionItems("ImageURL") & ""
                        rDesc = rsRedemptionItems("Description") & ""
                        
                        Dim rTypeLabel, rTypeColor
                        Select Case rType
                            Case "coupon":  rTypeLabel = "优惠券": rTypeColor = "#2196F3"
                            Case "sample":  rTypeLabel = "小样":   rTypeColor = "#9C27B0"
                            Case "bottle":  rTypeLabel = "瓶身":   rTypeColor = "#795548"
                            Case "discount": rTypeLabel = "折扣":  rTypeColor = "#F44336"
                            Case Else:       rTypeLabel = rType:   rTypeColor = "#607D8B"
                        End Select
                        
                        Dim canRedeem
                        canRedeem = (availablePoints >= rCost)
                    %>
                    <div class="redeem-card" style="border:1px solid #eee;border-radius:12px;overflow:hidden;transition:box-shadow 0.2s;display:flex;flex-direction:column;">
                        <div style="background:linear-gradient(135deg, #fff8e1, #ffe0b2);padding:20px;text-align:center;">
                            <% If rImg <> "" Then %>
                            <img src="<%= rImg %>" alt="<%= rName %>" style="max-width:100%;height:80px;object-fit:contain;">
                            <% Else %>
                            <i class="fas fa-<% If rType = "coupon" Then %>ticket-alt<% ElseIf rType = "sample" Then %>vial<% ElseIf rType = "bottle" Then %>wine-bottle<% Else %>gift<% End If %>" style="font-size:48px;color:<%= rTypeColor %>;"></i>
                            <% End If %>
                        </div>
                        <div style="padding:16px;flex:1;display:flex;flex-direction:column;">
                            <span style="display:inline-block;background:<%= rTypeColor %>;color:#fff;font-size:11px;padding:2px 8px;border-radius:4px;align-self:flex-start;margin-bottom:8px;"><%= rTypeLabel %></span>
                            <h4 style="margin:0 0 8px;font-size:15px;"><%= rName %></h4>
                            <% If rDesc <> "" Then %>
                            <p style="font-size:12px;color:#888;margin:0 0 12px;flex:1;"><%= rDesc %></p>
                            <% End If %>
                            <div style="display:flex;justify-content:space-between;align-items:center;">
                                <div>
                                    <span style="font-size:20px;font-weight:bold;color:#ff8f00;"><%= rCost %></span>
                                    <span style="font-size:13px;color:#999;">积分</span>
                                    <% If rValue > 0 Then %>
                                    <span style="font-size:12px;color:#999;">≈¥<%= FormatNumber(rValue, 0) %></span>
                                    <% End If %>
                                </div>
                                <form method="post" style="margin:0;" onsubmit="return confirm('确认使用 <%= rCost %> 积分兑换【<%= rName %>】吗？');">
                                    <%= GetCSRFTokenField() %>
                                    <input type="hidden" name="action" value="redeem">
                                    <input type="hidden" name="redemption_id" value="<%= rId %>">
                                    <button type="submit" <% If Not canRedeem Then %>disabled<% End If %> style="padding:8px 16px;border:none;border-radius:6px;font-weight:bold;cursor:pointer;<% If canRedeem Then %>background:linear-gradient(135deg, #ff8f00, #f57c00);color:#fff;<% Else %>background:#e0e0e0;color:#bbb;cursor:not-allowed;<% End If %>">
                                        <% If canRedeem Then %>立即兑换<% Else %>积分不足<% End If %>
                                    </button>
                                </form>
                            </div>
                            <div style="font-size:11px;color:#bbb;margin-top:8px;">剩余库存: <%= rStock %></div>
                        </div>
                    </div>
                    <%
                        rsRedemptionItems.MoveNext
                    Loop
                    %>
                </div>
                <% Else %>
                <div style="text-align:center;padding:40px;color:#999;">
                    <i class="fas fa-box-open" style="font-size:48px;display:block;margin-bottom:12px;color:#ddd;"></i>
                    <p>暂无可兑换商品</p>
                </div>
                <% End If %>
            </div>
        </div>
    </div>
</div>

<style>
/* 积分中心卡片 hover 效果 */
.redeem-card:hover {
    box-shadow: 0 4px 16px rgba(0,0,0,0.1);
    border-color: #ff8f00;
}
/* 响应式 */
@media (max-width: 768px) {
    .stats-grid { grid-template-columns: repeat(2, 1fr) !important; }
    .redeem-section [style*="grid-template-columns"] { grid-template-columns: 1fr !important; }
}
</style>

<%
' 清理
If Not rsLedger Is Nothing Then
    rsLedger.Close
    Set rsLedger = Nothing
End If
If Not rsRedemptionItems Is Nothing Then
    rsRedemptionItems.Close
    Set rsRedemptionItems = Nothing
End If
%>

<!--#include file="../includes/footer.asp"-->
<%
Call CloseConnection()
%>
