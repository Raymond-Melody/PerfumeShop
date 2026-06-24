<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"

' 检查是否登录
If Session("UserID") = "" Then
    Response.Redirect "/user/login.asp"
End If
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/member_utils.asp"-->
<%
Call OpenConnection()

Dim userId, rsUser, userLevel, userPoints
userId = Session("UserID")

' 获取会员等级积分信息
userLevel = MU_CalcUserLevel(userId)
userPoints = MU_GetUserPoints(userId)

' 获取优惠券数量
Dim couponCount
couponCount = 0
Dim rsCoupon
Set rsCoupon = conn.Execute("SELECT COUNT(*) FROM SiteSettings WHERE SettingKey LIKE 'Coupon_%' AND SettingKey LIKE '%_User" & userId & "'")
If Not rsCoupon Is Nothing Then
    If Not rsCoupon.EOF Then couponCount = rsCoupon(0)
    rsCoupon.Close
End If
Set rsCoupon = Nothing

' 获取用户信息
Set rsUser = ExecuteQuery("SELECT * FROM Users WHERE UserID = " & userId)

' 获取订单统计
Dim orderCount, totalSpent
orderCount = GetScalar("SELECT COUNT(*) FROM Orders WHERE UserID = " & userId)
totalSpent = GetScalar("SELECT ISNULL(SUM(TotalAmount), 0) FROM Orders WHERE UserID = " & userId & " AND Status NOT IN ('Cancelled')")
If IsNull(orderCount) Then orderCount = 0
If IsNull(totalSpent) Then totalSpent = 0

' 获取最近订单
Dim rsOrders
Set rsOrders = ExecuteQuery("SELECT TOP 5 * FROM Orders WHERE UserID = " & userId & " ORDER BY CreatedAt DESC")

' V14: 获取推荐统计
Dim referralStats, canGenerate, dailyLimitMsg
Set referralStats = MU_GetReferralStats(userId)
canGenerate = MU_CheckDailyReferralLimit(userId)
If Not canGenerate Then
    dailyLimitMsg = "今日推荐链接生成次数已达上限（5个/天），请明天再试"
Else
    dailyLimitMsg = ""
End If

' V14: 处理生成推荐链接请求
Dim genSuccessMsg, genErrorMsg
genSuccessMsg = ""
genErrorMsg = ""
If Request.ServerVariables("REQUEST_METHOD") = "POST" And Request.Form("action") = "generate_link" Then
    If Not ValidateCSRFToken() Then
        genErrorMsg = "安全验证失败，请刷新页面重试"
    ElseIf Not canGenerate Then
        genErrorMsg = dailyLimitMsg
    Else
        Call MU_CheckDailyReferralLimit(userId) ' 再次确认
        If Not MU_CheckDailyReferralLimit(userId) Then
            genErrorMsg = "今日推荐链接生成次数已达上限"
        Else
            ' 生成Token并存储
            Dim newToken, genResult
            newToken = MU_GenerateReferralToken(userId, 30, 1, "user")
            If newToken <> "" Then
                If MU_StoreReferralToken(newToken, "user") Then
                    genSuccessMsg = "推荐链接已生成！请复制下方链接发送给好友"
                    ' 刷新统计
                    Set referralStats = MU_GetReferralStats(userId)
                    canGenerate = MU_CheckDailyReferralLimit(userId)
                Else
                    genErrorMsg = "存储推荐链接失败: " & Session("LastDBError")
                End If
            Else
                genErrorMsg = "生成推荐Token失败"
            End If
        End If
    End If
End If

' 确保CSRF令牌存在
Call EnsureCSRFToken()
%>
<!--#include file="../includes/header.asp"-->

<div class="breadcrumb">
    <div class="container">
        <a href="/index.asp">首页</a>
        <span class="separator">/</span>
        <span>个人中心</span>
    </div>
</div>

<div class="container">
    <div class="user-center">
        <!-- 侧边栏 -->
        <aside class="user-sidebar">
            <div class="user-profile">
                <h3><%= HTMLEncode(Session("Username")) %></h3>
                <p><%= HTMLEncode(Session("Email")) %></p>
                <div style="margin-top:8px;">
                    <% Call MU_RenderLevelBadge(userId) %>
                </div>
                <div style="margin-top:5px;font-size:12px;color:#888;">
                    积分: <strong><%= userPoints %></strong> | 
                    等级折扣: <strong><%= FormatPercent(1 - MU_GetLevelDiscount(userLevel), 0) %></strong>
                </div>
            </div>
            
            <nav class="user-nav">
                <a href="/user/index.asp" class="active"><i class="fas fa-home"></i> 个人中心</a>
                <% If Not rsUser Is Nothing Then %>
                    <% If rsUser("UserRole") = "KOL" Then %>
                        <a href="/user/kol_products.asp" style="background: #fff0f6; color: #eb2f96; font-weight: bold;"><i class="fas fa-star"></i> KOL推荐管理</a>
                    <% End If %>
                <% End If %>
                <a href="/user/orders.asp"><i class="fas fa-list"></i> 我的订单</a>
                <a href="/user/settings.asp"><i class="fas fa-user-edit"></i> 账户设置</a>
                <a href="/user/addresses.asp"><i class="fas fa-map-marker-alt"></i> 收货地址</a>
                <a href="/user/favorites.asp"><i class="fas fa-heart"></i> 我的收藏</a>
                <a href="/user/index.asp#referral"><i class="fas fa-user-friends"></i> 推荐好友</a>
                <a href="/user/logout.asp"><i class="fas fa-sign-out-alt"></i> 退出登录</a>
            </nav>
        </aside>

        <!-- 主内容 -->
        <div class="user-main">
            <div class="welcome-section">
                <h1>欢迎回来，<%= HTMLEncode(Session("Username")) %>！</h1>
                <p>在这里管理您的订单和账户信息</p>
            </div>
            
            <!-- 统计卡片 -->
            <div class="stats-grid">
                <div class="stat-card">
                    <div class="stat-icon"><i class="fas fa-shopping-bag"></i></div>
                    <div class="stat-info">
                        <span class="stat-value"><%= orderCount %></span>
                        <span class="stat-label">订单数量</span>
                    </div>
                </div>
                <div class="stat-card">
                    <div class="stat-icon"><i class="fas fa-yen-sign"></i></div>
                    <div class="stat-info">
                        <span class="stat-value"><%= FormatMoney(totalSpent) %></span>
                        <span class="stat-label">累计消费</span>
                    </div>
                </div>
                <div class="stat-card">
                    <div class="stat-icon"><i class="fas fa-coins"></i></div>
                    <div class="stat-info">
                        <span class="stat-value"><%= userPoints %></span>
                        <span class="stat-label">积分余额</span>
                    </div>
                </div>
                <div class="stat-card">
                    <div class="stat-icon"><i class="fas fa-ticket-alt"></i></div>
                    <div class="stat-info">
                        <span class="stat-value"><%= couponCount %></span>
                        <span class="stat-label">优惠券</span>
                    </div>
                </div>
            </div>
            
            <!-- 快捷入口 -->
            <div class="quick-actions">
                <h2>快捷操作</h2>
                <div class="action-grid">
                    <a href="/products.asp" class="action-item">
                        <i class="fas fa-spray-can"></i>
                        <span>选购香水</span>
                    </a>
                    <a href="/customize.asp" class="action-item">
                        <i class="fas fa-magic"></i>
                        <span>开始定制</span>
                    </a>
                    <a href="/cart.asp" class="action-item">
                        <i class="fas fa-shopping-cart"></i>
                        <span>我的购物车</span>
                    </a>
                    <a href="/user/orders.asp" class="action-item">
                        <i class="fas fa-truck"></i>
                        <span>查看物流</span>
                    </a>
                </div>
            </div>
            
            <!-- V14: 推荐好友 -->
            <div class="referral-section" id="referral">
                <div class="section-header">
                    <h2><i class="fas fa-user-friends"></i> 推荐好友</h2>
                    <span style="font-size:13px;color:#999;">邀请好友加入，一起享受香氛定制</span>
                </div>
                
                <% If genSuccessMsg <> "" Then %>
                <div class="alert alert-success" style="margin-bottom:15px;">
                    <i class="fas fa-check-circle"></i> <%= genSuccessMsg %>
                </div>
                <% End If %>
                <% If genErrorMsg <> "" Then %>
                <div class="alert alert-error" style="margin-bottom:15px;">
                    <i class="fas fa-exclamation-circle"></i> <%= genErrorMsg %>
                </div>
                <% End If %>
                
                <div class="referral-stats" style="display:grid;grid-template-columns:repeat(4,1fr);gap:15px;margin-bottom:20px;">
                    <div class="stat-mini" style="background:#f8f9fa;padding:15px;border-radius:8px;text-align:center;">
                        <div style="font-size:24px;font-weight:bold;color:#333;"><%= referralStats("todayLinks") %></div>
                        <div style="font-size:12px;color:#888;">今日生成</div>
                    </div>
                    <div class="stat-mini" style="background:#f8f9fa;padding:15px;border-radius:8px;text-align:center;">
                        <div style="font-size:24px;font-weight:bold;color:#333;"><%= referralStats("activeLinks") %></div>
                        <div style="font-size:12px;color:#888;">有效链接</div>
                    </div>
                    <div class="stat-mini" style="background:#f8f9fa;padding:15px;border-radius:8px;text-align:center;">
                        <div style="font-size:24px;font-weight:bold;color:#4CAF50;"><%= referralStats("totalInvitees") %></div>
                        <div style="font-size:12px;color:#888;">已邀请</div>
                    </div>
                    <div class="stat-mini" style="background:#f8f9fa;padding:15px;border-radius:8px;text-align:center;">
                        <div style="font-size:24px;font-weight:bold;color:#999;"><%= 5 - CLng(referralStats("todayLinks")) %></div>
                        <div style="font-size:12px;color:#888;">今日剩余</div>
                    </div>
                </div>
                
                <% If canGenerate Then %>
                <form method="post" style="margin-bottom:20px;" id="genLinkForm">
                    <%= GetCSRFTokenField() %>
                    <input type="hidden" name="action" value="generate_link">
                    <button type="submit" class="btn btn-primary" style="padding:12px 24px;">
                        <i class="fas fa-plus-circle"></i> 生成推荐链接
                    </button>
                    <span style="font-size:12px;color:#999;margin-left:10px;">有效期30天 · 每日限5个</span>
                </form>
                <% Else %>
                <div class="alert alert-warning" style="margin-bottom:20px;background:#fff3cd;color:#856404;padding:12px;border-radius:8px;">
                    <i class="fas fa-info-circle"></i> <%= dailyLimitMsg %>
                </div>
                <% End If %>
                
                <%
                ' 显示有效推荐链接列表
                Dim rsLinks
                Set rsLinks = MU_GetActiveReferralLinks(userId)
                If Not rsLinks Is Nothing And Not rsLinks.EOF Then
                    Dim linkIndex
                    linkIndex = 0
                %>
                <div class="referral-links-list">
                    <h4 style="margin-bottom:12px;"><i class="fas fa-link" style="color:#4CAF50;"></i> 有效推荐链接</h4>
                    <%
                    Do While Not rsLinks.EOF
                        Dim linkCreatedAt, linkExpiresAt, linkUsed, linkMax, linkTokenHash, linkOriginalToken, linkUrl
                        linkCreatedAt = rsLinks("CreatedAt")
                        linkExpiresAt = rsLinks("ExpiresAt")
                        linkUsed = rsLinks("UsedCount")
                        linkMax = rsLinks("MaxUses")
                        linkTokenHash = rsLinks("TokenHash")
                        linkOriginalToken = rsLinks("OriginalToken")
                        linkIndex = linkIndex + 1
                        If Not IsNull(linkOriginalToken) And linkOriginalToken <> "" Then
                            Dim linkProtocol
                            If Request.ServerVariables("HTTPS") = "on" Then
                                linkProtocol = "https://"
                            Else
                                linkProtocol = "http://"
                            End If
                            linkUrl = linkProtocol & Request.ServerVariables("HTTP_HOST") & "/user/register.asp?token=" & Server.URLEncode(linkOriginalToken)
                        Else
                            linkUrl = ""
                        End If
                    %>
                    <div style="background:#f8f9fa;border-radius:8px;padding:12px 15px;margin-bottom:10px;border:1px solid #eee;" class="ref-link-card">
                        <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;">
                            <div style="flex:1;min-width:200px;">
                                <div style="font-size:11px;color:#999;margin-bottom:4px;">
                                    生成: <%= SafeFormatDateTime(linkCreatedAt, 2) %> · 过期: <%= SafeFormatDateTime(linkExpiresAt, 2) %> · 已用 <%= linkUsed %>/<%= linkMax %>
                                </div>
                                <% If linkUrl <> "" Then %>
                                <div style="display:flex;align-items:center;gap:6px;">
                                    <input type="text" readonly value="<%= linkUrl %>" id="refLink<%= linkIndex %>" style="flex:1;padding:6px 10px;border:1px solid #ddd;border-radius:4px;font-size:12px;background:#fff;color:#333;font-family:monospace;" onclick="this.select()">
                                    <button type="button" class="btn btn-sm" onclick="copyRefLink(<%= linkIndex %>)" style="background:#4CAF50;color:#fff;white-space:nowrap;padding:6px 12px;">
                                        <i class="fas fa-copy"></i> 复制
                                    </button>
                                    <button type="button" class="btn btn-sm" onclick="toggleShareMenu(<%= linkIndex %>)" style="background:#2196F3;color:#fff;white-space:nowrap;padding:6px 12px;" id="shareBtn<%= linkIndex %>">
                                        <i class="fas fa-share-alt"></i> 分享
                                    </button>
                                </div>
                                <div id="shareMenu<%= linkIndex %>" style="display:none;margin-top:8px;padding:10px;background:#fff;border:1px solid #e0e0e0;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.1);">
                                    <div style="font-size:12px;color:#888;margin-bottom:8px;">分享到社交平台：</div>
                                    <div style="display:flex;gap:8px;flex-wrap:wrap;">
                                        <a href="#" id="shareWeibo<%= linkIndex %>" onclick="return shareWeibo(<%= linkIndex %>)" target="_blank" rel="noopener" style="display:inline-flex;align-items:center;gap:4px;padding:6px 12px;background:#E6162D;color:#fff;border-radius:6px;font-size:12px;text-decoration:none;" title="分享到微博">
                                            <i class="fab fa-weibo"></i> 微博
                                        </a>
                                        <a href="#" id="shareQQ<%= linkIndex %>" onclick="return shareQQ(<%= linkIndex %>)" target="_blank" rel="noopener" style="display:inline-flex;align-items:center;gap:4px;padding:6px 12px;background:#12B7F5;color:#fff;border-radius:6px;font-size:12px;text-decoration:none;" title="分享到QQ">
                                            <i class="fab fa-qq"></i> QQ
                                        </a>
                                        <a href="javascript:void(0)" onclick="shareWeChat(<%= linkIndex %>)" style="display:inline-flex;align-items:center;gap:4px;padding:6px 12px;background:#07C160;color:#fff;border-radius:6px;font-size:12px;text-decoration:none;" title="分享到微信">
                                            <i class="fab fa-weixin"></i> 微信
                                        </a>
                                    </div>
                                </div>
                                <% Else %>
                                <span style="color:#999;font-size:12px;">链接已失效</span>
                                <% End If %>
                            </div>
                        </div>
                    </div>
                    <%
                        rsLinks.MoveNext
                    Loop
                    %>
                </div>
                <script>
                function copyRefLink(idx) {
                    var input = document.getElementById('refLink' + idx);
                    if (!input) return;
                    input.select();
                    input.setSelectionRange(0, 99999);
                    if (navigator.clipboard && navigator.clipboard.writeText) {
                        navigator.clipboard.writeText(input.value).then(function() {
                            showToast('推荐链接已复制到剪贴板');
                        });
                    } else {
                        document.execCommand('copy');
                        showToast('推荐链接已复制');
                    }
                }
                function toggleShareMenu(idx) {
                    var menu = document.getElementById('shareMenu' + idx);
                    if (!menu) return;
                    // 关闭其他已打开的菜单
                    document.querySelectorAll('[id^="shareMenu"]').forEach(function(m) {
                        if (m.id !== 'shareMenu' + idx) m.style.display = 'none';
                    });
                    menu.style.display = menu.style.display === 'none' ? 'block' : 'none';
                }
                function getLinkUrl(idx) {
                    var input = document.getElementById('refLink' + idx);
                    return input ? input.value : '';
                }
                function shareWeibo(idx) {
                    var url = getLinkUrl(idx);
                    var title = encodeURIComponent('我在香氛定制平台发现了超棒的定制香水，快来加入我们吧！');
                    var shareUrl = 'https://service.weibo.com/share/share.php?url=' + encodeURIComponent(url) + '&title=' + title;
                    var link = document.getElementById('shareWeibo' + idx);
                    if (link) link.href = shareUrl;
                    document.getElementById('shareMenu' + idx).style.display = 'none';
                    return true;
                }
                function shareQQ(idx) {
                    var url = getLinkUrl(idx);
                    var title = encodeURIComponent('香氛定制 - 会员推荐注册');
                    var desc = encodeURIComponent('我在香氛定制平台发现了超棒的定制香水，快来加入我们吧！');
                    var shareUrl = 'https://connect.qq.com/widget/shareqq/index.html?url=' + encodeURIComponent(url) + '&title=' + title + '&desc=' + desc;
                    var link = document.getElementById('shareQQ' + idx);
                    if (link) link.href = shareUrl;
                    document.getElementById('shareMenu' + idx).style.display = 'none';
                    return true;
                }
                function shareWeChat(idx) {
                    var url = getLinkUrl(idx);
                    if (navigator.clipboard && navigator.clipboard.writeText) {
                        navigator.clipboard.writeText(url).then(function() {
                            showToast('链接已复制，请打开微信粘贴发送给好友');
                        });
                    } else {
                        var input = document.getElementById('refLink' + idx);
                        input.select();
                        document.execCommand('copy');
                        showToast('链接已复制，请打开微信粘贴发送给好友');
                    }
                    document.getElementById('shareMenu' + idx).style.display = 'none';
                }
                function showToast(msg) {
                    var t = document.getElementById('refToast');
                    if (!t) {
                        t = document.createElement('div');
                        t.id = 'refToast';
                        t.style.cssText = 'position:fixed;top:20px;left:50%;transform:translateX(-50%);background:#333;color:#fff;padding:10px 24px;border-radius:8px;z-index:9999;font-size:14px;transition:opacity 0.3s;';
                        document.body.appendChild(t);
                    }
                    t.textContent = msg;
                    t.style.opacity = '1';
                    setTimeout(function() { t.style.opacity = '0'; }, 2000);
                }
                // 点击页面其他地方关闭分享菜单
                document.addEventListener('click', function(e) {
                    if (!e.target.closest('[id^="shareMenu"]') && !e.target.closest('[id^="shareBtn"]')) {
                        document.querySelectorAll('[id^="shareMenu"]').forEach(function(m) { m.style.display = 'none'; });
                    }
                });
                </script>
                <%
                    rsLinks.Close
                Else
                %>
                <div style="text-align:center;padding:30px;color:#999;background:#f8f9fa;border-radius:8px;">
                    <i class="fas fa-link" style="font-size:32px;display:block;margin-bottom:10px;color:#ccc;"></i>
                    <p>暂无有效推荐链接，点击上方按钮生成</p>
                </div>
                <%
                End If
                Set rsLinks = Nothing
                %>
            </div>
            
            <!-- 最近订单 -->
            <div class="recent-orders">
                <div class="section-header">
                    <h2>最近订单</h2>
                    <a href="/user/orders.asp">查看全部 <i class="fas fa-arrow-right"></i></a>
                </div>
                
                <% If Not rsOrders Is Nothing And Not rsOrders.EOF Then %>
                <div class="orders-list">
                    <% Do While Not rsOrders.EOF %>
                    <div class="order-item">
                        <div class="order-info">
                            <span class="order-no">订单号: <%= rsOrders("OrderNo") %></span>
                            <span class="order-date"><%= SafeFormatDateTime(rsOrders("CreatedAt"), 2) %></span>
                        </div>
                        <div class="order-amount">
                            <%= FormatMoney(rsOrders("TotalAmount")) %>
                        </div>
                        <div class="order-status">
                            <%
                            Dim statusClass, statusText
                            Select Case rsOrders("Status")
                                Case "Pending"
                                    statusClass = "pending"
                                    statusText = "待付款"
                                Case "Paid"
                                    statusClass = "paid"
                                    statusText = "已付款"
                                Case "Processing"
                                    statusClass = "processing"
                                    statusText = "制作中"
                                Case "Shipped"
                                    statusClass = "shipped"
                                    statusText = "已发货"
                                Case "Delivered"
                                    statusClass = "delivered"
                                    statusText = "已完成"
                                Case "Cancelled"
                                    statusClass = "cancelled"
                                    statusText = "已取消"
                                Case Else
                                    statusClass = ""
                                    statusText = rsOrders("Status")
                            End Select
                            %>
                            <span class="status-badge <%= statusClass %>"><%= statusText %></span>
                        </div>
                        <div class="order-action">
                            <a href="/user/order_detail.asp?order_id=<%= rsOrders("OrderID") %>" class="btn btn-sm btn-outline">查看详情</a>
                        </div>
                    </div>
                    <%
                        rsOrders.MoveNext
                    Loop
                    %>
                </div>
                <%
                Else
                %>
                <div class="empty-orders">
                    <i class="fas fa-inbox"></i>
                    <p>暂无订单记录</p>
                    <a href="/products.asp" class="btn btn-primary">去选购</a>
                </div>
                <%
                End If
                %>
            </div>
        </div>
    </div>
</div>

<%
If Not rsOrders Is Nothing Then
    rsOrders.Close
    Set rsOrders = Nothing
End If
If Not rsUser Is Nothing Then
    rsUser.Close
    Set rsUser = Nothing
End If
%>

<!--#include file="../includes/footer.asp"-->
<%
Call CloseConnection()
%>
