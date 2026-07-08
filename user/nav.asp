<%
' ============================================
' 统一用户中心侧边栏导航 (V18)
' 所有用户中心页面通过 include 引用此文件
' 确保导航项、图标、顺序、激活状态完全一致
' ============================================

' 判断当前页面名称，用于激活状态
Dim navCurrentPage
navCurrentPage = LCase(Request.ServerVariables("SCRIPT_NAME"))
navCurrentPage = Mid(navCurrentPage, InStrRev(navCurrentPage, "/") + 1)

' KOL 角色检测
Dim navIsKOL
navIsKOL = False
If Session("UserID") <> "" Then
    On Error Resume Next
    Dim navKolRs
    Set navKolRs = conn.Execute("SELECT UserRole FROM Users WHERE UserID = " & Session("UserID"))
    If Not navKolRs Is Nothing Then
        If Not navKolRs.EOF Then
            If navKolRs("UserRole") = "KOL" Then navIsKOL = True
        End If
        navKolRs.Close
    End If
    Set navKolRs = Nothing
    On Error GoTo 0
End If
%>

<aside class="user-sidebar">
    <div class="user-profile">
        <h3><%= HTMLEncode(Session("Username")) %></h3>
        <p><%= HTMLEncode(Session("Email")) %></p>
    </div>
    
    <nav class="user-nav">
        <a href="/user/index.asp" data-nav="index" class="<% If navCurrentPage = "index.asp" Then %>active<% End If %>">
            <i class="fas fa-home"></i> <% If FEATURE_I18N Then %><%= T("user_nav_center", Empty) %><% Else %>个人中心<% End If %>
        </a>
        
        <% If FEATURE_POINTS_SYSTEM Then %>
        <a href="/user/points.asp" data-nav="points" class="<% If navCurrentPage = "points.asp" Then %>active<% End If %> nav-points" style="color:#ff8f00;">
            <i class="fas fa-coins"></i> <% If FEATURE_I18N Then %>积分中心<% Else %>积分中心<% End If %>
        </a>
        <% End If %>
        
        <% If FEATURE_COUPON_SYSTEM Then %>
        <a href="/user/coupons.asp" data-nav="coupons" class="<% If navCurrentPage = "coupons.asp" Then %>active<% End If %> nav-coupons" style="color:#2e7d32;">
            <i class="fas fa-ticket-alt"></i> <% If FEATURE_I18N Then %>优惠券<% Else %>优惠券<% End If %>
        </a>
        <% End If %>
        
        <% If navIsKOL Then %>
        <a href="/user/kol_products.asp" data-nav="kol" class="<% If navCurrentPage = "kol_products.asp" Then %>active<% End If %> nav-kol" style="color:#eb2f96;font-weight:bold;">
            <i class="fas fa-star"></i> <% If FEATURE_I18N Then %><%= T("user_nav_kol", Empty) %><% Else %>KOL推荐管理<% End If %>
        </a>
        <% End If %>
        
        <a href="/user/orders.asp" data-nav="orders" class="<% If navCurrentPage = "orders.asp" Then %>active<% End If %>">
            <i class="fas fa-list"></i> <% If FEATURE_I18N Then %><%= T("user_nav_orders", Empty) %><% Else %>我的订单<% End If %>
        </a>
        
        <a href="/user/settings.asp" data-nav="settings" class="<% If navCurrentPage = "settings.asp" Then %>active<% End If %>">
            <i class="fas fa-user-edit"></i> <% If FEATURE_I18N Then %><%= T("user_nav_settings", Empty) %><% Else %>账户设置<% End If %>
        </a>
        
        <a href="/user/addresses.asp" data-nav="addresses" class="<% If navCurrentPage = "addresses.asp" Then %>active<% End If %>">
            <i class="fas fa-map-marker-alt"></i> <% If FEATURE_I18N Then %><%= T("user_nav_addresses", Empty) %><% Else %>收货地址<% End If %>
        </a>
        
        <a href="/user/favorites.asp" data-nav="favorites" class="<% If navCurrentPage = "favorites.asp" Then %>active<% End If %>">
            <i class="fas fa-heart"></i> <% If FEATURE_I18N Then %><%= T("user_nav_favorites", Empty) %><% Else %>我的收藏<% End If %>
        </a>
        
        <% If FEATURE_SUBSCRIPTION Then %>
        <a href="/user/subscription.asp" data-nav="subscription" class="<% If navCurrentPage = "subscription.asp" Then %>active<% End If %>">
            <i class="fas fa-box-open"></i> <% If FEATURE_I18N Then %><%= T("user_nav_subscription", Empty) %><% Else %>我的订阅<% End If %>
        </a>
        <% End If %>
        
        <% If FEATURE_COMMUNITY Then %>
        <a href="/user/my_reviews.asp" data-nav="reviews" class="<% If navCurrentPage = "my_reviews.asp" Then %>active<% End If %>">
            <i class="fas fa-comment-dots"></i> <% If FEATURE_I18N Then %><%= T("user_nav_reviews", Empty) %><% Else %>我的评价<% End If %>
        </a>
        <% End If %>
        
        <a href="/user/index.asp#referral" data-nav="referral">
            <i class="fas fa-user-friends"></i> <% If FEATURE_I18N Then %><%= T("user_nav_referral", Empty) %><% Else %>推荐好友<% End If %>
        </a>
        
        <a href="/user/logout.asp" class="nav-logout">
            <i class="fas fa-sign-out-alt"></i> <% If FEATURE_I18N Then %><%= T("user_nav_logout", Empty) %><% Else %>退出登录<% End If %>
        </a>
    </nav>
</aside>

<%
' 为 order_detail.asp 特殊处理：激活"我的订单"
If navCurrentPage = "order_detail.asp" Then
%>
<script>
(function(){
    var ordersLink = document.querySelector('.user-nav a[data-nav="orders"]');
    if (ordersLink) ordersLink.classList.add('active');
})();
</script>
<% End If %>