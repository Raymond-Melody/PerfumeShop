<div class="user-nav">
    <div class="user-info">
        <div class="user-name">
            <h4><%= HTMLEncode(Session("Username")) %></h4>
            <span class="user-level"><% If FEATURE_I18N Then %><%= T("user_nav_user_level", Empty) %><% Else %>普通会员<% End If %></span>
        </div>
    </div>
    
    <ul class="nav-menu">
        <li><a href="/user/index.asp"><i class="fas fa-home"></i> <% If FEATURE_I18N Then %><%= T("user_nav_center", Empty) %><% Else %>个人中心<% End If %></a></li>
        <li><a href="/user/orders.asp"><i class="fas fa-file-invoice"></i> <% If FEATURE_I18N Then %><%= T("user_nav_orders", Empty) %><% Else %>我的订单<% End If %></a></li>
        <li><a href="/user/settings.asp"><i class="fas fa-cog"></i> <% If FEATURE_I18N Then %><%= T("user_nav_settings", Empty) %><% Else %>账户设置<% End If %></a></li>
        <li><a href="/user/addresses.asp"><i class="fas fa-map-marker-alt"></i> <% If FEATURE_I18N Then %><%= T("user_nav_addresses", Empty) %><% Else %>收货地址<% End If %></a></li>
        <li><a href="/user/favorites.asp"><i class="fas fa-heart"></i> <% If FEATURE_I18N Then %><%= T("user_nav_favorites", Empty) %><% Else %>我的收藏<% End If %></a></li>
        <% If FEATURE_SUBSCRIPTION Then %>
        <li><a href="/user/subscription.asp"><i class="fas fa-box-open"></i> <% If FEATURE_I18N Then %><%= T("user_nav_subscription", Empty) %><% Else %>我的订阅<% End If %></a></li>
        <% End If %>
        <% If FEATURE_COMMUNITY Then %>
        <li><a href="/user/my_reviews.asp"><i class="fas fa-star"></i> <% If FEATURE_I18N Then %>我的评价<% Else %>我的评价<% End If %></a></li>
        <% End If %>
        <li><a href="/user/logout.asp"><i class="fas fa-sign-out-alt"></i> <% If FEATURE_I18N Then %><%= T("user_nav_logout", Empty) %><% Else %>退出登录<% End If %></a></li>
    </ul>
</div>