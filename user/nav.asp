<div class="user-nav">
    <div class="user-info">
        <div class="user-name">
            <h4><%= HTMLEncode(Session("Username")) %></h4>
            <span class="user-level">普通会员</span>
        </div>
    </div>
    
    <ul class="nav-menu">
        <li><a href="/user/index.asp"><i class="fas fa-home"></i> 个人中心</a></li>
        <li><a href="/user/orders.asp"><i class="fas fa-file-invoice"></i> 我的订单</a></li>
        <li><a href="/user/settings.asp"><i class="fas fa-cog"></i> 账户设置</a></li>
        <li><a href="/user/addresses.asp"><i class="fas fa-map-marker-alt"></i> 收货地址</a></li>
        <li><a href="/user/favorites.asp"><i class="fas fa-heart"></i> 我的收藏</a></li>
        <li><a href="/user/logout.asp"><i class="fas fa-sign-out-alt"></i> 退出登录</a></li>
    </ul>
</div>