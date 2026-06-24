<%
' ============================================
' V11 移动端导航组件
' 包含：汉堡菜单 + 侧边栏 + 底部导航
' ============================================
%>
<!-- 移动端顶部导航栏 -->
<nav class="mobile-nav" id="mobileNav">
    <div class="mobile-nav-left">
        <button class="hamburger-btn" id="hamburgerBtn" aria-label="打开菜单">
            <span></span>
            <span></span>
            <span></span>
        </button>
    </div>
    <a href="/" class="mobile-logo"><%= SITE_NAME %></a>
    <div class="mobile-nav-right">
        <a href="/cart.asp" class="touch-target" style="position:relative">
            <i class="fas fa-shopping-cart"></i>
            <% If Session("CartCount") > 0 Then %>
            <span class="badge" style="position:absolute;top:2px;right:-4px;background:#dc3545;color:#fff;font-size:0.6rem;padding:1px 5px;border-radius:10px;min-width:16px;text-align:center"><%= Session("CartCount") %></span>
            <% End If %>
        </a>
    </div>
</nav>

<!-- 移动端侧边菜单遮罩 -->
<div class="mobile-menu-overlay" id="mobileMenuOverlay"></div>

<!-- 移动端侧边菜单 -->
<nav class="mobile-menu" id="mobileMenu">
    <div class="mobile-menu-header">
        <h3><%= SITE_NAME %></h3>
        <button class="close-menu-btn" id="closeMenuBtn" aria-label="关闭菜单">
            <i class="fas fa-times"></i>
        </button>
    </div>
    <ul class="mobile-menu-list">
        <li><a href="/"><i class="fas fa-home"></i> 首页</a></li>
        <li><a href="/products.asp"><i class="fas fa-box"></i> 所有产品</a></li>
        <li><a href="/customize.asp"><i class="fas fa-magic"></i> 定制香水</a></li>
        <li><a href="/about.asp"><i class="fas fa-info-circle"></i> 关于我们</a></li>
        <li><a href="/contact.asp"><i class="fas fa-phone"></i> 联系我们</a></li>
        <li class="divider"></li>
        <% If Session("UserID") <> "" Then %>
            <li><a href="/user/orders.asp"><i class="fas fa-clipboard-list"></i> 我的订单</a></li>
            <li><a href="/user/favorites.asp"><i class="fas fa-heart"></i> 收藏夹</a></li>
            <li><a href="/user/settings.asp"><i class="fas fa-cog"></i> 账户设置</a></li>
            <li><a href="/user/logout.asp"><i class="fas fa-sign-out-alt"></i> 退出登录</a></li>
        <% Else %>
            <li><a href="/user/login.asp"><i class="fas fa-sign-in-alt"></i> 登录</a></li>
            <li><a href="/user/register.asp"><i class="fas fa-user-plus"></i> 注册</a></li>
        <% End If %>
    </ul>
    <div class="mobile-menu-footer">
        &copy; <%= Year(Now()) %> <%= SITE_NAME %>
    </div>
</nav>

<!-- 移动端底部导航栏 -->
<nav class="bottom-nav">
    <a href="/" class="bottom-nav-item">
        <i class="fas fa-home"></i>
        <span>首页</span>
    </a>
    <a href="/products.asp" class="bottom-nav-item">
        <i class="fas fa-th-large"></i>
        <span>分类</span>
    </a>
    <a href="/cart.asp" class="bottom-nav-item">
        <i class="fas fa-shopping-cart"></i>
        <span>购物车</span>
        <% If Session("CartCount") > 0 Then %>
        <span class="badge"><%= Session("CartCount") %></span>
        <% End If %>
    </a>
    <% If Session("UserID") <> "" Then %>
    <a href="/user/orders.asp" class="bottom-nav-item">
        <i class="fas fa-user"></i>
        <span>订单</span>
    </a>
    <% Else %>
    <a href="/user/login.asp" class="bottom-nav-item">
        <i class="fas fa-user"></i>
        <span>登录</span>
    </a>
    <% End If %>
</nav>

<script>
// V11 移动端导航交互逻辑
(function() {
    'use strict';
    
    var hamburgerBtn = document.getElementById('hamburgerBtn');
    var closeMenuBtn = document.getElementById('closeMenuBtn');
    var mobileMenu = document.getElementById('mobileMenu');
    var overlay = document.getElementById('mobileMenuOverlay');
    var body = document.body;
    
    // 设置body class适配固定导航
    body.classList.add('has-mobile-nav', 'has-bottom-nav');
    
    function openMenu() {
        if (!mobileMenu || !overlay) return;
        mobileMenu.classList.add('active');
        overlay.classList.add('active');
        if (hamburgerBtn) hamburgerBtn.classList.add('active');
        body.style.overflow = 'hidden';
    }
    
    function closeMenu() {
        if (!mobileMenu || !overlay) return;
        mobileMenu.classList.remove('active');
        overlay.classList.remove('active');
        if (hamburgerBtn) hamburgerBtn.classList.remove('active');
        body.style.overflow = '';
    }
    
    if (hamburgerBtn) hamburgerBtn.addEventListener('click', openMenu);
    if (closeMenuBtn) closeMenuBtn.addEventListener('click', closeMenu);
    if (overlay) overlay.addEventListener('click', closeMenu);
    
    // ESC键关闭菜单
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape' && mobileMenu && mobileMenu.classList.contains('active')) {
            closeMenu();
        }
    });
    
    // 当前页面高亮
    var currentPath = window.location.pathname.toLowerCase();
    var navLinks = document.querySelectorAll('.mobile-menu-list a, .bottom-nav-item');
    navLinks.forEach(function(link) {
        var href = link.getAttribute('href');
        if (href && currentPath === href.toLowerCase()) {
            link.classList.add('active');
        } else if (href && href !== '/' && currentPath.startsWith(href.toLowerCase())) {
            link.classList.add('active');
        }
    });
})();
</script>
