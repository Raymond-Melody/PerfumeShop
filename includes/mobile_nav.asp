<%
' ============================================
' V11 移动端导航组件
' 包含：汉堡菜单 + 侧边栏 + 底部导航
' ============================================
%>
<!-- 移动端顶部导航栏 -->
<nav class="mobile-nav" id="mobileNav">
    <div class="mobile-nav-left">
        <button class="hamburger-btn" id="hamburgerBtn" aria-label="<% If FEATURE_I18N Then %><%= T("mobile_menu_open", Empty) %><% Else %>打开菜单<% End If %>">
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
        <button class="close-menu-btn" id="closeMenuBtn" aria-label="<% If FEATURE_I18N Then %><%= T("mobile_menu_close", Empty) %><% Else %>关闭菜单<% End If %>">
            <i class="fas fa-times"></i>
        </button>
    </div>
    <ul class="mobile-menu-list">
        <li><a href="/"><i class="fas fa-home"></i> <% If FEATURE_I18N Then %><%= T("mobile_menu_home", Empty) %><% Else %>首页<% End If %></a></li>
        <li><a href="/products.asp"><i class="fas fa-box"></i> <% If FEATURE_I18N Then %><%= T("mobile_menu_all", Empty) %><% Else %>所有产品<% End If %></a></li>
        <li><a href="/customize.asp"><i class="fas fa-magic"></i> <% If FEATURE_I18N Then %><%= T("mobile_menu_customize", Empty) %><% Else %>定制香水<% End If %></a></li>
        <% If FEATURE_FLASH_SALE Then %>
        <li><a href="/flash_sale.asp"><i class="fas fa-bolt"></i> <% If FEATURE_I18N Then %><%= T("mobile_menu_flash", Empty) %><% Else %>限时秒杀<% End If %></a></li>
        <% End If %>
        <% If FEATURE_GROUP_BUY Then %>
        <li><a href="/group_buy.asp"><i class="fas fa-users"></i> <% If FEATURE_I18N Then %><%= T("mobile_menu_group", Empty) %><% Else %>拼团优惠<% End If %></a></li>
        <% End If %>
        <% If FEATURE_COMMUNITY Then %>
        <li><a href="/community.asp"><i class="fas fa-comments"></i> <% If FEATURE_I18N Then %><%= T("mobile_menu_community", Empty) %><% Else %>香氛社区<% End If %></a></li>
        <% End If %>
        <% If FEATURE_SUBSCRIPTION Then %>
        <li><a href="/subscribe.asp"><i class="fas fa-box-open"></i> <% If FEATURE_I18N Then %><%= T("mobile_menu_subscription", Empty) %><% Else %>订阅盒子<% End If %></a></li>
        <% End If %>
        <li><a href="/about.asp"><i class="fas fa-info-circle"></i> <% If FEATURE_I18N Then %><%= T("mobile_menu_about", Empty) %><% Else %>关于我们<% End If %></a></li>
        <li><a href="/contact.asp"><i class="fas fa-phone"></i> <% If FEATURE_I18N Then %><%= T("mobile_menu_contact", Empty) %><% Else %>联系我们<% End If %></a></li>
        <li class="divider"></li>
        <% If Session("UserID") <> "" Then %>
            <li><a href="/user/index.asp"><i class="fas fa-user-circle"></i> <% If FEATURE_I18N Then %><%= T("mobile_menu_profile", Empty) %><% Else %>个人中心<% End If %></a></li>
            <li><a href="/user/orders.asp"><i class="fas fa-clipboard-list"></i> <% If FEATURE_I18N Then %><%= T("mobile_menu_orders", Empty) %><% Else %>我的订单<% End If %></a></li>
            <li><a href="/user/coupons.asp"><i class="fas fa-ticket-alt"></i> <% If FEATURE_I18N Then %><%= T("mobile_menu_coupons", Empty) %><% Else %>优惠券<% End If %></a></li>
            <li><a href="/user/favorites.asp"><i class="fas fa-heart"></i> <% If FEATURE_I18N Then %><%= T("mobile_menu_favorites", Empty) %><% Else %>收藏夹<% End If %></a></li>
            <li><a href="/user/settings.asp"><i class="fas fa-cog"></i> <% If FEATURE_I18N Then %><%= T("mobile_menu_settings", Empty) %><% Else %>账户设置<% End If %></a></li>
            <li><a href="/user/logout.asp"><i class="fas fa-sign-out-alt"></i> <% If FEATURE_I18N Then %><%= T("mobile_menu_logout", Empty) %><% Else %>退出登录<% End If %></a></li>
        <% Else %>
            <li><a href="/user/login.asp"><i class="fas fa-sign-in-alt"></i> <% If FEATURE_I18N Then %><%= T("mobile_menu_login", Empty) %><% Else %>登录<% End If %></a></li>
            <li><a href="/user/register.asp"><i class="fas fa-user-plus"></i> <% If FEATURE_I18N Then %><%= T("mobile_menu_register", Empty) %><% Else %>注册<% End If %></a></li>
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
        <span><% If FEATURE_I18N Then %><%= T("mobile_bottom_home", Empty) %><% Else %>首页<% End If %></span>
    </a>
    <a href="/products.asp" class="bottom-nav-item">
        <i class="fas fa-th-large"></i>
        <span><% If FEATURE_I18N Then %><%= T("mobile_bottom_category", Empty) %><% Else %>分类<% End If %></span>
    </a>
    <a href="/cart.asp" class="bottom-nav-item">
        <i class="fas fa-shopping-cart"></i>
        <span><% If FEATURE_I18N Then %><%= T("mobile_bottom_cart", Empty) %><% Else %>购物车<% End If %></span>
        <% If Session("CartCount") > 0 Then %>
        <span class="badge"><%= Session("CartCount") %></span>
        <% End If %>
    </a>
    <% If Session("UserID") <> "" Then %>
    <a href="/user/orders.asp" class="bottom-nav-item">
        <i class="fas fa-user"></i>
        <span><% If FEATURE_I18N Then %><%= T("mobile_bottom_orders", Empty) %><% Else %>订单<% End If %></span>
    </a>
    <% Else %>
    <a href="/user/login.asp" class="bottom-nav-item">
        <i class="fas fa-user"></i>
        <span><% If FEATURE_I18N Then %><%= T("mobile_bottom_login", Empty) %><% Else %>登录<% End If %></span>
    </a>
    <% End If %>
</nav>

<script>
// V11 移动端导航交互逻辑（含桌面窄屏汉堡菜单支持）
(function() {
    'use strict';
    
    var hamburgerBtn = document.getElementById('hamburgerBtn');
    var desktopHamburgerBtn = document.getElementById('desktopHamburgerBtn');
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
        if (desktopHamburgerBtn) desktopHamburgerBtn.classList.add('active');
        body.style.overflow = 'hidden';
    }
    
    function closeMenu() {
        if (!mobileMenu || !overlay) return;
        mobileMenu.classList.remove('active');
        overlay.classList.remove('active');
        if (hamburgerBtn) hamburgerBtn.classList.remove('active');
        if (desktopHamburgerBtn) desktopHamburgerBtn.classList.remove('active');
        body.style.overflow = '';
    }
    
    if (hamburgerBtn) hamburgerBtn.addEventListener('click', openMenu);
    if (desktopHamburgerBtn) desktopHamburgerBtn.addEventListener('click', openMenu);
    if (closeMenuBtn) closeMenuBtn.addEventListener('click', closeMenu);
    if (overlay) overlay.addEventListener('click', closeMenu);
    
    // ESC键关闭菜单
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape' && mobileMenu && mobileMenu.classList.contains('active')) {
            closeMenu();
        }
    });
    
    // 当前页面高亮 - 增强版路径匹配
    function setActiveNav() {
        var currentPath = window.location.pathname.toLowerCase().replace(/\/$/, '') || '/';
        var currentFull = (currentPath + window.location.search).toLowerCase();
        
        // 侧边菜单active高亮
        var sideLinks = document.querySelectorAll('.mobile-menu-list a');
        var bestMatch = null, bestScore = 0;
        sideLinks.forEach(function(link) {
            link.classList.remove('active');
            var href = link.getAttribute('href');
            if (!href) return;
            var hrefNorm = href.toLowerCase();
            // 精确匹配优先
            if (currentPath === hrefNorm || currentFull === hrefNorm) {
                link.classList.add('active');
                bestMatch = link; bestScore = 100;
            } else if (hrefNorm !== '/' && currentPath.indexOf(hrefNorm) === 0) {
                var score = hrefNorm.length;
                if (score > bestScore) { bestMatch = link; bestScore = score; }
            }
        });
        if (bestMatch && bestScore < 100 && bestScore > 0) {
            bestMatch.classList.add('active');
        }
        
        // 底部导航active高亮
        var bottomLinks = document.querySelectorAll('.bottom-nav-item');
        var bestBottom = null, bestScoreB = 0;
        bottomLinks.forEach(function(link) {
            link.classList.remove('active');
            var href = link.getAttribute('href');
            if (!href) return;
            var hrefNorm = href.toLowerCase();
            if (currentPath === hrefNorm) {
                link.classList.add('active');
                bestBottom = link; bestScoreB = 100;
            } else if (hrefNorm !== '/' && currentPath.indexOf(hrefNorm) === 0) {
                var score = hrefNorm.length;
                if (score > bestScoreB) { bestBottom = link; bestScoreB = score; }
            }
        });
        if (bestBottom && bestScoreB < 100 && bestScoreB > 0) {
            bestBottom.classList.add('active');
        }
    }
    
    // 初始执行
    setActiveNav();
    
    // 监听popstate（浏览器前进后退）
    window.addEventListener('popstate', setActiveNav);
})();
</script>
