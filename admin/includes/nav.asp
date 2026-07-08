<!-- V9.0 全局管理中心导航 - V18统一导航标准 -->
<link rel="stylesheet" href="/css/admin.css">
<link rel="stylesheet" href="/css/responsive.css?v=18.0">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
<div class="admin-dashboard">
    <nav class="admin-navbar">
        <div class="admin-nav-container">
            <div style="display:flex;align-items:center;">
                <!-- 移动端汉堡按钮 -->
                <button class="admin-hamburger" id="adminHamburger" aria-label="<% If FEATURE_I18N Then %><%= T("admin_nav_menu", Empty) %><% Else %>菜单<% End If %>">
                    <span></span>
                    <span></span>
                    <span></span>
                </button>
                <a href="portal.asp" class="admin-nav-brand">
                    <i class="fas fa-cubes"></i>
                    <span><% If FEATURE_I18N Then %><%= T("admin_nav_brand", Empty) %><% Else %>V9.0 管理中心<% End If %></span>
                </a>
            </div>
            <ul class="admin-nav-menu">
                <li><a href="javascript:void(0)" onclick="location.reload()"><i class="fas fa-sync-alt"></i> <% If FEATURE_I18N Then %><%= T("admin_nav_refresh", Empty) %><% Else %>刷新<% End If %></a></li>
                <li><a href="../index.asp" target="_blank"><i class="fas fa-store"></i> <% If FEATURE_I18N Then %><%= T("admin_nav_visit_site", Empty) %><% Else %>访问前台<% End If %></a></li>
                <li><a href="logout.asp"><i class="fas fa-sign-out-alt"></i> <% If FEATURE_I18N Then %><%= T("admin_nav_logout", Empty) %><% Else %>退出登录<% End If %></a></li>
            </ul>
        </div>
    </nav>
    
    <!-- 移动端侧边栏遮罩 -->
    <div class="sidebar-overlay" id="sidebarOverlay"></div>
    
    <aside class="sidebar" id="adminSidebar">
        <ul class="sidebar-menu">
            <li><a href="portal.asp"><i class="fas fa-th-large"></i> <span><% If FEATURE_I18N Then %><%= T("admin_nav_back", Empty) %><% Else %>返回入口<% End If %></span></a></li>
            <li class="sidebar-section-title"><% If FEATURE_I18N Then %><%= T("admin_nav_business", Empty) %><% Else %>业务模块<% End If %></li>
            <li><a href="operation/index.asp"><i class="fas fa-chart-line"></i> <span><% If FEATURE_I18N Then %><%= T("admin_nav_operation", Empty) %><% Else %>运营管理<% End If %></span></a></li>
            <li><a href="operation/tier_management.asp"><i class="fas fa-layer-group"></i> <span>会员等级</span></a></li>
            <li><a href="semifinished/index.asp"><i class="fas fa-vial"></i> <span><% If FEATURE_I18N Then %><%= T("admin_nav_semifinished", Empty) %><% Else %>半成品生产<% End If %></span></a></li>
            <li><a href="prodcenter/index.asp"><i class="fas fa-industry"></i> <span><% If FEATURE_I18N Then %><%= T("admin_nav_prodcenter", Empty) %><% Else %>产品生产<% End If %></span></a></li>
            <li><a href="logistics/index.asp"><i class="fas fa-truck"></i> <span><% If FEATURE_I18N Then %><%= T("admin_nav_logistics", Empty) %><% Else %>物流管理<% End If %></span></a></li>
            <li><a href="purchase/index.asp"><i class="fas fa-shopping-cart"></i> <span><% If FEATURE_I18N Then %><%= T("admin_nav_purchase", Empty) %><% Else %>采购管理<% End If %></span></a></li>
            <li><a href="finance/index.asp"><i class="fas fa-dollar-sign"></i> <span><% If FEATURE_I18N Then %><%= T("admin_nav_finance", Empty) %><% Else %>财务管理<% End If %></span></a></li>
            <li><a href="techcenter/index.asp"><i class="fas fa-flask"></i> <span><% If FEATURE_I18N Then %><%= T("admin_nav_tech", Empty) %><% Else %>技术中心<% End If %></span></a></li>
            <li class="sidebar-section-title"><% If FEATURE_I18N Then %><%= T("admin_nav_sys_mgmt", Empty) %><% Else %>系统管理<% End If %></li>
            <li><a href="system/index.asp"><i class="fas fa-shield-alt"></i> <span><% If FEATURE_I18N Then %><%= T("admin_nav_system", Empty) %><% Else %>系统管理<% End If %></span></a></li>
            <li><a href="site_settings.asp"><i class="fas fa-cog"></i> <span><% If FEATURE_I18N Then %><%= T("admin_nav_site_settings", Empty) %><% Else %>站点设置<% End If %></span></a></li>
        </ul>
    </aside>
</div>

<!-- V18 管理后台导航交互脚本 -->
<script>
(function() {
    'use strict';
    
    var hamburger = document.getElementById('adminHamburger');
    var sidebar = document.getElementById('adminSidebar');
    var overlay = document.getElementById('sidebarOverlay');
    var body = document.body;
    
    // 移动端侧边栏切换
    function openSidebar() {
        if (!sidebar || !overlay) return;
        sidebar.classList.add('active');
        overlay.classList.add('active');
        if (hamburger) hamburger.classList.add('active');
        body.style.overflow = 'hidden';
    }
    
    function closeSidebar() {
        if (!sidebar || !overlay) return;
        sidebar.classList.remove('active');
        overlay.classList.remove('active');
        if (hamburger) hamburger.classList.remove('active');
        body.style.overflow = '';
    }
    
    if (hamburger) hamburger.addEventListener('click', function(e) {
        e.stopPropagation();
        if (sidebar && sidebar.classList.contains('active')) {
            closeSidebar();
        } else {
            openSidebar();
        }
    });
    if (overlay) overlay.addEventListener('click', closeSidebar);
    
    // 点击侧边栏链接后自动关闭
    if (sidebar) {
        sidebar.addEventListener('click', function(e) {
            if (e.target.tagName === 'A' || e.target.closest('a')) {
                setTimeout(closeSidebar, 150);
            }
        });
    }
    
    // ESC关闭
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape' && sidebar && sidebar.classList.contains('active')) {
            closeSidebar();
        }
    });
    
    // 窗口大小变化时重置
    window.addEventListener('resize', function() {
        if (window.innerWidth >= 769) {
            closeSidebar();
        }
    });
    
    // 当前页面高亮 - 侧边栏active检测（仅在服务器未设置active时执行）
    (function setActiveLink() {
        // 检查服务器端是否已设置active状态，避免冲突
        var hasServerActive = document.querySelector('.sidebar-menu a.active');
        if (hasServerActive) return;
        
        var currentPath = window.location.pathname.toLowerCase().replace(/\/$/, '') || '/';
        var links = document.querySelectorAll('.sidebar-menu a');
        var bestMatch = null, bestScore = 0;
        
        links.forEach(function(link) {
            link.classList.remove('active');
            var href = link.getAttribute('href');
            if (!href || href === '#' || href.indexOf('javascript:') === 0) return;
            
            var hrefNorm = href.toLowerCase().replace(/\/$/, '');
            var hrefPath = hrefNorm.split('?')[0];
            
            if (currentPath === hrefPath || currentPath === hrefNorm) {
                link.classList.add('active');
                bestMatch = link; bestScore = 100;
            } else if (hrefPath !== '/' && hrefPath !== '' && currentPath.indexOf(hrefPath) === 0) {
                var score = hrefPath.length;
                if (score > bestScore) { bestMatch = link; bestScore = score; }
            }
        });
        
        if (bestMatch && bestScore < 100 && bestScore > 0) {
            bestMatch.classList.add('active');
        }
    })();
})();
</script>
