<%
' ============================================
' 公共头部文件
' ============================================

' 确保CSRF令牌存在
Call EnsureCSRFToken()

' V14.6: 安全初始化 amphtmlLink（仅 product.asp 会设置此变量）
If IsEmpty(amphtmlLink) Or IsNull(amphtmlLink) Then amphtmlLink = ""
%>
<!DOCTYPE html>
<html lang="<%= I18N_HtmlLang() %>">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- V17.0 资源预加载 -->
    <link rel="dns-prefetch" href="//cdnjs.cloudflare.com">
    <link rel="dns-prefetch" href="//code.jquery.com">
    <link rel="preconnect" href="https://cdnjs.cloudflare.com" crossorigin>
    <link rel="preconnect" href="https://code.jquery.com" crossorigin>
    <link rel="preload" href="/css/design-tokens.css?v=17.0" as="style">
    <link rel="preload" href="/css/style.css?v=17.0" as="style">
    <link rel="preload" href="https://code.jquery.com/jquery-3.6.0.min.js" as="script" crossorigin>
    <!-- V14.6 PWA 增强 -->
    <meta name="color-scheme" content="light only">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="default">
    <meta name="apple-mobile-web-app-title" content="<% If FEATURE_I18N Then %><%= T("header_meta_title", Empty) %><% Else %>香氛定制<% End If %>">
    <link rel="apple-touch-icon" href="/images/icons/icon-192x192.png">
    <meta name="description" content="<% If FEATURE_I18N Then %><%= T("header_meta_description", Empty) %><% Else %>专业个性化定制香水，打造专属于你的独特香氛体验<% End If %>">
    <meta name="keywords" content="<% If FEATURE_I18N Then %><%= T("header_meta_keywords", Empty) %><% Else %>香水定制,个性化香水,定制香氛,专属香水<% End If %>">
    <title><%= SITE_NAME %></title>
    <link rel="icon" href="/favicon.ico" type="image/x-icon">
    <link rel="manifest" href="/manifest.json">
    <meta name="theme-color" content="#8B4513">
    <link rel="stylesheet" href="/css/design-tokens.css?v=17.0">
    <link rel="stylesheet" href="/css/style.css?v=17.0">
    <link rel="stylesheet" href="/css/pages.css?v=17.0">
    <link rel="stylesheet" href="/css/buttons.css?v=17.0">
    <link rel="stylesheet" href="/css/responsive.css?v=18.0">
    <link rel="stylesheet" href="/css/mobile-first.css?v=18.0" media="print" onload="this.media='all'">
    <link rel="stylesheet" href="/css/lazy-load.css?v=17.0" media="print" onload="this.media='all'">
    <link rel="stylesheet" href="/css/cart-animation.css?v=17.0" media="print" onload="this.media='all'">
    <link rel="stylesheet" href="/css/filter-optimization.css?v=17.0">
    <link rel="stylesheet" href="/css/theme.css?v=17.0">
    <link rel="stylesheet" href="/css/skeleton.css?v=17.0" media="print" onload="this.media='all'">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <% If amphtmlLink <> "" Then Response.Write amphtmlLink %>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="/js/mobile-gestures.js?v=18.0" defer></script>
    <!-- V17.0 搜索建议CSS -->
    <style nonce="<%= Session("csp_nonce") %>">
    .search-box { position: relative; }
    .search-suggestions { position: absolute; top: 100%; left: 0; right: 0; background: #fff; border: 1px solid #ddd; border-radius: 0 0 8px 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); z-index: 1000; display: none; max-height: 250px; overflow-y: auto; }
    .search-suggestions .s-item { padding: 10px 15px; cursor: pointer; color: #333; font-size: 14px; border-bottom: 1px solid #f0f0f0; transition: background 0.15s; }
    .search-suggestions .s-item:last-child { border-bottom: none; }
    .search-suggestions .s-item:hover, .search-suggestions .s-item.active { background: #f5f5f5; color: #8B4513; }
    .search-suggestions .s-item i { margin-right: 8px; color: #999; font-size: 12px; }
    .search-suggestions .s-footer { padding: 8px 15px; text-align: center; font-size: 12px; color: #999; background: #fafafa; border-top: 1px solid #eee; }
    .lang-switcher { margin-left: 15px; font-size: 13px; }
    .lang-link { color: rgba(255,255,255,0.8); text-decoration: none; padding: 2px 4px; transition: color 0.2s; }
    .lang-link:hover { color: #fff; }
    .lang-link.active { color: #fff; font-weight: bold; }
    .lang-divider { color: rgba(255,255,255,0.4); margin: 0 3px; }
    /* V18: Cookie 同意横幅 */
    .cookie-banner { position: fixed; bottom: 0; left: 0; right: 0; background: #2d3748; color: #fff; z-index: 9999; box-shadow: 0 -4px 20px rgba(0,0,0,0.15); animation: cookieSlideUp 0.5s ease; }
    @keyframes cookieSlideUp { from { transform: translateY(100%); } to { transform: translateY(0); } }
    .cookie-banner-inner { max-width: 1200px; margin: 0 auto; padding: 16px 24px; display: flex; align-items: center; gap: 20px; flex-wrap: wrap; }
    .cookie-icon { font-size: 1.5rem; opacity: 0.8; flex-shrink: 0; }
    .cookie-text { flex: 1; min-width: 250px; }
    .cookie-text p { margin: 0; font-size: 0.9rem; line-height: 1.5; opacity: 0.9; }
    .cookie-actions { display: flex; align-items: center; gap: 10px; flex-shrink: 0; flex-wrap: wrap; }
    .cookie-actions .btn-sm { padding: 8px 16px; font-size: 0.85rem; }
    .cookie-actions .btn-outline { background: transparent; border: 1px solid rgba(255,255,255,0.5); color: #fff; }
    .cookie-actions .btn-outline:hover { background: rgba(255,255,255,0.1); border-color: #fff; }
    .cookie-link { color: rgba(255,255,255,0.7); font-size: 0.85rem; text-decoration: underline; white-space: nowrap; }
    .cookie-link:hover { color: #fff; }
    @media (max-width: 640px) { .cookie-banner-inner { flex-direction: column; align-items: flex-start; gap: 12px; } .cookie-actions { width: 100%; } }
    </style>
    <script nonce="<%= Session("csp_nonce") %>">
    // V17.0 搜索建议功能
    var searchTimeout, searchSelectedIdx = -1;
    var searchSuggestionsFmt = '<% If FEATURE_I18N Then %><%= T("header_search_suggestions", Empty) %><% Else %>共 {0} 条建议<% End If %>';
    $(document).ready(function() {
        $('#searchKeyword').on('input', function() {
            var q = $(this).val().trim();
            clearTimeout(searchTimeout);
            searchSelectedIdx = -1;
            if (q.length < 2) { $('#searchSuggestions').hide(); return; }
            searchTimeout = setTimeout(function() {
                $.get('/api/search_suggestions.asp', { q: q, limit: 8 }, function(res) {
                    var html = '';
                    if (res && res.code === 0 && res.data && res.data.length > 0) {
                        for (var i = 0; i < res.data.length; i++) {
                            html += '<div class="s-item" data-value="' + $('<span>').text(res.data[i]).html() + '"><i class="fas fa-search"></i>' + $('<span>').text(res.data[i]).html() + '</div>';
                        }
                        html += '<div class="s-footer">' + searchSuggestionsFmt.replace('{0}', res.data.length) + '</div>';
                        $('#searchSuggestions').html(html).show();
                    } else {
                        $('#searchSuggestions').hide();
                    }
                });
            }, 300);
        }).on('keydown', function(e) {
            var items = $('#searchSuggestions .s-item');
            if (items.length === 0) return;
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                searchSelectedIdx = Math.min(searchSelectedIdx + 1, items.length - 1);
                items.removeClass('active');
                $(items[searchSelectedIdx]).addClass('active');
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                searchSelectedIdx = Math.max(searchSelectedIdx - 1, -1);
                items.removeClass('active');
                if (searchSelectedIdx >= 0) $(items[searchSelectedIdx]).addClass('active');
            } else if (e.key === 'Enter' && searchSelectedIdx >= 0) {
                e.preventDefault();
                $(items[searchSelectedIdx]).click();
            }
        });
        $(document).on('click', function(e) {
            if (!$(e.target).closest('#searchBox').length) $('#searchSuggestions').hide();
        });
        
        // V17.0: 点击搜索建议项
        $(document).on('click', '#searchSuggestions .s-item', function() {
            var val = $(this).data('value');
            if (val) {
                $('#searchKeyword').val(val);
                $('#searchKeyword').closest('form').submit();
            }
        });
    });
    </script>
    <script nonce="<%= Session("csp_nonce") %>">
    // 全局CSRF令牌配置
    var csrfToken = '<%= Session("CSRFToken") %>';
    
    // 为所有jQuery AJAX请求添加CSRF令牌
    $(document).ready(function() {
        $.ajaxSetup({
            beforeSend: function(xhr, settings) {
                // 对于POST请求添加CSRF头
                if (settings.type && settings.type.toUpperCase() === 'POST') {
                    xhr.setRequestHeader('X-CSRF-TOKEN', csrfToken);
                }
            }
        });
        
        // 拦截所有jQuery AJAX请求，确保CSRF令牌被添加到POST数据中
        var originalAjax = $.ajax;
        $.ajax = function(options) {
            // 如果是POST请求，确保csrf_token在data中
            if (options.type && options.type.toUpperCase() === 'POST') {
                // 处理data为对象的情况
                if (options.data && typeof options.data === 'object' && !(options.data instanceof FormData)) {
                    if (!(options.data instanceof Array)) {
                        options.data.csrf_token = csrfToken;
                    }
                }
                // 处理data为字符串的情况（如序列化的表单数据）
                else if (options.data && typeof options.data === 'string') {
                    if (options.data.indexOf('csrf_token=') === -1) {
                        options.data += '&csrf_token=' + encodeURIComponent(csrfToken);
                    }
                }
                // data为空的情况
                else if (!options.data) {
                    options.data = 'csrf_token=' + encodeURIComponent(csrfToken);
                }
            }
            return originalAjax.apply(this, arguments);
        };
    });
    </script>
</head>
<body>
    <%
    ' V18: GDPR Cookie 同意横幅
    If FEATURE_GDPR_COMPLIANCE Then
        Dim cookieConsent
        cookieConsent = Request.Cookies("cookie_consent")
        If cookieConsent = "" Then
    %>
    <div id="cookieBanner" class="cookie-banner">
        <div class="cookie-banner-inner">
            <div class="cookie-icon"><i class="fas fa-cookie-bite"></i></div>
            <div class="cookie-text">
                <p><% If FEATURE_I18N Then %><%= T("cookie_banner_text", Empty) %><% Else %>本网站使用 Cookie 来提升您的浏览体验、分析网站流量并提供个性化推荐。继续使用即表示您同意我们的 Cookie 政策。<% End If %></p>
            </div>
            <div class="cookie-actions">
                <button class="btn btn-primary btn-sm" onclick="acceptCookies('all')">
                    <% If FEATURE_I18N Then %><%= T("cookie_accept_all", Empty) %><% Else %>全部接受<% End If %>
                </button>
                <button class="btn btn-outline btn-sm" onclick="acceptCookies('essential')">
                    <% If FEATURE_I18N Then %><%= T("cookie_accept_essential", Empty) %><% Else %>仅必要<% End If %>
                </button>
                <a href="/user/privacy.asp" class="cookie-link">
                    <% If FEATURE_I18N Then %><%= T("cookie_learn_more", Empty) %><% Else %>了解更多<% End If %>
                </a>
            </div>
        </div>
    </div>
    <%
        End If
    End If
    %>
    <!-- V11 移动端导航组件 -->
    <!--#include file="mobile_nav.asp"-->
    
    <!-- 顶部公告栏 -->
    <div class="top-bar">
        <div class="container">
            <span><i class="fas fa-gift"></i> <% If FEATURE_I18N Then Response.Write T("site_slogan", Empty) Else Response.Write "专业个性化定制香水" End If %></span>
            <span class="top-bar-right">
                <i class="fas fa-phone"></i> <%= SITE_PHONE %>
                <% If FEATURE_I18N Then %>
                <span class="lang-switcher">
                    <a href="?lang=zh-CN" class="lang-link<%= IIF(I18N_GetLocale()="zh-CN"," active","") %>"><% If FEATURE_I18N Then %><%= T("header_lang_zh", Empty) %><% Else %>中文<% End If %></a>
                    <span class="lang-divider">|</span>
                    <a href="?lang=en-US" class="lang-link<%= IIF(I18N_GetLocale()="en-US"," active","") %>"><% If FEATURE_I18N Then %><%= T("header_lang_en", Empty) %><% Else %>EN<% End If %></a>
                </span>
                <% End If %>
            </span>
        </div>
    </div>

    <!-- 主导航 -->
    <header class="main-header">
        <div class="container">
            <div class="header-wrapper">
                <!-- Logo -->
                <div class="logo">
                    <a href="/index.asp">
                        <i class="fas fa-spray-can"></i>
                        <span><% If FEATURE_I18N Then Response.Write T("site_name", Empty) Else Response.Write "香氛定制" End If %></span>
                    </a>
                </div>

                <!-- 搜索框 -->
                <div class="search-box" id="searchBox">
                    <form action="/products.asp" method="get" autocomplete="off">
                        <input type="text" name="keyword" id="searchKeyword" placeholder="<% If FEATURE_I18N Then Response.Write T("header_search_placeholder", Empty) Else Response.Write "搜索您想要的香水..." End If %>" autocomplete="off">
                        <button type="submit"><i class="fas fa-search"></i></button>
                        <div class="search-suggestions" id="searchSuggestions"></div>
                    </form>
                </div>

                <!-- V18 桌面窄屏汉堡菜单按钮（992-1050px显示） -->
                <button class="hamburger-btn desktop-hamburger" id="desktopHamburgerBtn" aria-label="<% If FEATURE_I18N Then %><%= T("mobile_menu_open", Empty) %><% Else %>打开菜单<% End If %>">
                    <span></span>
                    <span></span>
                    <span></span>
                </button>

                <!-- 用户菜单 -->
                <div class="user-menu">
                    <%
                    If Session("UserID") <> "" Then
                    %>
                    <div class="user-info dropdown">
                        <a href="#"><i class="fas fa-user"></i> <%= Session("Username") %></a>
                        <div class="dropdown-menu">
                            <a href="/user/index.asp"><i class="fas fa-user-circle"></i> <% If FEATURE_I18N Then %><%= T("header_user_center", Empty) %><% Else %>个人中心<% End If %></a>
                            <a href="/user/orders.asp"><i class="fas fa-list"></i> <% If FEATURE_I18N Then %><%= T("header_my_orders", Empty) %><% Else %>我的订单<% End If %></a>
                            <a href="/user/settings.asp"><i class="fas fa-cog"></i> <% If FEATURE_I18N Then %><%= T("header_account_settings", Empty) %><% Else %>账户设置<% End If %></a>
                            <a href="/user/addresses.asp"><i class="fas fa-map-marker-alt"></i> <% If FEATURE_I18N Then %><%= T("header_shipping_address", Empty) %><% Else %>收货地址<% End If %></a>
                            <a href="/user/favorites.asp"><i class="fas fa-heart"></i> <% If FEATURE_I18N Then %><%= T("header_my_favorites", Empty) %><% Else %>我的收藏<% End If %></a>
                            <a href="/user/logout.asp"><i class="fas fa-sign-out-alt"></i> <% If FEATURE_I18N Then %><%= T("header_logout", Empty) %><% Else %>退出登录<% End If %></a>
                        </div>
                    </div>
                    <%
                    Else
                    %>
                    <a href="/user/login.asp" class="btn-login"><i class="fas fa-sign-in-alt"></i> <% If FEATURE_I18N Then Response.Write T("login", Empty) Else Response.Write "登录" End If %></a>
                    <a href="/user/register.asp" class="btn-register"><% If FEATURE_I18N Then Response.Write T("register", Empty) Else Response.Write "注册" End If %></a>
                    <%
                    End If
                    %>
                    <a href="/cart.asp" class="cart-icon">
                        <i class="fas fa-shopping-cart"></i>
                        <span class="cart-count" id="cartCount">0</span>
                    </a>
                </div>
            </div>
        </div>
    </header>

    <!-- 分类导航 -->
    <nav class="main-nav pc-only">
        <div class="container">
            <ul class="nav-list">
                <li><a href="/index.asp"><i class="fas fa-home"></i> <% If FEATURE_I18N Then Response.Write T("home", Empty) Else Response.Write "首页" End If %></a></li>
                <li><a href="/products.asp"><i class="fas fa-flask"></i> <% If FEATURE_I18N Then Response.Write T("products", Empty) Else Response.Write "全部香水" End If %></a></li>
                <li><a href="/customize.asp"><i class="fas fa-magic"></i> <% If FEATURE_I18N Then Response.Write T("customize", Empty) Else Response.Write "定制香水" End If %></a></li>
                <li><a href="/about.asp"><i class="fas fa-book"></i> <% If FEATURE_I18N Then Response.Write T("about", Empty) Else Response.Write "品牌故事" End If %></a></li>
                <li><a href="/contact.asp"><i class="fas fa-envelope"></i> <% If FEATURE_I18N Then Response.Write T("contact", Empty) Else Response.Write "联系我们" End If %></a></li>
                <% If FEATURE_FLASH_SALE Then %>
                <li><a href="/flash_sale.asp" class="nav-hot"><i class="fas fa-bolt"></i> <% If FEATURE_I18N Then %><%= T("nav_flash_sale", Empty) %><% Else %>秒杀<% End If %></a></li>
                <% End If %>
                <% If FEATURE_GROUP_BUY Then %>
                <li><a href="/group_buy.asp" class="nav-hot"><i class="fas fa-users"></i> <% If FEATURE_I18N Then %><%= T("nav_group_buy", Empty) %><% Else %>拼团<% End If %></a></li>
                <% End If %>
                <% If FEATURE_SUBSCRIPTION Then %>
                <li><a href="/subscribe.asp" class="nav-hot"><i class="fas fa-box-open"></i> <% If FEATURE_I18N Then %><%= T("nav_subscription", Empty) %><% Else %>订阅<% End If %></a></li>
                <% End If %>
                <% If FEATURE_COMMUNITY Then %>
                <li><a href="/community.asp" class="nav-hot"><i class="fas fa-comments"></i> <% If FEATURE_I18N Then %><%= T("nav_community", Empty) %><% Else %>社区<% End If %></a></li>
                <% End If %>
            </ul>
        </div>
    </nav>

    <!-- V18 桌面导航active状态检测 -->
    <script nonce="<%= Session("csp_nonce") %>">
    (function() {
        var currentPath = window.location.pathname.toLowerCase().replace(/\/$/, '') || '/';
        var navAnchors = document.querySelectorAll('.main-nav .nav-list > li > a');
        var bestMatch = null, bestScore = 0;
        navAnchors.forEach(function(a) {
            var href = a.getAttribute('href');
            if (!href) return;
            var hrefNorm = href.toLowerCase().replace(/\/$/, '');
            // 跳过带查询参数的链接的精确比较
            var hrefPath = hrefNorm.split('?')[0];
            if (currentPath === hrefPath || currentPath === hrefNorm) {
                a.classList.add('active');
                bestMatch = a; bestScore = 100;
            } else if (hrefPath !== '/' && hrefPath !== '' && currentPath.indexOf(hrefPath) === 0) {
                var score = hrefPath.length;
                if (score > bestScore) { bestMatch = a; bestScore = score; }
            }
        });
        if (bestMatch && bestScore < 100 && bestScore > 0) {
            bestMatch.classList.add('active');
        }
    })();
    </script>

    <!-- V18 动态溢出检测：导航栏超宽时自动切换汉堡菜单 -->
    <script nonce="<%= Session("csp_nonce") %>">
    (function() {
        'use strict';
        var mainNav = document.querySelector('.main-nav');
        var navList = document.querySelector('.nav-list');
        var desktopBtn = document.getElementById('desktopHamburgerBtn');
        if (!mainNav || !navList) return;

        var container = mainNav.querySelector('.container');
        var body = document.body;
        var checking = false;

        function checkOverflow() {
            if (checking) return;
            checking = true;
            // 使用 requestAnimationFrame 确保 DOM 已更新
            requestAnimationFrame(function() {
                var listWidth = navList.scrollWidth;
                var containerWidth = container ? container.clientWidth : mainNav.clientWidth;
                var overflows = listWidth > containerWidth + 2; // 2px 容差
                if (overflows) {
                    body.classList.add('nav-overflow');
                } else {
                    body.classList.remove('nav-overflow');
                }
                checking = false;
            });
        }

        // 初始检测
        checkOverflow();

        // 使用 ResizeObserver 监听容器大小变化
        if (container && window.ResizeObserver) {
            var ro = new ResizeObserver(function() {
                checkOverflow();
            });
            ro.observe(container);
        }

        // 回退方案：监听窗口 resize
        var resizeTimer;
        window.addEventListener('resize', function() {
            clearTimeout(resizeTimer);
            resizeTimer = setTimeout(checkOverflow, 150);
        });
    })();
    </script>

    <style nonce="<%= Session("csp_nonce") %>">
    /* V18: 秒杀/拼团等热链样式 */
    .nav-hot { color: #e65100 !important; font-weight: 600; position: relative; }
    .nav-hot::after { content: 'HOT'; position: absolute; top: -1px; right: -4px; font-size: 9px; background: linear-gradient(135deg, #ff416c, #ff4b2b); color: #fff; padding: 1px 5px; border-radius: 3px; line-height: 1.4; font-weight: 700; letter-spacing: 0.5px; pointer-events: none; z-index: 1; }
    .nav-hot:hover { color: #bf360c !important; background: rgba(230,81,0,0.06) !important; }
    </style>

    <!-- V11.1 PWA Service Worker注册 -->
    <script nonce="<%= Session("csp_nonce") %>">
    if ('serviceWorker' in navigator) {
        window.addEventListener('load', () => {
            navigator.serviceWorker.register('/sw.js')
                .then(reg => console.log('[PWA] SW registered:', reg.scope))
                .catch(err => console.error('[PWA] SW failed:', err));
        });
    }
    // V18: Cookie 同意处理
    function acceptCookies(level) {
        var d = new Date();
        d.setFullYear(d.getFullYear() + 1);
        document.cookie = 'cookie_consent=' + level + '; expires=' + d.toUTCString() + '; path=/; SameSite=Lax';
        var banner = document.getElementById('cookieBanner');
        if (banner) {
            banner.style.animation = 'cookieSlideDown 0.3s ease forwards';
            setTimeout(function() { banner.style.display = 'none'; }, 300);
        }
        // V18: 记录 Cookie 同意（通过 fetch 发送到后端）
        fetch('/api/cookie_consent.asp', {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: 'consent=' + encodeURIComponent(level) + '&csrf_token=' + encodeURIComponent(csrfToken)
        }).catch(function() {});
    }
    </script>
    <style nonce="<%= Session("csp_nonce") %>">
    @keyframes cookieSlideDown { from { transform: translateY(0); } to { transform: translateY(100%); } }
    </style>

    <!-- 主内容区 -->
    <main class="main-content">
