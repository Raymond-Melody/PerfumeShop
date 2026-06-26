<!--#include file="i18n.asp"-->
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
    <meta name="apple-mobile-web-app-title" content="<% If FEATURE_I18N Then Response.Write T("header_meta_title", Empty) Else %>香氛定制<% End If %>">
    <link rel="apple-touch-icon" href="/images/icons/icon-192x192.png">
    <meta name="description" content="<% If FEATURE_I18N Then Response.Write T("header_meta_description", Empty) Else %>专业个性化定制香水，打造专属于你的独特香氛体验<% End If %>">
    <meta name="keywords" content="<% If FEATURE_I18N Then Response.Write T("header_meta_keywords", Empty) Else %>香水定制,个性化香水,定制香氛,专属香水<% End If %>">
    <title><%= SITE_NAME %></title>
    <link rel="icon" href="/favicon.ico" type="image/x-icon">
    <link rel="manifest" href="/manifest.json">
    <meta name="theme-color" content="#8B4513">
    <link rel="stylesheet" href="/css/design-tokens.css?v=17.0">
    <link rel="stylesheet" href="/css/style.css?v=17.0">
    <link rel="stylesheet" href="/css/pages.css?v=17.0">
    <link rel="stylesheet" href="/css/buttons.css?v=17.0">
    <link rel="stylesheet" href="/css/responsive.css?v=17.0">
    <link rel="stylesheet" href="/css/lazy-load.css?v=17.0" media="print" onload="this.media='all'">
    <link rel="stylesheet" href="/css/cart-animation.css?v=17.0" media="print" onload="this.media='all'">
    <link rel="stylesheet" href="/css/filter-optimization.css?v=17.0">
    <link rel="stylesheet" href="/css/theme.css?v=17.0">
    <link rel="stylesheet" href="/css/skeleton.css?v=17.0" media="print" onload="this.media='all'">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <% If amphtmlLink <> "" Then Response.Write amphtmlLink %>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
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
                    <a href="?lang=zh-CN" class="lang-link<%= IIF(I18N_GetLocale()="zh-CN"," active","") %>"><% If FEATURE_I18N Then Response.Write T("header_lang_zh", Empty) Else %>中文<% End If %></a>
                    <span class="lang-divider">|</span>
                    <a href="?lang=en-US" class="lang-link<%= IIF(I18N_GetLocale()="en-US"," active","") %>"><% If FEATURE_I18N Then Response.Write T("header_lang_en", Empty) Else %>EN<% End If %></a>
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
                        <input type="text" name="keyword" id="searchKeyword" placeholder="<% If FEATURE_I18N Then Response.Write T("search_placeholder", Empty) Else Response.Write "搜索您想要的香水..." End If %>" autocomplete="off">
                        <button type="submit"><i class="fas fa-search"></i></button>
                        <div class="search-suggestions" id="searchSuggestions"></div>
                    </form>
                </div>

                <!-- 用户菜单 -->
                <div class="user-menu">
                    <%
                    If Session("UserID") <> "" Then
                    %>
                    <div class="user-info dropdown">
                        <a href="#"><i class="fas fa-user"></i> <%= Session("Username") %></a>
                        <div class="dropdown-menu">
                            <a href="/user/index.asp"><i class="fas fa-user-circle"></i> <% If FEATURE_I18N Then Response.Write T("header_user_center", Empty) Else %>个人中心<% End If %></a>
                            <a href="/user/orders.asp"><i class="fas fa-list"></i> <% If FEATURE_I18N Then Response.Write T("header_my_orders", Empty) Else %>我的订单<% End If %></a>
                            <a href="/user/settings.asp"><i class="fas fa-cog"></i> <% If FEATURE_I18N Then Response.Write T("header_account_settings", Empty) Else %>账户设置<% End If %></a>
                            <a href="/user/addresses.asp"><i class="fas fa-map-marker-alt"></i> <% If FEATURE_I18N Then Response.Write T("header_shipping_address", Empty) Else %>收货地址<% End If %></a>
                            <a href="/user/favorites.asp"><i class="fas fa-heart"></i> <% If FEATURE_I18N Then Response.Write T("header_my_favorites", Empty) Else %>我的收藏<% End If %></a>
                            <a href="/user/logout.asp"><i class="fas fa-sign-out-alt"></i> <% If FEATURE_I18N Then Response.Write T("header_logout", Empty) Else %>退出登录<% End If %></a>
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
                <%
                ' 动态生成商品类型导航
                ' 使用 On Error Resume Next 检查函数是否可用
                On Error Resume Next
                Dim navTypes, nIdx, navCode, navName, navIcon
                navTypes = GetActiveProductTypes()
                If Err.Number = 0 Then
                    If IsArray(navTypes) Then
                        For nIdx = 0 To UBound(navTypes, 1)
                            navCode = navTypes(nIdx, 0)
                            navName = navTypes(nIdx, 2)  ' NavName
                            navIcon = navTypes(nIdx, 4)  ' Icon
                            If navName <> "" Then  ' 只显示有NavName的类型
                %>
                <li><a href="/products.asp?type=<%= Server.URLEncode(navCode) %>"><i class="<%= Server.HTMLEncode(navIcon) %>"></i> <%= Server.HTMLEncode(navName) %></a></li>
                <%
                            End If
                        Next
                    End If
                End If
                On Error GoTo 0
                %>
                <li><a href="/about.asp"><i class="fas fa-book"></i> <% If FEATURE_I18N Then Response.Write T("about", Empty) Else Response.Write "品牌故事" End If %></a></li>
                <li><a href="/contact.asp"><i class="fas fa-envelope"></i> <% If FEATURE_I18N Then Response.Write T("contact", Empty) Else Response.Write "联系我们" End If %></a></li>
            </ul>
        </div>
    </nav>

    <!-- V11.1 PWA Service Worker注册 -->
    <script nonce="<%= Session("csp_nonce") %>">
    if ('serviceWorker' in navigator) {
        window.addEventListener('load', () => {
            navigator.serviceWorker.register('/sw.js')
                .then(reg => console.log('[PWA] SW registered:', reg.scope))
                .catch(err => console.error('[PWA] SW failed:', err));
        });
    }
    </script>

    <!-- 主内容区 -->
    <main class="main-content">
