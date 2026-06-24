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
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- V16.0 资源预加载 -->
    <link rel="dns-prefetch" href="//cdnjs.cloudflare.com">
    <link rel="dns-prefetch" href="//code.jquery.com">
    <link rel="preconnect" href="https://cdnjs.cloudflare.com" crossorigin>
    <link rel="preconnect" href="https://code.jquery.com" crossorigin>
    <link rel="preload" href="/css/design-tokens.css?v=16.2" as="style">
    <link rel="preload" href="/css/style.css?v=16.2" as="style">
    <link rel="preload" href="https://code.jquery.com/jquery-3.6.0.min.js" as="script" crossorigin>
    <!-- V14.6 PWA 增强 -->
    <meta name="color-scheme" content="light only">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="default">
    <meta name="apple-mobile-web-app-title" content="香氛定制">
    <link rel="apple-touch-icon" href="/images/icons/icon-192x192.png">
    <meta name="description" content="专业个性化定制香水，打造专属于你的独特香氛体验">
    <meta name="keywords" content="香水定制,个性化香水,定制香氛,专属香水">
    <title><%= SITE_NAME %></title>
    <link rel="icon" href="/favicon.ico" type="image/x-icon">
    <link rel="manifest" href="/manifest.json">
    <meta name="theme-color" content="#8B4513">
    <link rel="stylesheet" href="/css/design-tokens.css?v=16.2">
    <link rel="stylesheet" href="/css/style.css?v=16.2">
    <link rel="stylesheet" href="/css/pages.css?v=16.2">
    <link rel="stylesheet" href="/css/buttons.css?v=16.2">
    <link rel="stylesheet" href="/css/responsive.css?v=16.2">
    <link rel="stylesheet" href="/css/lazy-load.css?v=16.2" media="print" onload="this.media='all'">
    <link rel="stylesheet" href="/css/cart-animation.css?v=16.2" media="print" onload="this.media='all'">
    <link rel="stylesheet" href="/css/filter-optimization.css?v=16.2">
    <link rel="stylesheet" href="/css/theme.css?v=16.2">
    <link rel="stylesheet" href="/css/skeleton.css?v=16.2" media="print" onload="this.media='all'">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <% If amphtmlLink <> "" Then Response.Write amphtmlLink %>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
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
            <span><i class="fas fa-gift"></i> 首单立减50元 | 满299元免运费</span>
            <span class="top-bar-right">
                <i class="fas fa-phone"></i> <%= SITE_PHONE %>
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
                        <span>香氛定制</span>
                    </a>
                </div>

                <!-- 搜索框 -->
                <div class="search-box">
                    <form action="/products.asp" method="get">
                        <input type="text" name="keyword" placeholder="搜索您想要的香水...">
                        <button type="submit"><i class="fas fa-search"></i></button>
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
                            <a href="/user/index.asp"><i class="fas fa-user-circle"></i> 个人中心</a>
                            <a href="/user/orders.asp"><i class="fas fa-list"></i> 我的订单</a>
                            <a href="/user/settings.asp"><i class="fas fa-cog"></i> 账户设置</a>
                            <a href="/user/addresses.asp"><i class="fas fa-map-marker-alt"></i> 收货地址</a>
                            <a href="/user/favorites.asp"><i class="fas fa-heart"></i> 我的收藏</a>
                            <a href="/user/logout.asp"><i class="fas fa-sign-out-alt"></i> 退出登录</a>
                        </div>
                    </div>
                    <%
                    Else
                    %>
                    <a href="/user/login.asp" class="btn-login"><i class="fas fa-sign-in-alt"></i> 登录</a>
                    <a href="/user/register.asp" class="btn-register">注册</a>
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
                <li><a href="/index.asp"><i class="fas fa-home"></i> 首页</a></li>
                <li><a href="/products.asp"><i class="fas fa-flask"></i> 全部香水</a></li>
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
                <li><a href="/about.asp"><i class="fas fa-book"></i> 品牌故事</a></li>
                <li><a href="/contact.asp"><i class="fas fa-envelope"></i> 联系我们</a></li>
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
