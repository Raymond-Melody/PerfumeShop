    </main>

    <!-- 页脚 -->
    <footer class="main-footer">
        <div class="container">
            <div class="footer-grid">
                <!-- 关于我们 -->
                <div class="footer-section">
                    <h4>关于香氛定制</h4>
                    <p>我们致力于为每一位顾客打造独一无二的专属香水，让香氛成为你个性的延伸。</p>
                    <div class="social-links">
                        <a href="#"><i class="fab fa-weixin"></i></a>
                        <a href="#"><i class="fab fa-weibo"></i></a>
                        <a href="#"><i class="fas fa-envelope"></i></a>
                    </div>
                </div>

                <!-- 快速链接 -->
                <div class="footer-section">
                    <h4>快速链接</h4>
                    <ul>
                        <li><a href="/products.asp">全部产品</a></li>
                        <li><a href="/customize.asp">开始定制</a></li>
                        <li><a href="/about.asp">品牌故事</a></li>
                        <li><a href="/contact.asp">联系我们</a></li>
                    </ul>
                </div>

                <!-- 客户服务 -->
                <div class="footer-section">
                    <h4>客户服务</h4>
                    <ul>
                        <li><a href="/help/shipping.asp">配送说明</a></li>
                        <li><a href="/help/return.asp">退换政策</a></li>
                        <li><a href="/help/payment.asp">支付方式</a></li>
                        <li><a href="/help/faq.asp">常见问题</a></li>
                    </ul>
                </div>

                <!-- 联系方式 -->
                <div class="footer-section">
                    <h4>联系我们</h4>
                    <p><i class="fas fa-phone"></i> <%= SITE_PHONE %></p>
                    <p><i class="fas fa-envelope"></i> <%= SITE_EMAIL %></p>
                    <p><i class="fas fa-clock"></i> 周一至周日 9:00-21:00</p>
                </div>
            </div>

            <div class="footer-bottom">
                <p>&copy; 2024 香氛定制 版权所有 | <a href="/privacy.asp">隐私政策</a> | <a href="/terms.asp">服务条款</a></p>
            </div>
        </div>
    </footer>

    <!-- 返回顶部 -->
    <a href="#" class="back-to-top" id="backToTop">
        <i class="fas fa-arrow-up"></i>
    </a>

    <!-- 公共脚本 -->
    <!-- V12.0 图片懒加载 -->
    <script src="/js/lazy-load.js?v=14.6"></script>
    <!-- V12.0 主题切换 -->
    <script src="/js/theme-toggle.js?v=14.6" defer></script>
    <!-- V12.1 产品图库手势支持 -->
    <script src="/js/product-gallery.js?v=14.6" defer></script>
    <!-- V12.2 购物车动画 -->
    <script src="/js/cart-animation.js?v=14.6" defer></script>
    <!-- V12.3 筛选优化 -->
    <script src="/js/filter-optimization.js?v=14.6" defer></script>
    <!-- V13.2 骨架屏加载 -->
    <script src="/js/skeleton-loader.js?v=14.6"></script>
    <!-- V14.6 PWA 安装提示 -->
    <script src="/js/pwa-install.js?v=14.6" defer></script>
    <script src="/js/main.js?v=14.6"></script>
    
    <script nonce="<%= Session("csp_nonce") %>">
    // 更新购物车数量
    function updateCartCount() {
        $.get('/api/cart_count.asp', function(data) {
            $('#cartCount').text(data);
        });
    }
    $(document).ready(function() {
        updateCartCount();
    });
    </script>
</body>
</html>
