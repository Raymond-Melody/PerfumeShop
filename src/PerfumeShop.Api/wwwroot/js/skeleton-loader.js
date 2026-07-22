/**
 * PerfumeShop V14.6 - 骨架屏加载
 * 改善用户感知性能，消除白屏等待
 */

(function() {
    'use strict';
    
    // 骨架屏配置
    var config = {
        productList: {
            container: '.product-grid',
            count: 8,
            template: 'product'
        },
        productDetail: {
            container: '.product-detail',
            template: 'detail'
        },
        cart: {
            container: '.cart-items',
            count: 3,
            template: 'cart'
        }
    };
    
    // 创建产品骨架屏
    function createProductSkeletons(count) {
        var html = '';
        for (var i = 0; i < count; i++) {
            html += '<div class="product-skeleton">' +
                '<div class="skeleton-image skeleton"></div>' +
                '<div class="skeleton-title skeleton"></div>' +
                '<div class="skeleton-desc skeleton"></div>' +
                '<div class="skeleton-price skeleton"></div>' +
                '<div class="skeleton-btn skeleton"></div>' +
                '</div>';
        }
        return html;
    }
    
    // 创建产品详情骨架屏
    function createDetailSkeleton() {
        return '<div class="product-detail-skeleton">' +
            '<div class="skeleton-gallery skeleton"></div>' +
            '<div class="skeleton-title skeleton"></div>' +
            '<div class="skeleton-price skeleton"></div>' +
            '<div class="skeleton-desc skeleton"></div>' +
            '<div class="skeleton-desc skeleton short"></div>' +
            '<div class="skeleton-specs">' +
            '<div class="skeleton-spec-item">' +
            '<div class="skeleton-spec-label skeleton"></div>' +
            '<div class="skeleton-spec-value skeleton"></div>' +
            '</div>' +
            '<div class="skeleton-spec-item">' +
            '<div class="skeleton-spec-label skeleton"></div>' +
            '<div class="skeleton-spec-value skeleton"></div>' +
            '</div>' +
            '</div>' +
            '</div>';
    }
    
    // 创建购物车骨架屏
    function createCartSkeletons(count) {
        var html = '';
        for (var i = 0; i < count; i++) {
            html += '<div class="cart-skeleton">' +
                '<div class="skeleton-item">' +
                '<div class="skeleton-item-image skeleton"></div>' +
                '<div class="skeleton-item-info">' +
                '<div class="skeleton-item-title skeleton"></div>' +
                '<div class="skeleton-item-price skeleton"></div>' +
                '<div class="skeleton-item-qty skeleton"></div>' +
                '</div>' +
                '</div>' +
                '</div>';
        }
        return html;
    }
    
    // 显示骨架屏（仅用于AJAX动态加载的容器）
    function showSkeleton(container, type, count) {
        var el = document.querySelector(container);
        if (!el) return;
        
        // V13.2修复：如果容器已有服务端渲染的真实内容，不显示骨架屏
        if (hasServerRenderedContent(el)) return;
        
        var html = '';
        switch(type) {
            case 'product':
                html = '<div class="skeleton-grid">' + createProductSkeletons(count) + '</div>';
                break;
            case 'detail':
                html = createDetailSkeleton();
                break;
            case 'cart':
                html = createCartSkeletons(count);
                break;
        }
        
        // 使用prepend而非innerHTML，保留已有内容
        el.insertAdjacentHTML('afterbegin', html);
    }
    
    // 检测容器是否已有服务端渲染的真实内容
    function hasServerRenderedContent(el) {
        // 检查是否有真实产品卡片
        if (el.querySelector('.product-card, .cart-item')) return true;
        // 检查是否有产品图片、价格、标题等关键元素
        if (el.querySelector('.main-image, .product-gallery, .product-info-detail, .price-section')) return true;
        // 检查文本节点（排除空白）是否超过阈值，表明有丰富内容
        var textLen = (el.textContent || '').replace(/\s+/g, '').length;
        if (textLen > 50) return true;
        return false;
    }
    
    // 隐藏骨架屏
    function hideSkeleton(container) {
        var el = document.querySelector(container);
        if (!el) return;
        
        var skeletons = el.querySelectorAll('.skeleton');
        if (skeletons.length === 0) return;
        
        // 添加淡出动画
        skeletons.forEach(function(s) {
            s.classList.add('skeleton-fade-out');
        });
        
        // 动画结束后移除
        setTimeout(function() {
            // 等待真实内容加载完成后移除骨架屏
            var hasContent = el.querySelector('.product-card, .product-detail-content, .cart-item');
            if (hasContent) {
                var skeletonContainers = el.querySelectorAll('.product-skeleton, .product-detail-skeleton, .cart-skeleton');
                skeletonContainers.forEach(function(s) {
                    s.remove();
                });
            }
        }, 300);
    }
    
    // 自动检测并显示骨架屏（仅对AJAX加载的空容器生效）
    function autoShowSkeletons() {
        // 产品列表页：仅当使用AJAX动态加载且容器为空时显示
        var productList = document.querySelector('.product-grid, .products-grid');
        if (productList && productList.children.length === 0) {
            showSkeleton('.product-grid, .products-grid', 'product', config.productList.count);
        }
        
        // 产品详情页：Classic ASP服务端渲染，内容已在HTML中，不显示骨架屏
        // 仅当通过AJAX加载产品详情且容器确实为空时才显示
        var productDetail = document.querySelector('.product-detail');
        if (productDetail && productDetail.children.length === 0) {
            showSkeleton('.product-detail', 'detail');
        }
        
        // 购物车页：仅当购物车容器完全为空时显示
        var cartItems = document.querySelector('.cart-items');
        if (cartItems && cartItems.children.length === 0) {
            showSkeleton('.cart-items', 'cart', config.cart.count);
        }
    }
    
    // 监听DOM变化，自动隐藏骨架屏
    function watchContentLoad() {
        var observer = new MutationObserver(function(mutations) {
            mutations.forEach(function(mutation) {
                mutation.addedNodes.forEach(function(node) {
                    if (node.nodeType === 1) {
                        if (node.classList.contains('product-card') || 
                            node.classList.contains('product-detail-content') ||
                            node.classList.contains('cart-item')) {
                            
                            var container = node.parentElement;
                            if (container) {
                                hideSkeleton('.' + container.className.split(' ')[0]);
                            }
                        }
                    }
                });
            });
        });
        
        observer.observe(document.body, {
            childList: true,
            subtree: true
        });
    }
    
    // 页面加载完成后初始化
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', function() {
            autoShowSkeletons();
            watchContentLoad();
        });
    } else {
        autoShowSkeletons();
        watchContentLoad();
    }
    
    // 暴露全局方法
    window.showSkeleton = showSkeleton;
    window.hideSkeleton = hideSkeleton;
})();
