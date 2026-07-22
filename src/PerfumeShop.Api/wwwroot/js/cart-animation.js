/**
 * PerfumeShop V14.6 - 购物车动画
 * 添加商品飞入动画 + 购物车图标跳动 + 数量数字滚动
 */

(function() {
    'use strict';
    
    // 添加商品到购物车的飞入动画
    function flyToCart(button) {
        var cartIcon = document.querySelector('.cart-icon, .fa-shopping-cart');
        if (!cartIcon) return;
        
        // 获取按钮和购物车图标的位置
        var buttonRect = button.getBoundingClientRect();
        var cartRect = cartIcon.getBoundingClientRect();
        
        // 创建飞行动画元素
        var flyer = document.createElement('div');
        flyer.className = 'cart-flyer';
        flyer.innerHTML = '<i class="fas fa-shopping-cart"></i>';
        flyer.style.cssText = 'position:fixed;z-index:9999;pointer-events:none;' +
            'left:' + buttonRect.left + 'px;' +
            'top:' + buttonRect.top + 'px;' +
            'width:20px;height:20px;' +
            'transition:all 0.6s cubic-bezier(0.25, 0.46, 0.45, 0.94);';
        
        document.body.appendChild(flyer);
        
        // 触发重排
        flyer.offsetHeight;
        
        // 飞到购物车图标位置
        flyer.style.left = cartRect.left + 'px';
        flyer.style.top = cartRect.top + 'px';
        flyer.style.opacity = '0';
        flyer.style.transform = 'scale(0.3)';
        
        // 动画结束后移除元素并触发购物车跳动
        setTimeout(function() {
            document.body.removeChild(flyer);
            bounceCartIcon(cartIcon);
        }, 600);
    }
    
    // 购物车图标跳动动画
    function bounceCartIcon(cartIcon) {
        cartIcon.classList.add('cart-bounce');
        
        // 更新购物车数量
        updateCartCount(function() {
            setTimeout(function() {
                cartIcon.classList.remove('cart-bounce');
            }, 400);
        });
    }
    
    // 更新购物车数量（带数字滚动动画）
    function updateCartCount(callback) {
        var countElement = document.querySelector('.cart-count, #cartCount');
        if (!countElement) {
            if (callback) callback();
            return;
        }
        
        var currentCount = parseInt(countElement.textContent) || 0;
        
        // AJAX获取最新数量
        if (typeof $ !== 'undefined') {
            $.get('/api/cart_count.asp', function(data) {
                var newCount = parseInt(data) || 0;
                
                if (newCount !== currentCount) {
                    animateNumber(countElement, currentCount, newCount, callback);
                } else {
                    if (callback) callback();
                }
            }).fail(function() {
                if (callback) callback();
            });
        } else {
            if (callback) callback();
        }
    }
    
    // 数字滚动动画
    function animateNumber(element, from, to, callback) {
        var duration = 400;
        var startTime = null;
        
        function step(timestamp) {
            if (!startTime) startTime = timestamp;
            var progress = Math.min((timestamp - startTime) / duration, 1);
            
            // easeOutQuad缓动
            var eased = 1 - (1 - progress) * (1 - progress);
            var current = Math.floor(from + (to - from) * eased);
            
            element.textContent = current;
            
            if (progress < 1) {
                requestAnimationFrame(step);
            } else {
                element.textContent = to;
                if (callback) callback();
            }
        }
        
        requestAnimationFrame(step);
    }
    
    // 移除商品的淡出动画
    function fadeOutItem(item) {
        item.style.transition = 'all 0.3s ease';
        item.style.opacity = '0';
        item.style.transform = 'translateX(-20px)';
        
        setTimeout(function() {
            item.style.height = '0';
            item.style.margin = '0';
            item.style.padding = '0';
            item.style.overflow = 'hidden';
        }, 300);
    }
    
    // 监听添加购物车按钮
    function initCartAnimations() {
        // 监听所有"加入购物车"按钮
        var addToCartButtons = document.querySelectorAll('.btn-add-cart, .add-to-cart');
        addToCartButtons.forEach(function(button) {
            button.addEventListener('click', function(e) {
                flyToCart(this);
            });
        });
        
        // 监听删除按钮
        var removeButtons = document.querySelectorAll('.btn-remove, .remove-item');
        removeButtons.forEach(function(button) {
            button.addEventListener('click', function(e) {
                var cartItem = this.closest('.cart-item, .order-item');
                if (cartItem) {
                    fadeOutItem(cartItem);
                }
            });
        });
    }
    
    // 页面加载完成后初始化
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initCartAnimations);
    } else {
        initCartAnimations();
    }
    
    // 暴露全局方法
    window.flyToCart = flyToCart;
    window.bounceCartIcon = bounceCartIcon;
    window.fadeOutItem = fadeOutItem;
})();
