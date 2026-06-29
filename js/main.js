/**
 * PerfumeShop V17.0 - 主JavaScript模块
 * IIFE包装，避免全局作用域污染
 * 依赖: jQuery
 */
;(function(window, $, undefined) {
    'use strict';

    // ==================== 模块初始化 ====================
    $(document).ready(function() {
        initBackToTop();
        initDropdowns();
        initNavigation();
        updateCartCount();
    });

    // ==================== 返回顶部按钮 ====================
    function initBackToTop() {
        var $backToTop = $('#backToTop');
        if (!$backToTop.length) return;

        $(window).scroll(function() {
            if ($(this).scrollTop() > 300) {
                $backToTop.addClass('show');
            } else {
                $backToTop.removeClass('show');
            }
        });

        $backToTop.click(function(e) {
            e.preventDefault();
            $('html, body').animate({scrollTop: 0}, 500);
        });
    }

    // ==================== 下拉菜单 ====================
    function initDropdowns() {
        if (window.innerWidth <= 768) {
            $('.dropdown > a').click(function(e) {
                e.preventDefault();
                $(this).parent().toggleClass('active');
            });
        }
    }

    // ==================== 导航菜单 ====================
    function initNavigation() {
        $('.has-submenu').hover(
            function() {
                $(this).find('.submenu').stop().slideDown(200);
            },
            function() {
                $(this).find('.submenu').stop().slideUp(200);
            }
        );
    }

    // ==================== 购物车 ====================
    function updateCartCount() {
        $.get('/api/cart_count.asp', function(data) {
            var count = parseInt(data) || 0;
            $('#cartCount').text(count);
            if (count > 0) {
                $('#cartCount').show();
            } else {
                $('#cartCount').hide();
            }
        });
    }

    function addToCart(productId, options) {
        var data = { productId: productId };
        if (options) { $.extend(data, options); }

        $.post('/api/cart_add.asp', data, function(response) {
            if (response.success) {
                showMessage('success', '已添加到购物车');
                updateCartCount();
            } else {
                showMessage('error', response.message || '添加失败');
            }
        }, 'json').fail(function() {
            showMessage('error', '网络错误，请重试');
        });
    }

    // ==================== 消息提示 ====================
    function showMessage(type, message) {
        var iconClass = type === 'success' ? 'check-circle' : 'exclamation-circle';
        var $msg = $('<div class="toast-message ' + type + '">' +
            '<i class="fas fa-' + iconClass + '"></i> ' + message + '</div>');

        $('body').append($msg);
        setTimeout(function() { $msg.addClass('show'); }, 10);
        setTimeout(function() {
            $msg.removeClass('show');
            setTimeout(function() { $msg.remove(); }, 300);
        }, 3000);
    }

    // ==================== 图片懒加载 ====================
    function initLazyLoad() {
        var lazyImages = document.querySelectorAll('img[data-src]');
        if (!lazyImages.length) return;

        if ('IntersectionObserver' in window) {
            var observer = new IntersectionObserver(function(entries) {
                entries.forEach(function(entry) {
                    if (entry.isIntersecting) {
                        var img = entry.target;
                        img.src = img.dataset.src;
                        img.removeAttribute('data-src');
                        observer.unobserve(img);
                    }
                });
            });
            lazyImages.forEach(function(img) { observer.observe(img); });
        } else {
            lazyImages.forEach(function(img) { img.src = img.dataset.src; });
        }
    }

    // ==================== 工具函数 ====================
    function formatMoney(amount) {
        return '¥' + parseFloat(amount).toFixed(2);
    }

    function validateForm(formId) {
        var $form = $(formId);
        var valid = true;
        $form.find('[required]').each(function() {
            var $input = $(this);
            if (!$input.val().trim()) {
                $input.addClass('error');
                valid = false;
            } else {
                $input.removeClass('error');
            }
        });
        return valid;
    }

    function isValidEmail(email) {
        return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
    }

    function isValidPhone(phone) {
        return /^1[3-9]\d{9}$/.test(phone);
    }

    function debounce(func, wait) {
        var timeout;
        return function() {
            var context = this, args = arguments;
            clearTimeout(timeout);
            timeout = setTimeout(function() { func.apply(context, args); }, wait);
        };
    }

    function throttle(func, limit) {
        var inThrottle;
        return function() {
            var context = this, args = arguments;
            if (!inThrottle) {
                func.apply(context, args);
                inThrottle = true;
                setTimeout(function() { inThrottle = false; }, limit);
            }
        };
    }

    // ==================== 公共API暴露 ====================
    // 将需要外部调用的函数挂载到全局 PerfumeShop 命名空间
    window.PerfumeShop = {
        addToCart: addToCart,
        updateCartCount: updateCartCount,
        showMessage: showMessage,
        initLazyLoad: initLazyLoad,
        formatMoney: formatMoney,
        validateForm: validateForm,
        isValidEmail: isValidEmail,
        isValidPhone: isValidPhone,
        debounce: debounce,
        throttle: throttle
    };

})(window, jQuery);
