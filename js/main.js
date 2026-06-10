/**
 * 香氛定制电商网站 - 主JavaScript文件
 */

$(document).ready(function() {
    // 初始化
    initBackToTop();
    initDropdowns();
    initNavigation();
    updateCartCount();
});

/**
 * 返回顶部按钮
 */
function initBackToTop() {
    var $backToTop = $('#backToTop');
    
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

/**
 * 下拉菜单
 */
function initDropdowns() {
    // 移动端点击切换
    if (window.innerWidth <= 768) {
        $('.dropdown > a').click(function(e) {
            e.preventDefault();
            $(this).parent().toggleClass('active');
        });
    }
}

/**
 * 导航菜单
 */
function initNavigation() {
    // 子菜单悬停效果
    $('.has-submenu').hover(
        function() {
            $(this).find('.submenu').stop().slideDown(200);
        },
        function() {
            $(this).find('.submenu').stop().slideUp(200);
        }
    );
}

/**
 * 更新购物车数量
 */
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

/**
 * 添加到购物车
 */
function addToCart(productId, options) {
    var data = {
        productId: productId
    };
    
    if (options) {
        $.extend(data, options);
    }
    
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

/**
 * 显示消息提示
 */
function showMessage(type, message) {
    var $msg = $('<div class="toast-message ' + type + '">' +
        '<i class="fas fa-' + (type === 'success' ? 'check-circle' : 'exclamation-circle') + '"></i> ' +
        message + '</div>');
    
    $('body').append($msg);
    
    setTimeout(function() {
        $msg.addClass('show');
    }, 10);
    
    setTimeout(function() {
        $msg.removeClass('show');
        setTimeout(function() {
            $msg.remove();
        }, 300);
    }, 3000);
}

/**
 * 图片懒加载
 */
function initLazyLoad() {
    var lazyImages = document.querySelectorAll('img[data-src]');
    
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
        
        lazyImages.forEach(function(img) {
            observer.observe(img);
        });
    } else {
        // 降级处理
        lazyImages.forEach(function(img) {
            img.src = img.dataset.src;
        });
    }
}

/**
 * 格式化货币
 */
function formatMoney(amount) {
    return '¥' + parseFloat(amount).toFixed(2);
}

/**
 * 表单验证
 */
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

/**
 * 邮箱验证
 */
function isValidEmail(email) {
    var re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return re.test(email);
}

/**
 * 手机号验证
 */
function isValidPhone(phone) {
    var re = /^1[3-9]\d{9}$/;
    return re.test(phone);
}

/**
 * 防抖函数
 */
function debounce(func, wait) {
    var timeout;
    return function() {
        var context = this, args = arguments;
        clearTimeout(timeout);
        timeout = setTimeout(function() {
            func.apply(context, args);
        }, wait);
    };
}

/**
 * 节流函数
 */
function throttle(func, limit) {
    var inThrottle;
    return function() {
        var context = this, args = arguments;
        if (!inThrottle) {
            func.apply(context, args);
            inThrottle = true;
            setTimeout(function() {
                inThrottle = false;
            }, limit);
        }
    };
}

// Toast消息样式（动态添加）
var toastStyle = document.createElement('style');
toastStyle.textContent = `
    .toast-message {
        position: fixed;
        top: 100px;
        left: 50%;
        transform: translateX(-50%) translateY(-20px);
        padding: 15px 30px;
        background: #fff;
        border-radius: 8px;
        box-shadow: 0 5px 25px rgba(0,0,0,0.2);
        z-index: 10000;
        opacity: 0;
        transition: all 0.3s ease;
    }
    .toast-message.show {
        opacity: 1;
        transform: translateX(-50%) translateY(0);
    }
    .toast-message.success {
        border-left: 4px solid #28a745;
    }
    .toast-message.success i {
        color: #28a745;
    }
    .toast-message.error {
        border-left: 4px solid #dc3545;
    }
    .toast-message.error i {
        color: #dc3545;
    }
`;
document.head.appendChild(toastStyle);
