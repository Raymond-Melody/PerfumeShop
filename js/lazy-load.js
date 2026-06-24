/**
 * PerfumeShop V14.6 - 图片懒加载
 * 使用IntersectionObserver API实现按需加载
 */

(function() {
    'use strict';
    
    // 检查浏览器支持
    if (!('IntersectionObserver' in window)) {
        // 降级方案：直接加载所有图片
        loadAllImages();
        return;
    }
    
    // 懒加载配置
    var config = {
        rootMargin: '200px 0px', // 提前200px开始加载
        threshold: 0.01 // 1%可见时触发
    };
    
    // 创建观察器
    var observer = new IntersectionObserver(function(entries) {
        entries.forEach(function(entry) {
            if (entry.isIntersecting) {
                var img = entry.target;
                
                // 加载真实图片
                if (img.dataset.src) {
                    img.src = img.dataset.src;
                    img.removeAttribute('data-src');
                }
                
                // 加载srcset（响应式图片）
                if (img.dataset.srcset) {
                    img.srcset = img.dataset.srcset;
                    img.removeAttribute('data-srcset');
                }
                
                // 添加加载完成动画
                img.classList.add('loaded');
                
                // 停止观察
                observer.unobserve(img);
            }
        });
    }, config);
    
    // 观察所有懒加载图片
    function initLazyLoad() {
        var lazyImages = document.querySelectorAll('img[data-src]');
        lazyImages.forEach(function(img) {
            observer.observe(img);
        });
    }
    
    // 降级方案：直接加载所有图片
    function loadAllImages() {
        var lazyImages = document.querySelectorAll('img[data-src]');
        lazyImages.forEach(function(img) {
            if (img.dataset.src) {
                img.src = img.dataset.src;
            }
            if (img.dataset.srcset) {
                img.srcset = img.dataset.srcset;
            }
            img.classList.add('loaded');
        });
    }
    
    // 页面加载完成后初始化
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initLazyLoad);
    } else {
        initLazyLoad();
    }
    
    // 暴露全局方法（用于动态添加的图片）
    window.refreshLazyLoad = function() {
        initLazyLoad();
    };
})();
