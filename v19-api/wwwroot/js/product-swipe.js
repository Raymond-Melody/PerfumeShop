/**
 * PerfumeShop V14.6 - 产品图片手势滑动
 * 支持触摸滑动切换产品图片
 */

(function() {
    'use strict';
    
    // 手势滑动功能
    function initProductImageSwipe() {
        var productImages = document.querySelectorAll('.product-image-gallery');
        
        productImages.forEach(function(gallery) {
            var images = gallery.querySelectorAll('img');
            if (images.length < 2) return;
            
            var currentIndex = 0;
            var startX = 0;
            var startY = 0;
            var isSwiping = false;
            
            // 触摸开始
            gallery.addEventListener('touchstart', function(e) {
                startX = e.touches[0].clientX;
                startY = e.touches[0].clientY;
                isSwiping = true;
            }, { passive: true });
            
            // 触摸移动
            gallery.addEventListener('touchmove', function(e) {
                if (!isSwiping) return;
                
                var currentX = e.touches[0].clientX;
                var currentY = e.touches[0].clientY;
                var diffX = startX - currentX;
                var diffY = startY - currentY;
                
                // 水平滑动检测（避免垂直滚动）
                if (Math.abs(diffX) > Math.abs(diffY) && Math.abs(diffX) > 30) {
                    e.preventDefault();
                }
            }, { passive: false });
            
            // 触摸结束
            gallery.addEventListener('touchend', function(e) {
                if (!isSwiping) return;
                
                var endX = e.changedTouches[0].clientX;
                var diffX = startX - endX;
                
                // 滑动阈值：50px
                if (Math.abs(diffX) > 50) {
                    if (diffX > 0) {
                        // 向左滑动 - 下一张
                        currentIndex = (currentIndex + 1) % images.length;
                    } else {
                        // 向右滑动 - 上一张
                        currentIndex = (currentIndex - 1 + images.length) % images.length;
                    }
                    
                    // 显示当前图片
                    images.forEach(function(img, index) {
                        img.style.display = index === currentIndex ? 'block' : 'none';
                    });
                    
                    // 更新指示器
                    updateIndicators(gallery, currentIndex);
                }
                
                isSwiping = false;
            }, { passive: true });
        });
    }
    
    // 更新图片指示器
    function updateIndicators(gallery, currentIndex) {
        var indicators = gallery.querySelectorAll('.image-indicator');
        indicators.forEach(function(dot, index) {
            dot.classList.toggle('active', index === currentIndex);
        });
    }
    
    // 页面加载完成后初始化
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initProductImageSwipe);
    } else {
        initProductImageSwipe();
    }
})();
