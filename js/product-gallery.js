/**
 * PerfumeShop V14.6 - 产品详情页图片手势滑动
 * 支持触摸滑动切换产品图片 + 点击指示器切换
 */

(function() {
    'use strict';
    
    function initProductGallery() {
        var gallery = document.querySelector('.product-gallery');
        if (!gallery) return;
        
        var images = gallery.querySelectorAll('.gallery-image');
        if (images.length < 2) return;
        
        var currentIndex = 0;
        var startX = 0;
        var startY = 0;
        var isSwiping = false;
        
        // 创建指示器
        var indicators = document.createElement('div');
        indicators.className = 'gallery-indicators';
        images.forEach(function(img, index) {
            var dot = document.createElement('span');
            dot.className = 'indicator-dot' + (index === 0 ? ' active' : '');
            dot.addEventListener('click', function() {
                goToImage(index);
            });
            indicators.appendChild(dot);
        });
        gallery.appendChild(indicators);
        
        // 切换到指定图片
        function goToImage(index) {
            images.forEach(function(img, i) {
                img.classList.toggle('active', i === index);
            });
            
            var dots = indicators.querySelectorAll('.indicator-dot');
            dots.forEach(function(dot, i) {
                dot.classList.toggle('active', i === index);
            });
            
            currentIndex = index;
        }
        
        // 触摸事件
        gallery.addEventListener('touchstart', function(e) {
            startX = e.touches[0].clientX;
            startY = e.touches[0].clientY;
            isSwiping = true;
        }, { passive: true });
        
        gallery.addEventListener('touchmove', function(e) {
            if (!isSwiping) return;
            
            var diffX = startX - e.touches[0].clientX;
            var diffY = startY - e.touches[0].clientY;
            
            if (Math.abs(diffX) > Math.abs(diffY) && Math.abs(diffX) > 30) {
                e.preventDefault();
            }
        }, { passive: false });
        
        gallery.addEventListener('touchend', function(e) {
            if (!isSwiping) return;
            
            var endX = e.changedTouches[0].clientX;
            var diffX = startX - endX;
            
            if (Math.abs(diffX) > 50) {
                if (diffX > 0) {
                    // 向左滑动 - 下一张
                    goToImage((currentIndex + 1) % images.length);
                } else {
                    // 向右滑动 - 上一张
                    goToImage((currentIndex - 1 + images.length) % images.length);
                }
            }
            
            isSwiping = false;
        }, { passive: true });
        
        // 键盘导航
        document.addEventListener('keydown', function(e) {
            if (e.key === 'ArrowLeft') {
                goToImage((currentIndex - 1 + images.length) % images.length);
            } else if (e.key === 'ArrowRight') {
                goToImage((currentIndex + 1) % images.length);
            }
        });
    }
    
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initProductGallery);
    } else {
        initProductGallery();
    }
})();
