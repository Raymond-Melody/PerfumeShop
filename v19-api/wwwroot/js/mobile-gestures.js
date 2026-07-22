/**
 * PerfumeShop V18.0 - 移动端手势支持
 * 滑动切换产品图、下拉刷新、长按操作
 */
(function() {
    'use strict';

    // ===== 下拉刷新 =====
    var PullToRefresh = {
        threshold: 80,
        pulling: false,
        startY: 0,
        indicator: null,
        container: null,

        init: function(containerSelector, onRefresh) {
            var self = this;
            this.container = document.querySelector(containerSelector);
            if (!this.container) return;

            // 创建指示器
            this.indicator = document.createElement('div');
            this.indicator.className = 'pull-to-refresh';
            this.indicator.innerHTML = '<div class="spinner"></div>';
            this.container.insertBefore(this.indicator, this.container.firstChild);

            // 仅在顶部时启用
            this.container.addEventListener('touchstart', function(e) {
                if (window.scrollY <= 5) {
                    self.startY = e.touches[0].clientY;
                    self.pulling = true;
                }
            }, { passive: true });

            this.container.addEventListener('touchmove', function(e) {
                if (!self.pulling) return;
                var dy = e.touches[0].clientY - self.startY;
                if (dy > 0 && window.scrollY <= 5) {
                    var pullDist = Math.min(dy * 0.4, 80);
                    self.indicator.style.height = pullDist + 'px';
                    if (pullDist >= self.threshold) {
                        self.indicator.classList.add('active');
                    }
                }
            }, { passive: true });

            this.container.addEventListener('touchend', function() {
                if (!self.pulling) return;
                self.pulling = false;
                var pullDist = parseInt(self.indicator.style.height) || 0;
                if (pullDist >= self.threshold) {
                    self.indicator.style.height = '52px';
                    if (typeof onRefresh === 'function') {
                        onRefresh(function() {
                            self.indicator.style.height = '0';
                            self.indicator.classList.remove('active');
                        });
                    } else {
                        setTimeout(function() {
                            self.indicator.style.height = '0';
                            self.indicator.classList.remove('active');
                        }, 800);
                    }
                } else {
                    self.indicator.style.height = '0';
                    self.indicator.classList.remove('active');
                }
            });
        }
    };

    // ===== 图片滑动切换 =====
    var ImageSwiper = {
        init: function(containerSelector) {
            var containers = document.querySelectorAll(containerSelector);
            containers.forEach(function(container) {
                ImageSwiper.attach(container);
            });
        },

        attach: function(container) {
            var images = container.querySelectorAll('img');
            if (images.length < 2) return;

            var currentIndex = 0;
            var startX = 0;
            var isSwiping = false;
            var wrapper = container.querySelector('.swipe-wrapper');
            if (!wrapper) {
                wrapper = document.createElement('div');
                wrapper.className = 'swipe-wrapper';
                wrapper.style.cssText = 'display:flex;overflow-x:auto;scroll-snap-type:x mandatory;-webkit-overflow-scrolling:touch;scroll-behavior:smooth;';
                images.forEach(function(img) {
                    img.style.cssText = 'scroll-snap-align:start;flex-shrink:0;width:100%;';
                });
                while (container.firstChild) {
                    wrapper.appendChild(container.firstChild);
                }
                container.appendChild(wrapper);
            }

            // 指示器
            var dots = document.createElement('div');
            dots.className = 'swipe-dots';
            dots.style.cssText = 'display:flex;justify-content:center;gap:6px;padding:8px 0;';
            for (var i = 0; i < images.length; i++) {
                var dot = document.createElement('span');
                dot.className = 'swipe-dot';
                dot.style.cssText = 'width:8px;height:8px;border-radius:50%;background:' + (i === 0 ? '#8B4513' : '#ddd') + ';transition:background 0.3s;';
                dot.dataset.index = i;
                dot.addEventListener('click', function() {
                    var idx = parseInt(this.dataset.index);
                    wrapper.scrollTo({ left: wrapper.clientWidth * idx, behavior: 'smooth' });
                });
                dots.appendChild(dot);
            }
            container.appendChild(dots);

            // 滑动检测更新指示器
            wrapper.addEventListener('scroll', function() {
                var idx = Math.round(wrapper.scrollLeft / wrapper.clientWidth);
                if (idx !== currentIndex) {
                    currentIndex = idx;
                    var allDots = dots.querySelectorAll('.swipe-dot');
                    allDots.forEach(function(d, i) {
                        d.style.background = i === idx ? '#8B4513' : '#ddd';
                    });
                }
            });

            // 触摸滑动
            container.addEventListener('touchstart', function(e) {
                startX = e.touches[0].clientX;
                isSwiping = true;
            }, { passive: true });

            container.addEventListener('touchend', function(e) {
                if (!isSwiping) return;
                isSwiping = false;
                var endX = e.changedTouches[0].clientX;
                var diff = startX - endX;
                if (Math.abs(diff) > 50) {
                    if (diff > 0 && currentIndex < images.length - 1) {
                        currentIndex++;
                    } else if (diff < 0 && currentIndex > 0) {
                        currentIndex--;
                    }
                    wrapper.scrollTo({ left: wrapper.clientWidth * currentIndex, behavior: 'smooth' });
                }
            });
        }
    };

    // ===== 暴露到全局 =====
    window.PullToRefresh = PullToRefresh;
    window.ImageSwiper = ImageSwiper;

    // ===== 自动初始化（页面加载后）=====
    document.addEventListener('DOMContentLoaded', function() {
        // 产品图片滑动
        var productGalleries = document.querySelectorAll('.product-gallery, .product-images, .gallery-container');
        productGalleries.forEach(function(g) {
            ImageSwiper.attach(g);
        });
    });
})();
