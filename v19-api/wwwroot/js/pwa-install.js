/**
 * PerfumeShop V14.6 - PWA 安装提示
 * 监听 beforeinstallprompt 事件，在用户浏览足够页面后显示安装横幅
 */

(function() {
    'use strict';

    var INSTALL_KEY = 'perfumeshop_pwa_install';
    var DISMISS_KEY = 'perfumeshop_pwa_dismissed';
    var PAGE_COUNT_KEY = 'perfumeshop_page_views';
    var REQUIRED_PAGES = 3;       // 浏览 3 个页面后提示
    var DISMISS_DAYS = 30;        // 拒绝后 30 天不再提示

    var deferredPrompt = null;

    // 检查是否应该显示安装提示
    function shouldShowPrompt() {
        // 已安装
        if (window.matchMedia && window.matchMedia('(display-mode: standalone)').matches) {
            return false;
        }

        // 用户之前拒绝过，检查是否已过冷却期
        var dismissed = localStorage.getItem(DISMISS_KEY);
        if (dismissed) {
            var dismissedAt = parseInt(dismissed, 10);
            var cooldownMs = DISMISS_DAYS * 24 * 60 * 60 * 1000;
            if (Date.now() - dismissedAt < cooldownMs) {
                return false;
            }
        }

        // 检查页面浏览次数
        var pageCount = parseInt(localStorage.getItem(PAGE_COUNT_KEY) || '0', 10);
        return pageCount >= REQUIRED_PAGES;
    }

    // 递增页面浏览计数
    function incrementPageCount() {
        var count = parseInt(localStorage.getItem(PAGE_COUNT_KEY) || '0', 10);
        localStorage.setItem(PAGE_COUNT_KEY, (count + 1).toString());
    }

    // 创建安装横幅
    function createInstallBanner() {
        if (document.getElementById('pwa-install-banner')) return;

        var banner = document.createElement('div');
        banner.id = 'pwa-install-banner';
        banner.setAttribute('role', 'dialog');
        banner.setAttribute('aria-label', '安装应用');
        banner.innerHTML =
            '<div class="pwa-install-inner">' +
                '<div class="pwa-install-info">' +
                    '<div class="pwa-install-icon">' +
                        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">' +
                            '<path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2z"/>' +
                            '<path d="M8 12l2 2 4-4"/>' +
                        '</svg>' +
                    '</div>' +
                    '<div class="pwa-install-text">' +
                        '<strong>安装「香氛定制」</strong>' +
                        '<span>添加到主屏幕，获得更好的体验</span>' +
                    '</div>' +
                '</div>' +
                '<div class="pwa-install-actions">' +
                    '<button class="pwa-install-btn" id="pwaInstallBtn">安装</button>' +
                    '<button class="pwa-dismiss-btn" id="pwaDismissBtn" aria-label="关闭安装提示">&times;</button>' +
                '</div>' +
            '</div>';

        // 添加样式
        var style = document.createElement('style');
        style.textContent =
            '#pwa-install-banner{position:fixed;bottom:0;left:0;right:0;z-index:10000;' +
            'background:#fff;border-top:1px solid #e0d8d0;padding:12px 16px;' +
            'transform:translateY(100%);transition:transform 0.3s ease-out;' +
            'box-shadow:0 -2px 10px rgba(0,0,0,0.1);font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;}' +
            '#pwa-install-banner.visible{transform:translateY(0);}' +
            '.pwa-install-inner{display:flex;align-items:center;justify-content:space-between;max-width:600px;margin:0 auto;gap:12px;}' +
            '.pwa-install-info{display:flex;align-items:center;gap:12px;}' +
            '.pwa-install-icon{width:40px;height:40px;border-radius:10px;background:#8B4513;color:#fff;' +
            'display:flex;align-items:center;justify-content:center;flex-shrink:0;}' +
            '.pwa-install-icon svg{width:24px;height:24px;}' +
            '.pwa-install-text{display:flex;flex-direction:column;gap:2px;}' +
            '.pwa-install-text strong{font-size:14px;color:#1a1a1a;}' +
            '.pwa-install-text span{font-size:12px;color:#666;}' +
            '.pwa-install-actions{display:flex;align-items:center;gap:8px;}' +
            '.pwa-install-btn{padding:8px 20px;background:#8B4513;color:#fff;border:none;border-radius:6px;' +
            'font-size:14px;font-weight:600;cursor:pointer;transition:background 0.15s;}' +
            '.pwa-install-btn:hover{background:#6B3410;}' +
            '.pwa-dismiss-btn{width:28px;height:28px;border:none;background:none;color:#999;' +
            'font-size:20px;cursor:pointer;border-radius:50%;display:flex;align-items:center;justify-content:center;}' +
            '.pwa-dismiss-btn:hover{background:#f5f5f5;color:#666;}' +
            '@media(prefers-color-scheme:dark){' +
            '#pwa-install-banner{background:#1a1410;border-top-color:#3a3228;}' +
            '.pwa-install-text strong{color:#e8e0d8;}.pwa-install-text span{color:#b0a89e;}' +
            '.pwa-dismiss-btn{color:#7a7268;}.pwa-dismiss-btn:hover{background:#241e18;color:#b0a89e;}}';

        document.head.appendChild(style);
        document.body.appendChild(banner);

        // 动画显示
        requestAnimationFrame(function() {
            requestAnimationFrame(function() {
                banner.classList.add('visible');
            });
        });

        // 安装按钮
        document.getElementById('pwaInstallBtn').addEventListener('click', function() {
            if (deferredPrompt) {
                deferredPrompt.prompt();
                deferredPrompt.userChoice.then(function(choiceResult) {
                    if (choiceResult.outcome === 'accepted') {
                        localStorage.setItem(INSTALL_KEY, 'installed');
                    } else {
                        localStorage.setItem(DISMISS_KEY, Date.now().toString());
                    }
                    deferredPrompt = null;
                });
            }
            dismissBanner();
        });

        // 关闭按钮
        document.getElementById('pwaDismissBtn').addEventListener('click', function() {
            localStorage.setItem(DISMISS_KEY, Date.now().toString());
            dismissBanner();
        });
    }

    function dismissBanner() {
        var banner = document.getElementById('pwa-install-banner');
        if (banner) {
            banner.classList.remove('visible');
            setTimeout(function() {
                if (banner.parentNode) banner.parentNode.removeChild(banner);
            }, 300);
        }
    }

    // 监听 beforeinstallprompt
    window.addEventListener('beforeinstallprompt', function(e) {
        e.preventDefault();
        deferredPrompt = e;
        window.deferredPrompt = e;

        // 如果满足条件，显示安装横幅
        if (shouldShowPrompt()) {
            // 延迟 5 秒显示，避免干扰
            setTimeout(createInstallBanner, 5000);
        }
    });

    // 监听安装成功
    window.addEventListener('appinstalled', function() {
        localStorage.setItem(INSTALL_KEY, 'installed');
        deferredPrompt = null;
        dismissBanner();
        console.log('[PWA] App installed');
    });

    // 页面加载时递增浏览次数
    incrementPageCount();

})();
