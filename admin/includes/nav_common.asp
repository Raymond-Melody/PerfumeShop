<!-- V18 管理后台共享导航交互脚本 -->
<script>
(function() {
    'use strict';
    
    var hamburger = document.getElementById('adminHamburger');
    var sidebar = document.getElementById('adminSidebar');
    var overlay = document.getElementById('sidebarOverlay');
    var body = document.body;
    
    function openSidebar() {
        if (!sidebar || !overlay) return;
        sidebar.classList.add('active');
        overlay.classList.add('active');
        if (hamburger) hamburger.classList.add('active');
        body.style.overflow = 'hidden';
    }
    
    function closeSidebar() {
        if (!sidebar || !overlay) return;
        sidebar.classList.remove('active');
        overlay.classList.remove('active');
        if (hamburger) hamburger.classList.remove('active');
        body.style.overflow = '';
    }
    
    if (hamburger) hamburger.addEventListener('click', function(e) {
        e.stopPropagation();
        if (sidebar && sidebar.classList.contains('active')) {
            closeSidebar();
        } else {
            openSidebar();
        }
    });
    if (overlay) overlay.addEventListener('click', closeSidebar);
    
    if (sidebar) {
        sidebar.addEventListener('click', function(e) {
            if (e.target.tagName === 'A' || e.target.closest('a')) {
                setTimeout(closeSidebar, 150);
            }
        });
    }
    
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape' && sidebar && sidebar.classList.contains('active')) {
            closeSidebar();
        }
    });
    
    window.addEventListener('resize', function() {
        if (window.innerWidth >= 769) {
            closeSidebar();
        }
    });
    
    // 当前页面高亮 - 自动检测（仅在服务器未设置active时执行）
    (function setActiveLink() {
        // 检查服务器端是否已设置active状态，避免冲突
        var hasServerActive = document.querySelector('.sidebar-menu a.active, .sidebar-nav .nav-item.active');
        if (hasServerActive) return;
        
        var currentPath = window.location.pathname.toLowerCase().replace(/\/$/, '') || '/';
        var links = document.querySelectorAll('.sidebar-menu a, .sidebar-nav .nav-item');
        var bestMatch = null, bestScore = 0;
        
        links.forEach(function(link) {
            var href = link.getAttribute('href');
            if (!href || href === '#' || href.indexOf('javascript:') === 0) return;
            
            var hrefNorm = href.toLowerCase().replace(/\/$/, '');
            var hrefPath = hrefNorm.split('?')[0];
            
            if (currentPath === hrefPath || currentPath === hrefNorm) {
                link.classList.add('active');
                bestMatch = link; bestScore = 100;
            } else if (hrefPath !== '/' && hrefPath !== '' && currentPath.indexOf(hrefPath) === 0) {
                var score = hrefPath.length;
                if (score > bestScore) { bestMatch = link; bestScore = score; }
            }
        });
        
        if (bestMatch && bestScore < 100 && bestScore > 0) {
            bestMatch.classList.add('active');
        }
    })();
})();
</script>
