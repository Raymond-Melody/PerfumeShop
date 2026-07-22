/**
 * PerfumeShop V14.6 - 主题切换（暗黑模式）
 * 支持手动切换 + 系统偏好自动检测 + localStorage记忆
 */

(function() {
    'use strict';
    
    var THEME_KEY = 'perfumeshop_theme';
    var THEME_DARK = 'dark';
    var THEME_LIGHT = 'light';
    
    // 获取当前主题
    function getCurrentTheme() {
        // 1. 优先使用localStorage
        var stored = localStorage.getItem(THEME_KEY);
        if (stored) {
            return stored;
        }
        
        // 2. 检测系统偏好
        if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
            return THEME_DARK;
        }
        
        // 3. 默认亮色
        return THEME_LIGHT;
    }
    
    // 应用主题
    function applyTheme(theme) {
        document.documentElement.setAttribute('data-theme', theme);
        localStorage.setItem(THEME_KEY, theme);
        
        // 更新按钮图标
        var icon = document.querySelector('.theme-toggle i');
        if (icon) {
            icon.className = theme === THEME_DARK ? 'fas fa-sun' : 'fas fa-moon';
        }
        
        console.log('[Theme] Applied:', theme);
    }
    
    // 切换主题
    function toggleTheme() {
        var current = getCurrentTheme();
        var next = current === THEME_DARK ? THEME_LIGHT : THEME_DARK;
        applyTheme(next);
    }
    
    // 创建主题切换按钮
    function createThemeToggle() {
        // 检查是否已存在
        if (document.querySelector('.theme-toggle')) {
            return;
        }
        
        var btn = document.createElement('button');
        btn.className = 'theme-toggle touch-target';
        btn.setAttribute('aria-label', '切换主题');
        btn.title = '切换明暗主题';
        
        var icon = document.createElement('i');
        icon.className = getCurrentTheme() === THEME_DARK ? 'fas fa-sun' : 'fas fa-moon';
        
        btn.appendChild(icon);
        btn.addEventListener('click', toggleTheme);
        
        document.body.appendChild(btn);
        
        console.log('[Theme] Toggle button created');
    }
    
    // 监听系统主题变化
    function watchSystemTheme() {
        if (!window.matchMedia) return;
        
        var mediaQuery = window.matchMedia('(prefers-color-scheme: dark)');
        
        // 现代浏览器
        if (mediaQuery.addEventListener) {
            mediaQuery.addEventListener('change', function(e) {
                // 只在用户没有手动设置时跟随系统
                if (!localStorage.getItem(THEME_KEY)) {
                    applyTheme(e.matches ? THEME_DARK : THEME_LIGHT);
                }
            });
        }
    }
    
    // 初始化
    function init() {
        var theme = getCurrentTheme();
        applyTheme(theme);
        createThemeToggle();
        watchSystemTheme();
    }
    
    // 页面加载完成后初始化
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
    
    // 暴露全局方法
    window.setTheme = applyTheme;
    window.getTheme = getCurrentTheme;
    window.toggleTheme = toggleTheme;
})();
