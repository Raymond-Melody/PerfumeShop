/**
 * PerfumeShop V14.6 - 筛选优化
 * 筛选结果计数 + 一键清空 + 筛选面包屑 + 移动端优化
 */

(function() {
    'use strict';
    
    // 初始化筛选优化
    function initFilterOptimization() {
        var filterContainer = document.querySelector('.filter-sidebar, .filter-panel');
        if (!filterContainer) return;
        
        // 1. 添加筛选结果计数
        addFilterCount(filterContainer);
        
        // 2. 添加一键清空按钮
        addClearAllButton(filterContainer);
        
        // 3. 添加筛选面包屑
        addFilterBreadcrumbs(filterContainer);
        
        // 4. 监听筛选变化
        watchFilterChanges(filterContainer);
    }
    
    // 添加筛选结果计数
    function addFilterCount(container) {
        var header = container.querySelector('.filter-header');
        if (!header) return;
        
        var countDiv = document.createElement('div');
        countDiv.className = 'filter-count';
        countDiv.innerHTML = '已选 <span class="count-badge">0</span> 个筛选条件';
        
        header.appendChild(countDiv);
        updateFilterCount(container);
    }
    
    // 更新筛选计数
    function updateFilterCount(container) {
        var checkedBoxes = container.querySelectorAll('input[type="checkbox"]:checked, input[type="radio"]:checked');
        var countBadge = container.querySelector('.count-badge');
        
        if (countBadge) {
            countBadge.textContent = checkedBoxes.length;
        }
        
        // 更新清空按钮可见性
        var clearBtn = container.querySelector('.btn-clear-all');
        if (clearBtn) {
            clearBtn.style.display = checkedBoxes.length > 0 ? 'inline-block' : 'none';
        }
        
        // 更新面包屑
        updateFilterBreadcrumbs(container);
    }
    
    // 添加一键清空按钮
    function addClearAllButton(container) {
        var header = container.querySelector('.filter-header');
        if (!header) return;
        
        var clearBtn = document.createElement('button');
        clearBtn.className = 'btn-clear-all';
        clearBtn.innerHTML = '<i class="fas fa-times-circle"></i> 清空筛选';
        clearBtn.style.display = 'none'; // 初始隐藏
        
        clearBtn.addEventListener('click', function() {
            if (confirm('确定要清空所有筛选条件吗？')) {
                clearAllFilters(container);
            }
        });
        
        header.appendChild(clearBtn);
    }
    
    // 清空所有筛选
    function clearAllFilters(container) {
        var checkboxes = container.querySelectorAll('input[type="checkbox"]:checked');
        var radios = container.querySelectorAll('input[type="radio"]:checked');
        var selects = container.querySelectorAll('select');
        
        checkboxes.forEach(function(checkbox) {
            checkbox.checked = false;
        });
        
        radios.forEach(function(radio) {
            radio.checked = false;
        });
        
        selects.forEach(function(select) {
            select.selectedIndex = 0;
        });
        
        // 触发筛选更新
        updateFilterCount(container);
        
        // 重新加载产品列表（无筛选）
        if (typeof refreshProducts === 'function') {
            refreshProducts();
        } else {
            window.location.href = window.location.pathname;
        }
        
        // 显示成功提示
        showToast('已清空所有筛选条件');
    }
    
    // 添加筛选面包屑
    function addFilterBreadcrumbs(container) {
        var header = container.querySelector('.filter-header');
        if (!header) return;
        
        var breadcrumbs = document.createElement('div');
        breadcrumbs.className = 'filter-breadcrumbs';
        breadcrumbs.style.display = 'none';
        
        header.appendChild(breadcrumbs);
    }
    
    // 更新筛选面包屑
    function updateFilterBreadcrumbs(container) {
        var breadcrumbs = container.querySelector('.filter-breadcrumbs');
        if (!breadcrumbs) return;
        
        var checkedBoxes = container.querySelectorAll('input[type="checkbox"]:checked, input[type="radio"]:checked');
        
        if (checkedBoxes.length === 0) {
            breadcrumbs.style.display = 'none';
            return;
        }
        
        breadcrumbs.style.display = 'block';
        breadcrumbs.innerHTML = '';
        
        checkedBoxes.forEach(function(checkbox) {
            var label = findLabelForInput(checkbox);
            if (!label) return;
            
            var tag = document.createElement('span');
            tag.className = 'filter-tag';
            tag.innerHTML = label.textContent + ' <i class="fas fa-times"></i>';
            
            tag.addEventListener('click', function() {
                checkbox.checked = false;
                updateFilterCount(container);
                
                if (typeof refreshProducts === 'function') {
                    refreshProducts();
                }
                
                showToast('已移除筛选：' + label.textContent);
            });
            
            breadcrumbs.appendChild(tag);
        });
    }
    
    // 查找input对应的label
    function findLabelForInput(input) {
        // 方法1: label包裹input
        var label = input.closest('label');
        if (label) return label;
        
        // 方法2: for属性关联
        var id = input.id;
        if (id) {
            label = document.querySelector('label[for="' + id + '"]');
            if (label) return label;
        }
        
        // 方法3: 兄弟元素
        var siblings = input.parentElement.children;
        for (var i = 0; i < siblings.length; i++) {
            if (siblings[i].tagName === 'LABEL') {
                return siblings[i];
            }
        }
        
        return null;
    }
    
    // 监听筛选变化
    function watchFilterChanges(container) {
        var inputs = container.querySelectorAll('input[type="checkbox"], input[type="radio"], select');
        
        inputs.forEach(function(input) {
            var event = input.tagName === 'SELECT' ? 'change' : 'click';
            
            input.addEventListener(event, function() {
                setTimeout(function() {
                    updateFilterCount(container);
                }, 100);
            });
        });
    }
    
    // 显示提示消息
    function showToast(message) {
        var toast = document.createElement('div');
        toast.className = 'filter-toast';
        toast.textContent = message;
        toast.style.cssText = 'position:fixed;bottom:80px;left:50%;transform:translateX(-50%);' +
            'background:rgba(0,0,0,0.8);color:#fff;padding:10px 20px;border-radius:6px;' +
            'z-index:9999;animation:slideInUp 0.3s ease;';
        
        document.body.appendChild(toast);
        
        setTimeout(function() {
            toast.style.opacity = '0';
            toast.style.transition = 'opacity 0.3s';
            setTimeout(function() {
                document.body.removeChild(toast);
            }, 300);
        }, 2000);
    }
    
    // 页面加载完成后初始化
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initFilterOptimization);
    } else {
        initFilterOptimization();
    }
    
    // 暴露全局方法
    window.updateFilterCount = updateFilterCount;
    window.clearAllFilters = clearAllFilters;
})();
