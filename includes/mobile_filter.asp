<%
' ============================================
' V11 移动端侧边栏切换组件
' 在移动端隐藏侧边栏，添加筛选按钮
' ============================================
%>
<!-- 移动端筛选按钮（仅在小屏显示） -->
<div class="mobile-filter-toggle hide-desktop" id="mobileFilterToggle">
    <button class="btn btn-outline touch-target" onclick="toggleMobileFilter()">
        <i class="fas fa-filter"></i> 筛选
    </button>
</div>

<script>
function toggleMobileFilter() {
    var sidebar = document.querySelector('.sidebar');
    if (sidebar) {
        sidebar.classList.toggle('mobile-active');
    }
}
</script>

<style>
@media (max-width: 991px) {
    .sidebar {
        position: fixed;
        top: 0;
        left: calc(-1 * 280px);
        width: 280px;
        height: 100vh;
        height: 100dvh;
        z-index: 1003;
        transition: left 0.3s ease;
        overflow-y: auto;
        background: #fff;
        padding: 20px;
        box-shadow: 2px 0 12px rgba(0,0,0,0.15);
    }
    
    .sidebar.mobile-active {
        left: 0;
    }
    
    .sidebar-overlay {
        display: none;
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: rgba(0,0,0,0.4);
        z-index: 1002;
    }
    
    .sidebar-overlay.active {
        display: block;
    }
    
    .mobile-filter-toggle {
        position: fixed;
        bottom: 70px;
        right: 16px;
        z-index: 998;
    }
    
    .mobile-filter-toggle .btn {
        border-radius: 50%;
        width: 56px;
        height: 56px;
        padding: 0;
        display: flex;
        align-items: center;
        justify-content: center;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
    }
}
</style>
