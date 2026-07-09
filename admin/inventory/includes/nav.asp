<!-- 库存管理中心侧边导航 V18 -->
<link rel="stylesheet" href="/css/admin.css">
<link rel="stylesheet" href="/css/responsive.css?v=18.0">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
<%
Dim invCurrentPage
invCurrentPage = Request.ServerVariables("SCRIPT_NAME")
Dim invParts : invParts = Split(CStr(invCurrentPage), "/")
invCurrentPage = invParts(UBound(invParts))

' 计算导航项active状态（避免依赖外部IIf函数）
Function NavActive(checkPage)
    If invCurrentPage = checkPage Then NavActive = "active" Else NavActive = ""
End Function
%>
<div class="admin-dashboard">
    <!-- V18 桌面端顶部导航栏 -->
    <nav class="admin-navbar">
        <div class="admin-nav-container">
            <div style="display:flex;align-items:center;">
                <button class="admin-hamburger" id="adminHamburger" aria-label="菜单">
                    <span></span><span></span><span></span>
                </button>
                <a href="index.asp" class="admin-nav-brand">
                    <i class="fas fa-warehouse"></i>
                    <span>库存管理中心</span>
                </a>
            </div>
            <ul class="admin-nav-menu">
                <li><a href="javascript:void(0)" onclick="location.reload()"><i class="fas fa-sync-alt"></i> 刷新</a></li>
                <li><a href="/admin/portal.asp"><i class="fas fa-th-large"></i> 返回入口</a></li>
                <li><a href="/admin/logout.asp"><i class="fas fa-sign-out-alt"></i> 退出</a></li>
            </ul>
        </div>
    </nav>

    <!-- 移动端侧边栏遮罩 -->
    <div class="sidebar-overlay" id="sidebarOverlay"></div>

    <!-- V18 侧边栏导航 -->
    <aside class="sidebar" id="adminSidebar">
        <div class="sidebar-header" style="padding:20px;border-bottom:1px solid rgba(255,255,255,0.06);">
            <h3 style="color:#e0e0e0;margin:0;font-size:16px;display:flex;align-items:center;gap:10px;">
                <i class="fas fa-warehouse" style="color:#00BCD4;"></i> 库存管理中心
            </h3>
            <p style="color:#888;font-size:11px;margin:8px 0 0;">全品类 · 统一管控</p>
        </div>
        <ul class="sidebar-menu">
            <li class="sidebar-section-title">总览</li>
            <li><a href="/admin/inventory/index.asp" class="<%= NavActive("index.asp") %>"><i class="fas fa-chart-pie"></i> <span>库存仪表盘</span></a></li>
            <li class="sidebar-section-title">成品库存</li>
            <li><a href="/admin/prodcenter/product_inventory.asp?from=inventory"><i class="fas fa-box"></i> <span>成品库存</span></a></li>
            <li><a href="/admin/prodcenter/bottle_inventory.asp?from=inventory"><i class="fas fa-flask"></i> <span>瓶子库存</span></a></li>
            <li><a href="/admin/prodcenter/packaging_inventory.asp?from=inventory"><i class="fas fa-box-open"></i> <span>包装物库存</span></a></li>
            <li class="sidebar-section-title">半成品库存</li>
            <li><a href="/admin/semifinished/raw_material_inventory.asp?from=inventory"><i class="fas fa-boxes"></i> <span>原料库存</span></a></li>
            <li><a href="/admin/semifinished/base_note_inventory.asp?from=inventory"><i class="fas fa-database"></i> <span>基香库存</span></a></li>
            <li><a href="/admin/semifinished/note_inventory.asp?from=inventory"><i class="fas fa-layer-group"></i> <span>香调库存</span></a></li>
            <li class="sidebar-section-title">库存运营</li>
            <li><a href="/admin/inventory/stock_movements.asp" class="<%= NavActive("stock_movements.asp") %>"><i class="fas fa-history"></i> <span>库存流水</span></a></li>
            <li class="sidebar-section-title">预警</li>
            <li><a href="/admin/inventory/inventory_alerts.asp" class="<%= NavActive("inventory_alerts.asp") %>"><i class="fas fa-exclamation-triangle"></i> <span>库存预警</span></a></li>
        </ul>
        <div class="sidebar-footer" style="padding:15px 20px;border-top:1px solid rgba(255,255,255,0.06);">
            <a href="/admin/portal.asp" style="color:#666;text-decoration:none;font-size:13px;display:flex;align-items:center;gap:8px;">
                <i class="fas fa-arrow-left"></i> 返回管理中心
            </a>
        </div>
    </aside>
</div>
<!--#include file="../../includes/nav_common.asp"-->
