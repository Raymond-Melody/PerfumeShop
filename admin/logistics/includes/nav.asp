<!-- 物流管理中心侧边导航 V18 -->
<link rel="stylesheet" href="/css/admin.css">
<link rel="stylesheet" href="/css/responsive.css?v=18.0">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
<%
Dim logCurrentPage
logCurrentPage = Request.ServerVariables("SCRIPT_NAME")
logCurrentPage = Mid(logCurrentPage, InStrRev(logCurrentPage, "/") + 1)
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
                    <i class="fas fa-truck"></i>
                    <span>物流管理中心</span>
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
                <i class="fas fa-truck" style="color:#FF9800;"></i> 物流管理中心
            </h3>
            <p style="color:#888;font-size:11px;margin:8px 0 0;">发货 → 在途 → 签收</p>
        </div>
        <ul class="sidebar-menu">
            <li class="sidebar-section-title">总览</li>
            <li><a href="index.asp" class="<%= IIf(logCurrentPage="index.asp","active","") %>"><i class="fas fa-tachometer-alt"></i> <span>物流概览</span></a></li>
            <li class="sidebar-section-title">发货管理</li>
            <li><a href="shipping_orders.asp" class="<%= IIf(logCurrentPage="shipping_orders.asp","active","") %>"><i class="fas fa-clipboard-check"></i> <span>待发货订单</span></a></li>
            <li><a href="shipments.asp" class="<%= IIf(logCurrentPage="shipments.asp","active","") %>"><i class="fas fa-shipping-fast"></i> <span>发货单管理</span></a></li>
            <li class="sidebar-section-title">物流追踪</li>
            <li><a href="in_transit.asp" class="<%= IIf(logCurrentPage="in_transit.asp","active","") %>"><i class="fas fa-route"></i> <span>在途跟踪</span></a></li>
            <li><a href="delivery_confirm.asp" class="<%= IIf(logCurrentPage="delivery_confirm.asp","active","") %>"><i class="fas fa-check-double"></i> <span>签收确认</span></a></li>
            <li><a href="returns.asp" class="<%= IIf(logCurrentPage="returns.asp","active","") %>"><i class="fas fa-undo-alt"></i> <span>退货入库</span></a></li>
            <li class="sidebar-section-title">配置</li>
            <li><a href="shipping_companies.asp" class="<%= IIf(logCurrentPage="shipping_companies.asp","active","") %>"><i class="fas fa-building"></i> <span>物流公司</span></a></li>
            <li><a href="shipping_cost.asp" class="<%= IIf(logCurrentPage="shipping_cost.asp","active","") %>"><i class="fas fa-calculator"></i> <span>运费管理</span></a></li>
        </ul>
        <div class="sidebar-footer" style="padding:15px 20px;border-top:1px solid rgba(255,255,255,0.06);">
            <a href="/admin/portal.asp" style="color:#666;text-decoration:none;font-size:13px;display:flex;align-items:center;gap:8px;">
                <i class="fas fa-arrow-left"></i> 返回管理中心
            </a>
        </div>
    </aside>
</div>
<!--#include file="../../includes/nav_common.asp"-->
