<!-- 产品生产管理中心侧边导航 V18 -->
<link rel="stylesheet" href="/css/admin.css">
<link rel="stylesheet" href="/css/responsive.css?v=18.0">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
<%
Dim navFrom : navFrom = Request.QueryString("from")
If navFrom = "inventory" Then
    Server.Execute("/admin/inventory/includes/nav.asp")
Else
    Dim prodCurrentPage
    prodCurrentPage = Request.ServerVariables("SCRIPT_NAME")
    prodCurrentPage = Mid(prodCurrentPage, InStrRev(prodCurrentPage, "/") + 1)
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
                    <i class="fas fa-industry"></i>
                    <span>产品生产中心</span>
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
                <i class="fas fa-industry" style="color:#4CAF50;"></i> 产品生产中心
            </h3>
            <p style="color:#888;font-size:11px;margin:8px 0 0;">香调 → 成品</p>
        </div>
        <ul class="sidebar-menu">
            <li class="sidebar-section-title">总览</li>
            <li><a href="index.asp" class="<%= IIf(prodCurrentPage="index.asp","active","") %>"><i class="fas fa-tachometer-alt"></i> <span>生产概览</span></a></li>
            <li class="sidebar-section-title">生产管理</li>
            <li><a href="production_management.asp" class="<%= IIf(prodCurrentPage="production_management.asp","active","") %>"><i class="fas fa-clipboard-list"></i> <span>生产工单管理</span></a></li>
            <li><a href="prod_scheduling.asp" class="<%= IIf(prodCurrentPage="prod_scheduling.asp","active","") %>"><i class="fas fa-calendar-alt"></i> <span>排产调度</span></a></li>
            <li><a href="prod_workshop.asp" class="<%= IIf(prodCurrentPage="prod_workshop.asp","active","") %>"><i class="fas fa-cogs"></i> <span>车间作业</span></a></li>
            <li><a href="prod_qc.asp" class="<%= IIf(prodCurrentPage="prod_qc.asp","active","") %>"><i class="fas fa-check-circle"></i> <span>质量检验</span></a></li>
            <li><a href="prod_warehouse.asp" class="<%= IIf(prodCurrentPage="prod_warehouse.asp","active","") %>"><i class="fas fa-warehouse"></i> <span>成品入库</span></a></li>
            <li><a href="order_production.asp" class="<%= IIf(prodCurrentPage="order_production.asp","active","") %>"><i class="fas fa-shopping-cart"></i> <span>订单级生产追踪</span></a></li>
            <li class="sidebar-section-title">库存管理</li>
            <li><a href="product_inventory.asp" class="<%= IIf(prodCurrentPage="product_inventory.asp","active","") %>"><i class="fas fa-box"></i> <span>成品库存</span></a></li>
            <li><a href="bottle_inventory.asp" class="<%= IIf(prodCurrentPage="bottle_inventory.asp","active","") %>"><i class="fas fa-flask"></i> <span>瓶子库存</span></a></li>
            <li><a href="packaging_inventory.asp" class="<%= IIf(prodCurrentPage="packaging_inventory.asp","active","") %>"><i class="fas fa-box-open"></i> <span>包装物库存</span></a></li>
        </ul>
        <div class="sidebar-footer" style="padding:15px 20px;border-top:1px solid rgba(255,255,255,0.06);">
            <a href="/admin/portal.asp" style="color:#666;text-decoration:none;font-size:13px;display:flex;align-items:center;gap:8px;">
                <i class="fas fa-arrow-left"></i> 返回管理中心
            </a>
        </div>
    </aside>
</div>
<!--#include file="../../includes/nav_common.asp"-->
<% End If %>
