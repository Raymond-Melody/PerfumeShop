<!-- 半成品生产中心侧边导航 V18 -->
<link rel="stylesheet" href="/css/admin.css">
<link rel="stylesheet" href="/css/responsive.css?v=18.0">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
<%
Dim navFrom : navFrom = Request.QueryString("from")
If navFrom = "inventory" Then
    Server.Execute("/admin/inventory/includes/nav.asp")
Else
    Dim semiCurrentPage
    semiCurrentPage = Request.ServerVariables("SCRIPT_NAME")
    Dim scriptParts : scriptParts = Split(CStr(semiCurrentPage), "/")
    semiCurrentPage = scriptParts(UBound(scriptParts))
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
                    <i class="fas fa-vial"></i>
                    <span>半成品生产中心</span>
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
                <i class="fas fa-flask" style="color:#2196F3;"></i> 半成品生产中心
            </h3>
            <p style="color:#888;font-size:11px;margin:8px 0 0;">原料 → 基香 → 香调</p>
        </div>
        <ul class="sidebar-menu">
            <li class="sidebar-section-title">总览</li>
            <li><a href="index.asp" class="<%= IIf(semiCurrentPage="index.asp","active","") %>"><i class="fas fa-tachometer-alt"></i> <span>生产概览</span></a></li>
            <li class="sidebar-section-title">生产管理</li>
            <li><a href="accord_production.asp" class="<%= IIf(semiCurrentPage="accord_production.asp","active","") %>"><i class="fas fa-cogs"></i> <span>Accord生产</span></a></li>
            <li><a href="base_note_production.asp" class="<%= IIf(semiCurrentPage="base_note_production.asp","active","") %>"><i class="fas fa-vial"></i> <span>基香生产</span></a></li>
            <li><a href="workshop_transfer.asp" class="<%= IIf(semiCurrentPage="workshop_transfer.asp","active","") %>"><i class="fas fa-exchange-alt"></i> <span>车间调拨</span></a></li>
            <li class="sidebar-section-title">库存管理</li>
            <li><a href="inventory_dashboard.asp" class="<%= IIf(semiCurrentPage="inventory_dashboard.asp","active","") %>"><i class="fas fa-chart-pie"></i> <span>库存仪表盘</span></a></li>
            <li><a href="raw_material_inventory.asp" class="<%= IIf(semiCurrentPage="raw_material_inventory.asp","active","") %>"><i class="fas fa-boxes"></i> <span>原料库存</span></a></li>
            <li><a href="base_note_inventory.asp" class="<%= IIf(semiCurrentPage="base_note_inventory.asp","active","") %>"><i class="fas fa-database"></i> <span>基香库存</span></a></li>
            <li><a href="note_inventory.asp" class="<%= IIf(semiCurrentPage="note_inventory.asp","active","") %>"><i class="fas fa-layer-group"></i> <span>香调库存</span></a></li>
            <li><a href="material_outbound.asp" class="<%= IIf(semiCurrentPage="material_outbound.asp","active","") %>"><i class="fas fa-truck-loading"></i> <span>原料出库</span></a></li>
            <li class="sidebar-section-title">预警</li>
            <li><a href="inventory_alerts.asp" class="<%= IIf(semiCurrentPage="inventory_alerts.asp","active","") %>"><i class="fas fa-exclamation-triangle"></i> <span>库存预警</span></a></li>
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
