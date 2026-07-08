<!-- 半成品生产中心侧边导航 V18 -->
<link rel="stylesheet" href="/css/responsive.css?v=18.0">
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
<!-- V18 移动端顶部栏 -->
<div class="admin-mobile-topbar" style="display:none;position:fixed;top:0;left:0;right:0;height:48px;background:#1a1a2e;z-index:999;align-items:center;padding:0 12px;border-bottom:1px solid rgba(255,255,255,0.06);">
    <button class="admin-hamburger" id="adminHamburger" aria-label="菜单">
        <span></span><span></span><span></span>
    </button>
    <span style="color:#e0e0e0;font-size:14px;font-weight:600;margin-left:12px;">半成品生产中心</span>
</div>
<!-- 移动端侧边栏遮罩 -->
<div class="sidebar-overlay" id="sidebarOverlay"></div>

<div class="admin-sidebar" id="adminSidebar" style="position:fixed;left:0;top:60px;width:250px;height:calc(100vh - 60px);background:#1a1a2e;border-right:1px solid rgba(255,255,255,0.06);overflow-y:auto;z-index:100;">
    <div class="sidebar-header" style="padding:20px;border-bottom:1px solid rgba(255,255,255,0.06);">
        <h3 style="color:#e0e0e0;margin:0;font-size:16px;display:flex;align-items:center;gap:10px;">
            <i class="fas fa-flask" style="color:#2196F3;"></i> 半成品生产中心
        </h3>
        <p style="color:#888;font-size:11px;margin:8px 0 0;">原料 → 基香 → 香调</p>
    </div>
    
    <nav class="sidebar-nav" style="padding:15px 0;">
        <div class="nav-section">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">总览</div>
            <a href="index.asp" class="nav-item <%= IIf(semiCurrentPage="index.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-tachometer-alt"></i> 生产概览
            </a>
        </div>
        
        <div class="nav-section" style="margin-top:15px;">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">生产管理</div>
            <a href="accord_production.asp" class="nav-item <%= IIf(semiCurrentPage="accord_production.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-cogs"></i> Accord生产
            </a>
            <a href="base_note_production.asp" class="nav-item <%= IIf(semiCurrentPage="base_note_production.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-vial"></i> 基香生产
            </a>
            <a href="workshop_transfer.asp" class="nav-item <%= IIf(semiCurrentPage="workshop_transfer.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-exchange-alt"></i> 车间调拨
            </a>
        </div>
        
        <div class="nav-section" style="margin-top:15px;">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">库存管理</div>
            <a href="inventory_dashboard.asp" class="nav-item <%= IIf(semiCurrentPage="inventory_dashboard.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-chart-pie"></i> 库存仪表盘
            </a>
            <a href="raw_material_inventory.asp" class="nav-item <%= IIf(semiCurrentPage="raw_material_inventory.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-boxes"></i> 原料库存
            </a>
            <a href="base_note_inventory.asp" class="nav-item <%= IIf(semiCurrentPage="base_note_inventory.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-database"></i> 基香库存
            </a>
            <a href="note_inventory.asp" class="nav-item <%= IIf(semiCurrentPage="note_inventory.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-layer-group"></i> 香调库存
            </a>
            <a href="material_outbound.asp" class="nav-item <%= IIf(semiCurrentPage="material_outbound.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-truck-loading"></i> 原料出库
            </a>
        </div>
        
        <div class="nav-section" style="margin-top:15px;">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">预警</div>
            <a href="inventory_alerts.asp" class="nav-item <%= IIf(semiCurrentPage="inventory_alerts.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-exclamation-triangle"></i> 库存预警
            </a>
        </div>
    </nav>
    
    <div class="sidebar-footer" style="position:sticky;bottom:0;padding:15px 20px;border-top:1px solid rgba(255,255,255,0.06);background:#1a1a2e;">
        <a href="/admin/portal.asp" style="color:#888;text-decoration:none;font-size:13px;display:flex;align-items:center;gap:8px;">
            <i class="fas fa-arrow-left"></i> 返回管理中心
        </a>
    </div>
</div>

<style>
.admin-sidebar .nav-item:hover { background: rgba(255,255,255,0.05); color: #e0e0e0; border-left-color: #2196F3; }
.admin-sidebar .nav-item.active { background: rgba(33,150,243,0.12); color: #2196F3; border-left-color: #2196F3; }
/* V18 移动端body偏移适配顶部栏 */
@media (max-width: 768px) {
    body { padding-top: 48px !important; }
}
</style>
<!--#include file="../../includes/nav_common.asp"-->
<% End If %>
