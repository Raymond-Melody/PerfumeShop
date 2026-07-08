<!-- 物流管理中心侧边导航 V18 -->
<link rel="stylesheet" href="/css/responsive.css?v=18.0">
<%
Dim logCurrentPage
logCurrentPage = Request.ServerVariables("SCRIPT_NAME")
logCurrentPage = Mid(logCurrentPage, InStrRev(logCurrentPage, "/") + 1)
%>
<!-- V18 移动端顶部栏 -->
<div class="admin-mobile-topbar" style="display:none;position:fixed;top:0;left:0;right:0;height:48px;background:#1a1a2e;z-index:999;align-items:center;padding:0 12px;border-bottom:1px solid rgba(255,255,255,0.06);">
    <button class="admin-hamburger" id="adminHamburger" aria-label="菜单">
        <span></span><span></span><span></span>
    </button>
    <span style="color:#e0e0e0;font-size:14px;font-weight:600;margin-left:12px;">物流管理中心</span>
</div>
<!-- 移动端侧边栏遮罩 -->
<div class="sidebar-overlay" id="sidebarOverlay"></div>

<div class="admin-sidebar" id="adminSidebar" style="position:fixed;left:0;top:60px;width:250px;height:calc(100vh - 60px);background:#1a1a2e;border-right:1px solid rgba(255,255,255,0.06);overflow-y:auto;z-index:100;">
    <div class="sidebar-header" style="padding:20px;border-bottom:1px solid rgba(255,255,255,0.06);">
        <h3 style="color:#e0e0e0;margin:0;font-size:16px;display:flex;align-items:center;gap:10px;">
            <i class="fas fa-truck" style="color:#FF9800;"></i> 物流管理中心
        </h3>
        <p style="color:#888;font-size:11px;margin:8px 0 0;">发货 → 在途 → 签收</p>
    </div>
    
    <nav class="sidebar-nav" style="padding:15px 0;">
        <div class="nav-section">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">总览</div>
            <a href="index.asp" class="nav-item <%= IIf(logCurrentPage="index.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-tachometer-alt"></i> 物流概览
            </a>
        </div>
        
        <div class="nav-section" style="margin-top:15px;">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">发货管理</div>
            <a href="shipping_orders.asp" class="nav-item <%= IIf(logCurrentPage="shipping_orders.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-clipboard-check"></i> 待发货订单
            </a>
            <a href="shipments.asp" class="nav-item <%= IIf(logCurrentPage="shipments.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-shipping-fast"></i> 发货单管理
            </a>
        </div>
        
        <div class="nav-section" style="margin-top:15px;">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">物流追踪</div>
            <a href="in_transit.asp" class="nav-item <%= IIf(logCurrentPage="in_transit.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-route"></i> 在途跟踪
            </a>
            <a href="delivery_confirm.asp" class="nav-item <%= IIf(logCurrentPage="delivery_confirm.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-check-double"></i> 签收确认
            </a>
            <a href="returns.asp" class="nav-item <%= IIf(logCurrentPage="returns.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-undo-alt"></i> 退货入库
            </a>
        </div>
        
        <div class="nav-section" style="margin-top:15px;">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">配置</div>
            <a href="shipping_companies.asp" class="nav-item <%= IIf(logCurrentPage="shipping_companies.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-building"></i> 物流公司
            </a>
            <a href="shipping_cost.asp" class="nav-item <%= IIf(logCurrentPage="shipping_cost.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-calculator"></i> 运费管理
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
.admin-sidebar .nav-item:hover { background: rgba(255,255,255,0.05); color: #e0e0e0; border-left-color: #FF9800; }
.admin-sidebar .nav-item.active { background: rgba(255,152,0,0.12); color: #FF9800; border-left-color: #FF9800; }
/* V18 移动端body偏移适配顶部栏 */
@media (max-width: 768px) {
    body { padding-top: 48px !important; }
}
</style>
<!--#include file="../../includes/nav_common.asp"-->
