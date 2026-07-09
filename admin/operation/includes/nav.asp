<!-- 运营管理后台导航 V18 -->
<link rel="stylesheet" href="/css/admin.css">
<link rel="stylesheet" href="/css/responsive.css?v=18.0">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
<div class="admin-dashboard">
    <!-- 顶部导航栏 -->
    <nav class="admin-navbar">
        <div class="admin-nav-container">
            <div style="display:flex;align-items:center;">
                <button class="admin-hamburger" id="adminHamburger" aria-label="菜单">
                    <span></span>
                    <span></span>
                    <span></span>
                </button>
                <a href="index.asp" class="admin-nav-brand">
                    <i class="fas fa-chart-line"></i>
                    <span>运营管理中心</span>
                </a>
            </div>
            <ul class="admin-nav-menu">
                <li><a href="javascript:void(0)" onclick="location.reload()"><i class="fas fa-sync-alt"></i> 刷新</a></li>
                <li><a href="../portal.asp"><i class="fas fa-th-large"></i> 返回入口</a></li>
                <li><a href="../logout.asp"><i class="fas fa-sign-out-alt"></i> 退出</a></li>
            </ul>
        </div>
    </nav>
    
    <!-- 移动端侧边栏遮罩 -->
    <div class="sidebar-overlay" id="sidebarOverlay"></div>
    
    <!-- 侧边栏 -->
    <aside class="sidebar" id="adminSidebar">
        <div class="sidebar-header" style="padding:20px;border-bottom:1px solid rgba(255,255,255,0.06);">
            <h3 style="color:#e0e0e0;margin:0;font-size:16px;display:flex;align-items:center;gap:10px;">
                <i class="fas fa-chart-line" style="color:#667eea;"></i> 运营管理中心
            </h3>
            <p style="color:#888;font-size:11px;margin:8px 0 0;">订单 · 客户 · 商品 · 营销</p>
        </div>
        <ul class="sidebar-menu">
            <li><a href="index.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/operation/index.asp", "active", "") %>"><i class="fas fa-home"></i> <span>运营概览</span></a></li>
            
            <li class="sidebar-section-title">订单与客户</li>
            <li><a href="orders.asp" class="<%= IIf(InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "/admin/operation/orders") > 0, "active", "") %>"><i class="fas fa-shopping-cart"></i> <span>订单管理</span></a></li>
            <li><a href="customers.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/operation/customers.asp", "active", "") %>"><i class="fas fa-users"></i> <span>客户管理</span></a></li>
            <li><a href="points.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/operation/points.asp", "active", "") %>"><i class="fas fa-coins"></i> <span>积分管理</span></a></li>
            
            <li class="sidebar-section-title">售后与评价</li>
            <li><a href="after_sales.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/operation/after_sales.asp", "active", "") %>"><i class="fas fa-headset"></i> <span>售后管理</span></a></li>
            <li><a href="reviews_manage.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/operation/reviews_manage.asp", "active", "") %>"><i class="fas fa-star"></i> <span>评价审核</span></a></li>

            <li class="sidebar-section-title">商品与内容</li>
            <li><a href="products.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/operation/products.asp", "active", "") %>"><i class="fas fa-box"></i> <span>商品上下架</span></a></li>
            <li><a href="content_pages.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/operation/content_pages.asp", "active", "") %>"><i class="fas fa-file-alt"></i> <span>内容页面</span></a></li>
            
            <li class="sidebar-section-title">营销与推广</li>
            <li><a href="order_reviews.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/operation/order_reviews.asp", "active", "") %>"><i class="fas fa-star-half-alt"></i> <span>评价管理</span></a></li>
            <li><a href="marketing.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/operation/marketing.asp", "active", "") %>"><i class="fas fa-bullhorn"></i> <span>营销活动</span></a></li>
            <li><a href="coupon_management.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/operation/coupon_management.asp", "active", "") %>"><i class="fas fa-ticket-alt"></i> <span>优惠券管理</span></a></li>
            <li><a href="tier_management.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/operation/tier_management.asp", "active", "") %>"><i class="fas fa-layer-group"></i> <span>会员等级</span></a></li>
            <li><a href="flash_sale.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/operation/flash_sale.asp", "active", "") %>"><i class="fas fa-bolt"></i> <span>秒杀管理</span></a></li>
            <li><a href="group_buy.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/operation/group_buy.asp", "active", "") %>"><i class="fas fa-users"></i> <span>拼团管理</span></a></li>
            <li><a href="subscription_plans.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/operation/subscription_plans.asp", "active", "") %>"><i class="fas fa-box-open"></i> <span>订阅管理</span></a></li>
            <li><a href="referral_codes.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/operation/referral_codes.asp", "active", "") %>"><i class="fas fa-user-friends"></i> <span>推荐码管理</span></a></li>
            <li><a href="payment_switch.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/operation/payment_switch.asp", "active", "") %>"><i class="fas fa-toggle-on"></i> <span>支付开关</span></a></li>
        </ul>
        <div class="sidebar-footer" style="padding:15px 20px;border-top:1px solid rgba(255,255,255,0.06);">
            <a href="/admin/portal.asp" style="color:#666;text-decoration:none;font-size:13px;display:flex;align-items:center;gap:8px;">
                <i class="fas fa-arrow-left"></i> 返回管理中心
            </a>
        </div>
    </aside>
</div>
<!--#include file="../../includes/nav_common.asp"-->
<style>
/* 侧边栏深色主题增强 */
.sidebar-section-title {
    display: block;
    padding: 15px 20px 5px;
    font-size: 11px;
    color: #888;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    border-bottom: 1px solid rgba(255,255,255,0.06);
    margin: 5px 10px 5px;
}
.sidebar-menu li a {
    display: flex;
    align-items: center;
    padding: 10px 15px;
    color: #b0b0b0;
    text-decoration: none;
    transition: all 0.2s ease;
    border-radius: 4px;
    margin: 2px 10px;
}
.sidebar-menu li a:hover {
    background: rgba(255,255,255,0.05);
    color: #fff;
}
.sidebar-menu li a.active {
    background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%);
    color: #fff;
}
.sidebar-menu li a i {
    width: 20px;
    margin-right: 10px;
    text-align: center;
}
</style>
