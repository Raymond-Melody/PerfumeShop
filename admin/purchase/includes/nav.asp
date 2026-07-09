<!-- 采购管理中心侧边导航 V18 -->
<link rel="stylesheet" href="/css/admin.css">
<link rel="stylesheet" href="/css/responsive.css?v=18.0">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
<%
Dim purchaseFullPath : purchaseFullPath = Request.ServerVariables("SCRIPT_NAME")
Dim purParts : purParts = Split(CStr(purchaseFullPath), "/")
Dim purchaseCurrentPage : purchaseCurrentPage = purParts(UBound(purParts))
Dim isFixedBrand : isFixedBrand = False
If InStr(LCase(purchaseFullPath), "/fixed_brand/") > 0 Then isFixedBrand = True

' 辅助函数：判断是否为当前页面
Function IsActive(pageName)
    If purchaseCurrentPage = pageName Then IsActive = "active" Else IsActive = ""
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
                <a href="<%= IIf(isFixedBrand,"../index.asp","index.asp") %>" class="admin-nav-brand">
                    <i class="fas fa-shopping-cart"></i>
                    <span>采购管理中心</span>
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
                <i class="fas fa-shopping-cart" style="color:#FF9800;"></i> 采购管理中心
            </h3>
            <p style="color:#888;font-size:11px;margin:8px 0 0;">六大品类 · 成本传导</p>
        </div>
        <ul class="sidebar-menu">
            <li class="sidebar-section-title">总览</li>
            <li><a href="<%= IIf(isFixedBrand,"../index.asp","index.asp") %>" class="<%= IsActive("index.asp") %>"><i class="fas fa-home"></i> <span>采购概览</span></a></li>
            <li class="sidebar-section-title">采购业务</li>
            <li><a href="<%= IIf(isFixedBrand,"../purchase_orders.asp","purchase_orders.asp") %>" class="<%= IsActive("purchase_orders.asp") %>"><i class="fas fa-file-invoice"></i> <span>采购订单</span></a></li>
            <li><a href="<%= IIf(isFixedBrand,"../receiving.asp","receiving.asp") %>" class="<%= IsActive("receiving.asp") %>"><i class="fas fa-clipboard-check"></i> <span>收货入库</span></a></li>
            <li><a href="<%= IIf(isFixedBrand,"../replenishment.asp","replenishment.asp") %>" class="<%= IsActive("replenishment.asp") %>"><i class="fas fa-robot"></i> <span>智能补货</span></a></li>
            <li class="sidebar-section-title">供应商管理</li>
            <li><a href="<%= IIf(isFixedBrand,"../supplier_management.asp","supplier_management.asp") %>" class="<%= IsActive("supplier_management.asp") %>"><i class="fas fa-truck"></i> <span>供应商管理</span></a></li>
            <li><a href="<%= IIf(isFixedBrand,"../price_management.asp","price_management.asp") %>" class="<%= IsActive("price_management.asp") %>"><i class="fas fa-tags"></i> <span>价格管理</span></a></li>
            <li class="sidebar-section-title">专项采购</li>
            <li><a href="<%= IIf(isFixedBrand,"../purchase_orders.asp?new=1&order_type=Packaging","purchase_orders.asp?new=1&order_type=Packaging") %>"><i class="fas fa-box"></i> <span>包装物采购</span></a></li>
            <li><a href="<%= IIf(isFixedBrand,"../purchase_orders.asp?new=1&order_type=Bottle","purchase_orders.asp?new=1&order_type=Bottle") %>"><i class="fas fa-wine-bottle"></i> <span>瓶子采购</span></a></li>
            <li><a href="<%= IIf(isFixedBrand,"../purchase_orders.asp?new=1&order_type=Printing","purchase_orders.asp?new=1&order_type=Printing") %>"><i class="fas fa-print"></i> <span>印刷品采购</span></a></li>
            <li><a href="<%= IIf(isFixedBrand,"../purchase_orders.asp?new=1&order_type=SprayHead","purchase_orders.asp?new=1&order_type=SprayHead") %>"><i class="fas fa-spray-can"></i> <span>喷头采购</span></a></li>
            <li class="sidebar-section-title">品牌定香采购</li>
            <li><a href="<%= IIf(isFixedBrand,"index.asp","fixed_brand/index.asp") %>" class="<%= IIf(isFixedBrand,"active","") %>"><i class="fas fa-boxes" style="color:#E91E63;"></i> <span>品牌定香概览</span></a></li>
            <li><a href="<%= IIf(isFixedBrand,"product_management.asp","fixed_brand/product_management.asp") %>" class="<%= IIf(isFixedBrand And purchaseCurrentPage="product_management.asp","active","") %>" style="padding-left:35px;"><i class="fas fa-cubes" style="color:#E91E63;"></i> <span>产品管理</span></a></li>
            <li><a href="<%= IIf(isFixedBrand,"purchase_orders.asp","fixed_brand/purchase_orders.asp") %>" class="<%= IIf(isFixedBrand And purchaseCurrentPage="purchase_orders.asp","active","") %>" style="padding-left:35px;"><i class="fas fa-file-invoice" style="color:#E91E63;"></i> <span>采购订单</span></a></li>
            <li><a href="<%= IIf(isFixedBrand,"receiving.asp","fixed_brand/receiving.asp") %>" class="<%= IIf(isFixedBrand And purchaseCurrentPage="receiving.asp","active","") %>" style="padding-left:35px;"><i class="fas fa-clipboard-check" style="color:#E91E63;"></i> <span>收货入库</span></a></li>
            <li><a href="<%= IIf(isFixedBrand,"replenishment.asp","fixed_brand/replenishment.asp") %>" class="<%= IIf(isFixedBrand And purchaseCurrentPage="replenishment.asp","active","") %>" style="padding-left:35px;"><i class="fas fa-robot" style="color:#E91E63;"></i> <span>智能补货</span></a></li>
            <li><a href="<%= IIf(isFixedBrand,"cost_profit.asp","fixed_brand/cost_profit.asp") %>" class="<%= IIf(isFixedBrand And purchaseCurrentPage="cost_profit.asp","active","") %>" style="padding-left:35px;"><i class="fas fa-chart-pie" style="color:#E91E63;"></i> <span>成本利润分析</span></a></li>
            <li class="sidebar-section-title">成本管理</li>
            <li><a href="<%= IIf(isFixedBrand,"../order_cost_summary.asp","order_cost_summary.asp") %>" class="<%= IsActive("order_cost_summary.asp") %>"><i class="fas fa-calculator"></i> <span>订单成本汇总</span></a></li>
        </ul>
        <div class="sidebar-footer" style="padding:15px 20px;border-top:1px solid rgba(255,255,255,0.06);">
            <a href="/admin/portal.asp" style="color:#666;text-decoration:none;font-size:13px;display:flex;align-items:center;gap:8px;">
                <i class="fas fa-arrow-left"></i> 返回管理中心
            </a>
        </div>
    </aside>
</div>
<!--#include file="../../includes/nav_common.asp"-->
