<!-- V9.0 全局管理中心导航 -->
<!--#include file="../includes/i18n.asp"-->
<link rel="stylesheet" href="/css/admin.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
<div class="admin-dashboard">
    <nav class="admin-navbar">
        <div class="admin-nav-container">
            <a href="portal.asp" class="admin-nav-brand">
                <i class="fas fa-cubes"></i>
                <span><% If FEATURE_I18N Then Response.Write T("admin_nav_brand", Empty) Else %>V9.0 管理中心<% End If %></span>
            </a>
            <ul class="admin-nav-menu">
                <li><a href="javascript:void(0)" onclick="location.reload()"><i class="fas fa-sync-alt"></i> <% If FEATURE_I18N Then Response.Write T("admin_nav_refresh", Empty) Else %>刷新<% End If %></a></li>
                <li><a href="../index.asp" target="_blank"><i class="fas fa-store"></i> <% If FEATURE_I18N Then Response.Write T("admin_nav_visit_site", Empty) Else %>访问前台<% End If %></a></li>
                <li><a href="logout.asp"><i class="fas fa-sign-out-alt"></i> <% If FEATURE_I18N Then Response.Write T("admin_nav_logout", Empty) Else %>退出登录<% End If %></a></li>
            </ul>
        </div>
    </nav>
    
    <aside class="sidebar">
        <ul class="sidebar-menu">
            <li><a href="portal.asp"><i class="fas fa-th-large"></i> <span><% If FEATURE_I18N Then Response.Write T("admin_nav_back", Empty) Else %>返回入口<% End If %></span></a></li>
            <li class="sidebar-section-title"><% If FEATURE_I18N Then Response.Write T("admin_nav_business", Empty) Else %>业务模块<% End If %></li>
            <li><a href="operation/index.asp"><i class="fas fa-chart-line"></i> <span><% If FEATURE_I18N Then Response.Write T("admin_nav_operation", Empty) Else %>运营管理<% End If %></span></a></li>
            <li><a href="semifinished/index.asp"><i class="fas fa-vial"></i> <span><% If FEATURE_I18N Then Response.Write T("admin_nav_semifinished", Empty) Else %>半成品生产<% End If %></span></a></li>
            <li><a href="prodcenter/index.asp"><i class="fas fa-industry"></i> <span><% If FEATURE_I18N Then Response.Write T("admin_nav_prodcenter", Empty) Else %>产品生产<% End If %></span></a></li>
            <li><a href="logistics/index.asp"><i class="fas fa-truck"></i> <span><% If FEATURE_I18N Then Response.Write T("admin_nav_logistics", Empty) Else %>物流管理<% End If %></span></a></li>
            <li><a href="purchase/index.asp"><i class="fas fa-shopping-cart"></i> <span><% If FEATURE_I18N Then Response.Write T("admin_nav_purchase", Empty) Else %>采购管理<% End If %></span></a></li>
            <li><a href="finance/index.asp"><i class="fas fa-dollar-sign"></i> <span><% If FEATURE_I18N Then Response.Write T("admin_nav_finance", Empty) Else %>财务管理<% End If %></span></a></li>
            <li><a href="techcenter/index.asp"><i class="fas fa-flask"></i> <span><% If FEATURE_I18N Then Response.Write T("admin_nav_tech", Empty) Else %>技术中心<% End If %></span></a></li>
            <li class="sidebar-section-title"><% If FEATURE_I18N Then Response.Write T("admin_nav_sys_mgmt", Empty) Else %>系统管理<% End If %></li>
            <li><a href="system/index.asp"><i class="fas fa-shield-alt"></i> <span><% If FEATURE_I18N Then Response.Write T("admin_nav_system", Empty) Else %>系统管理<% End If %></span></a></li>
            <li><a href="site_settings.asp"><i class="fas fa-cog"></i> <span><% If FEATURE_I18N Then Response.Write T("admin_nav_site_settings", Empty) Else %>站点设置<% End If %></span></a></li>
        </ul>
    </aside>
</div>