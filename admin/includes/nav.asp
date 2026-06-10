<!-- V8 全局管理中心导航 -->
<link rel="stylesheet" href="/css/admin.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
<div class="admin-dashboard">
    <nav class="admin-navbar">
        <div class="admin-nav-container">
            <a href="portal.asp" class="admin-nav-brand">
                <i class="fas fa-cubes"></i>
                <span>V8 管理中心</span>
            </a>
            <ul class="admin-nav-menu">
                <li><a href="javascript:void(0)" onclick="location.reload()"><i class="fas fa-sync-alt"></i> 刷新</a></li>
                <li><a href="../index.asp" target="_blank"><i class="fas fa-store"></i> 访问前台</a></li>
                <li><a href="logout.asp"><i class="fas fa-sign-out-alt"></i> 退出登录</a></li>
            </ul>
        </div>
    </nav>
    
    <aside class="sidebar">
        <ul class="sidebar-menu">
            <li><a href="portal.asp"><i class="fas fa-th-large"></i> <span>返回入口</span></a></li>
            <li class="sidebar-section-title">业务模块</li>
            <li><a href="operation/index.asp"><i class="fas fa-chart-line"></i> <span>运营管理</span></a></li>
            <li><a href="semifinished/index.asp"><i class="fas fa-vial"></i> <span>半成品生产</span></a></li>
            <li><a href="prodcenter/index.asp"><i class="fas fa-industry"></i> <span>产品生产</span></a></li>
            <li><a href="logistics/index.asp"><i class="fas fa-truck"></i> <span>物流管理</span></a></li>
            <li><a href="purchase/index.asp"><i class="fas fa-shopping-cart"></i> <span>采购管理</span></a></li>
            <li><a href="finance/index.asp"><i class="fas fa-dollar-sign"></i> <span>财务管理</span></a></li>
            <li><a href="techcenter/index.asp"><i class="fas fa-flask"></i> <span>技术中心</span></a></li>
            <li class="sidebar-section-title">系统管理</li>
            <li><a href="system/index.asp"><i class="fas fa-shield-alt"></i> <span>系统管理</span></a></li>
            <li><a href="site_settings.asp"><i class="fas fa-cog"></i> <span>站点设置</span></a></li>
        </ul>
    </aside>
</div>