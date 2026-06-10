<!-- 系统管理后台导航 -->
<link rel="stylesheet" href="/css/admin.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
<div class="admin-dashboard">
    <nav class="admin-navbar">
        <div class="admin-nav-container">
            <a href="index.asp" class="admin-nav-brand">
                <i class="fas fa-cogs"></i>
                <span>站点技术管理</span>
            </a>
            <ul class="admin-nav-menu">
                <li><a href="javascript:void(0)" onclick="location.reload()"><i class="fas fa-sync-alt"></i> 刷新</a></li>
                <li><a href="../portal.asp"><i class="fas fa-th-large"></i> 返回入口</a></li>
                <li><a href="../logout.asp"><i class="fas fa-sign-out-alt"></i> 退出</a></li>
            </ul>
        </div>
    </nav>
    <!-- 侧边栏 -->
    <aside class="sidebar">
        <ul class="sidebar-menu">
            <li><a href="index.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/system/index.asp", "active", "") %>"><i class="fas fa-home"></i> <span>系统概览</span></a></li>
            <li class="sidebar-section-title">权限管理</li>
            <li><a href="admins.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/system/admins.asp", "active", "") %>"><i class="fas fa-users-cog"></i> <span>管理员管理</span></a></li>
            <li><a href="roles.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/system/roles.asp", "active", "") %>"><i class="fas fa-user-tag"></i> <span>角色管理</span></a></li>
            <li><a href="logs.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/system/logs.asp", "active", "") %>"><i class="fas fa-history"></i> <span>操作日志</span></a></li>
            <li class="sidebar-section-title">系统安全</li>
            <li><a href="security_audit.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/system/security_audit.asp", "active", "") %>"><i class="fas fa-shield-alt"></i> <span>安全审计</span></a></li>
            <li><a href="backup_center.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/system/backup_center.asp", "active", "") %>"><i class="fas fa-database"></i> <span>备份中心</span></a></li>
            <li><a href="ip_blacklist.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/system/ip_blacklist.asp", "active", "") %>"><i class="fas fa-ban"></i> <span>IP黑名单</span></a></li>
            <li><a href="login_monitor.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/system/login_monitor.asp", "active", "") %>"><i class="fas fa-user-shield"></i> <span>登录监控</span></a></li>
            <li class="sidebar-section-title">站点配置</li>
            <li><a href="site_settings.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/system/site_settings.asp", "active", "") %>"><i class="fas fa-sliders-h"></i> <span>站点设置</span></a></li>
            <li><a href="settings.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/system/settings.asp", "active", "") %>"><i class="fas fa-cog"></i> <span>管理配置</span></a></li>
            <li class="sidebar-section-title">帮助文档</li>
            <li><a href="V8_USER_MANUAL.asp" target="_blank" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/admin/system/V8_USER_MANUAL.asp", "active", "") %>"><i class="fas fa-book"></i> <span>V8使用说明书</span></a></li>
        </ul>
    </aside>
</div>
<style>
/* 侧边栏深色主题 */
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
