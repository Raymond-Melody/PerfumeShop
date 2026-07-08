<!-- 系统管理后台导航 V18 -->
<link rel="stylesheet" href="/css/admin.css">
<link rel="stylesheet" href="/css/responsive.css?v=18.0">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
<div class="admin-dashboard">
    <nav class="admin-navbar">
        <div class="admin-nav-container">
            <div style="display:flex;align-items:center;">
                <button class="admin-hamburger" id="adminHamburger" aria-label="菜单">
                    <span></span>
                    <span></span>
                    <span></span>
                </button>
                <a href="index.asp" class="admin-nav-brand">
                    <i class="fas fa-cogs"></i>
                    <span>站点技术管理</span>
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
        <ul class="sidebar-menu">
            <li><a href="index.asp"<% If LCase(Request.ServerVariables("SCRIPT_NAME") & "") = "/admin/system/index.asp" Then %> class="active"<% End If %>><i class="fas fa-home"></i> <span>系统概览</span></a></li>
            <li class="sidebar-section-title">权限管理</li>
            <li><a href="admins.asp"<% If LCase(Request.ServerVariables("SCRIPT_NAME") & "") = "/admin/system/admins.asp" Then %> class="active"<% End If %>><i class="fas fa-users-cog"></i> <span>管理员管理</span></a></li>
            <li><a href="roles.asp"<% If LCase(Request.ServerVariables("SCRIPT_NAME") & "") = "/admin/system/roles.asp" Then %> class="active"<% End If %>><i class="fas fa-user-tag"></i> <span>角色管理</span></a></li>
            <li><a href="logs.asp"<% If LCase(Request.ServerVariables("SCRIPT_NAME") & "") = "/admin/system/logs.asp" Then %> class="active"<% End If %>><i class="fas fa-history"></i> <span>操作日志</span></a></li>
            <li><a href="audit_logs.asp"<% If LCase(Request.ServerVariables("SCRIPT_NAME") & "") = "/admin/system/audit_logs.asp" Then %> class="active"<% End If %>><i class="fas fa-shield-alt"></i> <span>审计日志 V16</span></a></li>
            <li class="sidebar-section-title">系统安全</li>
            <li><a href="security_audit.asp"<% If LCase(Request.ServerVariables("SCRIPT_NAME") & "") = "/admin/system/security_audit.asp" Then %> class="active"<% End If %>><i class="fas fa-shield-alt"></i> <span>安全审计</span></a></li>
            <li><a href="backup_center.asp"<% If LCase(Request.ServerVariables("SCRIPT_NAME") & "") = "/admin/system/backup_center.asp" Then %> class="active"<% End If %>><i class="fas fa-database"></i> <span>备份中心</span></a></li>
            <li><a href="ip_blacklist.asp"<% If LCase(Request.ServerVariables("SCRIPT_NAME") & "") = "/admin/system/ip_blacklist.asp" Then %> class="active"<% End If %>><i class="fas fa-ban"></i> <span>IP黑名单</span></a></li>
            <li><a href="login_monitor.asp"<% If LCase(Request.ServerVariables("SCRIPT_NAME") & "") = "/admin/system/login_monitor.asp" Then %> class="active"<% End If %>><i class="fas fa-user-shield"></i> <span>登录监控</span></a></li>
            <li class="sidebar-section-title">站点配置</li>
            <li><a href="site_settings.asp"<% If LCase(Request.ServerVariables("SCRIPT_NAME") & "") = "/admin/system/site_settings.asp" Then %> class="active"<% End If %>><i class="fas fa-sliders-h"></i> <span>站点设置</span></a></li>
            <li><a href="settings.asp"<% If LCase(Request.ServerVariables("SCRIPT_NAME") & "") = "/admin/system/settings.asp" Then %> class="active"<% End If %>><i class="fas fa-cog"></i> <span>管理配置</span></a></li>
            <li class="sidebar-section-title">帮助文档</li>
            <li><a href="guide.asp"<% If LCase(Request.ServerVariables("SCRIPT_NAME") & "") = "/admin/system/guide.asp" Then %> class="active"<% End If %>><i class="fas fa-book-open"></i> <span>使用流程指南</span></a></li>
            <li><a href="version.asp"<% If LCase(Request.ServerVariables("SCRIPT_NAME") & "") = "/admin/system/version.asp" Then %> class="active"<% End If %>><i class="fas fa-code-branch"></i> <span>版本信息</span></a></li>
        </ul>
    </aside>
</div>
<!--#include file="../../includes/nav_common.asp"-->
