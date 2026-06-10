<!-- 产品技术管理中心导航 -->
<link rel="stylesheet" href="/css/admin.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
<%
' 获取当前页面名称用于高亮
Dim currentPage
currentPage = LCase(Request.ServerVariables("SCRIPT_NAME"))

' 判断当前用户角色（isManager 已在 auth.asp 中声明）
' TECH_MANAGER 显示全部功能，TECH_STAFF 隐藏删除相关功能

' 辅助函数：判断是否为当前页面
Function GetNavActiveClass(pageName)
    If InStr(currentPage, "/admin/techcenter/" & pageName) > 0 Then
        GetNavActiveClass = "active"
    Else
        GetNavActiveClass = ""
    End If
End Function
%>
<div class="admin-dashboard">
    <nav class="admin-navbar">
        <div class="admin-nav-container">
            <a href="index.asp" class="admin-nav-brand">
                <i class="fas fa-flask"></i>
                <span>产品技术管理中心</span>
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
            <!-- 技术概览（仪表板） -->
            <li class="nav-item">
                <a href="index.asp" class="<%= GetNavActiveClass("index.asp") %>">
                    <i class="fas fa-home"></i>
                    <span>技术概览</span>
                </a>
            </li>
            
            <!-- 分组：原料层 -->
            <li class="nav-group">
                <span class="group-title">原料层</span>
            </li>
            <li class="nav-item">
                <a href="base_note_management.asp" class="<%= GetNavActiveClass("base_note_management.asp") %>">
                    <i class="fas fa-wine-bottle"></i>
                    <span>基香管理</span>
                </a>
            </li>
            
            <!-- 分组：配方层 -->
            <li class="nav-group">
                <span class="group-title">配方层</span>
            </li>
            <li class="nav-item">
                <a href="note_management.asp" class="<%= GetNavActiveClass("note_management.asp") %>">
                    <i class="fas fa-leaf"></i>
                    <span>香调管理</span>
                </a>
            </li>
            <li class="nav-item">
                <a href="formula_management.asp" class="<%= GetNavActiveClass("formula_management.asp") %>">
                    <i class="fas fa-vial"></i>
                    <span>配方设置</span>
                </a>
            </li>
            <% If isManager Then %>
            <li class="nav-item">
                <a href="recipe_publish.asp" class="<%= GetNavActiveClass("recipe_publish.asp") %>">
                    <i class="fas fa-lock"></i>
                    <span>配方拆分发布 🔒</span>
                </a>
            </li>
            <% End If %>
            
            <!-- 分组：产品层 -->
            <li class="nav-group">
                <span class="group-title">产品层</span>
            </li>
            <li class="nav-item">
                <a href="product_settings.asp" class="<%= GetNavActiveClass("product_settings.asp") %>">
                    <i class="fas fa-box-open"></i>
                    <span>产品设置</span>
                </a>
            </li>
            <li class="nav-item">
                <a href="bottle_management.asp" class="<%= GetNavActiveClass("bottle_management.asp") %>">
                    <i class="fas fa-prescription-bottle"></i>
                    <span>瓶型管理</span>
                </a>
            </li>
            
            <!-- 分组：审核管理 -->
            <li class="nav-group">
                <span class="group-title">审核管理</span>
            </li>
            <li class="nav-item">
                <a href="kol_reviews.asp" class="<%= GetNavActiveClass("kol_reviews.asp") %>">
                    <i class="fas fa-user-check"></i>
                    <span>KOL审核</span>
                </a>
            </li>
        </ul>
    </aside>
</div>
<style>
/* 导航分组样式 */
.nav-group {
    margin-top: 15px;
    padding: 0 15px;
}
.nav-group .group-title {
    display: block;
    font-size: 11px;
    color: #888;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    padding: 8px 0;
    border-bottom: 1px solid #3a3a3a;
    margin-bottom: 5px;
}
.nav-item a {
    display: flex;
    align-items: center;
    padding: 10px 15px;
    color: #b0b0b0;
    text-decoration: none;
    transition: all 0.2s ease;
    border-radius: 4px;
    margin: 2px 10px;
}
.nav-item a:hover {
    background: rgba(255,255,255,0.05);
    color: #fff;
}
.nav-item a.active {
    background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%);
    color: #fff;
}
.nav-item a i {
    width: 20px;
    margin-right: 10px;
    text-align: center;
}
</style>
