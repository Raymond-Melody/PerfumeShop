<!-- 财务管理后台导航 V18 -->
<link rel="stylesheet" href="/css/admin.css">
<link rel="stylesheet" href="../operation/css/operation-dark.css">
<link rel="stylesheet" href="/css/responsive.css?v=18.0">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
<%
' 获取当前页面名称用于高亮
Dim currentPage
currentPage = LCase(Request.ServerVariables("SCRIPT_NAME"))

' 判断当前用户角色
Dim isManager
isManager = False
If Session("AdminRoleCode") = "FIN_MANAGER" Or Session("AdminRoleCode") = "SUPER_ADMIN" Then
    isManager = True
End If

' 辅助函数：判断是否为当前页面
Function IsActive(pageName)
    If InStr(currentPage, "/admin/finance/" & pageName) > 0 Then
        IsActive = "active"
    Else
        IsActive = ""
    End If
End Function
%>
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
                    <i class="fas fa-dollar-sign"></i>
                    <span>财务管理中心</span>
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
            <!-- 财务概览 -->
            <li class="nav-item">
                <a href="index.asp" class="<%= IsActive("index.asp") %>">
                    <i class="fas fa-home"></i>
                    <span>财务概览</span>
                </a>
            </li>
            
            <!-- 分组：财务中台 -->
            <li class="nav-group">
                <span class="group-title">财务中台</span>
            </li>
            <li class="nav-item">
                <a href="cost_management.asp" class="<%= IsActive("cost_management.asp") %>">
                    <i class="fas fa-coins"></i>
                    <span>成本管理</span>
                </a>
            </li>
            <li class="nav-item">
                <a href="expense_allocation.asp" class="<%= IsActive("expense_allocation.asp") %>">
                    <i class="fas fa-calculator"></i>
                    <span>费用分摊</span>
                </a>
            </li>
            <li class="nav-item">
                <a href="reconciliation.asp" class="<%= IsActive("reconciliation.asp") %>">
                    <i class="fas fa-balance-scale"></i>
                    <span>对账中心</span>
                </a>
            </li>
            <% If isManager Then %>
            <li class="nav-item">
                <a href="purchase_review.asp" class="<%= IsActive("purchase_review.asp") %>">
                    <i class="fas fa-clipboard-check"></i>
                    <span>采购审核</span>
                </a>
            </li>
            <% End If %>
            
            <!-- 分组：资金台账 -->
            <li class="nav-group">
                <span class="group-title">资金台账</span>
            </li>
            <li class="nav-item">
                <a href="fund_dashboard.asp" class="<%= IsActive("fund_dashboard.asp") %>">
                    <i class="fas fa-wallet"></i>
                    <span>资金看板</span>
                </a>
            </li>
            <li class="nav-item">
                <a href="cash_flow.asp" class="<%= IsActive("cash_flow.asp") %>">
                    <i class="fas fa-chart-line"></i>
                    <span>现金流预测</span>
                </a>
            </li>
            <li class="nav-item">
                <a href="transactions.asp" class="<%= IsActive("transactions.asp") %>">
                    <i class="fas fa-exchange-alt"></i>
                    <span>流水管理</span>
                </a>
            </li>
            <li class="nav-item">
                <a href="accounts_payable.asp" class="<%= IsActive("accounts_payable.asp") %>">
                    <i class="fas fa-file-invoice-dollar"></i>
                    <span>应付账款</span>
                </a>
            </li>
            <li class="nav-item">
                <a href="accounts_receivable.asp" class="<%= IsActive("accounts_receivable.asp") %>">
                    <i class="fas fa-hand-holding-usd"></i>
                    <span>应收账款</span>
                </a>
            </li>
            <li class="nav-item">
                <a href="payment_vouchers.asp" class="<%= IsActive("payment_vouchers.asp") %>">
                    <i class="fas fa-receipt"></i>
                    <span>付款凭证</span>
                </a>
            </li>
            
            <!-- 分组：报表分析 -->
            <li class="nav-group">
                <span class="group-title">报表分析</span>
            </li>
            <li class="nav-item">
                <a href="profit_report.asp" class="<%= IsActive("profit_report.asp") %>">
                    <i class="fas fa-chart-line"></i>
                    <span>经营利润表</span>
                </a>
            </li>
            <li class="nav-item">
                <a href="product_analysis.asp" class="<%= IsActive("product_analysis.asp") %>">
                    <i class="fas fa-box-open"></i>
                    <span>单品分析</span>
                </a>
            </li>
            <li class="nav-item">
                <a href="marketing_stats.asp" class="<%= IsActive("marketing_stats.asp") %>">
                    <i class="fas fa-bullhorn"></i>
                    <span>营销统计</span>
                </a>
            </li>
            <li class="nav-item">
                <a href="budget_management.asp" class="<%= IsActive("budget_management.asp") %>">
                    <i class="fas fa-chart-pie"></i>
                    <span>预算管理</span>
                </a>
            </li>
            <li class="nav-item">
                <a href="gl_report.asp" class="<%= IsActive("gl_report.asp") %>">
                    <i class="fas fa-book"></i>
                    <span>总账报表</span>
                </a>
            </li>
            
            <!-- 分组：系统配置 -->
            <li class="nav-group">
                <span class="group-title">系统配置</span>
            </li>
            <% If isManager Then %>
            <li class="nav-item">
                <a href="payment_config.asp" class="<%= IsActive("payment_config.asp") %>">
                    <i class="fas fa-credit-card"></i>
                    <span>支付配置</span>
                </a>
            </li>
            <% End If %>
            <li class="nav-item">
                <a href="revenue.asp" class="<%= IsActive("revenue.asp") %>">
                    <i class="fas fa-chart-bar"></i>
                    <span>收入统计</span>
                </a>
            </li>
            <li class="nav-item">
                <a href="reports.asp" class="<%= IsActive("reports.asp") %>">
                    <i class="fas fa-file-alt"></i>
                    <span>财务报表</span>
                </a>
            </li>
            <li class="nav-item">
                <a href="cost_centers.asp" class="<%= IsActive("cost_centers.asp") %>">
                    <i class="fas fa-building"></i>
                    <span>成本中心</span>
                </a>
            </li>
            
            <!-- 分组：综合报表 -->
            <li class="nav-group">
                <span class="group-title">综合报表</span>
            </li>
            <li class="nav-item">
                <a href="comprehensive_report.asp" class="<%= IsActive("comprehensive_report.asp") %>">
                    <i class="fas fa-chart-pie"></i>
                    <span>综合报表中心</span>
                </a>
            </li>
            
            <!-- 分组：风控管理 -->
            <li class="nav-group">
                <span class="group-title">风控管理</span>
            </li>
            <li class="nav-item">
                <a href="risk_control.asp" class="<%= IsActive("risk_control.asp") %>">
                    <i class="fas fa-shield-alt"></i>
                    <span>风控管理</span>
                </a>
            </li>
        </ul>
    </aside>
</div>
<!--#include file="../../includes/nav_common.asp"-->
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
