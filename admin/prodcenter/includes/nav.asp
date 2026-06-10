<!-- 产品生产管理中心侧边导航 -->
<%
Dim navFrom : navFrom = Request.QueryString("from")
If navFrom = "inventory" Then
    Server.Execute("/admin/inventory/includes/nav.asp")
Else
    Dim prodCurrentPage
    prodCurrentPage = Request.ServerVariables("SCRIPT_NAME")
    prodCurrentPage = Mid(prodCurrentPage, InStrRev(prodCurrentPage, "/") + 1)
%>
<div class="admin-sidebar" style="position:fixed;left:0;top:60px;width:250px;height:calc(100vh - 60px);background:#1a1a2e;border-right:1px solid rgba(255,255,255,0.06);overflow-y:auto;z-index:100;">
    <div class="sidebar-header" style="padding:20px;border-bottom:1px solid rgba(255,255,255,0.06);">
        <h3 style="color:#e0e0e0;margin:0;font-size:16px;display:flex;align-items:center;gap:10px;">
            <i class="fas fa-industry" style="color:#4CAF50;"></i> 产品生产中心
        </h3>
        <p style="color:#888;font-size:11px;margin:8px 0 0;">香调 → 成品</p>
    </div>
    
    <nav class="sidebar-nav" style="padding:15px 0;">
        <div class="nav-section">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">总览</div>
            <a href="index.asp" class="nav-item <%= IIf(prodCurrentPage="index.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-tachometer-alt"></i> 生产概览
            </a>
        </div>
        
        <div class="nav-section" style="margin-top:15px;">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">生产管理</div>
            <a href="production_management.asp" class="nav-item <%= IIf(prodCurrentPage="production_management.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-clipboard-list"></i> 生产工单管理
            </a>
            <a href="prod_scheduling.asp" class="nav-item <%= IIf(prodCurrentPage="prod_scheduling.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-calendar-alt"></i> 排产调度 <span style="font-size:9px;color:#4CAF50;font-weight:700;margin-left:2px;">批量</span>
            </a>
            <a href="prod_workshop.asp" class="nav-item <%= IIf(prodCurrentPage="prod_workshop.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-cogs"></i> 车间作业
            </a>
            <a href="prod_qc.asp" class="nav-item <%= IIf(prodCurrentPage="prod_qc.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-check-circle"></i> 质量检验
            </a>
            <a href="prod_warehouse.asp" class="nav-item <%= IIf(prodCurrentPage="prod_warehouse.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-warehouse"></i> 成品入库
            </a>
            <a href="order_production.asp" class="nav-item <%= IIf(prodCurrentPage="order_production.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-shopping-cart"></i> 订单级生产追踪
            </a>
        </div>
        
        <div class="nav-section" style="margin-top:15px;">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">库存管理</div>
            <a href="product_inventory.asp" class="nav-item <%= IIf(prodCurrentPage="product_inventory.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-box"></i> 成品库存
            </a>
            <a href="bottle_inventory.asp" class="nav-item <%= IIf(prodCurrentPage="bottle_inventory.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-flask"></i> 瓶子库存
            </a>
            <a href="packaging_inventory.asp" class="nav-item <%= IIf(prodCurrentPage="packaging_inventory.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-box-open"></i> 包装物库存
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
.admin-sidebar .nav-item:hover { background: rgba(255,255,255,0.05); color: #e0e0e0; border-left-color: #4CAF50; }
.admin-sidebar .nav-item.active { background: rgba(76,175,80,0.12); color: #4CAF50; border-left-color: #4CAF50; }
</style>
<% End If %>
