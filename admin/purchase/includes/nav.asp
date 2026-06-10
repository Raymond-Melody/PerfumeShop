<!-- 采购管理中心侧边导航 -->
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
<div class="admin-sidebar" style="position:fixed;left:0;top:60px;width:250px;height:calc(100vh - 60px);background:#1a1a2e;border-right:1px solid rgba(255,255,255,0.06);overflow-y:auto;z-index:100;">
    <div class="sidebar-header" style="padding:20px;border-bottom:1px solid rgba(255,255,255,0.06);">
        <h3 style="color:#e0e0e0;margin:0;font-size:16px;display:flex;align-items:center;gap:10px;">
            <i class="fas fa-shopping-cart" style="color:#FF9800;"></i> 采购管理中心
        </h3>
        <p style="color:#888;font-size:11px;margin:8px 0 0;">六大品类 · 成本传导</p>
    </div>
    
    <nav class="sidebar-nav" style="padding:15px 0;">
        <div class="nav-section">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">总览</div>
            <a href="<%= IIf(isFixedBrand,"../index.asp","index.asp") %>" class="nav-item <%= IsActive("index.asp") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-home"></i> 采购概览
            </a>
        </div>
        
        <div class="nav-section" style="margin-top:15px;">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">采购业务</div>
            <a href="<%= IIf(isFixedBrand,"../purchase_orders.asp","purchase_orders.asp") %>" class="nav-item <%= IsActive("purchase_orders.asp") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-file-invoice"></i> 采购订单
            </a>
            <a href="<%= IIf(isFixedBrand,"../receiving.asp","receiving.asp") %>" class="nav-item <%= IsActive("receiving.asp") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-clipboard-check"></i> 收货入库
            </a>
            <a href="<%= IIf(isFixedBrand,"../replenishment.asp","replenishment.asp") %>" class="nav-item <%= IsActive("replenishment.asp") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-robot"></i> 智能补货
            </a>
        </div>
        
        <div class="nav-section" style="margin-top:15px;">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">供应商管理</div>
            <a href="<%= IIf(isFixedBrand,"../supplier_management.asp","supplier_management.asp") %>" class="nav-item <%= IsActive("supplier_management.asp") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-truck"></i> 供应商管理
            </a>
            <a href="<%= IIf(isFixedBrand,"../price_management.asp","price_management.asp") %>" class="nav-item <%= IsActive("price_management.asp") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-tags"></i> 价格管理
            </a>
        </div>
        
        <div class="nav-section" style="margin-top:15px;">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">专项采购</div>
            <a href="<%= IIf(isFixedBrand,"../purchase_orders.asp?new=1&order_type=Packaging","purchase_orders.asp?new=1&order_type=Packaging") %>" class="nav-item" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-box"></i> 包装物采购
            </a>
            <a href="<%= IIf(isFixedBrand,"../purchase_orders.asp?new=1&order_type=Bottle","purchase_orders.asp?new=1&order_type=Bottle") %>" class="nav-item" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-wine-bottle"></i> 瓶子采购
            </a>
            <a href="<%= IIf(isFixedBrand,"../purchase_orders.asp?new=1&order_type=Printing","purchase_orders.asp?new=1&order_type=Printing") %>" class="nav-item" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-print"></i> 印刷品采购
            </a>
            <a href="<%= IIf(isFixedBrand,"../purchase_orders.asp?new=1&order_type=SprayHead","purchase_orders.asp?new=1&order_type=SprayHead") %>" class="nav-item" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-spray-can"></i> 喷头采购
            </a>
        </div>
        
        <div class="nav-section" style="margin-top:15px;">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">品牌定香采购</div>
            <a href="<%= IIf(isFixedBrand,"index.asp","fixed_brand/index.asp") %>" class="nav-item <%= IIf(isFixedBrand,"active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-boxes" style="color:#E91E63;"></i> 品牌定香概览
            </a>
            <a href="<%= IIf(isFixedBrand,"product_management.asp","fixed_brand/product_management.asp") %>" class="nav-item <%= IIf(isFixedBrand And purchaseCurrentPage="product_management.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px 12px 35px;color:#b0b0b0;text-decoration:none;font-size:13px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-cubes" style="color:#E91E63;"></i> 产品管理
            </a>
            <a href="<%= IIf(isFixedBrand,"purchase_orders.asp","fixed_brand/purchase_orders.asp") %>" class="nav-item <%= IIf(isFixedBrand And purchaseCurrentPage="purchase_orders.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px 12px 35px;color:#b0b0b0;text-decoration:none;font-size:13px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-file-invoice" style="color:#E91E63;"></i> 采购订单
            </a>
            <a href="<%= IIf(isFixedBrand,"receiving.asp","fixed_brand/receiving.asp") %>" class="nav-item <%= IIf(isFixedBrand And purchaseCurrentPage="receiving.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px 12px 35px;color:#b0b0b0;text-decoration:none;font-size:13px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-clipboard-check" style="color:#E91E63;"></i> 收货入库
            </a>
            <a href="<%= IIf(isFixedBrand,"replenishment.asp","fixed_brand/replenishment.asp") %>" class="nav-item <%= IIf(isFixedBrand And purchaseCurrentPage="replenishment.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px 12px 35px;color:#b0b0b0;text-decoration:none;font-size:13px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-robot" style="color:#E91E63;"></i> 智能补货
            </a>
            <a href="<%= IIf(isFixedBrand,"cost_profit.asp","fixed_brand/cost_profit.asp") %>" class="nav-item <%= IIf(isFixedBrand And purchaseCurrentPage="cost_profit.asp","active","") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px 12px 35px;color:#b0b0b0;text-decoration:none;font-size:13px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-chart-pie" style="color:#E91E63;"></i> 成本利润分析
            </a>
        </div>
        
        <div class="nav-section" style="margin-top:15px;">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">成本管理</div>
            <a href="<%= IIf(isFixedBrand,"../order_cost_summary.asp","order_cost_summary.asp") %>" class="nav-item <%= IsActive("order_cost_summary.asp") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-calculator"></i> 订单成本汇总
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
.admin-sidebar .nav-item.active { background: rgba(255,152,0,0.12); color: #FF9800; border-left-color: #FF9800; font-weight: 500; }
</style>
