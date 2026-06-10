<!-- 库存管理中心侧边导航 -->
<%
Dim invCurrentPage
invCurrentPage = Request.ServerVariables("SCRIPT_NAME")
Dim invParts : invParts = Split(CStr(invCurrentPage), "/")
invCurrentPage = invParts(UBound(invParts))

' 计算导航项active状态（避免依赖外部IIf函数）
Function NavActive(checkPage)
    If invCurrentPage = checkPage Then NavActive = "active" Else NavActive = ""
End Function
%>
<div class="admin-sidebar" style="position:fixed;left:0;top:60px;width:250px;height:calc(100vh - 60px);background:#1a1a2e;border-right:1px solid rgba(255,255,255,0.06);overflow-y:auto;z-index:100;">
    <div class="sidebar-header" style="padding:20px;border-bottom:1px solid rgba(255,255,255,0.06);">
        <h3 style="color:#e0e0e0;margin:0;font-size:16px;display:flex;align-items:center;gap:10px;">
            <i class="fas fa-warehouse" style="color:#00BCD4;"></i> 库存管理中心
        </h3>
        <p style="color:#888;font-size:11px;margin:8px 0 0;">全品类 · 统一管控</p>
    </div>
    
    <nav class="sidebar-nav" style="padding:15px 0;">
        <div class="nav-section">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">总览</div>
            <a href="/admin/inventory/index.asp" class="nav-item <%= NavActive("index.asp") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-chart-pie"></i> 库存仪表盘
            </a>
        </div>
        
        <div class="nav-section" style="margin-top:15px;">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">成品库存</div>
            <a href="/admin/prodcenter/product_inventory.asp?from=inventory" class="nav-item" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-box"></i> 成品库存
            </a>
            <a href="/admin/prodcenter/bottle_inventory.asp?from=inventory" class="nav-item" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-flask"></i> 瓶子库存
            </a>
            <a href="/admin/prodcenter/packaging_inventory.asp?from=inventory" class="nav-item" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-box-open"></i> 包装物库存
            </a>
        </div>
        
        <div class="nav-section" style="margin-top:15px;">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">半成品库存</div>
            <a href="/admin/semifinished/raw_material_inventory.asp?from=inventory" class="nav-item" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-boxes"></i> 原料库存
            </a>
            <a href="/admin/semifinished/base_note_inventory.asp?from=inventory" class="nav-item" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-database"></i> 基香库存
            </a>
            <a href="/admin/semifinished/note_inventory.asp?from=inventory" class="nav-item" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-layer-group"></i> 香调库存
            </a>
        </div>
        
        <div class="nav-section" style="margin-top:15px;">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">库存运营</div>
            <a href="/admin/inventory/stock_movements.asp" class="nav-item <%= NavActive("stock_movements.asp") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-history"></i> 库存流水
            </a>
        </div>
        
        <div class="nav-section" style="margin-top:15px;">
            <div class="nav-section-title" style="padding:8px 20px;color:#666;font-size:11px;text-transform:uppercase;letter-spacing:1px;">预警</div>
            <a href="/admin/inventory/inventory_alerts.asp" class="nav-item <%= NavActive("inventory_alerts.asp") %>" style="display:flex;align-items:center;gap:10px;padding:12px 20px;color:#b0b0b0;text-decoration:none;font-size:14px;transition:all 0.2s;border-left:3px solid transparent;">
                <i class="fas fa-exclamation-triangle"></i> 库存预警
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
.admin-sidebar .nav-item:hover { background: rgba(255,255,255,0.05); color: #e0e0e0; border-left-color: #00BCD4; }
.admin-sidebar .nav-item.active { background: rgba(0,188,212,0.12); color: #00BCD4; border-left-color: #00BCD4; }
</style>
