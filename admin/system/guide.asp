<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>使用流程指南 - 系统管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; font-family: 'Segoe UI','Microsoft YaHei',sans-serif; }
        
        .guide-section { margin-bottom: 30px; }
        .guide-section h2 { 
            color: #fff; font-size: 20px; margin: 30px 0 15px; 
            padding: 12px 20px; 
            background: linear-gradient(135deg, #1e2a3a, #1a2535); 
            border-left: 4px solid #00bcd4; border-radius: 0 8px 8px 0; 
        }
        .guide-section h3 { 
            color: #80deea; font-size: 16px; margin: 18px 0 10px; 
        }
        
        .guide-card { 
            background: linear-gradient(135deg, #2d2d44, #1e1e32);
            border: 1px solid rgba(255,255,255,0.06); border-radius: 12px; 
            padding: 22px; margin-bottom: 18px;
        }
        
        .flow-row { 
            display: flex; align-items: center; justify-content: flex-start; 
            flex-wrap: wrap; gap: 8px; padding: 15px; 
            background: rgba(0,188,212,0.05); border-radius: 8px; 
            margin: 12px 0; 
        }
        .flow-step { 
            background: linear-gradient(135deg, #00bcd4, #00838f); 
            color: #fff; padding: 7px 14px; border-radius: 20px; 
            font-size: 13px; font-weight: 600; white-space: nowrap; 
        }
        .flow-arrow { color: #00bcd4; font-size: 16px; font-weight: 700; }
        
        .guide-table { width: 100%; border-collapse: collapse; margin: 8px 0; }
        .guide-table th { 
            text-align: left; padding: 10px 14px; 
            background: rgba(0,188,212,0.08); color: #80deea; 
            font-size: 13px; border-bottom: 1px solid rgba(255,255,255,0.06); 
        }
        .guide-table td { 
            padding: 10px 14px; border-bottom: 1px solid rgba(255,255,255,0.03); 
            font-size: 13px; color: #d0d6e0; 
        }
        .guide-table tr:hover td { background: rgba(255,255,255,0.02); }
        
        .role-tag { display: inline-block; padding: 2px 10px; border-radius: 12px; font-size: 11px; margin-left: 6px; }
        .role-tag-user { background: rgba(76,175,80,0.2); color: #81c784; }
        .role-tag-admin { background: rgba(0,188,212,0.2); color: #80deea; }
        .role-tag-sys { background: rgba(255,152,0,0.2); color: #ffb74d; }
        
        .module-grid { 
            display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); 
            gap: 16px; margin-top: 12px; 
        }
        .module-card { 
            background: rgba(0,188,212,0.04); border: 1px solid rgba(0,188,212,0.1); 
            border-radius: 10px; padding: 16px; 
        }
        .module-card h4 { color: #00bcd4; margin-bottom: 6px; font-size: 15px; }
        .module-card p { font-size: 13px; color: #90a4ae; }
        .module-card .page-count { font-size: 12px; color: #546e7a; margin-top: 5px; }
        
        .note-warn { 
            color: #ffb74d; font-size: 13px; margin-top: 10px; 
            padding: 10px 14px; background: rgba(255,152,0,0.06); 
            border-radius: 8px; border-left: 3px solid rgba(255,152,0,0.3);
        }
        .note-warn code { background: rgba(255,255,255,0.08); padding: 2px 6px; border-radius: 4px; font-size: 12px; }
        
        .guide-footer { 
            text-align: center; color: #455a64; font-size: 12px; 
            padding: 20px 0; margin-top: 20px;
            border-top: 1px solid rgba(255,255,255,0.04);
        }
        
        @media (max-width: 768px) {
            .module-grid { grid-template-columns: 1fr; }
            .flow-row { justify-content: center; }
        }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-book-open"></i> 使用流程指南</h2>
            <div class="breadcrumb">
                <a href="index.asp">系统中心</a> / <span>使用指南</span>
            </div>
        </div>

        <p style="color:#78909c;font-size:14px;margin-bottom:25px;">
            <%= SITE_NAME %> · <%= SYS_VERSION %> 完整使用流程指南 · 涵盖全部10大模块
        </p>

        <!-- ========== 一、用户购物流程 ========== -->
        <div class="guide-section">
            <h2>一、用户购物流程（前台）</h2>
            
            <div class="guide-card">
                <h3>1.1 浏览与选购</h3>
                <div class="flow-row">
                    <span class="flow-step">首页 index.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">产品列表 products.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">产品详情 product.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">定制调香 customize.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">加入购物车</span>
                </div>
                <table class="guide-table">
                    <tr><th>页面</th><th>功能</th><th>角色</th></tr>
                    <tr><td>/index.asp</td><td>网站首页，展示推荐产品、热门分类<span class="role-tag role-tag-user">用户</span></td></tr>
                    <tr><td>/products.asp</td><td>产品列表页，支持分类筛选、搜索<span class="role-tag role-tag-user">用户</span></td></tr>
                    <tr><td>/product.asp</td><td>产品详情，查看香调、配料、价格<span class="role-tag role-tag-user">用户</span></td></tr>
                    <tr><td>/customize.asp</td><td>定制调香页面，选择香型组合<span class="role-tag role-tag-user">用户</span></td></tr>
                    <tr><td>/order_ingredients.asp</td><td>配料订购页面<span class="role-tag role-tag-user">用户</span></td></tr>
                </table>
            </div>

            <div class="guide-card">
                <h3>1.2 下单与支付</h3>
                <div class="flow-row">
                    <span class="flow-step">购物车 cart.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">结算 checkout.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">支付回调 payment_callback.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">订单成功 order_success.asp</span>
                </div>
                <table class="guide-table">
                    <tr><th>页面</th><th>功能</th><th>角色</th></tr>
                    <tr><td>/cart.asp</td><td>购物车管理（增删改）<span class="role-tag role-tag-user">用户</span></td></tr>
                    <tr><td>/checkout.asp</td><td>结算页，填写地址、选择支付<span class="role-tag role-tag-user">用户</span></td></tr>
                    <tr><td>/payment_callback.asp</td><td>支付异步回调处理<span class="role-tag role-tag-sys">系统</span></td></tr>
                    <tr><td>/order_success.asp</td><td>下单成功确认页<span class="role-tag role-tag-user">用户</span></td></tr>
                </table>
            </div>

            <div class="guide-card">
                <h3>1.3 用户中心</h3>
                <table class="guide-table">
                    <tr><th>页面</th><th>功能</th><th>角色</th></tr>
                    <tr><td>/user/index.asp</td><td>用户中心首页<span class="role-tag role-tag-user">用户</span></td></tr>
                    <tr><td>/user/orders.asp</td><td>我的订单列表<span class="role-tag role-tag-user">用户</span></td></tr>
                    <tr><td>/user/order_detail.asp</td><td>订单详情<span class="role-tag role-tag-user">用户</span></td></tr>
                    <tr><td>/user/favorites.asp</td><td>我的收藏<span class="role-tag role-tag-user">用户</span></td></tr>
                    <tr><td>/user/addresses.asp</td><td>收货地址管理<span class="role-tag role-tag-user">用户</span></td></tr>
                    <tr><td>/user/settings.asp</td><td>个人设置<span class="role-tag role-tag-user">用户</span></td></tr>
                    <tr><td>/user/register.asp</td><td>用户注册<span class="role-tag role-tag-user">用户</span></td></tr>
                    <tr><td>/user/login.asp</td><td>用户登录<span class="role-tag role-tag-user">用户</span></td></tr>
                </table>
            </div>
        </div>

        <!-- ========== 二、后台模块总览 ========== -->
        <div class="guide-section">
            <h2>二、后台管理系统总览</h2>
            <div class="module-grid">
                <div class="module-card">
                    <h4><i class="fas fa-chart-line"></i> 运营中心 (operation)</h4>
                    <p>订单处理、发货、退款、营销统计、综合报表</p>
                    <p class="page-count">22 页面</p>
                </div>
                <div class="module-card">
                    <h4><i class="fas fa-dollar-sign"></i> 财务中心 (finance)</h4>
                    <p>应收应付、付款凭证、利润表、成本中心、对账</p>
                    <p class="page-count">24 页面</p>
                </div>
                <div class="module-card">
                    <h4><i class="fas fa-shopping-cart"></i> 采购管理 (purchase)</h4>
                    <p>采购订单、收货入库、品牌定香、供应商管理</p>
                    <p class="page-count">21 页面</p>
                </div>
                <div class="module-card">
                    <h4><i class="fas fa-industry"></i> 生产中心 (prodcenter)</h4>
                    <p>生产排程、质量控制、成品入库、物料需求</p>
                    <p class="page-count">11 页面</p>
                </div>
                <div class="module-card">
                    <h4><i class="fas fa-flask"></i> 技术中心 (techcenter)</h4>
                    <p>配方管理、打样试香、原料清单</p>
                    <p class="page-count">11 页面</p>
                </div>
                <div class="module-card">
                    <h4><i class="fas fa-vial"></i> 半成品管理 (semifinished)</h4>
                    <p>半成品库存、调配记录、状态追踪</p>
                    <p class="page-count">10 页面</p>
                </div>
                <div class="module-card">
                    <h4><i class="fas fa-truck"></i> 物流管理 (logistics)</h4>
                    <p>发货确认、运输追踪、退货处理、运费管理</p>
                    <p class="page-count">8 页面</p>
                </div>
                <div class="module-card">
                    <h4><i class="fas fa-warehouse"></i> 库存管理 (inventory)</h4>
                    <p>库存预警、库存变动、库存总览</p>
                    <p class="page-count">3 页面</p>
                </div>
                <div class="module-card">
                    <h4><i class="fas fa-shield-alt"></i> 系统管理 (system)</h4>
                    <p>用户管理、角色权限、安全审计、备份中心</p>
                    <p class="page-count">15 页面</p>
                </div>
            </div>
        </div>

        <!-- ========== 三、财务中心 ========== -->
        <div class="guide-section">
            <h2>三、财务中心流程</h2>
            
            <div class="guide-card">
                <h3>3.1 应付账款 (AP)</h3>
                <div class="flow-row">
                    <span class="flow-step">采购收货</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">应付列表 accounts_payable.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">创建付款凭证 payment_vouchers.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">记录总账 GLTransactions</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">对账 reconciliation.asp</span>
                </div>
            </div>

            <div class="guide-card">
                <h3>3.2 应收账款 (AR)</h3>
                <div class="flow-row">
                    <span class="flow-step">用户下单</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">应收列表 accounts_receivable.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">创建收款凭证 payment_vouchers.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">记录总账</span>
                </div>
            </div>

            <div class="guide-card">
                <h3>3.3 财务报表</h3>
                <div class="flow-row">
                    <span class="flow-step">成本中心设置 cost_centers.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">费用归集 expense_allocation.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">利润表 profit_report.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">综合报表 comprehensive_report.asp</span>
                </div>
                <table class="guide-table">
                    <tr><th>页面</th><th>功能说明</th></tr>
                    <tr><td>cost_centers.asp</td><td>成本中心CRUD（采购/生产/物流/市场/行政/研发）</td></tr>
                    <tr><td>expense_allocation.asp</td><td>费用分配到各成本中心</td></tr>
                    <tr><td>profit_report.asp</td><td>经营利润表：净销售额→毛利→边际贡献，日月周维度</td></tr>
                    <tr><td>cost_management.asp</td><td>成本管理与分析</td></tr>
                    <tr><td>comprehensive_report.asp</td><td>综合财务报表</td></tr>
                    <tr><td>cash_flow.asp</td><td>现金流管理</td></tr>
                    <tr><td>budget_management.asp</td><td>预算编制与追踪</td></tr>
                    <tr><td>gl_report.asp</td><td>总账报表</td></tr>
                    <tr><td>marketing_stats.asp</td><td>营销费用统计</td></tr>
                    <tr><td>product_analysis.asp</td><td>产品盈利分析</td></tr>
                    <tr><td>fund_dashboard.asp</td><td>资金大盘</td></tr>
                </table>
            </div>
        </div>

        <!-- ========== 四、采购中心 ========== -->
        <div class="guide-section">
            <h2>四、采购中心流程</h2>
            
            <div class="guide-card">
                <h3>4.1 标准采购</h3>
                <div class="flow-row">
                    <span class="flow-step">供应商管理 supplier_management.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">创建采购单 purchase_order_new.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">采购列表 purchase_orders.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">收货入库 batch_detail.asp</span>
                </div>
            </div>

            <div class="guide-card">
                <h3>4.2 品牌定香采购（FixedBrand）</h3>
                <div class="flow-row">
                    <span class="flow-step">品牌定香首页 fixed_brand/index.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">补货分析 replenishment.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">采购下单</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">收货入库</span>
                </div>
                <div class="note-warn">
                    <i class="fas fa-exclamation-triangle"></i> 品牌定香模块使用 FixedBrand* 系列独立表，与标准采购流程隔离。
                </div>
            </div>
        </div>

        <!-- ========== 五、运营中心 ========== -->
        <div class="guide-section">
            <h2>五、运营中心流程</h2>
            <div class="guide-card">
                <h3>5.1 订单处理</h3>
                <div class="flow-row">
                    <span class="flow-step">新订单审核</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">确认付款</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">安排生产/发货</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">物流追踪</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">完成订单</span>
                </div>
            </div>
        </div>

        <!-- ========== 六、生产中心 ========== -->
        <div class="guide-section">
            <h2>六、生产中心流程</h2>
            <div class="guide-card">
                <h3>6.1 生产管理</h3>
                <div class="flow-row">
                    <span class="flow-step">生产排程 prod_scheduling.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">领料出库</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">质量检测 prod_qc.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">成品入库 prod_warehouse.asp</span>
                </div>
            </div>
        </div>

        <!-- ========== 七、物流中心 ========== -->
        <div class="guide-section">
            <h2>七、物流中心流程</h2>
            <div class="guide-card">
                <h3>7.1 发货管理</h3>
                <div class="flow-row">
                    <span class="flow-step">待发货列表</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">生成运单</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">发货确认 delivery_confirm.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">在途追踪</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">签收确认</span>
                </div>
            </div>
        </div>

        <!-- ========== 八、系统管理 ========== -->
        <div class="guide-section">
            <h2>八、系统管理流程</h2>
            <div class="guide-card">
                <h3>8.1 用户与权限</h3>
                <table class="guide-table">
                    <tr><th>页面</th><th>功能</th></tr>
                    <tr><td>/admin/system/admins.asp</td><td>创建/编辑/禁用管理员账号</td></tr>
                    <tr><td>/admin/system/roles.asp</td><td>定义角色（SUPER_ADMIN / FIN_MANAGER / PURCHASE_MANAGER 等）</td></tr>
                    <tr><td>/admin/system/logs.asp</td><td>查看操作日志</td></tr>
                    <tr><td>/admin/system/security_audit.asp</td><td>安全审计与登录监控</td></tr>
                </table>
            </div>

            <div class="guide-card">
                <h3>8.2 数据备份</h3>
                <div class="flow-row">
                    <span class="flow-step">进入备份中心 backup_center.asp</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">检测修复（创建存储过程）</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">立即备份（生成 .bak）</span>
                    <span class="flow-arrow">→</span>
                    <span class="flow-step">下载备份文件</span>
                </div>
                <div class="note-warn">
                    <i class="fas fa-exclamation-triangle"></i> 如 ASP 无法执行备份，需在 SSMS 中手动运行：
                    <code>EXEC usp_BackupDatabase '路径','文件名','PerfumeShop'</code>
                </div>
            </div>
        </div>

        <!-- ========== 九、API接口 ========== -->
        <div class="guide-section">
            <h2>九、前端 API 接口</h2>
            <div class="guide-card">
                <table class="guide-table">
                    <tr><th>接口</th><th>功能</th></tr>
                    <tr><td>/api/cart_add.asp</td><td>添加商品到购物车</td></tr>
                    <tr><td>/api/cart_update.asp</td><td>更新购物车数量</td></tr>
                    <tr><td>/api/cart_remove.asp</td><td>移除购物车商品</td></tr>
                    <tr><td>/api/cart_count.asp</td><td>获取购物车数量</td></tr>
                    <tr><td>/api/cart_clear.asp</td><td>清空购物车</td></tr>
                    <tr><td>/api/favorites.asp</td><td>收藏/取消收藏</td></tr>
                    <tr><td>/api/order_confirm.asp</td><td>确认订单</td></tr>
                    <tr><td>/api/order_cancel.asp</td><td>取消订单</td></tr>
                    <tr><td>/api/track.asp</td><td>物流追踪查询</td></tr>
                    <tr><td>/api/upload.asp</td><td>文件上传</td></tr>
                    <tr><td>/api/risk_check.asp</td><td>风控检查</td></tr>
                    <tr><td>/api/get_areas.asp</td><td>获取地区数据</td></tr>
                </table>
            </div>
        </div>

        <!-- ========== 十、技术架构 ========== -->
        <div class="guide-section">
            <h2>十、技术架构要点</h2>
            <div class="guide-card">
                <table class="guide-table">
                    <tr><th>组件</th><th>技术栈</th></tr>
                    <tr><td>后端语言</td><td>Classic ASP (VBScript), CodePage 65001 (UTF-8)</td></tr>
                    <tr><td>数据库</td><td>SQL Server 2017 (localhost\YOURPERFUME, 97张表)</td></tr>
                    <tr><td>数据库连接</td><td>SQLOLEDB Provider, Integrated Security (SSPI)</td></tr>
                    <tr><td>前端样式</td><td>CSS 自定义属性 (design-tokens.css), Font Awesome 6</td></tr>
                    <tr><td>安全机制</td><td>CSRF 令牌, Cookie 签名, 登录速率限制, 安全响应头</td></tr>
                    <tr><td>备份方式</td><td>SQL Server BACKUP DATABASE (完整备份), PowerShell 自动化</td></tr>
                </table>
            </div>

            <div class="guide-card">
                <h3>重要开发注意事项</h3>
                <div class="note-warn" style="margin-top:0;">
                    <p><strong>1.</strong> 列表查询使用 <code>conn.Execute()</code> 返回只进游标，<code>MoveLast / RecordCount</code> 需要用客户端游标 (<code>CursorLocation=3</code>)</p>
                    <p style="margin-top:8px;"><strong>2.</strong> IIS 应用池身份无 DDL 权限，CREATE TABLE / ALTER TABLE 需通过 sqlcmd 预执行</p>
                    <p style="margin-top:8px;"><strong>3.</strong> sqlcmd 管道传输含中文路径会乱码，备份时先用英文暂存路径再复制</p>
                </div>
            </div>
        </div>

        <div class="guide-footer">
            <%= SITE_NAME %> · <%= SYS_VERSION %> · 完整使用流程指南<br>
            数据库备份: PerfumeShop_full_20260610_145134.bak (13.11 MB) | 代码备份: PerfumeShop_code_20260610_145415.zip (13.25 MB)
        </div>
    </div>
</body>
</html>
