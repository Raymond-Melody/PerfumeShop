<%@ Language="VBScript" CodePage="65001" %>
<% Response.CodePage = 65001 : Response.CharSet = "UTF-8" %>
<%
' 香氛定制电商系统 V8 使用逻辑说明书
' V8版本核心功能模块覆盖：运营管理 / 产品技术中心 / 采购管理 / 生产中心 / 半成品管理 / 物流管理 / 财务管理 / 系统管理
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>香氛定制电商系统 V8 使用逻辑说明书</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: "Microsoft YaHei", "PingFang SC", "Hiragino Sans GB", sans-serif;
            line-height: 1.8;
            color: #333;
            background-color: #f5f5f5;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background-color: #fff;
            min-height: 100vh;
            box-shadow: 0 0 20px rgba(0,0,0,0.1);
        }
        .header {
            background: linear-gradient(135deg, #1a1a2e 0%, #8B4513 100%);
            color: white;
            padding: 40px 30px;
            text-align: center;
        }
        .header h1 { font-size: 2.5em; margin-bottom: 10px; text-shadow: 2px 2px 4px rgba(0,0,0,0.3); }
        .header .subtitle { font-size: 1.1em; opacity: 0.9; }
        .header .version-badge {
            display: inline-block;
            background: #00bcd4;
            color: #fff;
            padding: 4px 16px;
            border-radius: 20px;
            font-size: 0.9em;
            margin-top: 10px;
        }
        .nav {
            background-color: #1a1a2e;
            padding: 15px 30px;
            position: sticky;
            top: 0;
            z-index: 100;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
        .nav ul { list-style: none; display: flex; flex-wrap: wrap; justify-content: center; gap: 10px; }
        .nav li a {
            color: white;
            text-decoration: none;
            padding: 8px 15px;
            border-radius: 4px;
            transition: all 0.3s;
            font-size: 0.95em;
        }
        .nav li a:hover { background-color: #8B4513; }
        .content { padding: 30px; }
        .section { margin-bottom: 50px; padding-bottom: 40px; border-bottom: 2px solid #e0e0e0; }
        .section:last-child { border-bottom: none; }
        .section-title {
            color: #1a1a2e;
            font-size: 1.8em;
            margin-bottom: 25px;
            padding-bottom: 10px;
            border-bottom: 3px solid #8B4513;
            display: flex;
            align-items: center;
        }
        .section-title .icon {
            background-color: #8B4513;
            color: white;
            width: 40px;
            height: 40px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-right: 15px;
            font-weight: bold;
        }
        .subsection { margin: 25px 0; }
        .subsection-title {
            color: #8B4513;
            font-size: 1.3em;
            margin-bottom: 15px;
            padding-left: 15px;
            border-left: 4px solid #8B4513;
        }
        h4 { color: #34495e; font-size: 1.1em; margin: 20px 0 10px 0; }
        p { margin-bottom: 12px; text-align: justify; }
        ul, ol { margin: 15px 0 15px 30px; }
        li { margin-bottom: 8px; }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            font-size: 0.95em;
        }
        th {
            background-color: #1a1a2e;
            color: white;
            padding: 12px 15px;
            text-align: left;
            font-weight: 600;
        }
        td { padding: 12px 15px; border-bottom: 1px solid #ddd; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        tr:hover { background-color: #f0f0f0; }
        .code-block {
            background-color: #f4f4f4;
            border-left: 4px solid #8B4513;
            padding: 15px 20px;
            margin: 15px 0;
            font-family: "Consolas", "Monaco", monospace;
            font-size: 0.9em;
            overflow-x: auto;
            border-radius: 0 4px 4px 0;
        }
        .info-box { background-color: #e8f4f8; border-left: 4px solid #3498db; padding: 15px 20px; margin: 15px 0; border-radius: 0 4px 4px 0; }
        .warning-box { background-color: #fff3cd; border-left: 4px solid #ffc107; padding: 15px 20px; margin: 15px 0; border-radius: 0 4px 4px 0; }
        .tip-box { background-color: #d4edda; border-left: 4px solid #28a745; padding: 15px 20px; margin: 15px 0; border-radius: 0 4px 4px 0; }
        .step-list { counter-reset: step; list-style: none; margin-left: 0; }
        .step-list li {
            position: relative;
            padding-left: 50px;
            margin-bottom: 20px;
        }
        .step-list li::before {
            counter-increment: step;
            content: counter(step);
            position: absolute;
            left: 0;
            top: 0;
            width: 35px;
            height: 35px;
            background-color: #8B4513;
            color: white;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: bold;
        }
        .architecture-box {
            background: linear-gradient(135deg, #f5f7fa 0%, #e4e8ec 100%);
            border: 2px solid #8B4513;
            border-radius: 8px;
            padding: 25px;
            margin: 20px 0;
        }
        .arch-section { background-color: white; border: 1px solid #ddd; border-radius: 6px; padding: 15px; margin: 10px 0; }
        .arch-title { font-weight: bold; color: #1a1a2e; margin-bottom: 10px; font-size: 1.1em; }
        .module-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 15px; margin: 20px 0; }
        .module-card {
            background: white;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            padding: 18px;
            transition: all 0.3s;
        }
        .module-card:hover { box-shadow: 0 4px 12px rgba(0,0,0,0.1); transform: translateY(-2px); }
        .module-card h4 { color: #8B4513; margin: 0 0 8px 0; font-size: 1em; }
        .module-card p { font-size: 0.9em; color: #666; margin: 0; }
        .status-tag {
            display: inline-block;
            padding: 3px 10px;
            border-radius: 12px;
            font-size: 0.85em;
            font-weight: 500;
        }
        .status-pending { background-color: #fff3cd; color: #856404; }
        .status-paid { background-color: #d1ecf1; color: #0c5460; }
        .status-processing { background-color: #cce5ff; color: #004085; }
        .status-shipped { background-color: #d4edda; color: #155724; }
        .status-completed { background-color: #c3e6cb; color: #155724; }
        .flow-diagram {
            background: #f8f9fa;
            border: 1px dashed #8B4513;
            border-radius: 8px;
            padding: 20px;
            margin: 15px 0;
            text-align: center;
            font-size: 1.1em;
        }
        .flow-diagram .arrow { color: #8B4513; margin: 0 10px; font-size: 1.3em; }
        .footer {
            background-color: #1a1a2e;
            color: white;
            text-align: center;
            padding: 25px;
            font-size: 0.9em;
        }
        .footer p { margin: 5px 0; text-align: center; }
        @media (max-width: 768px) {
            .header h1 { font-size: 1.8em; }
            .nav ul { flex-direction: column; align-items: center; }
            .content { padding: 20px 15px; }
            .section-title { font-size: 1.4em; }
            table { font-size: 0.85em; }
            th, td { padding: 8px 10px; }
        }
        strong { color: #8B4513; }
        em { color: #1a1a2e; font-style: normal; font-weight: 600; }
        .v8-badge { background: #00bcd4; color: #fff; padding: 2px 10px; border-radius: 10px; font-size: 0.75em; margin-left: 8px; }
        .v7-badge { background: #888; color: #fff; padding: 2px 10px; border-radius: 10px; font-size: 0.75em; margin-left: 8px; }
    </style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>香氛定制电商系统 V8</h1>
        <p class="subtitle">使用逻辑说明书</p>
        <span class="version-badge">V8.0.10 稳定版</span>
    </div>

    <nav class="nav">
        <ul>
            <li><a href="#section1">一、系统架构</a></li>
            <li><a href="#section2">二、前端商城</a></li>
            <li><a href="#section3">三、后台模块总览</a></li>
            <li><a href="#section4">四、运营管理中心</a></li>
            <li><a href="#section5">五、产品技术中心</a></li>
            <li><a href="#section6">六、采购管理中心</a></li>
            <li><a href="#section7">七、生产与半成品管理</a></li>
            <li><a href="#section8">八、物流管理中心</a></li>
            <li><a href="#section9">九、财务管理中心</a></li>
            <li><a href="#section10">十、系统管理中心</a></li>
            <li><a href="#section11">十一、权限角色体系</a></li>
            <li><a href="#section12">十二、核心业务流程</a></li>
            <li><a href="#section13">十三、成本引擎</a></li>
            <li><a href="#section14">十四、安全机制</a></li>
        </ul>
    </nav>

    <div class="content">
        <!-- ====== 第一部分：系统架构 ====== -->
        <section class="section" id="section1">
            <h2 class="section-title"><span class="icon">1</span>系统整体架构</h2>

            <div class="subsection">
                <h3 class="subsection-title">技术栈</h3>
                <ul>
                    <li><strong>后端技术：</strong>ASP Classic（VBScript） - 成熟稳定的服务端脚本技术</li>
                    <li><strong>数据库：</strong>SQL Server Express（本地） - 已从Access MDB迁移至SQL Server</li>
                    <li><strong>前端交互：</strong>jQuery + Chart.js（图表）</li>
                    <li><strong>页面技术：</strong>HTML5 + CSS3 + 暗色主题UI</li>
                </ul>
                <div class="tip-box">
                    <strong>数据库迁移说明：</strong>V8版本已完成从Access MDB到SQL Server Express的数据库迁移，<br>
                    连接字符串在 <code>includes/connection.asp</code> 中配置（SQLOLEDB + Integrated Security=SSPI）。<br>
                    数据库名称：<code>PerfumeShop</code>，位于 <code>localhost\SQLEXPRESS</code>。
                </div>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">整体架构图</h3>
                <div class="architecture-box">
                    <div class="arch-section">
                        <div class="arch-title">前端商城（用户端）</div>
                        <p>面向终端消费者的购物平台：首页/商品浏览/商品详情/<br>定制香水/购物车/结算/用户中心</p>
                    </div>
                    <div style="text-align: center; margin: 10px 0; color: #8B4513; font-size: 1.5em;">&#8593;</div>
                    <div class="arch-section">
                        <div class="arch-title">API接口层</div>
                        <p>购物车操作(cart_*)、收藏夹、订单操作、地区查询、文件上传、风险管理</p>
                    </div>
                    <div style="text-align: center; margin: 10px 0; color: #8B4513; font-size: 1.5em;">&#8593;</div>
                    <div class="arch-section">
                        <div class="arch-title">共享服务层（includes/）</div>
                        <p>config / connection / cost_engine / payment_handler / promotion_engine / recommendation_engine / product_type_utils / share_utils / tracking_utils / upload_utils / email_utils / password_utils / error_handler / member_utils</p>
                    </div>
                    <div style="text-align: center; margin: 10px 0; color: #8B4513; font-size: 1.5em;">&#8593;</div>
                    <div class="arch-section">
                        <div class="arch-title">八大后台管理中心</div>
                        <ul>
                            <li><strong>运营管理</strong> - 订单/客户/商品/香调/营销</li>
                            <li><strong>产品技术中心</strong> <span class="v8-badge">V8新增</span> - 配方/香调/基香/瓶型/KOL审核</li>
                            <li><strong>采购管理中心</strong> <span class="v8-badge">V8新增</span> - 采购订单/供应商/价格管理</li>
                            <li><strong>生产中心</strong> <span class="v8-badge">V8新增</span> - 排产/质检/仓库/车间</li>
                            <li><strong>半成品管理</strong> <span class="v8-badge">V8新增</span> - Accord生产/基香生产/原料库存</li>
                            <li><strong>物流管理</strong> <span class="v8-badge">V8新增</span> - 发货/在途/签收/退换</li>
                            <li><strong>财务管理</strong> - 资金看板/成本/利润/预算/风控</li>
                            <li><strong>系统管理</strong> - 角色/管理员/日志/安全审计</li>
                        </ul>
                    </div>
                </div>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">数据库核心表</h3>
                <table>
                    <tr><th>表名</th><th>说明</th><th>V8新增/变更</th></tr>
                    <tr><td>Users</td><td>前台用户表</td><td>—</td></tr>
                    <tr><td>Products</td><td>商品表（固定品牌/定制/KOL三类）</td><td>新增 Engravable/EngravingPrice 字段</td></tr>
                    <tr><td>ProductTypeConfig</td><td>产品类型配置表</td><td>V8新增</td></tr>
                    <tr><td>FragranceNotes</td><td>香调表（前/中/后调）</td><td>—</td></tr>
                    <tr><td>BaseNotes</td><td>基香成分表</td><td>—</td></tr>
                    <tr><td>ProductNoteRatios</td><td>产品香调配比表（KOL预设）</td><td>V8加强</td></tr>
                    <tr><td>BottleStyles</td><td>瓶型表</td><td>—</td></tr>
                    <tr><td>ProductBottleStyles</td><td>产品瓶型关联表（多选+独立定价）</td><td>V8新增</td></tr>
                    <tr><td>Cart</td><td>购物车表</td><td>—</td></tr>
                    <tr><td>Orders</td><td>订单主表</td><td>新增 ShippingStatus/ShippingCompany/TrackingNumber</td></tr>
                    <tr><td>OrderDetails</td><td>订单明细表</td><td>—</td></tr>
                    <tr><td>OrderItems</td><td>订单项扩展表</td><td>V8新增</td></tr>
                    <tr><td>AdminUsers</td><td>管理员表</td><td>新增 Department 字段</td></tr>
                    <tr><td>AdminRoles</td><td>角色表</td><td>新增 ModuleAccess 字段</td></tr>
                    <tr><td>RolePermissions</td><td>操作级权限表</td><td>V8新增</td></tr>
                    <tr><td>AdminLogs</td><td>操作日志表</td><td>—</td></tr>
                    <tr><td>SiteSettings</td><td>站点配置表</td><td>新增安全策略字段</td></tr>
                    <tr><td>SupplierPrices</td><td>供应商价格表</td><td>V8新增</td></tr>
                    <tr><td>PurchaseOrders</td><td>采购订单表</td><td>V8新增（5张采购相关表）</td></tr>
                    <tr><td>ProductCosts</td><td>产品成本表</td><td>V8新增</td></tr>
                    <tr><td>CostCenterBudgets</td><td>成本中心预算表</td><td>V8新增</td></tr>
                    <tr><td>BudgetItems</td><td>预算明细表</td><td>V8新增</td></tr>
                    <tr><td>ReconciliationLogs</td><td>对账日志表</td><td>V8新增</td></tr>
                    <tr><td>RefundRecords</td><td>退款记录表</td><td>V8新增</td></tr>
                    <tr><td>ProductionOrders</td><td>生产订单表</td><td>V8优化</td></tr>
                    <tr><td>Shipments</td><td>发货记录表</td><td>V8新增</td></tr>
                    <tr><td>Recipes</td><td>配方表</td><td>V8新增</td></tr>
                    <tr><td>RecipeAccords</td><td>配方Accord关联表</td><td>V8新增</td></tr>
                    <tr><td>RecipeProducts</td><td>配方产品关联表</td><td>V8新增</td></tr>
                    <tr><td>RecipePublishLog</td><td>配方发布日志表</td><td>V8新增</td></tr>
                </table>
            </div>
        </section>

        <!-- ====== 第二部分：前端商城 ====== -->
        <section class="section" id="section2">
            <h2 class="section-title"><span class="icon">2</span>前端商城功能</h2>

            <div class="subsection">
                <h3 class="subsection-title">功能模块清单</h3>
                <table>
                    <tr><th>模块</th><th>功能说明</th><th>文件</th></tr>
                    <tr><td>首页</td><td>三栏商品展示（固定品牌/定制/KOL），轮播横幅，特色服务，推荐引擎</td><td>index.asp</td></tr>
                    <tr><td>商品列表</td><td>分类筛选/关键词搜索/价格区间/排序</td><td>products.asp</td></tr>
                    <tr><td>商品详情</td><td>三种展示模式：固定品牌配基香 / 定制选择香调容量瓶型 / KOL预设配比</td><td>product.asp</td></tr>
                    <tr><td>定制香水</td><td>前调/中调/后调选择 + 配比输入 + 容量瓶型选择 + 刻字服务</td><td>customize.asp</td></tr>
                    <tr><td>购物车</td><td>增删改查/全选单选/金额计算/KOL配比实时统计</td><td>cart.asp</td></tr>
                    <tr><td>结算</td><td>地址选择/支付方式/订单确认/成分汇总/刻字明细</td><td>checkout.asp</td></tr>
                    <tr><td>订单成功</td><td>完成支付后的成功提示</td><td>order_success.asp</td></tr>
                    <tr><td>用户中心</td><td>订单管理/个人信息/地址管理/收藏夹/密码修改</td><td>user/</td></tr>
                    <tr><td>登录注册</td><td>用户注册/登录/忘记密码</td><td>user/login.asp / register.asp</td></tr>
                </table>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">商品详情页三种展示模式</h3>
                <table>
                    <tr><th>模式</th><th>特点</th><th>适用商品</th></tr>
                    <tr><td>固定品牌配基香</td><td>展示品牌信息、基香成分、固定价格，不可定制</td><td>品牌合作香水</td></tr>
                    <tr><td>定制香水模式</td><td>选择前/中/后调、设置配比百分比、容量瓶型、刻字服务</td><td>个性化定制香水</td></tr>
                    <tr><td>KOL推荐模式</td><td>展示KOL信息、预设香调配比（可微调）、推荐语、价格构成</td><td>KOL合作推荐商品</td></tr>
                </table>
                <div class="info-box">
                    <strong>V8配比验证增强：</strong>定制和KOL产品支持前后端双重配比校验，包括总和100%校验和各香调的最小比例校验，<br>
                    最小比例值从 SiteSettings 表动态读取，可在产品技术中心配置。
                </div>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">刻字服务 V8增强</h3>
                <p>V8版本对所有产品类型（固定品牌/定制/KOL）全面开放刻字选项：</p>
                <ul>
                    <li><strong>Engravable字段：</strong>每个产品独立控制是否支持刻字</li>
                    <li><strong>EngravingPrice字段：</strong>配置刻字附加费用</li>
                    <li><strong>前端显示：</strong>产品详情页/购物车/结算页三页面独立显示刻字费用明细行</li>
                    <li><strong>价格计算：</strong>刻字费用实时计入总价（product.asp + cart_add.asp）</li>
                </ul>
            </div>
        </section>

        <!-- ====== 第三部分：后台模块总览 ====== -->
        <section class="section" id="section3">
            <h2 class="section-title"><span class="icon">3</span>后台管理中心总览</h2>

            <div class="subsection">
                <h3 class="subsection-title">八大管理中心</h3>
                <table>
                    <tr><th>管理中心</th><th>路径</th><th>功能定位</th><th>适用角色</th></tr>
                    <tr><td>运营管理中心</td><td>/admin/operation/</td><td>订单/客户/商品/香调/营销/支付开关</td><td>OP_MANAGER, OP_STAFF, CONTENT_ADMIN</td></tr>
                    <tr><td>产品技术中心 <span class="v8-badge">V8新增</span></td><td>/admin/techcenter/</td><td>配方发布/香调基香/瓶型管理/KOL审核/产品设置</td><td>TECH_ADMIN, TECH_STAFF</td></tr>
                    <tr><td>采购管理中心 <span class="v8-badge">V8新增</span></td><td>/admin/purchase/</td><td>采购订单/供应商/价格/收货</td><td>PURCHASE_MANAGER, PURCHASE_STAFF</td></tr>
                    <tr><td>生产中心 <span class="v8-badge">V8新增</span></td><td>/admin/prodcenter/</td><td>排产/质检/仓库/车间/成品库存</td><td>PROD_MANAGER, PROD_STAFF</td></tr>
                    <tr><td>半成品管理 <span class="v8-badge">V8新增</span></td><td>/admin/semifinished/</td><td>Accord生产/基香生产/原料库存/出库/车间流转</td><td>PROD_MANAGER, PROD_STAFF</td></tr>
                    <tr><td>物流管理中心 <span class="v8-badge">V8新增</span></td><td>/admin/logistics/</td><td>发货/在途跟踪/签收确认/退货</td><td>PROD_MANAGER, PROD_STAFF</td></tr>
                    <tr><td>财务管理中心</td><td>/admin/finance/</td><td>资金看板/成本/利润/预算/对账/风控</td><td>FIN_MANAGER, FIN_STAFF</td></tr>
                    <tr><td>系统管理中心</td><td>/admin/system/</td><td>角色/管理员/日志/安全/配置/备份</td><td>SUPER_ADMIN</td></tr>
                </table>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">统一入口页（portal.asp）</h3>
                <p>登录 <code>/admin/login.asp</code> 后进入 portal.asp 统一入口页：</p>
                <ul>
                    <li>系统根据管理员角色显示可访问的模块入口卡片</li>
                    <li><strong>SUPER_ADMIN</strong> 显示全部八个后台入口</li>
                    <li>单模块角色自动跳转到对应后台首页</li>
                    <li>角色分组规则：PROD系列可访问生产中心+半成品+物流</li>
                </ul>
            </div>
        </section>

        <!-- ====== 第四部分：运营管理中心 ====== -->
        <section class="section" id="section4">
            <h2 class="section-title"><span class="icon">4</span>运营管理中心</h2>
            <p>路径：<code>/admin/operation/</code> | 适用角色：OP_MANAGER, OP_STAFF, CONTENT_ADMIN</p>

            <div class="subsection">
                <h3 class="subsection-title">核心功能页面</h3>
                <table>
                    <tr><th>页面</th><th>功能</th><th>权限校验</th></tr>
                    <tr><td>index.asp</td><td>运营概览：今日订单/营收/新客/在售商品统计卡片</td><td>通用</td></tr>
                    <tr><td>orders.asp</td><td>订单列表：按状态/日期/关键词筛选，状态流转操作</td><td>OP_MANAGER, OP_STAFF</td></tr>
                    <tr><td>order_detail.asp</td><td>订单详情：完整信息/商品明细/收货地址/支付信息</td><td>OP_MANAGER, OP_STAFF</td></tr>
                    <tr><td>order_edit.asp</td><td>订单编辑：修改配送信息和状态</td><td>OP_MANAGER</td></tr>
                    <tr><td>customers.asp</td><td>客户列表：分页/搜索/状态管理</td><td>OP_MANAGER</td></tr>
                    <tr><td>customer_detail.asp</td><td>客户详情：基本信息/订单历史/消费统计</td><td>OP_MANAGER</td></tr>
                    <tr><td>product_types.asp</td><td>产品类型配置：三类商品类型的CRUD+前台显示控制</td><td>CONTENT_ADMIN</td></tr>
                    <tr><td>products.asp</td><td>商品列表：新增/编辑/上下架/前台状态显示</td><td>CONTENT_ADMIN</td></tr>
                    <tr><td>fragrances.asp</td><td>香调配置：前/中/后调CRUD+附加价格+建议配比</td><td>CONTENT_ADMIN</td></tr>
                    <tr><td>base_notes.asp</td><td>基香管理：基香成分CRUD+启用/禁用+香调关联</td><td>CONTENT_ADMIN</td></tr>
                    <tr><td>reviews_manage.asp</td><td>运营评审管理（跳转至技术中心）</td><td>OP_MANAGER</td></tr>
                    <tr><td>order_reviews.asp</td><td>订单评价管理</td><td>OP_MANAGER</td></tr>
                    <tr><td>campaign_edit.asp</td><td>营销活动编辑</td><td>OP_MANAGER</td></tr>
                    <tr><td>coupons.asp</td><td>优惠券管理：创建/使用条件/发放</td><td>OP_MANAGER</td></tr>
                    <tr><td>points.asp</td><td>积分管理：规则配置/积分调整/兑换管理</td><td>OP_MANAGER</td></tr>
                    <tr><td>payment_switch.asp</td><td>支付开关：控制支付方式在前台的显示/隐藏</td><td>OP_MANAGER</td></tr>
                    <tr><td>marketing.asp</td><td>营销活动列表</td><td>OP_MANAGER</td></tr>
                    <tr><td>after_sales.asp</td><td>售后管理：退货/退款处理</td><td>OP_MANAGER</td></tr>
                    <tr><td>content_pages.asp</td><td>内容页面管理</td><td>OP_MANAGER</td></tr>
                    <tr><td>performance_dashboard.asp</td><td>绩效看板</td><td>OP_MANAGER</td></tr>
                </table>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">订单状态流转</h3>
                <div class="flow-diagram">
                    <span class="status-tag status-pending">Pending</span> 待付款
                    <span class="arrow">&#8594;</span>
                    <span class="status-tag status-paid">Paid</span> 已支付
                    <span class="arrow">&#8594;</span>
                    <span class="status-tag status-processing">Processing</span> 处理中
                    <span class="arrow">&#8594;</span>
                    <span class="status-tag status-shipped">Shipped</span> 已发货
                    <span class="arrow">&#8594;</span>
                    <span class="status-tag status-completed">Completed</span> 已完成
                </div>
                <p>生产流程自动联动：订单支付后自动创建生产订单。</p>
            </div>
        </section>

        <!-- ====== 第五部分：产品技术中心 ====== -->
        <section class="section" id="section5">
            <h2 class="section-title"><span class="icon">5</span>产品技术中心 <span class="v8-badge">V8新增</span></h2>
            <p>路径：<code>/admin/techcenter/</code> | 适用角色：TECH_ADMIN, TECH_STAFF, SUPER_ADMIN</p>
            <div class="info-box">
                <strong>功能定位：</strong>产品技术中心是V8新增的核心模块，负责产品配方管理、香调基香配置、瓶型管理、KOL审核和产品设置。<br>
                它将原先分散在运营中心的配方/技术相关功能独立为专门的研发管理模块。
            </div>

            <div class="subsection">
                <h3 class="subsection-title">功能页面</h3>
                <table>
                    <tr><th>页面</th><th>功能</th></tr>
                    <tr><td>index.asp</td><td>技术概览：产品/香调/基香/配方统计，配方发布状态看板</td></tr>
                    <tr><td>formula_management.asp</td><td>配方管理（含基香多选）：Recipes表CRUD + NoteIngredients关联 + ReviewStatus审核状态</td></tr>
                    <tr><td>recipe_publish.asp</td><td>配方发布：配方审核后发布到供应链（RecipeAccords + RecipeProducts + RecipePublishLog）</td></tr>
                    <tr><td>note_management.asp</td><td>香调管理：FragranceNotes CRUD + 类型设置 + 基香关联（NoteIngredients复选框组+批量关联）</td></tr>
                    <tr><td>base_note_management.asp</td><td>基香管理：BaseNotes CRUD + 启用/禁用 + 配方用量配置</td></tr>
                    <tr><td>bottle_management.asp</td><td>瓶型管理：BottleStyles CRUD + Chart.js使用统计 + 产品关联查看</td></tr>
                    <tr><td>product_settings.asp</td><td>产品设置：产品CRUD + 香调/基香/刻字/瓶型多选配置 + 价格设置 + KOL配比编辑</td></tr>
                    <tr><td>kol_reviews.asp</td><td>KOL产品审核：待审核列表 + 详情弹窗 + 审核通过/拒绝</td></tr>
                    <tr><td>check_ratio_settings.asp</td><td>香调配比参数查看（最小比例动态配置入口）</td></tr>
                    <tr><td>create_formula_tables.asp</td><td>配方表初始化工具（首次使用前执行）</td></tr>
                    <tr><td>create_product_bottles_table.asp</td><td>产品-瓶型关联表初始化工具</td></tr>
                </table>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">配方管理逻辑</h3>
                <p>配方模块遵循完整的生命周期管理：</p>
                <div class="flow-diagram">
                    草稿(Draft) <span class="arrow">&#8594;</span> 待审核(Pending) <span class="arrow">&#8594;</span> 已审核(Approved) <span class="arrow">&#8594;</span> 已发布(Published)
                </div>
                <ul>
                    <li><strong>基香多选：</strong>配方关联多个基香成分，通过 NoteIngredients 表实现批量关联</li>
                    <li><strong>配方发布：</strong>已审核配方可发布至供应链，生成 RecipeAccords/RecipeProducts 记录，并记录到 RecipePublishLog</li>
                    <li><strong>配方统计：</strong>全部计数、按状态分类统计</li>
                </ul>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">瓶型多选与独立定价</h3>
                <ul>
                    <li><strong>多选关联：</strong>一个产品可关联多个瓶型（需先访问 create_product_bottles_table.asp 初始化表）</li>
                    <li><strong>独立定价：</strong>每个关联的瓶型可设置独立 CustomPrice，优先使用自定义价格，无则使用瓶型默认价格</li>
                    <li><strong>前端显示：</strong>仅展示产品关联的瓶型，未关联的瓶型不显示</li>
                    <li><strong>兼容性：</strong>旧产品无关联数据时自动使用全部可用瓶型兜底</li>
                    </ul>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">香调配比最小比例动态配置</h3>
                <p>V9.0.10功能迁移至产品技术中心：</p>
                <ul>
                    <li>在站点设置中配置前调/中调/后调的最小配比百分比（minTopPercent/minMiddlePercent/minBasePercent）</li>
                    <li>产品技术中心新增「香调配比参数」Tab，可集中查看和跳转配置</li>
                    <li>前后端双重校验：前端JS实时提示 + 后端VBScript验证</li>
                    <li>配比总和必须等于100%，每个香调不低于最小比例阈值</li>
                </ul>
            </div>
        </section>

        <!-- ====== 第六部分：采购管理中心 ====== -->
        <section class="section" id="section6">
            <h2 class="section-title"><span class="icon">6</span>采购管理中心 <span class="v8-badge">V8新增</span></h2>
            <p>路径：<code>/admin/purchase/</code> | 适用角色：PURCHASE_MANAGER, PURCHASE_STAFF</p>

            <div class="subsection">
                <h3 class="subsection-title">功能页面</h3>
                <table>
                    <tr><th>页面</th><th>功能</th></tr>
                    <tr><td>index.asp</td><td>采购概览：按品类（原料/包装/瓶子）的统计卡片、待审批/月度总额/供应商数/待收货</td></tr>
                    <tr><td>purchase_orders.asp</td><td>采购订单管理：完整CRUD + 状态流转（草稿→已提交→已审核→已下单→部分收货→已收货→已完成）</td></tr>
                    <tr><td>supplier_management.asp</td><td>供应商管理：CRUD + 联系人/联系方式/供应品类/评级/活跃状态</td></tr>
                    <tr><td>price_management.asp</td><td>价格管理：供应商价格表管理 + 最新价/历史价对比 + 批量更新</td></tr>
                    <tr><td>receiving.asp</td><td>收货管理：采购到货登记 + 质检确认 + 入库</td></tr>
                    <tr><td>bottle_purchase.asp</td><td>瓶子采购：专门针对BottleStyles品类的采购单管理</td></tr>
                    <tr><td>packaging_purchase.asp</td><td>包装物采购：专门针对包装品类采购单管理</td></tr>
                    <tr><td>base_note_receiving.asp</td><td>基香原料收货：基香原料到货处理</td></tr>
                </table>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">采购订单状态流转</h3>
                <div class="flow-diagram">
                    草稿(Draft) <span class="arrow">&#8594;</span> 已提交(Submitted) <span class="arrow">&#8594;</span> 已审核(Approved)
                    <span class="arrow">&#8594;</span> 已下单(Ordered) <span class="arrow">&#8594;</span> 部分收货(PartialReceived)
                    <span class="arrow">&#8594;</span> 已收货(Received) <span class="arrow">&#8594;</span> 已完成(Completed)
                </div>
                <ul>
                    <li><strong>采购品类：</strong>RawMaterial（原料）、Packaging（包装物）、Bottle（瓶子）三种</li>
                    <li><strong>价格引用：</strong>从 SupplierPrices 表获取最新供应商报价</li>
                    <li><strong>财务集成：</strong>采购审核环节与财务中心关联，支持采购审核集成</li>
                </ul>
            </div>
        </section>

        <!-- ====== 第七部分：生产与半成品管理 ====== -->
        <section class="section" id="section7">
            <h2 class="section-title"><span class="icon">7</span>生产中心与半成品管理 <span class="v8-badge">V8拆分重构</span></h2>
            <p>V8将原先的 <code>admin/production/</code> 模块拆分为 <code>admin/prodcenter/</code>（生产中心）和 <code>admin/semifinished/</code>（半成品管理）两个独立模块。</p>

            <div class="subsection">
                <h3 class="subsection-title">7.1 生产中心（/admin/prodcenter/）</h3>
                <table>
                    <tr><th>页面</th><th>功能</th></tr>
                    <tr><td>index.asp</td><td>生产概览：待排产/生产中/已完成统计</td></tr>
                    <tr><td>order_production.asp</td><td>订单生产管理：订单→生产任务创建，成分检查</td></tr>
                    <tr><td>prod_scheduling.asp</td><td>排产管理：生产排期/产线分配/优先级设置</td></tr>
                    <tr><td>prod_qc.asp</td><td>质检管理：质量检查记录/合格/不合格处理</td></tr>
                    <tr><td>prod_warehouse.asp</td><td>成品仓库：成品入库/出库/库存管理</td></tr>
                    <tr><td>prod_workshop.asp</td><td>车间管理：车间作业/工单/流转</td></tr>
                    <tr><td>product_inventory.asp</td><td>成品库存：产成品库存查询/盘点</td></tr>
                    <tr><td>bottle_inventory.asp</td><td>瓶型库存：瓶型库存管理</td></tr>
                    <tr><td>packaging_inventory.asp</td><td>包装库存：包装物库存管理</td></tr>
                    <tr><td>production_management.asp</td><td>生产管理总看板：多维度汇总</td></tr>
                </table>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">7.2 半成品管理（/admin/semifinished/）</h3>
                <table>
                    <tr><th>页面</th><th>功能</th></tr>
                    <tr><td>index.asp</td><td>半成品概览</td></tr>
                    <tr><td>accord_production.asp</td><td>Accord（合香）生产：调香工序管理</td></tr>
                    <tr><td>base_note_production.asp</td><td>基香生产：基香制作工序</td></tr>
                    <tr><td>note_inventory.asp</td><td>香调库存：各香调库存数量管理</td></tr>
                    <tr><td>base_note_inventory.asp</td><td>基香库存：基香成分库存管理</td></tr>
                    <tr><td>raw_material_inventory.asp</td><td>原料库存：原材料库存管理 + 安全库存预警</td></tr>
                    <tr><td>material_outbound.asp</td><td>原料出库：领料管理/出库登记</td></tr>
                    <tr><td>workshop_transfer.asp</td><td>车间流转：半成品在不同车间之间流转</td></tr>
                    <tr><td>inventory_alerts.asp</td><td>库存预警：低库存品项红色提醒 + 快速入库</td></tr>
                    <tr><td>inventory_dashboard.asp</td><td>库存看板：可视化库存仪表板</td></tr>
                </table>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">生产业务流程</h3>
                <ol class="step-list">
                    <li><strong>订单到达：</strong>用户在商城下单并支付，系统自动或手动创建生产任务</li>
                    <li><strong>排产：</strong>生产管理员在生产中心进行排产，分配产线和优先级</li>
                    <li><strong>原料准备：</strong>半成品管理中的原料出库，从香调库存调拨原料到生产线</li>
                    <li><strong>半成品生产：</strong>Accord（合香）和 Base Note（基香）生产环节</li>
                    <li><strong>成品组装：</strong>香水成品灌装、包装</li>
                    <li><strong>质检：</strong>质量检查，合格品进入成品仓库</li>
                    <li><strong>入库待发：</strong>成品入库，进入物流发货流程</li>
                    <li><strong>完工：</strong>物流签收后自动完结生产订单</li>
                </ol>
            </div>
        </section>

        <!-- ====== 第八部分：物流管理中心 ====== -->
        <section class="section" id="section8">
            <h2 class="section-title"><span class="icon">8</span>物流管理中心 <span class="v8-badge">V8新增</span></h2>
            <p>路径：<code>/admin/logistics/</code> | 适用角色：PROD_MANAGER, PROD_STAFF</p>

            <div class="subsection">
                <h3 class="subsection-title">功能页面</h3>
                <table>
                    <tr><th>页面</th><th>功能</th></tr>
                    <tr><td>index.asp</td><td>物流概览：待发货/运输中/待签收/退货统计</td></tr>
                    <tr><td>shipping_orders.asp</td><td>发货管理：订单发货操作，填写物流公司和运单号</td></tr>
                    <tr><td>shipments.asp</td><td>运输中列表：已发货在途的订单追踪</td></tr>
                    <tr><td>in_transit.asp</td><td>在途跟踪：运输状态更新</td></tr>
                    <tr><td>delivery_confirm.asp</td><td>签收确认：确认客户已收到商品</td></tr>
                    <tr><td>returns.asp</td><td>退货管理：退货申请/审核/退货入库</td></tr>
                    <tr><td>shipping_companies.asp</td><td>物流公司管理：承运商CRUD</td></tr>
                    <tr><td>shipping_cost.asp</td><td>运费管理：运费计算规则/对账</td></tr>
                </table>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">物流状态流转</h3>
                <div class="flow-diagram">
                    待发货 <span class="arrow">&#8594;</span> 已发货(运输中) <span class="arrow">&#8594;</span> 已签收 <span class="arrow">&#8594;</span> 已完成
                </div>
                <p>物流状态与订单状态联动：已发货(Shipped) → 签收后自动转为已完成(Completed)。</p>
            </div>
        </section>

        <!-- ====== 第九部分：财务管理中心 ====== -->
        <section class="section" id="section9">
            <h2 class="section-title"><span class="icon">9</span>财务管理中心</h2>
            <p>路径：<code>/admin/finance/</code> | 适用角色：FIN_MANAGER, FIN_STAFF</p>

            <div class="subsection">
                <h3 class="subsection-title">V8财务模块概述</h3>
                <p>财务管理中心在V8中经历了多轮迭代（V8.0 ~ V8.0.10），已发展为三层架构的完整财务子系统：</p>
                <ul>
                    <li><strong>基础层：</strong>成本管理/费用分摊/对账中心</li>
                    <li><strong>分析层：</strong>利润表/单品分析/营销统计</li>
                    <li><strong>管理层：</strong>资金看板/预算管理/风控预警</li>
                </ul>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">功能清单</h3>
                <div class="module-grid">
                    <div class="module-card"><h4>资金看板</h4><p>fund_dashboard.asp - 资金流水总览/收支趋势/账户余额</p></div>
                    <div class="module-card"><h4>流水管理</h4><p>transactions.asp - 交易记录/流水列表/综合查询</p></div>
                    <div class="module-card"><h4>成本管理</h4><p>cost_management.asp - 产品成本/采购成本/间接成本</p></div>
                    <div class="module-card"><h4>费用分摊</h4><p>expense_allocation.asp - 间接费用分摊规则/执行</p></div>
                    <div class="module-card"><h4>对账中心</h4><p>reconciliation.asp - 订单对账/支付对账/差异处理</p></div>
                    <div class="module-card"><h4>利润表</h4><p>profit_report.asp - 经营利润表/月度利润趋势</p></div>
                    <div class="module-card"><h4>单品分析</h4><p>product_analysis.asp - 单品成本/利润/销量分析</p></div>
                    <div class="module-card"><h4>营销统计</h4><p>marketing_stats.asp - ROI分析/活动效果/投放产出</p></div>
                    <div class="module-card"><h4>预算管理</h4><p>budget_management.asp - 预算编制/执行监控/偏差分析</p></div>
                    <div class="module-card"><h4>风控管理</h4><p>risk_control.asp - 异常预警/成本异动/收支预警</p></div>
                    <div class="module-card"><h4>应收应付</h4><p>accounts_receivable.asp / accounts_payable.asp - 客户应收/供应商应付</p></div>
                    <div class="module-card"><h4>支付配置</h4><p>payment_config.asp - 支付通道参数/手续费率</p></div>
                    <div class="module-card"><h4>采购审核</h4><p>purchase_review.asp / purchase_review_detail.asp - 采购订单财务审核</p></div>
                    <div class="module-card"><h4>全面报表</h4><p>comprehensive_report.asp - 多维度财务综合分析报表</p></div>
                </div>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">财务关键指标</h3>
                <ul>
                    <li><strong>总营收：</strong>SUM(Orders.TotalAmount) WHERE Status IN ('Paid','Processing','Shipped','Completed')</li>
                    <li><strong>总成本：</strong>SUM(Orders.CostAmount)</li>
                    <li><strong>总利润：</strong>总营收 - 总成本 - 总退款（RefundRecords）</li>
                    <li><strong>利润率：</strong>利润/营收 * 100%，三色预警（绿>20% / 黄10-20% / 红<10%）</li>
                    <li><strong>成本异动预警：</strong>单品成本月度波动>5%标记预警</li>
                </ul>
            </div>
        </section>

        <!-- ====== 第十部分：系统管理中心 ====== -->
        <section class="section" id="section10">
            <h2 class="section-title"><span class="icon">10</span>系统管理中心</h2>
            <p>路径：<code>/admin/system/</code> | 适用角色：SUPER_ADMIN</p>

            <div class="subsection">
                <h3 class="subsection-title">功能页面</h3>
                <table>
                    <tr><th>页面</th><th>功能</th></tr>
                    <tr><td>index.asp</td><td>系统概览：管理员数/角色数/今日日志/系统状态</td></tr>
                    <tr><td>roles.asp</td><td>角色管理：8种预设角色的CRUD + 操作级权限（RolePermissions表）</td></tr>
                    <tr><td>admins.asp</td><td>管理员列表</td></tr>
                    <tr><td>admin_add.asp</td><td>新增管理员</td></tr>
                    <tr><td>admin_edit.asp</td><td>编辑管理员信息/角色分配/状态</td></tr>
                    <tr><td>admin_reset.asp</td><td>密码重置（生成随机密码）</td></tr>
                    <tr><td>logs.asp</td><td>操作日志查询：多条件筛选（管理员/操作类型/时间范围）</td></tr>
                    <tr><td>login_monitor.asp</td><td>登录监控：登录尝试记录/失败分析/异常IP追踪</td></tr>
                    <tr><td>ip_blacklist.asp</td><td>IP黑名单：IP封禁/白名单/自动解封规则</td></tr>
                    <tr><td>settings.asp</td><td>管理配置：安全策略（密码最小长度/Session超时/登录重试/MFA）</td></tr>
                    <tr><td>site_settings.asp</td><td>站点设置：首页栏目显隐(ShowFixed/Custom/KOL) + 香调配比最小值</td></tr>
                    <tr><td>backup_center.asp</td><td>备份中心：数据库手动备份/自动备份策略/备份列表管理</td></tr>
                    <tr><td>security_audit.asp</td><td>安全审计：安全事件列表/合规检查</td></tr>
                    <tr><td>statistics.asp</td><td>数据统计：销售趋势/热销TOP5/用户活跃/支付分布</td></tr>
                </table>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">安全策略配置项</h3>
                <table>
                    <tr><th>配置项</th><th>说明</th><th>默认值</th></tr>
                    <tr><td>Security_PasswordMinLength</td><td>密码最小长度</td><td>8</td></tr>
                    <tr><td>Security_SessionTimeout</td><td>Session超时时间（分钟）</td><td>30</td></tr>
                    <tr><td>Security_LoginMaxAttempts</td><td>登录最大重试次数</td><td>5</td></tr>
                    <tr><td>Security_MFAEnabled</td><td>是否启用MFA多因素认证</td><td>0</td></tr>
                </table>
            </div>
        </section>

        <!-- ====== 第十一部分：权限角色体系 ====== -->
        <section class="section" id="section11">
            <h2 class="section-title"><span class="icon">11</span>权限角色体系</h2>

            <div class="subsection">
                <h3 class="subsection-title">V8角色清单</h3>
                <table>
                    <tr><th>角色代码</th><th>角色名称</th><th>可访问后台</th><th>主要权限</th></tr>
                    <tr><td>SUPER_ADMIN</td><td>超级管理员</td><td>全部八个后台</td><td>完全控制所有功能，角色管理、管理员管理、安全配置</td></tr>
                    <tr><td>OP_MANAGER</td><td>运营经理</td><td>运营管理</td><td>订单/客户/商品/营销全部操作权限</td></tr>
                    <tr><td>OP_STAFF</td><td>运营专员</td><td>运营管理</td><td>基础运营操作（订单查看/客户查看/商品查看）</td></tr>
                    <tr><td>CONTENT_ADMIN</td><td>内容管理员</td><td>运营管理（部分）</td><td>商品/分类/香调内容管理，无订单操作权限</td></tr>
                    <tr><td>PROD_MANAGER</td><td>生产经理</td><td>生产中心+半成品+物流</td><td>排产/质检/仓库/物流全部操作权限</td></tr>
                    <tr><td>PROD_STAFF</td><td>生产专员</td><td>生产中心+半成品+物流</td><td>基础生产操作（查看库存/查看订单）</td></tr>
                    <tr><td>FIN_MANAGER</td><td>财务经理</td><td>财务管理</td><td>收入/报表/支付配置/成本/预算全部权限</td></tr>
                    <tr><td>FIN_STAFF</td><td>财务专员</td><td>财务管理</td><td>基础财务查看（报表查看/统计数据）</td></tr>
                    <tr><td>TECH_ADMIN</td><td>技术管理员</td><td>产品技术中心</td><td>配方/香调/基香/瓶型/KOL审核全部权限</td></tr>
                    <tr><td>TECH_STAFF</td><td>技术专员</td><td>产品技术中心</td><td>基础技术操作</td></tr>
                    <tr><td>PURCHASE_MANAGER</td><td>采购经理</td><td>采购管理中心</td><td>采购订单/供应商/价格全部权限</td></tr>
                    <tr><td>PURCHASE_STAFF</td><td>采购专员</td><td>采购管理中心</td><td>基础采购操作</td></tr>
                </table>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">权限验证机制</h3>
                <p>V8采用两级权限验证体系：</p>
                <ul>
                    <li><strong>模块级权限：</strong>由 <code>admin/includes/role_auth.asp</code> 的 VerifyModuleAccess() 函数控制</li>
                    <li><strong>操作级权限：</strong>通过 RolePermissions 表实现细粒度操作权限（Create/Edit/Delete/View/Export/Approve）</li>
                    <li><strong>超级管理员豁免：</strong>SUPER_ADMIN 角色在所有模块拥有完全权限</li>
                    <li><strong>越权审计：</strong>无权限访问尝试自动记录到 AdminLogs，并跳转到 unauthorized.asp</li>
                </ul>
                <div class="code-block">
' 权限验证示例 - 在每个后台页面顶部包含<br>
&lt;!--#include file="includes/auth.asp"--&gt;<br>
&lt;!--#include file="../../includes/config.asp"--&gt;<br>
&lt;!--#include file="../../includes/connection.asp"--&gt;<br><br>
Call VerifyModuleAccess("techcenter", "view")  ' 如果无权限则跳转
                </div>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">角色后台映射</h3>
                <table>
                    <tr><th>角色前缀</th><th>门户映射</th></tr>
                    <tr><td>SUPER_ADMIN</td><td>operation,semifinished,prodcenter,production,logistics,purchase,finance,techcenter,system</td></tr>
                    <tr><td>OP_*</td><td>operation</td></tr>
                    <tr><td>PROD_*</td><td>production,semifinished,prodcenter,logistics（四合一门户）</td></tr>
                    <tr><td>FIN_*</td><td>finance</td></tr>
                    <tr><td>TECH_*</td><td>techcenter</td></tr>
                    <tr><td>PURCHASE_*</td><td>purchase</td></tr>
                    <tr><td>CONTENT_ADMIN</td><td>operation（仅内容管理部分）</td></tr>
                </table>
            </div>
        </section>

        <!-- ====== 第十二部分：核心业务流程 ====== -->
        <section class="section" id="section12">
            <h2 class="section-title"><span class="icon">12</span>核心业务流程</h2>

            <div class="subsection">
                <h3 class="subsection-title">12.1 商品上架完整流程</h3>
                <ol class="step-list">
                    <li><strong>配置基香成分</strong> - 产品技术中心 → 基香管理，添加/确认基香成分，确保状态为启用</li>
                    <li><strong>配置香调</strong> - 产品技术中心 → 香调管理，设置类型（前/中/后调）、附加价格、基香关联</li>
                    <li><strong>配置瓶型</strong> - 产品技术中心 → 瓶型管理，添加瓶型样式和默认价格</li>
                    <li><strong>创建商品</strong> - 产品技术中心 → 产品设置，填写基本信息/关联香调基香/设置价格/配置瓶型多选</li>
                    <li><strong>上传图片</strong> - 使用统一上传功能（支持客户端压缩），图片上限由ASP 0104修复解除</li>
                    <li><strong>设置刻字选项</strong> - 在商品编辑中开启 Engravable 并设置刻字价格</li>
                    <li><strong>上下架控制</strong> - 设置 IsActive 状态控制前台显示</li>
                </ol>
                <div class="tip-box">
                    <strong>KOL商品需额外步骤：</strong>创建后为KOL商品配置预设香调配比（ProductNoteRatios），<br>
                    然后在技术中心的KOL审核页面完成审核（Approved），商品才能在前台展示。
                </div>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">12.2 订单处理全流程</h3>
                <ol class="step-list">
                    <li><strong>用户下单&支付</strong> - 前端用户完成选购和支付，订单状态变为Paid</li>
                    <li><strong>运营审核</strong> - 运营中心查看订单详情，确认订单信息无误</li>
                    <li><strong>生产触发</strong> - 系统自动/手动创建生产任务（生产中心 → 订单生产管理）</li>
                    <li><strong>排产</strong> - 生产中心进行排产，分配产线和优先级</li>
                    <li><strong>原料出库</strong> - 半成品管理 → 原料出库，领料</li>
                    <li><strong>生产制造</strong> - Accord生产 + 基香生产 + 成品灌装包装</li>
                    <li><strong>质检</strong> - 生产中心 → 质检管理，合格后入库</li>
                    <li><strong>发货</strong> - 物流管理中心 → 发货管理，填写物流公司/运单号</li>
                    <li><strong>运输中</strong> - 物流中心追踪在途包裹</li>
                    <li><strong>签收</strong> - 物流中心确认签收，订单自动完成</li>
                </ol>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">12.3 采购管理流程</h3>
                <ol class="step-list">
                    <li><strong>需求识别</strong> - 从库存预警/安全库存阈值触发采购需求</li>
                    <li><strong>创建采购单</strong> - 采购管理中心 → 采购订单，选择品类（原料/包装/瓶子）</li>
                    <li><strong>提交审批</strong> - 采购单状态变为 Submitted</li>
                    <li><strong>财务审核</strong> - 财务管理中心 → 采购审核，审核通过后状态变为 Approved</li>
                    <li><strong>下单</strong> - 向供应商下单，状态变为 Ordered</li>
                    <li><strong>收货</strong> - 到货后在收货管理登记，部分收货/全部收货</li>
                    <li><strong>入库</strong> - 自动更新库存，采购单完成</li>
                    <li><strong>财务结算</strong> - 应付账款管理</li>
                </ol>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">12.4 成本传导链路</h3>
                <p>成本从原材料端自动传导至订单端（基于 cost_engine.asp）：</p>
                <div class="flow-diagram">
                    原材料采购价格 <span class="arrow">&#8594;</span> 香调(Accord/Note)成本
                    <span class="arrow">&#8594;</span> 产品BOM成本 <span class="arrow">&#8594;</span> 订单成本/利润
                </div>
                <ul>
                    <li>使用缓存字典批量预加载，提升计算性能</li>
                    <li>支持按产品/订单维度的成本核算</li>
                    <li>与采购价格管理系统联动，自动获取最新采购价</li>
                </ul>
            </div>
        </section>

        <!-- ====== 第十三部分：成本引擎 ====== -->
        <section class="section" id="section13">
            <h2 class="section-title"><span class="icon">13</span>成本引擎（Cost Engine）</h2>

            <div class="subsection">
                <h3 class="subsection-title">架构说明</h3>
                <p>文件位置：<code>includes/cost_engine.asp</code>（784行）</p>
                <p>成本引擎实现原材料 → 香调(Accord/Note) → 产品(Product) → 订单(Order) 的三级成本自动计算与传导。</p>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">核心函数</h3>
                <table>
                    <tr><th>函数</th><th>功能</th></tr>
                    <tr><td>CE_PreloadAllCostData()</td><td>批量预加载所有成本参考数据到字典缓存</td></tr>
                    <tr><td>CE_GetCachedNoteCost(noteId)</td><td>获取缓存的香调成本</td></tr>
                    <tr><td>CE_GetCachedProductBOM(productId)</td><td>获取缓存的物料清单成本</td></tr>
                    <tr><td>CE_GetCachedProductUnitCost(productId)</td><td>获取缓存的产品单位成本</td></tr>
                    <tr><td>CE_CalculateNoteCost(noteId)</td><td>实时计算香调成本</td></tr>
                    <tr><td>CE_CalculateProductCost(productId)</td><td>实时计算产品成本</td></tr>
                    <tr><td>CE_SafeNum(val)</td><td>安全数值转换</td></tr>
                </table>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">数据流</h3>
                <ul>
                    <li><strong>SupplierPrices</strong> → 最新采购价（按ItemCode+MAX(CreatedAt)）</li>
                    <li><strong>RawMaterialInventory</strong> → 原料单价</li>
                    <li><strong>AccordRecipes</strong> → Accord配方组成（原料+比例）</li>
                    <li><strong>NoteIngredients</strong> → 香调成分聚合（基香+比例）</li>
                    <li><strong>ProductNoteRatios</strong> → 产品香调配比</li>
                    <li><strong>BottleAdditions</strong> → 瓶身成本加成</li>
                    <li><strong>ProductExtraCosts</strong> → 包装/人工分摊</li>
                </ul>
            </div>
        </section>

        <!-- ====== 第十四部分：安全机制 ====== -->
        <section class="section" id="section14">
            <h2 class="section-title"><span class="icon">14</span>安全机制</h2>

            <div class="subsection">
                <h3 class="subsection-title">V8安全加固总览</h3>
                <p>V8版本在V7基础上进行了多次安全迭代（V8升级迭代4+5），形成了完整的安全防护体系：</p>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">1. CSRF防护</h3>
                <p>所有表单和AJAX请求携带CSRF令牌（存于Session中）：</p>
                <div class="code-block">
' 表单中包含CSRF令牌<br>
&lt;input type="hidden" name="csrf_token" value="&lt;%=Session("CSRFToken")%&gt;"&gt;<br><br>
' 服务器端验证<br>
If Request.Form("csrf_token") &lt;&gt; Session("CSRFToken") Then<br>
&nbsp;&nbsp;&nbsp;&nbsp;Response.Status = "403 Forbidden"<br>
&nbsp;&nbsp;&nbsp;&nbsp;Response.End<br>
End If
                </div>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">2. Cookie加密</h3>
                <p>管理员"记住我"Cookie使用安全令牌机制（ValidateSecureToken函数）：</p>
                <ul>
                    <li>Cookie值经过HMAC签名，防止篡改</li>
                    <li>有效期30天，安全性不减</li>
                    <li>自动从Cookie恢复Session会话</li>
                </ul>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">3. SQL注入防护</h3>
                <p>使用 SafeSQL() 函数过滤所有用户输入（定义在 connection.asp 中）：</p>
                <div class="code-block">
' 安全用法示例<br>
sql = "SELECT * FROM Users WHERE Username='" &amp; SafeSQL(username) &amp; "'"<br>
conn.Execute sql<br><br>
' 数值类型使用 CLng/CInt/IsNumeric 前置校验<br>
If IsNumeric(id) Then conn.Execute "UPDATE Products SET Price=" &amp; CDbl(price) &amp; " WHERE ProductID=" &amp; CInt(id)
                </div>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">4. XSS防护</h3>
                <p>所有输出使用 Server.HTMLEncode() 转义：</p>
                <div class="code-block">
Response.Write Server.HTMLEncode(userInput)<br>
Response.Write Server.HTMLEncode(rsProduct("ProductName"))
                </div>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">5. 安全响应头</h3>
                <p>系统在 config.asp 中配置了以下安全响应头：</p>
                <ul>
                    <li><strong>X-Content-Type-Options: nosniff</strong> - 防止MIME类型嗅探</li>
                    <li><strong>X-Frame-Options: SAMEORIGIN</strong> - 防止点击劫持</li>
                    <li><strong>X-XSS-Protection: 1; mode=block</strong> - 启用浏览器XSS过滤器</li>
                    <li><strong>Referrer-Policy: strict-origin-when-cross-origin</strong> - 控制Referer传递</li>
                </ul>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">6. 登录安全</h3>
                <ul>
                    <li><strong>登录锁定：</strong>连续5次登录失败自动锁定账户15分钟</li>
                    <li><strong>登录监控：</strong>login_monitor.asp 记录所有登录尝试（成功/失败）</li>
                    <li><strong>IP黑名单：</strong>ip_blacklist.asp 管理可疑IP的封禁和自动解封</li>
                    <li><strong>速率限制：</strong>防暴力破解</li>
                </ul>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">7. 操作审计</h3>
                <p>所有后台操作通过 LogAdminAction() 记录到 AdminLogs 表：</p>
                <ul>
                    <li>操作人用户名和ID</li>
                    <li>操作时间（精确到秒）</li>
                    <li>操作类型和内容描述</li>
                    <li>客户端IP地址和User-Agent</li>
                    <li>越权访问尝试单独记录</li>
                </ul>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">8. 数据备份策略</h3>
                <ul>
                    <li><strong>备份中心：</strong>backup_center.asp 支持手动/自动备份数据库</li>
                    <li><strong>备份脚本：</strong>database/backup_database.ps1 / auto_fix_backup.bat</li>
                    <li><strong>SQL Server备份：</strong>usp_BackupDatabase 存储过程</li>
                    <li><strong>定期备份建议：</strong>每日自动备份，保留最近30天</li>
                </ul>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">9. 安全配置管理</h3>
                <p>系统设置（settings.asp）提供安全策略集中配置：</p>
                <ul>
                    <li>密码最小长度（默认8位）</li>
                    <li>Session超时时间（默认30分钟）</li>
                    <li>登录最大重试次数（默认5次）</li>
                    <li>MFA多因素认证（默认关闭）</li>
                </ul>
            </div>

            <div class="subsection">
                <h3 class="subsection-title">10. 安全审计</h3>
                <ul>
                    <li><strong>审计页面：</strong>security_audit.asp 提供安全事件汇总和合规检查</li>
                    <li><strong>日志查询：</strong>logs.asp 支持多条件筛选和日志导出</li>
                    <li><strong>安全事件记录：</strong>敏感操作（角色变更/权限修改/密码重置）单独标记</li>
                </ul>
            </div>
        </section>
    </div>

    <div class="footer">
        <p><strong>香氛定制电商系统 V8</strong></p>
        <p>版本号：V8.0.10 | 最后一次更新：2026年5月17日</p>
        <p>技术栈：ASP Classic + SQL Server Express + jQuery + Chart.js</p>
        <p>数据库：PerfumeShop @ localhost\SQLEXPRESS (SQL Server)</p>
        <p style="margin-top: 15px; font-size: 0.85em; opacity: 0.8;">
            本说明书仅供内部使用，请勿外传 | 基于V8.0.10稳定版编制
        </p>
    </div>
</div>
</body>
</html>
