<%@ Language="VBScript" CodePage=65001 %>
<% Response.CodePage = 65001 : Response.CharSet = "UTF-8" %>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>香氛定制电商系统 V7 使用说明手册</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif;
            line-height: 1.8;
            color: #333;
            background-color: #f5f5f5;
        }
        
        /* 顶部导航栏 */
        .navbar {
            background-color: #2c3e50;
            padding: 15px 0;
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            z-index: 1000;
            box-shadow: 0 2px 10px rgba(0,0,0,0.3);
        }
        
        .nav-container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 0 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .nav-title {
            color: #fff;
            font-size: 18px;
            font-weight: bold;
        }
        
        .nav-links {
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
        }
        
        .nav-links a {
            color: #ecf0f1;
            text-decoration: none;
            padding: 8px 15px;
            border-radius: 4px;
            font-size: 14px;
            transition: all 0.3s;
        }
        
        .nav-links a:hover {
            background-color: #8B4513;
            color: #fff;
        }
        
        /* 主内容区 */
        .main-container {
            max-width: 1200px;
            margin: 80px auto 40px;
            padding: 20px;
        }
        
        /* 页面标题 */
        .page-header {
            text-align: center;
            padding: 40px 20px;
            background: linear-gradient(135deg, #8B4513 0%, #A0522D 100%);
            color: #fff;
            border-radius: 10px;
            margin-bottom: 30px;
        }
        
        .page-header h1 {
            font-size: 32px;
            margin-bottom: 10px;
        }
        
        .page-header .subtitle {
            font-size: 16px;
            opacity: 0.9;
        }
        
        /* 章节样式 */
        .chapter {
            background: #fff;
            padding: 30px;
            margin-bottom: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        
        .chapter h2 {
            color: #8B4513;
            font-size: 24px;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 3px solid #8B4513;
        }
        
        .chapter h3 {
            color: #A0522D;
            font-size: 18px;
            margin: 25px 0 15px;
        }
        
        .chapter h4 {
            color: #333;
            font-size: 16px;
            margin: 20px 0 10px;
            font-weight: bold;
        }
        
        .chapter p {
            margin-bottom: 15px;
            text-align: justify;
        }
        
        .chapter ul, .chapter ol {
            margin: 15px 0 15px 30px;
        }
        
        .chapter li {
            margin-bottom: 8px;
        }
        
        /* 分隔线 */
        .chapter-divider {
            border: none;
            height: 2px;
            background: linear-gradient(to right, transparent, #8B4513, transparent);
            margin: 40px 0;
        }
        
        /* 表格样式 */
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            font-size: 14px;
        }
        
        th {
            background-color: #8B4513;
            color: #fff;
            padding: 12px;
            text-align: left;
            font-weight: bold;
        }
        
        td {
            padding: 12px;
            border-bottom: 1px solid #ddd;
        }
        
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        
        tr:hover {
            background-color: #f0e6dc;
        }
        
        /* 代码/代码块 */
        code {
            background-color: #f4f4f4;
            padding: 2px 6px;
            border-radius: 3px;
            font-family: Consolas, Monaco, monospace;
            font-size: 13px;
            color: #c7254e;
        }
        
        .code-block {
            background-color: #2c3e50;
            color: #ecf0f1;
            padding: 15px;
            border-radius: 5px;
            overflow-x: auto;
            font-family: Consolas, Monaco, monospace;
            font-size: 13px;
            margin: 15px 0;
        }
        
        /* 提示框 */
        .tip-box {
            background-color: #e8f4f8;
            border-left: 4px solid #3498db;
            padding: 15px;
            margin: 15px 0;
            border-radius: 0 5px 5px 0;
        }
        
        .tip-box.warning {
            background-color: #fff3cd;
            border-left-color: #ffc107;
        }
        
        .tip-box.danger {
            background-color: #f8d7da;
            border-left-color: #dc3545;
        }
        
        .tip-box.success {
            background-color: #d4edda;
            border-left-color: #28a745;
        }
        
        /* FAQ样式 */
        .faq-item {
            margin-bottom: 20px;
            padding: 15px;
            background-color: #f8f9fa;
            border-radius: 5px;
        }
        
        .faq-question {
            font-weight: bold;
            color: #8B4513;
            margin-bottom: 8px;
        }
        
        .faq-answer {
            color: #555;
        }
        
        /* 流程图样式 */
        .flow-step {
            display: inline-block;
            background-color: #8B4513;
            color: #fff;
            padding: 8px 15px;
            border-radius: 20px;
            margin: 5px;
            font-size: 14px;
        }
        
        .flow-arrow {
            color: #8B4513;
            font-size: 20px;
            margin: 0 5px;
        }
        
        /* 底部 */
        .footer {
            text-align: center;
            padding: 30px;
            background-color: #2c3e50;
            color: #ecf0f1;
            border-radius: 10px;
            margin-top: 30px;
        }
        
        .footer p {
            margin: 5px 0;
        }
        
        /* 响应式 */
        @media (max-width: 768px) {
            .nav-links {
                display: none;
            }
            
            .page-header h1 {
                font-size: 24px;
            }
            
            .chapter {
                padding: 20px;
            }
            
            .chapter h2 {
                font-size: 20px;
            }
            
            table {
                font-size: 12px;
            }
            
            th, td {
                padding: 8px;
            }
        }
        
        /* 锚点偏移 */
        .anchor {
            display: block;
            position: relative;
            top: -70px;
            visibility: hidden;
        }
    </style>
</head>
<body>
    <!-- 顶部导航栏 -->
    <nav class="navbar">
        <div class="nav-container">
            <div class="nav-title">V7 使用手册</div>
            <div class="nav-links">
                <a href="#chapter1">第一章</a>
                <a href="#chapter2">第二章</a>
                <a href="#chapter3">第三章</a>
                <a href="#chapter4">第四章</a>
                <a href="#chapter5">第五章</a>
                <a href="#chapter6">第六章</a>
                <a href="#chapter7">第七章</a>
            </div>
        </div>
    </nav>

    <div class="main-container">
        <!-- 页面标题 -->
        <div class="page-header">
            <h1>香氛定制电商系统 V7 使用说明手册</h1>
            <div class="subtitle">Perfume Customization E-commerce System V7 User Manual</div>
        </div>

        <!-- 第一章：网站整体功能架构和模块介绍 -->
        <span class="anchor" id="chapter1"></span>
        <div class="chapter">
            <h2>第一章：网站整体功能架构和模块介绍</h2>
            
            <h3>1.1 技术栈</h3>
            <p>本系统采用经典的技术架构组合，确保稳定性和可维护性：</p>
            <ul>
                <li><strong>后端技术：</strong>ASP Classic (VBScript)</li>
                <li><strong>数据库：</strong>Microsoft Access MDB</li>
                <li><strong>前端框架：</strong>jQuery 3.6.0</li>
                <li><strong>页面技术：</strong>HTML5 / CSS3</li>
            </ul>
            
            <h3>1.2 系统架构说明</h3>
            <p>系统采用前后台分离的设计模式，包含以下主要组成部分：</p>
            
            <h4>前台商城</h4>
            <ul>
                <li><strong>商品浏览：</strong>分类展示、详情查看、筛选排序</li>
                <li><strong>购物车：</strong>添加商品、数量修改、删除、全选结算</li>
                <li><strong>结算系统：</strong>地址选择、支付方式、订单确认</li>
                <li><strong>用户中心：</strong>订单管理、个人信息、收货地址、收藏夹</li>
            </ul>
            
            <h4>四大后台管理系统</h4>
            <ul>
                <li><strong>运营管理后台：</strong>订单处理、客户管理、商品管理、营销活动</li>
                <li><strong>生产管理后台：</strong>库存管理、生产订单跟踪、配方管理</li>
                <li><strong>财务管理后台：</strong>收入统计、成本分析、支付审核、报表中心</li>
                <li><strong>系统管理后台：</strong>角色权限、管理员管理、操作日志、站点配置</li>
            </ul>
            
            <h3>1.3 数据库核心表结构</h3>
            <table>
                <tr>
                    <th>表名</th>
                    <th>说明</th>
                </tr>
                <tr>
                    <td>Users</td>
                    <td>用户账户信息表</td>
                </tr>
                <tr>
                    <td>Products</td>
                    <td>商品信息主表</td>
                </tr>
                <tr>
                    <td>FragranceNotes</td>
                    <td>香调配置表</td>
                </tr>
                <tr>
                    <td>BaseNotes</td>
                    <td>基香成分表</td>
                </tr>
                <tr>
                    <td>Cart</td>
                    <td>购物车数据表</td>
                </tr>
                <tr>
                    <td>Orders / OrderDetails</td>
                    <td>订单主表 / 订单明细表</td>
                </tr>
                <tr>
                    <td>Categories</td>
                    <td>商品分类表</td>
                </tr>
                <tr>
                    <td>Volumes / BottleStyles</td>
                    <td>容量规格表 / 瓶型表</td>
                </tr>
                <tr>
                    <td>ProductVolumePrices</td>
                    <td>商品容量价格表</td>
                </tr>
                <tr>
                    <td>ProductNoteRatios</td>
                    <td>商品香调配比表</td>
                </tr>
                <tr>
                    <td>UserAddresses / UserFavorites</td>
                    <td>用户地址表 / 收藏表</td>
                </tr>
                <tr>
                    <td>SiteSettings</td>
                    <td>站点配置表</td>
                </tr>
                <tr>
                    <td>AdminUsers / AdminRoles / AdminLogs</td>
                    <td>管理员表 / 角色表 / 操作日志表</td>
                </tr>
                <tr>
                    <td>NoteInventory / InventoryTransactions</td>
                    <td>库存表 / 库存变动记录表</td>
                </tr>
                <tr>
                    <td>ProductionOrders</td>
                    <td>生产订单表</td>
                </tr>
                <tr>
                    <td>MarketingCampaigns / Coupons</td>
                    <td>营销活动表 / 优惠券表</td>
                </tr>
                <tr>
                    <td>UserPoints</td>
                    <td>用户积分表</td>
                </tr>
                <tr>
                    <td>NoteIngredients</td>
                    <td>香调成分关联表</td>
                </tr>
            </table>
        </div>

        <hr class="chapter-divider">

        <!-- 第二章：前台用户功能使用指南 -->
        <span class="anchor" id="chapter2"></span>
        <div class="chapter">
            <h2>第二章：前台用户功能使用指南</h2>
            
            <h3>2.1 浏览商品</h3>
            
            <h4>首页三大栏目</h4>
            <ul>
                <li><strong>品牌定香：</strong>展示品牌经典香水系列</li>
                <li><strong>定制香水：</strong>用户可自定义香调配比的个性化产品</li>
                <li><strong>KOL推荐：</strong>网红/达人推荐的预设配比商品</li>
            </ul>
            <div class="tip-box">
                <strong>提示：</strong>三大栏目的显示/隐藏可在后台站点设置中控制。
            </div>
            
            <h4>商品列表功能</h4>
            <ul>
                <li><strong>分类筛选：</strong>按商品分类快速过滤</li>
                <li><strong>关键词搜索：</strong>支持商品名称模糊搜索</li>
                <li><strong>价格区间：</strong>自定义价格范围筛选</li>
                <li><strong>排序方式：</strong>价格、销量、上架时间等多维度排序</li>
                <li><strong>分页浏览：</strong>支持页码跳转和每页数量设置</li>
            </ul>
            
            <h4>商品详情三种模式</h4>
            <table>
                <tr>
                    <th>商品类型</th>
                    <th>详情页展示</th>
                </tr>
                <tr>
                    <td>品牌定香</td>
                    <td>选择规格（容量/瓶型）后直接购买</td>
                </tr>
                <tr>
                    <td>定制香水</td>
                    <td>选择前中后调 → 设置配比 → 选容量瓶型</td>
                </tr>
                <tr>
                    <td>KOL推荐</td>
                    <td>查看预设配比详情 → 可直接购买或微调</td>
                </tr>
            </table>
            
            <h3>2.2 注册登录</h3>
            
            <h4>用户注册</h4>
            <ul>
                <li><strong>用户名：</strong>3-20个字符，支持字母数字下划线</li>
                <li><strong>邮箱：</strong>有效的电子邮箱地址</li>
                <li><strong>密码：</strong>至少6个字符，建议包含字母和数字</li>
                <li><strong>姓名：</strong>真实姓名（用于收货）</li>
                <li><strong>电话：</strong>有效的手机号码</li>
            </ul>
            
            <h4>用户登录</h4>
            <ul>
                <li>支持用户名或邮箱登录</li>
                <li>密码输入错误5次将锁定账户15分钟</li>
                <li>提供"记住我"功能，Cookie有效期30天</li>
            </ul>
            
            <div class="tip-box warning">
                <strong>注意：</strong>连续5次登录失败将锁定账户15分钟，请妥善保管密码。
            </div>
            
            <h3>2.3 定制香水</h3>
            <p>定制香水是本系统的核心功能，流程如下：</p>
            <ol>
                <li><strong>选择香调：</strong>从前调、中调、后调中各选一种</li>
                <li><strong>设置配比：</strong>调整各香调比例（有最小值限制）</li>
                <li><strong>选择容量：</strong>30ml / 50ml / 100ml 等规格</li>
                <li><strong>选择瓶型：</strong>经典款 / 奢华款等</li>
                <li><strong>个性标签：</strong>添加自定义标签文字（可选）</li>
            </ol>
            
            <h3>2.4 购物车</h3>
            <ul>
                <li><strong>加入购物车：</strong>商品详情页点击加入</li>
                <li><strong>修改数量：</strong>在购物车页面直接修改</li>
                <li><strong>删除商品：</strong>单条删除或批量删除</li>
                <li><strong>全选结算：</strong>一键选择所有商品</li>
            </ul>
            <div class="tip-box">
                <strong>数据同步：</strong>未登录时购物车数据存储在Session中，登录后自动与账户购物车合并。
            </div>
            
            <h3>2.5 下单支付</h3>
            <ol>
                <li><strong>选择商品结算：</strong>从购物车选择要购买的商品</li>
                <li><strong>选择收货地址：</strong>使用已有地址或新增地址</li>
                <li><strong>选择支付方式：</strong>支付宝 / 微信支付 / 银行卡等</li>
                <li><strong>确认订单：</strong>查看订单详情，含成分汇总去重显示</li>
                <li><strong>完成支付：</strong>跳转支付页面完成付款</li>
            </ol>
            
            <h3>2.6 个人中心</h3>
            
            <h4>我的订单</h4>
            <ul>
                <li>按订单状态筛选（待付款/待发货/待收货/已完成）</li>
                <li>按日期范围筛选</li>
                <li>订单号/商品名称搜索</li>
                <li>查看订单详情和物流信息</li>
            </ul>
            
            <h4>其他功能</h4>
            <ul>
                <li><strong>个人信息：</strong>修改基本资料</li>
                <li><strong>收货地址：</strong>管理常用收货地址</li>
                <li><strong>我的收藏：</strong>查看收藏的商品</li>
                <li><strong>修改密码：</strong>定期更换账户密码</li>
            </ul>
        </div>

        <hr class="chapter-divider">

        <!-- 第三章：后台管理功能使用说明 -->
        <span class="anchor" id="chapter3"></span>
        <div class="chapter">
            <h2>第三章：后台管理功能使用说明</h2>
            
            <h3>3.1 登录入口</h3>
            <ul>
                <li><strong>登录地址：</strong><code>/admin/login.asp</code></li>
                <li><strong>统一入口：</strong><code>/admin/portal.asp</code> - 根据角色显示可访问模块</li>
            </ul>
            
            <h3>3.2 运营管理后台（/admin/operation/）</h3>
            
            <h4>运营概览</h4>
            <ul>
                <li>今日订单数量</li>
                <li>今日营收金额</li>
                <li>新增客户数</li>
                <li>在售商品数量</li>
            </ul>
            
            <h4>订单管理</h4>
            <ul>
                <li><strong>订单列表：</strong>查看所有订单，支持多条件筛选</li>
                <li><strong>状态更新：</strong>Pending → Paid → Processing → Shipped → Completed</li>
                <li><strong>订单详情：</strong>查看完整订单信息和商品明细</li>
            </ul>
            
            <h4>客户管理</h4>
            <ul>
                <li>客户列表查看</li>
                <li>账户状态管理（正常/禁用/VIP）</li>
                <li>客户详情编辑</li>
            </ul>
            
            <h4>商品管理</h4>
            <p>支持三类商品的CRUD操作：</p>
            <ul>
                <li><strong>品牌定香：</strong>配置基香和规格</li>
                <li><strong>定制产品：</strong>设置可调配的香调</li>
                <li><strong>KOL预设配比：</strong>配置推荐配方</li>
            </ul>
            
            <h4>香调配置</h4>
            <ul>
                <li>增删改查香调信息</li>
                <li>设置香调类型（前调/中调/后调/基香）</li>
                <li>配置附加价格</li>
                <li>设置建议配比</li>
                <li>与基香关联配置</li>
            </ul>
            
            <h4>基香管理</h4>
            <ul>
                <li>成分库管理</li>
                <li>启用/禁用状态控制</li>
                <li>与香调关联设置</li>
            </ul>
            
            <h4>KOL审核</h4>
            <ul>
                <li>审核KOL提交的商品配比</li>
                <li>状态流转：Pending → Approved / Rejected</li>
            </ul>
            
            <h4>其他功能</h4>
            <ul>
                <li>配方推荐管理</li>
                <li>营销活动配置</li>
                <li>积分管理</li>
                <li>支付开关控制</li>
            </ul>
            
            <h3>3.3 生产管理后台（/admin/production/）</h3>
            
            <h4>生产概览</h4>
            <ul>
                <li>待生产订单统计</li>
                <li>生产中订单统计</li>
                <li>已完成订单统计</li>
            </ul>
            
            <h4>库存管理</h4>
            <ul>
                <li>查看库存余量</li>
                <li>更新库存数量</li>
                <li>设置预警阈值</li>
                <li>入库操作</li>
                <li>查看变动历史</li>
            </ul>
            
            <h4>库存预警</h4>
            <ul>
                <li>低库存商品列表</li>
                <li>快速入库入口</li>
            </ul>
            
            <h4>生产订单</h4>
            <ul>
                <li>订单生产状态跟踪</li>
                <li>状态流转：待生产 → 生产中 → 已完成</li>
                <li>负责人分配</li>
            </ul>
            
            <h4>订单生产</h4>
            <ul>
                <li>待生产订单清单</li>
                <li>成分检查</li>
                <li>生产状态更新</li>
            </ul>
            
            <h3>3.4 财务管理后台（/admin/finance/）</h3>
            
            <h4>财务概览</h4>
            <ul>
                <li>累计营收金额</li>
                <li>本月营收金额</li>
                <li>待付款订单数</li>
            </ul>
            
            <h4>收入统计</h4>
            <ul>
                <li>日期范围筛选</li>
                <li>按支付方式统计</li>
                <li>日收入趋势图</li>
            </ul>
            
            <h4>财务报表</h4>
            <ul>
                <li>月度收入报表</li>
                <li>金额分布分析</li>
                <li>支付占比统计</li>
            </ul>
            
            <h4>支付配置</h4>
            <ul>
                <li>支付参数设置</li>
                <li>费率配置</li>
                <li>支付渠道管理</li>
            </ul>
            
            <h3>3.5 系统管理后台（/admin/system/）</h3>
            <div class="tip-box warning">
                <strong>注意：</strong>仅SUPER_ADMIN角色可访问系统管理后台。
            </div>
            
            <h4>角色管理</h4>
            <ul>
                <li>8种预设角色管理</li>
                <li>角色权限配置</li>
            </ul>
            
            <h4>管理员管理</h4>
            <ul>
                <li>管理员CRUD操作</li>
                <li>角色分配</li>
                <li>账户状态管理</li>
                <li>密码重置</li>
            </ul>
            
            <h4>操作日志</h4>
            <ul>
                <li>日志查询</li>
                <li>多条件筛选</li>
                <li>日志导出</li>
            </ul>
            
            <h4>站点设置</h4>
            <ul>
                <li>首页栏目显示/隐藏</li>
                <li>香调配比最小值设置</li>
            </ul>
            
            <h4>系统配置</h4>
            <ul>
                <li>数据库备份</li>
                <li>日志清理</li>
                <li>缓存管理</li>
            </ul>
            
            <h4>数据统计</h4>
            <ul>
                <li>销售趋势分析</li>
                <li>热销TOP5排行</li>
                <li>用户活跃度统计</li>
            </ul>
        </div>

        <hr class="chapter-divider">

        <!-- 第四章：权限分配和角色说明 -->
        <span class="anchor" id="chapter4"></span>
        <div class="chapter">
            <h2>第四章：权限分配和角色说明</h2>
            
            <h3>4.1 角色权限表</h3>
            <table>
                <tr>
                    <th>角色代码</th>
                    <th>角色名称</th>
                    <th>可访问后台</th>
                    <th>主要权限</th>
                </tr>
                <tr>
                    <td>SUPER_ADMIN</td>
                    <td>超级管理员</td>
                    <td>全部四个后台</td>
                    <td>完全控制，包括系统管理</td>
                </tr>
                <tr>
                    <td>OP_MANAGER</td>
                    <td>运营经理</td>
                    <td>运营管理</td>
                    <td>全部运营功能，包括审核</td>
                </tr>
                <tr>
                    <td>OP_STAFF</td>
                    <td>运营专员</td>
                    <td>运营管理</td>
                    <td>基础运营操作，无审核权限</td>
                </tr>
                <tr>
                    <td>PROD_MANAGER</td>
                    <td>生产经理</td>
                    <td>生产管理</td>
                    <td>全部生产功能，包括人员分配</td>
                </tr>
                <tr>
                    <td>PROD_STAFF</td>
                    <td>生产专员</td>
                    <td>生产管理</td>
                    <td>基础生产操作，状态更新</td>
                </tr>
                <tr>
                    <td>FIN_MANAGER</td>
                    <td>财务经理</td>
                    <td>财务管理</td>
                    <td>全部财务功能，报表导出</td>
                </tr>
                <tr>
                    <td>FIN_STAFF</td>
                    <td>财务专员</td>
                    <td>财务管理</td>
                    <td>基础财务查看，数据录入</td>
                </tr>
                <tr>
                    <td>CONTENT_ADMIN</td>
                    <td>内容管理员</td>
                    <td>运营管理(部分)</td>
                    <td>商品/分类/香调内容管理</td>
                </tr>
            </table>
            
            <h3>4.2 权限验证机制</h3>
            <p>系统采用多层权限验证机制，确保访问安全：</p>
            <ol>
                <li><strong>Session存储：</strong>登录时获取用户角色信息存入Session</li>
                <li><strong>模块验证：</strong>每个后台模块包含auth.asp进行权限检查</li>
                <li><strong>功能验证：</strong>调用<code>VerifyModuleAccess()</code>函数验证具体权限</li>
                <li><strong>越权处理：</strong>权限不足时跳转至unauthorized.asp并记录日志</li>
            </ol>
            
            <div class="code-block">
权限验证流程：<br>
用户登录 → 获取角色 → Session存储 → 访问模块 → auth.asp检查 → VerifyModuleAccess() → 通过/拒绝
            </div>
        </div>

        <hr class="chapter-divider">

        <!-- 第五章：关键业务流程 -->
        <span class="anchor" id="chapter5"></span>
        <div class="chapter">
            <h2>第五章：关键业务流程</h2>
            
            <h3>5.1 商品上架流程</h3>
            <div style="text-align: center; margin: 20px 0;">
                <span class="flow-step">配置基香</span>
                <span class="flow-arrow">→</span>
                <span class="flow-step">配置香调</span>
                <span class="flow-arrow">→</span>
                <span class="flow-step">创建商品</span>
                <span class="flow-arrow">→</span>
                <span class="flow-step">上传图片</span>
                <span class="flow-arrow">→</span>
                <span class="flow-step">上架</span>
            </div>
            <ol>
                <li><strong>配置基香：</strong>在基香管理中创建基础香调成分</li>
                <li><strong>配置香调：</strong>设置前中后调及各调可选香调</li>
                <li><strong>创建商品：</strong>选择商品类型，填写基本信息，关联基香，设置价格</li>
                <li><strong>上传图片：</strong>上传商品主图和详情图</li>
                <li><strong>上架：</strong>设置IsActive为True，商品前台可见</li>
            </ol>
            
            <h3>5.2 KOL商品发布流程</h3>
            <div style="text-align: center; margin: 20px 0;">
                <span class="flow-step">创建KOL商品</span>
                <span class="flow-arrow">→</span>
                <span class="flow-step">配置预设配比</span>
                <span class="flow-arrow">→</span>
                <span class="flow-step">提交审核</span>
                <span class="flow-arrow">→</span>
                <span class="flow-step">审核通过</span>
                <span class="flow-arrow">→</span>
                <span class="flow-step">自动展示</span>
            </div>
            
            <h3>5.3 订单处理流程</h3>
            <div style="text-align: center; margin: 20px 0;">
                <span class="flow-step">用户下单</span>
                <span class="flow-arrow">→</span>
                <span class="flow-step">支付</span>
                <span class="flow-arrow">→</span>
                <span class="flow-step">运营处理</span>
                <span class="flow-arrow">→</span>
                <span class="flow-step">生产</span>
                <span class="flow-arrow">→</span>
                <span class="flow-step">发货</span>
                <span class="flow-arrow">→</span>
                <span class="flow-step">完成</span>
            </div>
            <table>
                <tr>
                    <th>状态</th>
                    <th>说明</th>
                    <th>操作人</th>
                </tr>
                <tr>
                    <td>Pending</td>
                    <td>待付款</td>
                    <td>用户</td>
                </tr>
                <tr>
                    <td>Paid</td>
                    <td>已付款</td>
                    <td>系统自动</td>
                </tr>
                <tr>
                    <td>Processing</td>
                    <td>处理中</td>
                    <td>运营人员</td>
                </tr>
                <tr>
                    <td>生产中</td>
                    <td>生产部门处理</td>
                    <td>生产人员</td>
                </tr>
                <tr>
                    <td>Shipped</td>
                    <td>已发货</td>
                    <td>运营人员</td>
                </tr>
                <tr>
                    <td>Completed</td>
                    <td>已完成</td>
                    <td>用户确认/系统自动</td>
                </tr>
            </table>
            
            <h3>5.4 库存管理流程</h3>
            <ol>
                <li><strong>检查预警：</strong>定期查看库存预警列表</li>
                <li><strong>入库操作：</strong>对低库存商品进行入库</li>
                <li><strong>记录变动：</strong>系统自动记录库存变动历史</li>
                <li><strong>调整阈值：</strong>根据销售情况调整预警阈值</li>
            </ol>
            
            <h3>5.5 财务报表流程</h3>
            <ol>
                <li><strong>选择日期范围：</strong>设置要统计的起止日期</li>
                <li><strong>查看收入趋势：</strong>分析日/周/月收入变化</li>
                <li><strong>分析支付分布：</strong>各支付渠道占比</li>
                <li><strong>查看月度报表：</strong>生成并导出月度财务报表</li>
            </ol>
        </div>

        <hr class="chapter-divider">

        <!-- 第六章：常见问题解答 -->
        <span class="anchor" id="chapter6"></span>
        <div class="chapter">
            <h2>第六章：常见问题解答（FAQ）</h2>
            
            <div class="faq-item">
                <div class="faq-question">Q1: 登录提示"账户已锁定"怎么办？</div>
                <div class="faq-answer">
                    <strong>原因：</strong>连续5次输入错误密码。<br>
                    <strong>解决：</strong>账户将被锁定15分钟，请等待后重试，或联系超级管理员解锁。
                </div>
            </div>
            
            <div class="faq-item">
                <div class="faq-question">Q2: 购物车添加失败怎么办？</div>
                <div class="faq-answer">
                    <strong>原因：</strong>CSRF令牌过期。<br>
                    <strong>解决：</strong>刷新页面后重新添加商品。
                </div>
            </div>
            
            <div class="faq-item">
                <div class="faq-question">Q3: 后台无法访问某模块怎么办？</div>
                <div class="faq-answer">
                    <strong>原因：</strong>当前角色权限不足。<br>
                    <strong>解决：</strong>联系超级管理员检查角色权限配置。
                </div>
            </div>
            
            <div class="faq-item">
                <div class="faq-question">Q4: 商品详情不显示定制选项怎么办？</div>
                <div class="faq-answer">
                    <strong>原因1：</strong>商品类型不是"Custom"。<br>
                    <strong>原因2：</strong>未配置香调。<br>
                    <strong>解决：</strong>检查商品类型设置，并在香调配置中关联该商品。
                </div>
            </div>
            
            <div class="faq-item">
                <div class="faq-question">Q5: 首页某栏目不显示怎么办？</div>
                <div class="faq-answer">
                    <strong>原因：</strong>栏目在站点设置中被隐藏。<br>
                    <strong>解决：</strong>进入系统管理 → 站点设置，检查ShowFixedSection/ShowCustomSection/ShowKOLSection设置。
                </div>
            </div>
            
            <div class="faq-item">
                <div class="faq-question">Q6: 订单成分信息缺失怎么办？</div>
                <div class="faq-answer">
                    <strong>原因：</strong>商品未关联基香或基香未配置成分。<br>
                    <strong>解决：</strong>检查商品基香关联，确保基香成分配置完整。
                </div>
            </div>
            
            <div class="faq-item">
                <div class="faq-question">Q7: 库存预警不准确怎么办？</div>
                <div class="faq-answer">
                    <strong>原因：</strong>预警阈值设置不合理。<br>
                    <strong>解决：</strong>进入生产管理 → 库存预警，调整各商品的预警阈值。
                </div>
            </div>
            
            <div class="faq-item">
                <div class="faq-question">Q8: KOL商品审核后未显示怎么办？</div>
                <div class="faq-answer">
                    <strong>原因：</strong>审核状态或上架状态不正确。<br>
                    <strong>解决：</strong>确认商品审核状态为"Approved"且IsActive为True。
                </div>
            </div>
            
            <div class="faq-item">
                <div class="faq-question">Q9: 数据库连接失败怎么办？</div>
                <div class="faq-answer">
                    <strong>原因：</strong>数据库文件不存在或权限不足。<br>
                    <strong>解决：</strong>检查PerfumeShop.mdb文件是否存在，确保IIS进程对database目录有读写权限。
                </div>
            </div>
            
            <div class="faq-item">
                <div class="faq-question">Q10: 页面显示乱码怎么办？</div>
                <div class="faq-answer">
                    <strong>原因：</strong>编码设置不正确。<br>
                    <strong>解决：</strong>确保IIS配置为UTF-8编码，页面包含正确的charset声明。
                </div>
            </div>
        </div>

        <hr class="chapter-divider">

        <!-- 第七章：安全注意事项和最佳实践 -->
        <span class="anchor" id="chapter7"></span>
        <div class="chapter">
            <h2>第七章：安全注意事项和最佳实践</h2>
            
            <h3>7.1 安全机制</h3>
            <table>
                <tr>
                    <th>安全类型</th>
                    <th>实现方式</th>
                </tr>
                <tr>
                    <td>CSRF防护</td>
                    <td>Session令牌 + jQuery拦截器自动附加</td>
                </tr>
                <tr>
                    <td>Cookie加密</td>
                    <td>HMAC签名，有效期30天</td>
                </tr>
                <tr>
                    <td>SQL注入防护</td>
                    <td>SafeSQL函数过滤危险字符</td>
                </tr>
                <tr>
                    <td>XSS防护</td>
                    <td>HTMLEncode输出 + 安全响应头</td>
                </tr>
                <tr>
                    <td>登录速率限制</td>
                    <td>5次失败锁定15分钟</td>
                </tr>
                <tr>
                    <td>操作审计</td>
                    <td>完整的操作日志记录</td>
                </tr>
            </table>
            
            <h3>7.2 最佳实践（10条）</h3>
            
            <div class="tip-box success">
                <strong>1. 定期更换密码</strong><br>
                建议每90天更换一次管理员密码，避免使用简单密码。
            </div>
            
            <div class="tip-box success">
                <strong>2. 最小权限原则</strong><br>
                为每个管理员分配最小必要的权限，避免过度授权。
            </div>
            
            <div class="tip-box success">
                <strong>3. 定期审查日志</strong><br>
                每周检查操作日志，发现异常行为及时处理。
            </div>
            
            <div class="tip-box success">
                <strong>4. 定期备份数据库</strong><br>
                建议每日自动备份MDB文件，保留最近7天备份。
            </div>
            
            <div class="tip-box warning">
                <strong>5. 删除初始化脚本</strong><br>
                系统部署完成后，立即删除database/create_v7_role_system.asp等初始化脚本。
            </div>
            
            <div class="tip-box warning">
                <strong>6. 删除调试页面</strong><br>
                生产环境删除所有调试页面和测试数据。
            </div>
            
            <div class="tip-box warning">
                <strong>7. 限制目录访问</strong><br>
                配置IIS禁止直接访问database目录，仅允许脚本访问。
            </div>
            
            <div class="tip-box success">
                <strong>8. 启用HTTPS</strong><br>
                生产环境必须启用SSL证书，使用HTTPS协议访问。
            </div>
            
            <div class="tip-box success">
                <strong>9. 及时处理库存预警</strong><br>
                每日检查库存预警，避免因缺货影响订单处理。
            </div>
            
            <div class="tip-box success">
                <strong>10. 及时审核KOL商品</strong><br>
                KOL提交的商品应在24小时内完成审核。
            </div>
            
            <div class="tip-box danger">
                <strong>重要提醒：</strong><br>
                超级管理员账户密码务必妥善保管，建议启用双因素认证（如有条件）。<br>
                如发现账户异常，立即锁定账户并检查操作日志。
            </div>
        </div>

        <!-- 底部 -->
        <div class="footer">
            <p><strong>香氛定制电商系统 V7</strong></p>
            <p>版本号：V7.0.0</p>
            <p>最后更新日期：2026年4月8日</p>
            <p style="margin-top: 15px; font-size: 12px; opacity: 0.8;">
                本手册仅供内部使用，请勿外传
            </p>
        </div>
    </div>
</body>
</html>
