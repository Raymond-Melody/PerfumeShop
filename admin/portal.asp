<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="includes/role_auth.asp"-->
<%
Call OpenConnection()

' 获取当前管理员信息
Dim adminId, adminUsername, adminRoleName, adminRealName, adminDepartment
adminId = Session("AdminID")
adminUsername = Session("AdminUsername")
adminRoleName = Session("AdminRoleName")
adminRealName = Session("AdminRealName")
adminDepartment = Session("AdminDepartment")

' 如果没有真实姓名，使用用户名
If adminRealName = "" Then
    adminRealName = adminUsername
End If

' 获取可访问的后台列表
Dim accessiblePortals
Session("AdminRoleID") = ""  ' 强制重新加载角色信息（修复RoleCode缓存问题）
accessiblePortals = GetAccessiblePortals()

' 如果只有一个可访问后台，直接跳转
If InStr(accessiblePortals, ",") = 0 And accessiblePortals <> "" Then
    Response.Redirect accessiblePortals & "/index.asp"
    Response.End
End If

' 记录登录日志
Call LogAdminAction("访问统一入口", "portal", "", "", "")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>管理中心入口 - 香氛电商系统</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            background: #1a1a2e;
            min-height: 100vh;
            padding: 20px;
        }
        .container { 
            max-width: 1400px; 
            width: 100%; 
            margin: 0 auto;
        }
        .header { 
            text-align: center; 
            margin-bottom: 30px; 
            padding-top: 20px;
        }
        .header h1 { 
            font-size: 32px; 
            color: #fff;
            margin-bottom: 8px;
            font-weight: 700;
        }
        .header h1 span { color: #00bcd4; }
        .header p { 
            font-size: 15px; 
            color: #888;
        }
        .user-info {
            background: linear-gradient(135deg, #2d2d44, #1e1e32);
            border-radius: 16px;
            padding: 20px 30px;
            margin-bottom: 30px;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 25px;
            color: #e0e0e0;
            border: 1px solid rgba(255,255,255,0.06);
        }
        .user-avatar {
            width: 56px;
            height: 56px;
            background: linear-gradient(135deg, #00bcd4, #00838f);
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 26px;
            color: #fff;
        }
        .user-details h3 {
            font-size: 18px;
            margin-bottom: 4px;
            color: #fff;
        }
        .user-details p {
            font-size: 13px;
            color: #888;
        }
        .portal-grid { 
            display: grid; 
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); 
            gap: 20px;
            margin-bottom: 30px;
        }
        .section-title {
            grid-column: 1 / -1;
            font-size: 13px;
            color: #666;
            text-transform: uppercase;
            letter-spacing: 1px;
            padding: 10px 0 0;
            margin-bottom: -10px;
        }
        .portal-card { 
            background: linear-gradient(135deg, #2d2d44, #1e1e32);
            border-radius: 16px; 
            padding: 28px 25px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.2);
            transition: all 0.3s ease;
            cursor: pointer;
            text-decoration: none;
            color: inherit;
            display: block;
            border: 1px solid rgba(255,255,255,0.05);
            position: relative;
            overflow: hidden;
        }
        .portal-card::before {
            content: '';
            position: absolute;
            top: 0; left: 0;
            width: 4px;
            height: 100%;
            transition: all 0.3s ease;
        }
        .portal-card:hover { 
            transform: translateY(-6px); 
            box-shadow: 0 12px 40px rgba(0,0,0,0.4);
            border-color: rgba(0,188,212,0.3);
        }
        .portal-card:hover::before { width: 6px; }
        .portal-card.operation::before { background: #667eea; }
        .portal-card.semifinished::before { background: #FF9800; }
        .portal-card.prodcenter::before { background: #4CAF50; }
        .portal-card.finance::before { background: #00bcd4; }
        .portal-card.system::before { background: #E91E63; }
        .portal-card.techcenter::before { background: #9C27B0; }
        .portal-card.purchase::before { background: #FF5722; }
        .portal-card.logistics::before { background: #2196F3; }
        .portal-card.inventory::before { background: #00BCD4; }
        .portal-card.logout::before { background: #f44336; }
        
        .card-top { display: flex; align-items: center; gap: 16px; margin-bottom: 15px; }
        .portal-icon { 
            width: 64px; 
            height: 64px; 
            border-radius: 14px; 
            display: flex; 
            align-items: center; 
            justify-content: center; 
            font-size: 28px;
            color: white;
            flex-shrink: 0;
        }
        .portal-card.operation .portal-icon { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); }
        .portal-card.semifinished .portal-icon { background: linear-gradient(135deg, #FF9800 0%, #E65100 100%); }
        .portal-card.prodcenter .portal-icon { background: linear-gradient(135deg, #4CAF50 0%, #1B5E20 100%); }
        .portal-card.finance .portal-icon { background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); }
        .portal-card.system .portal-icon { background: linear-gradient(135deg, #E91E63 0%, #880E4F 100%); }
        .portal-card.techcenter .portal-icon { background: linear-gradient(135deg, #9C27B0 0%, #4A148C 100%); }
        .portal-card.purchase .portal-icon { background: linear-gradient(135deg, #FF5722 0%, #BF360C 100%); }
        .portal-card.logistics .portal-icon { background: linear-gradient(135deg, #2196F3 0%, #0D47A1 100%); }
        .portal-card.inventory .portal-icon { background: linear-gradient(135deg, #00BCD4 0%, #006064 100%); }
        .portal-card.logout .portal-icon { background: linear-gradient(135deg, #f44336 0%, #b71c1c 100%); }
        
        .card-title h3 { 
            font-size: 20px; 
            color: #fff;
            margin-bottom: 4px;
        }
        .card-title .badge-new {
            display: inline-block;
            padding: 2px 8px;
            background: rgba(76,175,80,0.3);
            color: #4CAF50;
            border-radius: 10px;
            font-size: 10px;
            font-weight: 700;
        }
        .card-title p { 
            color: #888; 
            font-size: 13px;
        }
        .portal-features {
            padding-top: 15px;
            border-top: 1px solid rgba(255,255,255,0.06);
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 4px;
        }
        .portal-features li {
            list-style: none;
            padding: 3px 0;
            color: #999;
            font-size: 12px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .portal-features li i {
            color: #4CAF50;
            margin-right: 5px;
            font-size: 10px;
        }
        .logout-card {
            background: rgba(244,67,54,0.08);
            border: 1px solid rgba(244,67,54,0.15);
            text-align: center;
        }
        .logout-card .card-top { justify-content: center; }
        .logout-card h3 { color: #EF9A9A; }
        .logout-card p { color: #888; font-size: 13px; }
        
        .footer-info {
            text-align: center;
            color: #555;
            font-size: 12px;
            padding: 20px 0;
        }
        .footer-info span { color: #00bcd4; }
        @media (max-width: 768px) {
            .header h1 { font-size: 24px; }
            .user-info { flex-direction: column; text-align: center; }
            .portal-grid { grid-template-columns: 1fr; }
            .portal-features { grid-template-columns: 1fr; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1><i class="fas fa-cubes"></i> <span>管理中心</span></h1>
            <p>V9.0 全模块统一入口 · 请选择要进入的管理后台</p>
        </div>
        
        <div class="user-info">
            <div class="user-avatar">
                <i class="fas fa-user"></i>
            </div>
            <div class="user-details">
                <h3><%= adminRealName %></h3>
                <p>
                    <i class="fas fa-id-badge"></i> <%= adminRoleName %> 
                    <% If adminDepartment <> "" Then %>
                    | <i class="fas fa-building"></i> <%= adminDepartment %>
                    <% End If %>
                </p>
            </div>
        </div>
        
        <div class="portal-grid">
            <% If InStr(accessiblePortals, "operation") > 0 Then %>
            <a href="operation/index.asp" class="portal-card operation">
                <div class="card-top">
                    <div class="portal-icon"><i class="fas fa-chart-line"></i></div>
                    <div class="card-title">
                        <h3>运营管理中心</h3>
                        <p>订单、客户、商品、评价</p>
                    </div>
                </div>
                <ul class="portal-features">
                    <li><i class="fas fa-check"></i> 订单管理</li>
                    <li><i class="fas fa-check"></i> 客户CRM画像</li>
                    <li><i class="fas fa-check"></i> 售后管理</li>
                    <li><i class="fas fa-check"></i> 评价审核</li>
                </ul>
            </a>
            <% End If %>

            <% If InStr(accessiblePortals, "semifinished") > 0 Then %>
            <a href="semifinished/index.asp" class="portal-card semifinished">
                <div class="card-top">
                    <div class="portal-icon"><i class="fas fa-vial"></i></div>
                    <div class="card-title">
                        <h3>半成品生产中心 <span class="badge-new">NEW</span></h3>
                        <p>原料、基香、香调生产</p>
                    </div>
                </div>
                <ul class="portal-features">
                    <li><i class="fas fa-check"></i> Accord生产</li>
                    <li><i class="fas fa-check"></i> 原料库存</li>
                    <li><i class="fas fa-check"></i> 基香&香调库存</li>
                    <li><i class="fas fa-check"></i> 车间调拨</li>
                </ul>
            </a>
            <% End If %>

            <% If InStr(accessiblePortals, "prodcenter") > 0 Then %>
            <a href="prodcenter/index.asp" class="portal-card prodcenter">
                <div class="card-top">
                    <div class="portal-icon"><i class="fas fa-industry"></i></div>
                    <div class="card-title">
                        <h3>产品生产管理中心 <span class="badge-new">NEW</span></h3>
                        <p>工单、排产、质检、入库</p>
                    </div>
                </div>
                <ul class="portal-features">
                    <li><i class="fas fa-check"></i> 生产工单</li>
                    <li><i class="fas fa-check"></i> 排产调度</li>
                    <li><i class="fas fa-check"></i> 质量检验</li>
                    <li><i class="fas fa-check"></i> 成品库存</li>
                </ul>
            </a>
            <% End If %>

            <% If InStr(accessiblePortals, "logistics") > 0 Then %>
            <a href="logistics/index.asp" class="portal-card logistics">
                <div class="card-top">
                    <div class="portal-icon"><i class="fas fa-truck"></i></div>
                    <div class="card-title">
                        <h3>物流管理中心 <span class="badge-new">NEW</span></h3>
                        <p>发货、在途、签收、退货</p>
                    </div>
                </div>
                <ul class="portal-features">
                    <li><i class="fas fa-check"></i> 发货单管理</li>
                    <li><i class="fas fa-check"></i> 在途跟踪</li>
                    <li><i class="fas fa-check"></i> 签收确认</li>
                    <li><i class="fas fa-check"></i> 退货入库</li>
                </ul>
            </a>
            <% End If %>

            <% If InStr(accessiblePortals, "purchase") > 0 Then %>
            <a href="purchase/index.asp" class="portal-card purchase">
                <div class="card-top">
                    <div class="portal-icon"><i class="fas fa-shopping-cart"></i></div>
                    <div class="card-title">
                        <h3>采购管理中心</h3>
                        <p>品牌定香产品采购与管理、供应商管理、价格管理</p>
                    </div>
                </div>
                <ul class="portal-features">
                    <li><i class="fas fa-check"></i> 采购订单</li>
                    <li><i class="fas fa-check"></i> 供应商管理</li>
                    <li><i class="fas fa-check"></i> 价格管理</li>
                    <li><i class="fas fa-check"></i> 采购分析</li>
                </ul>
            </a>
            <% End If %>

            <% If InStr(accessiblePortals, "finance") > 0 Then %>
            <a href="finance/index.asp" class="portal-card finance">
                <div class="card-top">
                    <div class="portal-icon"><i class="fas fa-dollar-sign"></i></div>
                    <div class="card-title">
                        <h3>财务管理中心</h3>
                        <p>应收应付、成本、总账</p>
                    </div>
                </div>
                <ul class="portal-features">
                    <li><i class="fas fa-check"></i> 应收应付</li>
                    <li><i class="fas fa-check"></i> 付款凭证</li>
                    <li><i class="fas fa-check"></i> 总账报表</li>
                    <li><i class="fas fa-check"></i> 现金流预测</li>
                </ul>
            </a>
            <% End If %>

            <% If InStr(accessiblePortals, "techcenter") > 0 Then %>
            <a href="techcenter/index.asp" class="portal-card techcenter">
                <div class="card-top">
                    <div class="portal-icon"><i class="fas fa-flask"></i></div>
                    <div class="card-title">
                        <h3>产品技术管理中心</h3>
                        <p>配方、基香、规格设置</p>
                    </div>
                </div>
                <ul class="portal-features">
                    <li><i class="fas fa-check"></i> 配方工艺</li>
                    <li><i class="fas fa-check"></i> 基香管理</li>
                    <li><i class="fas fa-check"></i> 香调调整</li>
                    <li><i class="fas fa-check"></i> 产品规格</li>
                </ul>
            </a>
            <% End If %>

            <% If InStr(accessiblePortals, "system") > 0 Then %>
            <a href="system/index.asp" class="portal-card system">
                <div class="card-top">
                    <div class="portal-icon"><i class="fas fa-shield-alt"></i></div>
                    <div class="card-title">
                        <h3>站点技术管理</h3>
                        <p>权限、安全、备份、日志</p>
                    </div>
                </div>
                <ul class="portal-features">
                    <li><i class="fas fa-check"></i> 管理员&角色</li>
                    <li><i class="fas fa-check"></i> 安全审计</li>
                    <li><i class="fas fa-check"></i> 备份中心</li>
                    <li><i class="fas fa-check"></i> 站点设置</li>
                </ul>
            </a>
            <% End If %>

            <% If InStr(accessiblePortals, "inventory") > 0 Then %>
            <a href="inventory/index.asp" class="portal-card inventory">
                <div class="card-top">
                    <div class="portal-icon"><i class="fas fa-warehouse"></i></div>
                    <div class="card-title">
                        <h3>库存管理中心 <span class="badge-new">NEW</span></h3>
                        <p>全品类统一库存管控</p>
                    </div>
                </div>
                <ul class="portal-features">
                    <li><i class="fas fa-check"></i> 库存仪表盘</li>
                    <li><i class="fas fa-check"></i> 全品类预警</li>
                    <li><i class="fas fa-check"></i> 库存流水</li>
                    <li><i class="fas fa-check"></i> 快捷入口</li>
                </ul>
            </a>
            <% End If %>

            <a href="logout.asp" class="portal-card logout logout-card">
                <div class="card-top">
                    <div class="portal-icon"><i class="fas fa-sign-out-alt"></i></div>
                    <div class="card-title">
                        <h3>安全退出</h3>
                        <p>结束本次管理会话</p>
                    </div>
                </div>
            </a>
        </div>

        <div class="footer-info">
            香氛电商系统 <span>V9.0</span> · 全模块统一管理平台
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
