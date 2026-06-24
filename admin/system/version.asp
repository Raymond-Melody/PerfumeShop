<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
Call OpenConnection()

' 获取数据库统计
Dim dbSize, tableCount, lastBackup
dbSize = 80  ' MB (approximate)
tableCount = GetScalar("SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE'")
lastBackup = "2026-06-10"

' 获取模块页面数
Dim modFinance, modOperation, modPurchase, modSystem, modProdcenter, modTechcenter, modSemifinished, modLogistics, modInventory
modFinance = 24
modOperation = 22
modPurchase = 21
modSystem = 15
modProdcenter = 11
modTechcenter = 11
modSemifinished = 10
modLogistics = 8
modInventory = 3
Dim totalAdminPages : totalAdminPages = modFinance + modOperation + modPurchase + modSystem + modProdcenter + modTechcenter + modSemifinished + modLogistics + modInventory
Dim totalUserApiPages : totalUserApiPages = 60
Dim totalAllPages : totalAllPages = totalAdminPages + totalUserApiPages
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>版本信息 - 系统管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; }
        .version-banner {
            background: linear-gradient(135deg, #1e2a3a 0%, #162231 100%);
            border: 1px solid rgba(0,188,212,0.2);
            border-radius: 16px;
            padding: 30px 35px;
            margin-bottom: 25px;
            display: flex;
            align-items: center;
            gap: 30px;
        }
        .version-badge {
            flex-shrink: 0;
            width: 80px; height: 80px;
            background: linear-gradient(135deg, #00bcd4, #00838f);
            border-radius: 20px;
            display: flex; align-items: center; justify-content: center;
            font-size: 28px; font-weight: 700; color: #fff;
        }
        .version-info h2 { color: #fff; font-size: 24px; margin-bottom: 5px; }
        .version-info .ver-sub { color: #78909c; font-size: 14px; }
        .version-info .ver-meta { display: flex; gap: 20px; margin-top: 12px; flex-wrap: wrap; }
        .ver-meta-item { 
            background: rgba(0,188,212,0.08); 
            padding: 6px 14px; border-radius: 20px; 
            font-size: 13px; color: #80deea; 
        }
        
        .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 25px; }
        @media (max-width: 768px) { .info-grid { grid-template-columns: 1fr; } }
        
        .info-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border: 1px solid rgba(255,255,255,0.06);
            border-radius: 12px; padding: 22px;
        }
        .info-card h3 { 
            color: #00bcd4; font-size: 16px; margin-bottom: 15px;
            padding-bottom: 10px; border-bottom: 1px solid rgba(255,255,255,0.06);
            display: flex; align-items: center; gap: 8px;
        }
        .info-card h3 i { font-size: 14px; }
        
        .module-bar { 
            display: flex; align-items: center; margin-bottom: 10px;
        }
        .module-name { 
            width: 130px; font-size: 13px; color: #b0b0b0; flex-shrink: 0;
        }
        .module-track { 
            flex: 1; height: 22px; background: rgba(255,255,255,0.04);
            border-radius: 11px; overflow: hidden; position: relative;
        }
        .module-fill { 
            height: 100%; border-radius: 11px; 
            display: flex; align-items: center; justify-content: flex-end;
            padding-right: 10px; font-size: 11px; color: #fff; font-weight: 600;
        }
        
        .table-mini { width: 100%; border-collapse: collapse; }
        .table-mini th { 
            text-align: left; padding: 8px 12px; font-size: 12px; color: #888; 
            border-bottom: 1px solid rgba(255,255,255,0.06); 
        }
        .table-mini td { 
            padding: 10px 12px; font-size: 13px; color: #d0d6e0;
            border-bottom: 1px solid rgba(255,255,255,0.03);
        }
        .table-mini tr:hover td { background: rgba(0,188,212,0.04); }
        
        .tag-ok { display: inline-block; padding: 2px 10px; background: rgba(76,175,80,0.2); color: #81c784; border-radius: 10px; font-size: 11px; }
        .tag-warn { display: inline-block; padding: 2px 10px; background: rgba(255,152,0,0.2); color: #ffb74d; border-radius: 10px; font-size: 11px; }
        
        .footer-ver { 
            text-align: center; color: #455a64; font-size: 12px; 
            padding: 20px 0; margin-top: 10px;
            border-top: 1px solid rgba(255,255,255,0.04);
        }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-code-branch"></i> 版本信息</h2>
            <div class="breadcrumb">
                <a href="index.asp">系统中心</a> / <span>版本信息</span>
            </div>
        </div>

        <!-- 版本概览 -->
        <div class="version-banner">
            <div class="version-badge"><%= SYS_VERSION %></div>
            <div class="version-info">
                <h2><%= SITE_NAME %> · 香氛定制系统</h2>
                <p class="ver-sub"><%= SYS_VERSION_NAME %></p>
                <div class="ver-meta">
                    <span class="ver-meta-item"><i class="fas fa-database"></i> SQL Server 2017</span>
                    <span class="ver-meta-item"><i class="fas fa-server"></i> IIS / Classic ASP</span>
                    <span class="ver-meta-item"><i class="fas fa-table"></i> <%= tableCount %> 张数据表</span>
                    <span class="ver-meta-item"><i class="fas fa-file-code"></i> <%= totalAllPages %> 个ASP页面</span>
                </div>
            </div>
        </div>

        <!-- 模块结构 -->
        <div class="info-grid">
            <div class="info-card">
                <h3><i class="fas fa-cubes"></i> 管理模块分布</h3>
                <%
                Dim modules(8, 1)
                modules(0, 0) = "finance"    : modules(0, 1) = 24
                modules(1, 0) = "operation"   : modules(1, 1) = 22
                modules(2, 0) = "purchase"    : modules(2, 1) = 21
                modules(3, 0) = "system"      : modules(3, 1) = 15
                modules(4, 0) = "prodcenter"  : modules(4, 1) = 11
                modules(5, 0) = "techcenter"  : modules(5, 1) = 11
                modules(6, 0) = "semifinished": modules(6, 1) = 10
                modules(7, 0) = "logistics"   : modules(7, 1) = 8
                modules(8, 0) = "inventory"   : modules(8, 1) = 3
                
                Dim maxPages : maxPages = 24
                Dim i, pct, color
                For i = 0 To 8
                    pct = Int(modules(i, 1) / maxPages * 100)
                    Select Case i
                        Case 0: color = "#00bcd4"
                        Case 1: color = "#667eea"
                        Case 2: color = "#FF5722"
                        Case 3: color = "#E91E63"
                        Case 4: color = "#4CAF50"
                        Case 5: color = "#9C27B0"
                        Case 6: color = "#FF9800"
                        Case 7: color = "#2196F3"
                        Case 8: color = "#00BCD4"
                    End Select
                %>
                <div class="module-bar">
                    <span class="module-name"><%= modules(i, 0) %></span>
                    <div class="module-track">
                        <div class="module-fill" style="width:<%= pct %>%;background:<%= color %>;">
                            <%= modules(i, 1) %>页
                        </div>
                    </div>
                </div>
                <% Next %>
                <div style="text-align:right;color:#666;font-size:12px;margin-top:8px;">
                    合计: <%= totalAdminPages %> 管理页面 + <%= totalUserApiPages %> 用户/API页面
                </div>
            </div>

            <div class="info-card">
                <h3><i class="fas fa-cog"></i> 系统配置</h3>
                <table class="table-mini">
                    <tr><th style="width:40%">配置项</th><th>值</th></tr>
                    <tr><td>免运费门槛</td><td>≥ <%= FREE_SHIPPING_AMOUNT %> 元</td></tr>
                    <tr><td>运费</td><td><%= SHIPPING_FEE %> 元</td></tr>
                    <tr><td>分页大小</td><td><%= PAGE_SIZE %> 条/页</td></tr>
                    <tr><td>Session 超时</td><td>60 分钟</td></tr>
                    <tr><td>安全响应头</td><td><span class="tag-ok">已启用</span></td></tr>
                    <tr><td>数据库大小</td><td>≈72 MB (数据) + 8 MB (日志)</td></tr>
                </table>
            </div>
        </div>

        <!-- 数据库表清单 -->
        <div class="info-card" style="margin-bottom:25px;">
            <h3><i class="fas fa-history"></i> V9.x 关键数据库变更</h3>
            <table class="table-mini">
                <tr><th>变更类型</th><th>表/对象名</th><th>说明</th></tr>
                <tr><td>新建表</td><td>FixedBrandProducts, FixedBrandPurchaseOrders, FixedBrandPurchaseDetails</td><td>品牌定香模块独立表体系</td></tr>
                <tr><td>新建表</td><td>FixedBrandInventory, FixedBrandReceipts, FixedBrandReceiptDetails</td><td>品牌定香库存与收货</td></tr>
                <tr><td>新建表</td><td>FixedBrandCostAllocation, OrderCostAllocation</td><td>成本分摊</td></tr>
                <tr><td>新建表</td><td>CostCenters, GLTransactions, AccountsPayable, AccountsReceivable</td><td>财务模块基础表</td></tr>
                <tr><td>添加列</td><td>PaymentRecords</td><td>新增 VoucherNo, PaymentType, CenterID, PayableID, ReceivableID</td></tr>
            </table>
        </div>

        <!-- 已知开发模式 -->
        <div class="info-card" style="margin-bottom:25px;">
            <h3><i class="fas fa-lightbulb"></i> 开发注意事项（供后续迭代参考）</h3>
            <table class="table-mini">
                <tr><th style="width:5%">#</th><th style="width:35%">模式</th><th>解决方案</th></tr>
                <tr>
                    <td>1</td>
                    <td>conn.Execute + MoveLast 报错 80040e24</td>
                    <td>使用 Server.CreateObject("ADODB.Recordset") + CursorLocation=3 (客户端游标)</td>
                </tr>
                <tr>
                    <td>2</td>
                    <td>ASP 中 CREATE TABLE 报权限错误</td>
                    <td>IIS 应用池无 DDL 权限，需通过 sqlcmd 预执行建表脚本</td>
                </tr>
                <tr>
                    <td>3</td>
                    <td>PowerShell 脚本含中文字符串解析失败</td>
                    <td>使用 Join-Path、Get-Location 构造路径，避免 .ps1 中直接写中文路径字面量</td>
                </tr>
                <tr>
                    <td>4</td>
                    <td>sqlcmd 管道传输中文路径乱码</td>
                    <td>备份时先输出到 C:\temp\（纯英文路径），再用 [System.IO.File]::Copy() 复制到目标</td>
                </tr>
            </table>
        </div>

        <!-- 备份记录 -->
        <div class="info-card" style="margin-bottom:25px;">
            <h3><i class="fas fa-archive"></i> 最新备份记录</h3>
            <table class="table-mini">
                <tr><th style="width:30%">备份项</th><th>文件名</th><th style="width:15%">大小</th><th style="width:15%">状态</th></tr>
                <tr>
                    <td>数据库完整备份</td>
                    <td>PerfumeShop_full_20260610_145134.bak</td>
                    <td>13.11 MB</td>
                    <td><span class="tag-ok">已验证</span></td>
                </tr>
                <tr>
                    <td>代码完整备份</td>
                    <td>PerfumeShop_code_20260610_145415.zip</td>
                    <td>13.25 MB</td>
                    <td><span class="tag-ok">已完成</span></td>
                </tr>
            </table>
            <p style="color:#666;font-size:12px;margin-top:10px;">
                <i class="fas fa-folder"></i> 备份路径: database\backups\<br>
                <i class="fas fa-clock"></i> 备份时间: <%= lastBackup %>
            </p>
        </div>

        <div class="footer-ver">
            <%= SITE_NAME %> · <%= SYS_VERSION %> · 版本快照生成于 <%= lastBackup %>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
