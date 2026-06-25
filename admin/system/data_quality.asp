<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<!--#include file="../../includes/dal.asp"-->
<!--#include file="includes/auth.asp"-->
<%
Call OpenConnection()

' V17: 数据质量检查 - 使用参数化DAL查询
Dim checks, i, qualityScore, totalChecks, passedChecks
totalChecks = 11
passedChecks = 0

' 定义检查项
ReDim checks(totalChecks - 1, 3) ' name, status(pass/fail/warn), count, detail

' 1. 价格异常检查（为0或负数的产品）
Dim priceIssueCount
priceIssueCount = CLng(DAL_GetScalar("SELECT COUNT(*) FROM Products WHERE Price <= 0 OR Price IS NULL", Null, 0))
If priceIssueCount = 0 Then passedChecks = passedChecks + 1
checks(0, 0) = "产品价格有效性"
checks(0, 1) = IIF(priceIssueCount = 0, "pass", "fail")
checks(0, 2) = priceIssueCount
checks(0, 3) = "检查价格 <= 0 或为 NULL 的产品"

' 2. 库存负数检查
Dim negativeStockCount
negativeStockCount = CLng(DAL_GetScalar("SELECT COUNT(*) FROM Products WHERE StockQuantity < 0", Null, 0))
If negativeStockCount = 0 Then passedChecks = passedChecks + 1
checks(1, 0) = "库存数量有效性"
checks(1, 1) = IIF(negativeStockCount = 0, "pass", "fail")
checks(1, 2) = negativeStockCount
checks(1, 3) = "检查 StockQuantity < 0 的产品"

' 3. 缺失产品名称
Dim missingNameCount
missingNameCount = CLng(DAL_GetScalar("SELECT COUNT(*) FROM Products WHERE ProductName IS NULL OR ProductName = ''", Null, 0))
If missingNameCount = 0 Then passedChecks = passedChecks + 1
checks(2, 0) = "产品名称完整性"
checks(2, 1) = IIF(missingNameCount = 0, "pass", "fail")
checks(2, 2) = missingNameCount
checks(2, 3) = "检查 ProductName 为空的产品"

' 4. 订单状态完整性
Dim orderStatusCount
orderStatusCount = CLng(DAL_GetScalar("SELECT COUNT(*) FROM Orders WHERE Status IS NULL OR Status = ''", Null, 0))
If orderStatusCount = 0 Then passedChecks = passedChecks + 1
checks(3, 0) = "订单状态完整性"
checks(3, 1) = IIF(orderStatusCount = 0, "pass", "fail")
checks(3, 2) = orderStatusCount
checks(3, 3) = "检查 Status 为空的订单"

' 5. 孤立订单（无关联用户）
Dim orphanOrderCount
orphanOrderCount = CLng(DAL_GetScalar("SELECT COUNT(*) FROM Orders o LEFT JOIN Users u ON o.UserID = u.UserID WHERE u.UserID IS NULL", Null, 0))
If orphanOrderCount = 0 Then passedChecks = passedChecks + 1
checks(4, 0) = "订单用户关联"
checks(4, 1) = IIF(orphanOrderCount = 0, "pass", "warn")
checks(4, 2) = orphanOrderCount
checks(4, 3) = "检查 Order.UserID 在 Users 中不存在"

' 6. 重复邮箱用户
Dim dupEmailCount
dupEmailCount = CLng(DAL_GetScalar("SELECT COUNT(*) FROM (SELECT Email FROM Users WHERE Email IS NOT NULL AND Email <> '' GROUP BY Email HAVING COUNT(*) > 1) AS dup", Null, 0))
If dupEmailCount = 0 Then passedChecks = passedChecks + 1
checks(5, 0) = "用户邮箱唯一性"
checks(5, 1) = IIF(dupEmailCount = 0, "pass", "warn")
checks(5, 2) = dupEmailCount
checks(5, 3) = "检查 Email 重复的用户"

' 7. 分类关联完整性
Dim orphanProductType
orphanProductType = CLng(DAL_GetScalar("SELECT COUNT(*) FROM Products p LEFT JOIN ProductTypes pt ON p.TypeCode = pt.TypeCode WHERE pt.TypeCode IS NULL AND p.TypeCode IS NOT NULL AND p.TypeCode <> ''", Null, 0))
If orphanProductType = 0 Then passedChecks = passedChecks + 1
checks(6, 0) = "产品分类关联"
checks(6, 1) = IIF(orphanProductType = 0, "pass", "warn")
checks(6, 2) = orphanProductType
checks(6, 3) = "检查 Product.TypeCode 在 ProductTypes 中不存在"

' 8. 购物车过期数据
Dim cartExpiredCount
cartExpiredCount = CLng(DAL_GetScalar("SELECT COUNT(*) FROM Cart WHERE CreatedAt < DATEADD(DAY, -30, GETDATE())", Null, 0))
If cartExpiredCount = 0 Then passedChecks = passedChecks + 1
checks(7, 0) = "购物车过期清理"
checks(7, 1) = IIF(cartExpiredCount = 0, "pass", "warn")
checks(7, 2) = cartExpiredCount
checks(7, 3) = "检查超过30天未清理的购物车记录"

' 9. 未支付的过期订单
Dim unpaidExpired
unpaidExpired = CLng(DAL_GetScalar("SELECT COUNT(*) FROM Orders WHERE Status='Pending' AND CreatedAt < DATEADD(DAY, -7, GETDATE())", Null, 0))
If unpaidExpired = 0 Then passedChecks = passedChecks + 1
checks(8, 0) = "过期未支付订单"
checks(8, 1) = IIF(unpaidExpired = 0, "pass", "warn")
checks(8, 2) = unpaidExpired
checks(8, 3) = "检查超过7天未支付的待付款订单"

' 10. 管理员审计日志完整性
Dim auditCount
auditCount = CLng(DAL_GetScalar("SELECT COUNT(*) FROM AdminAuditLog", Null, 0))
checks(9, 0) = "审计日志记录"
checks(9, 1) = IIF(auditCount > 0, "pass", "warn")
checks(9, 2) = auditCount
checks(9, 3) = "检查审计日志表是否有记录"

' 11. 图片缺失检查
Dim missingImageCount
missingImageCount = CLng(DAL_GetScalar("SELECT COUNT(*) FROM Products WHERE (ImageURL IS NULL OR ImageURL = '') AND IsActive <> 0", Null, 0))
If missingImageCount = 0 Then passedChecks = passedChecks + 1
checks(10, 0) = "产品图片完整性"
checks(10, 1) = IIF(missingImageCount = 0, "pass", "fail")
checks(10, 2) = missingImageCount
checks(10, 3) = "检查已激活产品是否缺少图片"

qualityScore = Round((passedChecks / totalChecks) * 100, 0)
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>数据质量控制 - 系统管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; }
        .score-circle { width: 120px; height: 120px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 0 auto 15px; font-size: 36px; font-weight: bold; color: #fff; }
        .score-excellent { background: linear-gradient(135deg, #43e97b 0%, #38f9d7 100%); }
        .score-good { background: linear-gradient(135deg, #fa709a 0%, #fee140 100%); }
        .score-poor { background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); }
        .check-item { display: flex; align-items: center; padding: 15px 20px; background: #2d2d44; border-radius: 8px; margin-bottom: 10px; border-left: 4px solid; transition: transform 0.2s; }
        .check-item:hover { transform: translateX(5px); }
        .check-pass { border-left-color: #43e97b; }
        .check-fail { border-left-color: #f5576c; }
        .check-warn { border-left-color: #ffd93d; }
        .check-icon { width: 30px; height: 30px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin-right: 15px; flex-shrink: 0; font-size: 14px; color: #fff; }
        .check-pass .check-icon { background: #43e97b; }
        .check-fail .check-icon { background: #f5576c; }
        .check-warn .check-icon { background: #ffd93d; color: #333; }
        .check-info { flex: 1; }
        .check-name { font-size: 15px; color: #fff; font-weight: 500; }
        .check-desc { font-size: 12px; color: #888; margin-top: 4px; }
        .check-count { font-size: 18px; font-weight: bold; padding: 0 15px; }
        .check-pass .check-count { color: #43e97b; }
        .check-fail .check-count { color: #f5576c; }
        .check-warn .check-count { color: #ffd93d; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-clipboard-check"></i> 数据质量控制</h2>
            <div class="breadcrumb">
                <a href="index.asp">系统中心</a> / <span>数据质量</span>
            </div>
        </div>
        
        <div class="stats-grid" style="grid-template-columns: repeat(2, 1fr); max-width: 500px; margin: 0 auto 30px;">
            <div class="stat-card">
                <% Dim scoreClass : scoreClass = "score-good"
                   If qualityScore >= 90 Then scoreClass = "score-excellent"
                   If qualityScore < 70 Then scoreClass = "score-poor" %>
                <div class="score-circle <%= scoreClass %>"><%= qualityScore %>%</div>
                <div class="stat-value" style="font-size: 16px; color: #e0e0e0;">数据质量评分</div>
            </div>
            <div class="stat-card">
                <div class="stat-value" style="font-size: 36px;"><%= passedChecks %> / <%= totalChecks %></div>
                <div class="stat-label" style="font-size: 14px;">通过检查项</div>
            </div>
        </div>
        
        <div class="dashboard-card">
            <h3><i class="fas fa-list"></i> 检查结果明细</h3>
            <% For i = 0 To totalChecks - 1 %>
            <div class="check-item check-<%= checks(i, 1) %>">
                <div class="check-icon">
                    <i class="fas <%= IIF(checks(i, 1)="pass", "fa-check", IIF(checks(i, 1)="fail", "fa-times", "fa-exclamation")) %>"></i>
                </div>
                <div class="check-info">
                    <div class="check-name"><%= Server.HTMLEncode(checks(i, 0)) %></div>
                    <div class="check-desc"><%= Server.HTMLEncode(checks(i, 3)) %></div>
                </div>
                <div class="check-count"><%= checks(i, 2) %></div>
            </div>
            <% Next %>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>