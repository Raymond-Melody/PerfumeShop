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

' 获取统计配置
Dim enableStats
enableStats = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey = 'EnableStatistics'")
If IsNull(enableStats) Then enableStats = "1"

' 获取销售统计
Dim totalOrders, totalRevenue, totalUsers, todayOrders
totalOrders = GetScalar("SELECT COUNT(*) FROM Orders")
totalRevenue = GetScalar("SELECT SUM(TotalAmount) FROM Orders WHERE Status = 'Paid'")
If IsNull(totalRevenue) Or totalRevenue = "" Then totalRevenue = 0 Else totalRevenue = CDbl(totalRevenue)
totalUsers = GetScalar("SELECT COUNT(*) FROM Users")
todayOrders = GetScalar("SELECT COUNT(*) FROM Orders WHERE CAST(CreatedAt AS DATE) = CAST(GETDATE() AS DATE)")

' 获取商品统计
Dim rsProductStats
Set rsProductStats = ExecuteQuery(_
    "SELECT TOP 5 p.ProductName, COUNT(od.DetailID) as SaleCount, SUM(od.Quantity) as TotalQty " & _
    "FROM OrderDetails od " & _
    "INNER JOIN Products p ON od.ProductID = p.ProductID " & _
    "GROUP BY p.ProductName " & _
    "ORDER BY SaleCount DESC")

' 获取香调统计
Dim rsNoteStats
Set rsNoteStats = ExecuteQuery(_
    "SELECT TOP 5 fn.NoteName, fn.NoteType, COUNT(cns.SelectionID) as UseCount " & _
    "FROM CartNoteSelections cns " & _
    "INNER JOIN FragranceNotes fn ON cns.NoteID = fn.NoteID " & _
    "GROUP BY fn.NoteName, fn.NoteType " & _
    "ORDER BY UseCount DESC")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>数据统计 - 后台管理</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; }
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin-bottom: 30px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 25px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); text-align: center; }
        .stat-icon { font-size: 36px; margin-bottom: 10px; }
        .stat-card.orders .stat-icon { color: #2196F3; }
        .stat-card.revenue .stat-icon { color: #4CAF50; }
        .stat-card.users .stat-icon { color: #FF9800; }
        .stat-card.today .stat-icon { color: #9C27B0; }
        .stat-value { font-size: 32px; font-weight: bold; color: #fff; }
        .stat-label { color: #888; margin-top: 5px; }
        
        .chart-container { display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; margin-bottom: 30px; }
        .chart-box { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 20px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
        .chart-title { font-size: 18px; font-weight: bold; margin-bottom: 15px; color: #e0e0e0; border-left: 4px solid #4CAF50; padding-left: 10px; }
        
        .rank-list { list-style: none; padding: 0; margin: 0; }
        .rank-item { display: flex; align-items: center; padding: 12px 0; border-bottom: 1px solid rgba(255,255,255,0.06); }
        .rank-number { width: 30px; height: 30px; border-radius: 50%; background: rgba(255,255,255,0.08); display: flex; align-items: center; justify-content: center; font-weight: bold; margin-right: 15px; color: #e0e0e0; }
        .rank-item:nth-child(1) .rank-number { background: #FFD700; color: white; }
        .rank-item:nth-child(2) .rank-number { background: #C0C0C0; color: white; }
        .rank-item:nth-child(3) .rank-number { background: #CD7F32; color: white; }
        .rank-info { flex: 1; }
        .rank-name { font-weight: bold; color: #e0e0e0; }
        .rank-meta { font-size: 12px; color: #888; }
        .rank-value { font-size: 18px; font-weight: bold; color: #4CAF50; }
        
        .settings-panel { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 20px; border-radius: 8px; margin-bottom: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
        .settings-panel h3 { color: #e0e0e0; }
        .settings-panel p { color: #b0b0b0; }
        .status-enabled { color: #4CAF50; font-weight: bold; }
        .status-disabled { color: #f44336; font-weight: bold; }
        .chart-box td { padding: 10px; border-bottom: 1px solid rgba(255,255,255,0.06); color: #e0e0e0; }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="admin-card">
            <div class="admin-card-header">
                <h2 class="admin-card-title"><i class="fas fa-chart-line"></i> 数据统计</h2>
                <a href="index.asp" class="admin-btn admin-btn-secondary"><i class="fas fa-arrow-left"></i> 返回首页</a>
            </div>
            
            <!-- 功能开关 -->
            <div class="settings-panel">
                <h3><i class="fas fa-cog"></i> 功能设置</h3>
                <p>数据统计功能: 
                    <% If enableStats = "1" Then %>
                    <span class="status-enabled"><i class="fas fa-check-circle"></i> 已启用</span>
                    <% Else %>
                    <span class="status-disabled"><i class="fas fa-times-circle"></i> 已禁用</span>
                    <% End If %>
                </p>
            </div>
            
            <!-- 核心指标 -->
            <div class="stats-grid">
                <div class="stat-card orders">
                    <div class="stat-icon"><i class="fas fa-shopping-cart"></i></div>
                    <div class="stat-value"><%= totalOrders %></div>
                    <div class="stat-label">总订单数</div>
                </div>
                <div class="stat-card revenue">
                    <div class="stat-icon"><i class="fas fa-yen-sign"></i></div>
                    <div class="stat-value">¥<%= FormatNumber(totalRevenue, 2) %></div>
                    <div class="stat-label">总营收</div>
                </div>
                <div class="stat-card users">
                    <div class="stat-icon"><i class="fas fa-users"></i></div>
                    <div class="stat-value"><%= totalUsers %></div>
                    <div class="stat-label">注册用户</div>
                </div>
                <div class="stat-card today">
                    <div class="stat-icon"><i class="fas fa-calendar-day"></i></div>
                    <div class="stat-value"><%= todayOrders %></div>
                    <div class="stat-label">今日订单</div>
                </div>
            </div>
            
            <!-- 排行榜 -->
            <div class="chart-container">
                <div class="chart-box">
                    <div class="chart-title">🏆 热销商品TOP5</div>
                    <ul class="rank-list">
                        <% If Not rsProductStats Is Nothing Then %>
                        <% Dim prodRank : prodRank = 0 %>
                        <% Do While Not rsProductStats.EOF %>
                        <% prodRank = prodRank + 1 %>
                        <li class="rank-item">
                            <div class="rank-number"><%= prodRank %></div>
                            <div class="rank-info">
                                <div class="rank-name"><%= rsProductStats("ProductName") %></div>
                                <div class="rank-meta">销售 <%= rsProductStats("SaleCount") %> 次</div>
                            </div>
                            <div class="rank-value"><%= rsProductStats("TotalQty") %> 件</div>
                        </li>
                        <% rsProductStats.MoveNext %>
                        <% Loop %>
                        <% rsProductStats.Close %>
                        <% End If %>
                    </ul>
                </div>
                
                <div class="chart-box">
                    <div class="chart-title">🎯 热门香调TOP5</div>
                    <ul class="rank-list">
                        <% If Not rsNoteStats Is Nothing Then %>
                        <% Dim noteRank : noteRank = 0 %>
                        <% Do While Not rsNoteStats.EOF %>
                        <% noteRank = noteRank + 1 %>
                        <li class="rank-item">
                            <div class="rank-number"><%= noteRank %></div>
                            <div class="rank-info">
                                <div class="rank-name"><%= rsNoteStats("NoteName") %></div>
                                <div class="rank-meta"><%= rsNoteStats("NoteType") %></div>
                            </div>
                            <div class="rank-value"><%= rsNoteStats("UseCount") %> 次</div>
                        </li>
                        <% rsNoteStats.MoveNext %>
                        <% Loop %>
                        <% rsNoteStats.Close %>
                        <% End If %>
                    </ul>
                </div>
            </div>
            
            <!-- 其他统计信息 -->
            <div class="chart-container">
                <div class="chart-box">
                    <div class="chart-title">📊 系统概览</div>
                    <table style="width: 100%; border-collapse: collapse;">
                        <tr>
                            <td style="padding: 10px; border-bottom: 1px solid #eee;">商品总数</td>
                            <td style="padding: 10px; border-bottom: 1px solid #eee; text-align: right; font-weight: bold;"><%= GetScalar("SELECT COUNT(*) FROM Products") %></td>
                        </tr>
                        <tr>
                            <td style="padding: 10px; border-bottom: 1px solid #eee;">香调种类</td>
                            <td style="padding: 10px; border-bottom: 1px solid #eee; text-align: right; font-weight: bold;"><%= GetScalar("SELECT COUNT(*) FROM FragranceNotes") %></td>
                        </tr>
                        <tr>
                            <td style="padding: 10px; border-bottom: 1px solid #eee;">库存预警</td>
                            <td style="padding: 10px; border-bottom: 1px solid #eee; text-align: right; font-weight: bold; color: #f44336;"><%= GetScalar("SELECT COUNT(*) FROM NoteInventory WHERE StockQuantity <= MinStockLevel") %></td>
                        </tr>
                        <tr>
                            <td style="padding: 10px;">生产中订单</td>
                            <td style="padding: 10px; text-align: right; font-weight: bold; color: #2196F3;"><%= GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE Status = '生产中'") %></td>
                        </tr>
                    </table>
                </div>
                
                <div class="chart-box">
                    <div class="chart-title">💰 财务概览</div>
                    <table style="width: 100%; border-collapse: collapse;">
                        <tr>
                            <td style="padding: 10px; border-bottom: 1px solid #eee;">待付款订单</td>
                            <td style="padding: 10px; border-bottom: 1px solid #eee; text-align: right; font-weight: bold;"><%= GetScalar("SELECT COUNT(*) FROM Orders WHERE Status = 'Pending'") %></td>
                        </tr>
                        <tr>
                            <td style="padding: 10px; border-bottom: 1px solid #eee;">已完成订单</td>
                            <td style="padding: 10px; border-bottom: 1px solid #eee; text-align: right; font-weight: bold; color: #4CAF50;"><%= GetScalar("SELECT COUNT(*) FROM Orders WHERE Status = 'Paid'") %></td>
                        </tr>
                        <tr>
                            <td style="padding: 10px; border-bottom: 1px solid #eee;">积分总发放</td>
                            <td style="padding: 10px; text-align: right; font-weight: bold;"><%= GetScalar("SELECT ISNULL(SUM(TotalPoints), 0) FROM UserPoints") %></td>
                        </tr>
                        <tr>
                            <td style="padding: 10px;">平均客单价</td>
                            <td style="padding: 10px; text-align: right; font-weight: bold;">¥<%= FormatNumber(CDbl("0" & GetScalar("SELECT ISNULL(AVG(TotalAmount), 0) FROM Orders")), 2) %></td>
                        </tr>
                    </table>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
