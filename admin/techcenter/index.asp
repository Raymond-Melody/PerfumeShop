<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<!--#include file="../../includes/dal.asp"-->
<!--#include file="../../includes/dal_techcenter.asp"-->
<%
Call OpenConnection()

' ========== 顶部统计卡片 ==========
' 产品总数
Dim totalProducts
totalProducts = GetScalar("SELECT COUNT(*) FROM Products")

' 香调总数
Dim totalFragranceNotes
totalFragranceNotes = GetScalar("SELECT COUNT(*) FROM FragranceNotes WHERE IsActive=1")

' 基香总数
Dim totalBaseNotes
totalBaseNotes = GetScalar("SELECT COUNT(*) FROM BaseNotes WHERE IsActive=1")

' 产品类型数
Dim totalProductTypes
totalProductTypes = GetScalar("SELECT COUNT(*) FROM ProductTypeConfig WHERE IsActive=1")

' ========== 供应链对接：配方发布状态 ==========
Dim accordPublished, productPublished, totalPublishLogs
accordPublished = DAL_TC_CountPublishedAccords()
productPublished = DAL_TC_CountPublishedProducts()
totalPublishLogs = DAL_TC_CountPublishLogs()

Dim totalRecipes, totalRecipeNotes
' V18: 使用 DAL
Dim totalRecipeCount
totalRecipeCount = DAL_TC_CountActiveRecipes()

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

' ========== 最近发布操作 ==========
Dim rsRecentPublish
Set rsRecentPublish = DAL_TC_GetRecentPublishLogs(5)

' ========== 底部最近更新 ==========
' 最近修改的5个产品
Dim rsRecentProducts
Set rsRecentProducts = ExecuteQuery("SELECT TOP 5 ProductID, ProductName, UpdatedAt FROM Products ORDER BY UpdatedAt DESC")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>技术概览 - 产品技术管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <style>
        /* 暗色主题基础 */
        body {
            background: #1a1a2e;
            color: #e0e0e0;
        }
        
        /* 顶部统计卡片 */
        .stats-section {
            margin-bottom: 30px;
        }
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 20px;
        }
        .stats-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 25px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            border: 1px solid rgba(255,255,255,0.05);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        .stats-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 6px 25px rgba(0,0,0,0.4);
        }
        .stats-header {
            display: flex;
            align-items: center;
            margin-bottom: 15px;
        }
        .stats-icon {
            width: 48px;
            height: 48px;
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 20px;
            margin-right: 12px;
        }
        .stats-icon.products { background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); }
        .stats-icon.fragrance { background: linear-gradient(135deg, #26c6da 0%, #0097a7 100%); }
        .stats-icon.base { background: linear-gradient(135deg, #4dd0e1 0%, #00acc1 100%); }
        .stats-icon.types { background: linear-gradient(135deg, #80deea 0%, #26c6da 100%); }
        .stats-label {
            font-size: 13px;
            color: #888;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .stats-value {
            font-size: 28px;
            font-weight: 700;
            color: #fff;
            margin-top: 5px;
        }
        
        /* 中部快捷入口 */
        .quick-section {
            margin-bottom: 30px;
        }
        .section-title {
            font-size: 18px;
            font-weight: 600;
            margin-bottom: 20px;
            color: #fff;
            display: flex;
            align-items: center;
        }
        .section-title i {
            margin-right: 10px;
            color: #00bcd4;
        }
        .quick-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
            gap: 20px;
        }
        .quick-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 30px 20px;
            text-align: center;
            border: 1px solid rgba(255,255,255,0.05);
            transition: all 0.3s ease;
            cursor: pointer;
            text-decoration: none;
            color: inherit;
            display: block;
        }
        .quick-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 10px 30px rgba(0,188,212,0.2);
            border-color: rgba(0,188,212,0.3);
        }
        .quick-icon {
            width: 60px;
            height: 60px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            margin: 0 auto 15px;
            font-size: 24px;
            color: white;
            transition: transform 0.3s ease;
        }
        .quick-card:hover .quick-icon {
            transform: scale(1.1);
        }
        .quick-icon.formula { background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); }
        .quick-icon.base { background: linear-gradient(135deg, #26c6da 0%, #0097a7 100%); }
        .quick-icon.note { background: linear-gradient(135deg, #4dd0e1 0%, #00acc1 100%); }
        .quick-icon.product { background: linear-gradient(135deg, #80deea 0%, #26c6da 100%); }
        .quick-icon.publish { background: linear-gradient(135deg, #FF9800 0%, #F57C00 100%); }
        .quick-title {
            font-size: 16px;
            font-weight: 600;
            color: #fff;
            margin-bottom: 8px;
        }
        .quick-desc {
            font-size: 12px;
            color: #888;
        }
        
        /* 底部最近更新 */
        .recent-section {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 20px;
            border: 1px solid rgba(255,255,255,0.05);
        }
        .recent-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 15px;
        }
        .recent-table th,
        .recent-table td {
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid rgba(255,255,255,0.05);
        }
        .recent-table th {
            font-size: 12px;
            color: #888;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            font-weight: 600;
        }
        .recent-table td {
            font-size: 14px;
            color: #e0e0e0;
        }
        .recent-table tr:hover td {
            background: rgba(255,255,255,0.02);
        }
        .product-name {
            display: flex;
            align-items: center;
        }
        .product-name i {
            color: #00bcd4;
            margin-right: 10px;
        }
        .update-time {
            color: #888;
            font-size: 12px;
        }
        .empty-state {
            text-align: center;
            padding: 40px;
            color: #666;
        }
        .empty-state i {
            font-size: 48px;
            margin-bottom: 15px;
            color: #444;
        }
        
        /* 响应式 */
        @media (max-width: 1200px) {
            .stats-grid { grid-template-columns: repeat(2, 1fr); }
            .quick-grid { grid-template-columns: repeat(2, 1fr); }
        }
        @media (max-width: 768px) {
            .stats-grid { grid-template-columns: 1fr; }
            .quick-grid { grid-template-columns: 1fr; }
        }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-flask"></i> 技术概览</h2>
            <div class="breadcrumb">
                <a href="index.asp">技术中心</a> / <span>概览</span>
            </div>
        </div>
        
        <!-- 顶部：统计卡片 -->
        <div class="stats-section">
            <div class="stats-grid">
                <div class="stats-card">
                    <div class="stats-header">
                        <div class="stats-icon products"><i class="fas fa-box-open"></i></div>
                        <div class="stats-label">产品总数</div>
                    </div>
                    <div class="stats-value"><%= totalProducts %></div>
                </div>
                <div class="stats-card">
                    <div class="stats-header">
                        <div class="stats-icon fragrance"><i class="fas fa-leaf"></i></div>
                        <div class="stats-label">香调总数</div>
                    </div>
                    <div class="stats-value"><%= totalFragranceNotes %></div>
                </div>
                <div class="stats-card">
                    <div class="stats-header">
                        <div class="stats-icon base"><i class="fas fa-wine-bottle"></i></div>
                        <div class="stats-label">基香总数</div>
                    </div>
                    <div class="stats-value"><%= totalBaseNotes %></div>
                </div>
                <div class="stats-card">
                    <div class="stats-header">
                        <div class="stats-icon types"><i class="fas fa-tags"></i></div>
                        <div class="stats-label">产品类型数</div>
                    </div>
                    <div class="stats-value"><%= totalProductTypes %></div>
                </div>
            </div>
        </div>
        
        <!-- 供应链对接：配方发布状态 -->
        <div class="stats-section">
            <div class="section-title"><i class="fas fa-link" style="color:#00bcd4;"></i> 配方发布状态（供应链对接）</div>
            <div class="stats-grid">
                <div class="stats-card">
                    <div class="stats-header">
                        <div class="stats-icon" style="background:linear-gradient(135deg,#FF9800,#F57C00);"><i class="fas fa-flask"></i></div>
                        <div class="stats-label">已发布香调配方</div>
                    </div>
                    <div class="stats-value"><%=accordPublished%></div>
                    <div class="stats-sub">→ 香调车间可见</div>
                </div>
                <div class="stats-card">
                    <div class="stats-header">
                        <div class="stats-icon" style="background:linear-gradient(135deg,#2196F3,#1976D2);"><i class="fas fa-industry"></i></div>
                        <div class="stats-label">已发布产品配方</div>
                    </div>
                    <div class="stats-value"><%=productPublished%></div>
                    <div class="stats-sub">→ 制造车间可见</div>
                </div>
                <div class="stats-card">
                    <div class="stats-header">
                        <div class="stats-icon" style="background:linear-gradient(135deg,#4CAF50,#388E3C);"><i class="fas fa-clipboard-list"></i></div>
                        <div class="stats-label">活跃配方总数</div>
                    </div>
                    <div class="stats-value"><%=totalRecipeCount%></div>
                    <div class="stats-sub">Recipes</div>
                </div>
                <div class="stats-card">
                    <div class="stats-header">
                        <div class="stats-icon" style="background:linear-gradient(135deg,#9C27B0,#7B1FA2);"><i class="fas fa-history"></i></div>
                        <div class="stats-label">审计日志总数</div>
                    </div>
                    <div class="stats-value"><%=totalPublishLogs%></div>
                    <div class="stats-sub">RecipePublishLog</div>
                </div>
            </div>
        </div>
        
        <!-- 中部：快捷入口 -->
        <div class="quick-section">
            <div class="section-title"><i class="fas fa-th-large"></i> 快捷入口</div>
            <div class="quick-grid">
                <a href="formula_management.asp" class="quick-card">
                    <div class="quick-icon formula"><i class="fas fa-vial"></i></div>
                    <div class="quick-title">配方设置</div>
                    <div class="quick-desc">管理产品配方与工艺</div>
                </a>
                <a href="base_note_management.asp" class="quick-card">
                    <div class="quick-icon base"><i class="fas fa-wine-bottle"></i></div>
                    <div class="quick-title">基香管理</div>
                    <div class="quick-desc">基香成分配置</div>
                </a>
                <a href="note_management.asp" class="quick-card">
                    <div class="quick-icon note"><i class="fas fa-leaf"></i></div>
                    <div class="quick-title">香调管理</div>
                    <div class="quick-desc">香调配比调整</div>
                </a>
                <a href="product_settings.asp" class="quick-card">
                    <div class="quick-icon product"><i class="fas fa-box-open"></i></div>
                    <div class="quick-title">产品设置</div>
                    <div class="quick-desc">用户定制/KOL推荐 产品规格配置</div>
                </a>
                <% If isManager Then %>
                <a href="recipe_publish.asp" class="quick-card">
                    <div class="quick-icon publish"><i class="fas fa-lock"></i></div>
                    <div class="quick-title">配方拆分发布</div>
                    <div class="quick-desc">拆分下发供应链配方</div>
                </a>
                <% End If %>
            </div>
        </div>
        
        <!-- 最近发布操作 -->
        <div class="recent-section" style="margin-bottom:20px;">
            <div class="section-title"><i class="fas fa-history"></i> 最近发布操作</div>
            <%
            Dim hasPublishData
            hasPublishData = False
            If Not (rsRecentPublish Is Nothing) Then
                If Not rsRecentPublish.EOF Then
                    hasPublishData = True
                End If
            End If
            If hasPublishData Then
            %>
            <table class="recent-table">
                <thead>
                    <tr>
                        <th>时间</th>
                        <th>类型</th>
                        <th>发布人</th>
                    </tr>
                </thead>
                <tbody>
                    <% Do While Not rsRecentPublish.EOF %>
                    <tr>
                        <td class="update-time">
                            <i class="far fa-clock"></i> <%= rsRecentPublish("PublishedAt") & "" %>
                        </td>
                        <td>
                            <% If CStr(rsRecentPublish("PublishType") & "") = "Accord" Then %>
                            <span style="color:#FF9800;"><i class="fas fa-flask"></i> 香调配方</span>
                            <% Else %>
                            <span style="color:#2196F3;"><i class="fas fa-industry"></i> 产品配方</span>
                            <% End If %>
                        </td>
                        <td><%= Server.HTMLEncode(rsRecentPublish("PublishedBy") & "") %></td>
                    </tr>
                    <% rsRecentPublish.MoveNext %>
                    <% Loop %>
                </tbody>
            </table>
            <% Else %>
            <div class="empty-state" style="padding:20px;">
                <i class="fas fa-inbox"></i>
                <p>暂无发布操作记录</p>
            </div>
            <% End If %>
        </div>
        
        <!-- 底部：最近更新 -->
        <div class="recent-section">
            <div class="section-title"><i class="fas fa-clock"></i> 最近更新</div>
            <% 
            Dim hasRecentData
            hasRecentData = False
            If Not (rsRecentProducts Is Nothing) Then
                If Not rsRecentProducts.EOF Then
                    hasRecentData = True
                End If
            End If
            If hasRecentData Then
            %>
            <table class="recent-table">
                <thead>
                    <tr>
                        <th>产品名称</th>
                        <th>更新时间</th>
                    </tr>
                </thead>
                <tbody>
                    <% Do While Not rsRecentProducts.EOF %>
                    <tr>
                        <td>
                            <div class="product-name">
                                <i class="fas fa-box"></i>
                                <%= HTMLEncode(rsRecentProducts("ProductName")) %>
                            </div>
                        </td>
                        <td class="update-time">
                            <i class="far fa-clock"></i> <%= FormatDateField(rsRecentProducts("UpdatedAt")) %>
                        </td>
                    </tr>
                    <% rsRecentProducts.MoveNext %>
                    <% Loop %>
                </tbody>
            </table>
            <% Else %>
            <div class="empty-state">
                <i class="fas fa-inbox"></i>
                <p>暂无最近更新的产品</p>
            </div>
            <% End If %>
        </div>
    </div>
</body>
</html>
<%
If Not rsRecentProducts Is Nothing Then
    rsRecentProducts.Close
    Set rsRecentProducts = Nothing
End If
If Not rsRecentPublish Is Nothing Then
    If rsRecentPublish.State = 1 Then rsRecentPublish.Close
    Set rsRecentPublish = Nothing
End If
Call CloseConnection()
%>
