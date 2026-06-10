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

' 自动创建 BottleInventory 表（如不存在）- 使用 IF OBJECT_ID 确保可靠创建
On Error Resume Next
conn.Execute "IF OBJECT_ID('BottleInventory','U') IS NULL CREATE TABLE BottleInventory (BottleID INT IDENTITY(1,1) PRIMARY KEY, BottleName NVARCHAR(100), StockQty DECIMAL(19,4) DEFAULT 0, SafetyStock DECIMAL(19,4) DEFAULT 0, UnitCost DECIMAL(19,4) DEFAULT 0, IsActive BIT DEFAULT 1, UpdatedAt DATETIME DEFAULT GETDATE())"
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0

' 自动检查 PackagingInventory 表
On Error Resume Next
conn.Execute "IF OBJECT_ID('PackagingInventory','U') IS NULL CREATE TABLE PackagingInventory (PackagingID INT IDENTITY(1,1) PRIMARY KEY, ItemName NVARCHAR(100), StockQty INT DEFAULT 0, SafetyStock INT DEFAULT 0, UnitCost DECIMAL(19,4) DEFAULT 0, IsActive BIT DEFAULT 1, UpdatedAt DATETIME DEFAULT GETDATE())"
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Function GetScalar(sql)
    Dim rs, val : val = 0
    On Error Resume Next
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        If Not rs Is Nothing Then
            If Not rs.EOF Then val = rs(0)
            If IsNull(val) Then val = 0
            rs.Close
        End If
    Else : Err.Clear
    End If
    Set rs = Nothing
    GetScalar = val
End Function

' ========== 成品库存统计 ==========
Dim prodTotal, prodTotalQty, prodLowStock, prodValue
prodTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductInventory"))
prodTotalQty = SafeNum(GetScalar("SELECT ISNULL(SUM(StockQty),0) FROM ProductInventory"))
prodLowStock = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0"))
prodValue = SafeNum(GetScalar("SELECT ISNULL(SUM(StockQty * UnitCost),0) FROM ProductInventory"))

' ========== 瓶子库存统计 ==========
Dim bottleTotal, bottleTotalQty, bottleLowStock, bottleValue
bottleTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM BottleInventory"))
bottleTotalQty = SafeNum(GetScalar("SELECT ISNULL(SUM(StockQty),0) FROM BottleInventory"))
bottleLowStock = SafeNum(GetScalar("SELECT COUNT(*) FROM BottleInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0"))
bottleValue = SafeNum(GetScalar("SELECT ISNULL(SUM(StockQty * UnitCost),0) FROM BottleInventory"))

' ========== 包装物库存统计 ==========
Dim packTotal, packTotalQty, packLowStock, packValue
packTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM PackagingInventory"))
packTotalQty = SafeNum(GetScalar("SELECT ISNULL(SUM(StockQty),0) FROM PackagingInventory"))
packLowStock = SafeNum(GetScalar("SELECT COUNT(*) FROM PackagingInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0"))
packValue = SafeNum(GetScalar("SELECT ISNULL(SUM(StockQty * UnitCost),0) FROM PackagingInventory"))

' ========== 原料库存统计 ==========
Dim rawTotal, rawTotalQty, rawLowStock, rawValue
rawTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM RawMaterialInventory"))
rawTotalQty = SafeNum(GetScalar("SELECT ISNULL(SUM(StockQty),0) FROM RawMaterialInventory"))
rawLowStock = SafeNum(GetScalar("SELECT COUNT(*) FROM RawMaterialInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0"))
rawValue = SafeNum(GetScalar("SELECT ISNULL(SUM(StockQty * UnitPrice),0) FROM RawMaterialInventory"))

' ========== 香调库存统计 ==========
Dim noteTotal, noteTotalQty, noteLowStock
noteTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM NoteInventory"))
noteTotalQty = SafeNum(GetScalar("SELECT ISNULL(SUM(StockQuantity),0) FROM NoteInventory"))
noteLowStock = SafeNum(GetScalar("SELECT COUNT(*) FROM NoteInventory WHERE StockQuantity <= MinStockLevel AND MinStockLevel > 0"))

' ========== 基香统计 ==========
Dim bnTotal, bnActive
bnTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM BaseNotes"))
bnActive = SafeNum(GetScalar("SELECT COUNT(*) FROM BaseNotes WHERE IsActive=1"))

' ========== 总计 ==========
Dim totalCategories, totalItems, totalLowStock, totalValue
totalCategories = 6
totalItems = prodTotal + bottleTotal + packTotal + rawTotal + noteTotal + bnTotal
totalLowStock = prodLowStock + bottleLowStock + packLowStock + rawLowStock + noteLowStock
totalValue = prodValue + bottleValue + packValue + rawValue

' ========== 全品类低库存预警 TOP 10 ==========
Dim rsAllAlerts
On Error Resume Next
Set rsAllAlerts = conn.Execute(_
    "SELECT Category, ItemName, StockQty, AlertStock, UnitName FROM (" & _
    "SELECT TOP 10 '成品' AS Category, p.ProductName AS ItemName, pi.StockQty, pi.SafetyStock AS AlertStock, '个' AS UnitName " & _
    "FROM ProductInventory pi LEFT JOIN Products p ON pi.ProductID=p.ProductID WHERE pi.SafetyStock>0 AND pi.StockQty<=pi.SafetyStock " & _
    "UNION ALL " & _
    "SELECT TOP 10 '瓶子' AS Category, BottleName AS ItemName, StockQty, SafetyStock AS AlertStock, '个' AS UnitName " & _
    "FROM BottleInventory WHERE SafetyStock>0 AND StockQty<=SafetyStock " & _
    "UNION ALL " & _
    "SELECT TOP 10 '包装物' AS Category, ItemName, StockQty, SafetyStock AS AlertStock, '个' AS UnitName " & _
    "FROM PackagingInventory WHERE SafetyStock>0 AND StockQty<=SafetyStock " & _
    "UNION ALL " & _
    "SELECT TOP 10 '原料' AS Category, ItemName, StockQty, SafetyStock AS AlertStock, Unit AS UnitName " & _
    "FROM RawMaterialInventory WHERE SafetyStock>0 AND StockQty<=SafetyStock " & _
    "UNION ALL " & _
    "SELECT TOP 10 '香调' AS Category, fn.NoteName AS ItemName, ni.StockQuantity AS StockQty, ni.MinStockLevel AS AlertStock, 'g' AS UnitName " & _
    "FROM NoteInventory ni INNER JOIN FragranceNotes fn ON ni.NoteID=fn.NoteID WHERE ni.MinStockLevel>0 AND ni.StockQuantity<=ni.MinStockLevel " & _
    ") AS alerts_tmp ORDER BY Category, CAST(StockQty AS FLOAT)/NULLIF(CAST(AlertStock AS FLOAT),0) ASC")
If Err.Number <> 0 Then
    Err.Clear
    Set rsAllAlerts = Nothing
End If
On Error GoTo 0

' ========== 近期入库统计 ==========
Dim todayIn, weekIn, monthIn
todayIn = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE WarehouseInAt >= CAST(GETDATE() AS DATE) AND Status='WarehouseIn'"))
weekIn = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE WarehouseInAt >= DATEADD(day,-7,GETDATE()) AND Status='WarehouseIn'"))
monthIn = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE WarehouseInAt >= DATEADD(month,-1,GETDATE()) AND Status='WarehouseIn'"))
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>库存仪表盘 - 库存管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --input-bg: #2d2d44; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #1a1a2e; color: var(--text); min-height: 100vh; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-header { margin-bottom: 25px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #00BCD4; }
        .page-subtitle { font-size: 13px; color: #888; margin-top: 5px; }
        
        .section-title { font-size: 16px; color: #e0e0e0; margin: 30px 0 15px; display: flex; align-items: center; gap: 8px; padding-bottom: 10px; border-bottom: 1px solid rgba(255,255,255,0.06); }
        .section-title i { font-size: 14px; }
        
        .stats-summary { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }
        .summary-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 22px; border-radius: 14px; border: 1px solid rgba(255,255,255,0.06); text-align: center; }
        .summary-card .num { font-size: 32px; font-weight: 700; display: block; }
        .summary-card .label { font-size: 12px; color: #888; margin-top: 6px; }
        .summary-card.total .num { color: #00BCD4; }
        .summary-card.items .num { color: #4CAF50; }
        .summary-card.alerts .num { color: #f44336; }
        .summary-card.value .num { color: #FF9800; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 20px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 20px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.06); display: flex; align-items: center; gap: 15px; }
        .stat-icon { width: 50px; height: 50px; border-radius: 12px; display: flex; align-items: center; justify-content: center; font-size: 22px; flex-shrink: 0; }
        .stat-icon.prod { background: rgba(76,175,80,0.15); color: #4CAF50; }
        .stat-icon.bottle { background: rgba(33,150,243,0.15); color: #2196F3; }
        .stat-icon.pack { background: rgba(255,152,0,0.15); color: #FF9800; }
        .stat-icon.raw { background: rgba(156,39,176,0.15); color: #9C27B0; }
        .stat-icon.note { background: rgba(0,188,212,0.15); color: #00BCD4; }
        .stat-icon.bn { background: rgba(233,30,99,0.15); color: #E91E63; }
        .stat-info .num { font-size: 22px; font-weight: bold; }
        .stat-info .detail { font-size: 11px; color: #888; margin-top: 2px; }
        .stat-info .detail span { font-weight: 600; }
        
        .alert-badge { display: inline-block; background: rgba(244,67,54,0.15); color: #e57373; padding: 2px 8px; border-radius: 10px; font-size: 11px; margin-top: 4px; }
        .ok-badge { display: inline-block; background: rgba(76,175,80,0.12); color: #81c784; padding: 2px 8px; border-radius: 10px; font-size: 11px; margin-top: 4px; }
        
        .info-cards { display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; margin-bottom: 25px; }
        .info-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; border: 1px solid rgba(255,255,255,0.06); overflow: hidden; }
        .info-card-header { padding: 14px 20px; font-weight: 600; font-size: 15px; color: #e0e0e0; border-bottom: 1px solid rgba(255,255,255,0.06); display: flex; align-items: center; gap: 8px; }
        .info-card-header.alert { background: rgba(244,67,54,0.08); }
        .info-card-header.info { background: rgba(0,188,212,0.08); }
        .info-card-body { padding: 16px 20px; }
        
        table { width: 100%; border-collapse: collapse; }
        th { padding: 10px 12px; text-align: left; font-weight: 600; font-size: 12px; color: #888; border-bottom: 1px solid rgba(255,255,255,0.04); }
        td { padding: 10px 12px; border-bottom: 1px solid rgba(255,255,255,0.03); color: #e0e0e0; font-size: 13px; }
        tr:hover td { background: rgba(255,255,255,0.03); }
        tr.critical td { background: rgba(244,67,54,0.06); }
        
        .urgency-badge { display: inline-block; padding: 2px 10px; border-radius: 10px; font-size: 11px; font-weight: 500; }
        .urgency-high { background: rgba(244,67,54,0.2); color: #e57373; }
        .urgency-medium { background: rgba(255,152,0,0.2); color: #ffb74d; }
        .urgency-low { background: rgba(76,175,80,0.2); color: #81c784; }
        
        .quick-links { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin-top: 10px; }
        .quick-link { display: flex; align-items: center; gap: 8px; padding: 12px 16px; background: rgba(255,255,255,0.04); border-radius: 8px; color: #b0b0b0; text-decoration: none; font-size: 13px; transition: all 0.2s; }
        .quick-link:hover { background: rgba(0,188,212,0.1); color: #00BCD4; }
        .quick-link i { font-size: 14px; width: 18px; text-align: center; }
        
        .text-center { text-align: center; }
        .text-muted { color: #888; font-size: 13px; }
        .text-right { text-align: right; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-warehouse"></i> 库存仪表盘</h2>
            <p class="page-subtitle">全品类库存统一监控 · <%=FormatDateTime(Now(), 1)%></p>
        </div>
        
        <!-- 总体指标 -->
        <div class="stats-summary">
            <div class="summary-card total">
                <span class="num"><%=totalCategories%></span>
                <span class="label">库存品类</span>
            </div>
            <div class="summary-card items">
                <span class="num"><%=totalItems%></span>
                <span class="label">库存条目总数</span>
            </div>
            <div class="summary-card alerts">
                <span class="num"><%=totalLowStock%></span>
                <span class="label">低库存预警</span>
            </div>
            <div class="summary-card value">
                <span class="num">¥<%=FormatNumber(totalValue,0)%></span>
                <span class="label">库存总值</span>
            </div>
        </div>
        
        <!-- 各品类卡片 -->
        <div class="section-title"><i class="fas fa-cubes" style="color:#00BCD4;"></i> 各品类库存概览</div>
        <div class="stats-grid">
            <!-- 成品 -->
            <div class="stat-card">
                <div class="stat-icon prod"><i class="fas fa-box"></i></div>
                <div class="stat-info">
                    <div class="num" style="color:#4CAF50;"><%=prodTotalQty%></div>
                    <div class="detail">成品库存 | <span><%=prodTotal%></span> 种</div>
                    <% If prodLowStock > 0 Then %><span class="alert-badge"><%=prodLowStock%> 低库存</span><% Else %><span class="ok-badge">正常</span><% End If %>
                </div>
            </div>
            <!-- 瓶子 -->
            <div class="stat-card">
                <div class="stat-icon bottle"><i class="fas fa-flask"></i></div>
                <div class="stat-info">
                    <div class="num" style="color:#2196F3;"><%=bottleTotalQty%></div>
                    <div class="detail">瓶子库存 | <span><%=bottleTotal%></span> 种</div>
                    <% If bottleLowStock > 0 Then %><span class="alert-badge"><%=bottleLowStock%> 低库存</span><% Else %><span class="ok-badge">正常</span><% End If %>
                </div>
            </div>
            <!-- 包装物 -->
            <div class="stat-card">
                <div class="stat-icon pack"><i class="fas fa-box-open"></i></div>
                <div class="stat-info">
                    <div class="num" style="color:#FF9800;"><%=packTotalQty%></div>
                    <div class="detail">包装物库存 | <span><%=packTotal%></span> 种</div>
                    <% If packLowStock > 0 Then %><span class="alert-badge"><%=packLowStock%> 低库存</span><% Else %><span class="ok-badge">正常</span><% End If %>
                </div>
            </div>
            <!-- 原料 -->
            <div class="stat-card">
                <div class="stat-icon raw"><i class="fas fa-boxes"></i></div>
                <div class="stat-info">
                    <div class="num" style="color:#9C27B0;"><%=rawTotalQty%></div>
                    <div class="detail">原料库存 | <span><%=rawTotal%></span> 种</div>
                    <% If rawLowStock > 0 Then %><span class="alert-badge"><%=rawLowStock%> 低库存</span><% Else %><span class="ok-badge">正常</span><% End If %>
                </div>
            </div>
            <!-- 香调 -->
            <div class="stat-card">
                <div class="stat-icon note"><i class="fas fa-layer-group"></i></div>
                <div class="stat-info">
                    <div class="num" style="color:#00BCD4;"><%=noteTotalQty%></div>
                    <div class="detail">香调库存 | <span><%=noteTotal%></span> 种</div>
                    <% If noteLowStock > 0 Then %><span class="alert-badge"><%=noteLowStock%> 低库存</span><% Else %><span class="ok-badge">正常</span><% End If %>
                </div>
            </div>
            <!-- 基香 -->
            <div class="stat-card">
                <div class="stat-icon bn"><i class="fas fa-database"></i></div>
                <div class="stat-info">
                    <div class="num" style="color:#E91E63;"><%=bnActive%></div>
                    <div class="detail">激活基香 | 共 <span><%=bnTotal%></span> 种</div>
                    <span class="ok-badge">已记录</span>
                </div>
            </div>
        </div>
        
        <!-- 详细面板 -->
        <div class="info-cards">
            <!-- 全品类低库存预警 -->
            <div class="info-card">
                <div class="info-card-header alert"><i class="fas fa-exclamation-triangle"></i> 全品类低库存预警 TOP 10</div>
                <div class="info-card-body">
                    <table>
                        <thead><tr><th>品类</th><th>名称</th><th>当前库存</th><th>安全库存</th><th>占比</th></tr></thead>
                        <tbody>
                        <%
                        Dim alertRow : alertRow = 0
                        If Not rsAllAlerts Is Nothing Then
                            Do While Not rsAllAlerts.EOF
                                alertRow = alertRow + 1
                                Dim aQty : aQty = SafeNum(rsAllAlerts("StockQty"))
                                Dim aSafety : aSafety = SafeNum(rsAllAlerts("AlertStock"))
                                Dim aPct : aPct = 0 : If aSafety > 0 Then aPct = (aQty / aSafety) * 100
                                Dim aUrgency
                                If aQty <= 0 Then
                                    aUrgency = "high"
                                ElseIf aPct < 50 Then
                                    aUrgency = "medium"
                                Else
                                    aUrgency = "low"
                                End If
                        %>
                            <tr class="<%=IIF(aUrgency="high","critical","")%>">
                                <td><span class="urgency-badge <%=IIF(aUrgency="high","urgency-high",IIF(aUrgency="medium","urgency-medium","urgency-low"))%>"><%=rsAllAlerts("Category")%></span></td>
                                <td><strong><%=Server.HTMLEncode(rsAllAlerts("ItemName") & "")%></strong></td>
                                <td style="color:<%=IIF(aQty<=0,"#f44336","#FF9800")%>;"><%=FormatNumber(aQty,1)%> <%=rsAllAlerts("UnitName")%></td>
                                <td><%=FormatNumber(aSafety,1)%> <%=rsAllAlerts("UnitName")%></td>
                                <td>
                                    <div style="display:flex;align-items:center;gap:6px;">
                                        <div style="flex:1;height:4px;background:rgba(255,255,255,0.1);border-radius:2px;">
                                            <div style="height:100%;border-radius:2px;width:<%=IIF(aPct>100,100,aPct)%>%;background:<%=IIF(aUrgency="high","#f44336",IIF(aUrgency="medium","#FF9800","#FFC107"))%>;"></div>
                                        </div>
                                        <span style="font-size:11px;color:#888;"><%=FormatNumber(aPct,0)%>%</span>
                                    </div>
                                </td>
                            </tr>
                        <%
                                rsAllAlerts.MoveNext
                            Loop
                            rsAllAlerts.Close
                        End If
                        Set rsAllAlerts = Nothing
                        If alertRow = 0 Then
                        %>
                            <tr><td colspan="5" class="text-center text-muted" style="padding:20px;color:#81c784;">✓ 所有品类库存正常</td></tr>
                        <% End If %>
                        </tbody>
                    </table>
                </div>
            </div>
            
            <!-- 入库统计 & 快捷入口 -->
            <div class="info-card">
                <div class="info-card-header info"><i class="fas fa-chart-bar"></i> 入库统计 & 快捷入口</div>
                <div class="info-card-body">
                    <div class="stats-grid" style="grid-template-columns:repeat(3,1fr); gap:10px; margin-bottom:15px;">
                        <div class="stat-card" style="flex-direction:column;align-items:center;text-align:center;gap:5px;padding:14px;">
                            <div class="num" style="font-size:24px;color:#4CAF50;"><%=todayIn%></div>
                            <div class="detail" style="font-size:11px;">今日入库</div>
                        </div>
                        <div class="stat-card" style="flex-direction:column;align-items:center;text-align:center;gap:5px;padding:14px;">
                            <div class="num" style="font-size:24px;color:#2196F3;"><%=weekIn%></div>
                            <div class="detail" style="font-size:11px;">近7天入库</div>
                        </div>
                        <div class="stat-card" style="flex-direction:column;align-items:center;text-align:center;gap:5px;padding:14px;">
                            <div class="num" style="font-size:24px;color:#FF9800;"><%=monthIn%></div>
                            <div class="detail" style="font-size:11px;">本月入库</div>
                        </div>
                    </div>
                    
                    <div class="section-title" style="margin-top:10px;margin-bottom:10px;font-size:14px;">
                        <i class="fas fa-link" style="color:#00BCD4;"></i> 快捷入口
                    </div>
                    <div class="quick-links">
                        <a href="/admin/prodcenter/product_inventory.asp?from=inventory" class="quick-link">
                            <i class="fas fa-box" style="color:#4CAF50;"></i> 成品库存
                        </a>
                        <a href="/admin/prodcenter/bottle_inventory.asp?from=inventory" class="quick-link">
                            <i class="fas fa-flask" style="color:#2196F3;"></i> 瓶子库存
                        </a>
                        <a href="/admin/prodcenter/packaging_inventory.asp?from=inventory" class="quick-link">
                            <i class="fas fa-box-open" style="color:#FF9800;"></i> 包装物库存
                        </a>
                        <a href="/admin/semifinished/raw_material_inventory.asp?from=inventory" class="quick-link">
                            <i class="fas fa-boxes" style="color:#9C27B0;"></i> 原料库存
                        </a>
                        <a href="/admin/semifinished/base_note_inventory.asp?from=inventory" class="quick-link">
                            <i class="fas fa-database" style="color:#E91E63;"></i> 基香库存
                        </a>
                        <a href="/admin/semifinished/note_inventory.asp?from=inventory" class="quick-link">
                            <i class="fas fa-layer-group" style="color:#00BCD4;"></i> 香调库存
                        </a>
                        <a href="/admin/inventory/stock_movements.asp" class="quick-link">
                            <i class="fas fa-history" style="color:#607D8B;"></i> 库存流水
                        </a>
                        <a href="/admin/inventory/inventory_alerts.asp" class="quick-link">
                            <i class="fas fa-exclamation-triangle" style="color:#f44336;"></i> 库存预警
                        </a>
                        <a href="/admin/prodcenter/prod_warehouse.asp?from=inventory" class="quick-link">
                            <i class="fas fa-warehouse" style="color:#8BC34A;"></i> 成品入库
                        </a>
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
