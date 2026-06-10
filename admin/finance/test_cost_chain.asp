<%@ Language="VBScript" CodePage="65001" %>
<% Response.Charset = "UTF-8" %>
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<!--#include file="../../includes/cost_engine.asp"-->
<!--#include file="../includes/role_auth.asp"-->
<%
Call OpenConnection()

' 预加载成本数据（避免N+1查询）
Call CE_PreloadAllCostData()

' 权限检查
If Session("AdminRoleCode") <> "FIN_MANAGER" And Session("AdminRoleCode") <> "SUPER_ADMIN" Then
    Response.Redirect "../unauthorized.asp"
End If

Call LogAdminAction("成本传导链验证", "finance", "test_cost_chain", "", "打开验证测试页面")

' ============================================
' 辅助函数
' ============================================
Function FormatPrice(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then
        FormatPrice = "0.00"
    Else
        FormatPrice = FormatNumber(CDbl(val), 4)
    End If
End Function

Function CountRows(sql)
    Dim rs, cnt
    cnt = 0
    Set rs = conn.Execute(sql)
    If Not rs Is Nothing Then
        If Not rs.EOF Then cnt = CLng(rs(0))
        rs.Close
    End If
    Set rs = Nothing
    CountRows = cnt
End Function

Function GetFirstVal(sql)
    Dim rs, val
    val = ""
    Set rs = conn.Execute(sql)
    If Not rs Is Nothing Then
        If Not rs.EOF Then val = rs(0)
        rs.Close
    End If
    Set rs = Nothing
    GetFirstVal = val
End Function
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>成本传导链验证测试 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-title { font-size: 24px; display: flex; align-items: center; gap: 12px; margin-bottom: 25px; }
        .page-title i { color: #00bcd4; }
        .breadcrumb { color: #888; font-size: 14px; margin-bottom: 20px; }
        .breadcrumb a { color: #00bcd4; text-decoration: none; }

        .chain-container { display: flex; flex-direction: column; gap: 20px; }

        .chain-node { 
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px; padding: 20px; border: 1px solid rgba(255,255,255,0.06);
            border-left: 4px solid;
        }
        .chain-node.level1 { border-left-color: #4CAF50; }
        .chain-node.level2 { border-left-color: #2196F3; }
        .chain-node.level3 { border-left-color: #FF9800; }
        .chain-node.level4 { border-left-color: #9C27B0; }
        .chain-node.level5 { border-left-color: #f44336; }

        .node-header { 
            display: flex; align-items: center; gap: 15px; margin-bottom: 15px;
            padding-bottom: 12px; border-bottom: 1px solid rgba(255,255,255,0.06);
        }
        .node-header .level-badge {
            padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 600;
        }
        .level1 .level-badge { background: rgba(76,175,80,0.2); color: #81c784; }
        .level2 .level-badge { background: rgba(33,150,243,0.2); color: #64b5f6; }
        .level3 .level-badge { background: rgba(255,152,0,0.2); color: #ffb74d; }
        .level4 .level-badge { background: rgba(156,39,176,0.2); color: #ce93d8; }
        .level5 .level-badge { background: rgba(244,67,54,0.2); color: #ef9a9a; }
        .node-header h3 { margin: 0; font-size: 16px; color: #e0e0e0; }
        .node-header .stats { margin-left: auto; font-size: 12px; color: #888; }

        .data-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px; margin-bottom: 15px; }
        .data-card { background: #1a1a2e; border-radius: 8px; padding: 12px; text-align: center; }
        .data-card .label { font-size: 11px; color: #888; margin-bottom: 5px; }
        .data-card .value { font-size: 22px; font-weight: 700; }
        .data-card .sub { font-size: 11px; color: #888; margin-top: 4px; }

        .data-table { width: 100%; border-collapse: collapse; font-size: 13px; }
        .data-table th {
            background: #1a1a2e; color: #888; font-weight: 600; padding: 8px 10px;
            text-align: left; border-bottom: 1px solid #3a3a4a;
        }
        .data-table td { padding: 6px 10px; border-bottom: 1px solid rgba(255,255,255,0.04); }
        .data-table tr:hover td { background: rgba(255,255,255,0.03); }
        .data-table .cost { font-family: 'Courier New', monospace; font-weight: 600; }
        .data-table .pass { color: #4CAF50; }
        .data-table .fail { color: #f44336; }
        .data-table .warn { color: #FF9800; }

        .summary-bar {
            display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 10px;
            margin-bottom: 25px;
        }
        .summary-item {
            background: #2d2d44; border-radius: 10px; padding: 15px; text-align: center;
            border: 1px solid rgba(255,255,255,0.06);
        }
        .summary-item .sum-label { font-size: 11px; color: #888; }
        .summary-item .sum-value { font-size: 20px; font-weight: 700; margin: 5px 0; }
        .summary-item .sum-sub { font-size: 11px; }
        .ok { color: #4CAF50; } .err { color: #f44336; } .warn-color { color: #FF9800; }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    <div class="main-content">
        <div class="breadcrumb">
            <a href="index.asp">财务中心</a> / <a href="cost_management.asp">成本管理</a> / <span>成本传导链验证</span>
        </div>

        <h2 class="page-title"><i class="fas fa-flask"></i> 成本自动传导链验证测试</h2>

        <%
        ' ======================================
        ' 总体状态摘要
        ' ======================================
        Dim totalRaw, totalSuppliers, totalNotes, totalProducts, totalOrders
        Dim hasSupplierPrices, hasRecipeAccords, hasNoteIngredients, hasProdNoteRatios
        totalRaw = CountRows("SELECT COUNT(*) FROM RawMaterialInventory")
        totalSuppliers = CountRows("SELECT COUNT(*) FROM SupplierPrices")
        totalNotes = CountRows("SELECT COUNT(*) FROM FragranceNotes WHERE IsActive=1")
        totalProducts = CountRows("SELECT COUNT(*) FROM Products WHERE IsActive=1")
        totalOrders = CountRows("SELECT COUNT(*) FROM Orders")
        hasSupplierPrices = CountRows("SELECT COUNT(*) FROM SupplierPrices WHERE IsActive=1 AND UnitPrice > 0")
        hasRecipeAccords = CountRows("SELECT COUNT(*) FROM RecipeAccords WHERE Status='Published'")
        hasNoteIngredients = CountRows("SELECT COUNT(*) FROM NoteIngredients")
        hasProdNoteRatios = CountRows("SELECT COUNT(*) FROM ProductNoteRatios")
        Dim prodWithCost, ordersWithCost
        prodWithCost = CountRows("SELECT COUNT(*) FROM Products WHERE IsActive=1 AND UnitCost > 0")
        ordersWithCost = CountRows("SELECT COUNT(*) FROM Orders WHERE CostAmount > 0")
        
        ' 预计算链路状态标识
        Dim l1ok, l2ok, l3ok, l4ok, l5ok
        l1ok = (hasSupplierPrices > 0)
        l2ok = (hasRecipeAccords > 0 Or hasNoteIngredients > 0)
        l3ok = (hasProdNoteRatios > 0)
        l4ok = (prodWithCost > 0)
        l5ok = (ordersWithCost > 0)
        
        ' 预计算链路状态样式
        Dim l1bg, l1fg, l1icon, l1txt
        Dim l2bg, l2fg, l2icon, l2txt
        Dim l3bg, l3fg, l3icon, l3txt
        Dim l4bg, l4fg, l4icon, l4txt
        Dim l5bg, l5fg, l5icon, l5txt
        
        If l1ok Then l1bg="rgba(76,175,80,0.15)":l1fg="#81c784":l1icon="check":l1txt="Level1 原料价格" Else l1bg="rgba(244,67,54,0.15)":l1fg="#ef9a9a":l1icon="times":l1txt="Level1 原料价格"
        If l2ok Then l2bg="rgba(76,175,80,0.15)":l2fg="#81c784":l2icon="check":l2txt="Level2 香调成本" Else l2bg="rgba(244,67,54,0.15)":l2fg="#ef9a9a":l2icon="times":l2txt="Level2 香调成本"
        If l3ok Then l3bg="rgba(76,175,80,0.15)":l3fg="#81c784":l3icon="check":l3txt="Level3 产品BOM" Else l3bg="rgba(244,67,54,0.15)":l3fg="#ef9a9a":l3icon="times":l3txt="Level3 产品BOM"
        If l4ok Then l4bg="rgba(76,175,80,0.15)":l4fg="#81c784":l4icon="check":l4txt="Level4 单位成本" Else l4bg="rgba(255,152,0,0.15)":l4fg="#ffb74d":l4icon="exclamation":l4txt="Level4 单位成本"
        If l5ok Then l5bg="rgba(76,175,80,0.15)":l5fg="#81c784":l5icon="check":l5txt="Level5 订单利润" Else l5bg="rgba(255,152,0,0.15)":l5fg="#ffb74d":l5icon="exclamation":l5txt="Level5 订单利润"
        
        ' 预计算 summary 样式
        Dim s4color, s4sub, s5color, s5sub, spColor
        If prodWithCost > 0 Then s4color="#4CAF50":s4sub="ok" Else s4color="#f44336":s4sub="err"
        If ordersWithCost > 0 Then s5color="#4CAF50":s5sub="ok" Else s5color="#f44336":s5sub="err"
        
        Dim pendingOrders
        pendingOrders = CountRows("SELECT COUNT(*) FROM Orders WHERE (CostAmount IS NULL OR CostAmount = 0) AND Status NOT IN ('Pending','Cancelled')")
        If pendingOrders > 0 Then spColor="#FF9800" Else spColor="#4CAF50"
        %>

        <!-- 总览 -->
        <div class="summary-bar">
            <div class="summary-item">
                <div class="sum-label">原料种类</div>
                <div class="sum-value"><%= totalRaw %></div>
                <div class="sum-sub">供应商价格: <%= totalSuppliers %></div>
            </div>
            <div class="summary-item">
                <div class="sum-label">香调总数</div>
                <div class="sum-value"><%= totalNotes %></div>
                <div class="sum-sub">Accord配方: <%= hasRecipeAccords %> | NoteIngredients: <%= hasNoteIngredients %></div>
            </div>
            <div class="summary-item">
                <div class="sum-label">产品总数</div>
                <div class="sum-value"><%= totalProducts %></div>
                                <div class="sub <%= s4sub %>">已更新成本: <%= prodWithCost %></div>
            </div>
            <div class="summary-item">
                <div class="sum-label">订单总数</div>
                <div class="sum-value"><%= totalOrders %></div>
                                <div class="sub <%= s5sub %>">已计算利润: <%= ordersWithCost %></div>
            </div>
        </div>

        <!-- 链路总状态 -->
        <div style="background:#2d2d44;border-radius:12px;padding:20px;margin-bottom:20px;border:1px solid rgba(255,255,255,0.06);">
            <h4 style="color:#e0e0e0;margin:0 0 15px 0;"><i class="fas fa-link"></i> 传导链路状态</h4>
            <div style="display:flex;flex-wrap:wrap;gap:10px;">
                <span style="padding:6px 14px;border-radius:6px;font-size:13px;background:<%= l1bg %>;color:<%= l1fg %>;">
                    <i class="fas fa-<%= l1icon %>-circle"></i> <%= l1txt %>
                </span>
                <span style="padding:6px 14px;border-radius:6px;font-size:13px;background:<%= l2bg %>;color:<%= l2fg %>;">
                    <i class="fas fa-<%= l2icon %>-circle"></i> <%= l2txt %>
                </span>
                <span style="padding:6px 14px;border-radius:6px;font-size:13px;background:<%= l3bg %>;color:<%= l3fg %>;">
                    <i class="fas fa-<%= l3icon %>-circle"></i> <%= l3txt %>
                </span>
                <span style="padding:6px 14px;border-radius:6px;font-size:13px;background:<%= l4bg %>;color:<%= l4fg %>;">
                    <i class="fas fa-<%= l4icon %>-circle"></i> <%= l4txt %>
                </span>
                <span style="padding:6px 14px;border-radius:6px;font-size:13px;background:<%= l5bg %>;color:<%= l5fg %>;">
                    <i class="fas fa-<%= l5icon %>-circle"></i> <%= l5txt %>
                </span>
            </div>
        </div>

        <%
        ' ======================================
        ' Level 1: 原料采购成本
        ' ======================================
        Dim rs1
        %>
        <div class="chain-node level1">
            <div class="node-header">
                <span class="level-badge">Level 1</span>
                <i class="fas fa-truck-loading" style="color:#4CAF50;"></i>
                <h3>采购原料成本</h3>
                <span class="stats">SupplierPrices → RawMaterialInventory.UnitPrice</span>
            </div>
            <div class="data-grid">
                <div class="data-card">
                    <div class="label">原料种类</div>
                    <div class="value" style="color:#4CAF50;"><%= totalRaw %></div>
                    <div class="sub">有库存的原料</div>
                </div>
                <div class="data-card">
                    <div class="label">有采购价的原料</div>
                    <div class="value" style="color:#2196F3;"><%= hasSupplierPrices %></div>
                    <div class="sub">SupplierPrices.IsActive=1</div>
                </div>
                <div class="data-card">
                    <div class="label">原料总库存量</div>
                    <%
                    Dim totalStock
                    totalStock = GetFirstVal("SELECT ISNULL(SUM(StockQty),0) FROM RawMaterialInventory")
                    %>
                    <div class="value" style="color:#FF9800;"><%= FormatNumber(totalStock, 0) %></div>
                    <div class="sub">合计</div>
                </div>
            </div>
            <% If hasSupplierPrices > 0 Then %>
            <table class="data-table">
                <tr><th>原料名称</th><th>ItemCode</th><th>库存单价</th><th>最新采购价</th><th>计算所得成本</th><th>状态</th></tr>
                <%
                Set rs1 = conn.Execute("SELECT TOP 10 MaterialID, ItemName, ItemCode, UnitPrice FROM RawMaterialInventory WHERE StockQty > 0 ORDER BY MaterialID")
                If Not rs1 Is Nothing Then
                    Do While Not rs1.EOF
                        Dim matId, matName, matCode, matPrice, latestPrice, calcCost
                        matId = rs1("MaterialID")
                        matName = rs1("ItemName")
                        matCode = rs1("ItemCode")
                        matPrice = rs1("UnitPrice")
                        calcCost = CE_GetCachedMaterialCost(matId)
                %>
                <tr>
                    <td><%= Server.HTMLEncode(matName) %></td>
                    <td style="color:#888;"><%= Server.HTMLEncode(matCode) %></td>
                    <td class="cost">¥<%= FormatPrice(matPrice) %></td>
                    <td class="cost"><%= FormatPrice(GetFirstVal("SELECT TOP 1 UnitPrice FROM SupplierPrices WHERE ItemCode='" & SafeSQL(matCode) & "' AND IsActive=1 ORDER BY CreatedAt DESC")) %></td>
                    <td class="cost"><%= FormatPrice(calcCost) %></td>
                    <td><span class="<%= IIf(calcCost > 0, "pass", "warn") %>"><%= IIf(calcCost > 0, "OK", "待配置") %></span></td>
                </tr>
                <%
                        rs1.MoveNext
                    Loop
                    rs1.Close
                End If
                Set rs1 = Nothing
                %>
            </table>
            <div style="text-align:right;color:#888;font-size:12px;margin-top:8px;">
                <%= CountRows("SELECT COUNT(*) FROM RawMaterialInventory WHERE StockQty > 0") %> 条记录 (显示前10条)
            </div>
            <% Else %>
            <div style="padding:15px;text-align:center;color:#ffb74d;background:rgba(255,152,0,0.1);border-radius:8px;">
                <i class="fas fa-exclamation-triangle"></i> 暂无采购价格数据。请先在采购管理中为原料添加采购价。
            </div>
            <% End If %>
        </div>

        <%
        ' ======================================
        ' Level 2: 香调成本
        ' ======================================
        %>
        <div class="chain-node level2">
            <div class="node-header">
                <span class="level-badge">Level 2</span>
                <i class="fas fa-flask" style="color:#2196F3;"></i>
                <h3>香调(Note)成本计算</h3>
                <span class="stats">RecipeAccords / NoteIngredients → CE_CalculateNoteCost()</span>
            </div>
            <div class="data-grid">
                <div class="data-card">
                    <div class="label">香调总数</div>
                    <div class="value" style="color:#2196F3;"><%= totalNotes %></div>
                    <div class="sub">FragranceNotes (启用)</div>
                </div>
                <div class="data-card">
                    <div class="label">有Accord配方</div>
                    <div class="value" style="color:#4CAF50;"><%= hasRecipeAccords %></div>
                    <div class="sub">路径A: RecipeAccords (Published)</div>
                </div>
                <div class="data-card">
                    <div class="label">有成分聚合</div>
                    <div class="value" style="color:#FF9800;"><%= hasNoteIngredients %></div>
                    <div class="sub">路径B: NoteIngredients</div>
                </div>
            </div>
            <table class="data-table">
                <tr>
                    <th>NoteID</th><th>香调名称</th><th>类型</th><th>PriceAddition</th>
                    <th>路径A(Accord)</th><th>路径B(Ingredients)</th><th>计算成本</th><th>状态</th>
                </tr>
                <%
                Dim rsNotes, hasAccordFlag, hasIngFlag
                Set rsNotes = conn.Execute("SELECT NoteID, NoteName, NoteType, PriceAddition FROM FragranceNotes WHERE IsActive=1 ORDER BY NoteID")
                If Not rsNotes Is Nothing Then
                    Do While Not rsNotes.EOF
                        Dim nid, nname, ntype, nprice, ncost
                        nid = rsNotes("NoteID")
                        nname = rsNotes("NoteName")
                        ntype = rsNotes("NoteType")
                        nprice = rsNotes("PriceAddition")
                        ncost = CE_GetCachedNoteCost(nid)
                        hasAccordFlag = CountRows("SELECT COUNT(*) FROM RecipeAccords WHERE NoteID=" & nid & " AND Status='Published'")
                        hasIngFlag = CountRows("SELECT COUNT(*) FROM NoteIngredients WHERE NoteID=" & nid)
                %>
                <tr>
                    <td><%= nid %></td>
                    <td><strong><%= Server.HTMLEncode(nname) %></strong></td>
                    <td style="color:#888;"><%= Server.HTMLEncode(ntype) %></td>
                    <td class="cost">¥<%= FormatPrice(nprice) %></td>
                    <td><%= IIf(hasAccordFlag > 0, "<i class='fas fa-check pass'></i>", "<i class='fas fa-times' style='color:#555;'></i>") %></td>
                    <td><%= IIf(hasIngFlag > 0, "<i class='fas fa-check pass'></i>", "<i class='fas fa-times' style='color:#555;'></i>") %></td>
                    <td class="cost">¥<%= FormatPrice(ncost) %></td>
                    <td><span class="<%= IIf(ncost > 0, "pass", "warn") %>"><%= IIf(ncost > 0, "OK", "待核算") %></span></td>
                </tr>
                <%
                        rsNotes.MoveNext
                    Loop
                    rsNotes.Close
                End If
                Set rsNotes = Nothing
                %>
            </table>
            <div style="text-align:right;color:#888;font-size:12px;margin-top:8px;">
                <i class="fas fa-info-circle"></i> 
                路径A优先: RecipeAccords(Accord配方×原料配比), 无Accord则走路径B: NoteIngredients(BaseNote聚合)
            </div>
        </div>

        <%
        ' ======================================
        ' Level 3: 产品BOM成本
        ' ======================================
        %>
        <div class="chain-node level3">
            <div class="node-header">
                <span class="level-badge">Level 3</span>
                <i class="fas fa-box" style="color:#FF9800;"></i>
                <h3>产品BOM成本</h3>
                <span class="stats">ProductNoteRatios → CE_CalculateProductBOMCost()</span>
            </div>
            <div class="data-grid">
                <div class="data-card">
                    <div class="label">产品总数</div>
                    <div class="value" style="color:#FF9800;"><%= totalProducts %></div>
                    <div class="sub">Products (启用)</div>
                </div>
                <div class="data-card">
                    <div class="label">有香调配比</div>
                    <div class="value" style="color:#4CAF50;"><%= hasProdNoteRatios %></div>
                    <div class="sub">ProductNoteRatios</div>
                </div>
                <div class="data-card">
                    <div class="label">已更新成本</div>
                    <div class="value" style="color:<%= s4color %>;"><%= prodWithCost %></div>
                    <div class="sub">UnitCost > 0</div>
                </div>
            </div>
            <table class="data-table">
                <tr>
                    <th>ProductID</th><th>产品名称</th><th>类型</th><th>香调配比</th>
                    <th>BOM成本(自动)</th><th>单位成本(自动)</th><th>状态</th>
                </tr>
                <%
                Dim rsProd
                Set rsProd = conn.Execute("SELECT TOP 15 ProductID, ProductName, ProductType FROM Products WHERE IsActive=1 ORDER BY ProductID")
                If Not rsProd Is Nothing Then
                    Do While Not rsProd.EOF
                        Dim pid2, pname2, ptype2, pbom, punit, hasRatio
                        pid2 = rsProd("ProductID")
                        pname2 = rsProd("ProductName")
                        ptype2 = rsProd("ProductType")
                        pbom = CE_GetCachedProductBOMCost(pid2)
                        punit = CE_GetCachedProductUnitCost(pid2)
                        hasRatio = CountRows("SELECT COUNT(*) FROM ProductNoteRatios WHERE ProductID=" & pid2)
                %>
                <tr>
                    <td><%= pid2 %></td>
                    <td><strong><%= Server.HTMLEncode(pname2) %></strong></td>
                    <td style="color:#888;"><%= Server.HTMLEncode(ptype2) %></td>
                    <td><%= IIf(hasRatio > 0, "<i class='fas fa-check pass'></i> (" & hasRatio & ")", "<i class='fas fa-times' style='color:#555;'></i>") %></td>
                    <td class="cost">¥<%= FormatPrice(pbom) %></td>
                    <td class="cost">¥<%= FormatPrice(punit) %></td>
                    <td><span class="<%= IIf(punit > 0, "pass", IIf(hasRatio > 0, "warn", "fail")) %>">
                        <%= IIf(punit > 0, "已传导", IIf(hasRatio > 0, "香调待算", "无配比")) %></span></td>
                </tr>
                <%
                        rsProd.MoveNext
                    Loop
                    rsProd.Close
                End If
                Set rsProd = Nothing
                %>
            </table>
            <div style="text-align:right;color:#888;font-size:12px;margin-top:8px;">
                显示前15个产品。使用"<a href="cost_management.asp?tab=chain" style="color:#00bcd4;">成本传导链</a>"中的"自动计算全部产品成本"来批量更新。
            </div>
        </div>

        <%
        ' ======================================
        ' Level 4: 产品单位成本明细
        ' ======================================
        %>
        <div class="chain-node level4">
            <div class="node-header">
                <span class="level-badge">Level 4</span>
                <i class="fas fa-cubes" style="color:#9C27B0;"></i>
                <h3>单位成本构成明细</h3>
                <span class="stats">BOM + 包装 + 人工/分摊</span>
            </div>
            <%
            Dim samplePid
            samplePid = GetFirstVal("SELECT TOP 1 ProductID FROM Products WHERE IsActive=1 ORDER BY ProductID")
            If IsNumeric(samplePid) And samplePid > 0 Then
                Dim bom, totalBottleCost, pkg, oth
                bom = CE_GetCachedProductBOMCost(samplePid)
                Set rs1 = conn.Execute("SELECT ISNULL(SUM(TotalCost),0) FROM ProductCosts WHERE ProductID=" & samplePid & " AND CostType='Packaging'")
                If Not rs1 Is Nothing Then pkg = rs1(0) : rs1.Close : Set rs1 = Nothing
                Set rs1 = conn.Execute("SELECT ISNULL(SUM(TotalCost),0) FROM ProductCosts WHERE ProductID=" & samplePid & " AND CostType='Other'")
                If Not rs1 Is Nothing Then oth = rs1(0) : rs1.Close : Set rs1 = Nothing
            %>
            <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;">
                <div class="data-card">
                    <div class="label">BOM成本(香调+瓶身)</div>
                    <div class="value" style="color:#FF9800;">¥<%= FormatPrice(bom) %></div>
                    <div class="sub">占 <%= IIf(bom > 0, FormatPercent(bom/(bom+pkg+oth)), "0%") %></div>
                </div>
                <div class="data-card">
                    <div class="label">包装成本</div>
                    <div class="value" style="color:#2196F3;">¥<%= FormatPrice(pkg) %></div>
                    <div class="sub">ProductCosts.Packaging</div>
                </div>
                <div class="data-card">
                    <div class="label">人工/分摊</div>
                    <div class="value" style="color:#9C27B0;">¥<%= FormatPrice(oth) %></div>
                    <div class="sub">ProductCosts.Other</div>
                </div>
                <div class="data-card">
                    <div class="label">单位总成本</div>
                    <div class="value" style="color:#4CAF50;">¥<%= FormatPrice(CE_GetCachedProductUnitCost(samplePid)) %></div>
                    <div class="sub">ProductID: <%= samplePid %></div>
                </div>
            </div>
            <div style="text-align:right;color:#888;font-size:12px;margin-top:8px;">
                示例: ProductID=<%= samplePid %> 的成本结构
            </div>
            <% Else %>
            <div style="padding:15px;text-align:center;color:#ffb74d;">暂无产品数据</div>
            <% End If %>
        </div>

        <%
        ' ======================================
        ' Level 5: 订单成本与利润
        ' ======================================
        %>
        <div class="chain-node level5">
            <div class="node-header">
                <span class="level-badge">Level 5</span>
                <i class="fas fa-shopping-cart" style="color:#f44336;"></i>
                <h3>订单成本与利润</h3>
                <span class="stats">OrderDetails × UnitCost → Orders.CostAmount/ProfitAmount</span>
            </div>
            <div class="data-grid">
                <div class="data-card">
                    <div class="label">订单总数</div>
                    <div class="value" style="color:#f44336;"><%= totalOrders %></div>
                    <div class="sub">全部订单</div>
                </div>
                <div class="data-card">
                    <div class="label">已计算利润</div>
                    <div class="value" style="color:<%= s5color %>;"><%= ordersWithCost %></div>
                    <div class="sub">CostAmount > 0</div>
                </div>
                <div class="data-card">
                    <div class="label">待计算订单</div>
                    <%
                    pendingOrders = CountRows("SELECT COUNT(*) FROM Orders WHERE (CostAmount IS NULL OR CostAmount = 0) AND Status NOT IN ('Pending','Cancelled')")
                    %>
                    <div class="value" style="color:<%= spColor %>;"><%= pendingOrders %></div>
                    <div class="sub">已确认但未算成本</div>
                </div>
            </div>
            <% If totalOrders > 0 Then %>
            <table class="data-table">
                <tr>
                    <th>OrderID</th><th>订单号</th><th>总金额</th><th>运费</th>
                    <th>计算成本</th><th>计算利润</th><th>状态</th>
                </tr>
                <%
                Dim rsOrders2
                Set rsOrders2 = conn.Execute("SELECT TOP 10 OrderID, OrderNo, TotalAmount, ShippingFee, Status FROM Orders ORDER BY OrderID DESC")
                If Not rsOrders2 Is Nothing Then
                    Do While Not rsOrders2.EOF
                        Dim oid2, ono, oamt, oship, ostatus, ocost, oprofit
                        oid2 = rsOrders2("OrderID")
                        ono = rsOrders2("OrderNo")
                        oamt = rsOrders2("TotalAmount")
                        oship = rsOrders2("ShippingFee")
                        ostatus = rsOrders2("Status")
                        ocost = GetFirstVal("SELECT ISNULL(CostAmount,0) FROM Orders WHERE OrderID=" & oid2)
                        oprofit = GetFirstVal("SELECT ISNULL(ProfitAmount,0) FROM Orders WHERE OrderID=" & oid2)
                %>
                <tr>
                    <td><%= oid2 %></td>
                    <td style="color:#888;"><%= Server.HTMLEncode(ono) %></td>
                    <td class="cost">¥<%= FormatPrice(oamt) %></td>
                    <td class="cost">¥<%= FormatPrice(oship) %></td>
                    <td class="cost">¥<%= FormatPrice(ocost) %></td>
                    <td class="cost">¥<%= FormatPrice(oprofit) %></td>
                    <td><span class="<%= IIf(ostatus="Paid" Or ostatus="Delivered", "pass", "warn") %>"><%= ostatus %></span></td>
                </tr>
                <%
                        rsOrders2.MoveNext
                    Loop
                    rsOrders2.Close
                End If
                Set rsOrders2 = Nothing
                %>
            </table>
            <div style="text-align:right;color:#888;font-size:12px;margin-top:8px;">
                最近10个订单。使用"<a href="cost_management.asp?tab=chain" style="color:#00bcd4;">成本传导链</a>"中的"自动更新全部订单利润"来批量更新。
            </div>
            <% Else %>
            <div style="padding:15px;text-align:center;color:#888;">暂无订单数据</div>
            <% End If %>
        </div>

        <!-- 操作按钮 -->
        <div style="margin-top:25px;display:flex;gap:15px;flex-wrap:wrap;">
            <form method="post" action="cost_management.asp?tab=chain" style="display:inline;">
                <input type="hidden" name="action" value="auto_calc_all">
                <button type="submit" class="btn btn-primary" onclick="return confirm('确认批量更新所有产品成本？')">
                    <i class="fas fa-calculator"></i> 批量更新产品成本
                </button>
            </form>
            <form method="post" action="cost_management.asp?tab=chain" style="display:inline;">
                <input type="hidden" name="action" value="auto_calc_orders">
                <button type="submit" class="btn btn-primary" onclick="return confirm('确认批量更新所有订单利润？')">
                    <i class="fas fa-receipt"></i> 批量更新订单利润
                </button>
            </form>
            <form method="get" action="test_cost_chain.asp" style="display:inline;">
                <button type="submit" class="btn btn-secondary">
                    <i class="fas fa-sync-alt"></i> 刷新当前页面
                </button>
            </form>
        </div>

        <!-- 说明 -->
        <div style="margin-top:25px;background:#1a1a2e;border-radius:12px;padding:20px;border:1px solid rgba(255,255,255,0.06);">
            <h4 style="color:#00bcd4;margin:0 0 12px 0;"><i class="fas fa-info-circle"></i> 成本传导链路说明</h4>
            <ol style="color:#b0b0b0;font-size:13px;line-height:2;">
                <li><strong style="color:#4CAF50;">采购单价</strong> → SupplierPrices.UnitPrice 或 RawMaterialInventory.UnitPrice</li>
                <li><strong style="color:#2196F3;">香调成本</strong> → 路径A: RecipeAccords → RecipeAccordMaterials 原料配比×单价; 路径B: NoteIngredients → BaseNote 聚合计算</li>
                <li><strong style="color:#FF9800;">产品BOM</strong> → ProductNoteRatios 香调配比×Note成本 + BottleStyles.PriceAddition</li>
                <li><strong style="color:#9C27B0;">单位总成本</strong> → BOM + ProductCosts(Packaging + Other) 包装和人工分摊</li>
                <li><strong style="color:#f44336;">订单利润</strong> → Orders.TotalAmount - CostAmount - ShippingFee</li>
            </ol>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>