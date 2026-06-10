<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<%
Call OpenConnection()

' 自动创建 OrderCostAllocation 表（如果不存在）
On Error Resume Next
conn.Execute "SELECT TOP 1 1 FROM OrderCostAllocation"
If Err.Number <> 0 Then Err.Clear : conn.Execute "CREATE TABLE OrderCostAllocation (AllocationID INT IDENTITY(1,1) PRIMARY KEY, OrderID INT, OrderNo NVARCHAR(100), CostType NVARCHAR(30), ItemCode NVARCHAR(50), ItemName NVARCHAR(200), UnitCost DECIMAL(19,4) DEFAULT 0, Quantity FLOAT DEFAULT 0, TotalCost DECIMAL(19,4) DEFAULT 0, BatchID INT, AllocatedAt DATETIME DEFAULT GETDATE(), CreatedAt DATETIME DEFAULT GETDATE())"
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Function SafeSQL(str)
    If IsNull(str) Or str = "" Then SafeSQL = "" Else SafeSQL = Replace(str, "'", "''")
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

Function GetCostTypeLabel(ct)
    Select Case ct
        Case "RawMaterial"  : GetCostTypeLabel = "原料成本"
        Case "Packaging"    : GetCostTypeLabel = "包装成本"
        Case "Bottle"       : GetCostTypeLabel = "瓶子成本"
        Case "Printing"     : GetCostTypeLabel = "印刷品成本"
        Case "SprayHead"    : GetCostTypeLabel = "喷头成本"
        Case "Product"      : GetCostTypeLabel = "产品成本"
        Case Else           : GetCostTypeLabel = ct
    End Select
End Function

' ========== 搜索与分页 ==========
Dim searchNo, searchSQL
searchNo = Trim(Request.QueryString("search"))
Dim whereClause : whereClause = ""
If searchNo <> "" Then
    whereClause = " WHERE oca.OrderNo LIKE '%" & SafeSQL(searchNo) & "%'"
End If

' ========== 汇总统计 ==========
Dim totalOrders, totalAllocated, totalCostSum
totalOrders = SafeNum(GetScalar("SELECT COUNT(DISTINCT OrderNo) FROM OrderCostAllocation"))
totalAllocated = SafeNum(GetScalar("SELECT COUNT(*) FROM OrderCostAllocation"))
totalCostSum = SafeNum(GetScalar("SELECT ISNULL(SUM(TotalCost),0) FROM OrderCostAllocation"))

' ========== 成本汇总列表（按OrderNo分组）==========
Dim sqlSummary
sqlSummary = "SELECT oca.OrderNo, " & _
    "ISNULL(SUM(CASE WHEN oca.CostType='RawMaterial' THEN oca.TotalCost END),0) AS RawCost, " & _
    "ISNULL(SUM(CASE WHEN oca.CostType='Bottle' THEN oca.TotalCost END),0) AS BottleCost, " & _
    "ISNULL(SUM(CASE WHEN oca.CostType='Packaging' THEN oca.TotalCost END),0) AS PackCost, " & _
    "ISNULL(SUM(CASE WHEN oca.CostType='Printing' THEN oca.TotalCost END),0) AS PrintCost, " & _
    "ISNULL(SUM(CASE WHEN oca.CostType='SprayHead' THEN oca.TotalCost END),0) AS SprayCost, " & _
    "ISNULL(SUM(CASE WHEN oca.CostType='Product' THEN oca.TotalCost END),0) AS ProductCost, " & _
    "ISNULL(SUM(oca.TotalCost),0) AS TotalOrderCost, " & _
    "COUNT(DISTINCT oca.CostType) AS CostTypeCount, " & _
    "MAX(oca.AllocatedAt) AS LastAllocated " & _
    "FROM OrderCostAllocation oca " & whereClause & _
    " GROUP BY oca.OrderNo ORDER BY MAX(oca.AllocatedAt) DESC"

Dim rsSummary
Set rsSummary = conn.Execute(sqlSummary)

' ========== 获取品类成本类型颜色 ==========
Function GetTypeColor(ct)
    Select Case ct
        Case "RawMaterial" : GetTypeColor = "#FF9800"
        Case "Packaging"   : GetTypeColor = "#2196F3"
        Case "Bottle"      : GetTypeColor = "#9C27B0"
        Case "Printing"    : GetTypeColor = "#00BCD4"
        Case "SprayHead"   : GetTypeColor = "#E91E63"
        Case "Product"     : GetTypeColor = "#4CAF50"
        Case Else          : GetTypeColor = "#888"
    End Select
End Function
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>订单成本汇总 - 采购中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --accent: #FF9800; --input-bg: #2d2d44; --card-bg: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: var(--bg); color: var(--text); min-height: 100vh; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-header { margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px; }
        .page-title { font-size: 24px; color: var(--text); display: flex; align-items: center; gap: 10px; }
        .page-title i { color: var(--accent); }
        .breadcrumb { font-size: 13px; color: #888; margin-bottom: 5px; }
        .breadcrumb a { color: var(--accent); text-decoration: none; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: var(--card-bg); padding: 20px; border-radius: 12px; text-align: center; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .num { font-size: 28px; font-weight: bold; display: block; }
        .stat-card .label { font-size: 12px; color: #888; display: block; margin-top: 5px; }
        
        .card { background: var(--card-bg); border-radius: 12px; margin-bottom: 25px; border: 1px solid rgba(255,255,255,0.06); }
        .card-header { padding: 14px 20px; background: rgba(255,152,0,0.08); border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 15px; color: var(--text); display: flex; justify-content: space-between; align-items: center; }
        .card-body { padding: 16px 20px; overflow-x: auto; }
        
        table { width: 100%; border-collapse: collapse; font-size: 14px; }
        th { background: rgba(255,255,255,0.03); padding: 12px 14px; text-align: left; border-bottom: 2px solid rgba(255,255,255,0.08); color: #aaa; font-weight: 600; white-space: nowrap; }
        td { padding: 12px 14px; border-bottom: 1px solid rgba(255,255,255,0.04); }
        tr:hover { background: rgba(255,152,0,0.03); }
        
        .text-muted { color: #888; }
        .text-center { text-align: center; }
        .text-right { text-align: right; }
        .cost-cell { font-family: 'Consolas', monospace; font-size: 13px; }
        .cost-zero { color: #555; }
        .cost-bar { display: inline-block; height: 6px; border-radius: 3px; vertical-align: middle; margin-right: 4px; }
        
        .search-box { display: flex; gap: 10px; }
        .search-box input { background: var(--input-bg); border: 1px solid rgba(255,255,255,0.1); color: var(--text); padding: 8px 14px; border-radius: 8px; font-size: 14px; width: 300px; }
        .search-box input:focus { border-color: var(--accent); outline: none; }

        /* 成本明细展开 */
        .order-row { cursor: pointer; }
        .order-row.active { background: rgba(255,152,0,0.08); }
        .detail-panel { display: none; background: rgba(0,0,0,0.2); padding: 16px 20px; }
        .detail-panel.show { display: table-row; }
        .detail-panel td { padding: 0; }
        .detail-inner { padding: 16px 20px; }
        .detail-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 10px; }
        .detail-item { background: rgba(255,255,255,0.03); padding: 10px 14px; border-radius: 8px; display: flex; justify-content: space-between; align-items: center; }
        .detail-item .item-name { font-size: 13px; }
        .detail-item .item-cost { font-weight: 600; font-size: 14px; }
        
        .type-badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    <div class="main-content">
        <div class="breadcrumb">
            <a href="index.asp"><i class="fas fa-home"></i> 采购中心</a> / 订单成本汇总
        </div>
        <div class="page-header">
            <div class="page-title"><i class="fas fa-calculator"></i> 订单成本汇总</div>
            <div class="search-box">
                <form method="get" style="display:flex;gap:10px;">
                    <input type="text" name="search" value="<%=Server.HTMLEncode(searchNo)%>" placeholder="搜索订单号...">
                    <button type="submit" class="btn btn-accent"><i class="fas fa-search"></i> 查询</button>
                    <% If searchNo <> "" Then %>
                    <a href="order_cost_summary.asp" class="btn btn-outline">清除</a>
                    <% End If %>
                </form>
            </div>
        </div>
        
        <!-- 统计卡片 -->
        <div class="stats-grid">
            <div class="stat-card">
                <span class="num" style="color:#FF9800;"><%=totalOrders%></span>
                <span class="label">已分摊订单数</span>
            </div>
            <div class="stat-card">
                <span class="num" style="color:#4CAF50;"><%=totalAllocated%></span>
                <span class="label">分摊明细条数</span>
            </div>
            <div class="stat-card">
                <span class="num" style="color:#E91E63;">¥<%=FormatNumber(totalCostSum, 2)%></span>
                <span class="label">累计分摊总额</span>
            </div>
            <div class="stat-card">
                <span class="num" style="color:#2196F3;"><%=SafeNum(GetScalar("SELECT AVG(CostAvg) FROM (SELECT SUM(TotalCost) AS CostAvg FROM OrderCostAllocation GROUP BY OrderNo) t"))%></span>
                <span class="label">平均每单成本(¥)</span>
            </div>
        </div>
        
        <!-- 成本汇总表 -->
        <div class="card">
            <div class="card-header">
                <span><i class="fas fa-list"></i> 订单成本明细</span>
                <span class="text-muted" style="font-weight:400;font-size:13px;">点击订单行展开细节</span>
            </div>
            <div class="card-body">
                <table>
                    <thead>
                        <tr>
                            <th>订单号</th>
                            <th>原料成本</th>
                            <th>瓶子成本</th>
                            <th>包装成本</th>
                            <th>印刷品成本</th>
                            <th>喷头成本</th>
                            <th>产品成本</th>
                            <th>总成本</th>
                            <th>品类数</th>
                            <th>最近分摊</th>
                        </tr>
                    </thead>
                    <tbody>
                    <%
                    Dim rowCount : rowCount = 0
                    If Not rsSummary Is Nothing Then
                        Do While Not rsSummary.EOF
                            rowCount = rowCount + 1
                            Dim orderNo : orderNo = rsSummary("OrderNo") & ""
                            Dim rawC, botC, pkgC, prtC, sprC, prodC, totalC
                            rawC = SafeNum(rsSummary("RawCost"))
                            botC = SafeNum(rsSummary("BottleCost"))
                            pkgC = SafeNum(rsSummary("PackCost"))
                            prtC = SafeNum(rsSummary("PrintCost"))
                            sprC = SafeNum(rsSummary("SprayCost"))
                            prodC = SafeNum(rsSummary("ProductCost"))
                            totalC = SafeNum(rsSummary("TotalOrderCost"))
                    %>
                        <tr class="order-row" onclick="toggleDetail('<%=Server.HTMLEncode(orderNo)%>', this)">
                            <td><strong style="color:#FF9800;"><%=Server.HTMLEncode(orderNo)%></strong></td>
                            <td class="cost-cell <%=IIF(rawC=0,"cost-zero","")%>">¥<%=FormatNumber(rawC,4)%></td>
                            <td class="cost-cell <%=IIF(botC=0,"cost-zero","")%>">¥<%=FormatNumber(botC,4)%></td>
                            <td class="cost-cell <%=IIF(pkgC=0,"cost-zero","")%>">¥<%=FormatNumber(pkgC,4)%></td>
                            <td class="cost-cell <%=IIF(prtC=0,"cost-zero","")%>">¥<%=FormatNumber(prtC,4)%></td>
                            <td class="cost-cell <%=IIF(sprC=0,"cost-zero","")%>">¥<%=FormatNumber(sprC,4)%></td>
                            <td class="cost-cell <%=IIF(prodC=0,"cost-zero","")%>">¥<%=FormatNumber(prodC,4)%></td>
                            <td class="cost-cell" style="font-weight:700;color:#FF9800;">¥<%=FormatNumber(totalC,4)%></td>
                            <td class="text-center"><%=rsSummary("CostTypeCount")%></td>
                            <td class="text-muted"><%=IIF(IsNull(rsSummary("LastAllocated")) Or rsSummary("LastAllocated")="","-",Left(rsSummary("LastAllocated"),10))%></td>
                        </tr>
                        <tr class="detail-panel" id="detail_<%=Server.HTMLEncode(orderNo)%>">
                            <td colspan="10">
                                <div class="detail-inner" id="detail_content_<%=Server.HTMLEncode(orderNo)%>">
                                    <div style="text-align:center;color:#888;padding:10px;">加载中...</div>
                                </div>
                            </td>
                        </tr>
                    <%
                            rsSummary.MoveNext
                        Loop
                        rsSummary.Close
                    End If
                    Set rsSummary = Nothing
                    If rowCount = 0 Then
                    %>
                        <tr><td colspan="10" class="text-center text-muted" style="padding:40px;">
                            <% If searchNo <> "" Then %>
                            未找到订单 "<%=Server.HTMLEncode(searchNo)%>" 的成本记录
                            <% Else %>
                            暂无成本分摊数据，收货入库后自动生成
                            <% End If %>
                        </td></tr>
                    <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    
    <script>
    var detailCache = {};
    
    function toggleDetail(orderNo, row) {
        var panel = document.getElementById('detail_' + orderNo);
        if (!panel) return;
        
        if (panel.classList.contains('show')) {
            panel.classList.remove('show');
            row.classList.remove('active');
        } else {
            panel.classList.add('show');
            row.classList.add('active');
            if (!detailCache[orderNo]) {
                loadDetail(orderNo);
            }
        }
    }
    
    function loadDetail(orderNo) {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', 'order_cost_detail.asp?orderno=' + encodeURIComponent(orderNo), true);
        xhr.onload = function() {
            if (xhr.status === 200) {
                detailCache[orderNo] = xhr.responseText;
                var el = document.getElementById('detail_content_' + orderNo);
                if (el) el.innerHTML = xhr.responseText;
            }
        };
        xhr.onerror = function() {
            var el = document.getElementById('detail_content_' + orderNo);
            if (el) el.innerHTML = '<div style="text-align:center;color:#e74c3c;padding:10px;">加载失败</div>';
        };
        xhr.send();
    }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
