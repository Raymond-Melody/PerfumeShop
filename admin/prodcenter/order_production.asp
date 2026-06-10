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

' 确保 OrderItems 表存在（V8新增表）
On Error Resume Next
conn.Execute "IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='OrderItems') CREATE TABLE OrderItems (OrderItemID INT IDENTITY(1,1) PRIMARY KEY, OrderID INT NOT NULL, ProductID INT NULL, Quantity INT DEFAULT 1, UnitPrice DECIMAL(10,2) DEFAULT 0, CreatedAt DATETIME DEFAULT GETDATE())"
Err.Clear
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
            If Not rs.EOF Then
                val = rs(0)
                rs.Close
            End If
        End If
    Else
        Err.Clear
    End If
    Set rs = Nothing
    GetScalar = val
End Function

Dim oSearch : oSearch = Replace(Request.QueryString("keyword"),"'","''")
Dim oDateFrom : oDateFrom = Trim(Request.QueryString("date_from"))
Dim oDateTo : oDateTo = Trim(Request.QueryString("date_to"))
Dim oFilter : oFilter = Request.QueryString("status")
Dim whereClause : whereClause = "1=1"
If oSearch <> "" Then whereClause = whereClause & " AND (o.OrderNo LIKE '%" & oSearch & "%' OR o.ShippingName LIKE '%" & oSearch & "%' OR EXISTS (SELECT 1 FROM OrderItems oi LEFT JOIN Products p ON oi.ProductID=p.ProductID WHERE oi.OrderID=o.OrderID AND p.ProductName LIKE '%" & oSearch & "%'))"
If oDateFrom <> "" Then whereClause = whereClause & " AND CAST(o.CreatedAt AS DATE) >= '" & oDateFrom & "'"
If oDateTo <> "" Then whereClause = whereClause & " AND CAST(o.CreatedAt AS DATE) <= '" & oDateTo & "'"

' 统计
Dim opTotal, opWithPO, opNoPO, opCompleted
opTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders o WHERE Status IN ('Paid','Processing','Shipped')"))
opWithPO = SafeNum(GetScalar("SELECT COUNT(DISTINCT o.OrderID) FROM Orders o INNER JOIN ProductionOrders p ON o.OrderID=p.OrderID WHERE o.Status IN ('Paid','Processing','Shipped')"))
opNoPO = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders o WHERE o.Status IN ('Paid','Processing') AND NOT EXISTS (SELECT 1 FROM ProductionOrders p WHERE p.OrderID=o.OrderID)"))
opCompleted = SafeNum(GetScalar("SELECT COUNT(DISTINCT o.OrderID) FROM Orders o INNER JOIN ProductionOrders p ON o.OrderID=p.OrderID WHERE p.Status='Completed'"))

Dim rsOrders
Set rsOrders = conn.Execute("SELECT o.OrderID, o.OrderNo, o.ShippingName, o.TotalAmount, o.Status AS OrderStatus, o.CreatedAt, " & _
    "(SELECT TOP 1 p.Status FROM ProductionOrders p WHERE p.OrderID=o.OrderID ORDER BY p.CreatedAt DESC) AS LatestPOStatus, " & _
    "(SELECT COUNT(*) FROM ProductionOrders p WHERE p.OrderID=o.OrderID) AS POCount, " & _
    "(SELECT TOP 1 p.ProductionID FROM ProductionOrders p WHERE p.OrderID=o.OrderID ORDER BY p.CreatedAt DESC) AS LatestPOID " & _
    "FROM Orders o WHERE " & whereClause & " AND o.Status IN ('Paid','Processing','Shipped') ORDER BY o.CreatedAt DESC")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>订单级生产追踪 - 产品生产中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --accent: #4CAF50; --input-bg: #2d2d44; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #1a1a2e; color: var(--text); min-height: 100vh; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-header { margin-bottom: 25px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px; }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: var(--accent); }
        .search-box { display: flex; gap: 8px; }
        .search-box input { padding: 8px 14px; background: var(--input-bg); border: 1px solid rgba(255,255,255,0.12); border-radius: 6px; color: #e0e0e0; font-size: 13px; width: 220px; }
        .search-box input:focus { outline: none; border-color: var(--accent); }
        .search-box button { padding: 8px 14px; background: var(--accent); border: none; border-radius: 6px; color: #fff; cursor: pointer; font-size: 13px; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 10px; padding: 16px; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .stat-label { font-size: 11px; color: #888; margin-bottom: 6px; }
        .stat-card .stat-value { font-size: 24px; font-weight: 700; }
        .stat-card .stat-sub { font-size: 11px; color: #666; margin-top: 4px; }
        .data-table { width: 100%; border-collapse: collapse; background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; overflow: hidden; border: 1px solid rgba(255,255,255,0.06); }
        .data-table th { background: rgba(76,175,80,0.12); color: #81c784; font-weight: 600; padding: 12px 14px; text-align: left; font-size: 13px; }
        .data-table td { padding: 10px 14px; color: #e0e0e0; font-size: 13px; border-bottom: 1px solid rgba(255,255,255,0.04); }
        .data-table tr:hover td { background: rgba(255,255,255,0.03); }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 10px; font-size: 11px; font-weight: 500; }
        .badge-paid { background: rgba(33,150,243,0.12); color: #64b5f6; }
        .badge-processing { background: rgba(255,152,0,0.12); color: #ffb74d; }
        .badge-shipped { background: rgba(156,39,176,0.12); color: #ba68c8; }
        .badge-no-po { background: rgba(158,158,158,0.15); color: #9e9e9e; }
        .badge-pending { background: rgba(255,152,0,0.12); color: #ffb74d; }
        .badge-progress { background: rgba(33,150,243,0.12); color: #64b5f6; }
        .badge-completed { background: rgba(76,175,80,0.12); color: #81c784; }
        .progress-timeline { display: flex; gap: 2px; align-items: center; }
        .timeline-dot { width: 12px; height: 12px; border-radius: 50%; background: rgba(255,255,255,0.15); flex-shrink: 0; }
        .timeline-dot.done { background: #4CAF50; }
        .timeline-dot.active { background: #2196F3; box-shadow: 0 0 6px rgba(33,150,243,0.5); }
        .timeline-line { flex: 1; height: 2px; background: rgba(255,255,255,0.15); }
        .timeline-line.done { background: #4CAF50; }
        .empty-state { text-align: center; padding: 60px 20px; color: #666; }
        .empty-state i { font-size: 48px; margin-bottom: 15px; color: rgba(255,255,255,0.05); }
</style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-shopping-cart"></i> 订单级生产追踪</h2>
            <form class="search-box" method="get" style="flex-wrap:wrap;">
                <input type="text" name="keyword" placeholder="搜索订单号/收货人/产品..." value="<%=Server.HTMLEncode(oSearch)%>" style="width:200px;">
                <input type="date" name="date_from" value="<%=oDateFrom%>" title="开始日期" style="width:130px;padding:7px 10px;background:var(--input-bg);border:1px solid rgba(255,255,255,0.12);border-radius:6px;color:#e0e0e0;font-size:12px;">
                <input type="date" name="date_to" value="<%=oDateTo%>" title="结束日期" style="width:130px;padding:7px 10px;background:var(--input-bg);border:1px solid rgba(255,255,255,0.12);border-radius:6px;color:#e0e0e0;font-size:12px;">
                <button type="submit"><i class="fas fa-search"></i> 搜索</button>
                <% If oSearch <> "" Or oDateFrom <> "" Or oDateTo <> "" Then %>
                <a href="order_production.asp" class="btn btn-sm btn-outline" style="font-size:12px;"><i class="fas fa-times"></i> 清除</a>
                <% End If %>
            </form>
        </div>

        <!-- 统计 -->
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-shopping-cart"></i> 活跃订单</div>
                <div class="stat-value" style="color:#2196F3;"><%=opTotal%></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-clipboard-list"></i> 已关联工单</div>
                <div class="stat-value" style="color:#4CAF50;"><%=opWithPO%></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-exclamation-circle"></i> 待排产</div>
                <div class="stat-value" style="color:#FF9800;"><%=opNoPO%></div>
                <div class="stat-sub">未关联生产工单</div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-check-circle"></i> 已完成</div>
                <div class="stat-value" style="color:#4CAF50;"><%=opCompleted%></div>
            </div>
        </div>

        <% If Not rsOrders Is Nothing And Not rsOrders.EOF Then %>
        <table class="data-table">
            <thead>
                <tr>
                    <th>订单号</th>
                    <th>收货人</th>
                    <th>金额</th>
                    <th>订单状态</th>
                    <th>生产状态</th>
                    <th>生产进度</th>
                    <th>下单时间</th>
                    <th>操作</th>
                </tr>
            </thead>
            <tbody>
                <%
                    Do While Not rsOrders.EOF
                        Dim oId, oNo, oName, oAmt, oStatus, oCreated, oPOCnt, oPOStatus, oPOId
                        oId = rsOrders("OrderID")
                        oNo = rsOrders("OrderNo") & ""
                        oName = rsOrders("ShippingName") & ""
                        oAmt = SafeNum(rsOrders("TotalAmount"))
                        oStatus = rsOrders("OrderStatus") & ""
                        oCreated = rsOrders("CreatedAt") & ""
                        oPOCnt = SafeNum(rsOrders("POCount"))
                        oPOStatus = rsOrders("LatestPOStatus") & ""
                        oPOId = rsOrders("LatestPOID") & ""
                %>
                <tr>
                    <td><strong><a href="order_production_detail.asp?id=<%=oId%>" style="color:#81c784;text-decoration:none;">#<%=Server.HTMLEncode(oNo)%></a></strong></td>
                    <td><%=Server.HTMLEncode(oName)%></td>
                    <td>¥<%=FormatNumber(oAmt,2)%></td>
                    <td><span class="badge badge-<%=LCase(oStatus)%>"><%=oStatus%></span></td>
                    <td>
                        <% If oPOCnt = 0 Then %>
                        <span class="badge badge-no-po">未排产</span>
                        <% Else
                            Dim prodStatusLabel, prodStatusClass
                            Select Case oPOStatus
                                Case "Pending"    : prodStatusLabel = "待生产" : prodStatusClass = "badge-processing"
                                Case "InProgress" : prodStatusLabel = "生产中" : prodStatusClass = "badge-processing"
                                Case "Completed"  : prodStatusLabel = "待质检" : prodStatusClass = "badge-paid"
                                Case "QC_Passed"  : prodStatusLabel = "已质检" : prodStatusClass = "badge-paid"
                                Case "QC_Fail"    : prodStatusLabel = "质检不合格" : prodStatusClass = "badge-processing"
                                Case "WarehouseIn": prodStatusLabel = "已入库" : prodStatusClass = "badge-shipped"
                                Case "ShippedOut" : prodStatusLabel = "已发货" : prodStatusClass = "badge-shipped"
                                Case Else         : prodStatusLabel = oPOStatus : prodStatusClass = "badge-paid"
                            End Select
                        %>
                        <span class="badge <%=prodStatusClass%>"><%=prodStatusLabel%></span>
                        <% End If %>
                    </td>
                    <td>
                        <% If oPOCnt > 0 Then
                            Dim svgPct : svgPct = 0
                            Select Case oPOStatus
                                Case "Pending"     : svgPct = 20
                                Case "InProgress"  : svgPct = 40
                                Case "Completed"   : svgPct = 60
                                Case "QC_Passed"   : svgPct = 75
                                Case "WarehouseIn" : svgPct = 90
                                Case "ShippedOut"  : svgPct = 100
                                Case Else          : svgPct = 0
                            End Select
                            Dim svgColor : svgColor = "#2196F3"
                            If svgPct >= 100 Then svgColor = "#4CAF50"
                            If oPOStatus = "QC_Fail" Then svgColor = "#f44336" : svgPct = 55
                        %>
                        <div style="display:flex;align-items:center;gap:6px;">
                        <svg width="60" height="6" style="flex-shrink:0;"><rect width="60" height="6" rx="3" fill="rgba(255,255,255,0.08)"/><rect width="<%=svgPct*0.6%>" height="6" rx="3" fill="<%=svgColor%>"/></svg>
                        <span style="font-size:11px;color:<%=svgColor%>;"><%=svgPct%>%</span>
                        </div>
                        <% Else %>
                        <span style="color:#666;">—</span>
                        <% End If %>
                    </td>
                    <td style="color:#888;"><%=IIf(oCreated="","-",Left(oCreated,10))%></td>
                    <td>
                        <% If oPOCnt = 0 Then %>
                        <a href="production_management.asp?orderId=<%=oId%>" class="btn btn-sm btn-outline"><i class="fas fa-plus"></i> 创建工单</a>
                        <% Else %>
                        <a href="order_production_detail.asp?id=<%=oId%>" class="btn btn-sm btn-outline"><i class="fas fa-search-location"></i> 追踪</a>
                        <% End If %>
                    </td>
                </tr>
                <%
                        rsOrders.MoveNext
                    Loop
                    rsOrders.Close : Set rsOrders = Nothing
                %>
            </tbody>
        </table>
        <% Else %>
        <div class="empty-state">
            <div><i class="fas fa-inbox"></i></div>
            <p>暂无匹配的订单数据</p>
        </div>
        <% End If %>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
