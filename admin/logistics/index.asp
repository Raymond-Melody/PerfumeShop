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
    Set rs = Nothing : GetScalar = val
End Function

' 确保物流字段存在
On Error Resume Next
conn.Execute "SELECT ShippingStatus FROM Orders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE Orders ADD ShippingStatus NVARCHAR(20) DEFAULT 'Pending'"
conn.Execute "SELECT ShippingCompany FROM Orders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE Orders ADD ShippingCompany NVARCHAR(50)"
conn.Execute "SELECT TrackingNumber FROM Orders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE Orders ADD TrackingNumber NVARCHAR(100)"
conn.Execute "SELECT ShippedAt FROM Orders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE Orders ADD ShippedAt DATETIME"
conn.Execute "SELECT DeliveredAt FROM Orders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE Orders ADD DeliveredAt DATETIME"
conn.Execute "SELECT ShippingNotes FROM Orders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE Orders ADD ShippingNotes NVARCHAR(MAX)"
On Error GoTo 0

Dim lDashIsFull : lDashIsFull = (Session("AdminRoleCode") = "SUPER_ADMIN" Or Session("AdminRoleCode") = "PROD_MANAGER")

' 物流概览统计
Dim lPending, lShipped, lInTransit, lDelivered, lTodayShipped, lTodayDelivered
lPending = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders WHERE Status IN ('Paid','Processing') AND (ShippingStatus='Pending' OR ShippingStatus IS NULL)"))
lShipped = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders WHERE ShippingStatus='Shipped'"))
lInTransit = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders WHERE ShippingStatus='InTransit'"))
lDelivered = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders WHERE ShippingStatus='Delivered'"))
lTodayShipped = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders WHERE ShippingStatus='Shipped' AND ShippedAt >= CAST(GETDATE() AS DATE)"))
lTodayDelivered = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders WHERE ShippingStatus='Delivered' AND DeliveredAt >= CAST(GETDATE() AS DATE)"))

' 待发货订单总金额
Dim lPendingAmt : lPendingAmt = SafeNum(GetScalar("SELECT ISNULL(SUM(TotalAmount),0) FROM Orders WHERE Status IN ('Paid','Processing') AND (ShippingStatus='Pending' OR ShippingStatus IS NULL)"))

' 近期发货
Dim rsRecentShip
Set rsRecentShip = conn.Execute("SELECT TOP 5 OrderID, OrderNo, ShippingName, ShippingCompany, TrackingNumber, ShippingStatus, ShippedAt FROM Orders WHERE ShippedAt IS NOT NULL ORDER BY ShippedAt DESC")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>物流概览 - 物流管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 12px; margin-bottom: 25px; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px; margin-bottom: 30px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; padding: 20px; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .stat-label { font-size: 12px; color: #888; margin-bottom: 8px; }
        .stat-card .stat-value { font-size: 28px; font-weight: 700; }
        .stat-card .stat-sub { font-size: 11px; color: #666; margin-top: 5px; }
        .stat-pending .stat-value { color: #FF9800; }
        .stat-shipped .stat-value { color: #2196F3; }
        .stat-transit .stat-value { color: #9C27B0; }
        .stat-delivered .stat-value { color: #4CAF50; }
        .stat-amount .stat-value { color: #4CAF50; }
        .data-table { width: 100%; border-collapse: collapse; background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; overflow: hidden; border: 1px solid rgba(255,255,255,0.06); }
        .data-table th { background: rgba(255,152,0,0.12); color: #ffb74d; font-weight: 600; padding: 12px 15px; text-align: left; font-size: 13px; }
        .data-table td { padding: 12px 15px; color: #e0e0e0; font-size: 14px; border-bottom: 1px solid rgba(255,255,255,0.04); }
        .data-table tr:hover td { background: rgba(255,255,255,0.03); }
        .section-title { font-size: 18px; color: #e0e0e0; margin-bottom: 15px; display: flex; align-items: center; gap: 10px; }
        .badge { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 500; }
        .badge-pending { background: rgba(255,152,0,0.15); color: #ffb74d; }
        .badge-shipped { background: rgba(33,150,243,0.15); color: #64b5f6; }
        .badge-intransit { background: rgba(156,39,176,0.15); color: #ba68c8; }
        .badge-delivered { background: rgba(76,175,80,0.15); color: #81c784; }
        .quick-link { display: flex; gap: 10px; margin-bottom: 25px; }
        .quick-link a { padding: 10px 18px; background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.1); border-radius: 8px; color: #b0b0b0; text-decoration: none; font-size: 13px; display: flex; align-items: center; gap: 8px; transition: all 0.2s; }
        .quick-link a:hover { background: rgba(255,152,0,0.1); border-color: rgba(255,152,0,0.3); color: #ffb74d; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <h2 class="page-title"><i class="fas fa-truck" style="color:#FF9800;"></i> 物流概览</h2>

        <div class="quick-link">
            <a href="shipping_orders.asp"><i class="fas fa-clipboard-check"></i> 待发货处理</a>
            <a href="shipments.asp"><i class="fas fa-shipping-fast"></i> 发货单管理</a>
            <a href="in_transit.asp"><i class="fas fa-route"></i> 在途跟踪</a>
            <a href="delivery_confirm.asp"><i class="fas fa-check-double"></i> 签收确认</a>
        </div>
        
        <div class="stats-grid">
            <div class="stat-card stat-pending">
                <div class="stat-label"><i class="fas fa-clipboard-list"></i> 待发货</div>
                <div class="stat-value"><%=lPending%></div>
                <div class="stat-sub">待处理订单数</div>
            </div>
            <div class="stat-card stat-amount">
                <div class="stat-label"><i class="fas fa-dollar-sign"></i> 待发货金额</div>
                <div class="stat-value" style="font-size:22px;">¥<%=FormatNumber(lPendingAmt,0)%></div>
            </div>
            <div class="stat-card stat-shipped">
                <div class="stat-label"><i class="fas fa-shipping-fast"></i> 已发货</div>
                <div class="stat-value"><%=lShipped%></div>
                <div class="stat-sub">今日发货 <%=lTodayShipped%></div>
            </div>
            <div class="stat-card stat-transit">
                <div class="stat-label"><i class="fas fa-route"></i> 运输中</div>
                <div class="stat-value"><%=lInTransit%></div>
            </div>
            <div class="stat-card stat-delivered">
                <div class="stat-label"><i class="fas fa-check-circle"></i> 已签收</div>
                <div class="stat-value"><%=lDelivered%></div>
                <div class="stat-sub">今日签收 <%=lTodayDelivered%></div>
            </div>
        </div>
        
        <div style="margin-top:25px;">
            <h3 class="section-title"><i class="fas fa-history" style="color:#FF9800;"></i> 近期发货记录</h3>
            <table class="data-table">
                <thead>
                    <tr>
                        <th>订单号</th>
                        <th>收货人</th>
                        <th>物流公司</th>
                        <th>运单号</th>
                        <th>状态</th>
                        <th>发货时间</th>
                    </tr>
                </thead>
                <tbody>
                    <% If Not rsRecentShip Is Nothing Then
                        Do While Not rsRecentShip.EOF %>
                    <tr>
                        <td><strong>#<%=Server.HTMLEncode(rsRecentShip("OrderNo") & "")%></strong></td>
                        <td><%=Server.HTMLEncode(rsRecentShip("ShippingName") & "")%></td>
                        <td><%=IIf(IsNull(rsRecentShip("ShippingCompany")) Or rsRecentShip("ShippingCompany")="","-",Server.HTMLEncode(CStr(rsRecentShip("ShippingCompany") & "")))%></td>
                        <td style="color:#81c784;"><%=IIf(IsNull(rsRecentShip("TrackingNumber")) Or rsRecentShip("TrackingNumber")="","-",Server.HTMLEncode(CStr(rsRecentShip("TrackingNumber") & "")))%></td>
                        <td><%
                            Select Case CStr(rsRecentShip("ShippingStatus") & "")
                                Case "Shipped": Response.Write "<span class='badge badge-shipped'>已发货</span>"
                                Case "InTransit": Response.Write "<span class='badge badge-intransit'>运输中</span>"
                                Case "Delivered": Response.Write "<span class='badge badge-delivered'>已签收</span>"
                                Case Else: Response.Write rsRecentShip("ShippingStatus")
                            End Select
                        %></td>
                        <td style="color:#888;"><%=IIf(IsNull(rsRecentShip("ShippedAt")) Or rsRecentShip("ShippedAt")="","-",rsRecentShip("ShippedAt"))%></td>
                    </tr>
                    <%      rsRecentShip.MoveNext
                        Loop
                        rsRecentShip.Close : Set rsRecentShip = Nothing
                    End If %>
                </tbody>
            </table>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
