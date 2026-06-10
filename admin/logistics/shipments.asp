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

On Error Resume Next
conn.Execute "SELECT ShippingStatus FROM Orders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE Orders ADD ShippingStatus NVARCHAR(20) DEFAULT 'Pending'"
conn.Execute "SELECT ShippedAt FROM Orders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE Orders ADD ShippedAt DATETIME"
conn.Execute "SELECT DeliveredAt FROM Orders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE Orders ADD DeliveredAt DATETIME"
On Error GoTo 0

' POST: 更新为运输中/已签收等
Dim shAction : shAction = Request.Form("action")
If shAction = "mark_intransit" Then
    Dim shId : shId = Request.Form("orderId")
    If IsNumeric(shId) Then
        conn.Execute "UPDATE Orders SET ShippingStatus='InTransit', UpdatedAt=GETDATE() WHERE OrderID=" & CLng(shId)
        Response.Redirect "shipments.asp?msg=已标记为运输中"
        Response.End
    End If
ElseIf shAction = "mark_delivered" Then
    shId = Request.Form("orderId")
    If IsNumeric(shId) Then
        conn.Execute "UPDATE Orders SET ShippingStatus='Delivered', Status='Completed', DeliveredAt=GETDATE(), UpdatedAt=GETDATE() WHERE OrderID=" & CLng(shId)
        Response.Redirect "shipments.asp?msg=已标记为签收，订单完成"
        Response.End
    End If
End If

Dim shFilter : shFilter = Request.QueryString("status")
Dim shSearch : shSearch = Replace(Request.QueryString("keyword"),"'","''")
Dim shWhere : shWhere = "(o.ShippingStatus='Shipped' OR o.ShippingStatus='InTransit' OR o.ShippingStatus='Delivered')"
If shFilter = "Shipped" Then shWhere = "o.ShippingStatus='Shipped'"
If shFilter = "InTransit" Then shWhere = "o.ShippingStatus='InTransit'"
If shFilter = "Delivered" Then shWhere = "o.ShippingStatus='Delivered'"
If shSearch <> "" Then shWhere = shWhere & " AND (o.OrderNo LIKE '%" & shSearch & "%' OR o.ShippingName LIKE '%" & shSearch & "%' OR o.TrackingNumber LIKE '%" & shSearch & "%')"

' 统计
Dim shTotal, shShipped, shTransit, shDelivered
shTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders o WHERE o.ShippingStatus IN ('Shipped','InTransit','Delivered')"))
shShipped = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders o WHERE o.ShippingStatus='Shipped'"))
shTransit = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders o WHERE o.ShippingStatus='InTransit'"))
shDelivered = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders o WHERE o.ShippingStatus='Delivered'"))

Dim rsSH
Set rsSH = conn.Execute("SELECT o.OrderID, o.OrderNo, o.ShippingName, o.TotalAmount, o.ShippingCompany, o.TrackingNumber, o.ShippingStatus, o.ShippingNotes, o.ShippedAt, o.DeliveredAt, o.ShippingFee FROM Orders o WHERE " & shWhere & " ORDER BY o.ShippedAt DESC")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>发货单管理 - 物流管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --accent: #2196F3; --input-bg: #2d2d44; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #1a1a2e; color: var(--text); min-height: 100vh; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-header { margin-bottom: 25px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px; }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: var(--accent); }
        .filter-tabs { display: flex; gap: 6px; }
        .filter-tab { padding: 6px 16px; border-radius: 20px; font-size: 12px; color: #888; text-decoration: none; border: 1px solid rgba(255,255,255,0.1); transition: all 0.2s; }
        .filter-tab:hover { color: #e0e0e0; border-color: rgba(255,255,255,0.25); }
        .filter-tab.active { background: rgba(33,150,243,0.15); color: #64b5f6; border-color: rgba(33,150,243,0.3); }
        .search-box { display: flex; gap: 8px; }
        .search-box input { padding: 8px 14px; background: var(--input-bg); border: 1px solid rgba(255,255,255,0.12); border-radius: 6px; color: #e0e0e0; font-size: 13px; width: 220px; }
        .search-box input:focus { outline: none; border-color: var(--accent); }
        .search-box button { padding: 8px 14px; background: var(--accent); border: none; border-radius: 6px; color: #fff; cursor: pointer; font-size: 13px; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 10px; padding: 16px; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .stat-label { font-size: 11px; color: #888; margin-bottom: 6px; }
        .stat-card .stat-value { font-size: 24px; font-weight: 700; color: var(--accent); }
        .data-table { width: 100%; border-collapse: collapse; background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; overflow: hidden; border: 1px solid rgba(255,255,255,0.06); }
        .data-table th { background: rgba(33,150,243,0.12); color: #64b5f6; font-weight: 600; padding: 12px 14px; text-align: left; font-size: 13px; }
        .data-table td { padding: 10px 14px; color: #e0e0e0; font-size: 13px; border-bottom: 1px solid rgba(255,255,255,0.04); }
        .data-table tr:hover td { background: rgba(255,255,255,0.03); }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 10px; font-size: 11px; font-weight: 500; }
        .badge-shipped { background: rgba(33,150,243,0.15); color: #64b5f6; }
        .badge-intransit { background: rgba(156,39,176,0.15); color: #ba68c8; }
        .badge-delivered { background: rgba(76,175,80,0.15); color: #81c784; }
        .alert-msg { padding: 10px 16px; border-radius: 8px; margin-bottom: 15px; font-size: 13px; }
        .alert-success { background: rgba(76,175,80,0.12); color: #81c784; border: 1px solid rgba(76,175,80,0.2); }
</style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-shipping-fast"></i> 发货单管理</h2>
            <form class="search-box" method="get">
                <input type="hidden" name="status" value="<%=shFilter%>">
                <input type="text" name="keyword" placeholder="搜索订单/运单/收货人..." value="<%=Server.HTMLEncode(shSearch)%>">
                <button type="submit"><i class="fas fa-search"></i></button>
            </form>
        </div>

        <% Dim shMsg : shMsg = Request.QueryString("msg")
        If shMsg <> "" Then %>
        <div class="alert-msg alert-success"><i class="fas fa-check-circle"></i> <%=Server.HTMLEncode(shMsg)%></div>
        <% End If %>

        <div class="filter-tabs" style="margin-bottom:20px;">
            <a href="shipments.asp" class="filter-tab <%=IIf(shFilter="","active","")%>">全部 (<%=shTotal%>)</a>
            <a href="shipments.asp?status=Shipped" class="filter-tab <%=IIf(shFilter="Shipped","active","")%>"><i class="fas fa-box"></i> 已发货 (<%=shShipped%>)</a>
            <a href="shipments.asp?status=InTransit" class="filter-tab <%=IIf(shFilter="InTransit","active","")%>"><i class="fas fa-truck"></i> 运输中 (<%=shTransit%>)</a>
            <a href="shipments.asp?status=Delivered" class="filter-tab <%=IIf(shFilter="Delivered","active","")%>"><i class="fas fa-check-circle"></i> 已签收 (<%=shDelivered%>)</a>
        </div>

        <table class="data-table">
            <thead>
                <tr>
                    <th>订单号</th>
                    <th>收货人</th>
                    <th>物流公司</th>
                    <th>运单号</th>
                    <th>运费</th>
                    <th>状态</th>
                    <th>发货时间</th>
                    <th>签收时间</th>
                    <th>操作</th>
                </tr>
            </thead>
            <tbody>
                <% If Not rsSH Is Nothing Then
                    Do While Not rsSH.EOF
                        Dim sId, sNo, sName, sCompany, sTracking, sStatus, sShipped, sDelivered, sFee, sNotes
                        sId = rsSH("OrderID")
                        sNo = rsSH("OrderNo") & ""
                        sName = rsSH("ShippingName") & ""
                        sCompany = rsSH("ShippingCompany") & ""
                        sTracking = rsSH("TrackingNumber") & ""
                        sStatus = rsSH("ShippingStatus") & ""
                        sShipped = rsSH("ShippedAt") & ""
                        sDelivered = rsSH("DeliveredAt") & ""
                        sFee = SafeNum(rsSH("ShippingFee"))
                %>
                <tr>
                    <td><strong>#<%=Server.HTMLEncode(sNo)%></strong></td>
                    <td><%=Server.HTMLEncode(sName)%></td>
                    <td><%=IIf(sCompany="","-",Server.HTMLEncode(sCompany))%></td>
                    <td style="color:#81c784;"><%=IIf(sTracking="","-",Server.HTMLEncode(sTracking))%></td>
                    <td>¥<%=IIf(sFee=0,"-",FormatNumber(sFee,2))%></td>
                    <td><%
                        Select Case sStatus
                            Case "Shipped": Response.Write "<span class='badge badge-shipped'>已发货</span>"
                            Case "InTransit": Response.Write "<span class='badge badge-intransit'>运输中</span>"
                            Case "Delivered": Response.Write "<span class='badge badge-delivered'>已签收</span>"
                            Case Else: Response.Write sStatus
                        End Select
                    %></td>
                    <td style="color:#888;"><%=IIf(sShipped="","-",sShipped)%></td>
                    <td style="color:#888;"><%=IIf(sDelivered="","-",sDelivered)%></td>
                    <td>
                        <% If sStatus = "Shipped" Then %>
                        <form method="post" style="display:inline;">
                            <input type="hidden" name="action" value="mark_intransit">
                            <input type="hidden" name="orderId" value="<%=sId%>">
                            <button type="submit" class="btn btn--purple btn--sm"><i class="fas fa-truck"></i> 运输中</button>
                        </form>
                        <% End If %>
                        <% If sStatus = "Shipped" Or sStatus = "InTransit" Then %>
                        <form method="post" style="display:inline;">
                            <input type="hidden" name="action" value="mark_delivered">
                            <input type="hidden" name="orderId" value="<%=sId%>">
                            <button type="submit" class="btn btn--success btn--sm"><i class="fas fa-check"></i> 签收</button>
                        </form>
                        <% End If %>
                    </td>
                </tr>
                <%      rsSH.MoveNext
                    Loop
                    rsSH.Close : Set rsSH = Nothing
                End If %>
            </tbody>
        </table>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
