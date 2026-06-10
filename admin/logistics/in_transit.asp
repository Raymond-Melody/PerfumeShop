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
conn.Execute "SELECT ShippedAt FROM Orders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE Orders ADD ShippedAt DATETIME"
On Error GoTo 0

' POST: 标记为签收
Dim trAction : trAction = Request.Form("action")
If trAction = "mark_delivered" Then
    Dim trId : trId = Request.Form("orderId")
    If IsNumeric(trId) Then
        conn.Execute "UPDATE Orders SET ShippingStatus='Delivered', Status='Completed', DeliveredAt=GETDATE(), UpdatedAt=GETDATE() WHERE OrderID=" & CLng(trId)
        Response.Redirect "in_transit.asp?msg=已标记为签收"
        Response.End
    End If
End If

Dim trSearch : trSearch = Replace(Request.QueryString("keyword"),"'","''")
Dim trWhere : trWhere = "o.ShippingStatus IN ('Shipped','InTransit')"
If trSearch <> "" Then trWhere = trWhere & " AND (o.OrderNo LIKE '%" & trSearch & "%' OR o.ShippingName LIKE '%" & trSearch & "%' OR o.TrackingNumber LIKE '%" & trSearch & "%')"

' 统计
Dim trTotal, trShipped, trInTransit, trOverdue
trTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders o WHERE o.ShippingStatus IN ('Shipped','InTransit')"))
trShipped = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders o WHERE o.ShippingStatus='Shipped'"))
trInTransit = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders o WHERE o.ShippingStatus='InTransit'"))
trOverdue = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders o WHERE o.ShippingStatus IN ('Shipped','InTransit') AND o.ShippedAt < DATEADD(DAY,-7,GETDATE())"))

Dim rsTR
Set rsTR = conn.Execute("SELECT o.OrderID, o.OrderNo, o.ShippingName, o.ShippingCompany, o.TrackingNumber, o.ShippingStatus, o.ShippedAt, o.ShippingFee, o.ShippingAddress, o.ShippingCity FROM Orders o WHERE " & trWhere & " ORDER BY o.ShippedAt")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>在途跟踪 - 物流管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --accent: #9C27B0; --input-bg: #2d2d44; }
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
        .stat-card .stat-value { font-size: 24px; font-weight: 700; color: var(--accent); }
        .data-table { width: 100%; border-collapse: collapse; background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; overflow: hidden; border: 1px solid rgba(255,255,255,0.06); }
        .data-table th { background: rgba(156,39,176,0.12); color: #ba68c8; font-weight: 600; padding: 12px 14px; text-align: left; font-size: 13px; }
        .data-table td { padding: 10px 14px; color: #e0e0e0; font-size: 13px; border-bottom: 1px solid rgba(255,255,255,0.04); }
        .data-table tr:hover td { background: rgba(255,255,255,0.03); }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 10px; font-size: 11px; font-weight: 500; }
        .badge-shipped { background: rgba(33,150,243,0.15); color: #64b5f6; }
        .badge-intransit { background: rgba(156,39,176,0.15); color: #ba68c8; }
        .badge-overdue { background: rgba(244,67,54,0.15); color: #ef9a9a; }
        .alert-msg { padding: 10px 16px; border-radius: 8px; margin-bottom: 15px; font-size: 13px; }
        .alert-success { background: rgba(76,175,80,0.12); color: #81c784; border: 1px solid rgba(76,175,80,0.2); }
        .timeline-bar { display: flex; align-items: center; gap: 0; width: 120px; }
        .t-dot { width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; }
        .t-dot.done { background: #4CAF50; }
        .t-dot.active { background: #9C27B0; box-shadow: 0 0 6px rgba(156,39,176,0.5); animation: pulse 1.5s infinite; }
        @keyframes pulse { 0%,100% { transform: scale(1); } 50% { transform: scale(1.3); } }
        .t-line { flex: 1; height: 2px; background: rgba(255,255,255,0.2); }
        .t-line.done { background: #4CAF50; }
</style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-route"></i> 在途跟踪</h2>
            <form class="search-box" method="get">
                <input type="text" name="keyword" placeholder="搜索订单/运单/收货人..." value="<%=Server.HTMLEncode(trSearch)%>">
                <button type="submit"><i class="fas fa-search"></i></button>
            </form>
        </div>

        <% Dim trMsg : trMsg = Request.QueryString("msg")
        If trMsg <> "" Then %>
        <div class="alert-msg alert-success"><i class="fas fa-check-circle"></i> <%=Server.HTMLEncode(trMsg)%></div>
        <% End If %>

        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-shipping-fast"></i> 总在途</div>
                <div class="stat-value"><%=trTotal%></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-box"></i> 刚发货</div>
                <div class="stat-value" style="color:#2196F3;"><%=trShipped%></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-truck"></i> 运输中</div>
                <div class="stat-value"><%=trInTransit%></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-exclamation-triangle"></i> 超7天未签收</div>
                <div class="stat-value" style="color:#f44336;"><%=trOverdue%></div>
            </div>
        </div>

        <table class="data-table">
            <thead>
                <tr>
                    <th>订单号</th>
                    <th>收货人</th>
                    <th>物流公司</th>
                    <th>运单号</th>
                    <th>天数</th>
                    <th>状态</th>
                    <th>发货时间</th>
                    <th>操作</th>
                </tr>
            </thead>
            <tbody>
                <% If Not rsTR Is Nothing Then
                    Do While Not rsTR.EOF
                        Dim tId, tNo, tName, tCompany, tTracking, tStatus, tShipped, tFee
                        tId = rsTR("OrderID")
                        tNo = rsTR("OrderNo") & ""
                        tName = rsTR("ShippingName") & ""
                        tCompany = rsTR("ShippingCompany") & ""
                        tTracking = rsTR("TrackingNumber") & ""
                        tStatus = rsTR("ShippingStatus") & ""
                        tShipped = rsTR("ShippedAt") & ""

                        Dim tDays : tDays = 0
                        If tShipped <> "" And IsDate(tShipped) Then tDays = DateDiff("d", CDate(tShipped), Date())
                %>
                <tr>
                    <td><strong>#<%=Server.HTMLEncode(tNo)%></strong></td>
                    <td><%=Server.HTMLEncode(tName)%></td>
                    <td><%=IIf(tCompany="","-",Server.HTMLEncode(tCompany))%></td>
                    <td style="color:#81c784;"><%=IIf(tTracking="","-",Server.HTMLEncode(tTracking))%></td>
                    <td style="color:<%=IIf(tDays>=7,"#f44336","#888")%>;"><%=IIf(tShipped="","-",tDays & "天")%></td>
                    <td><%=IIf(tStatus="Shipped","<span class='badge badge-shipped'>刚发货</span>","<span class='badge badge-intransit'>运输中</span>")%></td>
                    <td style="color:#888;"><%=IIf(tShipped="","-",tShipped)%></td>
                    <td>
                        <form method="post" style="display:inline;">
                            <input type="hidden" name="action" value="mark_delivered">
                            <input type="hidden" name="orderId" value="<%=tId%>">
                            <button type="submit" class="btn btn--success btn--sm"><i class="fas fa-check"></i> 确认签收</button>
                        </form>
                    </td>
                </tr>
                <%      rsTR.MoveNext
                    Loop
                    rsTR.Close : Set rsTR = Nothing
                End If %>
            </tbody>
        </table>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
