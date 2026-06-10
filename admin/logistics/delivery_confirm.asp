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
conn.Execute "SELECT DeliveredAt FROM Orders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE Orders ADD DeliveredAt DATETIME"
On Error GoTo 0

Dim dcSearch : dcSearch = Replace(Request.QueryString("keyword"),"'","''")
Dim dcWhere : dcWhere = "o.ShippingStatus='Delivered'"
If dcSearch <> "" Then dcWhere = dcWhere & " AND (o.OrderNo LIKE '%" & dcSearch & "%' OR o.ShippingName LIKE '%" & dcSearch & "%')"

' 统计
Dim dcToday, dcWeek, dcMonth, dcTotal
dcToday = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders WHERE ShippingStatus='Delivered' AND DeliveredAt >= CAST(GETDATE() AS DATE)"))
dcWeek = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders WHERE ShippingStatus='Delivered' AND DeliveredAt >= DATEADD(DAY,-7,GETDATE())"))
dcMonth = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders WHERE ShippingStatus='Delivered' AND DeliveredAt >= DATEADD(DAY,-30,GETDATE())"))
dcTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders WHERE ShippingStatus='Delivered'"))

' 签收率
Dim dcShippedTotal : dcShippedTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders WHERE ShippingStatus IS NOT NULL AND ShippingStatus <> 'Pending'"))
Dim dcRate : dcRate = IIf(dcShippedTotal > 0, Round((dcTotal / dcShippedTotal) * 100, 1), 0)

Dim rsDC
Set rsDC = conn.Execute("SELECT OrderID, OrderNo, ShippingName, ShippingCompany, TrackingNumber, ShippedAt, DeliveredAt, TotalAmount FROM Orders o WHERE " & dcWhere & " ORDER BY o.DeliveredAt DESC")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>签收确认 - 物流管理中心</title>
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
        .stat-card .stat-value { font-size: 24px; font-weight: 700; color: var(--accent); }
        .stat-card .stat-sub { font-size: 11px; color: #666; margin-top: 4px; }
        .data-table { width: 100%; border-collapse: collapse; background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; overflow: hidden; border: 1px solid rgba(255,255,255,0.06); }
        .data-table th { background: rgba(76,175,80,0.12); color: #81c784; font-weight: 600; padding: 12px 14px; text-align: left; font-size: 13px; }
        .data-table td { padding: 10px 14px; color: #e0e0e0; font-size: 13px; border-bottom: 1px solid rgba(255,255,255,0.04); }
        .data-table tr:hover td { background: rgba(255,255,255,0.03); }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 10px; font-size: 11px; font-weight: 500; }
        .badge-delivered { background: rgba(76,175,80,0.15); color: #81c784; }
        .delivery-rate { display: flex; align-items: center; gap: 10px; margin-bottom: 25px; padding: 16px; background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 10px; border: 1px solid rgba(255,255,255,0.06); }
        .rate-bar { flex: 1; height: 10px; border-radius: 5px; background: rgba(255,255,255,0.1); overflow: hidden; }
        .rate-fill { height: 100%; border-radius: 5px; background: linear-gradient(90deg, #4CAF50, #81c784); }
</style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-check-double"></i> 签收确认</h2>
            <form class="search-box" method="get">
                <input type="text" name="keyword" placeholder="搜索订单号/收货人..." value="<%=Server.HTMLEncode(dcSearch)%>">
                <button type="submit"><i class="fas fa-search"></i></button>
            </form>
        </div>

        <!-- 签收率 -->
        <div class="delivery-rate">
            <span style="font-size:14px;color:#e0e0e0;"><i class="fas fa-chart-line" style="color:#4CAF50;"></i> 签收率</span>
            <div class="rate-bar">
                <div class="rate-fill" style="width:<%=dcRate%>%;"></div>
            </div>
            <span style="font-weight:700;font-size:18px;color:#4CAF50;"><%=dcRate%>%</span>
            <span style="font-size:11px;color:#888;">(<%=dcTotal%>/<%=dcShippedTotal%>)</span>
        </div>

        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-calendar-check"></i> 今日签收</div>
                <div class="stat-value"><%=dcToday%></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-calendar-week"></i> 本周签收</div>
                <div class="stat-value"><%=dcWeek%></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-calendar-alt"></i> 本月签收</div>
                <div class="stat-value"><%=dcMonth%></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-check-circle"></i> 累计签收</div>
                <div class="stat-value"><%=dcTotal%></div>
            </div>
        </div>

        <table class="data-table">
            <thead>
                <tr>
                    <th>订单号</th>
                    <th>收货人</th>
                    <th>物流公司</th>
                    <th>运单号</th>
                    <th>金额</th>
                    <th>发货时间</th>
                    <th>签收时间</th>
                    <th>耗时</th>
                </tr>
            </thead>
            <tbody>
                <% If Not rsDC Is Nothing Then
                    Do While Not rsDC.EOF
                        Dim dId, dNo, dName, dCompany, dTracking, dShipped, dDelivered, dAmt
                        dId = rsDC("OrderID")
                        dNo = rsDC("OrderNo") & ""
                        dName = rsDC("ShippingName") & ""
                        dCompany = rsDC("ShippingCompany") & ""
                        dTracking = rsDC("TrackingNumber") & ""
                        dShipped = rsDC("ShippedAt") & ""
                        dDelivered = rsDC("DeliveredAt") & ""
                        dAmt = SafeNum(rsDC("TotalAmount"))

                        Dim dTransit : dTransit = "-"
                        If dShipped <> "" And dDelivered <> "" And IsDate(dShipped) And IsDate(dDelivered) Then
                            dTransit = DateDiff("d", CDate(dShipped), CDate(dDelivered)) & "天"
                        End If
                %>
                <tr>
                    <td><strong>#<%=Server.HTMLEncode(dNo)%></strong></td>
                    <td><%=Server.HTMLEncode(dName)%></td>
                    <td><%=IIf(dCompany="","-",Server.HTMLEncode(dCompany))%></td>
                    <td style="color:#81c784;"><%=IIf(dTracking="","-",Server.HTMLEncode(dTracking))%></td>
                    <td>¥<%=FormatNumber(dAmt,2)%></td>
                    <td style="color:#888;"><%=IIf(dShipped="","-",dShipped)%></td>
                    <td style="color:#81c784;"><%=IIf(dDelivered="","-",dDelivered)%></td>
                    <td style="color:<%=IIf(dTransit="1天" Or dTransit="2天","#81c784","#888")%>;"><%=dTransit%></td>
                </tr>
                <%      rsDC.MoveNext
                    Loop
                    rsDC.Close : Set rsDC = Nothing
                End If %>
            </tbody>
        </table>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
