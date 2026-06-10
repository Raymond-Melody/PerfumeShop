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
conn.Execute "SELECT ShippingCompany FROM Orders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE Orders ADD ShippingCompany NVARCHAR(50)"
conn.Execute "SELECT TrackingNumber FROM Orders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE Orders ADD TrackingNumber NVARCHAR(100)"
conn.Execute "SELECT ShippedAt FROM Orders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE Orders ADD ShippedAt DATETIME"
On Error GoTo 0

' POST: 发货
Dim soAction : soAction = Request.Form("action")
If soAction = "ship" Then
    Dim sId, soCompany, soTracking, soNotes
    sId = Request.Form("orderId")
    soCompany = Replace(Request.Form("shippingCompany"),"'","''")
    soTracking = Replace(Request.Form("trackingNumber"),"'","''")
    soNotes = Replace(Request.Form("shippingNotes"),"'","''")
    If IsNumeric(sId) Then
        conn.Execute "UPDATE Orders SET ShippingStatus='Shipped', ShippingCompany='" & soCompany & "', TrackingNumber='" & soTracking & "', ShippingNotes='" & soNotes & "', ShippedAt=GETDATE(), UpdatedAt=GETDATE() WHERE OrderID=" & CLng(sId)
        Response.Redirect "shipping_orders.asp?msg=发货成功"
        Response.End
    End If
ElseIf soAction = "batch_ship" Then
    Dim bsIds, bsIdsArr, bsI, bsCnt : bsCnt = 0
    bsIds = Request.Form("batchIds")
    If bsIds <> "" Then
        bsIdsArr = Split(bsIds, ",")
        For bsI = 0 To UBound(bsIdsArr)
            If IsNumeric(Trim(bsIdsArr(bsI))) Then
                conn.Execute "UPDATE Orders SET ShippingStatus='Shipped', ShippedAt=GETDATE(), UpdatedAt=GETDATE() WHERE OrderID=" & CLng(Trim(bsIdsArr(bsI))) & " AND (ShippingStatus='Pending' OR ShippingStatus IS NULL)"
                bsCnt = bsCnt + 1
            End If
        Next
        Response.Redirect "shipping_orders.asp?msg=批量发货成功，" & bsCnt & "笔订单已发"
        Response.End
    End If
End If

Dim soSearch : soSearch = Replace(Request.QueryString("keyword"),"'","''")
Dim soWhere : soWhere = "o.Status IN ('Paid','Processing','Shipped') AND (o.ShippingStatus='Pending' OR o.ShippingStatus IS NULL)"
If soSearch <> "" Then soWhere = soWhere & " AND (o.OrderNo LIKE '%" & soSearch & "%' OR o.ShippingName LIKE '%" & soSearch & "%')"

Dim soCount : soCount = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders o WHERE " & soWhere))
Dim soAmt : soAmt = SafeNum(GetScalar("SELECT ISNULL(SUM(TotalAmount),0) FROM Orders o WHERE " & soWhere))

Dim rsSO
Set rsSO = conn.Execute("SELECT o.OrderID, o.OrderNo, o.ShippingName, o.TotalAmount, o.Status, o.ShippingAddress, o.ShippingCity, o.ShippingPhone, o.ShippingFee, " & _
    "(SELECT TOP 1 p.Status FROM ProductionOrders p WHERE p.OrderID=o.OrderID ORDER BY p.CreatedAt DESC) AS POStatus " & _
    "FROM Orders o WHERE " & soWhere & " ORDER BY o.CreatedAt")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>待发货订单 - 物流管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --accent: #FF9800; --input-bg: #2d2d44; }
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
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 10px; padding: 16px; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .stat-label { font-size: 11px; color: #888; margin-bottom: 6px; }
        .stat-card .stat-value { font-size: 24px; font-weight: 700; color: var(--accent); }
        .data-table { width: 100%; border-collapse: collapse; background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; overflow: hidden; border: 1px solid rgba(255,255,255,0.06); }
        .data-table th { background: rgba(255,152,0,0.12); color: #ffb74d; font-weight: 600; padding: 12px 14px; text-align: left; font-size: 13px; }
        .data-table td { padding: 10px 14px; color: #e0e0e0; font-size: 13px; border-bottom: 1px solid rgba(255,255,255,0.04); }
        .data-table tr:hover td { background: rgba(255,255,255,0.03); }
        .modal { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.6); z-index: 999; align-items: center; justify-content: center; }
        .modal.active { display: flex; }
        .modal-content { background: #2d2d44; border-radius: 12px; padding: 24px; width: 480px; max-width: 90vw; border: 1px solid rgba(255,255,255,0.1); }
        .modal-title { font-size: 18px; color: #e0e0e0; margin-bottom: 16px; }
        .form-group { margin-bottom: 14px; }
        .form-group label { display: block; font-size: 12px; color: #888; margin-bottom: 5px; }
        .form-group input, .form-group textarea { width: 100%; padding: 9px 12px; background: var(--input-bg); border: 1px solid rgba(255,255,255,0.12); border-radius: 6px; color: #e0e0e0; font-size: 13px; }
        .form-group textarea { resize: vertical; min-height: 60px; }
        .form-group input:focus, .form-group textarea:focus { outline: none; border-color: var(--accent); }
        .modal-actions { display: flex; gap: 10px; justify-content: flex-end; margin-top: 20px; }
        .alert-msg { padding: 10px 16px; border-radius: 8px; margin-bottom: 15px; font-size: 13px; }
        .alert-success { background: rgba(76,175,80,0.12); color: #81c784; border: 1px solid rgba(76,175,80,0.2); }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 10px; font-size: 11px; font-weight: 500; }
        .badge-paid { background: rgba(33,150,243,0.12); color: #64b5f6; }
        .badge-processing { background: rgba(255,152,0,0.12); color: #ffb74d; }
</style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-clipboard-check"></i> 待发货订单</h2>
            <form class="search-box" method="get">
                <input type="text" name="keyword" placeholder="搜索订单号/收货人..." value="<%=Server.HTMLEncode(soSearch)%>">
                <button type="submit"><i class="fas fa-search"></i> 搜索</button>
            </form>
        </div>

        <% Dim soMsg : soMsg = Request.QueryString("msg")
        If soMsg <> "" Then %>
        <div class="alert-msg alert-success"><i class="fas fa-check-circle"></i> <%=Server.HTMLEncode(soMsg)%></div>
        <% End If %>

        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-clipboard-list"></i> 待发货订单</div>
                <div class="stat-value"><%=soCount%></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-dollar-sign"></i> 待发货金额</div>
                <div class="stat-value" style="font-size:22px;">¥<%=FormatNumber(soAmt,0)%></div>
            </div>
        </div>

        <div style="margin-bottom:15px; display:flex; gap:10px;">
            <button class="btn btn-primary" onclick="batchShip()"><i class="fas fa-rocket"></i> 批量发货</button>
        </div>

        <form id="shipForm" method="post">
            <input type="hidden" name="action" id="shipAction">
            <input type="hidden" name="orderId" id="shipOrderId">
            <input type="hidden" name="batchIds" id="batchIds">
            <table class="data-table">
                <thead>
                    <tr>
                        <th style="width:40px;"><input type="checkbox" id="selectAll" onchange="toggleAll(this)"></th>
                        <th>订单号</th>
                        <th>收货人</th>
                        <th>金额</th>
                        <th>收货地址</th>
                        <th>支付状态</th>
                        <th>生产状态</th>
                        <th>操作</th>
                    </tr>
                </thead>
                <tbody>
                    <% If Not rsSO Is Nothing Then
                        Do While Not rsSO.EOF
                            Dim sOId, sONo, sOName, oAmt, sOAddr, sOCity, sOPhone, sOFee, sOStatus, sOPOStatus
                            sOId = rsSO("OrderID")
                            sONo = rsSO("OrderNo") & ""
                            sOName = rsSO("ShippingName") & ""
                            oAmt = SafeNum(rsSO("TotalAmount"))
                            sOAddr = (rsSO("ShippingCity") & "") & " " & (rsSO("ShippingAddress") & "")
                            sOPhone = rsSO("ShippingPhone") & ""
                            sOFee = SafeNum(rsSO("ShippingFee"))
                            sOStatus = rsSO("Status") & ""
                            sOPOStatus = rsSO("POStatus") & ""
                    %>
                    <tr>
                        <td><input type="checkbox" class="order-check" value="<%=sOId%>"></td>
                        <td><strong>#<%=Server.HTMLEncode(sONo)%></strong></td>
                        <td><%=Server.HTMLEncode(sOName)%><br><small style="color:#888;"><%=Server.HTMLEncode(sOPhone)%></small></td>
                        <td>¥<%=FormatNumber(oAmt,2)%></td>
                        <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#888;"><%=Server.HTMLEncode(sOAddr)%></td>
                        <td><span class="badge badge-<%=LCase(sOStatus)%>"><%=sOStatus%></span></td>
                        <td style="color:<%=IIf(sOPOStatus="","#888","#81c784")%>;"><%=IIf(sOPOStatus="","未排产",sOPOStatus)%></td>
                        <td>
                            <button type="button" class="btn btn-outline" style="font-size:11px;" onclick="openShipModal(<%=sOId%>,'<%=Server.HTMLEncode(Replace(sONo,"'","\'"))%>','<%=Server.HTMLEncode(Replace(sOName,"'","\'"))%>')"><i class="fas fa-shipping-fast"></i> 发货</button>
                        </td>
                    </tr>
                    <%      rsSO.MoveNext
                        Loop
                        rsSO.Close : Set rsSO = Nothing
                    End If %>
                </tbody>
            </table>
        </form>
    </div>

    <!-- 发货 Modal -->
    <div id="shipModal" class="modal">
        <div class="modal-content">
            <div class="modal-title"><i class="fas fa-shipping-fast"></i> 确认发货 — <span id="shipOrderTitle"></span></div>
            <form method="post">
                <input type="hidden" name="action" value="ship">
                <input type="hidden" name="orderId" id="shipModalOrderId">
                <div class="form-group">
                    <label>物流公司 *</label>
                    <input type="text" name="shippingCompany" id="shipCompany" placeholder="如：顺丰速运、中通快递" required>
                </div>
                <div class="form-group">
                    <label>运单号 *</label>
                    <input type="text" name="trackingNumber" id="shipTracking" placeholder="快递单号" required>
                </div>
                <div class="form-group">
                    <label>备注</label>
                    <textarea name="shippingNotes" placeholder="发货备注..."></textarea>
                </div>
                <div class="modal-actions">
                    <button type="button" class="btn btn-outline" onclick="document.getElementById('shipModal').classList.remove('active')">取消</button>
                    <button type="submit" class="btn btn-primary">确认发货</button>
                </div>
            </form>
        </div>
    </div>

    <script>
    var selectedOrder = [];
    function openShipModal(id, orderNo, name) {
        document.getElementById('shipModalOrderId').value = id;
        document.getElementById('shipOrderTitle').textContent = '#' + orderNo + ' — ' + name;
        document.getElementById('shipModal').classList.add('active');
    }
    function toggleAll(cb) {
        var checks = document.querySelectorAll('.order-check');
        checks.forEach(function(c){ c.checked = cb.checked; });
    }
    function batchShip() {
        var checks = document.querySelectorAll('.order-check:checked');
        if (checks.length === 0) { alert('请至少选择一条订单'); return; }
        var ids = [];
        checks.forEach(function(c){ ids.push(c.value); });
        document.getElementById('shipAction').value = 'batch_ship';
        document.getElementById('batchIds').value = ids.join(',');
        document.getElementById('shipForm').submit();
    }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
