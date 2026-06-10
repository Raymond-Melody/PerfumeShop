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
conn.Execute "SELECT ShippingFee FROM Orders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE Orders ADD ShippingFee DECIMAL(19,4) DEFAULT 0"
On Error GoTo 0

' POST: 更新运费
Dim sfAction : sfAction = Request.Form("action")
If sfAction = "update_fee" Then
    Dim sfId, sfFee
    sfId = Request.Form("orderId")
    sfFee = SafeNum(Request.Form("shippingFee"))
    If IsNumeric(sfId) Then
        conn.Execute "UPDATE Orders SET ShippingFee = " & sfFee & " WHERE OrderID = " & CLng(sfId)
        Response.Redirect "shipping_cost.asp?msg=运费已更新"
        Response.End
    End If
End If

Dim sfSearch : sfSearch = Replace(Request.QueryString("keyword"),"'","''")
Dim sfWhere : sfWhere = "ShippingStatus IS NOT NULL OR ShippingFee > 0"
If sfSearch <> "" Then sfWhere = "(" & sfWhere & ") AND (OrderNo LIKE '%" & sfSearch & "%' OR ShippingName LIKE '%" & sfSearch & "%')"

' 统计
Dim sfTotalFee, sfAvgFee, sfCount
sfTotalFee = SafeNum(GetScalar("SELECT ISNULL(SUM(ShippingFee),0) FROM Orders WHERE ShippingFee > 0"))
sfCount = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders WHERE ShippingFee > 0"))
If sfCount > 0 Then sfAvgFee = sfTotalFee / sfCount Else sfAvgFee = 0

' 本月运费
Dim sfMonthFee : sfMonthFee = SafeNum(GetScalar("SELECT ISNULL(SUM(ShippingFee),0) FROM Orders WHERE ShippingFee > 0 AND ShippedAt >= DATEADD(MONTH,DATEDIFF(MONTH,0,GETDATE()),0)"))

Dim rsSF
Set rsSF = conn.Execute("SELECT OrderID, OrderNo, ShippingName, ShippingCompany, ShippingStatus, ShippingFee, ShippedAt FROM Orders o WHERE " & sfWhere & " ORDER BY COALESCE(ShippedAt, CAST('2000-01-01' AS DATETIME)) DESC")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>运费管理 - 物流管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --accent: #FF5722; --input-bg: #2d2d44; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: var(--bg); color: var(--text); min-height: 100vh; }
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
        .data-table th { background: rgba(255,87,34,0.12); color: #ff8a65; font-weight: 600; padding: 12px 14px; text-align: left; font-size: 13px; }
        .data-table td { padding: 10px 14px; color: #e0e0e0; font-size: 13px; border-bottom: 1px solid rgba(255,255,255,0.04); }
        .data-table tr:hover td { background: rgba(255,255,255,0.03); }
        .modal { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.6); z-index: 999; align-items: center; justify-content: center; }
        .modal.active { display: flex; }
        .modal-content { background: #2d2d44; border-radius: 12px; padding: 24px; width: 380px; max-width: 90vw; border: 1px solid rgba(255,255,255,0.1); }
        .modal-title { font-size: 18px; color: #e0e0e0; margin-bottom: 16px; }
        .form-group { margin-bottom: 14px; }
        .form-group label { display: block; font-size: 12px; color: #888; margin-bottom: 5px; }
        .form-group input { width: 100%; padding: 9px 12px; background: var(--input-bg); border: 1px solid rgba(255,255,255,0.12); border-radius: 6px; color: #e0e0e0; font-size: 13px; }
        .form-group input:focus { outline: none; border-color: var(--accent); }
        .modal-actions { display: flex; gap: 10px; justify-content: flex-end; margin-top: 20px; }
        .alert-msg { padding: 10px 16px; border-radius: 8px; margin-bottom: 15px; font-size: 13px; }
        .alert-success { background: rgba(76,175,80,0.12); color: #81c784; border: 1px solid rgba(76,175,80,0.2); }
</style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-calculator"></i> 运费管理</h2>
            <form class="search-box" method="get">
                <input type="text" name="keyword" placeholder="搜索订单号..." value="<%=Server.HTMLEncode(sfSearch)%>">
                <button type="submit"><i class="fas fa-search"></i></button>
            </form>
        </div>

        <% Dim sfMsg : sfMsg = Request.QueryString("msg")
        If sfMsg <> "" Then %>
        <div class="alert-msg alert-success"><i class="fas fa-check-circle"></i> <%=Server.HTMLEncode(sfMsg)%></div>
        <% End If %>

        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-money-bill-wave"></i> 累计运费</div>
                <div class="stat-value">¥<%=FormatNumber(sfTotalFee,0)%></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-chart-line"></i> 平均运费</div>
                <div class="stat-value">¥<%=FormatNumber(sfAvgFee,2)%></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-calendar-alt"></i> 本月运费</div>
                <div class="stat-value">¥<%=FormatNumber(sfMonthFee,0)%></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-shipping-fast"></i> 已计费订单</div>
                <div class="stat-value"><%=sfCount%></div>
            </div>
        </div>

        <table class="data-table">
            <thead>
                <tr>
                    <th>订单号</th>
                    <th>收货人</th>
                    <th>物流公司</th>
                    <th>运费</th>
                    <th>发货状态</th>
                    <th>发货时间</th>
                    <th>操作</th>
                </tr>
            </thead>
            <tbody>
                <% If Not rsSF Is Nothing Then
                    Do While Not rsSF.EOF
                        Dim fId, fNo, fName, fCompany, fStatus, fFee, fShipped
                        fId = rsSF("OrderID")
                        fNo = rsSF("OrderNo") & ""
                        fName = rsSF("ShippingName") & ""
                        fCompany = rsSF("ShippingCompany") & ""
                        fStatus = rsSF("ShippingStatus") & ""
                        fFee = SafeNum(rsSF("ShippingFee"))
                        fShipped = rsSF("ShippedAt") & ""
                %>
                <tr>
                    <td><strong>#<%=Server.HTMLEncode(fNo)%></strong></td>
                    <td><%=Server.HTMLEncode(fName)%></td>
                    <td><%=IIf(fCompany="","-",Server.HTMLEncode(fCompany))%></td>
                    <td style="color:<%=IIf(fFee>0,"#ff8a65","#888")%>;"><%=IIf(fFee=0,"未设置","¥" & FormatNumber(fFee,2))%></td>
                    <td style="color:#888;"><%=IIf(fStatus="","-",fStatus)%></td>
                    <td style="color:#888;"><%=IIf(fShipped="","-",fShipped)%></td>
                    <td>
                        <button class="btn btn-sm btn-outline" onclick="openFee(<%=fId%>,<%=fFee%>,'<%=Server.HTMLEncode(Replace(fNo,"'","\'"))%>')"><i class="fas fa-edit"></i> 设置运费</button>
                    </td>
                </tr>
                <%      rsSF.MoveNext
                    Loop
                    rsSF.Close : Set rsSF = Nothing
                End If %>
            </tbody>
        </table>
    </div>

    <!-- 运费 Modal -->
    <div id="feeModal" class="modal">
        <div class="modal-content">
            <div class="modal-title"><i class="fas fa-calculator"></i> 设置运费 — <span id="feeOrderNo"></span></div>
            <form method="post">
                <input type="hidden" name="action" value="update_fee">
                <input type="hidden" name="orderId" id="feeOrderId">
                <div class="form-group">
                    <label>运费金额 (元)</label>
                    <input type="number" name="shippingFee" id="feeAmount" step="0.01" min="0" required>
                </div>
                <div class="modal-actions">
                    <button type="button" class="btn btn-outline" onclick="document.getElementById('feeModal').classList.remove('active')">取消</button>
                    <button type="submit" class="btn btn-primary">保存</button>
                </div>
            </form>
        </div>
    </div>

    <script>
    function openFee(id, fee, orderNo) {
        document.getElementById('feeOrderId').value = id;
        document.getElementById('feeAmount').value = fee;
        document.getElementById('feeOrderNo').textContent = '#' + orderNo;
        document.getElementById('feeModal').classList.add('active');
    }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
