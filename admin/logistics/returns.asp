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
conn.Execute "SELECT RefundAmount FROM Orders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE Orders ADD RefundAmount DECIMAL(19,4) DEFAULT 0"
On Error GoTo 0

' POST
Dim retAction : retAction = Request.Form("action")
If retAction = "mark_return" Then
    Dim retId, retNotes, retFee
    retId = Request.Form("orderId")
    retNotes = Replace(Request.Form("returnNotes"),"'","''")
    retFee = SafeNum(Request.Form("refundAmount"))
    If IsNumeric(retId) Then
        conn.Execute "UPDATE Orders SET Status='Refunded', ShippingStatus='Returned', RefundAmount=" & retFee & ", ShippingNotes=ISNULL(ShippingNotes,'') + ' [退货: " & retNotes & "]', UpdatedAt=GETDATE() WHERE OrderID=" & CLng(retId)
        Response.Redirect "returns.asp?msg=退货已处理"
        Response.End
    End If
End If

Dim retSearch : retSearch = Replace(Request.QueryString("keyword"),"'","''")
Dim retWhere : retWhere = "o.Status IN ('Refunded','Cancelled') OR o.ShippingStatus='Returned'"
If retSearch <> "" Then retWhere = "(" & retWhere & ") AND (o.OrderNo LIKE '%" & retSearch & "%' OR o.ShippingName LIKE '%" & retSearch & "%')"

' 统计
Dim retTotal, retMonth, retAmount
retTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders WHERE Status='Refunded' OR ShippingStatus='Returned'"))
retMonth = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders WHERE (Status='Refunded' OR ShippingStatus='Returned') AND UpdatedAt >= DATEADD(MONTH,DATEDIFF(MONTH,0,GETDATE()),0)"))
retAmount = SafeNum(GetScalar("SELECT ISNULL(SUM(ISNULL(RefundAmount,TotalAmount)),0) FROM Orders WHERE Status='Refunded' OR ShippingStatus='Returned'"))

Dim rsRet
Set rsRet = conn.Execute("SELECT o.OrderID, o.OrderNo, o.ShippingName, o.TotalAmount, o.RefundAmount, o.Status, o.ShippingStatus, o.ShippingCompany, o.TrackingNumber, o.UpdatedAt, o.ShippingNotes FROM Orders o WHERE " & retWhere & " ORDER BY o.UpdatedAt DESC")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>退货入库 - 物流管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --accent: #f44336; --input-bg: #2d2d44; }
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
        .data-table th { background: rgba(244,67,54,0.12); color: #ef9a9a; font-weight: 600; padding: 12px 14px; text-align: left; font-size: 13px; }
        .data-table td { padding: 10px 14px; color: #e0e0e0; font-size: 13px; border-bottom: 1px solid rgba(255,255,255,0.04); }
        .data-table tr:hover td { background: rgba(255,255,255,0.03); }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 10px; font-size: 11px; font-weight: 500; }
        .badge-refunded { background: rgba(244,67,54,0.15); color: #ef9a9a; }
        .badge-returned { background: rgba(255,152,0,0.15); color: #ffb74d; }
        .modal { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.6); z-index: 999; align-items: center; justify-content: center; }
        .modal.active { display: flex; }
        .modal-content { background: #2d2d44; border-radius: 12px; padding: 24px; width: 450px; max-width: 90vw; border: 1px solid rgba(255,255,255,0.1); }
        .modal-title { font-size: 18px; color: #e0e0e0; margin-bottom: 16px; }
        .form-group { margin-bottom: 14px; }
        .form-group label { display: block; font-size: 12px; color: #888; margin-bottom: 5px; }
        .form-group input, .form-group textarea { width: 100%; padding: 9px 12px; background: var(--input-bg); border: 1px solid rgba(255,255,255,0.12); border-radius: 6px; color: #e0e0e0; font-size: 13px; }
        .form-group textarea { resize: vertical; min-height: 60px; }
        .form-group input:focus, .form-group textarea:focus { outline: none; border-color: var(--accent); }
        .modal-actions { display: flex; gap: 10px; justify-content: flex-end; margin-top: 20px; }
        .alert-msg { padding: 10px 16px; border-radius: 8px; margin-bottom: 15px; font-size: 13px; }
        .alert-success { background: rgba(76,175,80,0.12); color: #81c784; border: 1px solid rgba(76,175,80,0.2); }
</style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-undo-alt"></i> 退货入库管理</h2>
            <form class="search-box" method="get">
                <input type="text" name="keyword" placeholder="搜索订单号..." value="<%=Server.HTMLEncode(retSearch)%>">
                <button type="submit"><i class="fas fa-search"></i></button>
            </form>
        </div>

        <% Dim retMsg : retMsg = Request.QueryString("msg")
        If retMsg <> "" Then %>
        <div class="alert-msg alert-success"><i class="fas fa-check-circle"></i> <%=Server.HTMLEncode(retMsg)%></div>
        <% End If %>

        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-undo"></i> 累计退货</div>
                <div class="stat-value"><%=retTotal%></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-calendar-alt"></i> 本月退货</div>
                <div class="stat-value"><%=retMonth%></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-money-bill-wave"></i> 退款总额</div>
                <div class="stat-value" style="font-size:20px;">¥<%=FormatNumber(retAmount,0)%></div>
            </div>
        </div>

        <table class="data-table">
            <thead>
                <tr>
                    <th>订单号</th>
                    <th>收货人</th>
                    <th>订单金额</th>
                    <th>退款金额</th>
                    <th>状态</th>
                    <th>物流公司</th>
                    <th>运单号</th>
                    <th>处理时间</th>
                    <th>备注</th>
                </tr>
            </thead>
            <tbody>
                <% If Not rsRet Is Nothing Then
                    Do While Not rsRet.EOF
                        Dim rId, rNo, rName, rAmt, rRefund, rStatus, rShipStatus, rCompany, rTracking, rUpdated, rNotes
                        rId = rsRet("OrderID")
                        rNo = rsRet("OrderNo") & ""
                        rName = rsRet("ShippingName") & ""
                        rAmt = SafeNum(rsRet("TotalAmount"))
                        rRefund = SafeNum(rsRet("RefundAmount"))
                        rStatus = rsRet("Status") & ""
                        rShipStatus = rsRet("ShippingStatus") & ""
                        rCompany = rsRet("ShippingCompany") & ""
                        rTracking = rsRet("TrackingNumber") & ""
                        rUpdated = rsRet("UpdatedAt") & ""
                        rNotes = rsRet("ShippingNotes") & ""
                %>
                <tr>
                    <td><strong>#<%=Server.HTMLEncode(rNo)%></strong></td>
                    <td><%=Server.HTMLEncode(rName)%></td>
                    <td>¥<%=FormatNumber(rAmt,2)%></td>
                    <td style="color:#ef9a9a;">¥<%=IIf(rRefund=0,"-",FormatNumber(rRefund,2))%></td>
                    <td>
                        <% If rStatus = "Refunded" Then %><span class="badge badge-refunded">已退款</span>
                        <% ElseIf rShipStatus = "Returned" Then %><span class="badge badge-returned">已退货</span>
                        <% Else %><%=rStatus%><% End If %>
                    </td>
                    <td><%=IIf(rCompany="","-",Server.HTMLEncode(rCompany))%></td>
                    <td style="color:#81c784;"><%=IIf(rTracking="","-",Server.HTMLEncode(rTracking))%></td>
                    <td style="color:#888;"><%=IIf(rUpdated="","-",rUpdated)%></td>
                    <td style="color:#888;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;"><%=IIf(rNotes="","-",Server.HTMLEncode(rNotes))%></td>
                </tr>
                <%      rsRet.MoveNext
                    Loop
                    rsRet.Close : Set rsRet = Nothing
                End If %>
            </tbody>
        </table>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
