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

' V10.1: 移除运行时DDL，表应由 deploy.asp 预先创建
Function TableExists_AP(tblName)
    Dim rs, exists : exists = False
    On Error Resume Next
    Set rs = conn.Execute("SELECT 1 FROM sys.tables WHERE name='" & Replace(tblName,"'","''") & "'")
    If Err.Number = 0 Then
        If Not rs Is Nothing Then
            If Not rs.EOF Then exists = True
            rs.Close
        End If
    End If
    Err.Clear : Set rs = Nothing
    On Error GoTo 0
    TableExists_AP = exists
End Function

If Not TableExists_AP("AccountsPayable") Then
    Response.Write "<div style='padding:40px;color:#f44336;background:#1a1a2e;font-family:sans-serif;'><h2>表缺失</h2><p>AccountsPayable 表不存在，请先运行 <a href='/setup/deploy.asp' style='color:#00bcd4;'>系统部署工具</a> 创建数据库表。</p></div>"
    Response.End
End If

Dim filterStatus : filterStatus = Request.QueryString("status")

' 统计
Dim apTotal : apTotal = GetScalar("SELECT COUNT(*) FROM AccountsPayable")
Dim apPending : apPending = GetScalar("SELECT COUNT(*) FROM AccountsPayable WHERE Status='Pending'")
Dim apOverdue : apOverdue = GetScalar("SELECT COUNT(*) FROM AccountsPayable WHERE Status='Pending' AND DueDate < CAST(GETDATE() AS DATE)")
Dim apAmount : apAmount = GetScalar("SELECT ISNULL(SUM(Amount-PaidAmount),0) FROM AccountsPayable WHERE Status IN ('Pending','Partial')")

' 应付列表
Dim whereSQL : whereSQL = "1=1"
If filterStatus <> "" Then whereSQL = "Status='" & Replace(filterStatus,"'","''") & "'"

Dim rsAP, apCount : apCount = 0
On Error Resume Next
Set rsAP = Server.CreateObject("ADODB.Recordset")
If Err.Number = 0 Then
    rsAP.CursorLocation = 3 ' adUseClient - 支持 MoveLast/RecordCount
    rsAP.Open "SELECT ap.*, po.PurchaseNo FROM AccountsPayable ap LEFT JOIN PurchaseOrders po ON ap.PurchaseID=po.PurchaseID WHERE " & whereSQL & " ORDER BY ap.DueDate ASC, ap.CreatedAt DESC", conn, 1, 1
End If
If Err.Number <> 0 Then Set rsAP = Nothing : Err.Clear
On Error GoTo 0
Dim apList() : ReDim apList(0, 10)
If Not rsAP Is Nothing Then
    If Not rsAP.EOF Then
        rsAP.MoveLast : apCount = rsAP.RecordCount : rsAP.MoveFirst
        ReDim apList(apCount - 1, 10)
        Dim api : api = 0
        Do While Not rsAP.EOF
            apList(api, 0) = rsAP("PayableID")
            apList(api, 1) = rsAP("SupplierName") & ""
            apList(api, 2) = rsAP("PayableNo") & ""
            apList(api, 3) = rsAP("Amount")
            apList(api, 4) = rsAP("PaidAmount")
            apList(api, 5) = rsAP("Status") & ""
            apList(api, 6) = rsAP("DueDate") & ""
            apList(api, 7) = rsAP("InvoiceNo") & ""
            apList(api, 8) = rsAP("Notes") & ""
            apList(api, 9) = rsAP("PurchaseNo") & ""
            apList(api, 10) = rsAP("CreatedAt") & ""
            api = api + 1
            rsAP.MoveNext
        Loop
    End If
    rsAP.Close : Set rsAP = Nothing
End If

' 应付明细
Dim detailID : detailID = SafeNum(Request.QueryString("id"))
Dim rsDetail, detailInfo
If detailID > 0 Then
On Error Resume Next
Set rsDetail = conn.Execute("SELECT ap.*, po.PurchaseNo, pr.ReceiptNo, pr.ReceiptDate FROM AccountsPayable ap LEFT JOIN PurchaseOrders po ON ap.PurchaseID=po.PurchaseID LEFT JOIN PurchaseReceipts pr ON ap.PurchaseID=pr.PurchaseID WHERE ap.PayableID=" & detailID)
If Err.Number <> 0 Then Set rsDetail = Nothing : Err.Clear
On Error GoTo 0
    If Not rsDetail Is Nothing And Not rsDetail.EOF Then
        Set detailInfo = rsDetail
    End If
End If

Function BadgeClass(st)
    Select Case st
        Case "Pending" : BadgeClass = "status-pending"
        Case "Partial" : BadgeClass = "status-partial"
        Case "Paid" : BadgeClass = "status-paid"
        Case "Overdue" : BadgeClass = "status-overdue"
        Case Else : BadgeClass = "status-draft"
    End Select
End Function
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>应付账款 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; font-family: 'Segoe UI',Arial,sans-serif; }
        .main-content { margin-left: 250px; padding: 30px; min-height: 100vh; }
        .page-header { margin-bottom: 25px; }
        .page-title { color: #fff; font-size: 24px; margin: 0 0 8px; }
        .breadcrumb { color: #888; font-size: 13px; }
        .breadcrumb a { color: #00bcd4; text-decoration: none; }
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); padding: 20px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); text-align: center; }
        .stat-card .num { font-size: 28px; font-weight: 700; }
        .stat-card .label { font-size: 13px; color: #888; margin-top: 5px; }
        .stat-card.total { border-top: 4px solid #00bcd4; } .stat-card.total .num { color: #00bcd4; }
        .stat-card.pending { border-top: 4px solid #FF9800; } .stat-card.pending .num { color: #FF9800; }
        .stat-card.overdue { border-top: 4px solid #f44336; } .stat-card.overdue .num { color: #f44336; }
        .stat-card.amount { border-top: 4px solid #4CAF50; } .stat-card.amount .num { color: #4CAF50; }
        .tabs { display: flex; gap: 8px; margin-bottom: 20px; flex-wrap: wrap; }
        .tab { padding: 10px 20px; background: #2d2d44; border-radius: 8px; color: #888; text-decoration: none; font-size: 14px; transition: all 0.2s; }
        .tab:hover, .tab.active { background: linear-gradient(135deg, #00bcd4, #00838f); color: #fff; }
        .card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); margin-bottom: 20px; overflow: hidden; }
        .card-header { padding: 15px 20px; border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; }
        .card-body { padding: 20px; }
        table { width: 100%; border-collapse: collapse; }
        th { text-align: left; padding: 10px 12px; background: rgba(0,188,212,0.06); border-bottom: 1px solid rgba(255,255,255,0.06); font-size: 13px; color: #999; }
        td { padding: 10px 12px; border-bottom: 1px solid rgba(255,255,255,0.04); font-size: 14px; }
        tr:hover { background: rgba(255,255,255,0.02); }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; }
        .status-pending { background: rgba(255,152,0,0.2); color: #FFB74D; }
        .status-partial { background: rgba(33,150,243,0.2); color: #64B5F6; }
        .status-paid { background: rgba(76,175,80,0.2); color: #81C784; }
        .status-overdue { background: rgba(244,67,54,0.2); color: #EF9A9A; }
        .amount { font-weight: 600; }
        .amount.positive { color: #4CAF50; }
        .amount.negative { color: #f44336; }
        .detail-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }
        .detail-item label { display: block; color: #888; font-size: 12px; margin-bottom: 3px; }
        .detail-item .value { font-size: 15px; font-weight: 500; }
        @media (max-width: 768px) { .main-content { margin-left: 0; } .stats-grid { grid-template-columns: 1fr 1fr; } .detail-grid { grid-template-columns: 1fr; } }
    </style>
</head>
<body data-theme="operation-dark">
<!--#include file="includes/nav.asp"-->
<div class="main-content">
    <div class="page-header">
        <h2 class="page-title"><i class="fas fa-file-invoice-dollar"></i> 应付账款</h2>
        <div class="breadcrumb"><a href="index.asp">财务中心</a> / 应付账款</div>
    </div>
    
    <div class="stats-grid">
        <div class="stat-card total"><div class="num"><%= apTotal %></div><div class="label">应付总计</div></div>
        <div class="stat-card pending"><div class="num"><%= apPending %></div><div class="label">待付款</div></div>
        <div class="stat-card overdue"><div class="num"><%= apOverdue %></div><div class="label">已逾期</div></div>
        <div class="stat-card amount"><div class="num">¥<%= FormatNumber(SafeNum(apAmount), 0) %></div><div class="label">未付金额</div></div>
    </div>
    
    <div class="tabs">
        <a href="accounts_payable.asp" class="tab <%= IIf(filterStatus = "", "active", "") %>">全部</a>
        <a href="?status=Pending" class="tab <%= IIf(filterStatus = "Pending", "active", "") %>">待付款</a>
        <a href="?status=Partial" class="tab <%= IIf(filterStatus = "Partial", "active", "") %>">部分付款</a>
        <a href="?status=Paid" class="tab <%= IIf(filterStatus = "Paid", "active", "") %>">已付清</a>
    </div>

    <% If IsObject(detailInfo) Then %>
    <div class="card">
        <div class="card-header">应付详情 #<%= detailInfo("PayableNo") %></div>
        <div class="card-body">
            <div class="detail-grid">
                <div class="detail-item"><label>供应商</label><div class="value"><%= detailInfo("SupplierName") %></div></div>
                <div class="detail-item"><label>采购单号</label><div class="value"><%= detailInfo("PurchaseNo") & "" %></div></div>
                <div class="detail-item"><label>应付金额</label><div class="value amount positive">¥<%= FormatNumber(SafeNum(detailInfo("Amount")), 2) %></div></div>
                <div class="detail-item"><label>已付金额</label><div class="value amount">¥<%= FormatNumber(SafeNum(detailInfo("PaidAmount")), 2) %></div></div>
                <div class="detail-item"><label>到期日</label><div class="value"><%= detailInfo("DueDate") & "" %></div></div>
                <div class="detail-item"><label>状态</label><div class="value"><span class="badge <%= BadgeClass(detailInfo("Status") & "") %>"><%= detailInfo("Status") %></span></div></div>
                <div class="detail-item"><label>发票号</label><div class="value"><%= detailInfo("InvoiceNo") & "" %></div></div>
                <div class="detail-item"><label>收货单号</label><div class="value"><%= detailInfo("ReceiptNo") & "" %></div></div>
            </div>
            <div style="margin-top:20px"><a href="accounts_payable.asp" class="tab"><i class="fas fa-arrow-left"></i> 返回列表</a></div>
        </div>
    </div>
    <% End If %>

    <div class="card">
        <div class="card-header"><i class="fas fa-list"></i> 应付账款列表</div>
        <div class="card-body">
            <table>
                <thead><tr><th>编号</th><th>供应商</th><th>采购单</th><th>应付金额</th><th>已付金额</th><th>未付余额</th><th>到期日</th><th>状态</th><th>操作</th></tr></thead>
                <tbody>
                    <% If apCount > 0 Then
                        For api = 0 To apCount - 1
                            Dim apBal : apBal = SafeNum(apList(api, 3)) - SafeNum(apList(api, 4))
                    %>
                    <tr>
                        <td><strong><%= apList(api, 2) %></strong></td>
                        <td><%= apList(api, 1) %></td>
                        <td><%= apList(api, 9) %></td>
                        <td class="amount">¥<%= FormatNumber(SafeNum(apList(api, 3)), 2) %></td>
                        <td class="amount positive">¥<%= FormatNumber(SafeNum(apList(api, 4)), 2) %></td>
                        <td class="amount <%= IIf(apBal > 0, "negative", "") %>">¥<%= FormatNumber(SafeNum(apBal), 2) %></td>
                        <td><%= apList(api, 6) %></td>
                        <td><span class="badge <%= BadgeClass(apList(api, 5)) %>"><%= apList(api, 5) %></span></td>
                        <td><a href="?id=<%= apList(api, 0) %>" class="tab" style="font-size:12px;padding:5px 10px;">详情</a></td>
                    </tr>
                    <% Next
                    Else %>
                    <tr><td colspan="9" style="text-align:center;padding:40px;color:#666;">暂无应付账款数据</td></tr>
                    <% End If %>
                </tbody>
            </table>
        </div>
    </div>
</div>
</body>
</html>
<%
If IsObject(detailInfo) Then detailInfo.Close : Set detailInfo = Nothing
Call CloseConnection()
%>
