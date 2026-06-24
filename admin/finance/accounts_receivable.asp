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
Function TableExists_AR(tblName)
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
    TableExists_AR = exists
End Function

If Not TableExists_AR("AccountsReceivable") Then
    Response.Write "<div style='padding:40px;color:#f44336;background:#1a1a2e;font-family:sans-serif;'><h2>表缺失</h2><p>AccountsReceivable 表不存在，请先运行 <a href='/setup/deploy.asp' style='color:#00bcd4;'>系统部署工具</a> 创建数据库表。</p></div>"
    Response.End
End If

Dim filterStatus : filterStatus = Request.QueryString("status")

Dim arTotal : arTotal = GetScalar("SELECT COUNT(*) FROM AccountsReceivable")
Dim arPending : arPending = GetScalar("SELECT COUNT(*) FROM AccountsReceivable WHERE Status='Pending'")
Dim arOverdue : arOverdue = GetScalar("SELECT COUNT(*) FROM AccountsReceivable WHERE Status='Pending' AND DueDate < CAST(GETDATE() AS DATE)")
Dim arAmount : arAmount = GetScalar("SELECT ISNULL(SUM(Amount-ReceivedAmount),0) FROM AccountsReceivable WHERE Status IN ('Pending','Partial')")

Dim whereSQL : whereSQL = "1=1"
If filterStatus <> "" Then whereSQL = "Status='" & Replace(filterStatus,"'","''") & "'"

Dim rsAR, arCount : arCount = 0
On Error Resume Next
Set rsAR = Server.CreateObject("ADODB.Recordset")
If Err.Number = 0 Then
    rsAR.CursorLocation = 3 ' adUseClient - 支持 MoveLast/RecordCount
    rsAR.Open "SELECT ar.*, o.OrderNo FROM AccountsReceivable ar LEFT JOIN Orders o ON ar.OrderID=o.OrderID WHERE " & whereSQL & " ORDER BY ar.DueDate ASC, ar.CreatedAt DESC", conn, 1, 1
End If
If Err.Number <> 0 Then Set rsAR = Nothing : Err.Clear
On Error GoTo 0
Dim arList() : ReDim arList(0, 10)
If Not rsAR Is Nothing Then
    If Not rsAR.EOF Then
        rsAR.MoveLast : arCount = rsAR.RecordCount : rsAR.MoveFirst
        ReDim arList(arCount - 1, 10)
        Dim ari : ari = 0
        Do While Not rsAR.EOF
            arList(ari, 0) = rsAR("ReceivableID")
            arList(ari, 1) = rsAR("CustomerName") & ""
            arList(ari, 2) = rsAR("ReceivableNo") & ""
            arList(ari, 3) = rsAR("Amount")
            arList(ari, 4) = rsAR("ReceivedAmount")
            arList(ari, 5) = rsAR("Status") & ""
            arList(ari, 6) = rsAR("DueDate") & ""
            arList(ari, 7) = rsAR("Notes") & ""
            arList(ari, 8) = rsAR("OrderNo") & ""
            arList(ari, 9) = rsAR("CreatedAt") & ""
            ari = ari + 1
            rsAR.MoveNext
        Loop
    End If
    rsAR.Close : Set rsAR = Nothing
End If

Function BadgeClass(st)
    Select Case st
        Case "Pending" : BadgeClass = "status-pending"
        Case "Partial" : BadgeClass = "status-partial"
        Case "Received" : BadgeClass = "status-paid"
        Case "Overdue" : BadgeClass = "status-overdue"
        Case Else : BadgeClass = "status-draft"
    End Select
End Function
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>应收账款 - 财务管理中心</title>
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
        .status-received { background: rgba(76,175,80,0.2); color: #81C784; }
        .status-overdue { background: rgba(244,67,54,0.2); color: #EF9A9A; }
        .amount { font-weight: 600; }
        .amount.positive { color: #4CAF50; }
        .amount.negative { color: #f44336; }
        @media (max-width: 768px) { .main-content { margin-left: 0; } .stats-grid { grid-template-columns: 1fr 1fr; } }
    </style>
</head>
<body data-theme="operation-dark">
<!--#include file="includes/nav.asp"-->
<div class="main-content">
    <div class="page-header">
        <h2 class="page-title"><i class="fas fa-hand-holding-usd"></i> 应收账款</h2>
        <div class="breadcrumb"><a href="index.asp">财务中心</a> / 应收账款</div>
    </div>
    
    <div class="stats-grid">
        <div class="stat-card total"><div class="num"><%= arTotal %></div><div class="label">应收总计</div></div>
        <div class="stat-card pending"><div class="num"><%= arPending %></div><div class="label">待收款</div></div>
        <div class="stat-card overdue"><div class="num"><%= arOverdue %></div><div class="label">已逾期</div></div>
        <div class="stat-card amount"><div class="num">¥<%= FormatNumber(SafeNum(arAmount), 0) %></div><div class="label">未收金额</div></div>
    </div>
    
    <div class="tabs">
        <a href="accounts_receivable.asp" class="tab <%= IIf(filterStatus = "", "active", "") %>">全部</a>
        <a href="?status=Pending" class="tab <%= IIf(filterStatus = "Pending", "active", "") %>">待收款</a>
        <a href="?status=Partial" class="tab <%= IIf(filterStatus = "Partial", "active", "") %>">部分收款</a>
        <a href="?status=Received" class="tab <%= IIf(filterStatus = "Received", "active", "") %>">已收讫</a>
    </div>
    
    <div class="card">
        <div class="card-header"><i class="fas fa-list"></i> 应收账款列表</div>
        <div class="card-body">
            <table>
                <thead><tr><th>编号</th><th>客户</th><th>关联订单</th><th>应收金额</th><th>已收金额</th><th>未收余额</th><th>到期日</th><th>状态</th><th>备注</th></tr></thead>
                <tbody>
                    <% If arCount > 0 Then
                        For ari = 0 To arCount - 1
                            Dim arBal : arBal = SafeNum(arList(ari, 3)) - SafeNum(arList(ari, 4))
                            Dim arSt : arSt = arList(ari, 5)
                    %>
                    <tr>
                        <td><strong><%= arList(ari, 2) %></strong></td>
                        <td><%= arList(ari, 1) %></td>
                        <td><%= arList(ari, 8) %></td>
                        <td class="amount">¥<%= FormatNumber(SafeNum(arList(ari, 3)), 2) %></td>
                        <td class="amount positive">¥<%= FormatNumber(SafeNum(arList(ari, 4)), 2) %></td>
                        <td class="amount <%= IIf(arBal > 0, "negative", "") %>">¥<%= FormatNumber(SafeNum(arBal), 2) %></td>
                        <td><%= arList(ari, 6) %></td>
                        <td><span class="badge <%= BadgeClass(arSt) %>"><%= arSt %></span></td>
                        <td><%= Left(arList(ari, 7), 20) %></td>
                    </tr>
                    <% Next
                    Else %>
                    <tr><td colspan="9" style="text-align:center;padding:40px;color:#666;">暂无应收账款数据</td></tr>
                    <% End If %>
                </tbody>
            </table>
        </div>
    </div>
</div>
</body>
</html>
<%
Call CloseConnection()
%>
