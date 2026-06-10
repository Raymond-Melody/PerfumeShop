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

' 确保 PaymentRecords 有 PaymentType 列
On Error Resume Next
conn.Execute "SELECT PaymentType FROM PaymentRecords WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PaymentRecords ADD PaymentType NVARCHAR(30) DEFAULT ''Receipt''"
conn.Execute "SELECT PayableID FROM PaymentRecords WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PaymentRecords ADD PayableID INT"
conn.Execute "SELECT ReceivableID FROM PaymentRecords WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PaymentRecords ADD ReceivableID INT"
On Error GoTo 0

' 资金概况
Dim cashIn : cashIn = GetScalar("SELECT ISNULL(SUM(Amount),0) FROM PaymentRecords WHERE PaymentType='Receipt' AND Status='Completed' AND CAST(CreatedAt AS DATE) >= DATEADD(DAY,-30,CAST(GETDATE() AS DATE))")
Dim cashOut : cashOut = GetScalar("SELECT ISNULL(SUM(Amount),0) FROM PaymentRecords WHERE PaymentType='Payment' AND Status='Completed' AND CAST(CreatedAt AS DATE) >= DATEADD(DAY,-30,CAST(GETDATE() AS DATE))")

' 应收应付余额
Dim apBalance : apBalance = GetScalar("SELECT ISNULL(SUM(Amount-PaidAmount),0) FROM AccountsPayable WHERE Status IN ('Pending','Partial')")
Dim arBalance : arBalance = GetScalar("SELECT ISNULL(SUM(Amount-ReceivedAmount),0) FROM AccountsReceivable WHERE Status IN ('Pending','Partial')")

' 到期应付(30天内)
Dim apDue30 : apDue30 = GetScalar("SELECT ISNULL(SUM(Amount-PaidAmount),0) FROM AccountsPayable WHERE Status IN ('Pending','Partial') AND DueDate >= CAST(GETDATE() AS DATE) AND DueDate <= DATEADD(DAY,30,CAST(GETDATE() AS DATE))")

' 到期应收(30天内)
Dim arDue30 : arDue30 = GetScalar("SELECT ISNULL(SUM(Amount-ReceivedAmount),0) FROM AccountsReceivable WHERE Status IN ('Pending','Partial') AND DueDate >= CAST(GETDATE() AS DATE) AND DueDate <= DATEADD(DAY,30,CAST(GETDATE() AS DATE))")

' 逾期应付
Dim apOverdue : apOverdue = GetScalar("SELECT ISNULL(SUM(Amount-PaidAmount),0) FROM AccountsPayable WHERE Status IN ('Pending','Partial') AND DueDate < CAST(GETDATE() AS DATE)")

' 逾期应收
Dim arOverdue : arOverdue = GetScalar("SELECT ISNULL(SUM(Amount-ReceivedAmount),0) FROM AccountsReceivable WHERE Status IN ('Pending','Partial') AND DueDate < CAST(GETDATE() AS DATE)")

' 资金账户余额
Dim fundBalance : fundBalance = GetScalar("SELECT ISNULL(SUM(TotalBalance),0) FROM FundAccounts WHERE IsActive=1")

' 净现金流
Dim netCashFlow : netCashFlow = SafeNum(cashIn) - SafeNum(cashOut)
Dim projectedCash : projectedCash = SafeNum(fundBalance) + SafeNum(arDue30) - SafeNum(apDue30)
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>现金流预测 - 财务管理中心</title>
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
        .green-num { color: #4CAF50; }
        .red-num { color: #f44336; }
        .blue-num { color: #2196F3; }
        .cyan-num { color: #00bcd4; }
        .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        .grid-3 { display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; }
        .card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); margin-bottom: 20px; overflow: hidden; }
        .card-header { padding: 15px 20px; border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; }
        .card-body { padding: 20px; }
        .flow-indicator { display: flex; align-items: center; gap: 20px; padding: 25px; background: linear-gradient(135deg, rgba(0,188,212,0.08), rgba(33,150,243,0.08)); border-radius: 12px; margin-bottom: 20px; }
        .flow-item { flex: 1; text-align: center; }
        .flow-item .value { font-size: 24px; font-weight: 700; }
        .flow-item .label { font-size: 12px; color: #888; margin-top: 4px; }
        .flow-arrow { font-size: 24px; color: #888; }
        .alert-box { padding: 15px; border-radius: 8px; margin-bottom: 15px; }
        .alert-warning { background: rgba(255,152,0,0.15); border: 1px solid rgba(255,152,0,0.3); color: #FFB74D; }
        .alert-danger { background: rgba(244,67,54,0.15); border: 1px solid rgba(244,67,54,0.3); color: #EF9A9A; }
        .alert-success { background: rgba(76,175,80,0.15); border: 1px solid rgba(76,175,80,0.3); color: #81C784; }
        @media (max-width: 768px) { .main-content { margin-left: 0; } .stats-grid, .grid-2, .grid-3 { grid-template-columns: 1fr; } }
    </style>
</head>
<body data-theme="operation-dark">
<!--#include file="includes/nav.asp"-->
<div class="main-content">
    <div class="page-header">
        <h2 class="page-title"><i class="fas fa-chart-line"></i> 现金流预测</h2>
        <div class="breadcrumb"><a href="index.asp">财务中心</a> / 现金流预测</div>
    </div>
    
    <!-- 资金总览仪表 -->
    <div class="stats-grid">
        <div class="stat-card" style="border-top:4px solid #00bcd4;"><div class="num cyan-num">¥<%= FormatNumber(SafeNum(fundBalance), 0) %></div><div class="label">资金账户余额</div></div>
        <div class="stat-card" style="border-top:4px solid #4CAF50;"><div class="num green-num">¥<%= FormatNumber(SafeNum(cashIn), 0) %></div><div class="label">近30天收入</div></div>
        <div class="stat-card" style="border-top:4px solid #f44336;"><div class="num red-num">¥<%= FormatNumber(SafeNum(cashOut), 0) %></div><div class="label">近30天支出</div></div>
        <div class="stat-card" style="border-top:4px solid <%= IIf(SafeNum(netCashFlow)>=0,"#4CAF50","#f44336") %>;"><div class="num <%= IIf(SafeNum(netCashFlow)>=0,"green-num","red-num") %>">¥<%= FormatNumber(Abs(SafeNum(netCashFlow)), 0) %></div><div class="label"><%= IIf(SafeNum(netCashFlow)>=0,"净流入","净流出") %></div></div>
    </div>
    
    <!-- 现金流预测管道 -->
    <div class="flow-indicator">
        <div class="flow-item"><div class="value green-num">¥<%= FormatNumber(SafeNum(fundBalance), 0) %></div><div class="label">当前余额</div></div>
        <div class="flow-arrow"><i class="fas fa-plus-circle" style="color:#4CAF50;"></i></div>
        <div class="flow-item"><div class="value green-num">¥<%= FormatNumber(SafeNum(arDue30), 0) %></div><div class="label">30天到期应收</div></div>
        <div class="flow-arrow"><i class="fas fa-minus-circle" style="color:#f44336;"></i></div>
        <div class="flow-item"><div class="value red-num">¥<%= FormatNumber(SafeNum(apDue30), 0) %></div><div class="label">30天到期应付</div></div>
        <div class="flow-arrow"><i class="fas fa-equals"></i></div>
        <div class="flow-item"><div class="value <%= IIf(SafeNum(projectedCash)>=0,"cyan-num","red-num") %>">¥<%= FormatNumber(SafeNum(projectedCash), 0) %></div><div class="label">预测余额</div></div>
    </div>
    
    <!-- 预警 -->
    <% If SafeNum(apOverdue) > 0 Then %>
    <div class="alert-danger"><i class="fas fa-exclamation-triangle"></i> 逾期应付：¥<%= FormatNumber(SafeNum(apOverdue), 0) %>，请尽快安排付款</div>
    <% End If %>
    <% If SafeNum(arOverdue) > 0 Then %>
    <div class="alert-warning"><i class="fas fa-exclamation-circle"></i> 逾期应收：¥<%= FormatNumber(SafeNum(arOverdue), 0) %>，请跟催收款</div>
    <% End If %>
    <% If SafeNum(projectedCash) < 0 Then %>
    <div class="alert-danger"><i class="fas fa-skull-crossbones"></i> 资金预警：30天预测余额为负，需采取措施</div>
    <% End If %>
    
    <div class="grid-3">
        <div class="card">
            <div class="card-header"><i class="fas fa-file-invoice-dollar"></i> 应付概况</div>
            <div class="card-body">
                <div style="margin-bottom:10px"><span style="color:#888">未付余额：</span><strong class="red-num">¥<%= FormatNumber(SafeNum(apBalance), 0) %></strong></div>
                <div style="margin-bottom:10px"><span style="color:#888">30天到期：</span><strong style="color:#FF9800;">¥<%= FormatNumber(SafeNum(apDue30), 0) %></strong></div>
                <div><span style="color:#888">已逾期：</span><strong class="red-num">¥<%= FormatNumber(SafeNum(apOverdue), 0) %></strong></div>
            </div>
        </div>
        <div class="card">
            <div class="card-header"><i class="fas fa-hand-holding-usd"></i> 应收概况</div>
            <div class="card-body">
                <div style="margin-bottom:10px"><span style="color:#888">未收余额：</span><strong class="green-num">¥<%= FormatNumber(SafeNum(arBalance), 0) %></strong></div>
                <div style="margin-bottom:10px"><span style="color:#888">30天到期：</span><strong style="color:#2196F3;">¥<%= FormatNumber(SafeNum(arDue30), 0) %></strong></div>
                <div><span style="color:#888">已逾期：</span><strong style="color:#FF9800;">¥<%= FormatNumber(SafeNum(arOverdue), 0) %></strong></div>
            </div>
        </div>
        <div class="card">
            <div class="card-header"><i class="fas fa-lightbulb"></i> 决策建议</div>
            <div class="card-body">
                <% If SafeNum(projectedCash) < 0 Then %>
                <p style="color:#EF9A9A;"><i class="fas fa-exclamation-triangle"></i> 预测资金缺口 ¥<%= FormatNumber(Abs(SafeNum(projectedCash)), 0) %>，建议：</p>
                <ul style="color:#b0b0b0;font-size:13px;padding-left:20px;line-height:1.8;">
                    <li>加速催收逾期应收</li>
                    <li>协商延期支付应付</li>
                    <li>考虑短期融资</li>
                </ul>
                <% ElseIf SafeNum(projectedCash) < SafeNum(fundBalance) * 0.3 Then %>
                <p style="color:#FFB74D;"><i class="fas fa-exclamation-circle"></i> 资金偏紧，建议：</p>
                <ul style="color:#b0b0b0;font-size:13px;padding-left:20px;line-height:1.8;">
                    <li>控制非必要支出</li>
                    <li>优化付款节奏</li>
                </ul>
                <% Else %>
                <p style="color:#81C784;"><i class="fas fa-check-circle"></i> 资金状况良好</p>
                <p style="color:#b0b0b0;font-size:13px;">预测余额可覆盖30天应付需求</p>
                <% End If %>
            </div>
        </div>
    </div>
    
    <div style="text-align:center;margin-top:20px;">
        <a href="fund_dashboard.asp" class="btn" style="padding:10px 25px;background:#2d2d44;color:#e0e0e0;border-radius:8px;text-decoration:none;">前往资金看板 <i class="fas fa-arrow-right"></i></a>
    </div>
</div>
</body>
</html>
<%
Call CloseConnection()
%>
