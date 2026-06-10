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
conn.Execute "SELECT TOP 1 1 FROM GLTransactions"
If Err.Number <> 0 Then Err.Clear : conn.Execute "CREATE TABLE GLTransactions (GLID INT IDENTITY(1,1) PRIMARY KEY, GLNo NVARCHAR(50), TransactionDate DATETIME2 DEFAULT GETDATE(), AccountCode NVARCHAR(50), AccountName NVARCHAR(200), DebitAmount DECIMAL(19,4) DEFAULT 0, CreditAmount DECIMAL(19,4) DEFAULT 0, Balance DECIMAL(19,4) DEFAULT 0, CenterID INT NULL, RefType NVARCHAR(30), RefID INT NULL, RefNo NVARCHAR(100), Description NVARCHAR(500), CreatedBy NVARCHAR(50), CreatedAt DATETIME2 DEFAULT GETDATE())"
On Error GoTo 0

' 确保 GLTransactions 表有所需列（兼容旧表）
On Error Resume Next
conn.Execute "SELECT DebitAmount FROM GLTransactions WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE GLTransactions ADD DebitAmount DECIMAL(19,4) DEFAULT 0"
conn.Execute "SELECT CreditAmount FROM GLTransactions WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE GLTransactions ADD CreditAmount DECIMAL(19,4) DEFAULT 0"
conn.Execute "SELECT AccountCode FROM GLTransactions WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE GLTransactions ADD AccountCode NVARCHAR(50)"
conn.Execute "SELECT AccountName FROM GLTransactions WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE GLTransactions ADD AccountName NVARCHAR(200)"
On Error GoTo 0

Dim period : period = Request.QueryString("period")
If period = "" Then period = Year(Now) & "-" & Right("0"&Month(Now),2)

Dim whereDate : whereDate = "1=1"
If period <> "all" Then whereDate = "LEFT(CONVERT(VARCHAR,TransactionDate,23),7)='" & Replace(period,"'","''") & "'"

' 总账汇总
Dim totalDebit : totalDebit = GetScalar("SELECT ISNULL(SUM(DebitAmount),0) FROM GLTransactions WHERE " & whereDate)
Dim totalCredit : totalCredit = GetScalar("SELECT ISNULL(SUM(CreditAmount),0) FROM GLTransactions WHERE " & whereDate)
Dim glCount : glCount = GetScalar("SELECT COUNT(*) FROM GLTransactions WHERE " & whereDate)

' 按科目汇总
Dim rsSummary, sumCount : sumCount = 0
On Error Resume Next
Set rsSummary = conn.Execute("SELECT AccountCode, AccountName, SUM(DebitAmount) AS TotalDebit, SUM(CreditAmount) AS TotalCredit, COUNT(*) AS TxCount FROM GLTransactions WHERE " & whereDate & " GROUP BY AccountCode, AccountName ORDER BY AccountCode")
If Err.Number <> 0 Then Set rsSummary = Nothing : Err.Clear
On Error GoTo 0
Dim sumRows
If Not rsSummary Is Nothing Then
    If Not rsSummary.EOF Then
        sumRows = rsSummary.GetRows()
        sumCount = UBound(sumRows, 2) + 1
        ReDim sumList(sumCount - 1, 4)
        Dim si : si = 0
        For si = 0 To sumCount - 1
            sumList(si, 0) = sumRows(0, si) & ""
            sumList(si, 1) = sumRows(1, si) & ""
            sumList(si, 2) = sumRows(2, si)
            sumList(si, 3) = sumRows(3, si)
            sumList(si, 4) = sumRows(4, si)
        Next
    End If
    rsSummary.Close : Set rsSummary = Nothing
End If

' 最近流水
Dim rsGL, glRows, glList(), glLCount : glLCount = 0
On Error Resume Next
Set rsGL = conn.Execute("SELECT TOP 50 * FROM GLTransactions ORDER BY TransactionDate DESC, GLID DESC")
If Err.Number <> 0 Then Set rsGL = Nothing : Err.Clear
On Error GoTo 0
If Not rsGL Is Nothing Then
    If Not rsGL.EOF Then
        glRows = rsGL.GetRows()
        glLCount = UBound(glRows, 2) + 1
        ReDim glList(glLCount - 1, 6)
        Dim gli : gli = 0
        For gli = 0 To glLCount - 1
            glList(gli, 0) = glRows(1, gli) & ""
            glList(gli, 1) = glRows(2, gli) & ""
            glList(gli, 2) = glRows(4, gli) & ""
            glList(gli, 3) = glRows(5, gli)
            glList(gli, 4) = glRows(6, gli)
            glList(gli, 5) = glRows(12, gli) & ""
            glList(gli, 6) = glRows(13, gli) & ""
        Next
    End If
    rsGL.Close : Set rsGL = Nothing
End If
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>总账报表 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
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
        .stat-card .num { font-size: 28px; font-weight: 700; color: #00bcd4; }
        .stat-card .label { font-size: 13px; color: #888; margin-top: 5px; }
        .filter-bar { display: flex; gap: 10px; margin-bottom: 20px; align-items: center; }
        .filter-bar select, .filter-bar input { padding: 10px 15px; background: #2d2d44; border: 1px solid rgba(255,255,255,0.1); border-radius: 8px; color: #e0e0e0; }
        .filter-bar button { padding: 10px 20px; background: linear-gradient(135deg, #00bcd4, #00838f); border: none; border-radius: 8px; color: #fff; cursor: pointer; }
        .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        .card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); margin-bottom: 20px; overflow: hidden; }
        .card-header { padding: 15px 20px; border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; }
        .card-body { padding: 20px; }
        table { width: 100%; border-collapse: collapse; }
        th { text-align: left; padding: 10px 12px; background: rgba(0,188,212,0.06); border-bottom: 1px solid rgba(255,255,255,0.06); font-size: 13px; color: #999; }
        td { padding: 10px 12px; border-bottom: 1px solid rgba(255,255,255,0.04); font-size: 14px; }
        tr:hover { background: rgba(255,255,255,0.02); }
        .amount { font-weight: 600; }
        .amount.debit { color: #4CAF50; }
        .amount.credit { color: #f44336; }
        .text-right { text-align: right; }
        @media (max-width: 768px) { .main-content { margin-left: 0; } .grid-2 { grid-template-columns: 1fr; } .stats-grid { grid-template-columns: 1fr 1fr; } }
    </style>
</head>
<body>
<!--#include file="includes/nav.asp"-->
<div class="main-content">
    <div class="page-header">
        <h2 class="page-title"><i class="fas fa-book"></i> 总账报表</h2>
        <div class="breadcrumb"><a href="index.asp">财务中心</a> / 总账报表</div>
    </div>
    
    <div class="stats-grid">
        <div class="stat-card"><div class="num"><%= glCount %></div><div class="label">分录总数</div></div>
        <div class="stat-card"><div class="num" style="color:#4CAF50;">¥<%= FormatNumber(SafeNum(totalDebit), 0) %></div><div class="label">借方合计</div></div>
        <div class="stat-card"><div class="num" style="color:#f44336;">¥<%= FormatNumber(SafeNum(totalCredit), 0) %></div><div class="label">贷方合计</div></div>
        <div class="stat-card"><div class="num" style="color:#00bcd4;">¥<%= FormatNumber(SafeNum(totalDebit) - SafeNum(totalCredit), 0) %></div><div class="label">借贷差额</div></div>
    </div>
    
    <div class="filter-bar">
        <form method="get" style="display:flex;gap:10px;align-items:center;">
            <label style="color:#888;">期间：</label>
            <select name="period" onchange="this.form.submit()">
                <option value="all" <%= IIf(period="all","selected","") %>>全部</option>
                <%
                Dim ym, ymLabel
                For ym = 0 To 11
                    ymLabel = DateAdd("m", -ym, Now)
                    ymLabel = Year(ymLabel) & "-" & Right("0"&Month(ymLabel),2)
                %>
                <option value="<%= ymLabel %>" <%= IIf(period=ymLabel,"selected","") %>><%= ymLabel %></option>
                <% Next %>
            </select>
        </form>
    </div>
    
    <div class="grid-2">
        <!-- 科目汇总 -->
        <div class="card">
            <div class="card-header"><i class="fas fa-layer-group"></i> 科目汇总</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>科目代码</th><th>科目名称</th><th>借方</th><th>贷方</th><th>笔数</th></tr></thead>
                    <tbody>
                        <% If sumCount > 0 Then
                            For si = 0 To sumCount - 1
                        %>
                        <tr>
                            <td><strong><%= sumList(si, 0) %></strong></td>
                            <td><%= sumList(si, 1) %></td>
                            <td class="amount debit text-right">¥<%= FormatNumber(SafeNum(sumList(si, 2)), 2) %></td>
                            <td class="amount credit text-right">¥<%= FormatNumber(SafeNum(sumList(si, 3)), 2) %></td>
                            <td class="text-right"><%= sumList(si, 4) %></td>
                        </tr>
                        <% Next
                        Else %>
                        <tr><td colspan="5" style="text-align:center;padding:40px;color:#666;">该期间暂无总账数据</td></tr>
                        <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
        
        <!-- 最近流水 -->
        <div class="card">
            <div class="card-header"><i class="fas fa-stream"></i> 最近流水</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>凭证号</th><th>日期</th><th>科目</th><th>借方</th><th>贷方</th></tr></thead>
                    <tbody>
                        <% If glLCount > 0 Then
                            For gli = 0 To glLCount - 1
                        %>
                        <tr>
                            <td><strong><%= glList(gli, 0) %></strong></td>
                            <td><%= glList(gli, 1) %></td>
                            <td><%= glList(gli, 2) %></td>
                            <td class="amount debit text-right">¥<%= FormatNumber(SafeNum(glList(gli, 3)), 2) %></td>
                            <td class="amount credit text-right">¥<%= FormatNumber(SafeNum(glList(gli, 4)), 2) %></td>
                        </tr>
                        <% Next
                        Else %>
                        <tr><td colspan="5" style="text-align:center;padding:40px;color:#666;">暂无流水数据</td></tr>
                        <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</div>
</body>
</html>
<%
Call CloseConnection()
%>
