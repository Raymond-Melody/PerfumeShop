<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
' 安全数值转换函数
Function SafeNum(val)
    If IsNull(val) Or IsEmpty(val) Or val = "" Then
        SafeNum = 0
    Else
        On Error Resume Next
        SafeNum = CDbl(val)
        If Err.Number <> 0 Then
            SafeNum = 0
            Err.Clear
        End If
        On Error GoTo 0
    End If
End Function

Call OpenConnection()

' 获取日期范围
Dim startDate, endDate
startDate = Request.QueryString("startDate")
endDate = Request.QueryString("endDate")
If startDate = "" Then startDate = SafeFormatDateTime(DateAdd("d", -30, Date()), 2)
If endDate = "" Then endDate = SafeFormatDateTime(Date(), 2)

' 获取统计数据
Dim totalRevenue, totalOrders, avgOrderValue
totalRevenue = GetScalar("SELECT CAST(IIF(SUM(TotalAmount) IS NULL, 0, SUM(TotalAmount)) AS FLOAT) FROM Orders WHERE Status = 'Paid' AND OrderDate BETWEEN #" & startDate & "# AND #" & endDate & "#")
totalOrders = GetScalar("SELECT COUNT(*) FROM Orders WHERE Status = 'Paid' AND OrderDate BETWEEN #" & startDate & "# AND #" & endDate & "'")
If CDbl("0" & totalOrders) > 0 Then avgOrderValue = CDbl("0" & totalRevenue) / CDbl("0" & totalOrders) Else avgOrderValue = 0

' 获取按支付方式统计
Dim rsPaymentStats
Set rsPaymentStats = ExecuteQuery(_
    "SELECT PaymentMethod, COUNT(*) as OrderCount, SUM(CAST(TotalAmount AS FLOAT)) as Amount " & _
    "FROM Orders WHERE Status = 'Paid' AND OrderDate BETWEEN #" & startDate & "# AND #" & endDate & "# " & _
    "GROUP BY PaymentMethod")

' 获取按日统计
Dim rsDailyStats, payIcon, percent
Set rsDailyStats = ExecuteQuery(_
    "SELECT OrderDate, COUNT(*) as OrderCount, SUM(CAST(TotalAmount AS FLOAT)) as Amount " & _
    "FROM Orders WHERE Status = 'Paid' AND OrderDate BETWEEN #" & startDate & "# AND #" & endDate & "# " & _
    "GROUP BY OrderDate ORDER BY OrderDate DESC")

Call LogAdminAction("查看收入统计", "finance", "Orders", "", startDate & " 至 " & endDate)
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>收入统计 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; }
        .stats-cards { display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 30px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.3); text-align: center; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card i { font-size: 40px; color: #00bcd4; margin-bottom: 15px; }
        .stat-card h3 { font-size: 36px; margin: 10px 0; color: #e0e0e0; }
        .stat-card p { color: #888; margin: 0; font-size: 14px; }
        .filter-bar { background: #2d2d44; padding: 20px; border-radius: 12px; margin-bottom: 25px; box-shadow: 0 4px 15px rgba(0,0,0,0.3); border: 1px solid rgba(255,255,255,0.06); }
        .filter-bar form { display: flex; gap: 15px; align-items: center; }
        .filter-bar input { padding: 10px 15px; border: 2px solid rgba(255,255,255,0.15); border-radius: 8px; background: #2d2d44; color: #e0e0e0; }
        .filter-bar input:focus { border-color: #00bcd4; outline: none; }
        .charts-container { display: grid; grid-template-columns: 1fr 1fr; gap: 25px; margin-bottom: 25px; }
        .chart-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 25px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.3); border: 1px solid rgba(255,255,255,0.06); }
        .chart-card h3 { margin-bottom: 20px; color: #e0e0e0; }
        .data-table { width: 100%; background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.3); }
        .data-table th { background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); color: white; padding: 15px; text-align: left; }
        .data-table td { padding: 12px 15px; border-bottom: 1px solid rgba(255,255,255,0.06); color: #e0e0e0; }
        .data-table tr:hover { background: rgba(255,255,255,0.05); }
        .progress-bar { height: 20px; background: rgba(255,255,255,0.1); border-radius: 10px; overflow: hidden; }
        .progress-fill { height: 100%; background: linear-gradient(90deg, #00bcd4, #00838f); border-radius: 10px; }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-chart-bar"></i> 收入统计</h2>
            <div class="breadcrumb">
                <a href="index.asp">财务中心</a> / <span>收入统计</span>
            </div>
        </div>
        
        <div class="filter-bar">
            <form method="get" action="revenue.asp">
                <label>开始日期:</label>
                <input type="date" name="startDate" value="<%= startDate %>">
                <label>结束日期:</label>
                <input type="date" name="endDate" value="<%= endDate %>">
                <button type="submit" class="admin-btn admin-btn-primary"><i class="fas fa-search"></i> 查询</button>
            </form>
        </div>
        
        <div class="stats-cards">
            <div class="stat-card">
                <i class="fas fa-yen-sign"></i>
                <h3>¥<%= FormatNumber(CDbl("0" & totalRevenue), 2) %></h3>
                <p>总收入</p>
            </div>
            <div class="stat-card">
                <i class="fas fa-shopping-cart"></i>
                <h3><%= totalOrders %></h3>
                <p>订单数</p>
            </div>
            <div class="stat-card">
                <i class="fas fa-receipt"></i>
                <h3>¥<%= FormatNumber(CDbl("0" & avgOrderValue), 2) %></h3>
                <p>客单价</p>
            </div>
        </div>
        
        <div class="charts-container">
            <div class="chart-card">
                <h3><i class="fas fa-credit-card"></i> 按支付方式统计</h3>
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>支付方式</th>
                            <th>订单数</th>
                            <th>金额</th>
                            <th>占比</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% If Not rsPaymentStats Is Nothing Then %>
                        <% Do While Not rsPaymentStats.EOF %>
                        <tr>
                            <td>
                                <% 
                                Select Case rsPaymentStats("PaymentMethod")
                                    Case "alipay": payIcon = "<i class='fab fa-alipay' style='color: #1677ff;'></i> 支付宝"
                                    Case "wechat": payIcon = "<i class='fab fa-weixin' style='color: #07c160;'></i> 微信支付"
                                    Case "bank": payIcon = "<i class='fas fa-university' style='color: #ff6b6b;'></i> 银行转账"
                                    Case Else: payIcon = rsPaymentStats("PaymentMethod")
                                End Select
                                Response.Write payIcon
                                %>
                            </td>
                            <td><%= rsPaymentStats("OrderCount") %></td>
                            <td>¥<%= FormatNumber(CDbl("0" & rsPaymentStats("Amount")), 2) %></td>
                            <td>
                                <% 
                                If totalRevenue > 0 Then percent = (CDbl("0" & rsPaymentStats("Amount")) / totalRevenue) * 100 Else percent = 0
                                %>
                                <div class="progress-bar" style="width: 100px; display: inline-block; vertical-align: middle;">
                                    <div class="progress-fill" style="width: <%= percent %>%"></div>
                                </div>
                                <span style="margin-left: 10px;"><%= FormatNumber(CDbl("0" & percent), 1) %>%</span>
                            </td>
                        </tr>
                        <% rsPaymentStats.MoveNext %>
                        <% Loop %>
                        <% rsPaymentStats.Close %>
                        <% End If %>
                    </tbody>
                </table>
            </div>
            
            <div class="chart-card">
                <h3><i class="fas fa-calendar-alt"></i> 按日统计</h3>
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>日期</th>
                            <th>订单数</th>
                            <th>金额</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% If Not rsDailyStats Is Nothing Then %>
                        <% Do While Not rsDailyStats.EOF %>
                        <tr>
                            <td><%= SafeFormatDateTime(rsDailyStats("OrderDate"), 2) %></td>
                            <td><%= rsDailyStats("OrderCount") %></td>
                            <td>¥<%= FormatNumber(CDbl("0" & rsDailyStats("Amount")), 2) %></td>
                        </tr>
                        <% rsDailyStats.MoveNext %>
                        <% Loop %>
                        <% rsDailyStats.Close %>
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
