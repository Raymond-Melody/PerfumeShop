<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<!--#include file="../../includes/dal.asp"-->
<!--#include file="../../includes/dal_finance.asp"-->
<!--#include file="../../includes/audit_utils.asp"-->
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

' 日期范围处理
Dim startDate, endDate
startDate = Trim(Request.QueryString("startDate"))
endDate = Trim(Request.QueryString("endDate"))
If startDate = "" Then startDate = SafeFormatDateTime(DateAdd("m", -6, Date()), 2)
If endDate = "" Then endDate = SafeFormatDateTime(Date(), 2)

Dim safeStart, safeEnd
safeStart = SafeSQL(startDate)
safeEnd = SafeSQL(endDate)

' V17: 使用参数化DAL查询
Dim rsMonthly
Set rsMonthly = DAL_Fin_GetMonthlyRevenue(startDate, endDate)

' V17: 使用参数化DAL查询订单金额分布
Call DAL_Fin_GetOrderDistribution(startDate, endDate, cnt0_50, cnt50_100, cnt100_200, cnt200_500, cnt500plus)

Dim totalDistribution
totalDistribution = CLng("0" & cnt0_50) + CLng("0" & cnt50_100) + CLng("0" & cnt100_200) + CLng("0" & cnt200_500) + CLng("0" & cnt500plus)
If totalDistribution = 0 Then totalDistribution = 1

' V17: 使用参数化DAL查询支付方式统计
Dim rsPayment
Set rsPayment = DAL_Fin_GetPaymentStats(startDate, endDate)

' V17: 使用参数化DAL查询总收入
Dim grandTotal
grandTotal = DAL_Fin_GetTotalRevenue(startDate, endDate)

' V17: 替换为审计日志
Call AuditLog("view", "finance", 0, "Reports", safeStart & " 至 " & safeEnd)
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>财务报表 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; }
        .filter-bar { background: #2d2d44; padding: 20px; border-radius: 12px; margin-bottom: 25px; box-shadow: 0 4px 15px rgba(0,0,0,0.3); border: 1px solid rgba(255,255,255,0.06); }
        .filter-bar form { display: flex; gap: 15px; align-items: center; flex-wrap: wrap; }
        .filter-bar input { padding: 10px 15px; border: 2px solid rgba(255,255,255,0.15); border-radius: 8px; background: #2d2d44; color: #e0e0e0; }
        .filter-bar input:focus { border-color: #00bcd4; outline: none; }
        .report-section { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; padding: 25px; box-shadow: 0 4px 15px rgba(0,0,0,0.3); margin-bottom: 25px; border: 1px solid rgba(255,255,255,0.06); }
        .report-section h3 { font-size: 18px; color: #e0e0e0; margin-bottom: 20px; display: flex; align-items: center; gap: 10px; padding-bottom: 15px; border-bottom: 2px solid rgba(255,255,255,0.06); }
        .data-table { width: 100%; border-collapse: collapse; }
        .data-table th { background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); color: white; padding: 15px; text-align: left; }
        .data-table td { padding: 12px 15px; border-bottom: 1px solid rgba(255,255,255,0.06); color: #e0e0e0; }
        .data-table tr:hover { background: rgba(255,255,255,0.05); }
        .data-table tfoot td { font-weight: bold; background: #1e1e32; color: #e0e0e0; }
        .progress-bar { height: 20px; background: rgba(255,255,255,0.1); border-radius: 10px; overflow: hidden; display: inline-block; vertical-align: middle; }
        .progress-fill { height: 100%; border-radius: 10px; }
        .progress-fill.blue { background: linear-gradient(90deg, #00bcd4, #00838f); }
        .progress-fill.green { background: linear-gradient(90deg, #4CAF50, #388e3c); }
        .progress-fill.orange { background: linear-gradient(90deg, #ff9800, #f57c00); }
        .charts-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 25px; margin-bottom: 25px; }
        .no-data { text-align: center; padding: 40px; color: #888; }
        .no-data i { font-size: 48px; margin-bottom: 15px; display: block; color: #555; }
        .amount-highlight { font-weight: bold; color: #4CAF50; }
        .count-highlight { font-weight: bold; color: #00bcd4; }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-file-alt"></i> 财务报表中心</h2>
            <div class="breadcrumb">
                <a href="index.asp">财务中心</a> / <span>财务报表</span>
            </div>
        </div>
        
        <!-- 日期筛选 -->
        <div class="filter-bar">
            <form method="get" action="reports.asp">
                <label>开始日期:</label>
                <input type="date" name="startDate" value="<%= Server.HTMLEncode(startDate) %>">
                <label>结束日期:</label>
                <input type="date" name="endDate" value="<%= Server.HTMLEncode(endDate) %>">
                <button type="submit" class="admin-btn admin-btn-primary"><i class="fas fa-search"></i> 查询</button>
                <a href="reports.asp" class="admin-btn admin-btn-secondary"><i class="fas fa-redo"></i> 重置</a>
            </form>
        </div>
        
        <!-- 月度收入汇总 -->
        <div class="report-section">
            <h3><i class="fas fa-calendar-alt" style="color: #00bcd4;"></i> 月度收入汇总</h3>
            <% If Not rsMonthly Is Nothing Then %>
            <%
            Dim hasMonthlyData : hasMonthlyData = False
            Dim monthTotalOrders : monthTotalOrders = 0
            Dim monthTotalRevenue : monthTotalRevenue = 0
            %>
            <table class="data-table">
                <thead>
                    <tr>
                        <th>月份</th>
                        <th>订单数</th>
                        <th>总金额</th>
                        <th>月均客单价</th>
                    </tr>
                </thead>
                <tbody>
                    <% Do While Not rsMonthly.EOF %>
                    <% hasMonthlyData = True %>
                    <%
                    Dim mYear, mMonth, mCount, mRevenue, mAvg
                    mYear = rsMonthly("Y")
                    mMonth = rsMonthly("M")
                    mCount = rsMonthly("OrderCount")
                    mRevenue = rsMonthly("Revenue")
                    If IsNull(mRevenue) Then mRevenue = 0 Else mRevenue = CDbl(mRevenue)
                    If mCount > 0 Then mAvg = mRevenue / mCount Else mAvg = 0
                    monthTotalOrders = monthTotalOrders + CLng("0" & mCount)
                    monthTotalRevenue = monthTotalRevenue + CDbl("0" & mRevenue)
                    %>
                    <tr>
                        <td><%= mYear %>年<%= mMonth %>月</td>
                        <td class="count-highlight"><%= mCount %></td>
                        <td class="amount-highlight">¥<%= FormatNumber(CDbl("0" & mRevenue), 2) %></td>
                        <td>¥<%= FormatNumber(CDbl("0" & mAvg), 2) %></td>
                    </tr>
                    <% rsMonthly.MoveNext %>
                    <% Loop %>
                </tbody>
                <% If hasMonthlyData Then %>
                <tfoot>
                    <tr>
                        <td>合计</td>
                        <td class="count-highlight"><%= monthTotalOrders %></td>
                        <td class="amount-highlight">¥<%= FormatNumber(CDbl("0" & monthTotalRevenue), 2) %></td>
                        <td>¥<%= FormatNumber(CDbl("0" & IIf(monthTotalOrders > 0, monthTotalRevenue / monthTotalOrders, 0)), 2) %></td>
                    </tr>
                </tfoot>
                <% End If %>
            </table>
            <% If Not hasMonthlyData Then %>
            <div class="no-data"><i class="fas fa-inbox"></i>所选日期范围内暂无订单数据</div>
            <% End If %>
            <% rsMonthly.Close %>
            <% Else %>
            <div class="no-data"><i class="fas fa-inbox"></i>暂无数据</div>
            <% End If %>
        </div>
        
        <div class="charts-grid">
            <!-- 订单金额分布 -->
            <div class="report-section" style="margin-bottom: 0;">
                <h3><i class="fas fa-chart-bar" style="color: #ff9800;"></i> 订单金额分布</h3>
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>金额区间</th>
                            <th>订单数</th>
                            <th>占比</th>
                        </tr>
                    </thead>
                    <tbody>
                        <%
                        Dim ranges, rangeCounts, i, pctDist, barColors
                        ranges = Array("0 - 50元", "50 - 100元", "100 - 200元", "200 - 500元", "500元以上")
                        rangeCounts = Array(CLng("0" & cnt0_50), CLng("0" & cnt50_100), CLng("0" & cnt100_200), CLng("0" & cnt200_500), CLng("0" & cnt500plus))
                        barColors = Array("#00bcd4", "#4CAF50", "#ff9800", "#e91e63", "#9C27B0")
                        For i = 0 To UBound(ranges)
                            pctDist = Round((rangeCounts(i) / totalDistribution) * 100, 1)
                        %>
                        <tr>
                            <td><%= ranges(i) %></td>
                            <td class="count-highlight"><%= rangeCounts(i) %></td>
                            <td>
                                <div class="progress-bar" style="width: 120px;">
                                    <div style="height: 100%; width: <%= pctDist %>%; background: <%= barColors(i) %>; border-radius: 10px;"></div>
                                </div>
                                <span style="margin-left: 8px;"><%= pctDist %>%</span>
                            </td>
                        </tr>
                        <% Next %>
                    </tbody>
                </table>
            </div>
            
            <!-- 支付方式占比 -->
            <div class="report-section" style="margin-bottom: 0;">
                <h3><i class="fas fa-credit-card" style="color: #4CAF50;"></i> 支付方式占比</h3>
                <% If Not rsPayment Is Nothing Then %>
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
                        <% 
                        Dim hasPayData, payMethod, payLabel, payTotal, payPct
                        hasPayData = False
                        Do While Not rsPayment.EOF 
                            hasPayData = True
                            payMethod = rsPayment("PaymentMethod") & ""
                            Select Case LCase(payMethod)
                                Case "alipay"
                                    payLabel = "<i class='fab fa-alipay' style='color: #1677ff;'></i> 支付宝"
                                Case "wechat"
                                    payLabel = "<i class='fab fa-weixin' style='color: #07c160;'></i> 微信支付"
                                Case "bank"
                                    payLabel = "<i class='fas fa-university' style='color: #ff6b6b;'></i> 银行转账"
                                Case Else
                                    payLabel = Server.HTMLEncode(payMethod)
                            End Select
                            
                            payTotal = rsPayment("Total")
                            If IsNull(payTotal) Then payTotal = 0 Else payTotal = CDbl(payTotal)
                            If CDbl(grandTotal) > 0 Then payPct = Round((CDbl(payTotal) / CDbl(grandTotal)) * 100, 1) Else payPct = 0
                        %>
                        <tr>
                            <td><%= payLabel %></td>
                            <td class="count-highlight"><%= rsPayment("Cnt") %></td>
                            <td class="amount-highlight">¥<%= FormatNumber(CDbl("0" & payTotal), 2) %></td>
                            <td>
                                <div class="progress-bar" style="width: 100px;">
                                    <div class="progress-fill blue" style="width: <%= payPct %>%;"></div>
                                </div>
                                <span style="margin-left: 8px;"><%= payPct %>%</span>
                            </td>
                        </tr>
                        <% 
                            rsPayment.MoveNext
                        Loop 
                        %>
                    </tbody>
                </table>
                <% If Not hasPayData Then %>
                <div class="no-data"><i class="fas fa-inbox"></i>暂无支付数据</div>
                <% End If %>
                <% rsPayment.Close %>
                <% Else %>
                <div class="no-data"><i class="fas fa-inbox"></i>暂无支付数据</div>
                <% End If %>
            </div>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
