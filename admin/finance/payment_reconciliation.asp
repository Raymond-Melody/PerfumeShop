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

' ============================================
' V18.0 支付对账工具
' 自动比对系统订单与支付网关记录，发现差异
' ============================================

' 筛选参数
Dim dateFrom, dateTo, payMethod, showStatus
dateFrom = Trim(Request.QueryString("dateFrom"))
dateTo = Trim(Request.QueryString("dateTo"))
payMethod = Trim(Request.QueryString("pm"))
showStatus = Trim(Request.QueryString("status"))

If dateFrom = "" Then dateFrom = DateAdd("d", -7, Date())
If dateTo = "" Then dateTo = Date()

' ============================================
' 统计汇总
' ============================================
Dim totalOrders, totalMatched, totalUnmatched, totalDiffCount, totalDiffAmount
totalOrders = 0 : totalMatched = 0 : totalUnmatched = 0 : totalDiffCount = 0 : totalDiffAmount = 0

Function SafeCLng(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeCLng = 0 Else SafeCLng = CLng(val)
End Function
Function SafeCDbl(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeCDbl = 0 Else SafeCDbl = CDbl(val)
End Function

' 总订单数（已付款/处理中的订单）
totalOrders = SafeCLng(GetScalar("SELECT COUNT(*) FROM Orders WHERE Status IN ('Paid','Processing','Shipped','Completed') " & _
    "AND CAST(CreatedAt AS DATE) >= '" & SafeSQL(dateFrom) & "' AND CAST(CreatedAt AS DATE) <= '" & SafeSQL(dateTo) & "'"))

' 匹配/未匹配统计基于 AppLogs 中的回调记录
totalMatched = SafeCLng(GetScalar("SELECT COUNT(DISTINCT o.OrderID) FROM Orders o " & _
    "INNER JOIN AppLogs al ON al.LogMessage LIKE '%OrderID=' + CAST(o.OrderID AS VARCHAR) + '%' " & _
    "AND al.LogSource = 'cookie_consent' " & _
    "WHERE o.Status IN ('Paid','Processing','Shipped','Completed') " & _
    "AND CAST(o.CreatedAt AS DATE) >= '" & SafeSQL(dateFrom) & "' AND CAST(o.CreatedAt AS DATE) <= '" & SafeSQL(dateTo) & "'"))

If totalMatched < 0 Then totalMatched = 0
totalUnmatched = totalOrders - totalMatched
If totalUnmatched < 0 Then totalUnmatched = 0

' ============================================
' 对账明细查询
' ============================================
Dim sqlWhere, rsOrders, orderSQL
sqlWhere = " WHERE o.Status IN ('Paid','Processing','Shipped','Completed') " & _
    "AND CAST(o.CreatedAt AS DATE) >= '" & SafeSQL(dateFrom) & "' AND CAST(o.CreatedAt AS DATE) <= '" & SafeSQL(dateTo) & "'"

If payMethod <> "" And IsNumeric(payMethod) Then
    sqlWhere = sqlWhere & " AND o.PaymentMethod = " & CLng(payMethod)
End If

If showStatus = "matched" Then
    ' 已有回调记录
    sqlWhere = sqlWhere & " AND EXISTS (SELECT 1 FROM AppLogs al WHERE al.LogMessage LIKE '%OrderID=' + CAST(o.OrderID AS VARCHAR) + '%')"
ElseIf showStatus = "unmatched" Then
    sqlWhere = sqlWhere & " AND NOT EXISTS (SELECT 1 FROM AppLogs al WHERE al.LogMessage LIKE '%OrderID=' + CAST(o.OrderID AS VARCHAR) + '%')"
End If

orderSQL = "SELECT TOP 100 o.OrderID, o.OrderNo, o.TotalAmount, o.PaymentMethod, o.Status, o.CreatedAt, " & _
    "o.ShippingName, u.Username, " & _
    "CASE WHEN EXISTS (SELECT 1 FROM AppLogs al WHERE al.LogMessage LIKE '%OrderID=' + CAST(o.OrderID AS VARCHAR) + '%') THEN 1 ELSE 0 END AS HasCallback " & _
    "FROM Orders o LEFT JOIN Users u ON o.UserID = u.UserID" & sqlWhere & " ORDER BY o.CreatedAt DESC"

Set rsOrders = ExecuteQuery(orderSQL)
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>支付对账 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; }
        .main-content { padding: 24px; max-width: 1400px; margin: 0 auto; }
        .page-header { margin-bottom: 24px; }
        .page-title { font-size: 22px; font-weight: 600; color: #fff; margin: 0; }
        .breadcrumb { font-size: 13px; color: #888; margin-top: 6px; }
        .breadcrumb a { color: #00bcd4; text-decoration: none; }
        
        /* 摘要卡片 */
        .summary-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 24px; }
        .summary-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 10px; padding: 20px; border: 1px solid rgba(255,255,255,0.05); }
        .summary-card .card-label { font-size: 12px; color: #888; text-transform: uppercase; }
        .summary-card .card-value { font-size: 28px; font-weight: 700; margin-top: 6px; }
        .summary-card .card-icon { float: right; font-size: 28px; opacity: 0.3; }
        .card-blue .card-value { color: #2196F3; }
        .card-green .card-value { color: #4CAF50; }
        .card-orange .card-value { color: #FF9800; }
        .card-red .card-value { color: #F44336; }
        
        /* 筛选栏 */
        .filter-bar { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 10px; padding: 16px 20px; margin-bottom: 24px; border: 1px solid rgba(255,255,255,0.05); }
        .filter-bar form { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
        .filter-bar label { font-size: 13px; color: #888; }
        .filter-bar input[type="date"], .filter-bar select { background: rgba(0,0,0,0.3); border: 1px solid rgba(255,255,255,0.1); color: #e0e0e0; padding: 8px 12px; border-radius: 6px; font-size: 13px; }
        .filter-bar .btn { padding: 8px 16px; background: #2196F3; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; }
        .filter-bar .btn:hover { background: #1976D2; }
        .filter-bar .btn-outline { background: transparent; border: 1px solid rgba(255,255,255,0.2); }
        .filter-bar .btn-outline:hover { background: rgba(255,255,255,0.05); }
        
        /* 表格 */
        .data-table { width: 100%; border-collapse: collapse; background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 10px; overflow: hidden; border: 1px solid rgba(255,255,255,0.05); }
        .data-table th { background: rgba(0,0,0,0.3); color: #aaa; font-size: 12px; font-weight: 600; text-transform: uppercase; padding: 12px 16px; text-align: left; }
        .data-table td { padding: 12px 16px; font-size: 13px; border-top: 1px solid rgba(255,255,255,0.03); }
        .data-table tr:hover td { background: rgba(255,255,255,0.02); }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 600; }
        .badge-matched { background: rgba(76,175,80,0.15); color: #4CAF50; }
        .badge-unmatched { background: rgba(244,67,54,0.15); color: #F44336; }
        .badge-diff { background: rgba(255,152,0,0.15); color: #FF9800; }
        .text-muted { color: #666; }
        .text-right { text-align: right; }
        .empty-state { text-align: center; padding: 48px 24px; color: #666; }
        .empty-state i { font-size: 48px; margin-bottom: 12px; display: block; opacity: 0.4; }
        
        @media (max-width: 1024px) { .summary-grid { grid-template-columns: repeat(2, 1fr); } }
        @media (max-width: 640px) { .summary-grid { grid-template-columns: 1fr; } .filter-bar form { flex-direction: column; align-items: stretch; } }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-balance-scale"></i> 支付对账</h2>
            <div class="breadcrumb">
                <a href="index.asp">财务中心</a> / <span>支付对账</span>
            </div>
        </div>
        
        <!-- 摘要卡片 -->
        <div class="summary-grid">
            <div class="summary-card">
                <div class="card-icon"><i class="fas fa-receipt"></i></div>
                <div class="card-label">待对账订单</div>
                <div class="card-value card-blue"><%= totalOrders %></div>
            </div>
            <div class="summary-card">
                <div class="card-icon"><i class="fas fa-check-circle"></i></div>
                <div class="card-label">已匹配</div>
                <div class="card-value card-green"><%= totalMatched %></div>
            </div>
            <div class="summary-card">
                <div class="card-icon"><i class="fas fa-exclamation-circle"></i></div>
                <div class="card-label">未匹配</div>
                <div class="card-value card-orange"><%= totalUnmatched %></div>
            </div>
            <div class="summary-card">
                <div class="card-icon"><i class="fas fa-yen-sign"></i></div>
                <div class="card-label">缺陷说明</div>
                <div class="card-value card-red" style="font-size:14px;font-weight:500;">支付网关回调记录与系统订单对照</div>
            </div>
        </div>
        
        <!-- 筛选栏 -->
        <div class="filter-bar">
            <form method="get" action="payment_reconciliation.asp">
                <label>日期范围:</label>
                <input type="date" name="dateFrom" value="<%= dateFrom %>">
                <span style="color:#888;">至</span>
                <input type="date" name="dateTo" value="<%= dateTo %>">
                <label>支付方式:</label>
                <select name="pm">
                    <option value="">全部</option>
                    <option value="1" <%= IIf(payMethod="1","selected","") %>>微信支付</option>
                    <option value="2" <%= IIf(payMethod="2","selected","") %>>支付宝</option>
                    <option value="3" <%= IIf(payMethod="3","selected","") %>>PayPal</option>
                    <option value="4" <%= IIf(payMethod="4","selected","") %>>货到付款</option>
                </select>
                <label>状态:</label>
                <select name="status">
                    <option value="">全部</option>
                    <option value="matched" <%= IIf(showStatus="matched","selected","") %>>已匹配</option>
                    <option value="unmatched" <%= IIf(showStatus="unmatched","selected","") %>>未匹配</option>
                </select>
                <button type="submit" class="btn"><i class="fas fa-filter"></i> 筛选</button>
                <a href="payment_reconciliation.asp" class="btn btn-outline"><i class="fas fa-redo"></i> 重置</a>
            </form>
        </div>
        
        <!-- 对账明细表格 -->
        <%
        If Not rsOrders Is Nothing And Not rsOrders.EOF Then
            Dim pmName, hasCallback, rowIndex
            rowIndex = 0
        %>
        <div style="overflow-x:auto;">
            <table class="data-table">
                <thead>
                    <tr>
                        <th>订单号</th>
                        <th>用户</th>
                        <th>支付方式</th>
                        <th>金额</th>
                        <th>订单状态</th>
                        <th>创建时间</th>
                        <th>对账状态</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    Do While Not rsOrders.EOF
                        rowIndex = rowIndex + 1
                        ' 支付方式名称
                        Select Case SafeCLng(rsOrders("PaymentMethod"))
                            Case PAYMENT_METHOD_WECHAT: pmName = "微信支付"
                            Case PAYMENT_METHOD_ALIPAY: pmName = "支付宝"
                            Case PAYMENT_METHOD_PAYPAL: pmName = "PayPal"
                            Case PAYMENT_METHOD_COD: pmName = "货到付款"
                            Case Else: pmName = "未知"
                        End Select
                        hasCallback = SafeCLng(rsOrders("HasCallback"))
                    %>
                    <tr>
                        <td>
                            <a href="/admin/operation/order_detail.asp?id=<%= rsOrders("OrderID") %>" style="color:#2196F3;text-decoration:none;">
                                #<%= rsOrders("OrderNo") %>
                            </a>
                        </td>
                        <td><%= rsOrders("Username") & "" %></td>
                        <td><%= pmName %></td>
                        <td class="text-right">¥<%= FormatNumber(SafeCDbl(rsOrders("TotalAmount")), 2) %></td>
                        <td><%= rsOrders("Status") & "" %></td>
                        <td class="text-muted"><%= Left(rsOrders("CreatedAt") & "", 16) %></td>
                        <td>
                            <% If hasCallback > 0 Then %>
                                <span class="badge badge-matched"><i class="fas fa-check"></i> 已匹配</span>
                            <% Else %>
                                <span class="badge badge-unmatched"><i class="fas fa-question"></i> 未匹配</span>
                            <% End If %>
                        </td>
                    </tr>
                    <%
                        rsOrders.MoveNext
                    Loop
                    %>
                </tbody>
            </table>
        </div>
        <%
        Else
        %>
        <div class="empty-state">
            <i class="fas fa-inbox"></i>
            <p>当前筛选条件下无对账记录</p>
        </div>
        <%
        End If
        If Not rsOrders Is Nothing Then
            rsOrders.Close
            Set rsOrders = Nothing
        End If
        %>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
