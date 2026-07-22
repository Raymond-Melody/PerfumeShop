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

' V8：供应商对账数据
On Error Resume Next
conn.Execute "SELECT PayableID FROM AccountsPayable WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "CREATE TABLE AccountsPayable (PayableID INT IDENTITY(1,1) PRIMARY KEY, PurchaseID INT, SupplierID INT NOT NULL, SupplierName NVARCHAR(200), PayableNo NVARCHAR(50), Amount DECIMAL(19,4) DEFAULT 0, PaidAmount DECIMAL(19,4) DEFAULT 0, Status NVARCHAR(20) DEFAULT 'Pending', DueDate DATE, InvoiceNo NVARCHAR(100), Notes NVARCHAR(MAX), CreatedAt DATETIME2 DEFAULT GETDATE(), UpdatedAt DATETIME2 DEFAULT GETDATE())"
On Error GoTo 0

Dim apForRecon : apForRecon = GetScalar("SELECT ISNULL(SUM(Amount-PaidAmount),0) FROM AccountsPayable WHERE Status IN ('Pending','Partial')")

' 权限检查
Dim canEdit
If Session("AdminRoleCode") = "FIN_MANAGER" OR Session("AdminRoleCode") = "SUPER_ADMIN" Then
    canEdit = True
Else
    canEdit = False
End If

' 获取当前Tab
Dim currentTab
currentTab = Request.QueryString("tab")
If currentTab = "" Then currentTab = "auto"

' 处理表单提交
Dim action
action = Request.Form("action")

' 自动对账处理
If action = "run_reconciliation" AND canEdit Then
    If Not ValidateCSRFToken() Then
        Response.Redirect "reconciliation.asp?tab=auto&error=安全验证失败"
        Response.End
    End If
    
    Dim startDate, endDate
    startDate = Request.Form("startDate")
    endDate = Request.Form("endDate")
    
    If startDate = "" OR endDate = "" Then
        Response.Redirect "reconciliation.asp?tab=auto&error=请选择对账日期范围"
        Response.End
    End If
    
    ' 执行对账逻辑
    Call RunReconciliation(startDate, endDate)
    Response.Redirect "reconciliation.asp?tab=results&msg=对账完成"
    Response.End
End If

' 手动处理异常项
If action = "resolve_item" AND canEdit Then
    If Not ValidateCSRFToken() Then
        Response.Redirect "reconciliation.asp?tab=manual&error=安全验证失败"
        Response.End
    End If
    
    Dim logID, resolution
    logID = Request.Form("logID")
    resolution = Request.Form("resolution")
    
    If logID <> "" AND resolution <> "" Then
        Dim resolveSQL
        resolveSQL = "UPDATE ReconciliationLogs SET Status='Resolved', Resolution='" & SafeSQL(resolution) & "', ResolvedBy='" & SafeSQL(Session("AdminName")) & "', ResolvedAt= GETDATE() WHERE LogID=" & CLng(logID)
        ExecuteNonQuery resolveSQL
        Call LogAdminAction("处理对账异常", "finance", "ReconciliationLogs", logID, resolution)
    End If
    
    Response.Redirect "reconciliation.asp?tab=manual&msg=处理完成"
    Response.End
End If

' V20 导出CSV
If Request.QueryString("action") = "export_csv" Then
    Response.ContentType = "text/csv; charset=utf-8"
    Response.AddHeader "Content-Disposition", "attachment; filename=reconciliation_" & Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2) & ".csv"
    Response.CodePage = 65001
    Response.BinaryWrite ChrB(&HEF) & ChrB(&HBB) & ChrB(&HBF) ' UTF-8 BOM
    
    Dim expFilter : expFilter = Request.Form("export_filter")
    Dim expSQL, expRS, expLine
    expSQL = "SELECT LogID, ReconcileDate, OrderNo, OrderAmount, PaymentAmount, Difference, Status, Resolution FROM ReconciliationLogs"
    If expFilter <> "" Then expSQL = expSQL & " WHERE Status='" & SafeSQL(expFilter) & "'"
    expSQL = expSQL & " ORDER BY CreatedAt DESC"
    Set expRS = ExecuteQuery(expSQL)
    
    ' CSV Header
    Response.Write "ID,对账日期,订单号,订单金额,支付金额,差异,状态,处理说明" & vbCrLf
    If Not expRS Is Nothing And IsObject(expRS) Then
        Do While Not expRS.EOF
            Response.Write expRS("LogID") & "," & Chr(34) & FormatDateTime(expRS("ReconcileDate"), 2) & Chr(34) & ","
            Response.Write Chr(34) & expRS("OrderNo") & Chr(34) & "," & FormatNumber(SafeNum(expRS("OrderAmount")), 2) & ","
            Response.Write FormatNumber(SafeNum(expRS("PaymentAmount")), 2) & "," & FormatNumber(SafeNum(expRS("Difference")), 2) & ","
            Response.Write Chr(34) & StatusLabelCN(expRS("Status")) & Chr(34) & "," & Chr(34) & expRS("Resolution") & Chr(34) & vbCrLf
            expRS.MoveNext
        Loop
        expRS.Close : Set expRS = Nothing
    End If
    Response.End
End If

' 退款核销处理
If action = "process_refund" AND canEdit Then
    If Not ValidateCSRFToken() Then
        Response.Redirect "reconciliation.asp?tab=refund&error=安全验证失败"
        Response.End
    End If
    
    Dim refundID
    refundID = Request.Form("refundID")
    
    If refundID <> "" Then
        Call ProcessRefundReconciliation(refundID)
    End If
    
    Response.Redirect "reconciliation.asp?tab=refund&msg=退款核销完成"
    Response.End
End If

' 对账主函数（V20 幂等模式：先删后插）
Sub RunReconciliation(startDate, endDate)
    Dim orderSQL, paymentSQL
    Dim orderRS, paymentRS
    Dim matchCount, shortCount, overCount, missingCount
    matchCount = 0
    shortCount = 0
    overCount = 0
    missingCount = 0
    
    ' V20 幂等：清除同区间旧对账记录，支持重新运行
    Call ExecuteNonQuery("DELETE FROM ReconciliationLogs WHERE CreatedAt >= '" & startDate & "' AND CreatedAt < DATEADD(day, 1, '" & endDate & "')")
    
    ' 获取已付款订单
    orderSQL = "SELECT OrderID, OrderNo, TotalAmount FROM Orders WHERE Status='Paid' AND CreatedAt >= '" & startDate & "' AND CreatedAt < DATEADD(day, 1, '" & endDate & "')"
    Set orderRS = ExecuteQuery(orderSQL)
    
    ' 获取支付记录
    paymentSQL = "SELECT OrderID, OrderNo, TransactionNo, Amount FROM PaymentRecords WHERE CreatedAt >= '" & startDate & "' AND CreatedAt < DATEADD(day, 1, '" & endDate & "')"
    Set paymentRS = ExecuteQuery(paymentSQL)
    
    ' 创建字典存储订单和支付记录
    Dim orderDict, paymentDict, orderKey, paymentKey
    Dim oID, oNo, oAmount, diff, paySQL, payRS, payAmount
    Dim pOrderNo, pAmount, checkOrderSQL, checkOrderRS
    Set orderDict = Server.CreateObject("Scripting.Dictionary")
    Set paymentDict = Server.CreateObject("Scripting.Dictionary")
    
    ' 填充订单字典
    If Not orderRS Is Nothing Then
        Do While Not orderRS.EOF
            orderKey = CStr(orderRS("OrderNo")) & "_" & CStr(orderRS("TotalAmount"))
            If Not orderDict.Exists(orderKey) Then
                orderDict.Add orderKey, orderRS("OrderID")
            End If
            orderRS.MoveNext
        Loop
        orderRS.Close
        Set orderRS = Nothing
    End If
    
    ' 填充支付字典
    If Not paymentRS Is Nothing Then
        Do While Not paymentRS.EOF
            paymentKey = CStr(paymentRS("OrderNo")) & "_" & CStr(paymentRS("Amount"))
            If Not paymentDict.Exists(paymentKey) Then
                paymentDict.Add paymentKey, paymentRS("OrderID")
            End If
            paymentRS.MoveNext
        Loop
        paymentRS.Close
        Set paymentRS = Nothing
    End If
    
    ' 重新获取数据进行匹配
    Set orderRS = ExecuteQuery(orderSQL)
    
    If Not orderRS Is Nothing Then
        Do While Not orderRS.EOF
            oID = orderRS("OrderID")
            oNo = orderRS("OrderNo")
            oAmount = CDbl(orderRS("TotalAmount"))
            
            ' 查找对应支付记录
            paySQL = "SELECT Amount FROM PaymentRecords WHERE OrderNo='" & SafeSQL(oNo) & "'"
            Set payRS = ExecuteQuery(paySQL)
            
            If Not payRS Is Nothing AND Not payRS.EOF Then
                payAmount = CDbl(payRS("Amount"))
                diff = payAmount - oAmount
                
                If diff = 0 Then
                    Call SaveReconciliationLog(oID, oNo, oAmount, payAmount, diff, "Matched", "")
                    matchCount = matchCount + 1
                ElseIf diff < 0 Then
                    Call SaveReconciliationLog(oID, oNo, oAmount, payAmount, diff, "ShortPay", "支付金额不足")
                    shortCount = shortCount + 1
                Else
                    Call SaveReconciliationLog(oID, oNo, oAmount, payAmount, diff, "OverPay", "支付金额超额")
                    overCount = overCount + 1
                End If
                payRS.Close
                Set payRS = Nothing
            Else
                ' 未找到支付记录
                Call SaveReconciliationLog(oID, oNo, oAmount, 0, -oAmount, "Missing", "未找到支付记录")
                missingCount = missingCount + 1
            End If
            
            orderRS.MoveNext
        Loop
        orderRS.Close
        Set orderRS = Nothing
    End If
    
    ' 检查支付记录中有但订单中没有的（未达账项）
    Set paymentRS = ExecuteQuery(paymentSQL)
    If Not paymentRS Is Nothing Then
        Do While Not paymentRS.EOF
            pOrderNo = paymentRS("OrderNo")
            pAmount = CDbl(paymentRS("Amount"))
            
            ' 检查订单是否存在
            checkOrderSQL = "SELECT OrderID FROM Orders WHERE OrderNo='" & SafeSQL(pOrderNo) & "'"
            Set checkOrderRS = ExecuteQuery(checkOrderSQL)
            
            If checkOrderRS Is Nothing OR checkOrderRS.EOF Then
                ' 支付记录存在但订单不存在
                Call SaveReconciliationLog(0, pOrderNo, 0, pAmount, pAmount, "Missing", "未找到对应订单")
                missingCount = missingCount + 1
            End If
            
            If Not checkOrderRS Is Nothing Then
                checkOrderRS.Close
                Set checkOrderRS = Nothing
            End If
            
            paymentRS.MoveNext
        Loop
        paymentRS.Close
        Set paymentRS = Nothing
    End If
    
    ' 记录对账结果到Session用于显示
    Session("ReconcileResult") = "匹配:" & matchCount & ",短款:" & shortCount & ",长款:" & overCount & ",未达:" & missingCount
    
    Set orderDict = Nothing
    Set paymentDict = Nothing
End Sub

' 保存对账日志
Sub SaveReconciliationLog(orderID, orderNo, orderAmount, paymentAmount, difference, status, remark)
    Dim insertSQL
    insertSQL = "INSERT INTO ReconciliationLogs (ReconcileDate, OrderID, OrderNo, OrderAmount, PaymentAmount, Difference, Status, Resolution, CreatedAt) VALUES (GETDATE(), " & orderID & ", '" & SafeSQL(orderNo) & "', " & orderAmount & ", " & paymentAmount & ", " & difference & ", '" & SafeSQL(status) & "', '" & SafeSQL(remark) & "', GETDATE())"
    ExecuteNonQuery insertSQL
End Sub

' 处理退款核销
Sub ProcessRefundReconciliation(refundID)
    Dim refundSQL, refundRS
    refundSQL = "SELECT * FROM RefundRecords WHERE RefundID=" & CLng(refundID)
    Set refundRS = ExecuteQuery(refundSQL)
    
    If Not refundRS Is Nothing AND Not refundRS.EOF Then
        Dim orderID, orderNo, refundAmount
        orderID = refundRS("OrderID")
        orderNo = refundRS("OrderNo")
        refundAmount = CDbl(refundRS("RefundAmount"))
        
        ' 在PaymentRecords中生成负向流水
        Dim paymentSQL
        paymentSQL = "INSERT INTO PaymentRecords (OrderID, OrderNo, TransactionNo, PaymentMethod, TransactionType, Amount, Fee, NetAmount, Status, Category, Remark, CreatedAt) VALUES (" & orderID & ", '" & SafeSQL(orderNo) & "', 'REFUND_" & refundID & "', 'Refund', 'Refund', -" & refundAmount & ", 0, -" & refundAmount & ", 'Completed', 'Refund', '退款核销', GETDATE())"
        ExecuteNonQuery paymentSQL
        
        ' 更新退款记录状态
        Dim updateRefundSQL
        updateRefundSQL = "UPDATE RefundRecords SET Status='Completed', CompletedAt= GETDATE() WHERE RefundID=" & CLng(refundID)
        ExecuteNonQuery updateRefundSQL
        
        ' 记录到对账日志
        Call SaveReconciliationLog(orderID, orderNo, 0, -refundAmount, -refundAmount, "Refund", "退款核销记录")
        
        refundRS.Close
    End If
    Set refundRS = Nothing
End Sub

' V20 获取对账统计数据（单次查询，避免多次GetScalar调用）
Function GetReconciliationStats()
    Dim stats(5), i, rs, s, c
    For i = 0 To 5 : stats(i) = 0 : Next
    
    Dim sql : sql = "SELECT Status, COUNT(*) AS Cnt, ISNULL(SUM(Difference),0) AS TotalDiff FROM ReconciliationLogs GROUP BY Status"
    Set rs = ExecuteQuery(sql)
    If Not rs Is Nothing And IsObject(rs) Then
        Do While Not rs.EOF
            s = CStr(rs(0).Value & "")
            c = CLng(rs(1).Value & "")
            Select Case s
                Case "Matched"  : stats(0) = c
                Case "ShortPay" : stats(1) = c
                Case "OverPay"  : stats(2) = c
                Case "Missing"  : stats(3) = c
                Case "Resolved" : stats(4) = c
            End Select
            rs.MoveNext
        Loop
        rs.Close : Set rs = Nothing
    End If
    
    ' stats(5) = 总记录数
    stats(5) = stats(0) + stats(1) + stats(2) + stats(3) + stats(4)
    GetReconciliationStats = stats
End Function

' V20 状态中文标签
Function StatusLabelCN(status)
    Select Case CStr(status & "")
        Case "Matched"  : StatusLabelCN = "匹配"
        Case "ShortPay" : StatusLabelCN = "短款"
        Case "OverPay"  : StatusLabelCN = "长款"
        Case "Missing"  : StatusLabelCN = "未达"
        Case "Resolved" : StatusLabelCN = "已解决"
        Case "Refund"   : StatusLabelCN = "退款"
        Case Else        : StatusLabelCN = CStr(status & "")
    End Select
End Function

' V20 汇总金额数据
Function GetTotalsByStatus()
    Dim totals(4), rs, s, a
    Dim i : For i = 0 To 4 : totals(i) = 0 : Next
    
    Dim sql : sql = "SELECT Status, ISNULL(SUM(OrderAmount),0), ISNULL(SUM(PaymentAmount),0), ISNULL(SUM(Difference),0) FROM ReconciliationLogs GROUP BY Status"
    Set rs = ExecuteQuery(sql)
    If Not rs Is Nothing And IsObject(rs) Then
        Do While Not rs.EOF
            s = CStr(rs(0).Value & "")
            a = CDbl(rs(1).Value & "")
            ' Store only total amounts indexed by status match
            rs.MoveNext
        Loop
        rs.Close : Set rs = Nothing
    End If
    GetTotalsByStatus = totals
End Function

Call LogAdminAction("访问对账中心", "finance", "", "", "")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>对账中心 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        .tab-container { margin-bottom: 25px; }
        .tab-nav { display: flex; gap: 5px; border-bottom: 2px solid #3a3a4a; margin-bottom: 25px; }
        .tab-nav a { 
            padding: 15px 25px; color: #888; text-decoration: none; 
            border-bottom: 3px solid transparent; transition: all 0.3s;
            display: flex; align-items: center; gap: 8px;
        }
        .tab-nav a:hover { color: #e0e0e0; background: #2d2d44; }
        .tab-nav a.active { color: #00bcd4; border-bottom-color: #00bcd4; background: #2d2d44; }
        
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin-bottom: 30px; }
        .stat-card { 
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; padding: 25px; 
            text-align: center; border: 1px solid rgba(255,255,255,0.06);
        }
        .stat-card.matched { border-left: 4px solid #4CAF50; }
        .stat-card.short { border-left: 4px solid #ff5252; }
        .stat-card.over { border-left: 4px solid #ff5252; }
        .stat-card.missing { border-left: 4px solid #ffa726; }
        .stat-card.resolved { border-left: 4px solid #9e9e9e; }
        .stat-value { font-size: 36px; font-weight: bold; margin-bottom: 8px; }
        .stat-card.matched .stat-value { color: #4CAF50; }
        .stat-card.short .stat-value { color: #ff5252; }
        .stat-card.over .stat-value { color: #ff5252; }
        .stat-card.missing .stat-value { color: #ffa726; }
        .stat-card.resolved .stat-value { color: #9e9e9e; }
        .stat-label { color: #888; font-size: 14px; }
        
        .reconcile-form { 
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; padding: 30px; 
            border: 1px solid rgba(255,255,255,0.06); max-width: 600px;
        }
        .form-row { display: flex; gap: 20px; margin-bottom: 20px; }
        .form-group { flex: 1; }
        .form-group label { display: block; margin-bottom: 8px; color: #b0b0b0; font-weight: 500; }
        .form-group input { 
            width: 100%; padding: 12px 15px; border: 2px solid #3a3a4a; 
            border-radius: 8px; font-size: 14px; background: #1a1a2e; color: #e0e0e0;
        }
        .btn-primary { 
            background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); 
            color: white; padding: 15px 40px; border: none; border-radius: 8px; 
            font-size: 16px; cursor: pointer; display: inline-flex; align-items: center; gap: 10px;
        }
        .btn-primary:hover { opacity: 0.9; }
        .btn-primary:disabled { opacity: 0.5; cursor: not-allowed; }
        
        .progress-container { 
            background: #1a1a2e; border-radius: 10px; padding: 20px; 
            margin-top: 20px; display: none;
        }
        .progress-bar { 
            height: 20px; background: #3a3a4a; border-radius: 10px; overflow: hidden;
        }
        .progress-fill { 
            height: 100%; background: linear-gradient(90deg, #00bcd4, #00838f); 
            width: 0%; transition: width 0.3s; border-radius: 10px;
        }
        .progress-text { text-align: center; margin-top: 10px; color: #888; }
        
        .filter-bar { 
            display: flex; gap: 15px; margin-bottom: 20px; flex-wrap: wrap;
            background: #2d2d44; padding: 15px; border-radius: 8px;
        }
        .filter-bar select, .filter-bar input { 
            padding: 10px 15px; border: 1px solid rgba(255,255,255,0.15); 
            border-radius: 6px; background: #1a1a2e; color: #e0e0e0;
        }
        
        .data-table { width: 100%; border-collapse: collapse; background: #2a2a3a; border-radius: 12px; overflow: hidden; }
        .data-table th, .data-table td { padding: 15px; text-align: left; border-bottom: 1px solid #3a3a4a; }
        .data-table th { background: #1a1a2e; color: #b0b0b0; font-weight: 600; }
        .data-table td { color: #e0e0e0; }
        .data-table tr:hover { background: #323242; }
        
        .status-badge { 
            display: inline-block; padding: 4px 12px; border-radius: 12px; 
            font-size: 12px; font-weight: 500;
        }
        .status-matched { background: #1b5e20; color: #81c784; }
        .status-shortpay { background: #5e1b1b; color: #e57373; }
        .status-overpay { background: #5e1b1b; color: #e57373; }
        .status-missing { background: #5e4b1b; color: #ffb74d; }
        .status-resolved { background: #424242; color: #bdbdbd; }
        .status-refund { background: #1a237e; color: #7986cb; }
        
        .amount-positive { color: #4CAF50; }
        .amount-negative { color: #ff5252; }
        
        .action-form { display: flex; gap: 10px; align-items: center; }
        .action-form input[type="text"] { 
            padding: 8px 12px; border: 1px solid #3a3a4a; 
            border-radius: 6px; background: #1a1a2e; color: #e0e0e0; flex: 1;
        }
        .btn-small { 
            padding: 8px 15px; border: none; border-radius: 6px; 
            cursor: pointer; font-size: 12px;
        }
        .btn-resolve { background: #4CAF50; color: white; }
        .btn-process { background: #00bcd4; color: white; }
        
        .readonly-mask { 
            position: relative; pointer-events: none; opacity: 0.7;
        }
        .readonly-mask::after { 
            content: "只读模式"; position: absolute; top: 50%; left: 50%; 
            transform: translate(-50%, -50%); background: rgba(0,0,0,0.8);
            padding: 10px 20px; border-radius: 6px; color: #888;
        }
        
        .report-section { 
            background: #2a2a3a; border-radius: 12px; padding: 25px; 
            margin-bottom: 25px; border: 1px solid #3a3a4a;
        }
        .report-title { 
            font-size: 18px; color: #e0e0e0; margin-bottom: 20px; 
            display: flex; align-items: center; gap: 10px;
        }
        .report-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; }
        .report-item { text-align: center; padding: 20px; background: #1a1a2e; border-radius: 8px; }
        .report-value { font-size: 28px; font-weight: bold; color: #00bcd4; margin-bottom: 5px; }
        .report-label { color: #888; font-size: 13px; }
        
        /* V20 预设按钮样式 */
        .preset-btn { 
            padding: 8px 14px; border: 1px solid #00bcd4; border-radius: 20px;
            background: transparent; color: #00bcd4; cursor: pointer;
            font-size: 13px; transition: all 0.2s;
        }
        .preset-btn:hover { background: rgba(0,188,212,0.15); }
        
        /* V20 导出按钮 */
        .btn-export { 
            background: #4CAF50; color: white; padding: 10px 18px; border: none;
            border-radius: 6px; cursor: pointer; font-size: 13px; display: inline-flex;
            align-items: center; gap: 6px; margin-left: 10px;
        }
        .btn-export:hover { background: #43A047; }
        
        /* V20 汇总行 */
        .total-row td { font-weight: bold; background: rgba(0,188,212,0.08) !important; }
        
        @media (max-width: 1200px) {
            .stats-grid { grid-template-columns: repeat(2, 1fr); }
            .report-grid { grid-template-columns: repeat(2, 1fr); }
        }
        @media (max-width: 768px) {
            .stats-grid { grid-template-columns: 1fr; }
            .report-grid { grid-template-columns: 1fr; }
            .form-row { flex-direction: column; }
        }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-balance-scale"></i> 对账中心</h2>
            <div class="breadcrumb">
                <a href="index.asp">财务中心</a> / <span>对账中心</span>
            </div>
        </div>
        
        <% If Request.QueryString("msg") <> "" Then %>
        <div class="alert alert-success"><i class="fas fa-check-circle"></i> <%= Server.HTMLEncode(Request.QueryString("msg")) %></div>
        <% End If %>
        
        <% If Request.QueryString("error") <> "" Then %>
        <div class="alert alert-error"><i class="fas fa-times-circle"></i> <%= Server.HTMLEncode(Request.QueryString("error")) %></div>
        <% End If %>
        
        <% If Session("ReconcileResult") <> "" Then %>
        <div class="alert alert-info">
            <i class="fas fa-info-circle"></i> 对账结果：<%= Server.HTMLEncode(Session("ReconcileResult")) %>
        </div>
        <% Session("ReconcileResult") = "" %>
        <% End If %>
        
        <!-- Tab导航 -->
        <div class="tab-container">
            <div class="tab-nav">
                <a href="?tab=auto" class="<%= IIf(currentTab="auto", "active", "") %>"><i class="fas fa-robot"></i> 自动对账</a>
                <a href="?tab=results" class="<%= IIf(currentTab="results", "active", "") %>"><i class="fas fa-list-alt"></i> 对账结果</a>
                <a href="?tab=manual" class="<%= IIf(currentTab="manual", "active", "") %>"><i class="fas fa-hand-pointer"></i> 手动处理</a>
                <a href="?tab=refund" class="<%= IIf(currentTab="refund", "active", "") %>"><i class="fas fa-undo-alt"></i> 退款核销</a>
                <a href="?tab=report" class="<%= IIf(currentTab="report", "active", "") %>"><i class="fas fa-chart-pie"></i> 对账报告</a>
            </div>
        </div>
        
        <!-- Tab 1: 自动对账 -->
        <div class="tab-content <%= IIf(currentTab="auto", "active", "") %>">
            <div class="stats-grid">
                <% Dim stats : stats = GetReconciliationStats() %>
                <div class="stat-card matched">
                    <div class="stat-value"><%= stats(0) %></div>
                    <div class="stat-label">匹配成功</div>
                </div>
                <div class="stat-card short">
                    <div class="stat-value"><%= stats(1) %></div>
                    <div class="stat-label">短款异常</div>
                </div>
                <div class="stat-card over">
                    <div class="stat-value"><%= stats(2) %></div>
                    <div class="stat-label">长款异常</div>
                </div>
                <div class="stat-card missing">
                    <div class="stat-value"><%= stats(3) %></div>
                    <div class="stat-label">未达账项</div>
                </div>
                <div class="stat-card" style="border-left:4px solid #00bcd4;">
                    <div class="stat-value" style="color:#00bcd4;">¥<%= FormatNumber(apForRecon, 0) %></div>
                    <div class="stat-label">供应商待付余额</div>
                </div>
            </div>
            
            <div class="reconcile-form <%= IIf(NOT canEdit, "readonly-mask", "") %>">
                <h3 style="color: #e0e0e0; margin-bottom: 15px;"><i class="fas fa-play-circle"></i> 开始对账</h3>
                
                <!-- V20 快捷日期选择 -->
                <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:18px;">
                    <button type="button" class="preset-btn" onclick="setPresetDate(0,0)" title="今天">📅 今天</button>
                    <button type="button" class="preset-btn" onclick="setPresetDate(-1,-1)" title="昨天">📆 昨天</button>
                    <button type="button" class="preset-btn" onclick="setPresetDate(-7,0)" title="最近7天">📊 近7天</button>
                    <button type="button" class="preset-btn" onclick="setPresetDate(-30,0)" title="最近30天">📈 近30天</button>
                    <button type="button" class="preset-btn" onclick="setMonthPreset(0)" title="本月">🗓 本月</button>
                    <button type="button" class="preset-btn" onclick="setMonthPreset(-1)" title="上月">📋 上月</button>
                </div>
                
                <form method="post" action="reconciliation.asp?tab=auto" id="reconcileForm">
                    <%= GetCSRFTokenField() %>
                    <input type="hidden" name="action" value="run_reconciliation">
                    
                    <div class="form-row">
                        <div class="form-group">
                            <label><i class="fas fa-calendar-alt"></i> 开始日期</label>
                            <input type="date" name="startDate" required value="<%= Year(Now()) & "-" & Right("0" & Month(Now()), 2) & "-" & Right("0" & Day(Now()), 2) %>">
                        </div>
                        <div class="form-group">
                            <label><i class="fas fa-calendar-alt"></i> 结束日期</label>
                            <input type="date" name="endDate" required value="<%= Year(Now()) & "-" & Right("0" & Month(Now()), 2) & "-" & Right("0" & Day(Now()), 2) %>">
                        </div>
                    </div>
                    
                    <button type="submit" class="btn-primary" <%= IIf(NOT canEdit, "disabled", "") %>>
                        <i class="fas fa-play"></i> 开始对账
                    </button>
                </form>
                
                <div class="progress-container" id="progressContainer">
                    <div class="progress-bar">
                        <div class="progress-fill" id="progressFill"></div>
                    </div>
                    <div class="progress-text" id="progressText">准备对账...</div>
                </div>
            </div>
        </div>
        
        <!-- Tab 2: 对账结果查看 -->
        <div class="tab-content <%= IIf(currentTab="results", "active", "") %>">
            <div class="filter-bar">
                <select id="statusFilter" onchange="filterResults()">
                    <option value="">全部状态</option>
                    <option value="Matched">匹配成功</option>
                    <option value="ShortPay">短款异常</option>
                    <option value="OverPay">长款异常</option>
                    <option value="Missing">未达账项</option>
                    <option value="Resolved">已解决</option>
                </select>
                <input type="date" id="dateFilter" onchange="filterResults()" placeholder="选择日期">
                <input type="text" id="orderFilter" onkeyup="filterResults()" placeholder="搜索订单号...">
                
                <!-- V20 导出按钮 -->
                <a href="reconciliation.asp?action=export_csv" class="btn-export" target="_blank" style="text-decoration:none;display:inline-flex;align-items:center;gap:6px;"><i class="fas fa-download"></i> 导出CSV</a>
            </div>
            
            <table class="data-table">
                <thead>
                    <tr>
                        <th>对账日期</th>
                        <th>订单号</th>
                        <th>订单金额</th>
                        <th>支付金额</th>
                        <th>差异</th>
                        <th>状态</th>
                        <th>处理说明</th>
                    </tr>
                </thead>
                <tbody>
                    <% 
                    Dim logSQL, logRS, diff, diffClass
                    logSQL = "SELECT TOP 100 * FROM ReconciliationLogs ORDER BY CreatedAt DESC"
                    Set logRS = ExecuteQuery(logSQL)
                    
                    If Not logRS Is Nothing Then
                        Do While Not logRS.EOF
                            diff = CDbl(logRS("Difference"))
                            If diff > 0 Then
                                diffClass = "amount-positive"
                            ElseIf diff < 0 Then
                                diffClass = "amount-negative"
                            Else
                                diffClass = ""
                            End If
                    %>
                    <tr class="log-row" data-status="<%= logRS("Status") %>" data-order="<%= logRS("OrderNo") %>">
                        <td><%= FormatDateField(logRS("ReconcileDate")) %></td>
                        <td><%= Server.HTMLEncode(logRS("OrderNo")) %></td>
                        <td>¥<%= FormatNumber(SafeNum(logRS("OrderAmount")), 2) %></td>
                        <td>¥<%= FormatNumber(SafeNum(logRS("PaymentAmount")), 2) %></td>
                        <td class="<%= diffClass %>">¥<%= FormatNumber(SafeNum(diff), 2) %></td>
                        <td>
                            <% Select Case logRS("Status")
                                Case "Matched" %>
                                <span class="status-badge status-matched"><i class="fas fa-check"></i> 匹配</span>
                            <% Case "ShortPay" %>
                                <span class="status-badge status-shortpay"><i class="fas fa-arrow-down"></i> 短款</span>
                            <% Case "OverPay" %>
                                <span class="status-badge status-overpay"><i class="fas fa-arrow-up"></i> 长款</span>
                            <% Case "Missing" %>
                                <span class="status-badge status-missing"><i class="fas fa-question"></i> 未达</span>
                            <% Case "Resolved" %>
                                <span class="status-badge status-resolved"><i class="fas fa-check-double"></i> 已解决</span>
                            <% Case "Refund" %>
                                <span class="status-badge status-refund"><i class="fas fa-undo"></i> 退款</span>
                            <% End Select %>
                        </td>
                        <td><%= Server.HTMLEncode(logRS("Resolution") & "") %></td>
                    </tr>
                    <% 
                            logRS.MoveNext
                        Loop
                        logRS.Close
                        Set logRS = Nothing
                    End If
                    %>
                </tbody>
            </table>
        </div>
        
        <!-- Tab 3: 手动处理 -->
        <div class="tab-content <%= IIf(currentTab="manual", "active", "") %>">
            <div class="filter-bar">
                <select onchange="location.href='?tab=manual&filter='+this.value">
                    <option value="">全部异常</option>
                    <option value="unresolved" <%= IIf(Request.QueryString("filter")="unresolved", "selected", "") %>>未解决</option>
                    <option value="short" <%= IIf(Request.QueryString("filter")="short", "selected", "") %>>仅短款</option>
                    <option value="over" <%= IIf(Request.QueryString("filter")="over", "selected", "") %>>仅长款</option>
                    <option value="missing" <%= IIf(Request.QueryString("filter")="missing", "selected", "") %>>仅未达</option>
                </select>
            </div>
            
            <table class="data-table">
                <thead>
                    <tr>
                        <th>订单号</th>
                        <th>订单金额</th>
                        <th>支付金额</th>
                        <th>差异</th>
                        <th>状态</th>
                        <th>处理说明</th>
                        <th>操作</th>
                    </tr>
                </thead>
                <tbody>
                    <% 
                    Dim manualSQL, manualRS, filterWhere
                    filterWhere = " WHERE Status IN ('ShortPay', 'OverPay', 'Missing')"
                    
                    If Request.QueryString("filter") = "unresolved" Then
                        filterWhere = " WHERE Status IN ('ShortPay', 'OverPay', 'Missing')"
                    ElseIf Request.QueryString("filter") = "short" Then
                        filterWhere = " WHERE Status='ShortPay'"
                    ElseIf Request.QueryString("filter") = "over" Then
                        filterWhere = " WHERE Status='OverPay'"
                    ElseIf Request.QueryString("filter") = "missing" Then
                        filterWhere = " WHERE Status='Missing'"
                    End If
                    
                    manualSQL = "SELECT * FROM ReconciliationLogs" & filterWhere & " ORDER BY CreatedAt DESC"
                    Set manualRS = ExecuteQuery(manualSQL)
                    
                    If Not manualRS Is Nothing Then
                        Do While Not manualRS.EOF
                    %>
                    <tr>
                        <td><%= Server.HTMLEncode(manualRS("OrderNo")) %></td>
                        <td>¥<%= FormatNumber(SafeNum(manualRS("OrderAmount")), 2) %></td>
                        <td>¥<%= FormatNumber(SafeNum(manualRS("PaymentAmount")), 2) %></td>
                        <td class="<%= IIf(CDbl(manualRS("Difference"))<0, "amount-negative", "amount-positive") %>">¥<%= FormatNumber(manualRS("Difference"), 2) %></td>
                        <td>
                            <% Select Case manualRS("Status")
                                Case "ShortPay" %>
                                <span class="status-badge status-shortpay">短款</span>
                            <% Case "OverPay" %>
                                <span class="status-badge status-overpay">长款</span>
                            <% Case "Missing" %>
                                <span class="status-badge status-missing">未达</span>
                            <% End Select %>
                        </td>
                        <td>
                            <% If canEdit Then %>
                            <form method="post" action="reconciliation.asp?tab=manual" class="action-form">
                                <%= GetCSRFTokenField() %>
                                <input type="hidden" name="action" value="resolve_item">
                                <input type="hidden" name="logID" value="<%= manualRS("LogID") %>">
                                <input type="text" name="resolution" placeholder="填写处理说明..." required>
                                <button type="submit" class="btn-small btn-resolve"><i class="fas fa-check"></i> 解决</button>
                            </form>
                            <% Else %>
                            <span style="color: #888;"><%= Server.HTMLEncode(manualRS("Resolution") & "") %></span>
                            <% End If %>
                        </td>
                        <td>
                            <% If manualRS("ResolvedBy") <> "" Then %>
                            <small style="color: #888;">
                                <%= Server.HTMLEncode(manualRS("ResolvedBy")) %><br>
                                <%= FormatDateField(manualRS("ResolvedAt")) %>
                            </small>
                            <% Else %>
                            <span style="color: #666;">-</span>
                            <% End If %>
                        </td>
                    </tr>
                    <% 
                            manualRS.MoveNext
                        Loop
                        manualRS.Close
                        Set manualRS = Nothing
                    End If
                    %>
                </tbody>
            </table>
        </div>
        
        <!-- Tab 4: 退款核销 -->
        <div class="tab-content <%= IIf(currentTab="refund", "active", "") %>">
            <table class="data-table">
                <thead>
                    <tr>
                        <th>退款单号</th>
                        <th>订单号</th>
                        <th>退款金额</th>
                        <th>退款原因</th>
                        <th>状态</th>
                        <th>申请时间</th>
                        <th>操作</th>
                    </tr>
                </thead>
                <tbody>
                    <% 
                    Dim refundSQL, refundRS
                    refundSQL = "SELECT * FROM RefundRecords WHERE Status IN ('Approved', 'Completed') ORDER BY CreatedAt DESC"
                    Set refundRS = ExecuteQuery(refundSQL)
                    
                    If Not refundRS Is Nothing Then
                        Do While Not refundRS.EOF
                    %>
                    <tr>
                        <td><%= Server.HTMLEncode(refundRS("RefundNo")) %></td>
                        <td><%= Server.HTMLEncode(refundRS("OrderNo")) %></td>
                        <td class="amount-negative">-¥<%= FormatNumber(refundRS("RefundAmount"), 2) %></td>
                        <td><%= Server.HTMLEncode(Left(refundRS("RefundReason") & "", 30)) %></td>
                        <td>
                            <% If refundRS("Status") = "Approved" Then %>
                                <span class="status-badge status-missing">待核销</span>
                            <% Else %>
                                <span class="status-badge status-resolved">已核销</span>
                            <% End If %>
                        </td>
                        <td><%= FormatDateField(refundRS("CreatedAt")) %></td>
                        <td>
                            <% If refundRS("Status") = "Approved" AND canEdit Then %>
                            <form method="post" action="reconciliation.asp?tab=refund" style="display: inline;">
                                <%= GetCSRFTokenField() %>
                                <input type="hidden" name="action" value="process_refund">
                                <input type="hidden" name="refundID" value="<%= refundRS("RefundID") %>">
                                <button type="submit" class="btn-small btn-process"><i class="fas fa-check"></i> 核销</button>
                            </form>
                            <% ElseIf refundRS("Status") = "Completed" Then %>
                            <small style="color: #888;">
                                <%= FormatDateField(refundRS("CompletedAt")) %>
                            </small>
                            <% End If %>
                        </td>
                    </tr>
                    <% 
                            refundRS.MoveNext
                        Loop
                        refundRS.Close
                        Set refundRS = Nothing
                    End If
                    %>
                </tbody>
            </table>
        </div>
        
        <!-- Tab 5: 对账报告 -->
        <div class="tab-content <%= IIf(currentTab="report", "active", "") %>">
            <div class="report-section">
                <div class="report-title"><i class="fas fa-calendar-day"></i> 日报视图（<%= Year(Now()) & "-" & Right("0" & Month(Now()), 2) & "-" & Right("0" & Day(Now()), 2) %>）</div>
                <div class="report-grid">
                    <div class="report-item">
                        <div class="report-value">
                            <%= GetScalar("SELECT COUNT(*) FROM ReconciliationLogs WHERE CAST(CreatedAt AS DATE)=CAST(GETDATE() AS DATE)") %>
                        </div>
                        <div class="report-label">今日对账笔数</div>
                    </div>
                    <div class="report-item">
                        <div class="report-value">
                            <% 
                            Dim todayTotal, todayMatched
                            todayTotal = GetScalar("SELECT COUNT(*) FROM ReconciliationLogs WHERE CAST(CreatedAt AS DATE)=CAST(GETDATE() AS DATE)")
                            todayMatched = GetScalar("SELECT COUNT(*) FROM ReconciliationLogs WHERE CAST(CreatedAt AS DATE)=CAST(GETDATE() AS DATE) AND Status='Matched'")
                            If todayTotal > 0 Then
                                Response.Write FormatNumber(todayMatched / todayTotal * 100, 1) & "%"
                            Else
                                Response.Write "0%"
                            End If
                            %>
                        </div>
                        <div class="report-label">匹配率</div>
                    </div>
                    <div class="report-item">
                        <div class="report-value">
                            <% 
                            Dim todayException
                            todayException = GetScalar("SELECT COUNT(*) FROM ReconciliationLogs WHERE CAST(CreatedAt AS DATE)=CAST(GETDATE() AS DATE) AND Status IN ('ShortPay', 'OverPay', 'Missing')")
                            If todayTotal > 0 Then
                                Response.Write FormatNumber(todayException / todayTotal * 100, 1) & "%"
                            Else
                                Response.Write "0%"
                            End If
                            %>
                        </div>
                        <div class="report-label">异常率</div>
                    </div>
                    <div class="report-item">
                        <div class="report-value">
                            <%= GetScalar("SELECT COUNT(*) FROM ReconciliationLogs WHERE Status IN ('ShortPay', 'OverPay', 'Missing')") %>
                        </div>
                        <div class="report-label">待处理异常</div>
                    </div>
                </div>
            </div>
            
            <div class="report-section">
                <div class="report-title"><i class="fas fa-calendar-week"></i> 周报视图（最近7天）</div>
                <div class="report-grid">
                    <div class="report-item">
                        <div class="report-value">
                            <%= GetScalar("SELECT COUNT(*) FROM ReconciliationLogs WHERE CreatedAt >= DATEADD(day, -7, CAST(GETDATE() AS DATE))") %>
                        </div>
                        <div class="report-label">本周对账笔数</div>
                    </div>
                    <div class="report-item">
                        <div class="report-value">
                            ¥<%= FormatNumber(CDbl("0" & GetScalar("SELECT IIF(SUM(OrderAmount) IS NULL, 0, SUM(OrderAmount)) FROM ReconciliationLogs WHERE CreatedAt >= DATEADD(day, -7, CAST(GETDATE() AS DATE))")), 2) %>
                        </div>
                        <div class="report-label">本周对账金额</div>
                    </div>
                    <div class="report-item">
                        <div class="report-value">
                            ¥<%= FormatNumber(CDbl("0" & GetScalar("SELECT IIF(SUM(Difference) IS NULL, 0, SUM(Difference)) FROM ReconciliationLogs WHERE CreatedAt >= DATEADD(day, -7, CAST(GETDATE() AS DATE)) AND Status IN ('ShortPay', 'OverPay')")), 2) %>
                        </div>
                        <div class="report-label">差异金额</div>
                    </div>
                    <div class="report-item">
                        <div class="report-value">
                            <%= GetScalar("SELECT COUNT(*) FROM ReconciliationLogs WHERE CreatedAt >= DATEADD(day, -7, CAST(GETDATE() AS DATE)) AND Status='Resolved'") %>
                        </div>
                        <div class="report-label">已解决异常</div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    
    <script>
        // 对账进度动画
        document.getElementById('reconcileForm')?.addEventListener('submit', function(e) {
            var container = document.getElementById('progressContainer');
            var fill = document.getElementById('progressFill');
            var text = document.getElementById('progressText');
            
            container.style.display = 'block';
            
            var progress = 0;
            var interval = setInterval(function() {
                progress += Math.random() * 15;
                if (progress >= 100) {
                    progress = 100;
                    clearInterval(interval);
                    text.textContent = '对账完成，正在保存结果...';
                } else {
                    text.textContent = '正在对账... ' + Math.floor(progress) + '%';
                }
                fill.style.width = progress + '%';
            }, 200);
        });
        
        // 结果筛选
        function filterResults() {
            var statusFilter = document.getElementById('statusFilter').value;
            var orderFilter = document.getElementById('orderFilter').value.toLowerCase();
            var rows = document.querySelectorAll('.log-row');
            
            rows.forEach(function(row) {
                var status = row.getAttribute('data-status');
                var orderNo = row.getAttribute('data-order').toLowerCase();
                
                var statusMatch = !statusFilter || status === statusFilter;
                var orderMatch = !orderFilter || orderNo.indexOf(orderFilter) > -1;
                
                row.style.display = statusMatch && orderMatch ? '' : 'none';
            });
        }
        
        // V20 快捷日期预设
        function fmtDate(d) { return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0'); }
        function setPresetDate(startOffset, endOffset) {
            var now = new Date();
            var s = new Date(now); s.setDate(s.getDate() + startOffset);
            var e = new Date(now); e.setDate(e.getDate() + endOffset);
            var ds = document.querySelectorAll('input[type=date]');
            if(ds.length>=2){ ds[0].value=fmtDate(s); ds[1].value=fmtDate(e);
                ds.forEach(function(d){ d.dispatchEvent(new Event('input',{bubbles:true})); d.dispatchEvent(new Event('change',{bubbles:true})); }); }
        }
        function setMonthPreset(monthOffset) {
            var now = new Date(); var s = new Date(now.getFullYear(), now.getMonth() + monthOffset, 1);
            var e = new Date(now.getFullYear(), now.getMonth() + monthOffset + 1, 0);
            var ds = document.querySelectorAll('input[type=date]');
            if(ds.length>=2){ ds[0].value=fmtDate(s); ds[1].value=fmtDate(e);
                ds.forEach(function(d){ d.dispatchEvent(new Event('input',{bubbles:true})); d.dispatchEvent(new Event('change',{bubbles:true})); }); }
        }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
