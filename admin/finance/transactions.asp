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

' V8：支付类型扩展
On Error Resume Next
conn.Execute "SELECT PaymentType FROM PaymentRecords WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PaymentRecords ADD PaymentType NVARCHAR(30) DEFAULT 'Receipt'"
conn.Execute "SELECT VoucherNo FROM PaymentRecords WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PaymentRecords ADD VoucherNo NVARCHAR(50)"
On Error GoTo 0

' ============================================
' 权限检查
' ============================================
Dim canEdit
If Session("AdminRoleCode") = "FIN_MANAGER" OR Session("AdminRoleCode") = "SUPER_ADMIN" Then
    canEdit = True
Else
    canEdit = False
End If

' ============================================
' 获取大额交易阈值
' ============================================
Dim largeAmountThreshold
largeAmountThreshold = CDbl("0" & GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey = 'LargeTransactionThreshold'"))
If largeAmountThreshold = 0 Then largeAmountThreshold = 5000

' ============================================
' 处理表单提交
' ============================================
Dim action
action = Request.Form("action")

' 更新分类标签
If action = "update_category" AND canEdit Then
    If Not ValidateCSRFToken() Then
        Response.Redirect "transactions.asp?tab=records&error=安全验证失败"
        Response.End
    End If
    
    Dim recordId, newCategory
    recordId = Request.Form("recordId")
    newCategory = Request.Form("category")
    
    If IsNumeric(recordId) AND newCategory <> "" Then
        Dim updateCatSQL
        updateCatSQL = "UPDATE PaymentRecords SET Category = '" & SafeSQL(newCategory) & "', UpdatedAt = GETDATE() WHERE RecordID = " & CLng(recordId)
        If ExecuteNonQuery(updateCatSQL) Then
            Call LogAdminAction("更新交易分类", "finance", "PaymentRecords", recordId, "分类改为: " & newCategory)
            Response.Redirect "transactions.asp?tab=records&msg=分类更新成功"
            Response.End
        Else
            Response.Redirect "transactions.asp?tab=records&error=分类更新失败"
            Response.End
        End If
    End If
End If

' 创建退款申请
If action = "create_refund" AND canEdit Then
    If Not ValidateCSRFToken() Then
        Response.Redirect "transactions.asp?tab=refunds&error=安全验证失败"
        Response.End
    End If
    
    Dim refundOrderId, refundAmount, refundReason
    refundOrderId = Request.Form("orderId")
    refundAmount = Request.Form("refundAmount")
    refundReason = Request.Form("refundReason")
    
    If IsNumeric(refundOrderId) AND IsNumeric(refundAmount) AND refundReason <> "" Then
        ' 获取订单信息
        Dim orderRs, orderNo, orderAmount, existingRefund
        Set orderRs = ExecuteQuery("SELECT OrderNo, TotalAmount, RefundAmount FROM Orders WHERE OrderID = " & CLng(refundOrderId))
        
        If Not orderRs Is Nothing AND Not orderRs.EOF Then
            orderNo = orderRs("OrderNo")
            orderAmount = CDbl("0" & orderRs("TotalAmount"))
            existingRefund = CDbl("0" & orderRs("RefundAmount"))
            orderRs.Close
            Set orderRs = Nothing
            
            ' 验证退款金额
            Dim maxRefund
            maxRefund = orderAmount - existingRefund
            
            If CDbl(refundAmount) > 0 AND CDbl(refundAmount) <= maxRefund Then
                ' 生成退款单号 REF+YYYYMMDDHHmmss
                Dim refundNo
                refundNo = "REF" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2)
                refundNo = refundNo & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
                
                Dim insertRefundSQL
                insertRefundSQL = "INSERT INTO RefundRecords (OrderID, OrderNo, RefundNo, RefundAmount, RefundReason, Status, CreatedAt, UpdatedAt) VALUES ("
                insertRefundSQL = insertRefundSQL & CLng(refundOrderId) & ", '" & SafeSQL(orderNo) & "', '" & SafeSQL(refundNo) & "', " & CDbl(refundAmount) & ", '" & SafeSQL(refundReason) & "', 'Pending', GETDATE(), GETDATE())"
                
                If ExecuteNonQuery(insertRefundSQL) Then
                    Call LogAdminAction("创建退款申请", "finance", "RefundRecords", "", "退款单号: " & refundNo & ", 金额: " & refundAmount)
                    Response.Redirect "transactions.asp?tab=refunds&msg=退款申请创建成功"
                    Response.End
                Else
                    Response.Redirect "transactions.asp?tab=refunds&error=退款申请创建失败"
                    Response.End
                End If
            Else
                Response.Redirect "transactions.asp?tab=refunds&error=退款金额无效"
                Response.End
            End If
        Else
            If Not orderRs Is Nothing Then
                orderRs.Close
                Set orderRs = Nothing
            End If
            Response.Redirect "transactions.asp?tab=refunds&error=订单不存在"
            Response.End
        End If
    Else
        Response.Redirect "transactions.asp?tab=refunds&error=请填写完整信息"
        Response.End
    End If
End If

' 审批退款
If action = "approve_refund" AND canEdit Then
    If Not ValidateCSRFToken() Then
        Response.Redirect "transactions.asp?tab=refunds&error=安全验证失败"
        Response.End
    End If
    
    Dim refundId, approveAction, rejectReason
    refundId = Request.Form("refundId")
    approveAction = Request.Form("approveAction")
    rejectReason = Request.Form("rejectReason")
    
    If IsNumeric(refundId) Then
        Dim updateRefundSQL, newStatus
        
        If approveAction = "approve" Then
            newStatus = "Approved"
            updateRefundSQL = "UPDATE RefundRecords SET Status = 'Approved', ApprovedBy = '" & SafeSQL(Session("AdminUserName")) & "', ApprovedAt = GETDATE(), UpdatedAt = GETDATE() WHERE RefundID = " & CLng(refundId)
        Else
            newStatus = "Rejected"
            updateRefundSQL = "UPDATE RefundRecords SET Status = 'Rejected', RefundReason = RefundReason & ' [拒绝原因: " & SafeSQL(rejectReason) & "]', UpdatedAt = GETDATE() WHERE RefundID = " & CLng(refundId)
        End If
        
        If ExecuteNonQuery(updateRefundSQL) Then
            Call LogAdminAction("审批退款", "finance", "RefundRecords", refundId, "状态改为: " & newStatus)
            Response.Redirect "transactions.asp?tab=refunds&msg=审批操作成功"
            Response.End
        Else
            Response.Redirect "transactions.asp?tab=refunds&error=审批操作失败"
            Response.End
        End If
    End If
End If

' 执行退款
If action = "execute_refund" AND canEdit Then
    If Not ValidateCSRFToken() Then
        Response.Redirect "transactions.asp?tab=refunds&error=安全验证失败"
        Response.End
    End If
    
    Dim execRefundId
    execRefundId = Request.Form("refundId")
    
    If IsNumeric(execRefundId) Then
        ' 获取退款信息
        Dim refundRs, refundOrderNo  ' 其他变量已声明: refundOrderId, refundAmount, refundNo
        Set refundRs = ExecuteQuery("SELECT * FROM RefundRecords WHERE RefundID = " & CLng(execRefundId) & " AND Status = 'Approved' AND CostWriteBack = 0")
        
        If Not refundRs Is Nothing AND Not refundRs.EOF Then
            refundOrderId = refundRs("OrderID")
            refundNo = refundRs("RefundNo")
            refundAmount = CDbl("0" & refundRs("RefundAmount"))
            refundOrderNo = refundRs("OrderNo")
            refundRs.Close
            Set refundRs = Nothing
            
            ' 开始事务处理
            On Error Resume Next
            
            ' 1. 更新订单退款金额
            Dim updateOrderSQL
            updateOrderSQL = "UPDATE Orders SET RefundAmount = RefundAmount + " & refundAmount & ", UpdatedAt = GETDATE() WHERE OrderID = " & CLng(refundOrderId)
            ExecuteNonQuery(updateOrderSQL)
            
            ' 2. 在 PaymentRecords 插入负向记录
            Dim insertPaymentSQL
            insertPaymentSQL = "INSERT INTO PaymentRecords (OrderID, OrderNo, TransactionNo, PaymentMethod, TransactionType, Amount, Fee, NetAmount, Status, Category, Remark, CreatedAt, UpdatedAt) VALUES ("
            insertPaymentSQL = insertPaymentSQL & CLng(refundOrderId) & ", '" & SafeSQL(refundOrderNo) & "', '" & SafeSQL(refundNo) & "', 'System', 'Refund', " & (-refundAmount) & ", 0, " & (-refundAmount) & ", 'Completed', '退款支出', '退款单: " & SafeSQL(refundNo) & "', GETDATE(), GETDATE())"
            ExecuteNonQuery(insertPaymentSQL)
            
            ' 3. 更新退款记录状态
            Dim completeRefundSQL
            completeRefundSQL = "UPDATE RefundRecords SET Status = 'Completed', CompletedAt = GETDATE(), CostWriteBack = 1, UpdatedAt = GETDATE() WHERE RefundID = " & CLng(execRefundId)
            ExecuteNonQuery(completeRefundSQL)
            
            If Err.Number = 0 Then
                Call LogAdminAction("执行退款", "finance", "RefundRecords", execRefundId, "退款单号: " & refundNo)
                Response.Redirect "transactions.asp?tab=refunds&msg=退款执行成功"
                Response.End
            Else
                Response.Redirect "transactions.asp?tab=refunds&error=退款执行失败: " & Server.HTMLEncode(Err.Description)
                Response.End
            End If
        Else
            If Not refundRs Is Nothing Then
                refundRs.Close
                Set refundRs = Nothing
            End If
            Response.Redirect "transactions.asp?tab=refunds&error=退款记录不存在或状态不正确"
            Response.End
        End If
    End If
End If

' 标记异常已处理
If action = "resolve_anomaly" AND canEdit Then
    If Not ValidateCSRFToken() Then
        Response.Redirect "transactions.asp?tab=anomalies&error=安全验证失败"
        Response.End
    End If
    
    Dim anomalyId
    anomalyId = Request.Form("anomalyId")
    
    ' 这里可以创建异常处理记录表，暂时使用Session标记
    Session("ResolvedAnomaly_" & anomalyId) = True
    
    Response.Redirect "transactions.asp?tab=anomalies&msg=异常已标记为已处理"
    Response.End
End If

' ============================================
' 获取当前Tab
' ============================================
Dim currentTab
currentTab = Request.QueryString("tab")
If currentTab = "" Then currentTab = "records"

' ============================================
' Tab 1: 支付流水记录参数
' ============================================
Dim filterStartDate, filterEndDate, filterType, filterMethod, filterStatus
Dim filterMinAmount, filterMaxAmount, filterSearch, pageNum

filterStartDate = Request.QueryString("startDate")
filterEndDate = Request.QueryString("endDate")
filterType = Request.QueryString("type")
filterMethod = Request.QueryString("method")
filterStatus = Request.QueryString("status")
filterMinAmount = Request.QueryString("minAmount")
filterMaxAmount = Request.QueryString("maxAmount")
filterSearch = Request.QueryString("search")
pageNum = Request.QueryString("page")
If Not IsNumeric(pageNum) OR pageNum = "" Then pageNum = 1
pageNum = CLng(pageNum)

' ============================================
' Tab 2: 退款管理参数
' ============================================
Dim refundStatus, refundPageNum
refundStatus = Request.QueryString("refundStatus")
refundPageNum = Request.QueryString("refundPage")
If Not IsNumeric(refundPageNum) OR refundPageNum = "" Then refundPageNum = 1
refundPageNum = CLng(refundPageNum)

' ============================================
' 智能分类函数
' ============================================
Function AutoClassify(transactionType, remark)
    Dim category
    category = ""
    
    Select Case transactionType
        Case "Payment"
            category = "订单收入"
        Case "Refund"
            category = "退款支出"
        Case "Fee"
            category = "平台扣费"
        Case "Transfer"
            category = "资金划转"
    End Select
    
    ' 检查备注关键词
    If InStr(LCase(remark), "推广") > 0 OR InStr(LCase(remark), "广告") > 0 Then
        category = "营销费用"
    End If
    
    AutoClassify = category
End Function

' ============================================
' 构建流水查询SQL
' ============================================
Function BuildRecordSQL(isCount)
    Dim sql
    
    If isCount Then
        sql = "SELECT COUNT(*) FROM PaymentRecords WHERE 1=1"
    Else
        sql = "SELECT * FROM PaymentRecords WHERE 1=1"
    End If
    
    ' 时间范围筛选
    If filterStartDate <> "" Then
        sql = sql & " AND CreatedAt >= #" & SafeSQL(filterStartDate) & " 00:00:00#"
    End If
    If filterEndDate <> "" Then
        sql = sql & " AND CreatedAt <= #" & SafeSQL(filterEndDate) & " 23:59:59#"
    End If
    
    ' 交易类型筛选
    If filterType <> "" Then
        sql = sql & " AND TransactionType = '" & SafeSQL(filterType) & "'"
    End If
    
    ' 支付方式筛选
    If filterMethod <> "" Then
        sql = sql & " AND PaymentMethod = '" & SafeSQL(filterMethod) & "'"
    End If
    
    ' 状态筛选
    If filterStatus <> "" Then
        sql = sql & " AND Status = '" & SafeSQL(filterStatus) & "'"
    End If
    
    ' 金额范围筛选
    If IsNumeric(filterMinAmount) Then
        sql = sql & " AND ABS(Amount) >= " & CDbl(filterMinAmount)
    End If
    If IsNumeric(filterMaxAmount) Then
        sql = sql & " AND ABS(Amount) <= " & CDbl(filterMaxAmount)
    End If
    
    ' 搜索
    If filterSearch <> "" Then
        sql = sql & " AND (OrderNo LIKE '%" & SafeSQL(filterSearch) & "%' OR TransactionNo LIKE '%" & SafeSQL(filterSearch) & "%')"
    End If
    
    ' 排序
    If Not isCount Then
        sql = sql & " ORDER BY CreatedAt DESC"
    End If
    
    BuildRecordSQL = sql
End Function

' ============================================
' 构建退款查询SQL
' ============================================
Function BuildRefundSQL(isCount)
    Dim sql
    
    If isCount Then
        sql = "SELECT COUNT(*) FROM RefundRecords WHERE 1=1"
    Else
        sql = "SELECT * FROM RefundRecords WHERE 1=1"
    End If
    
    ' 状态筛选
    If refundStatus <> "" Then
        sql = sql & " AND Status = '" & SafeSQL(refundStatus) & "'"
    End If
    
    ' 排序
    If Not isCount Then
        sql = sql & " ORDER BY CreatedAt DESC"
    End If
    
    BuildRefundSQL = sql
End Function

Call LogAdminAction("查看交易管理", "finance", "", "", "Tab: " & currentTab)
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>流水管理与退款 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        /* Tab 导航样式 */
        .tab-nav { display: flex; gap: 5px; margin-bottom: 25px; border-bottom: 2px solid #3a3a4a; padding-bottom: 0; }
        .tab-nav a { 
            padding: 15px 30px; color: #888; text-decoration: none; 
            border-bottom: 3px solid transparent; margin-bottom: -2px;
            transition: all 0.3s; font-weight: 500;
        }
        .tab-nav a:hover { color: #e0e0e0; }
        .tab-nav a.active { color: #00bcd4; border-bottom-color: #00bcd4; }
        .tab-nav a i { margin-right: 8px; }
        
        /* 筛选栏样式 */
        .filter-bar { 
            background: #2d2d44; padding: 20px; border-radius: 12px; 
            margin-bottom: 25px; border: 1px solid rgba(255,255,255,0.06);
        }
        .filter-row { display: flex; gap: 15px; flex-wrap: wrap; align-items: end; }
        .filter-group { display: flex; flex-direction: column; gap: 5px; }
        .filter-group label { color: #888; font-size: 12px; font-weight: 500; }
        .filter-group input, .filter-group select { 
            padding: 10px 15px; border: 1px solid rgba(255,255,255,0.15); border-radius: 8px; 
            background: #1a1a2e; color: #e0e0e0; font-size: 14px; min-width: 140px;
        }
        .filter-group input:focus, .filter-group select:focus { 
            outline: none; border-color: #00bcd4; 
        }
        .btn-filter { 
            background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); 
            color: white; padding: 10px 25px; border: none; border-radius: 8px; 
            cursor: pointer; font-size: 14px; transition: opacity 0.3s;
        }
        .btn-filter:hover { opacity: 0.9; }
        .btn-reset { 
            background: #3a3a4a; color: #e0e0e0; padding: 10px 25px; 
            border: none; border-radius: 8px; cursor: pointer; font-size: 14px;
        }
        .btn-reset:hover { background: #4a4a5a; }
        
        /* 汇总栏样式 */
        .summary-bar { 
            display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; 
            margin-bottom: 25px;
        }
        .summary-card { 
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 20px; border-radius: 12px; 
            border: 1px solid rgba(255,255,255,0.06); text-align: center;
        }
        .summary-card .label { color: #888; font-size: 13px; margin-bottom: 8px; }
        .summary-card .value { font-size: 24px; font-weight: bold; }
        .summary-card.income .value { color: #4CAF50; }
        .summary-card.expense .value { color: #f44336; }
        .summary-card.fee .value { color: #ff9800; }
        .summary-card.net .value { color: #00bcd4; }
        
        /* 表格样式 */
        .data-table { width: 100%; border-collapse: collapse; background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; overflow: hidden; }
        .data-table th, .data-table td { padding: 15px; text-align: left; border-bottom: 1px solid rgba(255,255,255,0.06); }
        .data-table th { background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); color: white; font-weight: 600; font-size: 13px; }
        .data-table td { color: #e0e0e0; font-size: 14px; }
        .data-table tr:hover { background: rgba(255,255,255,0.05); }
        .data-table tr.anomaly { background: rgba(255, 193, 7, 0.1); }
        .data-table tr.anomaly:hover { background: rgba(255, 193, 7, 0.2); }
        
        /* 金额颜色 */
        .amount-income { color: #4CAF50; font-weight: 600; }
        .amount-expense { color: #f44336; font-weight: 600; }
        
        /* 状态徽章 */
        .status-badge { 
            display: inline-block; padding: 4px 12px; border-radius: 12px; 
            font-size: 12px; font-weight: 500;
        }
        .status-pending { background: #424242; color: #9e9e9e; }
        .status-approved { background: #1b5e20; color: #81c784; }
        .status-rejected { background: #5e1b1b; color: #e57373; }
        .status-completed { background: #0d47a1; color: #64b5f6; }
        .status-success { background: #1b5e20; color: #81c784; }
        .status-failed { background: #5e1b1b; color: #e57373; }
        
        /* 分类标签 */
        .category-tag { 
            display: inline-block; padding: 3px 10px; border-radius: 10px; 
            font-size: 12px; background: #3a3a4a; color: #b0b0b0;
        }
        .category-income { background: #1b5e20; color: #81c784; }
        .category-expense { background: #5e1b1b; color: #e57373; }
        .category-fee { background: #5d4037; color: #ffcc80; }
        .category-marketing { background: #4a148c; color: #ce93d8; }
        
        /* 分页样式 */
        .pagination { display: flex; justify-content: center; gap: 5px; margin-top: 25px; }
        .pagination a, .pagination span { 
            padding: 10px 15px; border-radius: 8px; text-decoration: none; 
            color: #888; background: #2a2a3a; border: 1px solid #3a3a4a;
        }
        .pagination a:hover { background: #3a3a4a; color: #e0e0e0; }
        .pagination .current { background: #00bcd4; color: white; border-color: #00bcd4; }
        

        /* 模态框样式 */
        .modal { 
            display: none; position: fixed; top: 0; left: 0; right: 0; bottom: 0; 
            background: rgba(0,0,0,0.7); z-index: 1000; align-items: center; justify-content: center;
        }
        .modal.active { display: flex; }
        .modal-content { 
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 16px; width: 90%; max-width: 600px; 
            max-height: 90vh; overflow-y: auto; border: 1px solid rgba(255,255,255,0.06);
        }
        .modal-header { 
            padding: 20px 25px; border-bottom: 1px solid rgba(255,255,255,0.06); 
            display: flex; justify-content: space-between; align-items: center;
        }
        .modal-header h3 { margin: 0; color: #e0e0e0; }
        .modal-close { 
            background: none; border: none; color: #888; font-size: 24px; 
            cursor: pointer; padding: 0; width: 30px; height: 30px;
        }
        .modal-close:hover { color: #e0e0e0; }
        .modal-body { padding: 25px; }
        .modal-footer { 
            padding: 20px 25px; border-top: 1px solid rgba(255,255,255,0.06); 
            display: flex; justify-content: flex-end; gap: 10px;
        }
        
        /* 表单样式 */
        .form-group { margin-bottom: 20px; }
        .form-group label { 
            display: block; margin-bottom: 8px; color: #b0b0b0; font-weight: 500; 
        }
        .form-group input, .form-group select, .form-group textarea { 
            width: 100%; padding: 12px 15px; border: 1px solid #3a3a4a; 
            border-radius: 8px; background: #1a1a2e; color: #e0e0e0; font-size: 14px;
            box-sizing: border-box;
        }
        .form-group input:focus, .form-group select:focus, .form-group textarea:focus { 
            outline: none; border-color: #00bcd4; 
        }
        .form-group textarea { resize: vertical; min-height: 100px; }
        .form-group .help-text { font-size: 12px; color: #888; margin-top: 5px; }
        
        /* 异常监控样式 */
        .anomaly-card { 
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; padding: 20px; 
            margin-bottom: 15px; border: 1px solid rgba(255,255,255,0.06); border-left: 4px solid #ffc107;
        }
        .anomaly-card.resolved { border-left-color: #4CAF50; opacity: 0.7; }
        .anomaly-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
        .anomaly-type { 
            background: rgba(255, 193, 7, 0.2); color: #ffc107; 
            padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 500;
        }
        .anomaly-time { color: #888; font-size: 13px; }
        .anomaly-body { color: #e0e0e0; margin-bottom: 10px; }
        .anomaly-amount { font-size: 20px; font-weight: bold; color: #f44336; }
        
        /* 空状态 */
        .empty-state { text-align: center; padding: 60px 20px; color: #888; }
        .empty-state i { font-size: 48px; margin-bottom: 15px; color: #555; }
        
        /* 响应式 */
        @media (max-width: 1200px) {
            .summary-bar { grid-template-columns: repeat(2, 1fr); }
            .filter-row { flex-direction: column; align-items: stretch; }
            .filter-group { width: 100%; }
            .filter-group input, .filter-group select { width: 100%; }
        }
        @media (max-width: 768px) {
            .summary-bar { grid-template-columns: 1fr; }
            .tab-nav { flex-wrap: wrap; }
            .tab-nav a { padding: 10px 20px; font-size: 13px; }
            .data-table { font-size: 12px; }
            .data-table th, .data-table td { padding: 10px 8px; }
        }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-exchange-alt"></i> 流水管理与退款</h2>
            <div class="breadcrumb">
                <a href="index.asp">财务中心</a> / <span>流水管理</span>
            </div>
        </div>
        
        <% If Request.QueryString("msg") <> "" Then %>
        <div class="alert alert-success"><i class="fas fa-check-circle"></i> <%= Server.HTMLEncode(Request.QueryString("msg")) %></div>
        <% End If %>
        
        <% If Request.QueryString("error") <> "" Then %>
        <div class="alert alert-error"><i class="fas fa-times-circle"></i> <%= Server.HTMLEncode(Request.QueryString("error")) %></div>
        <% End If %>
        
        <!-- Tab 导航 -->
        <div class="tab-nav">
            <a href="?tab=records" class="<%= IIF(currentTab="records", "active", "") %>"><i class="fas fa-list-alt"></i> 支付流水</a>
            <a href="?tab=refunds" class="<%= IIF(currentTab="refunds", "active", "") %>"><i class="fas fa-undo"></i> 退款管理</a>
            <a href="?tab=anomalies" class="<%= IIF(currentTab="anomalies", "active", "") %>"><i class="fas fa-exclamation-triangle"></i> 异常监控</a>
        </div>
        
        <% If currentTab = "records" Then %>
        <!-- ======================================== -->
        <!-- Tab 1: 支付流水记录 -->
        <!-- ======================================== -->
        
        <!-- 筛选栏 -->
        <div class="filter-bar">
            <form method="get" action="">
                <input type="hidden" name="tab" value="records">
                <div class="filter-row">
                    <div class="filter-group">
                        <label>开始日期</label>
                        <input type="date" name="startDate" value="<%= Server.HTMLEncode(filterStartDate) %>">
                    </div>
                    <div class="filter-group">
                        <label>结束日期</label>
                        <input type="date" name="endDate" value="<%= Server.HTMLEncode(filterEndDate) %>">
                    </div>
                    <div class="filter-group">
                        <label>交易类型</label>
                        <select name="type">
                            <option value="">全部</option>
                            <option value="Payment" <%= IIF(filterType="Payment", "selected", "") %>>支付</option>
                            <option value="Refund" <%= IIF(filterType="Refund", "selected", "") %>>退款</option>
                            <option value="Transfer" <%= IIF(filterType="Transfer", "selected", "") %>>转账</option>
                            <option value="Fee" <%= IIF(filterType="Fee", "selected", "") %>>手续费</option>
                        </select>
                    </div>
                    <div class="filter-group">
                        <label>支付方式</label>
                        <select name="method">
                            <option value="">全部</option>
                            <option value="Alipay" <%= IIF(filterMethod="Alipay", "selected", "") %>>支付宝</option>
                            <option value="Wechat" <%= IIF(filterMethod="Wechat", "selected", "") %>>微信支付</option>
                            <option value="Bank" <%= IIF(filterMethod="Bank", "selected", "") %>>银行转账</option>
                            <option value="Stripe" <%= IIF(filterMethod="Stripe", "selected", "") %>>Stripe</option>
                            <option value="UnionPay" <%= IIF(filterMethod="UnionPay", "selected", "") %>>银联</option>
                            <option value="PayPal" <%= IIF(filterMethod="PayPal", "selected", "") %>>PayPal</option>
                            <option value="System" <%= IIF(filterMethod="System", "selected", "") %>>系统</option>
                        </select>
                    </div>
                    <div class="filter-group">
                        <label>状态</label>
                        <select name="status">
                            <option value="">全部</option>
                            <option value="Completed" <%= IIF(filterStatus="Completed", "selected", "") %>>已完成</option>
                            <option value="Pending" <%= IIF(filterStatus="Pending", "selected", "") %>>处理中</option>
                            <option value="Failed" <%= IIF(filterStatus="Failed", "selected", "") %>>失败</option>
                        </select>
                    </div>
                    <div class="filter-group">
                        <label>最小金额</label>
                        <input type="number" name="minAmount" value="<%= Server.HTMLEncode(filterMinAmount) %>" placeholder="0" step="0.01">
                    </div>
                    <div class="filter-group">
                        <label>最大金额</label>
                        <input type="number" name="maxAmount" value="<%= Server.HTMLEncode(filterMaxAmount) %>" placeholder="不限" step="0.01">
                    </div>
                    <div class="filter-group">
                        <label>搜索</label>
                        <input type="text" name="search" value="<%= Server.HTMLEncode(filterSearch) %>" placeholder="订单号/流水号">
                    </div>
                    <button type="submit" class="btn-filter"><i class="fas fa-search"></i> 筛选</button>
                    <a href="?tab=records" class="btn-reset"><i class="fas fa-undo"></i> 重置</a>
                </div>
            </form>
        </div>
        
        <% 
        ' 获取统计数据
        Dim totalIncome, totalExpense, totalFee, totalNet
        Dim statSQL
        
        statSQL = "SELECT SUM(IIF(TransactionType='Payment' AND Amount>0, Amount, 0)) as Income, " & _
                  "SUM(IIF(TransactionType='Refund' OR Amount<0, ABS(Amount), 0)) as Expense, " & _
                  "SUM(IIF(Fee IS NOT NULL, Fee, 0)) as FeeTotal, " & _
                  "SUM(NetAmount) as NetTotal FROM PaymentRecords WHERE 1=1"
        
        If filterStartDate <> "" Then
            statSQL = statSQL & " AND CreatedAt >= #" & SafeSQL(filterStartDate) & " 00:00:00#"
        End If
        If filterEndDate <> "" Then
            statSQL = statSQL & " AND CreatedAt <= #" & SafeSQL(filterEndDate) & " 23:59:59#"
        End If
        If filterType <> "" Then
            statSQL = statSQL & " AND TransactionType = '" & SafeSQL(filterType) & "'"
        End If
        If filterMethod <> "" Then
            statSQL = statSQL & " AND PaymentMethod = '" & SafeSQL(filterMethod) & "'"
        End If
        If filterStatus <> "" Then
            statSQL = statSQL & " AND Status = '" & SafeSQL(filterStatus) & "'"
        End If
        If filterSearch <> "" Then
            statSQL = statSQL & " AND (OrderNo LIKE '%" & SafeSQL(filterSearch) & "%' OR TransactionNo LIKE '%" & SafeSQL(filterSearch) & "%')"
        End If
        
        Dim statRs
        Set statRs = ExecuteQuery(statSQL)
        If Not statRs Is Nothing Then
            If Not statRs.EOF Then
                totalIncome = SafeNum(statRs("Income"))
                totalExpense = SafeNum(statRs("Expense"))
                totalFee = SafeNum(statRs("FeeTotal"))
                totalNet = SafeNum(statRs("NetTotal"))
            End If
            statRs.Close
        End If
        Set statRs = Nothing
        %>
        
        <!-- 汇总栏 -->
        <div class="summary-bar">
            <div class="summary-card income">
                <div class="label"><i class="fas fa-arrow-up"></i> 总收入</div>
                <div class="value">¥<%= FormatNumber(totalIncome, 2) %></div>
            </div>
            <div class="summary-card expense">
                <div class="label"><i class="fas fa-arrow-down"></i> 总支出</div>
                <div class="value">¥<%= FormatNumber(totalExpense, 2) %></div>
            </div>
            <div class="summary-card fee">
                <div class="label"><i class="fas fa-percentage"></i> 总手续费</div>
                <div class="value">¥<%= FormatNumber(totalFee, 2) %></div>
            </div>
            <div class="summary-card net">
                <div class="label"><i class="fas fa-wallet"></i> 净额</div>
                <div class="value">¥<%= FormatNumber(totalNet, 2) %></div>
            </div>
        </div>
        
        <% 
        ' 获取记录总数和分页
        Dim recordCount, totalPages
        recordCount = CLng("0" & GetScalar(BuildRecordSQL(True)))
        totalPages = Int((recordCount + PAGE_SIZE - 1) / PAGE_SIZE)
        If totalPages < 1 Then totalPages = 1
        If pageNum > totalPages Then pageNum = totalPages
        If pageNum < 1 Then pageNum = 1
        
        ' 获取记录列表
        Dim recordRs
        Dim recordSQL
        recordSQL = BuildRecordSQL(False)
        
        ' Access分页使用TOP语法
        Dim startRow, endRow
        startRow = (pageNum - 1) * PAGE_SIZE
        endRow = pageNum * PAGE_SIZE
        
        ' 使用子查询实现分页
        recordSQL = "SELECT TOP " & PAGE_SIZE & " * FROM (" & recordSQL & ") AS T WHERE RecordID NOT IN (SELECT TOP " & startRow & " RecordID FROM (" & BuildRecordSQL(False) & ") AS T2)"
        
        Set recordRs = ExecuteQuery(recordSQL)
        %>
        
        <!-- 记录表格 -->
        <table class="data-table">
            <thead>
                <tr>
                    <th>日期</th>
                    <th>订单号</th>
                    <th>流水号</th>
                    <th>类型</th>
                    <th>支付方式</th>
                    <th>金额</th>
                    <th>手续费</th>
                    <th>净额</th>
                    <th>分类标签</th>
                    <th>状态</th>
                    <% If canEdit Then %><th>操作</th><% End If %>
                </tr>
            </thead>
            <tbody>
                <% 
                ' 循环变量声明移到循环外（VBScript限制）
                Dim recordAmount, isAnomaly, transHour
                If Not recordRs Is Nothing Then
                    Do While Not recordRs.EOF
                        recordAmount = CDbl("0" & recordRs("Amount"))
                        
                        ' 检测异常：大额交易或非工作时间
                        isAnomaly = False
                        If Abs(recordAmount) > largeAmountThreshold Then isAnomaly = True
                        transHour = Hour(recordRs("CreatedAt"))
                        If transHour >= 22 OR transHour < 8 Then isAnomaly = True
                %>
                <tr <%= IIF(isAnomaly, "class='anomaly'", "") %>>
                    <td><%= SafeFormatDateTime(recordRs("CreatedAt"), 2) %><br><small style="color:#888;"><%= SafeFormatDateTime(recordRs("CreatedAt"), 4) %></small></td>
                    <td><%= Server.HTMLEncode(recordRs("OrderNo")) %></td>
                    <td><%= Server.HTMLEncode(recordRs("TransactionNo")) %></td>
                    <td>
                        <% Select Case recordRs("TransactionType")
                           Case "Payment": Response.Write "<span style='color:#4CAF50;'>支付</span>"
                           Case "Refund": Response.Write "<span style='color:#f44336;'>退款</span>"
                           Case "Transfer": Response.Write "<span style='color:#2196F3;'>转账</span>"
                           Case "Fee": Response.Write "<span style='color:#ff9800;'>手续费</span>"
                           Case Else: Response.Write Server.HTMLEncode(recordRs("TransactionType"))
                        End Select %>
                    </td>
                    <td>
                        <% Select Case recordRs("PaymentMethod")
                           Case "Alipay": Response.Write "<i class='fab fa-alipay' style='color:#1677ff;'></i> 支付宝"
                           Case "Wechat": Response.Write "<i class='fab fa-weixin' style='color:#07c160;'></i> 微信"
                           Case "Bank": Response.Write "<i class='fas fa-university' style='color:#ff6b6b;'></i> 银行"
                           Case "Stripe": Response.Write "<i class='fab fa-stripe' style='color:#635bff;'></i> Stripe"
                           Case "UnionPay": Response.Write "<i class='fas fa-credit-card' style='color:#c00;'></i> 银联"
                           Case "PayPal": Response.Write "<i class='fab fa-paypal' style='color:#003087;'></i> PayPal"
                           Case "System": Response.Write "<i class='fas fa-cog' style='color:#888;'></i> 系统"
                           Case Else: Response.Write Server.HTMLEncode(recordRs("PaymentMethod"))
                        End Select %>
                    </td>
                    <td class="<%= IIF(recordAmount >= 0, "amount-income", "amount-expense") %>">
                        <%= IIF(recordAmount >= 0, "+", "") %><%= FormatNumber(recordAmount, 2) %>
                    </td>
                    <td><%= FormatNumber(CDbl("0" & recordRs("Fee")), 2) %></td>
                    <td><%= FormatNumber(CDbl("0" & recordRs("NetAmount")), 2) %></td>
                    <td>
                        <% 
                        Dim displayCategory
                        displayCategory = recordRs("Category")
                        If IsNull(displayCategory) OR displayCategory = "" Then
                            displayCategory = AutoClassify(recordRs("TransactionType"), recordRs("Remark"))
                        End If
                        
                        Dim catClass
                        Select Case displayCategory
                            Case "订单收入": catClass = "category-income"
                            Case "退款支出": catClass = "category-expense"
                            Case "平台扣费": catClass = "category-fee"
                            Case "营销费用": catClass = "category-marketing"
                            Case Else: catClass = ""
                        End Select
                        %>
                        <span class="category-tag <%= catClass %>"><%= Server.HTMLEncode(displayCategory) %></span>
                    </td>
                    <td>
                        <% Select Case recordRs("Status")
                           Case "Completed": Response.Write "<span class='status-badge status-completed'>已完成</span>"
                           Case "Pending": Response.Write "<span class='status-badge status-pending'>处理中</span>"
                           Case "Failed": Response.Write "<span class='status-badge status-failed'>失败</span>"
                           Case "Success": Response.Write "<span class='status-badge status-success'>成功</span>"
                           Case Else: Response.Write "<span class='status-badge status-pending'>" & Server.HTMLEncode(recordRs("Status")) & "</span>"
                        End Select %>
                    </td>
                    <% If canEdit Then %>
                    <td>
                        <button class="btn btn-sm btn-secondary" onclick="openCategoryModal(<%= recordRs("RecordID") %>, '<%= SafeOutput(displayCategory) %>')">
                            <i class="fas fa-tag"></i> 改分类
                        </button>
                    </td>
                    <% End If %>
                </tr>
                <% 
                        recordRs.MoveNext
                    Loop
                    recordRs.Close
                End If
                Set recordRs = Nothing
                %>
            </tbody>
        </table>
        
        <% If recordCount = 0 Then %>
        <div class="empty-state">
            <i class="fas fa-inbox"></i>
            <p>暂无支付流水记录</p>
        </div>
        <% End If %>
        
        <!-- 分页 -->
        <% If totalPages > 1 Then %>
        <div class="pagination">
            <% If pageNum > 1 Then %>
            <a href="?tab=records&page=<%= pageNum-1 %>&startDate=<%= Server.HTMLEncode(filterStartDate) %>&endDate=<%= Server.HTMLEncode(filterEndDate) %>&type=<%= Server.HTMLEncode(filterType) %>&method=<%= Server.HTMLEncode(filterMethod) %>&status=<%= Server.HTMLEncode(filterStatus) %>&minAmount=<%= Server.HTMLEncode(filterMinAmount) %>&maxAmount=<%= Server.HTMLEncode(filterMaxAmount) %>&search=<%= Server.HTMLEncode(filterSearch) %>"><i class="fas fa-chevron-left"></i></a>
            <% End If %>
            
            <% 
            Dim pageStart, pageEnd, p
            pageStart = pageNum - 2
            pageEnd = pageNum + 2
            If pageStart < 1 Then pageStart = 1
            If pageEnd > totalPages Then pageEnd = totalPages
            
            For p = pageStart To pageEnd
            %>
            <% If p = pageNum Then %>
            <span class="current"><%= p %></span>
            <% Else %>
            <a href="?tab=records&page=<%= p %>&startDate=<%= Server.HTMLEncode(filterStartDate) %>&endDate=<%= Server.HTMLEncode(filterEndDate) %>&type=<%= Server.HTMLEncode(filterType) %>&method=<%= Server.HTMLEncode(filterMethod) %>&status=<%= Server.HTMLEncode(filterStatus) %>&minAmount=<%= Server.HTMLEncode(filterMinAmount) %>&maxAmount=<%= Server.HTMLEncode(filterMaxAmount) %>&search=<%= Server.HTMLEncode(filterSearch) %>"><%= p %></a>
            <% End If %>
            <% Next %>
            
            <% If pageNum < totalPages Then %>
            <a href="?tab=records&page=<%= pageNum+1 %>&startDate=<%= Server.HTMLEncode(filterStartDate) %>&endDate=<%= Server.HTMLEncode(filterEndDate) %>&type=<%= Server.HTMLEncode(filterType) %>&method=<%= Server.HTMLEncode(filterMethod) %>&status=<%= Server.HTMLEncode(filterStatus) %>&minAmount=<%= Server.HTMLEncode(filterMinAmount) %>&maxAmount=<%= Server.HTMLEncode(filterMaxAmount) %>&search=<%= Server.HTMLEncode(filterSearch) %>"><i class="fas fa-chevron-right"></i></a>
            <% End If %>
        </div>
        <% End If %>
        
        <% End If ' End Tab 1 %>
        
        <% If currentTab = "refunds" Then %>
        <!-- ======================================== -->
        <!-- Tab 2: 退款管理 -->
        <!-- ======================================== -->
        
        <!-- 操作栏 -->
        <div style="margin-bottom: 25px; display: flex; justify-content: space-between; align-items: center;">
            <div class="filter-group" style="margin:0;">
                <select id="refundStatusFilter" onchange="location.href='?tab=refunds&refundStatus='+this.value" style="min-width: 150px;">
                    <option value="">全部状态</option>
                    <option value="Pending" <%= IIF(refundStatus="Pending", "selected", "") %>>待审批</option>
                    <option value="Approved" <%= IIF(refundStatus="Approved", "selected", "") %>>已批准</option>
                    <option value="Rejected" <%= IIF(refundStatus="Rejected", "selected", "") %>>已拒绝</option>
                    <option value="Completed" <%= IIF(refundStatus="Completed", "selected", "") %>>已完成</option>
                </select>
            </div>
            <% If canEdit Then %>
            <button class="btn btn-primary" onclick="openRefundModal()">
                <i class="fas fa-plus"></i> 新建退款申请
            </button>
            <% End If %>
        </div>
        
        <% 
        ' 获取退款记录总数和分页
        Dim refundCount, refundTotalPages
        refundCount = CLng("0" & GetScalar(BuildRefundSQL(True)))
        refundTotalPages = Int((refundCount + PAGE_SIZE - 1) / PAGE_SIZE)
        If refundTotalPages < 1 Then refundTotalPages = 1
        If refundPageNum > refundTotalPages Then refundPageNum = refundTotalPages
        If refundPageNum < 1 Then refundPageNum = 1
        
        ' 获取退款记录列表
        Dim refundSQL
        refundSQL = BuildRefundSQL(False)
        
        ' Access分页
        Dim refundStartRow
        refundStartRow = (refundPageNum - 1) * PAGE_SIZE
        
        refundSQL = "SELECT TOP " & PAGE_SIZE & " * FROM (" & refundSQL & ") AS T WHERE RefundID NOT IN (SELECT TOP " & refundStartRow & " RefundID FROM (" & BuildRefundSQL(False) & ") AS T2)"
        
        Set refundRs = ExecuteQuery(refundSQL)
        %>
        
        <!-- 退款记录表格 -->
        <table class="data-table">
            <thead>
                <tr>
                    <th>退款单号</th>
                    <th>订单号</th>
                    <th>退款金额</th>
                    <th>退款原因</th>
                    <th>状态</th>
                    <th>申请人</th>
                    <th>申请时间</th>
                    <th>审批人</th>
                    <th>审批时间</th>
                    <% If canEdit Then %><th>操作</th><% End If %>
                </tr>
            </thead>
            <tbody>
                <% 
                If Not refundRs Is Nothing Then
                    Dim refundReason2
                    Do While Not refundRs.EOF
                %>
                <tr>
                    <td><code><%= Server.HTMLEncode(refundRs("RefundNo")) %></code></td>
                    <td><%= Server.HTMLEncode(refundRs("OrderNo")) %></td>
                    <td class="amount-expense">-¥<%= FormatNumber(CDbl("0" & refundRs("RefundAmount")), 2) %></td>
                    <td>
                        <% 
                        refundReason2 = refundRs("RefundReason")
                        If Len(refundReason2) > 30 Then
                            Response.Write Server.HTMLEncode(Left(refundReason2, 30)) & "..."
                        Else
                            Response.Write Server.HTMLEncode(refundReason2)
                        End If
                        %>
                    </td>
                    <td>
                        <% Select Case refundRs("Status")
                           Case "Pending": Response.Write "<span class='status-badge status-pending'>待审批</span>"
                           Case "Approved": Response.Write "<span class='status-badge status-approved'>已批准</span>"
                           Case "Rejected": Response.Write "<span class='status-badge status-rejected'>已拒绝</span>"
                           Case "Completed": Response.Write "<span class='status-badge status-completed'>已完成</span>"
                           Case Else: Response.Write "<span class='status-badge status-pending'>" & Server.HTMLEncode(refundRs("Status")) & "</span>"
                        End Select %>
                    </td>
                    <td>-</td>
                    <td><%= SafeFormatDateTime(refundRs("CreatedAt"), 2) %></td>
                    <td><%= IIF(IsNull(refundRs("ApprovedBy")), "-", Server.HTMLEncode(refundRs("ApprovedBy"))) %></td>
                    <td><%= IIF(IsNull(refundRs("ApprovedAt")), "-", SafeFormatDateTime(refundRs("ApprovedAt"), 2)) %></td>
                    <% If canEdit Then %>
                    <td>
                        <% If refundRs("Status") = "Pending" Then %>
                        <button class="btn btn-sm btn-success" onclick="openApproveModal(<%= refundRs("RefundID") %>, 'approve')">
                            <i class="fas fa-check"></i> 通过
                        </button>
                        <button class="btn btn-sm btn-danger" onclick="openApproveModal(<%= refundRs("RefundID") %>, 'reject')">
                            <i class="fas fa-times"></i> 拒绝
                        </button>
                        <% ElseIf refundRs("Status") = "Approved" AND refundRs("CostWriteBack") = 0 Then %>
                        <button class="btn btn-sm btn-primary" onclick="executeRefund(<%= refundRs("RefundID") %>, '<%= SafeOutput(refundRs("RefundNo")) %>')">
                            <i class="fas fa-play"></i> 执行退款
                        </button>
                        <% Else %>
                        <span style="color:#888;font-size:12px;">无操作</span>
                        <% End If %>
                    </td>
                    <% End If %>
                </tr>
                <% 
                        refundRs.MoveNext
                    Loop
                    refundRs.Close
                End If
                Set refundRs = Nothing
                %>
            </tbody>
        </table>
        
        <% If refundCount = 0 Then %>
        <div class="empty-state">
            <i class="fas fa-inbox"></i>
            <p>暂无退款记录</p>
        </div>
        <% End If %>
        
        <!-- 分页 -->
        <% If refundTotalPages > 1 Then %>
        <div class="pagination">
            <% If refundPageNum > 1 Then %>
            <a href="?tab=refunds&refundPage=<%= refundPageNum-1 %>&refundStatus=<%= Server.HTMLEncode(refundStatus) %>"><i class="fas fa-chevron-left"></i></a>
            <% End If %>
            
            <% 
            Dim refundPageStart, refundPageEnd, rp
            refundPageStart = refundPageNum - 2
            refundPageEnd = refundPageNum + 2
            If refundPageStart < 1 Then refundPageStart = 1
            If refundPageEnd > refundTotalPages Then refundPageEnd = refundTotalPages
            
            For rp = refundPageStart To refundPageEnd
            %>
            <% If rp = refundPageNum Then %>
            <span class="current"><%= rp %></span>
            <% Else %>
            <a href="?tab=refunds&refundPage=<%= rp %>&refundStatus=<%= Server.HTMLEncode(refundStatus) %>"><%= rp %></a>
            <% End If %>
            <% Next %>
            
            <% If refundPageNum < refundTotalPages Then %>
            <a href="?tab=refunds&refundPage=<%= refundPageNum+1 %>&refundStatus=<%= Server.HTMLEncode(refundStatus) %>"><i class="fas fa-chevron-right"></i></a>
            <% End If %>
        </div>
        <% End If %>
        
        <% End If ' End Tab 2 %>
        
        <% If currentTab = "anomalies" Then %>
        <!-- ======================================== -->
        <!-- Tab 3: 异常监控 -->
        <!-- ======================================== -->
        
        <!-- 异常统计 -->
        <div class="summary-bar" style="grid-template-columns: repeat(3, 1fr); margin-bottom: 25px;">
            <div class="summary-card" style="border-left: 4px solid #f44336;">
                <div class="label"><i class="fas fa-money-bill-wave"></i> 大额交易 (>¥<%= largeAmountThreshold %>)</div>
                <div class="value" style="color: #f44336;">
                    <%= GetScalar("SELECT COUNT(*) FROM PaymentRecords WHERE ABS(Amount) > " & largeAmountThreshold) %>
                </div>
            </div>
            <div class="summary-card" style="border-left: 4px solid #ff9800;">
                <div class="label"><i class="fas fa-moon"></i> 非工作时间交易</div>
                <div class="value" style="color: #ff9800;">
                    <%= GetScalar("SELECT COUNT(*) FROM PaymentRecords WHERE Hour(CreatedAt) >= 22 OR Hour(CreatedAt) < 8") %>
                </div>
            </div>
            <div class="summary-card" style="border-left: 4px solid #9c27b0;">
                <div class="label"><i class="fas fa-clone"></i> 重复交易嫌疑</div>
                <div class="value" style="color: #9c27b0;">
                    <% 
                    ' 检测同一分钟内同一订单的多笔交易
                    Dim dupCount
                    dupCount = GetScalar("SELECT COUNT(*) FROM (SELECT OrderID, Format(CreatedAt, 'yyyy-mm-dd hh:nn') as TimeSlot FROM PaymentRecords GROUP BY OrderID, Format(CreatedAt, 'yyyy-mm-dd hh:nn') HAVING COUNT(*) > 1) AS Dups")
                    Response.Write dupCount
                    %>
                </div>
            </div>
        </div>
        
        <h3 style="color: #e0e0e0; margin-bottom: 20px;"><i class="fas fa-exclamation-circle" style="color: #ffc107;"></i> 异常交易列表</h3>
        
        <% 
        ' 获取异常交易记录
        ' 1. 大额交易
        Dim largeTransSQL
        largeTransSQL = "SELECT *, '大额交易' as AnomalyType, '金额超过阈值 ¥" & largeAmountThreshold & "' as AnomalyReason FROM PaymentRecords WHERE ABS(Amount) > " & largeAmountThreshold
        
        ' 2. 非工作时间交易
        Dim offHourSQL
        offHourSQL = "SELECT *, '非工作时间' as AnomalyType, '交易时间在 22:00-08:00 之间' as AnomalyReason FROM PaymentRecords WHERE Hour(CreatedAt) >= 22 OR Hour(CreatedAt) < 8"
        
        ' 3. 重复交易
        Dim dupSQL
        dupSQL = "SELECT p.*, '重复交易' as AnomalyType, '同一分钟内存在多笔交易' as AnomalyReason FROM PaymentRecords p INNER JOIN (SELECT OrderID, Format(CreatedAt, 'yyyy-mm-dd hh:nn') as TimeSlot FROM PaymentRecords GROUP BY OrderID, Format(CreatedAt, 'yyyy-mm-dd hh:nn') HAVING COUNT(*) > 1) AS dups ON p.OrderID = dups.OrderID AND Format(p.CreatedAt, 'yyyy-mm-dd hh:nn') = dups.TimeSlot"
        
        ' 合并查询
        Dim anomalySQL
        anomalySQL = largeTransSQL & " UNION ALL " & offHourSQL & " UNION ALL " & dupSQL & " ORDER BY CreatedAt DESC"
        
        Dim anomalyRs
        Set anomalyRs = ExecuteQuery(anomalySQL)
        
        Dim anomalyCount, isResolved
        anomalyCount = 0
        %>
        
        <% 
        If Not anomalyRs Is Nothing Then
            Do While Not anomalyRs.EOF
                anomalyCount = anomalyCount + 1
                isResolved = False
                If Session("ResolvedAnomaly_" & anomalyRs("RecordID")) = True Then
                    isResolved = True
                End If
        %>
        <div class="anomaly-card <%= IIF(isResolved, "resolved", "") %>">
            <div class="anomaly-header">
                <div>
                    <span class="anomaly-type"><i class="fas fa-exclamation-triangle"></i> <%= Server.HTMLEncode(anomalyRs("AnomalyType")) %></span>
                    <span style="margin-left: 15px; color: #e0e0e0;"><%= Server.HTMLEncode(anomalyRs("OrderNo")) %></span>
                </div>
                <span class="anomaly-time"><%= SafeFormatDateTime(anomalyRs("CreatedAt"), 2) & " " & SafeFormatDateTime(anomalyRs("CreatedAt"), 4) %></span>
            </div>
            <div class="anomaly-body">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <div>
                        <p style="margin: 0 0 10px 0; color: #888;"><%= Server.HTMLEncode(anomalyRs("AnomalyReason")) %></p>
                        <p style="margin: 0;">
                            流水号: <code><%= Server.HTMLEncode(anomalyRs("TransactionNo")) %></code> | 
                            支付方式: <%= Server.HTMLEncode(anomalyRs("PaymentMethod")) %> |
                            类型: <%= Server.HTMLEncode(anomalyRs("TransactionType")) %>
                        </p>
                    </div>
                    <div class="anomaly-amount">
                        ¥<%= FormatNumber(CDbl("0" & anomalyRs("Amount")), 2) %>
                    </div>
                </div>
            </div>
            <% If canEdit AND Not isResolved Then %>
            <div style="margin-top: 15px; padding-top: 15px; border-top: 1px solid #3a3a4a;">
                <form method="post" style="display: inline;">
                    <%= GetCSRFTokenField() %>
                    <input type="hidden" name="action" value="resolve_anomaly">
                    <input type="hidden" name="anomalyId" value="<%= anomalyRs("RecordID") %>">
                    <button type="submit" class="btn btn-sm btn-success">
                        <i class="fas fa-check"></i> 标记为已处理
                    </button>
                </form>
            </div>
            <% End If %>
        </div>
        <% 
                anomalyRs.MoveNext
            Loop
            anomalyRs.Close
        End If
        Set anomalyRs = Nothing
        %>
        
        <% If anomalyCount = 0 Then %>
        <div class="empty-state">
            <i class="fas fa-check-circle" style="color: #4CAF50;"></i>
            <p>未发现异常交易，系统运行正常</p>
        </div>
        <% End If %>
        
        <% End If ' End Tab 3 %>
        
    </div>
    
    <% If canEdit Then %>
    <!-- ======================================== -->
    <!-- 模态框：修改分类 -->
    <!-- ======================================== -->
    <div id="categoryModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-tag"></i> 修改分类标签</h3>
                <button class="modal-close" onclick="closeCategoryModal()">&times;</button>
            </div>
            <form method="post" action="">
                <%= GetCSRFTokenField() %>
                <input type="hidden" name="action" value="update_category">
                <input type="hidden" name="recordId" id="categoryRecordId" value="">
                <div class="modal-body">
                    <div class="form-group">
                        <label>选择分类</label>
                        <select name="category" id="categorySelect" style="width: 100%;">
                            <option value="订单收入">订单收入</option>
                            <option value="退款支出">退款支出</option>
                            <option value="平台扣费">平台扣费</option>
                            <option value="营销费用">营销费用</option>
                            <option value="资金划转">资金划转</option>
                            <option value="其他收入">其他收入</option>
                            <option value="其他支出">其他支出</option>
                        </select>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" onclick="closeCategoryModal()">取消</button>
                    <button type="submit" class="btn btn-primary">保存</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- ======================================== -->
    <!-- 模态框：新建退款申请 -->
    <!-- ======================================== -->
    <div id="refundModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-undo"></i> 新建退款申请</h3>
                <button class="modal-close" onclick="closeRefundModal()">&times;</button>
            </div>
            <form method="post" action="" onsubmit="return validateRefundForm()">
                <%= GetCSRFTokenField() %>
                <input type="hidden" name="action" value="create_refund">
                <div class="modal-body">
                    <div class="form-group">
                        <label>选择订单 <span style="color: #f44336;">*</span></label>
                        <select name="orderId" id="refundOrderId" required onchange="loadOrderInfo(this.value)" style="width: 100%;">
                            <option value="">请选择订单</option>
                            <% 
                            Dim orderListRs
                            Set orderListRs = ExecuteQuery("SELECT OrderID, OrderNo, TotalAmount, RefundAmount FROM Orders WHERE Status = 'Paid' AND (TotalAmount - IIF(RefundAmount IS NULL, 0, RefundAmount)) > 0 ORDER BY OrderID DESC")
                            If Not orderListRs Is Nothing Then
                                Dim remainingAmount
                            Do While Not orderListRs.EOF
                                    remainingAmount = CDbl("0" & orderListRs("TotalAmount")) - CDbl("0" & orderListRs("RefundAmount"))
                            %>
                            <option value="<%= orderListRs("OrderID") %>" data-amount="<%= remainingAmount %>">
                                <%= Server.HTMLEncode(orderListRs("OrderNo")) %> - ¥<%= FormatNumber(remainingAmount, 2) %> 可退
                            </option>
                            <% 
                                    orderListRs.MoveNext
                                Loop
                                orderListRs.Close
                            End If
                            Set orderListRs = Nothing
                            %>
                        </select>
                        <div class="help-text">只显示已支付且有余额的订单</div>
                    </div>
                    <div class="form-group">
                        <label>退款金额 <span style="color: #f44336;">*</span></label>
                        <input type="number" name="refundAmount" id="refundAmount" step="0.01" min="0.01" required style="width: 100%;">
                        <div class="help-text" id="maxRefundHint">请先选择订单</div>
                    </div>
                    <div class="form-group">
                        <label>退款原因 <span style="color: #f44336;">*</span></label>
                        <textarea name="refundReason" id="refundReason" required placeholder="请详细说明退款原因..."></textarea>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" onclick="closeRefundModal()">取消</button>
                    <button type="submit" class="btn btn-primary">提交申请</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- ======================================== -->
    <!-- 模态框：审批退款 -->
    <!-- ======================================== -->
    <div id="approveModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3 id="approveModalTitle"><i class="fas fa-gavel"></i> 审批退款</h3>
                <button class="modal-close" onclick="closeApproveModal()">&times;</button>
            </div>
            <form method="post" action="">
                <%= GetCSRFTokenField() %>
                <input type="hidden" name="action" value="approve_refund">
                <input type="hidden" name="refundId" id="approveRefundId" value="">
                <input type="hidden" name="approveAction" id="approveAction" value="">
                <div class="modal-body">
                    <div class="form-group" id="rejectReasonGroup" style="display: none;">
                        <label>拒绝原因 <span style="color: #f44336;">*</span></label>
                        <textarea name="rejectReason" id="rejectReason" placeholder="请说明拒绝原因..."></textarea>
                    </div>
                    <div id="approveConfirmText" style="color: #e0e0e0; padding: 20px; background: #1a1a2e; border-radius: 8px;">
                        <i class="fas fa-question-circle" style="color: #00bcd4;"></i> 确定要通过此退款申请吗？
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" onclick="closeApproveModal()">取消</button>
                    <button type="submit" class="btn btn-primary" id="approveSubmitBtn">确认</button>
                </div>
            </form>
        </div>
    </div>
    <% End If %>
    
    <!-- ======================================== -->
    <!-- JavaScript -->
    <!-- ======================================== -->
    <script>
        // 分类模态框
        function openCategoryModal(recordId, currentCategory) {
            document.getElementById('categoryRecordId').value = recordId;
            document.getElementById('categorySelect').value = currentCategory || '订单收入';
            document.getElementById('categoryModal').classList.add('active');
        }
        
        function closeCategoryModal() {
            document.getElementById('categoryModal').classList.remove('active');
        }
        
        // 退款申请模态框
        function openRefundModal() {
            document.getElementById('refundModal').classList.add('active');
        }
        
        function closeRefundModal() {
            document.getElementById('refundModal').classList.remove('active');
            document.getElementById('refundOrderId').value = '';
            document.getElementById('refundAmount').value = '';
            document.getElementById('refundReason').value = '';
            document.getElementById('maxRefundHint').textContent = '请先选择订单';
        }
        
        // 加载订单信息
        function loadOrderInfo(orderId) {
            var select = document.getElementById('refundOrderId');
            var option = select.options[select.selectedIndex];
            var maxAmount = option.getAttribute('data-amount');
            
            if (maxAmount) {
                document.getElementById('refundAmount').max = maxAmount;
                document.getElementById('maxRefundHint').textContent = '最大可退金额: ¥' + parseFloat(maxAmount).toFixed(2);
            } else {
                document.getElementById('maxRefundHint').textContent = '请先选择订单';
            }
        }
        
        // 验证退款表单
        function validateRefundForm() {
            var orderId = document.getElementById('refundOrderId').value;
            var amount = parseFloat(document.getElementById('refundAmount').value);
            var maxAmount = parseFloat(document.getElementById('refundAmount').max);
            var reason = document.getElementById('refundReason').value.trim();
            
            if (!orderId) {
                alert('请选择订单');
                return false;
            }
            
            if (!amount || amount <= 0) {
                alert('请输入有效的退款金额');
                return false;
            }
            
            if (maxAmount && amount > maxAmount) {
                alert('退款金额不能超过最大可退金额 ¥' + maxAmount.toFixed(2));
                return false;
            }
            
            if (!reason) {
                alert('请填写退款原因');
                return false;
            }
            
            return true;
        }
        
        // 审批模态框
        function openApproveModal(refundId, action) {
            document.getElementById('approveRefundId').value = refundId;
            document.getElementById('approveAction').value = action;
            
            var title = document.getElementById('approveModalTitle');
            var confirmText = document.getElementById('approveConfirmText');
            var rejectGroup = document.getElementById('rejectReasonGroup');
            var submitBtn = document.getElementById('approveSubmitBtn');
            
            if (action === 'approve') {
                title.innerHTML = '<i class="fas fa-check-circle"></i> 通过退款申请';
                confirmText.innerHTML = '<i class="fas fa-question-circle" style="color: #4CAF50;"></i> 确定要通过此退款申请吗？通过后需要执行退款操作。';
                confirmText.style.display = 'block';
                rejectGroup.style.display = 'none';
                submitBtn.className = 'btn btn-success';
                submitBtn.innerHTML = '<i class="fas fa-check"></i> 确认通过';
                document.getElementById('rejectReason').required = false;
            } else {
                title.innerHTML = '<i class="fas fa-times-circle"></i> 拒绝退款申请';
                confirmText.style.display = 'none';
                rejectGroup.style.display = 'block';
                submitBtn.className = 'btn btn-danger';
                submitBtn.innerHTML = '<i class="fas fa-times"></i> 确认拒绝';
                document.getElementById('rejectReason').required = true;
            }
            
            document.getElementById('approveModal').classList.add('active');
        }
        
        function closeApproveModal() {
            document.getElementById('approveModal').classList.remove('active');
            document.getElementById('rejectReason').value = '';
        }
        
        // 执行退款确认
        function executeRefund(refundId, refundNo) {
            if (confirm('确定要执行退款 [' + refundNo + '] 吗？\n\n执行后将：\n1. 更新订单退款金额\n2. 在支付记录中插入负向记录\n3. 标记退款已完成\n\n此操作不可撤销！')) {
                var form = document.createElement('form');
                form.method = 'post';
                form.innerHTML = '<%= GetCSRFTokenField() %><input type="hidden" name="action" value="execute_refund"><input type="hidden" name="refundId" value="' + refundId + '">';
                document.body.appendChild(form);
                form.submit();
            }
        }
        
        // 点击模态框外部关闭
        document.querySelectorAll('.modal').forEach(function(modal) {
            modal.addEventListener('click', function(e) {
                if (e.target === this) {
                    this.classList.remove('active');
                }
            });
        });
        
        // ESC键关闭模态框
        document.addEventListener('keydown', function(e) {
            if (e.key === 'Escape') {
                document.querySelectorAll('.modal.active').forEach(function(modal) {
                    modal.classList.remove('active');
                });
            }
        });
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
