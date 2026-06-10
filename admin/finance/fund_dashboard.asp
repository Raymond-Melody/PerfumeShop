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

' 安全除法函数
Function SafeDiv(numerator, denominator)
    Dim n, d
    n = SafeNum(numerator)
    d = SafeNum(denominator)
    If d = 0 Then
        SafeDiv = 0
    Else
        SafeDiv = n / d
    End If
End Function

' 安全格式化函数
Function SafeFormat(val, decimals)
    SafeFormat = FormatNumber(SafeNum(val), decimals)
End Function

' 安全百分比函数
Function SafePercent(numerator, denominator)
    If SafeNum(denominator) = 0 Then
        SafePercent = "0.00%"
    Else
        SafePercent = FormatNumber(SafeNum(numerator) / SafeNum(denominator) * 100, 2) & "%"
    End If
End Function

Call OpenConnection()

' ============================================
' 权限检查
' ============================================
Dim isFinManager, isFinStaff
isFinManager = (Session("AdminRoleCode") = "FIN_MANAGER" OR Session("AdminRoleCode") = "SUPER_ADMIN")
isFinStaff = (Session("AdminRoleCode") = "FIN_STAFF")

' 如果不是财务经理也不是财务专员，拒绝访问
If NOT isFinManager AND NOT isFinStaff Then
    Response.Redirect "index.asp?error=权限不足"
    Response.End
End If

' ============================================
' 处理AJAX请求
' ============================================
Dim ajaxAction
ajaxAction = Request.Form("ajax_action")

If ajaxAction <> "" Then
    Response.ContentType = "application/json"
    
    If NOT ValidateCSRFToken() Then
        Response.Write "{"
        Response.Write chr(34) & "success" & chr(34) & ":false," & chr(34) & "message" & chr(34) & ":" & chr(34) & "CSRF validation failed" & chr(34) & "}"
        Response.End
    End If
    
    If NOT isFinManager Then
        Response.Write "{"
        Response.Write chr(34) & "success" & chr(34) & ":false," & chr(34) & "message" & chr(34) & ":" & chr(34) & "Permission denied" & chr(34) & "}"
        Response.End
    End If
    
    Select Case ajaxAction
        Case "add_account"
            Call AddAccount()
        Case "edit_account"
            Call EditAccount()
        Case "update_balance"
            Call UpdateBalance()
        Case "set_threshold"
            Call SetThreshold()
        Case "toggle_status"
            Call ToggleStatus()
        Case "get_pending_detail"
            Call GetPendingDetail()
    End Select
    
    Response.End
End If

' ============================================
' AJAX处理函数
' ============================================

' 添加账户
Sub AddAccount()
    Dim accountName, accountType, initialBalance, alertThreshold
    accountName = Trim(Request.Form("accountName"))
    accountType = Trim(Request.Form("accountType"))
    initialBalance = SafeNum(Request.Form("initialBalance"))
    alertThreshold = SafeNum(Request.Form("alertThreshold"))
    
    If accountName = "" Then
        Response.Write "{" & chr(34) & "success" & chr(34) & ":false," & chr(34) & "message" & chr(34) & ":" & chr(34) & "Account name cannot be empty" & chr(34) & "}"
        Exit Sub
    End If
    
    Dim sql
    sql = "INSERT INTO FundAccounts (AccountName, AccountType, TotalBalance, AvailableBalance, FrozenAmount, PendingSettlement, AlertThreshold, IsActive, CreatedAt, UpdatedAt) VALUES ('" & _
          SafeSQL(accountName) & "', '" & SafeSQL(accountType) & "', " & initialBalance & ", " & initialBalance & ", 0, 0, " & alertThreshold & ", -1, GETDATE(), GETDATE())"
    
    If ExecuteNonQuery(sql) Then
        Call LogAdminAction("添加资金账户", "finance", "FundAccounts", "", accountName)
        Response.Write "{" & chr(34) & "success" & chr(34) & ":true," & chr(34) & "message" & chr(34) & ":" & chr(34) & "Added successfully" & chr(34) & "}"
    Else
        Response.Write "{" & chr(34) & "success" & chr(34) & ":false," & chr(34) & "message" & chr(34) & ":" & chr(34) & "Add failed" & chr(34) & "}"
    End If
End Sub

' 编辑账户
Sub EditAccount()
    Dim accountId, accountName, accountType
    accountId = CLng("0" & Request.Form("accountId"))
    accountName = Trim(Request.Form("accountName"))
    accountType = Trim(Request.Form("accountType"))
    
    If accountId = 0 OR accountName = "" Then
        Response.Write "{" & chr(34) & "success" & chr(34) & ":false," & chr(34) & "message" & chr(34) & ":" & chr(34) & "Invalid parameters" & chr(34) & "}"
        Exit Sub
    End If
    
    Dim sql
    sql = "UPDATE FundAccounts SET AccountName = '" & SafeSQL(accountName) & "', AccountType = '" & _
          SafeSQL(accountType) & "', UpdatedAt = GETDATE() WHERE AccountID = " & accountId
    
    If ExecuteNonQuery(sql) Then
        Call LogAdminAction("编辑资金账户", "finance", "FundAccounts", CStr(accountId), accountName)
        Response.Write "{" & chr(34) & "success" & chr(34) & ":true," & chr(34) & "message" & chr(34) & ":" & chr(34) & "Updated successfully" & chr(34) & "}"
    Else
        Response.Write "{" & chr(34) & "success" & chr(34) & ":false," & chr(34) & "message" & chr(34) & ":" & chr(34) & "Update failed" & chr(34) & "}"
    End If
End Sub

' 更新余额
Sub UpdateBalance()
    Dim accountId, newBalance
    accountId = CLng("0" & Request.Form("accountId"))
    newBalance = SafeNum(Request.Form("newBalance"))
    
    If accountId = 0 Then
        Response.Write "{" & chr(34) & "success" & chr(34) & ":false," & chr(34) & "message" & chr(34) & ":" & chr(34) & "Invalid parameters" & chr(34) & "}"
        Exit Sub
    End If
    
    Dim sql, rs, frozenAmount
    ' 先获取当前冻结金额
    Set rs = ExecuteQuery("SELECT FrozenAmount FROM FundAccounts WHERE AccountID = " & accountId)
    If rs Is Nothing OR rs.EOF Then
        Response.Write "{" & chr(34) & "success" & chr(34) & ":false," & chr(34) & "message" & chr(34) & ":" & chr(34) & "Account does not exist" & chr(34) & "}"
        Exit Sub
    End If
    frozenAmount = SafeNum(rs("FrozenAmount"))
    rs.Close
    Set rs = Nothing
    
    ' 更新余额（可用余额 = 新余额 - 冻结金额）
    Dim availableBalance
    availableBalance = newBalance - frozenAmount
    If availableBalance < 0 Then availableBalance = 0
    
    sql = "UPDATE FundAccounts SET TotalBalance = " & newBalance & ", AvailableBalance = " & _
          availableBalance & ", LastSyncAt = GETDATE(), UpdatedAt = GETDATE() WHERE AccountID = " & accountId
    
    If ExecuteNonQuery(sql) Then
        Call LogAdminAction("更新账户余额", "finance", "FundAccounts", CStr(accountId), "余额更新为" & newBalance)
        Response.Write "{" & chr(34) & "success" & chr(34) & ":true," & chr(34) & "message" & chr(34) & ":" & chr(34) & "Balance updated successfully" & chr(34) & "}"
    Else
        Response.Write "{" & chr(34) & "success" & chr(34) & ":false," & chr(34) & "message" & chr(34) & ":" & chr(34) & "Update failed" & chr(34) & "}"
    End If
End Sub

' 设置预警阈值
Sub SetThreshold()
    Dim accountId, threshold
    accountId = CLng("0" & Request.Form("accountId"))
    threshold = SafeNum(Request.Form("threshold"))
    
    If accountId = 0 Then
        Response.Write "{" & chr(34) & "success" & chr(34) & ":false," & chr(34) & "message" & chr(34) & ":" & chr(34) & "Invalid parameters" & chr(34) & "}"
        Exit Sub
    End If
    
    Dim sql
    sql = "UPDATE FundAccounts SET AlertThreshold = " & threshold & ", UpdatedAt = GETDATE() WHERE AccountID = " & accountId
    
    If ExecuteNonQuery(sql) Then
        Call LogAdminAction("Set threshold", "finance", "FundAccounts", CStr(accountId), "Threshold: " & threshold)
        Response.Write "{" & chr(34) & "success" & chr(34) & ":true," & chr(34) & "message" & chr(34) & ":" & chr(34) & "Threshold set successfully" & chr(34) & "}"
    Else
        Response.Write "{" & chr(34) & "success" & chr(34) & ":false," & chr(34) & "message" & chr(34) & ":" & chr(34) & "Setting failed" & chr(34) & "}"
    End If
End Sub

' 停用/启用账户
Sub ToggleStatus()
    Dim accountId, isActive
    accountId = CLng("0" & Request.Form("accountId"))
    isActive = (Request.Form("isActive") = "1")
    
    If accountId = 0 Then
        Response.Write "{" & chr(34) & "success" & chr(34) & ":false," & chr(34) & "message" & chr(34) & ":" & chr(34) & "Invalid parameters" & chr(34) & "}"
        Exit Sub
    End If
    
    Dim sql, newStatus
    newStatus = IIF(isActive, "-1", "0")
    sql = "UPDATE FundAccounts SET IsActive = " & newStatus & ", UpdatedAt = GETDATE() WHERE AccountID = " & accountId
    
    If ExecuteNonQuery(sql) Then
        Dim actionDesc
        actionDesc = IIF(isActive, "启用", "停用")
        Call LogAdminAction(actionDesc & " account", "finance", "FundAccounts", CStr(accountId), "")
        Response.Write "{" & chr(34) & "success" & chr(34) & ":true," & chr(34) & "message" & chr(34) & ":" & chr(34) & actionDesc & " successful" & chr(34) & "}"
    Else
        Response.Write "{" & chr(34) & "success" & chr(34) & ":false," & chr(34) & "message" & chr(34) & ":" & chr(34) & "Operation failed" & chr(34) & "}"
    End If
End Sub

' 获取待结算明细
Sub GetPendingDetail()
    Dim detailType
    detailType = Request.Form("detailType")
    
    Dim sql, rs, result
    ' 循环变量声明移到循环外（VBScript限制）
    Dim firstOrder, firstWithdraw, firstFrozen, firstPending
    result = "{" & chr(34) & "success" & chr(34) & ":true," & chr(34) & "data" & chr(34) & ":["
    
    Select Case detailType
        Case "orders"
            ' 交易在途 - Orders.Status='Paid' 未结算
            sql = "SELECT OrderID, OrderNo, TotalAmount, PaymentMethod, CreatedAt FROM Orders WHERE Status = 'Paid' ORDER BY CreatedAt DESC"
            Set rs = ExecuteQuery(sql)
            If Not rs Is Nothing Then
                firstOrder = True
                Do While Not rs.EOF
                    If Not firstOrder Then result = result & ","
                    result = result & "{" & chr(34) & "id" & chr(34) & ":" & rs("OrderID") & "," & chr(34) & "no" & chr(34) & ":" & chr(34) & rs("OrderNo") & chr(34) & "," & chr(34) & "amount" & chr(34) & ":" & SafeNum(rs("TotalAmount")) & "," & chr(34) & "method" & chr(34) & ":" & chr(34) & rs("PaymentMethod") & chr(34) & "," & chr(34) & "time" & chr(34) & ":" & chr(34) & FormatDateField(rs("CreatedAt")) & chr(34) & "}"
                    firstOrder = False
                    rs.MoveNext
                Loop
                rs.Close
            End If
            Set rs = Nothing
            
        Case "withdraw"
            ' 提现中 - PaymentRecords WHERE Status='Pending' AND TransactionType='Transfer'
            sql = "SELECT RecordID, OrderNo, Amount, CreatedAt FROM PaymentRecords WHERE Status = 'Pending' AND TransactionType = 'Transfer' ORDER BY CreatedAt DESC"
            Set rs = ExecuteQuery(sql)
            If Not rs Is Nothing Then
                firstWithdraw = True
                Do While Not rs.EOF
                    If Not firstWithdraw Then result = result & ","
                    result = result & "{" & chr(34) & "id" & chr(34) & ":" & rs("RecordID") & "," & chr(34) & "no" & chr(34) & ":" & chr(34) & rs("OrderNo") & chr(34) & "," & chr(34) & "amount" & chr(34) & ":" & SafeNum(rs("Amount")) & "," & chr(34) & "time" & chr(34) & ":" & chr(34) & FormatDateField(rs("CreatedAt")) & chr(34) & "}"
                    firstWithdraw = False
                    rs.MoveNext
                Loop
                rs.Close
            End If
            Set rs = Nothing
            
        Case "frozen"
            ' 不可用保证金 - FundAccounts.FrozenAmount > 0
            sql = "SELECT AccountID, AccountName, FrozenAmount FROM FundAccounts WHERE FrozenAmount > 0 ORDER BY FrozenAmount DESC"
            Set rs = ExecuteQuery(sql)
            If Not rs Is Nothing Then
                firstFrozen = True
                Do While Not rs.EOF
                    If Not firstFrozen Then result = result & ","
                    result = result & "{" & chr(34) & "id" & chr(34) & ":" & rs("AccountID") & "," & chr(34) & "name" & chr(34) & ":" & chr(34) & rs("AccountName") & chr(34) & "," & chr(34) & "amount" & chr(34) & ":" & SafeNum(rs("FrozenAmount")) & "}"
                    firstFrozen = False
                    rs.MoveNext
                Loop
                rs.Close
            End If
            Set rs = Nothing
    End Select
    
    result = result & "]}"
    Response.Write result
End Sub

' ============================================
' 获取统计数据
' ============================================

' 账户总览统计
Dim totalBookBalance, totalAvailableBalance, totalFrozenAmount, totalPendingSettlement
Dim rsTemp

Set rsTemp = conn.Execute("SELECT SUM(TotalBalance) FROM FundAccounts WHERE IsActive = 1")
totalBookBalance = 0
If Not rsTemp.EOF Then If Not IsNull(rsTemp.Fields(0).Value) Then totalBookBalance = SafeNum(rsTemp.Fields(0).Value)
rsTemp.Close

Set rsTemp = conn.Execute("SELECT SUM(AvailableBalance) FROM FundAccounts WHERE IsActive = 1")
totalAvailableBalance = 0
If Not rsTemp.EOF Then If Not IsNull(rsTemp.Fields(0).Value) Then totalAvailableBalance = SafeNum(rsTemp.Fields(0).Value)
rsTemp.Close

Set rsTemp = conn.Execute("SELECT SUM(FrozenAmount) FROM FundAccounts WHERE IsActive = 1")
totalFrozenAmount = 0
If Not rsTemp.EOF Then If Not IsNull(rsTemp.Fields(0).Value) Then totalFrozenAmount = SafeNum(rsTemp.Fields(0).Value)
rsTemp.Close

Set rsTemp = conn.Execute("SELECT SUM(PendingSettlement) FROM FundAccounts WHERE IsActive = 1")
totalPendingSettlement = 0
If Not rsTemp.EOF Then If Not IsNull(rsTemp.Fields(0).Value) Then totalPendingSettlement = SafeNum(rsTemp.Fields(0).Value)
rsTemp.Close

' 待结算资金穿透统计
Dim pendingOrdersAmount, pendingWithdrawAmount

Set rsTemp = conn.Execute("SELECT SUM(CAST(TotalAmount AS FLOAT)) FROM Orders WHERE Status = 'Paid'")
pendingOrdersAmount = 0
If Not rsTemp.EOF Then If Not IsNull(rsTemp.Fields(0).Value) Then pendingOrdersAmount = SafeNum(rsTemp.Fields(0).Value)
rsTemp.Close

Set rsTemp = conn.Execute("SELECT SUM(Amount) FROM PaymentRecords WHERE Status = 'Pending' AND TransactionType = 'Transfer'")
pendingWithdrawAmount = 0
If Not rsTemp.EOF Then If Not IsNull(rsTemp.Fields(0).Value) Then pendingWithdrawAmount = SafeNum(rsTemp.Fields(0).Value)
rsTemp.Close

Set rsTemp = Nothing

' 大额转账检测（同一分钟内超过10000元）
Dim largeTransferCount
largeTransferCount = CLng("0" & GetScalar("SELECT COUNT(*) FROM (SELECT DATEPART('n', CreatedAt) AS MinutePart, SUM(Amount) AS TotalAmt FROM PaymentRecords WHERE TransactionType = 'Transfer' AND Amount >= 10000 GROUP BY DATEPART('n', CreatedAt) HAVING SUM(Amount) >= 10000) AS T"))

Call LogAdminAction("查看资金看板", "finance", "", "", "")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>资金看板 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        /* 鏆楄壊涓婚鍩虹 */
        body { background: #1a1a2e; color: #e0e0e0; }
        
        /* 顶部统计卡片 */
        .stats-overview {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 20px;
            margin-bottom: 30px;
        }
        
        .stat-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 16px;
            padding: 25px;
            box-shadow: 0 8px 32px rgba(0,0,0,0.3);
            border: 1px solid rgba(255,255,255,0.06);
            position: relative;
            overflow: hidden;
        }
        
        .stat-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 4px;
        }
        
        .stat-card.total::before { background: linear-gradient(90deg, #00bcd4, #00838f); }
        .stat-card.available::before { background: linear-gradient(90deg, #11998e, #38ef7d); }
        .stat-card.frozen::before { background: linear-gradient(90deg, #fc4a1a, #f7b733); }
        .stat-card.pending::before { background: linear-gradient(90deg, #4facfe, #00f2fe); }
        
        .stat-card.highlight-warning {
            border-color: #ff4757;
            animation: pulse-warning 2s infinite;
        }
        
        @keyframes pulse-warning {
            0%, 100% { box-shadow: 0 0 0 0 rgba(255, 71, 87, 0.4); }
            50% { box-shadow: 0 0 20px 5px rgba(255, 71, 87, 0.2); }
        }
        
        .stat-label {
            font-size: 14px;
            color: #888;
            margin-bottom: 10px;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .stat-value {
            font-size: 32px;
            font-weight: 700;
            color: #fff;
            margin-bottom: 5px;
        }
        
        .stat-sub {
            font-size: 12px;
            color: #888;
        }
        
        .stat-card.available .stat-value { color: #38ef7d; }
        .stat-card.frozen .stat-value { color: #f7b733; }
        .stat-card.pending .stat-value { color: #00f2fe; }
        
        /* 区块标题 */
        .section-title {
            font-size: 18px;
            font-weight: 600;
            color: #e0e0e0;
            margin: 30px 0 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid rgba(255,255,255,0.06);
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        /* 账户卡片网格 */
        .accounts-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
            margin-bottom: 30px;
        }
        
        .account-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 20px;
            border: 1px solid rgba(255,255,255,0.06);
            transition: all 0.3s ease;
            position: relative;
        }
        
        .account-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 10px 30px rgba(0,0,0,0.4);
        }
        
        .account-card.alert {
            border-color: #ff4757;
            animation: pulse-red 2s infinite;
        }
        
        @keyframes pulse-red {
            0%, 100% { box-shadow: 0 0 0 0 rgba(255, 71, 87, 0.4); }
            50% { box-shadow: 0 0 15px 3px rgba(255, 71, 87, 0.3); }
        }
        
        .account-card.inactive {
            opacity: 0.6;
        }
        
        .account-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 15px;
        }
        
        .account-name {
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 16px;
            font-weight: 600;
        }
        
        .account-icon {
            width: 40px;
            height: 40px;
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 20px;
        }
        
        .icon-alipay { background: #1677ff20; color: #1677ff; }
        .icon-wechat { background: #07c16020; color: #07c160; }
        .icon-bank { background: #ff6b6b20; color: #ff6b6b; }
        .icon-paypal { background: #00308720; color: #0070ba; }
        .icon-cash { background: #ffa72620; color: #ffa726; }
        .icon-other { background: #9c27b020; color: #9c27b0; }
        
        .account-status {
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: 500;
        }
        
        .status-normal { background: #1b5e20; color: #81c784; }
        .status-warning { background: #5e1b1b; color: #ff4757; }
        .status-inactive { background: #424242; color: #9e9e9e; }
        
        .account-balance {
            margin: 15px 0;
        }
        
        .balance-main {
            font-size: 24px;
            font-weight: 700;
            color: #fff;
        }
        
        .balance-sub {
            display: flex;
            gap: 20px;
            margin-top: 10px;
            font-size: 12px;
            color: #888;
        }
        
        .balance-sub span {
            display: flex;
            align-items: center;
            gap: 5px;
        }
        
        .balance-sub .frozen { color: #f7b733; }
        .balance-sub .available { color: #38ef7d; }
        
        .account-footer {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-top: 15px;
            padding-top: 15px;
            border-top: 1px solid #3a3a4a;
            font-size: 12px;
            color: #666;
        }
        
        .account-actions {
            display: flex;
            gap: 8px;
        }
        
        .btn-icon {
            width: 28px;
            height: 28px;
            border-radius: 6px;
            border: none;
            background: #3a3a4a;
            color: #888;
            cursor: pointer;
            transition: all 0.2s;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        
        .btn-icon:hover {
            background: #4a4a5a;
            color: #fff;
        }
        
        .btn-icon.edit:hover { background: #00bcd4; }
        .btn-icon.balance:hover { background: #11998e; }
        .btn-icon.threshold:hover { background: #f7b733; color: #333; }
        .btn-icon.toggle:hover { background: #ff4757; }
        .btn-icon.toggle.enable:hover { background: #38ef7d; color: #333; }
        
        /* 待结算资金区块 */
        .pending-section {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
            margin-bottom: 30px;
        }
        
        .pending-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 20px;
            border: 1px solid rgba(255,255,255,0.06);
            cursor: pointer;
            transition: all 0.3s;
        }
        
        .pending-card:hover {
            border-color: #00bcd4;
            transform: translateY(-3px);
        }
        
        .pending-header {
            display: flex;
            align-items: center;
            gap: 10px;
            margin-bottom: 15px;
            color: #888;
            font-size: 14px;
        }
        
        .pending-amount {
            font-size: 28px;
            font-weight: 700;
            color: #fff;
            margin-bottom: 10px;
        }
        
        .pending-desc {
            font-size: 12px;
            color: #888;
        }
        
        .pending-card.orders .pending-amount { color: #4facfe; }
        .pending-card.withdraw .pending-amount { color: #f7b733; }
        .pending-card.frozen .pending-amount { color: #ff6b6b; }
        
        /* 预警区块 */
        .alert-section {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 20px;
            border: 1px solid rgba(255,255,255,0.06);
        }
        
        .alert-table {
            width: 100%;
            border-collapse: collapse;
        }
        
        .alert-table th,
        .alert-table td {
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid rgba(255,255,255,0.06);
        }
        
        .alert-table th {
            color: #888;
            font-weight: 500;
            font-size: 13px;
        }
        
        .alert-table td {
            color: #e0e0e0;
            font-size: 14px;
        }
        
        .alert-table tr:hover {
            background: #323242;
        }
        
        .alert-level {
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: 500;
        }
        
        .level-high { background: #5e1b1b; color: #ff4757; }
        .level-medium { background: #5e4b1b; color: #ffa726; }
        .level-low { background: #1b3e5e; color: #4facfe; }
        
        .diff-negative {
            color: #ff4757;
            font-weight: 600;
        }
        
        /* 异常监控 */
        .abnormal-alert {
            background: linear-gradient(135deg, #5e1b1b 0%, #3e1b1b 100%);
            border: 1px solid #ff4757;
            border-radius: 12px;
            padding: 15px 20px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 15px;
        }
        
        .abnormal-alert i {
            font-size: 24px;
            color: #ff4757;
            animation: blink 1s infinite;
        }
        
        @keyframes blink {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.5; }
        }
        
        .abnormal-content h4 {
            margin: 0 0 5px;
            color: #ff4757;
        }
        
        .abnormal-content p {
            margin: 0;
            color: #ff8a8a;
            font-size: 13px;
        }
        
        /* 操作按钮 */
        .action-bar {
            display: flex;
            justify-content: flex-end;
            margin-bottom: 20px;
        }
        
        .btn-primary {
            background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%);
            color: white;
            padding: 12px 24px;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 8px;
            transition: opacity 0.2s;
        }
        
        .btn-primary:hover {
            opacity: 0.9;
        }
        
        /* 模态框 */
        .modal {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0,0,0,0.7);
            z-index: 1000;
            align-items: center;
            justify-content: center;
        }
        
        .modal.active {
            display: flex;
        }
        
        .modal-content {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 16px;
            width: 90%;
            max-width: 500px;
            border: 1px solid rgba(255,255,255,0.06);
            overflow: hidden;
        }
        
        .modal-header {
            padding: 20px;
            border-bottom: 1px solid rgba(255,255,255,0.06);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .modal-header h3 {
            margin: 0;
            color: #e0e0e0;
        }
        
        .modal-close {
            background: none;
            border: none;
            color: #888;
            font-size: 20px;
            cursor: pointer;
        }
        
        .modal-close:hover {
            color: #fff;
        }
        
        .modal-body {
            padding: 20px;
        }
        
        .form-group {
            margin-bottom: 20px;
        }
        
        .form-group label {
            display: block;
            margin-bottom: 8px;
            color: #b0b0b0;
            font-size: 14px;
        }
        
        .form-group input,
        .form-group select {
            width: 100%;
            padding: 12px 15px;
            border: 2px solid #3a3a4a;
            border-radius: 8px;
            background: #1a1a2e;
            color: #e0e0e0;
            font-size: 14px;
        }
        
        .form-group input:focus,
        .form-group select:focus {
            border-color: #00bcd4;
            outline: none;
        }
        
        .modal-footer {
            padding: 15px 20px;
            border-top: 1px solid rgba(255,255,255,0.06);
            display: flex;
            justify-content: flex-end;
            gap: 10px;
        }
        
        .btn-secondary {
            background: #3a3a4a;
            color: #e0e0e0;
            padding: 10px 20px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
        }
        
        .btn-secondary:hover {
            background: #4a4a5a;
        }
        
        /* 明细弹窗 */
        .detail-list {
            max-height: 400px;
            overflow-y: auto;
        }
        
        .detail-item {
            padding: 12px;
            border-bottom: 1px solid rgba(255,255,255,0.06);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .detail-item:last-child {
            border-bottom: none;
        }
        
        .detail-info h4 {
            margin: 0 0 5px;
            color: #e0e0e0;
            font-size: 14px;
        }
        
        .detail-info p {
            margin: 0;
            color: #888;
            font-size: 12px;
        }
        
        .detail-amount {
            font-size: 16px;
            font-weight: 600;
            color: #fff;
        }
        
        /* 响应式 */
        @media (max-width: 1200px) {
            .stats-overview { grid-template-columns: repeat(2, 1fr); }
            .accounts-grid { grid-template-columns: repeat(2, 1fr); }
            .pending-section { grid-template-columns: 1fr; }
        }
        
        @media (max-width: 768px) {
            .stats-overview { grid-template-columns: 1fr; }
            .accounts-grid { grid-template-columns: 1fr; }
        }
        
        /* 空状态 */
        .empty-state {
            text-align: center;
            padding: 40px;
            color: #888;
        }
        
        .empty-state i {
            font-size: 48px;
            margin-bottom: 15px;
            color: #555;
        }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-wallet"></i> 资金看板</h2>
            <div class="breadcrumb">
                <a href="index.asp">财务中心</a> / <span>资金看板</span>
            </div>
        </div>
        
        <% If Request.QueryString("msg") <> "" Then %>
        <div class="alert alert-success"><i class="fas fa-check-circle"></i> <%= Server.HTMLEncode(Request.QueryString("msg")) %></div>
        <% End If %>
        
        <% If Request.QueryString("error") <> "" Then %>
        <div class="alert alert-error"><i class="fas fa-times-circle"></i> <%= Server.HTMLEncode(Request.QueryString("error")) %></div>
        <% End If %>
        
        <% If isFinManager Then %>
        <div class="action-bar">
            <button class="btn-primary" onclick="openModal('addAccountModal')">
                <i class="fas fa-plus"></i> 添加账户
            </button>
        </div>
        <% End If %>
        
        <!-- 区块1:账户总览卡片 -->
        <div class="stats-overview">
            <div class="stat-card total">
                <div class="stat-label">
                    <i class="fas fa-coins"></i> 账面总余额
                </div>
                <div class="stat-value">¥<%= FormatNumber(SafeNum(totalBookBalance), 2) %></div>
                <div class="stat-sub">所有账户合计</div>
            </div>
            <div class="stat-card available <%= IIF(totalAvailableBalance < totalBookBalance * 0.2, "highlight-warning", "") %>">
                <div class="stat-label">
                    <i class="fas fa-check-circle"></i> 可用余额
                </div>
                <div class="stat-value">¥<%= FormatNumber(SafeNum(totalAvailableBalance), 2) %></div>
                <div class="stat-sub">占总账面<%= SafeFormat(SafeDiv(totalAvailableBalance, totalBookBalance) * 100, 1) %>%</div>
            </div>
            <div class="stat-card frozen">
                <div class="stat-label">
                    <i class="fas fa-lock"></i> 冻结金额
                </div>
                <div class="stat-value">¥<%= FormatNumber(SafeNum(totalFrozenAmount), 2) %></div>
                <div class="stat-sub">保证金/冻结资金</div>
            </div>
            <div class="stat-card pending">
                <div class="stat-label">
                    <i class="fas fa-clock"></i> 待结算金额
                </div>
                <div class="stat-value">¥<%= FormatNumber(SafeNum(totalPendingSettlement), 2) %></div>
                <div class="stat-sub">在途资金</div>
            </div>
        </div>
        
        <!-- 区块2:各账户明细 -->
        <h3 class="section-title"><i class="fas fa-credit-card"></i> 账户明细</h3>
        <div class="accounts-grid">
            <% 
            Dim rsAccounts, accountIconClass, isAlert, isInactive
            Set rsAccounts = ExecuteQuery("SELECT * FROM FundAccounts ORDER BY IsActive DESC, AccountID DESC")
            
            If rsAccounts Is Nothing Then
            %>
            <div class="empty-state" style="grid-column: span 3;">
                <i class="fas fa-wallet"></i>
                <p>暂无资金账户</p>
                <% If isFinManager Then %>
                <p>点击右上角"添加账户"按钮创建</p>
                <% End If %>
            </div>
            <% 
            ElseIf rsAccounts.EOF Then
                ' 空结果集也需要关闭
                rsAccounts.Close
                Set rsAccounts = Nothing
            %>
            <div class="empty-state" style="grid-column: span 3;">
                <i class="fas fa-wallet"></i>
                <p>暂无资金账户</p>
                <% If isFinManager Then %>
                <p>点击右上角"添加账户"按钮创建</p>
                <% End If %>
            </div>
            <% 
            Else
                Do While Not rsAccounts.EOF
                    ' 确定图标样式
                    Select Case LCase(rsAccounts("AccountType"))
                        Case "alipay": accountIconClass = "icon-alipay"
                        Case "wechat": accountIconClass = "icon-wechat"
                        Case "bank": accountIconClass = "icon-bank"
                        Case "paypal": accountIconClass = "icon-paypal"
                        Case "cash": accountIconClass = "icon-cash"
                        Case Else: accountIconClass = "icon-other"
                    End Select
                    
                    isAlert = (SafeNum(rsAccounts("AvailableBalance")) < SafeNum(rsAccounts("AlertThreshold")))
                    isInactive = (rsAccounts("IsActive") = 0)
            %>
            <div class="account-card <%= IIF(isAlert, "alert", "") %> <%= IIF(isInactive, "inactive", "") %>" data-id="<%= rsAccounts("AccountID") %>">
                <div class="account-header">
                    <div class="account-name">
                        <div class="account-icon <%= accountIconClass %>">
                            <i class="fas <%= IIF(rsAccounts("AccountType")="alipay", "fa-alipay", IIF(rsAccounts("AccountType")="wechat", "fa-weixin", IIF(rsAccounts("AccountType")="bank", "fa-university", IIF(rsAccounts("AccountType")="paypal", "fa-paypal", "fa-wallet")))) %>"></i>
                        </div>
                        <%= Server.HTMLEncode(rsAccounts("AccountName")) %>
                    </div>
                    <span class="account-status <%= IIF(isInactive, "status-inactive", IIF(isAlert, "status-warning", "status-normal")) %>">
                        <%= IIF(isInactive, "停用", IIF(isAlert, "预警", "正常")) %>
                    </span>
                </div>
                
                <div class="account-balance">
                    <div class="balance-main">¥<%= FormatNumber(SafeNum(rsAccounts("TotalBalance")), 2) %></div>
                    <div class="balance-sub">
                        <span class="frozen"><i class="fas fa-lock"></i> 冻结 ¥<%= FormatNumber(SafeNum(rsAccounts("FrozenAmount")), 2) %></span>
                        <span class="available"><i class="fas fa-unlock"></i> 可用 ¥<%= FormatNumber(SafeNum(rsAccounts("AvailableBalance")), 2) %></span>
                    </div>
                </div>
                
                <div class="account-footer">
                    <span><i class="fas fa-sync"></i> <%= IIF(IsNull(rsAccounts("LastSyncAt")), "未同步", FormatDateField(rsAccounts("LastSyncAt"))) %></span>
                    <% If isFinManager Then %>
                    <div class="account-actions">
                        <button class="btn-icon edit" title="编辑" onclick="editAccount(<%= rsAccounts("AccountID") %>, '<%= SafeOutput(rsAccounts("AccountName")) %>', '<%= SafeOutput(rsAccounts("AccountType")) %>')">
                            <i class="fas fa-edit"></i>
                        </button>
                        <button class="btn-icon balance" title="更新余额" onclick="updateBalance(<%= rsAccounts("AccountID") %>, <%= SafeNum(rsAccounts("TotalBalance")) %>)">
                            <i class="fas fa-sync-alt"></i>
                        </button>
                        <button class="btn-icon threshold" title="设置阈值" onclick="setThreshold(<%= rsAccounts("AccountID") %>, <%= SafeNum(rsAccounts("AlertThreshold")) %>)">
                            <i class="fas fa-bell"></i>
                        </button>
                        <button class="btn-icon toggle <%= IIF(isInactive, "enable", "") %>" title="<%= IIF(isInactive, "启用", "停用") %>" onclick="toggleStatus(<%= rsAccounts("AccountID") %>, <%= IIF(isInactive, "1", "0") %>)">
                            <i class="fas <%= IIF(isInactive, "fa-play", "fa-pause") %>"></i>
                        </button>
                    </div>
                    <% End If %>
                </div>
            </div>
            <% 
                rsAccounts.MoveNext
                Loop
                rsAccounts.Close
            End If
            Set rsAccounts = Nothing
            %>
        </div>
        
        <!-- 区块3:待结算资金穿透 -->
        <h3 class="section-title"><i class="fas fa-hourglass-half"></i> 待结算资金穿透</h3>
        <div class="pending-section">
            <div class="pending-card orders" onclick="showPendingDetail('orders')">
                <div class="pending-header">
                    <i class="fas fa-shopping-cart"></i> 交易在途
                </div>
                <div class="pending-amount">¥<%= FormatNumber(SafeNum(pendingOrdersAmount), 2) %></div>
                <div class="pending-desc">Orders.Status='Paid' 未结算金额</div>
            </div>
            <div class="pending-card withdraw" onclick="showPendingDetail('withdraw')">
                <div class="pending-header">
                    <i class="fas fa-money-bill-wave"></i> 提现中
                </div>
                <div class="pending-amount">¥<%= FormatNumber(SafeNum(pendingWithdrawAmount), 2) %></div>
                <div class="pending-desc">PaymentRecords 提现待处理</div>
            </div>
            <div class="pending-card frozen" onclick="showPendingDetail('frozen')">
                <div class="pending-header">
                    <i class="fas fa-shield-alt"></i> 不可用保证金
                </div>
                <div class="pending-amount">¥<%= FormatNumber(SafeNum(totalFrozenAmount), 2) %></div>
                <div class="pending-desc">FundAccounts.FrozenAmount 明细</div>
            </div>
        </div>
        
        <!-- 区块4:收支预警 -->
        <h3 class="section-title"><i class="fas fa-exclamation-triangle"></i> 收支预警</h3>
        
        <% If largeTransferCount > 0 Then %>
        <div class="abnormal-alert">
            <i class="fas fa-radiation"></i>
            <div class="abnormal-content">
                <h4>异常监控提醒</h4>
                <p>检测到 <%= largeTransferCount %> 笔大额转账（单笔≥10000元），请核实</p>
            </div>
        </div>
        <% End If %>
        
        <div class="alert-section">
            <table class="alert-table">
                <thead>
                    <tr>
                        <th>账户名称</th>
                        <th>当前余额</th>
                        <th>预警阈值</th>
                        <th>差额</th>
                        <th>预警级别</th>
                    </tr>
                </thead>
                <tbody>
                    <% 
                    Dim rsAlerts, alertCount
                    alertCount = 0
                    Set rsAlerts = ExecuteQuery("SELECT * FROM FundAccounts WHERE IsActive = 1 AND AvailableBalance < AlertThreshold ORDER BY (AlertThreshold - AvailableBalance) DESC")
                    
                    Dim hasAlertData
                    hasAlertData = False
                    If Not (rsAlerts Is Nothing) Then
                        If Not rsAlerts.EOF Then
                            hasAlertData = True
                        End If
                    End If
                    
                    If Not hasAlertData Then
                    %>
                    <tr>
                        <td colspan="5" class="empty-state">
                            <i class="fas fa-check-circle" style="font-size: 24px; color: #38ef7d;"></i>
                            <p>暂无预警，所有账户余额正常</p>
                        </td>
                    </tr>
                    <% 
                    Else
                        Dim diffAmount, alertLevel, alertLevelClass
                        Do While Not rsAlerts.EOF
                            alertCount = alertCount + 1
                            diffAmount = SafeNum(rsAlerts("AlertThreshold")) - SafeNum(rsAlerts("AvailableBalance"))
                            
                            ' 预警级别判断
                            If diffAmount > SafeNum(rsAlerts("AlertThreshold")) * 0.5 Then
                                alertLevel = "高危"
                                alertLevelClass = "level-high"
                            ElseIf diffAmount > SafeNum(rsAlerts("AlertThreshold")) * 0.2 Then
                                alertLevel = "中危"
                                alertLevelClass = "level-medium"
                            Else
                                alertLevel = "低危"
                                alertLevelClass = "level-low"
                            End If
                    %>
                    <tr>
                        <td><%= Server.HTMLEncode(rsAlerts("AccountName")) %></td>
                        <td>¥<%= FormatNumber(SafeNum(rsAlerts("AvailableBalance")), 2) %></td>
                        <td>¥<%= FormatNumber(SafeNum(rsAlerts("AlertThreshold")), 2) %></td>
                        <td class="diff-negative">-¥<%= FormatNumber(SafeNum(diffAmount), 2) %></td>
                        <td><span class="alert-level <%= alertLevelClass %>"><%= alertLevel %></span></td>
                    </tr>
                    <% 
                        rsAlerts.MoveNext
                        Loop
                        rsAlerts.Close
                    End If
                    Set rsAlerts = Nothing
                    %>
                </tbody>
            </table>
        </div>
    </div>
    
    <% If isFinManager Then %>
    <!-- 添加账户模态框 -->
    <div class="modal" id="addAccountModal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-plus-circle"></i> 添加资金账户</h3>
                <button class="modal-close" onclick="closeModal('addAccountModal')">&times;</button>
            </div>
            <div class="modal-body">
                <form id="addAccountForm">
                    <%= GetCSRFTokenField() %>
                    <div class="form-group">
                        <label>账户名称</label>
                        <input type="text" name="accountName" required placeholder="如：支付宝主账户">
                    </div>
                    <div class="form-group">
                        <label>账户类型</label>
                        <select name="accountType" required>
                            <option value="alipay">支付宝</option>
                            <option value="wechat">微信支付</option>
                            <option value="bank">银行账户</option>
                            <option value="paypal">PayPal</option>
                            <option value="cash">现金</option>
                            <option value="other">其他</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>初始余额</label>
                        <input type="number" name="initialBalance" step="0.01" value="0" placeholder="0.00">
                    </div>
                    <div class="form-group">
                        <label>预警阈值</label>
                        <input type="number" name="alertThreshold" step="0.01" value="1000" placeholder="1000.00">
                    </div>
                </form>
            </div>
            <div class="modal-footer">
                <button class="btn-secondary" onclick="closeModal('addAccountModal')">取消</button>
                <button class="btn-primary" onclick="submitAddAccount()">确认添加</button>
            </div>
        </div>
    </div>
    
    <!-- 编辑账户模态框 -->
    <div class="modal" id="editAccountModal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-edit"></i> 编辑账户</h3>
                <button class="modal-close" onclick="closeModal('editAccountModal')">&times;</button>
            </div>
            <div class="modal-body">
                <form id="editAccountForm">
                    <%= GetCSRFTokenField() %>
                    <input type="hidden" name="accountId" id="editAccountId">
                    <div class="form-group">
                        <label>账户名称</label>
                        <input type="text" name="accountName" id="editAccountName" required>
                    </div>
                    <div class="form-group">
                        <label>账户类型</label>
                        <select name="accountType" id="editAccountType" required>
                            <option value="alipay">支付宝</option>
                            <option value="wechat">微信支付</option>
                            <option value="bank">银行账户</option>
                            <option value="paypal">PayPal</option>
                            <option value="cash">现金</option>
                            <option value="other">其他</option>
                        </select>
                    </div>
                </form>
            </div>
            <div class="modal-footer">
                <button class="btn-secondary" onclick="closeModal('editAccountModal')">取消</button>
                <button class="btn-primary" onclick="submitEditAccount()">保存修改</button>
            </div>
        </div>
    </div>
    
    <!-- 更新余额模态框 -->
    <div class="modal" id="updateBalanceModal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-sync-alt"></i> 更新余额</h3>
                <button class="modal-close" onclick="closeModal('updateBalanceModal')">&times;</button>
            </div>
            <div class="modal-body">
                <form id="updateBalanceForm">
                    <%= GetCSRFTokenField() %>
                    <input type="hidden" name="accountId" id="balanceAccountId">
                    <div class="form-group">
                        <label>最新余额</label>
                        <input type="number" name="newBalance" id="newBalance" step="0.01" required placeholder="0.00">
                    </div>
                </form>
            </div>
            <div class="modal-footer">
                <button class="btn-secondary" onclick="closeModal('updateBalanceModal')">取消</button>
                <button class="btn-primary" onclick="submitUpdateBalance()">确认更新</button>
            </div>
        </div>
    </div>
    
    <!-- 设置阈值模态框 -->
    <div class="modal" id="setThresholdModal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-bell"></i> 设置预警阈值</h3>
                <button class="modal-close" onclick="closeModal('setThresholdModal')">&times;</button>
            </div>
            <div class="modal-body">
                <form id="setThresholdForm">
                    <%= GetCSRFTokenField() %>
                    <input type="hidden" name="accountId" id="thresholdAccountId">
                    <div class="form-group">
                        <label>预警阈值</label>
                        <input type="number" name="threshold" id="thresholdValue" step="0.01" required placeholder="1000.00">
                        <small style="color: #666; display: block; margin-top: 5px;">当可用余额低于此值时将触发预警</small>
                    </div>
                </form>
            </div>
            <div class="modal-footer">
                <button class="btn-secondary" onclick="closeModal('setThresholdModal')">取消</button>
                <button class="btn-primary" onclick="submitSetThreshold()">确认设置</button>
            </div>
        </div>
    </div>
    <% End If %>
    
    <!-- 明细弹窗 -->
    <div class="modal" id="detailModal">
        <div class="modal-content" style="max-width: 600px;">
            <div class="modal-header">
                <h3 id="detailTitle"><i class="fas fa-list"></i> 明细</h3>
                <button class="modal-close" onclick="closeModal('detailModal')">&times;</button>
            </div>
            <div class="modal-body">
                <div class="detail-list" id="detailContent">
                    <!-- 动态加载 -->
                </div>
            </div>
        </div>
    </div>
    
    <script>
        // 模态框控制
        function openModal(modalId) {
            document.getElementById(modalId).classList.add('active');
        }
        
        function closeModal(modalId) {
            document.getElementById(modalId).classList.remove('active');
        }
        
        // 点击模态框背景关闭
        document.querySelectorAll('.modal').forEach(function(modal) {
            modal.addEventListener('click', function(e) {
                if (e.target === this) {
                    this.classList.remove('active');
                }
            });
        });
        
        <% If isFinManager Then %>
        // 编辑账户
        function editAccount(id, name, type) {
            document.getElementById('editAccountId').value = id;
            document.getElementById('editAccountName').value = name;
            document.getElementById('editAccountType').value = type;
            openModal('editAccountModal');
        }
        
        // 更新余额
        function updateBalance(id, currentBalance) {
            document.getElementById('balanceAccountId').value = id;
            document.getElementById('newBalance').value = currentBalance;
            openModal('updateBalanceModal');
        }
        
        // 设置阈值
        function setThreshold(id, currentThreshold) {
            document.getElementById('thresholdAccountId').value = id;
            document.getElementById('thresholdValue').value = currentThreshold;
            openModal('setThresholdModal');
        }
        
        // 停用/启用
        function toggleStatus(id, isActive) {
            var action = isActive ? '启用' : '停用';
            if (!confirm('确定要' + action + '该账户吗？')) return;
            
            var formData = new FormData();
            formData.append('ajax_action', 'toggle_status');
            formData.append('accountId', id);
            formData.append('isActive', isActive);
            formData.append('csrf_token', document.querySelector('input[name="csrf_token"]').value);
            
            fetch('fund_dashboard.asp', {
                method: 'POST',
                body: formData
            })
            .then(function(res) { return res.json(); })
            .then(function(data) {
                if (data.success) {
                    alert(data.message);
                    location.reload();
                } else {
                    alert(data.message);
                }
            })
            .catch(function(err) {
                alert('操作失败，请重试');
            });
        }
        
        // 提交添加账户
        function submitAddAccount() {
            var form = document.getElementById('addAccountForm');
            var formData = new FormData(form);
            formData.append('ajax_action', 'add_account');
            
            fetch('fund_dashboard.asp', {
                method: 'POST',
                body: formData
            })
            .then(function(res) { return res.json(); })
            .then(function(data) {
                if (data.success) {
                    alert(data.message);
                    location.reload();
                } else {
                    alert(data.message);
                }
            })
            .catch(function(err) {
                alert('添加失败，请重试');
            });
        }
        
        // 提交编辑账户
        function submitEditAccount() {
            var form = document.getElementById('editAccountForm');
            var formData = new FormData(form);
            formData.append('ajax_action', 'edit_account');
            
            fetch('fund_dashboard.asp', {
                method: 'POST',
                body: formData
            })
            .then(function(res) { return res.json(); })
            .then(function(data) {
                if (data.success) {
                    alert(data.message);
                    location.reload();
                } else {
                    alert(data.message);
                }
            })
            .catch(function(err) {
                alert('更新失败，请重试');
            });
        }
        
        // 提交更新余额
        function submitUpdateBalance() {
            var form = document.getElementById('updateBalanceForm');
            var formData = new FormData(form);
            formData.append('ajax_action', 'update_balance');
            
            fetch('fund_dashboard.asp', {
                method: 'POST',
                body: formData
            })
            .then(function(res) { return res.json(); })
            .then(function(data) {
                if (data.success) {
                    alert(data.message);
                    location.reload();
                } else {
                    alert(data.message);
                }
            })
            .catch(function(err) {
                alert('更新失败，请重试');
            });
        }
        
        // 提交设置阈值
        function submitSetThreshold() {
            var form = document.getElementById('setThresholdForm');
            var formData = new FormData(form);
            formData.append('ajax_action', 'set_threshold');
            
            fetch('fund_dashboard.asp', {
                method: 'POST',
                body: formData
            })
            .then(function(res) { return res.json(); })
            .then(function(data) {
                if (data.success) {
                    alert(data.message);
                    location.reload();
                } else {
                    alert(data.message);
                }
            })
            .catch(function(err) {
                alert('设置失败，请重试');
            });
        }
        <% End If %>
        
        // 显示待结算明细
        function showPendingDetail(type) {
            var titles = {
                'orders': '交易在途明细',
                'withdraw': '提现中明细',
                'frozen': '不可用保证金明细'
            };
            document.getElementById('detailTitle').innerHTML = '<i class="fas fa-list"></i> ' + titles[type];
            document.getElementById('detailContent').innerHTML = '<div class="empty-state"><i class="fas fa-spinner fa-spin"></i><p>加载中...</p></div>';
            openModal('detailModal');
            
            var formData = new FormData();
            formData.append('ajax_action', 'get_pending_detail');
            formData.append('detailType', type);
            <% If isFinManager Then %>
            formData.append('csrf_token', document.querySelector('input[name="csrf_token"]').value);
            <% End If %>
            
            fetch('fund_dashboard.asp', {
                method: 'POST',
                body: formData
            })
            .then(function(res) { return res.json(); })
            .then(function(data) {
                if (data.success && data.data.length > 0) {
                    var html = '';
                    data.data.forEach(function(item) {
                        var title, subtitle, amount;
                        if (type === 'orders') {
                            title = '订单 ' + item.no;
                            subtitle = item.method + ' | ' + item.time;
                            amount = '+¥' + parseFloat(item.amount).toFixed(2);
                        } else if (type === 'withdraw') {
                            title = item.no ? '提现 ' + item.no : '提现申请';
                            subtitle = item.time;
                            amount = '-¥' + parseFloat(item.amount).toFixed(2);
                        } else {
                            title = item.name;
                            subtitle = '冻结资金';
                            amount = '¥' + parseFloat(item.amount).toFixed(2);
                        }
                        
                        html += '<div class="detail-item">' +
                                '<div class="detail-info">' +
                                '<h4>' + title + '</h4>' +
                                '<p>' + subtitle + '</p>' +
                                '</div>' +
                                '<div class="detail-amount">' + amount + '</div>' +
                                '</div>';
                    });
                    document.getElementById('detailContent').innerHTML = html;
                } else {
                    document.getElementById('detailContent').innerHTML = '<div class="empty-state"><i class="fas fa-inbox"></i><p>暂无数据</p></div>';
                }
            })
            .catch(function(err) {
                document.getElementById('detailContent').innerHTML = '<div class="empty-state"><i class="fas fa-exclamation-circle"></i><p>加载失败</p></div>';
            });
        }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
