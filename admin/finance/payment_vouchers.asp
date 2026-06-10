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
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE GLTransactions (GLID INT IDENTITY(1,1) PRIMARY KEY, GLNo NVARCHAR(50), TransactionDate DATETIME, AccountCode NVARCHAR(50), AccountName NVARCHAR(100), DebitAmount DECIMAL(18,2) DEFAULT 0, CreditAmount DECIMAL(18,2) DEFAULT 0, CenterID INT, RefType NVARCHAR(50), RefID INT, RefNo NVARCHAR(50), Description NVARCHAR(500), CreatedBy NVARCHAR(50), CreatedAt DATETIME DEFAULT GETDATE())"
End If
Err.Clear

conn.Execute "SELECT PaymentType FROM PaymentRecords WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PaymentRecords ADD PaymentType NVARCHAR(30) DEFAULT 'Receipt'"
conn.Execute "SELECT VoucherNo FROM PaymentRecords WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PaymentRecords ADD VoucherNo NVARCHAR(50)"
On Error GoTo 0

' POST 处理
Dim action : action = Request.Form("action")
If action = "create_voucher" Then
    Dim pvType : pvType = Request.Form("paymentType")
    Dim pvRefID : pvRefID = SafeNum(Request.Form("refID"))
    Dim pvAmount : pvAmount = SafeNum(Request.Form("amount"))
    Dim pvMethod : pvMethod = Replace(Request.Form("method"),"'","''")
    Dim pvNotes : pvNotes = Replace(Request.Form("notes"),"'","''")
    Dim pvCenterID : pvCenterID = SafeNum(Request.Form("centerID"))
    
    If pvAmount > 0 Then
        Dim vNo : vNo = "PV" & Year(Now) & Right("0"&Month(Now),2) & Right("0"&Day(Now),2) & Right("0"&Hour(Now),2) & Right("0"&Minute(Now),2)
        Dim pvSql
        pvSql = "INSERT INTO PaymentRecords (TransactionNo, PaymentType, " & _
            IIf(pvType="Payment", "PayableID", "ReceivableID") & ", Amount, PaymentMethod, Remark, TransactionType, Status, CenterID, VoucherNo, CreatedAt) VALUES ('" & _
            vNo & "','" & Replace(pvType,"'","''") & "'," & pvRefID & "," & pvAmount & ",'" & pvMethod & "','" & pvNotes & "','" & IIf(pvType="Payment","支出","收入") & "','Completed'," & pvCenterID & ",'" & vNo & "',GETDATE())"
        
        conn.Execute pvSql
        If Err.Number <> 0 Then Err.Clear
        
        ' 更新应付/应收
        If pvType = "Payment" Then
            conn.Execute "UPDATE AccountsPayable SET PaidAmount = PaidAmount + " & pvAmount & ", Status = CASE WHEN PaidAmount + " & pvAmount & " >= Amount THEN 'Paid' ELSE 'Partial' END, UpdatedAt=GETDATE() WHERE PayableID=" & pvRefID
        Else
            conn.Execute "UPDATE AccountsReceivable SET ReceivedAmount = ReceivedAmount + " & pvAmount & ", Status = CASE WHEN ReceivedAmount + " & pvAmount & " >= Amount THEN 'Received' ELSE 'Partial' END, UpdatedAt=GETDATE() WHERE ReceivableID=" & pvRefID
        End If
        
        ' 记录总账
        conn.Execute "INSERT INTO GLTransactions (GLNo, TransactionDate, AccountCode, AccountName, " & IIf(pvType="Payment","DebitAmount","CreditAmount") & ", CenterID, RefType, RefID, RefNo, Description, CreatedBy, CreatedAt) VALUES ('" & _
            vNo & "',GETDATE(),'" & IIf(pvType="Payment","5001","4001") & "','" & IIf(pvType="Payment","应付账款","应收账款") & "'," & pvAmount & "," & pvCenterID & ",'" & IIf(pvType="Payment","APPayment","ARReceipt") & "'," & pvRefID & ",'" & vNo & "','" & pvNotes & "','" & Replace(Session("AdminUsername"),"'","''") & "',GETDATE())"
        If Err.Number <> 0 Then Err.Clear
        
        Response.Redirect "payment_vouchers.asp?msg=" & Server.URLEncode("凭证创建成功：" & vNo)
        Response.End
    End If
End If

' 统计数据
Dim pvCount : pvCount = GetScalar("SELECT COUNT(*) FROM PaymentRecords WHERE VoucherNo IS NOT NULL")
Dim pvToday : pvToday = GetScalar("SELECT COUNT(*) FROM PaymentRecords WHERE VoucherNo IS NOT NULL AND CAST(CreatedAt AS DATE)=CAST(GETDATE() AS DATE)")
Dim pvTotalIn : pvTotalIn = GetScalar("SELECT ISNULL(SUM(Amount),0) FROM PaymentRecords WHERE VoucherNo IS NOT NULL AND PaymentType='Receipt'")
Dim pvTotalOut : pvTotalOut = GetScalar("SELECT ISNULL(SUM(Amount),0) FROM PaymentRecords WHERE VoucherNo IS NOT NULL AND PaymentType='Payment'")

' 凭证列表
Dim rsPV, pvList(), pvPCount : pvPCount = 0
Set rsPV = conn.Execute("SELECT * FROM PaymentRecords WHERE VoucherNo IS NOT NULL ORDER BY CreatedAt DESC")
If Not rsPV Is Nothing Then
    If Not rsPV.EOF Then
        rsPV.MoveLast : pvPCount = rsPV.RecordCount : rsPV.MoveFirst
        ReDim pvList(pvPCount - 1, 7)
        Dim pvi : pvi = 0
        Do While Not rsPV.EOF
            pvList(pvi, 0) = rsPV("VoucherNo") & ""
            pvList(pvi, 1) = rsPV("PaymentType") & ""
            pvList(pvi, 2) = rsPV("Amount")
            pvList(pvi, 3) = rsPV("PaymentMethod") & ""
            pvList(pvi, 4) = rsPV("Remark") & ""
            pvList(pvi, 5) = rsPV("Status") & ""
            pvList(pvi, 6) = rsPV("CreatedAt") & ""
            pvList(pvi, 7) = rsPV("TransactionType") & ""
            pvi = pvi + 1
            rsPV.MoveNext
        Loop
    End If
    rsPV.Close : Set rsPV = Nothing
End If

' 应付列表(用于付款选择)
Dim rsAPSel : Set rsAPSel = conn.Execute("SELECT PayableID, PayableNo, SupplierName, Amount-PaidAmount AS Balance FROM AccountsPayable WHERE Status IN ('Pending','Partial') ORDER BY DueDate")
' 应收列表
Dim rsARSel : Set rsARSel = conn.Execute("SELECT ReceivableID, ReceivableNo, CustomerName, Amount-ReceivedAmount AS Balance FROM AccountsReceivable WHERE Status IN ('Pending','Partial') ORDER BY DueDate")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>付款凭证 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
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
        .stat-card.cyan { border-top: 4px solid #00bcd4; } .stat-card.cyan .num { color: #00bcd4; }
        .stat-card.green { border-top: 4px solid #4CAF50; } .stat-card.green .num { color: #4CAF50; }
        .stat-card.orange { border-top: 4px solid #FF9800; } .stat-card.orange .num { color: #FF9800; }
        .stat-card.blue { border-top: 4px solid #2196F3; } .stat-card.blue .num { color: #2196F3; }
        .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        .card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); margin-bottom: 20px; overflow: hidden; }
        .card-header { padding: 15px 20px; border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; display: flex; justify-content: space-between; align-items: center; }
        .card-body { padding: 20px; }
        table { width: 100%; border-collapse: collapse; }
        th { text-align: left; padding: 10px; background: rgba(0,188,212,0.06); border-bottom: 1px solid rgba(255,255,255,0.06); font-size: 13px; color: #999; }
        td { padding: 10px; border-bottom: 1px solid rgba(255,255,255,0.04); font-size: 14px; }
        tr:hover { background: rgba(255,255,255,0.02); }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; }
        .badge-success { background: rgba(76,175,80,0.2); color: #81C784; }
        .badge-info { background: rgba(0,188,212,0.2); color: #80DEEA; }
        .amount { font-weight: 600; }
        .amount.in { color: #4CAF50; }
        .amount.out { color: #f44336; }
.modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.7); z-index: 1000; }
        .modal-content { background: linear-gradient(135deg, #2d2d44, #1e1e32); width: 90%; max-width: 550px; margin: 60px auto; padding: 30px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.1); }
        .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }
        .modal-header h3 { color: #fff; margin: 0; }
        .close-btn { background: none; border: none; color: #888; font-size: 24px; cursor: pointer; }
        .form-group { margin-bottom: 15px; }
        .form-group label { display: block; color: #b0b0b0; margin-bottom: 6px; font-size: 13px; }
        .form-group input, .form-group select, .form-group textarea { width: 100%; padding: 10px; background: #1e1e32; border: 1px solid rgba(255,255,255,0.1); border-radius: 6px; color: #e0e0e0; font-size: 14px; box-sizing: border-box; }
        .form-group textarea { height: 80px; resize: vertical; }
        .form-actions { text-align: right; margin-top: 20px; }
        .msg-success { padding: 12px 20px; border-radius: 6px; margin-bottom: 16px; background: rgba(76,175,80,0.15); color: #4CAF50; border: 1px solid rgba(76,175,80,0.3); }
        @media (max-width: 768px) { .main-content { margin-left: 0; } .grid-2 { grid-template-columns: 1fr; } .stats-grid { grid-template-columns: 1fr 1fr; } }
    </style>
</head>
<body>
<!--#include file="includes/nav.asp"-->
<div class="main-content">
    <div class="page-header">
        <h2 class="page-title"><i class="fas fa-receipt"></i> 付款凭证</h2>
        <div class="breadcrumb"><a href="index.asp">财务中心</a> / 付款凭证</div>
    </div>
    <% If Request.QueryString("msg") <> "" Then %>
    <div class="msg-success"><i class="fas fa-check-circle"></i> <%= Server.HTMLEncode(Request.QueryString("msg")) %></div>
    <% End If %>
    
    <div class="stats-grid">
        <div class="stat-card cyan"><div class="num"><%= pvCount %></div><div class="label">凭证总数</div></div>
        <div class="stat-card blue"><div class="num"><%= pvToday %></div><div class="label">今日凭证</div></div>
        <div class="stat-card green"><div class="num">¥<%= FormatNumber(pvTotalIn, 0) %></div><div class="label">收款总额</div></div>
        <div class="stat-card orange"><div class="num">¥<%= FormatNumber(pvTotalOut, 0) %></div><div class="label">付款总额</div></div>
    </div>
    
    <div class="grid-2">
        <!-- 应付(付款)凭证创建 -->
        <div class="card">
            <div class="card-header"><i class="fas fa-money-bill-wave"></i> 创建付款凭证</div>
            <div class="card-body">
                <button class="btn btn-primary" onclick="openVoucher('Payment')"><i class="fas fa-plus"></i> 新建付款</button>
                <% If Not rsAPSel Is Nothing Then %>
                <table style="margin-top:15px">
                    <thead><tr><th>编号</th><th>供应商</th><th>余额</th></tr></thead>
                    <tbody>
                    <% Do While Not rsAPSel.EOF %>
                    <tr>
                        <td><%= rsAPSel("PayableNo") %></td>
                        <td><%= rsAPSel("SupplierName") %></td>
                        <td class="amount out">¥<%= FormatNumber(SafeNum(rsAPSel("Balance")), 2) %></td>
                    </tr>
                    <% rsAPSel.MoveNext : Loop
                    rsAPSel.Close : Set rsAPSel = Nothing %>
                    </tbody>
                </table>
                <% End If %>
            </div>
        </div>
        
        <!-- 应收(收款)凭证创建 -->
        <div class="card">
            <div class="card-header"><i class="fas fa-hand-holding-usd"></i> 创建收款凭证</div>
            <div class="card-body">
                <button class="btn btn-primary" onclick="openVoucher('Receipt')"><i class="fas fa-plus"></i> 新建收款</button>
                <% If Not rsARSel Is Nothing Then %>
                <table style="margin-top:15px">
                    <thead><tr><th>编号</th><th>客户</th><th>余额</th></tr></thead>
                    <tbody>
                    <% Do While Not rsARSel.EOF %>
                    <tr>
                        <td><%= rsARSel("ReceivableNo") %></td>
                        <td><%= rsARSel("CustomerName") %></td>
                        <td class="amount in">¥<%= FormatNumber(SafeNum(rsARSel("Balance")), 2) %></td>
                    </tr>
                    <% rsARSel.MoveNext : Loop
                    rsARSel.Close : Set rsARSel = Nothing %>
                    </tbody>
                </table>
                <% End If %>
            </div>
        </div>
    </div>
    
    <!-- 凭证记录列表 -->
    <div class="card">
        <div class="card-header"><i class="fas fa-history"></i> 凭证记录</div>
        <div class="card-body">
            <table>
                <thead><tr><th>凭证号</th><th>类型</th><th>金额</th><th>支付方式</th><th>备注</th><th>时间</th></tr></thead>
                <tbody>
                    <% If pvPCount > 0 Then
                        For pvi = 0 To pvPCount - 1
                    %>
                    <tr>
                        <td><strong><%= pvList(pvi, 0) %></strong></td>
                        <td><span class="badge <%= IIf(pvList(pvi,1)="Receipt","badge-success","badge-info") %>"><%= IIf(pvList(pvi,1)="Receipt","收款","付款") %></span></td>
                        <td class="amount <%= IIf(pvList(pvi,1)="Receipt","in","out") %>"><%= IIf(pvList(pvi,1)="Receipt","+","-") %>¥<%= FormatNumber(SafeNum(pvList(pvi,2)), 2) %></td>
                        <td><%= pvList(pvi, 3) %></td>
                        <td><%= Left(pvList(pvi,4), 30) %></td>
                        <td><%= pvList(pvi, 6) %></td>
                    </tr>
                    <% Next
                    Else %>
                    <tr><td colspan="6" style="text-align:center;padding:40px;color:#666;">暂无凭证记录</td></tr>
                    <% End If %>
                </tbody>
            </table>
        </div>
    </div>
</div>

<!-- 凭证创建弹窗 -->
<div id="voucherModal" class="modal">
    <div class="modal-content">
        <div class="modal-header">
            <h3 id="vModalTitle"><i class="fas fa-plus-circle"></i> 创建付款凭证</h3>
            <button class="close-btn" onclick="closeModal()">&times;</button>
        </div>
        <form method="post" action="payment_vouchers.asp">
            <input type="hidden" name="action" value="create_voucher">
            <input type="hidden" name="paymentType" id="pvType">
            <div class="form-group"><label>关联 *</label>
                <select name="refID" id="pvRefID" required><option value="">请选择</option></select>
            </div>
            <div class="form-group"><label>金额 *</label><input type="number" name="amount" step="0.01" required></div>
            <div class="form-group"><label>支付方式</label>
                <select name="method"><option value="银行转账">银行转账</option><option value="现金">现金</option><option value="支付宝">支付宝</option><option value="微信支付">微信支付</option></select>
            </div>
            <div class="form-group"><label>成本中心</label>
                <select name="centerID">
                    <option value="0">无</option>
                    <% 
                    Dim rsCCOpt : Set rsCCOpt = conn.Execute("SELECT CenterID, CenterName FROM CostCenters WHERE IsActive=1 ORDER BY CenterName")
                    If Not rsCCOpt Is Nothing Then
                        Do While Not rsCCOpt.EOF
                            Response.Write "<option value=""" & rsCCOpt("CenterID") & """>" & rsCCOpt("CenterName") & "</option>"
                            rsCCOpt.MoveNext
                        Loop
                        rsCCOpt.Close
                    End If
                    Set rsCCOpt = Nothing
                    %>
                </select>
            </div>
            <div class="form-group"><label>备注</label><textarea name="notes"></textarea></div>
            <div class="form-actions">
                <button type="button" class="btn btn--neutral" onclick="closeModal()">取消</button>
                <button type="submit" class="btn btn-primary"><i class="fas fa-check"></i> 确认</button>
            </div>
        </form>
    </div>
</div>
<script>
function openVoucher(type){
    document.getElementById('pvType').value = type;
    var sel = document.getElementById('pvRefID');
    sel.innerHTML = '<option value="">请选择</option>';
    var title = type === 'Payment' ? '创建付款凭证' : '创建收款凭证';
    document.getElementById('vModalTitle').innerHTML = '<i class="fas fa-plus-circle"></i> ' + title;
    
    // 动态加载选项
    var items = type === 'Payment' 
        ? [<% If Not rsAPSel Is Nothing Then rsAPSel.MoveFirst : Do While Not rsAPSel.EOF : Response.Write "{id:" & rsAPSel("PayableID") & ",name:'" & rsAPSel("PayableNo") & " " & rsAPSel("SupplierName") & " (" & FormatNumber(SafeNum(rsAPSel("Balance")),0) & ")'}," : rsAPSel.MoveNext : Loop : rsAPSel.Close : Set rsAPSel = Nothing End If %>]
        : [<% If Not rsARSel Is Nothing Then rsARSel.MoveFirst : Do While Not rsARSel.EOF : Response.Write "{id:" & rsARSel("ReceivableID") & ",name:'" & rsARSel("ReceivableNo") & " " & rsARSel("CustomerName") & " (" & FormatNumber(SafeNum(rsARSel("Balance")),0) & ")'}," : rsARSel.MoveNext : Loop : rsARSel.Close : Set rsARSel = Nothing End If %>];
    
    items.forEach(function(item){
        if(item.id){
            var opt = document.createElement('option');
            opt.value = item.id;
            opt.textContent = item.name;
            sel.appendChild(opt);
        }
    });
    document.getElementById('voucherModal').style.display='block';
}
function closeModal(){ document.getElementById('voucherModal').style.display='none'; }
window.onclick=function(e){ if(e.target.classList.contains('modal')) e.target.style.display='none'; }
</script>
</body>
</html>
<%
Call CloseConnection()
%>
