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

Function GetScalar(sql)
    Dim rs, val : val = ""
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

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

' 自动创建 AfterSales 表
On Error Resume Next
conn.Execute "SELECT TOP 1 * FROM AfterSales WHERE 1=0"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE AfterSales (" & _
        "AfterSalesID INT IDENTITY(1,1) PRIMARY KEY," & _
        "OrderID INT NOT NULL," & _
        "UserID INT NOT NULL," & _
        "RequestType NVARCHAR(20) NOT NULL," & _
        "Reason NVARCHAR(500)," & _
        "Status NVARCHAR(20) DEFAULT 'pending'," & _
        "RefundAmount DECIMAL(10,2)," & _
        "AdminNotes NVARCHAR(500)," & _
        "ProcessedBy INT NULL," & _
        "ProcessedAt DATETIME NULL," & _
        "CreatedAt DATETIME DEFAULT GETDATE()" & _
        ")"
End If
On Error GoTo 0

Dim msg : msg = ""
Dim msgType : msgType = ""

' 处理 POST
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim action : action = Request.Form("action")
    
    If action = "process" Then
        Dim asID : asID = CInt(Request.Form("asID"))
        Dim newStatus : newStatus = Request.Form("status")
        Dim adminNotes : adminNotes = Request.Form("adminNotes")
        
        conn.Execute "UPDATE AfterSales SET Status='" & Replace(newStatus,"'","''") & "', " & _
            "AdminNotes='" & Replace(adminNotes,"'","''") & "', " & _
            "ProcessedBy=" & Session("AdminID") & ", " & _
            "ProcessedAt=GETDATE() WHERE AfterSalesID=" & asID
        
        ' 如果是退款审批,更新订单状态
        If newStatus = "approved" Then
            Dim asType : asType = GetScalar("SELECT RequestType FROM AfterSales WHERE AfterSalesID=" & asID)
            Dim asOrderID : asOrderID = GetScalar("SELECT OrderID FROM AfterSales WHERE AfterSalesID=" & asID)
            If asType = "refund" And asOrderID <> "" Then
                conn.Execute "UPDATE Orders SET Status='refunding' WHERE OrderID=" & asOrderID
            End If
        End If
        
        msg = "售后申请已处理"
        msgType = "success"
    End If
End If

' 筛选
Dim filterStatus : filterStatus = Request.QueryString("status")
Dim sqlWhere : sqlWhere = "WHERE 1=1"
If filterStatus <> "" Then sqlWhere = sqlWhere & " AND a.Status='" & Replace(filterStatus,"'","''") & "'"

' 获取售后列表
Dim rsAfterSales
Set rsAfterSales = ExecuteQuery("SELECT a.*, o.OrderNo, o.TotalAmount, u.Username " & _
    "FROM AfterSales a LEFT JOIN Orders o ON a.OrderID=o.OrderID LEFT JOIN Users u ON a.UserID=u.UserID " & _
    sqlWhere & " ORDER BY a.CreatedAt DESC")

' 统计
Dim pendingCnt : pendingCnt = SafeNum(GetScalar("SELECT COUNT(*) FROM AfterSales WHERE Status='pending'"))
Dim approvedCnt : approvedCnt = SafeNum(GetScalar("SELECT COUNT(*) FROM AfterSales WHERE Status='approved'"))
Dim rejectedCnt : rejectedCnt = SafeNum(GetScalar("SELECT COUNT(*) FROM AfterSales WHERE Status='rejected'"))
Dim totalCnt : totalCnt = SafeNum(GetScalar("SELECT COUNT(*) FROM AfterSales"))
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>售后管理 - 运营管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; font-family: 'Segoe UI',Arial,sans-serif; }
        .main-content { margin-left: 260px; padding: 30px; min-height: 100vh; }
        .page-header { margin-bottom: 25px; }
        .page-title { color: #fff; font-size: 24px; margin: 0 0 8px; }
        .breadcrumb { color: #888; font-size: 13px; }
        .breadcrumb a { color: #00bcd4; text-decoration: none; }
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); padding: 20px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); text-align: center; cursor: pointer; transition: all 0.2s; }
        .stat-card:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(0,0,0,0.3); }
        .stat-card .num { font-size: 28px; font-weight: 700; color: #00bcd4; }
        .stat-card .label { font-size: 13px; color: #888; margin-top: 5px; }
        .stat-card.pending .num { color: #FF9800; }
        .stat-card.approved .num { color: #4CAF50; }
        .stat-card.rejected .num { color: #f44336; }
        .alert { padding: 12px 20px; border-radius: 8px; margin-bottom: 20px; font-weight: 500; }
        .alert-success { background: rgba(76,175,80,0.15); color: #4CAF50; border: 1px solid rgba(76,175,80,0.3); }
        .admin-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); margin-bottom: 20px; overflow: hidden; }
        .admin-card-header { padding: 15px 20px; border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 16px; display: flex; justify-content: space-between; align-items: center; }
        .admin-card-body { padding: 20px; }
        .admin-table { width: 100%; border-collapse: collapse; }
        .admin-table th { text-align: left; padding: 12px; background: rgba(0,188,212,0.06); border-bottom: 1px solid rgba(255,255,255,0.06); font-size: 13px; color: #999; }
        .admin-table td { padding: 12px; border-bottom: 1px solid rgba(255,255,255,0.04); font-size: 14px; }
        .admin-table tr:hover { background: rgba(255,255,255,0.03); }
        .status-badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
        .status-pending { background: rgba(255,152,0,0.2); color: #FFB74D; }
        .status-approved { background: rgba(76,175,80,0.2); color: #A5D6A7; }
        .status-rejected { background: rgba(244,67,54,0.2); color: #EF9A9A; }
        .status-refund { background: rgba(233,30,99,0.2); color: #F48FB1; }
        .status-return { background: rgba(33,150,243,0.2); color: #90CAF9; }
        .status-exchange { background: rgba(156,39,176,0.2); color: #CE93D8; }
        .modal-overlay { display: none; position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.7); z-index: 1000; justify-content: center; align-items: center; }
        .modal-overlay.show { display: flex; }
        .modal { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 16px; border: 1px solid rgba(255,255,255,0.1); padding: 30px; width: 90%; max-width: 500px; }
        .modal h3 { margin: 0 0 15px; font-size: 18px; color: #fff; }
        .form-group { margin-bottom: 15px; }
        .form-group label { display: block; color: #999; font-size: 13px; margin-bottom: 5px; }
        .form-group select, .form-group textarea { width: 100%; padding: 10px; border: 1px solid rgba(255,255,255,0.15); border-radius: 8px; background: #1a1a2e; color: #e0e0e0; font-size: 14px; }
        .form-group textarea { min-height: 80px; resize: vertical; }
        .modal-actions { display: flex; gap: 10px; justify-content: flex-end; margin-top: 20px; }
        .empty { text-align: center; padding: 40px; color: #888; }
        .empty i { font-size: 40px; display: block; margin-bottom: 10px; opacity: 0.3; }
        @media (max-width: 768px) { .main-content { margin-left: 0; } .stats-grid { grid-template-columns: 1fr 1fr; } }
    </style>
</head>
<body data-theme="operation-dark">
<!--#include file="includes/nav.asp"-->
<div class="main-content">
    <div class="page-header">
        <h2 class="page-title"><i class="fas fa-headset"></i> 售后管理</h2>
        <div class="breadcrumb"><a href="index.asp">运营中心</a> / 售后管理</div>
    </div>

    <% If msg <> "" Then %>
    <div class="alert alert-<%= msgType %>"><i class="fas fa-info-circle"></i> <%= msg %></div>
    <% End If %>

    <div class="stats-grid">
        <div class="stat-card pending" onclick="location.href='?status=pending'"><div class="num"><%= pendingCnt %></div><div class="label">待处理</div></div>
        <div class="stat-card approved" onclick="location.href='?status=approved'"><div class="num"><%= approvedCnt %></div><div class="label">已通过</div></div>
        <div class="stat-card rejected" onclick="location.href='?status=rejected'"><div class="num"><%= rejectedCnt %></div><div class="label">已拒绝</div></div>
        <div class="stat-card" onclick="location.href='?'"><div class="num"><%= totalCnt %></div><div class="label">全部记录</div></div>
    </div>

    <div class="admin-card">
        <div class="admin-card-header">
            <span><i class="fas fa-list"></i> 售后申请列表 <%= IIf(filterStatus<>"","(" & filterStatus & ")","") %></span>
            <a href="?" class="btn btn-primary btn-sm" style="<%= IIf(filterStatus="","display:none;","") %>"><i class="fas fa-list"></i> 显示全部</a>
        </div>
        <div class="admin-card-body">
            <table class="admin-table">
                <thead>
                    <tr>
                        <th>ID</th><th>订单号</th><th>客户</th><th>类型</th><th>原因</th><th>金额</th><th>状态</th><th>申请时间</th><th>操作</th>
                    </tr>
                </thead>
                <tbody>
                    <% If Not rsAfterSales Is Nothing Then
                    Dim hasRows : hasRows = False
                    Do While Not rsAfterSales.EOF
                        hasRows = True
                        Dim typeBadge
                        Select Case rsAfterSales("RequestType")
                            Case "refund": typeBadge = "status-refund"
                            Case "return": typeBadge = "status-return"
                            Case "exchange": typeBadge = "status-exchange"
                            Case Else: typeBadge = ""
                        End Select
                        Dim statusBadge
                        Select Case rsAfterSales("Status")
                            Case "pending": statusBadge = "status-pending"
                            Case "approved": statusBadge = "status-approved"
                            Case "rejected": statusBadge = "status-rejected"
                            Case Else: statusBadge = ""
                        End Select
                        Dim requestTypeCN
                        Select Case rsAfterSales("RequestType")
                            Case "refund": requestTypeCN = "退款"
                            Case "return": requestTypeCN = "退货"
                            Case "exchange": requestTypeCN = "换货"
                            Case Else: requestTypeCN = rsAfterSales("RequestType")
                        End Select
                    %>
                    <tr>
                        <td><%= rsAfterSales("AfterSalesID") %></td>
                        <td><a href="order_detail.asp?id=<%= rsAfterSales("OrderID") %>" style="color:#00bcd4;"><%= rsAfterSales("OrderNo") %></a></td>
                        <td><%= rsAfterSales("Username") %></td>
                        <td><span class="status-badge <%= typeBadge %>"><%= requestTypeCN %></span></td>
                        <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;"><%= rsAfterSales("Reason") %></td>
                        <td><%= IIf(IsNull(rsAfterSales("RefundAmount")),FormatNumber(SafeNum(rsAfterSales("TotalAmount")),2),FormatNumber(rsAfterSales("RefundAmount"),2)) %> 元</td>
                        <td><span class="status-badge <%= statusBadge %>"><%= rsAfterSales("Status") %></span></td>
                        <td style="font-size:13px;color:#999;"><%= rsAfterSales("CreatedAt") %></td>
                        <td>
                            <% If rsAfterSales("Status") = "pending" Then %>
                            <button class="btn btn-primary btn-sm" onclick="openProcessModal(<%= rsAfterSales("AfterSalesID") %>, '<%= rsAfterSales("OrderNo") %>', '<%= rsAfterSales("Username") %>', '<%= requestTypeCN %>', '<%= FormatNumber(SafeNum(rsAfterSales("RefundAmount")),2) %>')">
                                <i class="fas fa-gavel"></i> 处理
                            </button>
                            <% Else %>
                            <span style="font-size:12px;color:#888;"><%= rsAfterSales("AdminNotes") %></span>
                            <% End If %>
                        </td>
                    </tr>
                    <%
                        rsAfterSales.MoveNext
                    Loop
                    rsAfterSales.Close
                    If Not hasRows Then
                    %>
                    <tr><td colspan="9" class="empty"><i class="fas fa-check-circle" style="color:#4CAF50;"></i>暂无售后申请</td></tr>
                    <% End If
                    Else %>
                    <tr><td colspan="9" class="empty"><i class="fas fa-check-circle" style="color:#4CAF50;"></i>暂无售后申请</td></tr>
                    <% End If %>
                </tbody>
            </table>
        </div>
    </div>
</div>

<!-- 处理弹窗 -->
<div class="modal-overlay" id="processModal">
    <div class="modal">
        <h3><i class="fas fa-gavel"></i> 处理售后申请</h3>
        <div style="font-size:13px;color:#999;margin-bottom:15px;">订单: <strong id="modalOrderNo" style="color:#00bcd4;"></strong> | 客户: <strong id="modalUser" style="color:#e0e0e0;"></strong> | 类型: <strong id="modalType"></strong> | 金额: <strong id="modalAmount" style="color:#FF9800;"></strong></div>
        <form method="post">
            <input type="hidden" name="action" value="process">
            <input type="hidden" name="asID" id="modalAsID">
            <div class="form-group">
                <label>处理结果</label>
                <select name="status" id="modalStatus" required>
                    <option value="approved">通过</option>
                    <option value="rejected">拒绝</option>
                </select>
            </div>
            <div class="form-group">
                <label>处理备注</label>
                <textarea name="adminNotes" placeholder="输入处理备注..."></textarea>
            </div>
            <div class="modal-actions">
                <button type="button" class="btn btn-danger" onclick="closeProcessModal()">取消</button>
                <button type="submit" class="btn btn-success"><i class="fas fa-check"></i> 确认处理</button>
            </div>
        </form>
    </div>
</div>

<script>
function openProcessModal(id, orderNo, username, type, amount) {
    document.getElementById('modalAsID').value = id;
    document.getElementById('modalOrderNo').innerText = orderNo;
    document.getElementById('modalUser').innerText = username;
    document.getElementById('modalType').innerText = type;
    document.getElementById('modalAmount').innerText = amount + ' 元';
    document.getElementById('processModal').classList.add('show');
}
function closeProcessModal() {
    document.getElementById('processModal').classList.remove('show');
}
document.getElementById('processModal').addEventListener('click', function(e) {
    if (e.target === this) closeProcessModal();
});
</script>
</body>
</html>
<%
Call CloseConnection()
%>
