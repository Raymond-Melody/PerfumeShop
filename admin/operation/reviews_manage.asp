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

Dim msg : msg = ""

' POST 批量操作
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim bulkAction : bulkAction = Request.Form("bulk_action")
    If bulkAction <> "" Then
        Dim selectedIDs : selectedIDs = Request.Form("review_ids")
        If selectedIDs <> "" Then
            Dim idsArr : idsArr = Split(selectedIDs, ",")
            Dim newStatus, id
            Select Case bulkAction
                Case "approve": newStatus = "Approved"
                Case "reject": newStatus = "Rejected"
                Case "pending": newStatus = "Pending"
            End Select
            For Each id In idsArr
                If IsNumeric(id) And Trim(id) <> "" Then
                    conn.Execute "UPDATE ProductReviews SET [Status]='" & newStatus & "', UpdatedAt=GETDATE() WHERE ReviewID=" & CLng(id)
                End If
            Next
            msg = "已批量更新 " & UBound(idsArr)+1 & " 条评价"
        End If
    End If
End If

' 筛选
Dim filterStatus : filterStatus = Request.QueryString("status")
Dim keyword : keyword = Request.QueryString("keyword")

Dim sqlWhere : sqlWhere = "WHERE 1=1"
If filterStatus <> "" Then sqlWhere = sqlWhere & " AND r.[Status]='" & Replace(filterStatus,"'","''") & "'"
If keyword <> "" Then sqlWhere = sqlWhere & " AND (u.Username LIKE '%" & Replace(keyword,"'","''") & "%' OR r.Comment LIKE '%" & Replace(keyword,"'","''") & "%')"

' 获取评价列表
Dim rsReviews
Set rsReviews = ExecuteQuery("SELECT TOP 100 r.*, u.Username, o.OrderNo " & _
    "FROM ProductReviews r LEFT JOIN Users u ON r.UserID=u.UserID LEFT JOIN Orders o ON r.OrderID=o.OrderID " & _
    sqlWhere & " ORDER BY r.CreatedAt DESC")

' 统计
Dim pendingCount : pendingCount = CLng("0" & GetScalar("SELECT COUNT(*) FROM ProductReviews WHERE [Status]='Pending'"))
Dim approvedCount : approvedCount = CLng("0" & GetScalar("SELECT COUNT(*) FROM ProductReviews WHERE [Status]='Approved'"))
Dim rejectedCount : rejectedCount = CLng("0" & GetScalar("SELECT COUNT(*) FROM ProductReviews WHERE [Status]='Rejected'"))
Dim totalCount : totalCount = pendingCount + approvedCount + rejectedCount

' 近期趋势 - 7天各状态
Dim recentPending : recentPending = CLng("0" & GetScalar("SELECT COUNT(*) FROM ProductReviews WHERE [Status]='Pending' AND CreatedAt>=DATEADD(DAY,-7,GETDATE())"))
Dim recentApproved : recentApproved = CLng("0" & GetScalar("SELECT COUNT(*) FROM ProductReviews WHERE [Status]='Approved' AND CreatedAt>=DATEADD(DAY,-7,GETDATE())"))
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>评价审核 - 运营管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
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
        .stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); padding: 20px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); text-align: center; cursor: pointer; transition: all 0.2s; }
        .stat-card:hover { transform: translateY(-2px); }
        .stat-card .num { font-size: 28px; font-weight: 700; color: #00bcd4; }
        .stat-card .num small { font-size: 14px; color: #888; font-weight: normal; }
        .stat-card .label { font-size: 13px; color: #888; margin-top: 5px; }
        .stat-card.pending .num { color: #FF9800; }
        .stat-card.approved .num { color: #4CAF50; }
        .stat-card.rejected .num { color: #f44336; }
        .msg { padding: 12px 20px; border-radius: 8px; margin-bottom: 20px; font-weight: 500; }
        .msg-success { background: rgba(76,175,80,0.15); color: #4CAF50; border: 1px solid rgba(76,175,80,0.3); }
        .filter-bar { background: linear-gradient(135deg, #2d2d44, #1e1e32); padding: 20px; border-radius: 12px; margin-bottom: 20px; border: 1px solid rgba(255,255,255,0.05); }
        .filter-bar form { display: flex; gap: 15px; align-items: center; flex-wrap: wrap; }
        .filter-bar select, .filter-bar input { padding: 10px 15px; border: 1px solid rgba(255,255,255,0.12); border-radius: 8px; background: #1a1a2e; color: #e0e0e0; font-size: 14px; }
        .filter-bar select:focus, .filter-bar input:focus { border-color: #00bcd4; outline: none; }
        .card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); margin-bottom: 20px; overflow: hidden; }
        .card-header { padding: 15px 20px; border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 16px; display: flex; justify-content: space-between; align-items: center; }
        .card-body { padding: 20px; }
        table { width: 100%; border-collapse: collapse; }
        th { text-align: left; padding: 12px; background: rgba(0,188,212,0.06); border-bottom: 1px solid rgba(255,255,255,0.06); font-size: 13px; color: #999; }
        td { padding: 12px; border-bottom: 1px solid rgba(255,255,255,0.04); font-size: 14px; }
        tr:hover { background: rgba(255,255,255,0.03); }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
        .badge-pending { background: rgba(255,152,0,0.2); color: #FFB74D; }
        .badge-approved { background: rgba(76,175,80,0.2); color: #A5D6A7; }
        .badge-rejected { background: rgba(244,67,54,0.2); color: #EF9A9A; }
        .stars { color: #FFC107; letter-spacing: 2px; }
        .stars .empty { color: rgba(255,255,255,0.15); }
        .comment-text { max-width: 250px; font-size: 13px; color: #ccc; line-height: 1.5; overflow: hidden; }
        .bulk-bar { display: flex; align-items: center; gap: 10px; padding: 12px 15px; background: rgba(0,188,212,0.08); border-radius: 8px; margin-bottom: 15px; display: none; }
        .bulk-bar.show { display: flex; }
        .no-data { text-align: center; padding: 40px; color: #666; }
        .no-data i { font-size: 40px; display: block; margin-bottom: 10px; opacity: 0.3; }
        .row-check { width: 18px; height: 18px; accent-color: #00bcd4; cursor: pointer; }
        @media (max-width: 768px) { .main-content { margin-left: 0; } .stats-grid { grid-template-columns: 1fr 1fr; } }
    </style>
</head>
<body data-theme="operation-dark">
<!--#include file="includes/nav.asp"-->
<div class="main-content">
    <div class="page-header">
        <h2 class="page-title"><i class="fas fa-star"></i> 评价审核管理</h2>
        <div class="breadcrumb"><a href="index.asp">运营中心</a> / 评价审核</div>
    </div>

    <% If msg <> "" Then %>
    <div class="msg msg-success"><i class="fas fa-check-circle"></i> <%= msg %></div>
    <% End If %>

    <div class="stats-grid">
        <div class="stat-card pending" onclick="location.href='?status=Pending'"><div class="num"><%= pendingCount %> <small>+<%= recentPending %></small></div><div class="label">待审核 (7日新增)</div></div>
        <div class="stat-card approved" onclick="location.href='?status=Approved'"><div class="num"><%= approvedCount %> <small>+<%= recentApproved %></small></div><div class="label">已通过 (7日新增)</div></div>
        <div class="stat-card rejected" onclick="location.href='?status=Rejected'"><div class="num"><%= rejectedCount %></div><div class="label">已拒绝</div></div>
        <div class="stat-card" onclick="location.href='?'"><div class="num"><%= totalCount %></div><div class="label">全部评价</div></div>
    </div>

    <div class="filter-bar">
        <form method="get" action="reviews_manage.asp">
            <select name="status">
                <option value="">全部状态</option>
                <option value="Pending" <%= IIf(filterStatus="Pending","selected","") %>>待审核</option>
                <option value="Approved" <%= IIf(filterStatus="Approved","selected","") %>>已通过</option>
                <option value="Rejected" <%= IIf(filterStatus="Rejected","selected","") %>>已拒绝</option>
            </select>
            <input type="text" name="keyword" value="<%= keyword %>" placeholder="搜索用户名/评价内容...">
            <button type="submit" class="btn btn-ghost"><i class="fas fa-search"></i> 筛选</button>
            <a href="?" class="btn btn-ghost"><i class="fas fa-undo"></i> 重置</a>
        </form>
    </div>

    <form method="post" id="bulkForm">
        <div class="bulk-bar" id="bulkBar">
            <span style="font-size:13px;color:#888;">已选 <strong id="selectedCount">0</strong> 条</span>
            <select name="bulk_action" style="padding:8px 12px;border-radius:6px;background:#1a1a2e;border:1px solid rgba(255,255,255,0.12);color:#e0e0e0;">
                <option value="">批量操作...</option>
                <option value="approve">批量通过</option>
                <option value="reject">批量拒绝</option>
                <option value="pending">重置为待审核</option>
            </select>
            <button type="submit" class="btn btn-success" onclick="return confirm('确认执行批量操作？')"><i class="fas fa-check-double"></i> 执行</button>
            <button type="button" class="btn btn-ghost" onclick="clearSelection()">取消选择</button>
        </div>
        <input type="hidden" name="review_ids" id="selectedIDs">

        <div class="card">
            <div class="card-header">
                <span><i class="fas fa-list"></i> 评价列表</span>
                <button type="button" class="btn btn-ghost" onclick="selectAllPending()"><i class="fas fa-check-square"></i> 全选待审</button>
            </div>
            <div class="card-body">
                <table>
                    <thead>
                        <tr>
                            <th style="width:35px;"><input type="checkbox" class="row-check" id="checkAll" onclick="toggleAll(this)"></th>
                            <th>ID</th><th>用户</th><th>订单</th><th>评分</th><th>评价内容</th><th>时间</th><th>状态</th><th>操作</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% If Not rsReviews Is Nothing Then
                        Dim hasRows : hasRows = False
                        Do While Not rsReviews.EOF
                            hasRows = True
                            Dim rStatus : rStatus = rsReviews("Status")
                            Dim rStatusBadge
                            Select Case rStatus
                                Case "Pending": rStatusBadge = "badge-pending"
                                Case "Approved": rStatusBadge = "badge-approved"
                                Case "Rejected": rStatusBadge = "badge-rejected"
                                Case Else: rStatusBadge = "badge-pending"
                            End Select
                            Dim rRating : rRating = CInt("0" & rsReviews("Rating") & "")
                            Dim rComment : rComment = rsReviews("Comment") & ""
                        %>
                        <tr>
                            <td><input type="checkbox" class="row-check review-check" value="<%= rsReviews("ReviewID") %>"></td>
                            <td>#<%= rsReviews("ReviewID") %></td>
                            <td><%= rsReviews("Username") %></td>
                            <td><%= IIf(IsNull(rsReviews("OrderNo")),"—",rsReviews("OrderNo")) %></td>
                            <td>
                                <div class="stars">
                                    <% Dim s
                                    For s = 1 To 5
                                        If s <= rRating Then Response.Write "<i class='fas fa-star'></i>" Else Response.Write "<i class='fas fa-star empty'></i>"
                                    Next %>
                                </div>
                            </td>
                            <td>
                                <div class="comment-text">
                                    <%= IIf(Len(rComment)>60,Left(rComment,60)&"...",rComment) %>
                                    <% If rComment = "" Then %><span style="color:#666;">（无文字）</span><% End If %>
                                </div>
                            </td>
                            <td style="font-size:13px;color:#999;"><%= rsReviews("CreatedAt") %></td>
                            <td><span class="badge <%= rStatusBadge %>"><%= rStatus %></span></td>
                            <td>
                                <form method="post" style="display:inline;">
                                    <input type="hidden" name="review_ids" value="<%= rsReviews("ReviewID") %>">
                                    <% If rStatus <> "Approved" Then %>
                                    <button type="submit" name="bulk_action" value="approve" class="btn btn-success" style="padding:5px 10px;font-size:11px;"><i class="fas fa-check"></i></button>
                                    <% End If %>
                                    <% If rStatus <> "Rejected" Then %>
                                    <button type="submit" name="bulk_action" value="reject" class="btn btn-danger" style="padding:5px 10px;font-size:11px;"><i class="fas fa-times"></i></button>
                                    <% End If %>
                                </form>
                            </td>
                        </tr>
                        <% 
                            rsReviews.MoveNext
                        Loop
                        rsReviews.Close
                        If Not hasRows Then
                        %>
                        <tr><td colspan="9" class="no-data"><i class="fas fa-comment-slash"></i>暂无评价数据</td></tr>
                        <% End If
                        Else %>
                        <tr><td colspan="9" class="no-data"><i class="fas fa-comment-slash"></i>暂无评价数据</td></tr>
                        <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
    </form>
</div>

<script>
function toggleAll(cb) {
    var checks = document.querySelectorAll('.review-check');
    checks.forEach(function(c) { c.checked = cb.checked; });
    updateBulkBar();
}
function updateBulkBar() {
    var checks = document.querySelectorAll('.review-check:checked');
    var ids = [];
    checks.forEach(function(c) { ids.push(c.value); });
    document.getElementById('selectedIDs').value = ids.join(',');
    document.getElementById('selectedCount').innerText = ids.length;
    document.getElementById('bulkBar').className = ids.length > 0 ? 'bulk-bar show' : 'bulk-bar';
}
function clearSelection() {
    document.querySelectorAll('.review-check').forEach(function(c) { c.checked = false; });
    document.getElementById('checkAll').checked = false;
    updateBulkBar();
}
function selectAllPending() {
    document.querySelectorAll('.review-check').forEach(function(c) {
        var row = c.closest('tr');
        var statusCell = row ? row.querySelector('.badge-pending') : null;
        c.checked = !!statusCell;
    });
    updateBulkBar();
}
document.querySelectorAll('.review-check').forEach(function(c) { c.addEventListener('change', updateBulkBar); });
document.getElementById('bulkForm').addEventListener('submit', function(e) {
    var action = this.querySelector('[name="bulk_action"]').value;
    if (!action) { e.preventDefault(); alert('请选择批量操作类型'); }
});
</script>
</body>
</html>
<%
Call CloseConnection()
%>
