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

' 检查并处理POST请求（审核操作）
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' CSRF验证
    If Not ValidateCSRFToken() Then
        Session("ErrorMessage") = "安全验证失败，请重新操作"
        Response.Redirect "order_reviews.asp"
        Response.End
    End If
    
    Dim reviewId, action, newStatus, updateSql
    reviewId = Request.Form("review_id")
    action = Request.Form("action")
    
    If reviewId <> "" And IsNumeric(reviewId) Then
        Select Case action
            Case "approve"
                newStatus = "Approved"
            Case "reject"
                newStatus = "Rejected"
            Case Else
                newStatus = ""
        End Select
        
        If newStatus <> "" Then
            updateSql = "UPDATE ProductReviews SET [Status] = '" & SafeSQL(newStatus) & "', UpdatedAt = GETDATE() WHERE ReviewID = " & CLng(reviewId)
            If ExecuteNonQuery(updateSql) Then
                Session("SuccessMessage") = "评价审核操作成功"
                Call LogAdminAction("审核评价", "review", reviewId, newStatus, "")
            Else
                Session("ErrorMessage") = "评价审核操作失败: " & Session("LastDBError")
            End If
        End If
    End If
    
    Response.Redirect "order_reviews.asp" & BuildQueryString()
    Response.End
End If

' 构建查询字符串（保留筛选条件）
Function BuildQueryString()
    Dim qs, params
    params = ""
    If statusFilter <> "" Then params = params & "&status=" & Server.URLEncode(statusFilter)
    If keyword <> "" Then params = params & "&keyword=" & Server.URLEncode(keyword)
    If pageNum > 1 Then params = params & "&page=" & pageNum
    If params <> "" Then
        BuildQueryString = "?" & Mid(params, 2)
    Else
        BuildQueryString = ""
    End If
End Function

' 获取筛选参数
Dim statusFilter, keyword, pageNum, pageSize
statusFilter = Request.QueryString("status")
keyword = Request.QueryString("keyword")
pageSize = 20

' 获取页码
If Request.QueryString("page") <> "" And IsNumeric(Request.QueryString("page")) Then
    pageNum = CLng(Request.QueryString("page"))
    If pageNum < 1 Then pageNum = 1
Else
    pageNum = 1
End If

' 构建查询条件
Dim whereClause
whereClause = "WHERE 1=1"

If statusFilter <> "" Then
    whereClause = whereClause & " AND r.[Status] = '" & SafeSQL(statusFilter) & "'"
End If

If keyword <> "" Then
    whereClause = whereClause & " AND (u.Username LIKE '%" & SafeSQL(keyword) & "%' OR r.Comment LIKE '%" & SafeSQL(keyword) & "%' OR o.OrderNo LIKE '%" & SafeSQL(keyword) & "%')"
End If

' 获取总记录数
Dim totalRecords, totalPages
totalRecords = 0
Dim rsCount
Set rsCount = ExecuteQuery("SELECT COUNT(*) FROM ProductReviews r LEFT JOIN Users u ON r.UserID = u.UserID LEFT JOIN Orders o ON r.OrderID = o.OrderID " & whereClause)
If Not rsCount Is Nothing Then
    If Not rsCount.EOF Then
        totalRecords = rsCount(0).Value
        If IsNull(totalRecords) Then totalRecords = 0
    End If
    rsCount.Close
    Set rsCount = Nothing
End If

totalPages = Int((totalRecords + pageSize - 1) / pageSize)
If totalPages < 1 Then totalPages = 1
If pageNum > totalPages Then pageNum = totalPages

' 预先计算统计数据（必须在rsReviews打开前执行，因为Access不支持多活动记录集）
Dim pendingCount, pendingVal, approvedCount, approvedVal, rejectedCount, rejectedVal, totalCount, totalVal
pendingVal = GetScalar("SELECT COUNT(*) FROM ProductReviews WHERE [Status] = 'Pending'")
pendingCount = CLng("0" & pendingVal)
approvedVal = GetScalar("SELECT COUNT(*) FROM ProductReviews WHERE [Status] = 'Approved'")
approvedCount = CLng("0" & approvedVal)
rejectedVal = GetScalar("SELECT COUNT(*) FROM ProductReviews WHERE [Status] = 'Rejected'")
rejectedCount = CLng("0" & rejectedVal)
totalVal = GetScalar("SELECT COUNT(*) FROM ProductReviews")
totalCount = CLng("0" & totalVal)

' 获取评价列表（使用Recordset分页）
Dim rsReviews
Set rsReviews = Server.CreateObject("ADODB.Recordset")
rsReviews.CursorType = 1  ' adOpenKeyset
rsReviews.LockType = 1    ' adLockReadOnly
rsReviews.PageSize = pageSize

Dim reviewSql
reviewSql = "SELECT r.*, u.Username, u.FullName, o.OrderNo FROM ProductReviews r LEFT JOIN Users u ON r.UserID = u.UserID LEFT JOIN Orders o ON r.OrderID = o.OrderID " & whereClause & " ORDER BY r.CreatedAt DESC"

rsReviews.Open reviewSql, conn

If Not rsReviews.EOF Then
    rsReviews.AbsolutePage = pageNum
End If

' 记录访问日志
Call LogAdminAction("查看评价列表", "operation_reviews", "", "", "")

' 获取提示消息
Dim successMsg, errorMsg
successMsg = Session("SuccessMessage")
errorMsg = Session("ErrorMessage")
Session("SuccessMessage") = ""
Session("ErrorMessage") = ""
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>评价管理 - 运营管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        .filter-bar { background: white; padding: 20px; border-radius: 10px; margin-bottom: 25px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
        .filter-form { display: flex; gap: 15px; flex-wrap: wrap; align-items: flex-end; }
        .filter-group { display: flex; flex-direction: column; gap: 5px; }
        .filter-group label { font-size: 13px; color: #666; font-weight: 500; }
        .filter-group select, .filter-group input { padding: 10px 15px; border: 2px solid #e0e0e0; border-radius: 8px; font-size: 14px; min-width: 150px; }
        .filter-group select:focus, .filter-group input:focus { border-color: #667eea; outline: none; }
        
        .reviews-table { width: 100%; background: white; border-radius: 10px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
        .reviews-table th { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 15px; text-align: left; font-weight: 500; }
        .reviews-table td { padding: 15px; border-bottom: 1px solid #f0f0f0; vertical-align: top; }
        .reviews-table tr:hover { background: #f8f9fa; }
        .reviews-table tr:last-child td { border-bottom: none; }
        
        .review-id { font-weight: 600; color: #667eea; font-size: 14px; }
        .customer-info { display: flex; flex-direction: column; }
        .customer-name { font-weight: 500; color: #333; }
        .customer-username { font-size: 12px; color: #999; margin-top: 3px; }
        
        .order-no { font-family: monospace; color: #666; font-size: 13px; }
        
        .rating-stars { color: #ffc107; font-size: 14px; }
        .rating-stars .empty { color: #ddd; }
        
        .comment-preview { max-width: 250px; font-size: 13px; color: #555; line-height: 1.5; }
        .comment-preview.empty { color: #999; font-style: italic; }
        
        .review-time { font-size: 12px; color: #999; }
        
        .status-badge { display: inline-block; padding: 6px 14px; border-radius: 20px; font-size: 12px; font-weight: 500; }
        .status-pending { background: #fff3e0; color: #e65100; }
        .status-approved { background: #e8f5e9; color: #2e7d32; }
        .status-rejected { background: #ffebee; color: #c62828; }
        
        .action-btns { display: flex; gap: 8px; flex-wrap: wrap; }
        .action-btn { padding: 6px 12px; border-radius: 6px; font-size: 12px; text-decoration: none; transition: all 0.3s; border: none; cursor: pointer; }
        .action-btn.approve { background: #e8f5e9; color: #2e7d32; }
        .action-btn.approve:hover { background: #2e7d32; color: white; }
        .action-btn.reject { background: #ffebee; color: #c62828; }
        .action-btn.reject:hover { background: #c62828; color: white; }
        .action-btn.view { background: #e3f2fd; color: #1976d2; }
        .action-btn.view:hover { background: #1976d2; color: white; }
        
        .pagination { display: flex; justify-content: center; gap: 10px; margin-top: 25px; }
        .pagination a, .pagination span { padding: 10px 15px; background: white; border-radius: 8px; text-decoration: none; color: #666; box-shadow: 0 2px 5px rgba(0,0,0,0.08); }
        .pagination a:hover { background: #667eea; color: white; }
        .pagination span.active { background: #667eea; color: white; }
        .pagination .disabled { color: #ccc; cursor: not-allowed; }
        
        .stats-summary { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }
        .summary-item { background: white; padding: 20px; border-radius: 10px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
        .summary-number { font-size: 24px; font-weight: bold; margin-bottom: 5px; }
        .summary-label { font-size: 13px; color: #666; }
        
        .empty-state { text-align: center; padding: 60px 20px; background: white; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
        .empty-state i { font-size: 64px; color: #ddd; margin-bottom: 20px; }
        .empty-state h3 { color: #666; margin-bottom: 10px; }
        .empty-state p { color: #999; }
        
        .alert { padding: 15px 20px; border-radius: 8px; margin-bottom: 20px; }
        .alert-success { background: #e8f5e9; color: #2e7d32; border-left: 4px solid #4CAF50; }
        .alert-error { background: #ffebee; color: #c62828; border-left: 4px solid #f44336; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-star"></i> 评价管理</h2>
            <div class="breadcrumb">
                <a href="index.asp">运营中心</a> / <span>评价管理</span>
            </div>
        </div>
        
        <% If successMsg <> "" Then %>
        <div class="alert alert-success"><i class="fas fa-check-circle"></i> <%= SafeOutput(successMsg) %></div>
        <% End If %>
        <% If errorMsg <> "" Then %>
        <div class="alert alert-error"><i class="fas fa-exclamation-circle"></i> <%= SafeOutput(errorMsg) %></div>
        <% End If %>
        
        <!-- 统计概览 -->
        <div class="stats-summary">
            <div class="summary-item">
                <div class="summary-number" style="color: #ff9800;"><%= pendingCount %></div>
                <div class="summary-label">待审核</div>
            </div>
            <div class="summary-item">
                <div class="summary-number" style="color: #4CAF50;"><%= approvedCount %></div>
                <div class="summary-label">已通过</div>
            </div>
            <div class="summary-item">
                <div class="summary-number" style="color: #f44336;"><%= rejectedCount %></div>
                <div class="summary-label">已拒绝</div>
            </div>
            <div class="summary-item">
                <div class="summary-number" style="color: #333;"><%= totalCount %></div>
                <div class="summary-label">全部评价</div>
            </div>
        </div>
        
        <!-- 筛选栏 -->
        <div class="filter-bar">
            <form class="filter-form" method="get" action="order_reviews.asp">
                <div class="filter-group">
                    <label>评价状态</label>
                    <select name="status">
                        <option value="">全部状态</option>
                        <option value="Pending" <%= IIf(statusFilter="Pending", "selected", "") %>>待审核</option>
                        <option value="Approved" <%= IIf(statusFilter="Approved", "selected", "") %>>已通过</option>
                        <option value="Rejected" <%= IIf(statusFilter="Rejected", "selected", "") %>>已拒绝</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label>关键词搜索</label>
                    <input type="text" name="keyword" value="<%= SafeOutput(keyword) %>" placeholder="用户名/评价内容/订单号">
                </div>
                <div class="filter-group">
                    <button type="submit" class="admin-btn admin-btn-primary"><i class="fas fa-search"></i> 筛选</button>
                    <a href="order_reviews.asp" class="admin-btn admin-btn-secondary"><i class="fas fa-undo"></i> 重置</a>
                </div>
            </form>
        </div>
        
        <!-- 评价列表 -->
        <% If totalRecords = 0 Then %>
        <div class="empty-state">
            <i class="fas fa-comment-slash"></i>
            <h3>暂无评价数据</h3>
            <p>当前没有符合条件的用户评价记录</p>
        </div>
        <% Else %>
        <table class="reviews-table">
            <thead>
                <tr>
                    <th>评价ID</th>
                    <th>用户</th>
                    <th>订单号</th>
                    <th>评分</th>
                    <th>评价内容</th>
                    <th>评价时间</th>
                    <th>状态</th>
                    <th>操作</th>
                </tr>
            </thead>
            <tbody>
                <% 
                Dim i, rowCount
                rowCount = 0
                Do While Not rsReviews.EOF And rowCount < pageSize
                    rowCount = rowCount + 1
                %>
                <tr>
                    <td>
                        <div class="review-id">#<%= rsReviews("ReviewID") %></div>
                    </td>
                    <td>
                        <div class="customer-info">
                            <span class="customer-name"><%= SafeOutput(IIf(IsNull(rsReviews("FullName")) Or rsReviews("FullName")="", rsReviews("Username"), rsReviews("FullName"))) %></span>
                            <% If Not IsNull(rsReviews("Username")) And rsReviews("Username") <> "" Then %>
                            <span class="customer-username">@<%= SafeOutput(rsReviews("Username")) %></span>
                            <% End If %>
                        </div>
                    </td>
                    <td>
                        <div class="order-no"><%= SafeOutput(IIf(IsNull(rsReviews("OrderNo")) Or rsReviews("OrderNo")="", "-", rsReviews("OrderNo"))) %></div>
                    </td>
                    <td>
                        <div class="rating-stars">
                            <% 
                            Dim rating, starIdx
                            If IsNull(rsReviews("Rating")) Or rsReviews("Rating") = "" Then
                                rating = 0
                            Else
                                rating = CInt(rsReviews("Rating"))
                            End If
                            For starIdx = 1 To 5
                                If starIdx <= rating Then
                                    Response.Write "<i class='fas fa-star'></i>"
                                Else
                                    Response.Write "<i class='fas fa-star empty'></i>"
                                End If
                            Next
                            %>
                        </div>
                    </td>
                    <td>
                        <% 
                        Dim commentText
                        If IsNull(rsReviews("Comment")) Or rsReviews("Comment") = "" Then
                            commentText = ""
                        Else
                            commentText = CStr(rsReviews("Comment"))
                        End If
                        If commentText = "" Then
                        %>
                        <div class="comment-preview empty">（无评价内容）</div>
                        <% Else %>
                        <div class="comment-preview" title="<%= SafeOutput(commentText) %>">
                            <% 
                            If Len(commentText) > 50 Then
                                Response.Write SafeOutput(Left(commentText, 50)) & "..."
                            Else
                                Response.Write SafeOutput(commentText)
                            End If
                            %>
                        </div>
                        <% End If %>
                    </td>
                    <td>
                        <div class="review-time">
                            <% If IsNull(rsReviews("CreatedAt")) Or rsReviews("CreatedAt") = "" Or IsEmpty(rsReviews("CreatedAt")) Then %>
                            -
                            <% Else %>
                            <%= SafeFormatDateTime(rsReviews("CreatedAt"), 2) %>
                            <% End If %>
                        </div>
                    </td>
                    <td>
                        <% 
                        Dim reviewStatus
                        If IsNull(rsReviews("Status")) Or rsReviews("Status") = "" Then
                            reviewStatus = "Pending"
                        Else
                            reviewStatus = rsReviews("Status")
                        End If
                        Select Case reviewStatus
                            Case "Pending" 
                        %>
                            <span class="status-badge status-pending">待审核</span>
                        <% Case "Approved" %>
                            <span class="status-badge status-approved">已通过</span>
                        <% Case "Rejected" %>
                            <span class="status-badge status-rejected">已拒绝</span>
                        <% Case Else %>
                            <span class="status-badge status-pending">待审核</span>
                        <% End Select %>
                    </td>
                    <td>
                        <div class="action-btns">
                            <% If reviewStatus = "Pending" Then %>
                            <form method="post" action="order_reviews.asp<%= BuildQueryString() %>" style="display:inline;">
                                <%= GetCSRFTokenField() %>
                                <input type="hidden" name="review_id" value="<%= rsReviews("ReviewID") %>">
                                <input type="hidden" name="action" value="approve">
                                <button type="submit" class="action-btn approve" onclick="return confirm('确定要通过这条评价吗？')"><i class="fas fa-check"></i> 通过</button>
                            </form>
                            <form method="post" action="order_reviews.asp<%= BuildQueryString() %>" style="display:inline;">
                                <%= GetCSRFTokenField() %>
                                <input type="hidden" name="review_id" value="<%= rsReviews("ReviewID") %>">
                                <input type="hidden" name="action" value="reject">
                                <button type="submit" class="action-btn reject" onclick="return confirm('确定要拒绝这条评价吗？')"><i class="fas fa-times"></i> 拒绝</button>
                            </form>
                            <% ElseIf reviewStatus = "Rejected" Then %>
                            <form method="post" action="order_reviews.asp<%= BuildQueryString() %>" style="display:inline;">
                                <%= GetCSRFTokenField() %>
                                <input type="hidden" name="review_id" value="<%= rsReviews("ReviewID") %>">
                                <input type="hidden" name="action" value="approve">
                                <button type="submit" class="action-btn approve" onclick="return confirm('确定要重新通过这条评价吗？')"><i class="fas fa-check"></i> 通过</button>
                            </form>
                            <% Else %>
                            <form method="post" action="order_reviews.asp<%= BuildQueryString() %>" style="display:inline;">
                                <%= GetCSRFTokenField() %>
                                <input type="hidden" name="review_id" value="<%= rsReviews("ReviewID") %>">
                                <input type="hidden" name="action" value="reject">
                                <button type="submit" class="action-btn reject" onclick="return confirm('确定要取消通过这条评价吗？')"><i class="fas fa-times"></i> 取消</button>
                            </form>
                            <% End If %>
                        </div>
                    </td>
                </tr>
                <% 
                    rsReviews.MoveNext
                Loop 
                %>
            </tbody>
        </table>
        
        <!-- 分页 -->
        <% If totalPages > 1 Then %>
        <div class="pagination">
            <% If pageNum > 1 Then %>
            <a href="order_reviews.asp?page=<%= pageNum - 1 %><%= IIf(statusFilter<>"", "&status=" & Server.URLEncode(statusFilter), "") %><%= IIf(keyword<>"", "&keyword=" & Server.URLEncode(keyword), "") %>"><i class="fas fa-chevron-left"></i></a>
            <% Else %>
            <span class="disabled"><i class="fas fa-chevron-left"></i></span>
            <% End If %>
            
            <% 
            Dim startPage, endPage
            startPage = IIf(pageNum - 2 < 1, 1, pageNum - 2)
            endPage = IIf(startPage + 4 > totalPages, totalPages, startPage + 4)
            If endPage - startPage < 4 And startPage > 1 Then
                startPage = IIf(endPage - 4 < 1, 1, endPage - 4)
            End If
            
            For i = startPage To endPage
            %>
            <% If i = pageNum Then %>
            <span class="active"><%= i %></span>
            <% Else %>
            <a href="order_reviews.asp?page=<%= i %><%= IIf(statusFilter<>"", "&status=" & Server.URLEncode(statusFilter), "") %><%= IIf(keyword<>"", "&keyword=" & Server.URLEncode(keyword), "") %>"><%= i %></a>
            <% End If %>
            <% Next %>
            
            <% If pageNum < totalPages Then %>
            <a href="order_reviews.asp?page=<%= pageNum + 1 %><%= IIf(statusFilter<>"", "&status=" & Server.URLEncode(statusFilter), "") %><%= IIf(keyword<>"", "&keyword=" & Server.URLEncode(keyword), "") %>"><i class="fas fa-chevron-right"></i></a>
            <% Else %>
            <span class="disabled"><i class="fas fa-chevron-right"></i></span>
            <% End If %>
        </div>
        <% End If %>
        <% End If %>
    </div>
</body>
</html>
<%
If Not rsReviews Is Nothing Then
    If rsReviews.State = 1 Then rsReviews.Close
    Set rsReviews = Nothing
End If
Call CloseConnection()
%>
