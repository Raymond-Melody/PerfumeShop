<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/dal.asp"-->
<%
Call OpenConnection()

' 检查用户是否登录
If Session("UserID") = "" Or IsNull(Session("UserID")) Then
    Response.Redirect "/user/login.asp?return=" & Server.URLEncode(Request.ServerVariables("URL"))
End If

Dim userId, action, reviewId, reviewMsg, reviewErr
userId = CLng(Session("UserID"))
action = Request.QueryString("action")
reviewId = 0
If Request.QueryString("review_id") <> "" Then
    If IsNumeric(Request.QueryString("review_id")) Then
        reviewId = CLng(Request.QueryString("review_id"))
    End If
End If
reviewMsg = ""
reviewErr = ""

' ---- 删除评价 ----
If action = "delete" And reviewId > 0 Then
    ' 验证是该用户的评价
    Dim ownerCheck
    ownerCheck = DAL_GetScalar("SELECT COUNT(*) FROM ProductReviews WHERE ReviewID=@RID AND UserID=@UID", _
        Array(Array("@RID", DAL_adInteger, 0, reviewId), Array("@UID", DAL_adInteger, 0, userId)), 0)
    If CLng(ownerCheck) > 0 Then
        DAL_Execute "UPDATE ProductReviews SET IsActive = 0, UpdatedAt = GETDATE() WHERE ReviewID = @RID", _
            Array(Array("@RID", DAL_adInteger, 0, reviewId))
        reviewMsg = "评价已删除"
    Else
        reviewErr = "无权操作该评价"
    End If
End If

' ---- 分页参数 ----
Dim page, pageSize, pageInfo
page = 1
If Request.QueryString("page") <> "" Then
    If IsNumeric(Request.QueryString("page")) Then
        page = CLng(Request.QueryString("page"))
        If page < 1 Then page = 1
    End If
End If
pageSize = 10

' ---- 查询用户评价 ----
Dim reviewSQL, rsReviews
reviewSQL = "SELECT pr.*, p.ProductName, p.ProductImage, p.ProductID AS ProdID " & _
    "FROM ProductReviews pr " & _
    "LEFT JOIN Products p ON pr.ProductID = p.ProductID " & _
    "WHERE pr.UserID = @UID AND pr.IsActive = 1 " & _
    "ORDER BY pr.CreatedAt DESC"
Set rsReviews = DAL_GetListPaged(reviewSQL, _
    Array(Array("@UID", DAL_adInteger, 0, userId)), page, pageSize, pageInfo)

' ---- 统计 ----
Dim totalUserReviews
totalUserReviews = CLng(DAL_GetScalar("SELECT COUNT(*) FROM ProductReviews WHERE UserID=@UID AND IsActive=1", _
    Array(Array("@UID", DAL_adInteger, 0, userId)), 0))
%>
<!--#include file="../includes/header.asp"-->

<!-- 面包屑导航 -->
<div class="breadcrumb">
    <div class="container">
        <a href="/index.asp"><% If FEATURE_I18N Then %><%= T("breadcrumb_home", Empty) %><% Else %>首页<% End If %></a>
        <span class="separator">/</span>
        <a href="/user/index.asp"><% If FEATURE_I18N Then %><%= T("user_nav_center", Empty) %><% Else %>个人中心<% End If %></a>
        <span class="separator">/</span>
        <span><% If FEATURE_I18N Then %><%= T("user_my_reviews_title", Empty) %><% Else %>我的评价<% End If %></span>
    </div>
</div>

<div class="container">
    <div class="user-center">
        <!--#include file="nav.asp"-->

        <!-- 主内容区 -->
        <div class="user-main">
            <div class="welcome-section">
                <h1><i class="fas fa-star"></i> <% If FEATURE_I18N Then %><%= T("user_my_reviews_title", Empty) %><% Else %>我的评价<% End If %></h1>
                <p>共 <strong><%= totalUserReviews %></strong> 条评价 | <a href="/products.asp"><% If FEATURE_I18N Then %><%= T("user_favorites_browse", Empty) %><% Else %>浏览商品<% End If %> &raquo;</a></p>
            </div>

            <!-- 消息提示 -->
        <% If reviewMsg <> "" Then %>
        <div class="review-alert review-alert-success"><%= reviewMsg %></div>
        <% End If %>
        <% If reviewErr <> "" Then %>
        <div class="review-alert review-alert-error"><%= reviewErr %></div>
        <% End If %>

        <!-- 评价列表 -->
        <%
        Dim hasReviews
        hasReviews = False
        If Not rsReviews Is Nothing Then
            If Not rsReviews.EOF Then
                hasReviews = True
            End If
        End If
        If hasReviews Then
        %>
        <div class="my-review-list">
            <%
            Do While Not rsReviews.EOF
                Dim prodImg
                prodImg = rsReviews("ProductImage")
                If IsNull(prodImg) Or prodImg = "" Then prodImg = "/images/default-product.svg"
            %>
            <div class="my-review-card">
                <div class="my-review-product">
                    <img src="<%= prodImg %>" alt="<%= Server.HTMLEncode(rsReviews("ProductName")) %>" class="my-review-img" loading="lazy" onerror="this.src='/images/default-product.svg'">
                    <div class="my-review-product-info">
                        <a href="/product.asp?id=<%= rsReviews("ProdID") %>" class="my-review-product-name"><%= Server.HTMLEncode(rsReviews("ProductName")) %></a>
                        <div class="my-review-stars">
                            <%= String(rsReviews("Rating"), "★") %><%= String(5 - CInt(rsReviews("Rating")), "☆") %>
                            <span class="my-review-date"><%= FormatDateTime(rsReviews("CreatedAt"), 2) %></span>
                        </div>
                    </div>
                </div>
                <% If Not IsNull(rsReviews("Title")) And rsReviews("Title") <> "" Then %>
                <h4 class="my-review-title"><%= Server.HTMLEncode(rsReviews("Title")) %></h4>
                <% End If %>
                <p class="my-review-content"><%= Server.HTMLEncode(rsReviews("Content")) %></p>
                <div class="my-review-meta">
                    <% If rsReviews("IsVerifiedPurchase") Then %>
                    <span class="verified-badge"><i class="fas fa-check-circle"></i> 已验证购买</span>
                    <% End If %>
                    <span class="review-likes"><i class="far fa-heart"></i> <%= rsReviews("LikeCount") %></span>
                </div>
                <div class="my-review-actions">
                    <a href="/product.asp?id=<%= rsReviews("ProdID") %>#review-<%= rsReviews("ReviewID") %>" class="btn btn-sm btn-outline"><i class="fas fa-external-link-alt"></i> 查看</a>
                    <a href="javascript:void(0)" onclick="confirmDelete(<%= rsReviews("ReviewID") %>)" class="btn btn-sm btn-outline btn-danger"><i class="fas fa-trash-alt"></i> 删除</a>
                </div>
            </div>
            <%
                rsReviews.MoveNext
            Loop
            %>
        </div>

        <!-- 分页 -->
        <% If pageInfo("totalPages") > 1 Then %>
        <div class="pagination">
            <% If pageInfo("hasPrev") Then %>
            <a href="?page=<%= page - 1 %>" class="page-btn">&laquo; 上一页</a>
            <% End If %>
            <%
            Dim pg, pgStart, pgEnd
            pgStart = page - 2
            If pgStart < 1 Then pgStart = 1
            pgEnd = pgStart + 4
            If pgEnd > pageInfo("totalPages") Then
                pgEnd = pageInfo("totalPages")
                pgStart = pgEnd - 4
                If pgStart < 1 Then pgStart = 1
            End If
            For pg = pgStart To pgEnd
            %>
            <a href="?page=<%= pg %>" class="page-btn <%= IIf(pg = page, "active", "") %>"><%= pg %></a>
            <% Next %>
            <% If pageInfo("hasNext") Then %>
            <a href="?page=<%= page + 1 %>" class="page-btn">下一页 &raquo;</a>
            <% End If %>
        </div>
        <% End If %>

        <% Else %>
        <div class="empty-state">
            <i class="far fa-comment-dots"></i>
            <h3>暂无评价</h3>
            <p>你还没有对任何产品进行评价，快去探索并分享你的感受吧！</p>
            <a href="/products.asp" class="btn btn-primary">浏览商品</a>
        </div>
        <% End If %>

        <%
        If Not rsReviews Is Nothing Then
            rsReviews.Close
            Set rsReviews = Nothing
        End If
        %>
    </div><!-- /.user-main -->
</div><!-- /.user-center -->
</div><!-- /.container -->

<!-- 删除确认脚本 -->
<script nonce="<%= Session("csp_nonce") %>">
function confirmDelete(reviewId) {
    if (confirm('确定要删除这条评价吗？此操作不可撤销。')) {
        window.location.href = 'my_reviews.asp?action=delete&review_id=' + reviewId;
    }
}
</script>

<!-- 评价模块样式 -->
<style nonce="<%= Session("csp_nonce") %>">

.review-alert { padding: 12px 16px; border-radius: 8px; margin-bottom: 16px; font-size: .9rem; }
.review-alert-success { background: #e8f5e9; color: #2e7d32; border: 1px solid #c8e6c9; }
.review-alert-error { background: #fce4ec; color: #c62828; border: 1px solid #f8bbd0; }

.review-stats-mini { text-align: center; padding: 8px 0; }
.stat-item { display: flex; flex-direction: column; align-items: center; }
.stat-num { font-size: 2rem; font-weight: 700; color: #e91e63; line-height: 1.2; }
.stat-label { font-size: .8rem; color: #999; }

.my-review-list { display: flex; flex-direction: column; gap: 12px; }
.my-review-card { background: #fff; border-radius: 12px; padding: 16px 20px; box-shadow: 0 1px 4px rgba(0,0,0,.05); }
.my-review-product { display: flex; align-items: center; gap: 12px; margin-bottom: 10px; }
.my-review-img { width: 56px; height: 56px; border-radius: 8px; object-fit: cover; background: #f5f5f5; flex-shrink: 0; }
.my-review-product-info { min-width: 0; }
.my-review-product-name { font-weight: 600; font-size: .95rem; color: #333; text-decoration: none; display: block; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.my-review-product-name:hover { color: #e91e63; }
.my-review-stars { font-size: .85rem; color: #ff9800; margin-top: 2px; }
.my-review-date { color: #aaa; margin-left: 8px; font-size: .75rem; }
.my-review-title { font-size: .95rem; color: #444; margin: 0 0 6px; }
.my-review-content { font-size: .88rem; color: #666; line-height: 1.6; margin: 0 0 10px; }
.my-review-meta { display: flex; align-items: center; gap: 12px; margin-bottom: 10px; font-size: .8rem; }
.verified-badge { display: inline-flex; align-items: center; gap: 3px; background: #e8f5e9; color: #2e7d32; padding: 2px 8px; border-radius: 10px; font-size: .75rem; }
.review-likes { color: #999; }
.my-review-actions { display: flex; gap: 8px; }
.btn-sm { padding: 4px 12px; font-size: .8rem; border-radius: 6px; text-decoration: none; display: inline-flex; align-items: center; gap: 4px; cursor: pointer; }
.btn-outline { background: #fff; border: 1px solid #ddd; color: #666; }
.btn-outline:hover { border-color: #e91e63; color: #e91e63; }
.btn-danger { border-color: #fce4ec; color: #c62828; }
.btn-danger:hover { background: #fce4ec; }
.btn-primary { background: #e91e63; color: #fff; border: none; padding: 8px 20px; border-radius: 6px; text-decoration: none; display: inline-flex; align-items: center; gap: 6px; font-size: .9rem; }
.btn-primary:hover { background: #c2185b; }
.btn-block { display: block; text-align: center; justify-content: center; }

.empty-state { text-align: center; padding: 60px 20px; color: #bbb; }
.empty-state i { font-size: 3rem; display: block; margin-bottom: 12px; }
.empty-state h3 { font-size: 1.1rem; color: #888; margin: 0 0 8px; }
.empty-state p { font-size: .9rem; margin: 0 0 20px; }

.pagination { display: flex; justify-content: center; align-items: center; gap: 6px; margin-top: 20px; }
.pagination .page-btn { display: inline-flex; align-items: center; justify-content: center; min-width: 32px; height: 32px; padding: 0 8px; background: #fff; border: 1px solid #ddd; border-radius: 6px; font-size: .85rem; color: #666; text-decoration: none; transition: all .2s; }
.pagination .page-btn:hover { border-color: #e91e63; color: #e91e63; }
.pagination .page-btn.active { background: #e91e63; border-color: #e91e63; color: #fff; }

.products-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 18px; padding-bottom: 12px; border-bottom: 1px solid #eee; }
.products-title { font-size: 1.3rem; color: #333; margin: 0; display: flex; align-items: center; gap: 8px; }
.products-count { font-size: .85rem; color: #999; }
.cta-box { background: linear-gradient(135deg, #fce4ec, #f3e5f5); border-radius: 10px; padding: 16px; text-align: center; margin-top: 12px; }
.cta-box h3 { font-size: .95rem; margin: 0 0 6px; color: #333; }
.cta-box p { font-size: .8rem; color: #777; margin: 0 0 10px; }
</style>

<!--#include file="../includes/footer.asp"-->
<%
Call CloseConnection()
%>
